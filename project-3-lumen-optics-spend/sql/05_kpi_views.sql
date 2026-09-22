/*
================================================================================
Project 3 -- Lumen Optics Manufacturing: Photonics Spend Scorecard
Script:  05_kpi_views.sql
Purpose: The analytical payload. Everything here reads its costs from
         dbo.fn_POLineCost so there is exactly one definition of what a
         purchase cost and what it should have cost.

THE PRICE EROSION MODEL (the centrepiece)

    Photonic components follow learning curves: unit prices decline as volumes
    rise and processes mature. For a vendor-part pair first bought at P0, the
    price the market should be offering today is

        ExpectedPrice = P0 x (1 - CategoryErosion) ^ YearsElapsed

    and the gap between that and the price actually being paid is margin the
    category gave up and Lumen did not collect:

        ErosionGapPerUnit  = LastPaidPrice - ExpectedPrice
        AnnualOpportunity  = ErosionGapPerUnit x TrailingTwelveMonthVolume

    This is the finding a variance report cannot produce. Purchase price
    variance compares what was paid against what was AGREED, so a supplier who
    renews flat every year scores zero variance forever while quietly keeping
    everything the learning curve was supposed to hand back.

WHAT THIS MODEL DOES NOT DO -- stated here rather than discovered later
    It measures the failure to erode FROM THE FIRST OBSERVED PRICE. It cannot
    say whether that first price was fair: a part bought badly in 2023 and
    eroded perfectly since will look healthy here. Answering that needs a
    should-cost build-up from materials and process, which is a different
    exercise with different inputs, and the case study says so rather than
    letting the reader assume this covers it.

LEVERAGE IS PART OF THE ANSWER, NOT A FOOTNOTE
    The largest opportunity is often on a sole-source component with a
    nine-month requalification, where there is no credible threat to make and
    therefore no negotiation to win. Ranking on opportunity alone sends a
    category manager into meetings they cannot win, which is the fastest way to
    lose a sourcing team's trust in a report. The queue therefore scores
    leverage explicitly and routes the unwinnable ones somewhere useful.

DATA DISCLOSURE: Lumen Optics Manufacturing is fictional; all data is synthetic.
================================================================================
*/

USE LumenSpend;
GO

IF OBJECT_ID('dbo.vw_RenegotiationQueue', 'V') IS NOT NULL DROP VIEW dbo.vw_RenegotiationQueue;
IF OBJECT_ID('dbo.fn_RenegotiationQueue','IF') IS NOT NULL DROP FUNCTION dbo.fn_RenegotiationQueue;
IF OBJECT_ID('dbo.vw_VendorScorecard', 'V')    IS NOT NULL DROP VIEW dbo.vw_VendorScorecard;
IF OBJECT_ID('dbo.fn_VendorScorecard', 'IF')   IS NOT NULL DROP FUNCTION dbo.fn_VendorScorecard;
IF OBJECT_ID('dbo.vw_PriceErosion', 'V')       IS NOT NULL DROP VIEW dbo.vw_PriceErosion;
IF OBJECT_ID('dbo.fn_PriceErosion', 'IF')      IS NOT NULL DROP FUNCTION dbo.fn_PriceErosion;
IF OBJECT_ID('dbo.vw_SpendKPIMonthly', 'V')    IS NOT NULL DROP VIEW dbo.vw_SpendKPIMonthly;
IF OBJECT_ID('dbo.fn_SpendKPI', 'IF')          IS NOT NULL DROP FUNCTION dbo.fn_SpendKPI;
GO

/*
--------------------------------------------------------------------------------
fn_PriceErosion(@AsOf, @MinLines, @MinDays)

One row per vendor-part pair with enough history to judge. The thresholds are
parameters, not constants: a pair seen three times over two months has a price
"trend" that is noise, and including it would fill the priority queue with
accidents.

Prices are volume-weighted within the first and last quarter of the pair's
history rather than taken from single transactions, because one small rush
order at a bad price should not set the baseline the whole opportunity is
measured from.
--------------------------------------------------------------------------------
*/
CREATE FUNCTION dbo.fn_PriceErosion (@AsOf DATE, @MinLines INT, @MinDays INT)
RETURNS TABLE
AS RETURN
(
    WITH L AS (
        SELECT c.PartKey, c.VendorKey, c.OrderDate, c.OrderQty, c.UnitPrice,
               c.ExtendedPrice, c.LandedCost, c.QtyAccepted, c.RejectedValue,
               c.PPVAmount, c.IsOnContract
        FROM dbo.fn_POLineCost(@AsOf) c
        -- duplicated requisitions would double-weight a price point
        WHERE c.PONumber NOT LIKE 'PO-D%'
    ),
    Span AS (
        SELECT PartKey, VendorKey,
               FirstDate = MIN(OrderDate), LastDate = MAX(OrderDate),
               Lines = COUNT(*), TotalQty = SUM(OrderQty),
               TotalSpend = SUM(ExtendedPrice), TotalLanded = SUM(LandedCost)
        FROM L GROUP BY PartKey, VendorKey
    ),
    Windowed AS (
        SELECT s.*,
               -- volume-weighted price over the first and last 90 days of the
               -- pair's own history, so a single outlier cannot set the baseline
               FirstPrice = (SELECT SUM(l2.UnitPrice * l2.OrderQty) / NULLIF(SUM(l2.OrderQty), 0)
                             FROM L l2 WHERE l2.PartKey = s.PartKey AND l2.VendorKey = s.VendorKey
                               AND l2.OrderDate <= DATEADD(DAY, 90, s.FirstDate)),
               LastPrice  = (SELECT SUM(l2.UnitPrice * l2.OrderQty) / NULLIF(SUM(l2.OrderQty), 0)
                             FROM L l2 WHERE l2.PartKey = s.PartKey AND l2.VendorKey = s.VendorKey
                               AND l2.OrderDate >= DATEADD(DAY, -90, s.LastDate)),
               -- trailing twelve months of volume is what an opportunity annualises against
               TTMQty   = (SELECT ISNULL(SUM(l2.OrderQty), 0)
                           FROM L l2 WHERE l2.PartKey = s.PartKey AND l2.VendorKey = s.VendorKey
                             AND l2.OrderDate > DATEADD(MONTH, -12, @AsOf)),
               TTMSpend = (SELECT ISNULL(SUM(l2.ExtendedPrice), 0)
                           FROM L l2 WHERE l2.PartKey = s.PartKey AND l2.VendorKey = s.VendorKey
                             AND l2.OrderDate > DATEADD(MONTH, -12, @AsOf))
        FROM Span s
        -- The 180-day floor is enforced HERE, not left to the caller's
        -- arguments. Below it the first-90 and last-90 day windows overlap and
        -- the same orders are averaged into both the baseline and the current
        -- price, which yields an erosion rate near zero for any pair -- a
        -- confident, meaningless number.
        WHERE s.Lines >= @MinLines
          AND DATEDIFF(DAY, s.FirstDate, s.LastDate) >= @MinDays
          AND DATEDIFF(DAY, s.FirstDate, s.LastDate) >= 180
    ),
    -- Years and actual erosion are ROUNDED here, and everything derived from
    -- them below uses the rounded values. Publishing a figure computed from an
    -- unrounded intermediate means a reader redoing the arithmetic from the
    -- printed numbers gets a different answer than the report.
    Rounded AS (
        SELECT w.*,
               YearsR  = CAST(DATEDIFF(DAY, w.FirstDate, w.LastDate) / 365.25 AS DECIMAL(9,3)),
               ActualR = CAST(100.0 * (1 - POWER(w.LastPrice / NULLIF(w.FirstPrice, 0),
                             365.25 / NULLIF(DATEDIFF(DAY, w.FirstDate, w.LastDate), 0))) AS DECIMAL(9,3))
        FROM Windowed w
    )
    SELECT
        w.PartKey, w.VendorKey,
        w.FirstDate, w.LastDate, w.Lines, w.TotalQty, w.TotalSpend, w.TotalLanded,
        w.TTMQty, w.TTMSpend,
        YearsElapsed = w.YearsR,
        FirstPrice = CAST(w.FirstPrice AS DECIMAL(12,4)),
        LastPrice  = CAST(w.LastPrice  AS DECIMAL(12,4)),
        b.AnnualErosionPct AS BenchmarkErosionPct,
        b.ToleranceBandPct,
        -- what the market should be quoting by now
        ExpectedPrice = CAST(w.FirstPrice * POWER(1.0 - b.AnnualErosionPct/100.0, w.YearsR) AS DECIMAL(12,4)),
        -- what the supplier actually delivered, annualised
        ActualErosionPct = w.ActualR,
        ErosionGapPerUnit = CAST(w.LastPrice - CAST(w.FirstPrice * POWER(1.0 - b.AnnualErosionPct/100.0, w.YearsR) AS DECIMAL(12,4)) AS DECIMAL(12,4)),
        AnnualOpportunity = CAST((w.LastPrice - CAST(w.FirstPrice * POWER(1.0 - b.AnnualErosionPct/100.0, w.YearsR) AS DECIMAL(12,4))) * w.TTMQty AS DECIMAL(16,2)),
        -- share of the expected erosion actually realised; the headline KPI
        ErosionCapturePct = CAST(CASE WHEN b.AnnualErosionPct = 0 THEN NULL
             ELSE 100.0 * w.ActualR / b.AnnualErosionPct END AS DECIMAL(9,2)),
        AsOfDate = @AsOf
    FROM Rounded w
    JOIN dbo.Dim_Part p ON p.PartKey = w.PartKey
    JOIN dbo.Ref_PriceErosionBenchmark b ON b.Category = p.Category
    WHERE w.FirstPrice > 0 AND w.LastPrice > 0
);
GO

CREATE VIEW dbo.vw_PriceErosion AS
SELECT e.*, p.PartNumber, p.PartName, p.Category, p.Criticality,
       p.QualifiedSupplierCount, p.QualificationMonths,
       v.VendorID, v.VendorName, v.VendorTier, v.LumenRevenueSharePct
FROM dbo.fn_PriceErosion('2025-12-31', 10, 500) e
JOIN dbo.Dim_Part   p ON p.PartKey   = e.PartKey
JOIN dbo.Dim_Vendor v ON v.VendorKey = e.VendorKey;
GO

/*
--------------------------------------------------------------------------------
fn_VendorScorecard(@AsOf) -- price, quality, delivery and compliance in one row
per vendor.

CostPerAcceptedUnit cannot be averaged across vendors, because vendors supply
different parts at different base prices -- an average across a mixed basket
compares nothing. The scorecard therefore reports the INDEX: for each part a
vendor shares with another supplier, how its cost per accepted unit compares
with the best available on that same part. That is a like-for-like number.
--------------------------------------------------------------------------------
*/
CREATE FUNCTION dbo.fn_VendorScorecard (@AsOf DATE)
RETURNS TABLE
AS RETURN
(
    -- Duplicates are excluded by BEHAVIOUR, via the shared definition in
    -- fn_DuplicatePOLines, not by the generator's 'PO-D' prefix. The prefix
    -- test returned the identical 26 lines on this dataset, which is precisely
    -- why it survived review: it agreed with the right answer while depending
    -- on a marker that exists only because this data is synthetic.
    WITH C AS (
        SELECT c.* FROM dbo.fn_POLineCost(@AsOf) c
        WHERE NOT EXISTS (SELECT 1 FROM dbo.fn_DuplicatePOLines(@AsOf) d
                          WHERE d.POLineKey = c.POLineKey)
    ),
    PartVendor AS (
        SELECT PartKey, VendorKey,
               Landed = SUM(LandedCost),
               Accepted = SUM(ISNULL(QtyAccepted, 0)),
               CPA = SUM(LandedCost) / NULLIF(SUM(ISNULL(QtyAccepted,0)), 0)
        FROM C GROUP BY PartKey, VendorKey
        HAVING SUM(ISNULL(QtyAccepted,0)) > 0
    ),
    PartBest AS (
        SELECT PartKey, BestCPA = MIN(CPA), Suppliers = COUNT(*)
        FROM PartVendor GROUP BY PartKey
    ),
    -- The like-for-like index is a PAIR-grain quantity, so it has to be
    -- aggregated at pair grain. Summing pv.Landed across the line-level join
    -- instead adds each pair's landed cost once per PO line in that pair.
    -- That fan-out survives review because numerator and denominator both
    -- inflate: the result stays above 100, stays plausible, and is quietly a
    -- line-count-weighted average of the index rather than the index. It moved
    -- 33 of 36 vendors' rank positions, and CostIndexVsBest carries 40% of the
    -- weight in usp_VendorScorecard's CompositeScore.
    VendorIndex AS (
        SELECT pv.VendorKey,
               Landed         = SUM(pv.Landed),
               BestEquivalent = SUM(pb.BestCPA * pv.Accepted),
               SharedParts    = COUNT(CASE WHEN pb.Suppliers > 1 THEN pv.PartKey END)
        FROM PartVendor pv
        JOIN PartBest pb ON pb.PartKey = pv.PartKey
        GROUP BY pv.VendorKey
    )
    SELECT
        c.VendorKey,
        Lines        = COUNT(*),
        TotalSpend   = CAST(SUM(c.ExtendedPrice) AS DECIMAL(16,2)),
        LandedSpend  = CAST(SUM(c.LandedCost) AS DECIMAL(16,2)),
        FreightSpend = CAST(SUM(c.FreightAmount) AS DECIMAL(16,2)),
        ExpediteSpend= CAST(SUM(c.ExpediteFee) AS DECIMAL(16,2)),
        ExpeditePctOfSpend = CAST(100.0 * SUM(c.ExpediteFee) / NULLIF(SUM(c.ExtendedPrice), 0) AS DECIMAL(6,2)),
        OnContractPct = CAST(100.0 * SUM(CAST(c.IsOnContract AS INT)) / COUNT(*) AS DECIMAL(6,2)),
        PPVAmount    = CAST(SUM(c.PPVAmount) AS DECIMAL(16,2)),
        PPVPct       = CAST(100.0 * SUM(c.PPVAmount) / NULLIF(SUM(c.ContractedExtended), 0) AS DECIMAL(6,3)),
        QtyReceived  = SUM(ISNULL(c.QtyReceived, 0)),
        QtyAccepted  = SUM(ISNULL(c.QtyAccepted, 0)),
        AcceptanceRatePct = CAST(100.0 * SUM(ISNULL(c.QtyAccepted,0))
                               / NULLIF(SUM(ISNULL(c.QtyReceived,0)), 0) AS DECIMAL(6,2)),
        RejectedValue = CAST(SUM(c.RejectedValue) AS DECIMAL(16,2)),
        OnTimePct    = CAST(100.0 * SUM(CASE WHEN c.IsOnTime = 1 THEN 1 ELSE 0 END)
                          / NULLIF(SUM(CASE WHEN c.IsOnTime IS NULL THEN 0 ELSE 1 END), 0) AS DECIMAL(6,2)),
        AvgDaysLate  = CAST(AVG(CAST(c.DaysLate AS FLOAT)) AS DECIMAL(9,2)),
        -- like-for-like cost position: 100 = best available on the same parts,
        -- 108 = eight per cent dearer per usable unit than the best alternative.
        -- MAX() over a value that is constant within the group is just "read the
        -- pre-aggregated figure" -- the pair-grain arithmetic already happened
        -- in VendorIndex and must not be redone across lines.
        CostIndexVsBest = CAST(100.0 * MAX(vi.Landed) / NULLIF(MAX(vi.BestEquivalent), 0) AS DECIMAL(9,2)),
        SharedParts = MAX(vi.SharedParts),
        AsOfDate = @AsOf
    FROM C c
    -- LEFT, not INNER. Inner-joining to PartVendor made the vendor's own line
    -- count and spend depend on a HAVING clause about ACCEPTANCE: a vendor
    -- whose every delivery was rejected, or whose orders are all still in
    -- transit, would have disappeared from the scorecard entirely rather than
    -- appearing with a null cost index. It drops nothing on this dataset --
    -- 6,301 lines either way -- which is exactly why it would not have been
    -- noticed here.
    LEFT JOIN VendorIndex vi ON vi.VendorKey = c.VendorKey
    GROUP BY c.VendorKey
);
GO

CREATE VIEW dbo.vw_VendorScorecard AS
SELECT s.*, v.VendorID, v.VendorName, v.Country, v.VendorTier, v.LumenRevenueSharePct
FROM dbo.fn_VendorScorecard('2025-12-31') s
JOIN dbo.Dim_Vendor v ON v.VendorKey = s.VendorKey;
GO

/*
--------------------------------------------------------------------------------
fn_SpendKPI(@AsOf) -- the scorecard, one row.
--------------------------------------------------------------------------------
*/
CREATE FUNCTION dbo.fn_SpendKPI (@AsOf DATE)
RETURNS TABLE
AS RETURN
(
    WITH C AS (SELECT * FROM dbo.fn_POLineCost(@AsOf)
               WHERE OrderDate > DATEADD(MONTH, -12, @AsOf)),
    E AS (SELECT * FROM dbo.fn_PriceErosion(@AsOf, 10, 500))
    SELECT
        AsOfDate = @AsOf,
        Lines        = (SELECT COUNT(*) FROM C),
        TotalSpend   = CAST((SELECT SUM(ExtendedPrice) FROM C) AS DECIMAL(16,2)),
        LandedSpend  = CAST((SELECT SUM(LandedCost) FROM C) AS DECIMAL(16,2)),
        PPVAmount    = CAST((SELECT SUM(PPVAmount) FROM C) AS DECIMAL(16,2)),
        PPVPct       = CAST(100.0 * (SELECT SUM(PPVAmount) FROM C)
                          / NULLIF((SELECT SUM(ContractedExtended) FROM C), 0) AS DECIMAL(9,3)),
        MaverickSpendPct = CAST(100.0 * (SELECT SUM(CASE WHEN IsOnContract = 0 THEN ExtendedPrice ELSE 0 END) FROM C)
                          / NULLIF((SELECT SUM(ExtendedPrice) FROM C), 0) AS DECIMAL(9,2)),
        AcceptanceRatePct = CAST(100.0 * (SELECT SUM(ISNULL(QtyAccepted,0)) FROM C)
                          / NULLIF((SELECT SUM(ISNULL(QtyReceived,0)) FROM C), 0) AS DECIMAL(9,2)),
        RejectedValue = CAST((SELECT SUM(RejectedValue) FROM C) AS DECIMAL(16,2)),
        OnTimeDeliveryPct = CAST(100.0 * (SELECT SUM(CASE WHEN IsOnTime = 1 THEN 1 ELSE 0 END) FROM C)
                          / NULLIF((SELECT SUM(CASE WHEN IsOnTime IS NULL THEN 0 ELSE 1 END) FROM C), 0) AS DECIMAL(9,2)),
        ExpediteSpendPct = CAST(100.0 * (SELECT SUM(ExpediteFee) FROM C)
                          / NULLIF((SELECT SUM(ExtendedPrice) FROM C), 0) AS DECIMAL(9,2)),
        -- spend-weighted capture, not a simple average: a pair with $4m of
        -- spend and one with $40k should not count equally
        ErosionCapturePct = CAST((SELECT SUM(ErosionCapturePct * TTMSpend) / NULLIF(SUM(TTMSpend), 0)
                                  FROM E WHERE ErosionCapturePct IS NOT NULL) AS DECIMAL(9,2)),
        ErosionOpportunity = CAST((SELECT SUM(CASE WHEN AnnualOpportunity > 0 THEN AnnualOpportunity ELSE 0 END)
                                   FROM E) AS DECIMAL(16,2)),
        SingleSourceSpendPct = CAST(100.0 *
             (SELECT SUM(c2.ExtendedPrice) FROM C c2 JOIN dbo.Dim_Part p2 ON p2.PartKey = c2.PartKey
              WHERE p2.QualifiedSupplierCount = 1)
             / NULLIF((SELECT SUM(ExtendedPrice) FROM C), 0) AS DECIMAL(9,2)),
        Top5VendorSharePct = CAST(100.0 *
             (SELECT SUM(t.S) FROM (SELECT TOP 5 SUM(c3.ExtendedPrice) AS S FROM C c3
                                    GROUP BY c3.VendorKey ORDER BY SUM(c3.ExtendedPrice) DESC) t)
             / NULLIF((SELECT SUM(ExtendedPrice) FROM C), 0) AS DECIMAL(9,2))
);
GO

CREATE VIEW dbo.vw_SpendKPIMonthly AS
SELECT k.*
FROM (SELECT DISTINCT MonthEndDate FROM dbo.Dim_Date
      WHERE MonthEndDate BETWEEN '2023-12-31' AND '2025-12-31') m
CROSS APPLY dbo.fn_SpendKPI(m.MonthEndDate) k;
GO

/*
--------------------------------------------------------------------------------
fn_RenegotiationQueue(@AsOf) -- THE OPERATIONAL CONTROL.

The unit of work is the VENDOR-PART pair, because that is what gets
renegotiated: a category manager opens a conversation about a specific
component with a specific supplier, not about a vendor in the abstract.

RANKING is by annual opportunity, but ACTION is decided by leverage, and the
two are deliberately separate. The largest opportunity in this book sits on
sole-source parts where there is no credible threat to make; ranking on money
alone would send a category manager into a meeting they cannot win.

LEVERAGE SCORE (0-100) combines the only four things that actually decide a
photonics negotiation:
    - how many qualified suppliers exist (the only real alternative)
    - how long requalification takes (whether the alternative is reachable)
    - how much of the supplier's revenue we represent (whether they care)
    - whether there is a contract at all (whether there is anything to reopen)

BOUNDED BY CAPACITY. A category manager can run perhaps eight serious
negotiations a quarter. A list of four hundred is not a plan, so the queue
stays complete for audit and IsThisQuarter marks what is actually workable.
--------------------------------------------------------------------------------
*/
CREATE FUNCTION dbo.fn_RenegotiationQueue (@AsOf DATE, @NegotiationsPerBuyerPerQuarter INT)
RETURNS TABLE
AS RETURN
(
    WITH E AS (SELECT * FROM dbo.fn_PriceErosion(@AsOf, 10, 500)),
    Quality AS (
        SELECT c.PartKey, c.VendorKey,
               Accepted = SUM(ISNULL(c.QtyAccepted,0)),
               Received = SUM(ISNULL(c.QtyReceived,0)),
               RejectedValue = SUM(c.RejectedValue),
               PPV = SUM(c.PPVAmount),
               OffContractSpend = SUM(CASE WHEN c.IsOnContract = 0 THEN c.ExtendedPrice ELSE 0 END),
               -- The buyer who actually placed the most spend on this pair,
               -- not MIN(key): most pairs are touched by more than one buyer,
               -- and handing the negotiation to whoever sorts first is an
               -- accident, not an assignment.
               --
               -- The duplicate exclusion below is NOT optional. The outer CTE
               -- scopes itself with `PONumber NOT LIKE 'PO-D%'`; this
               -- correlated subquery opens fn_POLineCost again and originally
               -- applied only the date predicate, so every other column on the
               -- row was computed on de-duplicated spend while the buyer was
               -- chosen on spend inflated by duplicated requisitions.
               --
               -- It changed a real assignment: on LOM-0043 / VEN-029 the
               -- duplicate line was the entire margin between two buyers, and
               -- the negotiation was handed to BUY-03 when BUY-04 had actually
               -- placed the most spend. One pair in 185 -- which is why it
               -- survived: a filter dropped in a nested scope is invisible
               -- unless you diff the two populations.
               PrimaryBuyer = (SELECT TOP 1 c2.BuyerKey
                               FROM dbo.fn_POLineCost(@AsOf) c2
                               WHERE c2.PartKey = c.PartKey AND c2.VendorKey = c.VendorKey
                                 AND c2.PONumber NOT LIKE 'PO-D%'
                                 AND c2.OrderDate > DATEADD(MONTH, -12, @AsOf)
                               GROUP BY c2.BuyerKey
                               ORDER BY SUM(c2.ExtendedPrice) DESC, c2.BuyerKey)
        FROM dbo.fn_POLineCost(@AsOf) c
        WHERE c.PONumber NOT LIKE 'PO-D%' AND c.OrderDate > DATEADD(MONTH, -12, @AsOf)
        GROUP BY c.PartKey, c.VendorKey
    ),
    Scored AS (
        SELECT
            e.PartKey, e.VendorKey, p.PartNumber, p.PartName, p.Category, p.Criticality,
            p.QualifiedSupplierCount, p.QualificationMonths,
            v.VendorID, v.VendorName, v.VendorTier, v.LumenRevenueSharePct,
            b.BuyerID, b.BuyerName, b.Team,
            e.FirstPrice, e.LastPrice, e.ExpectedPrice, e.BenchmarkErosionPct,
            e.ActualErosionPct, e.ErosionCapturePct, e.ErosionGapPerUnit,
            e.TTMQty, e.TTMSpend,
            AnnualOpportunity = CASE WHEN e.AnnualOpportunity > 0 THEN e.AnnualOpportunity ELSE 0 END,
            q.RejectedValue, q.PPV, q.OffContractSpend,
            AcceptanceRatePct = CAST(100.0 * q.Accepted / NULLIF(q.Received, 0) AS DECIMAL(6,2)),
            -- leverage, 0-100. Each term is independently defensible and the
            -- weights are stated here rather than buried in a spreadsheet.
            LeverageScore = CAST(
                  CASE WHEN p.QualifiedSupplierCount >= 3 THEN 40
                       WHEN p.QualifiedSupplierCount = 2 THEN 26 ELSE 4 END          -- alternatives exist
                + CASE WHEN p.QualificationMonths <= 2 THEN 20
                       WHEN p.QualificationMonths <= 5 THEN 13
                       WHEN p.QualificationMonths <= 8 THEN 6 ELSE 0 END             -- reachable in time
                + CASE WHEN v.LumenRevenueSharePct >= 15 THEN 30
                       WHEN v.LumenRevenueSharePct >= 8  THEN 20
                       WHEN v.LumenRevenueSharePct >= 4  THEN 11 ELSE 3 END          -- do they care
                + CASE WHEN q.OffContractSpend > 0 THEN 10 ELSE 6 END                -- something to reopen
                AS DECIMAL(5,1))
        FROM E e
        JOIN dbo.Dim_Part   p ON p.PartKey   = e.PartKey
        JOIN dbo.Dim_Vendor v ON v.VendorKey = e.VendorKey
        JOIN Quality q ON q.PartKey = e.PartKey AND q.VendorKey = e.VendorKey
        JOIN dbo.Dim_Buyer  b ON b.BuyerKey = q.PrimaryBuyer
        WHERE e.TTMSpend > 0
    ),
    Ranked AS (
        SELECT s.*,
            PriorityRank = ROW_NUMBER() OVER (ORDER BY s.AnnualOpportunity DESC, s.TTMSpend DESC, s.PartNumber),
            BuyerRank    = ROW_NUMBER() OVER (PARTITION BY s.BuyerID
                                              ORDER BY s.AnnualOpportunity DESC, s.TTMSpend DESC, s.PartNumber),
            ActionCode = CASE
                -- quality first: if the money is being lost to rejects, a price
                -- conversation is the wrong conversation
                WHEN s.AcceptanceRatePct < 95.0
                     AND s.RejectedValue > s.AnnualOpportunity              THEN 'FIX_QUALITY'
                -- spend with no contract at all: there is nothing to renegotiate
                -- until there is something to renegotiate
                WHEN s.OffContractSpend > s.TTMSpend * 0.5                  THEN 'PUT_ON_CONTRACT'
                -- real leverage and real money
                WHEN s.LeverageScore >= 55 AND s.AnnualOpportunity >= 20000 THEN 'RENEGOTIATE'
                -- sole source, big money, and a second source is reachable
                WHEN s.QualifiedSupplierCount = 1 AND s.AnnualOpportunity >= 40000
                     AND s.QualificationMonths <= 8                         THEN 'DUAL_SOURCE'
                -- sole source, big money, and qualification is a project in its
                -- own right: start it now or accept the price for another year
                WHEN s.QualifiedSupplierCount = 1 AND s.AnnualOpportunity >= 40000 THEN 'QUALIFY_ALTERNATE'
                WHEN s.LeverageScore >= 40 AND s.AnnualOpportunity >= 8000   THEN 'RENEGOTIATE'
                -- Above a material threshold, ask anyway. Weak leverage is a
                -- reason to expect a hard conversation, not a reason to skip
                -- one: a six-figure annual gap justifies the meeting even when
                -- the supplier can afford to say no.
                WHEN s.AnnualOpportunity >= 100000                            THEN 'RENEGOTIATE'
                -- no leverage and not enough money to buy any: saying so is more
                -- useful than listing it forever
                ELSE                                                             'ACCEPT' END
        FROM Scored s
    )
    SELECT
        r.PriorityRank, r.BuyerRank,
        r.PartNumber, r.PartName, r.Category, r.Criticality,
        r.QualifiedSupplierCount, r.QualificationMonths,
        IsSingleSource = CAST(CASE WHEN r.QualifiedSupplierCount = 1 THEN 1 ELSE 0 END AS BIT),
        r.VendorID, r.VendorName, r.VendorTier, r.LumenRevenueSharePct,
        r.BuyerID, r.BuyerName, r.Team,
        r.FirstPrice, r.LastPrice, r.ExpectedPrice,
        r.BenchmarkErosionPct, r.ActualErosionPct, r.ErosionCapturePct,
        r.ErosionGapPerUnit, r.TTMQty, r.TTMSpend,
        r.AnnualOpportunity, r.RejectedValue, r.PPV, r.OffContractSpend,
        r.AcceptanceRatePct, r.LeverageScore, r.ActionCode,
        IsThisQuarter = CAST(CASE
            WHEN r.ActionCode = 'ACCEPT' THEN 0
            WHEN r.BuyerRank <= @NegotiationsPerBuyerPerQuarter THEN 1 ELSE 0 END AS BIT),
        RecommendedAction = CASE r.ActionCode
            WHEN 'FIX_QUALITY'      THEN 'Quality first: the loss here is rejected material, not price. Open a corrective action with the supplier before discussing cost.'
            WHEN 'PUT_ON_CONTRACT'  THEN 'Most of this spend has no agreement in force. Put it under contract before attempting to negotiate a price against nothing.'
            WHEN 'RENEGOTIATE'      THEN 'Renegotiate: a credible alternative exists and the supplier cares about our volume. Open with the category erosion benchmark.'
            WHEN 'DUAL_SOURCE'      THEN 'Sole source with real money at stake and a reachable requalification. Qualify a second supplier, then reopen the price.'
            WHEN 'QUALIFY_ALTERNATE'THEN 'Sole source with a long requalification. Start qualifying an alternate now: this is a project, not a negotiation, and the price will not move until it lands.'
            ELSE                         'Accept for now: no credible leverage and not enough value to create any. Revisit if volume or the supplier base changes.'
        END,
        AsOfDate = @AsOf
    FROM Ranked r
);
GO

CREATE VIEW dbo.vw_RenegotiationQueue AS
SELECT * FROM dbo.fn_RenegotiationQueue('2025-12-31', 8);
GO

DECLARE @missing VARCHAR(400) = '';
IF OBJECT_ID('dbo.fn_PriceErosion', 'IF')      IS NULL SET @missing += 'fn_PriceErosion ';
IF OBJECT_ID('dbo.vw_PriceErosion', 'V')       IS NULL SET @missing += 'vw_PriceErosion ';
IF OBJECT_ID('dbo.fn_VendorScorecard', 'IF')   IS NULL SET @missing += 'fn_VendorScorecard ';
IF OBJECT_ID('dbo.vw_VendorScorecard', 'V')    IS NULL SET @missing += 'vw_VendorScorecard ';
IF OBJECT_ID('dbo.fn_SpendKPI', 'IF')          IS NULL SET @missing += 'fn_SpendKPI ';
IF OBJECT_ID('dbo.vw_SpendKPIMonthly', 'V')    IS NULL SET @missing += 'vw_SpendKPIMonthly ';
IF OBJECT_ID('dbo.fn_RenegotiationQueue','IF') IS NULL SET @missing += 'fn_RenegotiationQueue ';
IF OBJECT_ID('dbo.vw_RenegotiationQueue', 'V') IS NULL SET @missing += 'vw_RenegotiationQueue ';
IF @missing <> '' THROW 51005, 'FAILED to create KPI objects -- scroll up for the compile error.', 1;
PRINT 'KPI objects created and verified.';
GO
