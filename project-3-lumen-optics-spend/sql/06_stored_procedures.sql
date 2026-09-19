/*
================================================================================
Project 3 -- Lumen Optics Manufacturing: Photonics Spend Scorecard
Script:  06_stored_procedures.sql
Purpose: The reusable interface. Everything a person, a workbook, a Power BI
         refresh or a scheduled job asks of this database goes through one of
         these, so no consumer needs to know how landed cost is built or how
         the erosion benchmark is applied.

DESIGN RULES APPLIED HERE
  1. Every procedure takes @AsOf and defaults it to the project reporting date.
     A report that cannot be re-run for last quarter cannot be audited.
  2. No dynamic SQL. @GroupBy resolves through a CASE expression rather than
     being concatenated into a string: no injection surface, one cached plan.
  3. SET NOCOUNT ON everywhere. "(6327 rows affected)" arriving ahead of a
     result set breaks Power Query and ADO.NET consumers in tedious ways.
  4. The scorecard returns LONG format -- one row per metric carrying its own
     target, threshold and RAG status -- so the Excel and DAX layers read the
     thresholds rather than re-implementing them. A threshold implemented twice
     is a threshold that will eventually disagree with itself.

DATA DISCLOSURE: Lumen Optics Manufacturing is fictional; all data is synthetic.
================================================================================
*/

USE LumenSpend;
GO

IF OBJECT_ID('dbo.usp_SpendScorecard', 'P')      IS NOT NULL DROP PROCEDURE dbo.usp_SpendScorecard;
IF OBJECT_ID('dbo.usp_RenegotiationQueue', 'P')  IS NOT NULL DROP PROCEDURE dbo.usp_RenegotiationQueue;
IF OBJECT_ID('dbo.usp_VendorScorecard', 'P')     IS NOT NULL DROP PROCEDURE dbo.usp_VendorScorecard;
IF OBJECT_ID('dbo.usp_PriceErosionDetail', 'P')  IS NOT NULL DROP PROCEDURE dbo.usp_PriceErosionDetail;
IF OBJECT_ID('dbo.usp_CategorySummary', 'P')     IS NOT NULL DROP PROCEDURE dbo.usp_CategorySummary;
IF OBJECT_ID('dbo.usp_PartPriceHistory', 'P')    IS NOT NULL DROP PROCEDURE dbo.usp_PartPriceHistory;
GO

/*
--------------------------------------------------------------------------------
usp_SpendScorecard -- one row per metric, RAG-graded against Ref_SpendTargets.

The RAG rule is stated once, here, and read by every consumer:
    LowerBetter   Green value <= Target · Amber value <= Warning · Red otherwise
    HigherBetter  Green value >= Target · Amber value >= Warning · Red otherwise

Each row carries a Commentary that names the cohort underneath it where one
exists, because on this book several metrics are green in aggregate and failing
inside. A scorecard that reports only the aggregate is telling the truth and
still misleading.
--------------------------------------------------------------------------------
*/
CREATE PROCEDURE dbo.usp_SpendScorecard
    @AsOf DATE = '2025-12-31'
AS
BEGIN
    SET NOCOUNT ON;

    -- The worst vendor-CATEGORY pair, not the worst vendor. A supplier can sit
    -- at 98% overall and 91% on one category, and the vendor-level figure hides
    -- exactly the thing worth acting on -- the same aggregate-hides-cohort trap
    -- this whole scorecard exists to expose.
    DECLARE @WorstAcceptVendor VARCHAR(60), @WorstAcceptPct DECIMAL(6,2);
    SELECT TOP 1
        @WorstAcceptVendor = v.VendorID + ' on ' + p.Category,
        @WorstAcceptPct    = CAST(100.0 * SUM(ISNULL(c.QtyAccepted,0))
                                 / NULLIF(SUM(ISNULL(c.QtyReceived,0)),0) AS DECIMAL(6,2))
    FROM dbo.fn_POLineCost(@AsOf) c
    JOIN dbo.Dim_Vendor v ON v.VendorKey = c.VendorKey
    JOIN dbo.Dim_Part   p ON p.PartKey   = c.PartKey
    GROUP BY v.VendorID, p.Category
    HAVING SUM(ISNULL(c.QtyReceived,0)) > 2000
    ORDER BY 100.0 * SUM(ISNULL(c.QtyAccepted,0)) / NULLIF(SUM(ISNULL(c.QtyReceived,0)),0);

    DECLARE @WorstOnTimeVendor VARCHAR(10), @WorstOnTimePct DECIMAL(6,2);
    SELECT TOP 1 @WorstOnTimeVendor = v.VendorID, @WorstOnTimePct = s.OnTimePct
    FROM dbo.fn_VendorScorecard(@AsOf) s
    JOIN dbo.Dim_Vendor v ON v.VendorKey = s.VendorKey
    WHERE s.Lines > 50
    ORDER BY s.OnTimePct;

    ;WITH K AS (SELECT * FROM dbo.fn_SpendKPI(@AsOf)),
    Measured AS (
        SELECT MetricName = 'ErosionCapturePct', MetricValue = CAST(ErosionCapturePct AS DECIMAL(10,2)),
               Commentary = 'Share of the expected category erosion actually realised. The gap is worth '
                          + FORMAT(ErosionOpportunity, 'C0', 'en-US') + ' a year and no price variance report can see it.' FROM K
        UNION ALL SELECT 'MaverickSpendPct', CAST(MaverickSpendPct AS DECIMAL(10,2)),
               'Spend placed with no agreement in force at the order date. There is nothing to negotiate against until it is on contract.' FROM K
        UNION ALL SELECT 'PPVPct', CAST(PPVPct AS DECIMAL(10,2)),
               'Measured against the agreement in force. A supplier who renews flat scores zero here forever -- read it beside erosion capture, never alone.' FROM K
        UNION ALL SELECT 'AcceptanceRatePct', CAST(AcceptanceRatePct AS DECIMAL(10,2)),
               'Portfolio figure, and it hides the cohort: worst vendor-category with material volume is '
               + ISNULL(@WorstAcceptVendor,'n/a') + ' at ' + ISNULL(CAST(@WorstAcceptPct AS VARCHAR(10)),'n/a') + '%.' FROM K
        UNION ALL SELECT 'OnTimeDeliveryPct', CAST(OnTimeDeliveryPct AS DECIMAL(10,2)),
               'Portfolio figure. Worst vendor with material volume is ' + ISNULL(@WorstOnTimeVendor,'n/a')
               + ' at ' + ISNULL(CAST(@WorstOnTimePct AS VARCHAR(10)),'n/a') + '%.' FROM K
        UNION ALL SELECT 'ExpediteSpendPct', CAST(ExpediteSpendPct AS DECIMAL(10,2)),
               'Expedite fees as a share of spend. Caused by planning, not by the supplier, and negotiable only by fixing the plan.' FROM K
        UNION ALL SELECT 'SingleSourceSpendPct', CAST(SingleSourceSpendPct AS DECIMAL(10,2)),
               'Spend on parts with one qualified supplier. This is where the largest erosion gaps sit and the least leverage exists.' FROM K
        UNION ALL SELECT 'Top5VendorSharePct', CAST(Top5VendorSharePct AS DECIMAL(10,2)),
               'Concentration of spend in the five largest suppliers.' FROM K
    )
    SELECT
        AsOfDate = @AsOf,
        m.MetricName, m.MetricValue, t.TargetValue, t.WarningValue, t.Direction, t.Unit,
        RAGStatus = CASE
            WHEN m.MetricValue IS NULL THEN 'No Data'
            WHEN t.Direction = 'LowerBetter' THEN
                 CASE WHEN m.MetricValue <= t.TargetValue THEN 'Green'
                      WHEN m.MetricValue <= t.WarningValue THEN 'Amber' ELSE 'Red' END
            ELSE CASE WHEN m.MetricValue >= t.TargetValue THEN 'Green'
                      WHEN m.MetricValue >= t.WarningValue THEN 'Amber' ELSE 'Red' END END,
        VarianceToTarget = CAST(CASE WHEN t.Direction = 'LowerBetter'
                                     THEN m.MetricValue - t.TargetValue
                                     ELSE t.TargetValue - m.MetricValue END AS DECIMAL(10,2)),
        t.Description, m.Commentary
    FROM Measured m
    JOIN dbo.Ref_SpendTargets t ON t.MetricName = m.MetricName
    ORDER BY CASE m.MetricName
        WHEN 'ErosionCapturePct' THEN 1 WHEN 'MaverickSpendPct' THEN 2 WHEN 'PPVPct' THEN 3
        WHEN 'AcceptanceRatePct' THEN 4 WHEN 'OnTimeDeliveryPct' THEN 5 WHEN 'ExpediteSpendPct' THEN 6
        WHEN 'SingleSourceSpendPct' THEN 7 ELSE 8 END;
END;
GO

/*
--------------------------------------------------------------------------------
usp_RenegotiationQueue -- the worklist a category manager actually opens.

@BuyerID NULL returns every buyer, which is the sourcing director's view.
@ThisQuarterOnly = 1 applies the capacity rule; 0 returns the full ranked list
for audit and for anyone who wants to see what was left out.
--------------------------------------------------------------------------------
*/
CREATE PROCEDURE dbo.usp_RenegotiationQueue
    @AsOf DATE = '2025-12-31',
    @BuyerID VARCHAR(10) = NULL,
    @ThisQuarterOnly BIT = 1,
    @ActionCode VARCHAR(20) = NULL,
    @NegotiationsPerBuyer INT = 8
AS
BEGIN
    SET NOCOUNT ON;

    SELECT q.*
    FROM dbo.fn_RenegotiationQueue(@AsOf, @NegotiationsPerBuyer) q
    WHERE (@BuyerID IS NULL OR q.BuyerID = @BuyerID)
      AND (@ThisQuarterOnly = 0 OR q.IsThisQuarter = 1)
      AND (@ActionCode IS NULL OR q.ActionCode = @ActionCode)
    ORDER BY q.BuyerName, q.BuyerRank;
END;
GO

/*
--------------------------------------------------------------------------------
usp_VendorScorecard -- price, quality, delivery and compliance per supplier.

@MinSpend filters out the long tail, because a supplier with two orders has a
delivery percentage that is arithmetic rather than evidence.
--------------------------------------------------------------------------------
*/
CREATE PROCEDURE dbo.usp_VendorScorecard
    @AsOf DATE = '2025-12-31',
    @MinSpend DECIMAL(16,2) = 0
AS
BEGIN
    SET NOCOUNT ON;

    SELECT
        s.VendorKey, v.VendorID, v.VendorName, v.Country, v.VendorTier,
        v.LumenRevenueSharePct,
        s.Lines, s.TotalSpend, s.LandedSpend, s.FreightSpend, s.ExpediteSpend,
        s.ExpeditePctOfSpend, s.OnContractPct, s.PPVAmount, s.PPVPct,
        s.QtyReceived, s.QtyAccepted, s.AcceptanceRatePct, s.RejectedValue,
        s.OnTimePct, s.AvgDaysLate,
        s.CostIndexVsBest, s.SharedParts,
        -- a single figure a sourcing director can sort on: 100 is the best
        -- available position on every dimension, and each term is visible above
        CompositeScore = CAST(
              0.40 * (200 - s.CostIndexVsBest)/100.0 * 50          -- cost position
            + 0.30 * ISNULL(s.AcceptanceRatePct, 0)                -- quality
            + 0.20 * ISNULL(s.OnTimePct, 0)                        -- delivery
            + 0.10 * ISNULL(s.OnContractPct, 0)                    -- compliance
            AS DECIMAL(6,2)),
        s.AsOfDate
    FROM dbo.fn_VendorScorecard(@AsOf) s
    JOIN dbo.Dim_Vendor v ON v.VendorKey = s.VendorKey
    WHERE s.TotalSpend >= @MinSpend
    ORDER BY s.TotalSpend DESC;
END;
GO

/*
--------------------------------------------------------------------------------
usp_PriceErosionDetail -- the evidence behind a single negotiation.

This is what a category manager takes into the meeting: what the first price
was, what the curve says it should be now, what is actually being paid, and how
much a year the difference is worth.
--------------------------------------------------------------------------------
*/
CREATE PROCEDURE dbo.usp_PriceErosionDetail
    @AsOf DATE = '2025-12-31',
    @Category VARCHAR(30) = NULL,
    @VendorID VARCHAR(10) = NULL,
    @MinOpportunity DECIMAL(16,2) = 0
AS
BEGIN
    SET NOCOUNT ON;

    SELECT
        e.PartNumber, e.PartName, e.Category, e.Criticality,
        e.VendorID, e.VendorName, e.VendorTier, e.LumenRevenueSharePct,
        e.QualifiedSupplierCount, e.QualificationMonths,
        e.FirstDate, e.LastDate, e.YearsElapsed, e.Lines,
        e.FirstPrice, e.LastPrice, e.ExpectedPrice,
        e.BenchmarkErosionPct, e.ActualErosionPct, e.ErosionCapturePct,
        e.ErosionGapPerUnit, e.TTMQty, e.TTMSpend, e.AnnualOpportunity,
        Verdict = CASE
            WHEN e.ErosionCapturePct IS NULL         THEN 'No benchmark for this category'
            WHEN e.ErosionCapturePct >= 100          THEN 'Beating the curve -- protect this relationship'
            WHEN e.ErosionCapturePct >= 60           THEN 'Tracking the curve'
            WHEN e.ErosionCapturePct >= 20           THEN 'Partial: roughly half the expected erosion captured'
            WHEN e.ActualErosionPct  <  0            THEN 'Price has RISEN while the category declined'
            ELSE                                          'Flat: the curve moved and the price did not' END
    FROM dbo.vw_PriceErosion e
    WHERE (@Category IS NULL OR e.Category = @Category)
      AND (@VendorID IS NULL OR e.VendorID = @VendorID)
      AND e.AnnualOpportunity >= @MinOpportunity
    ORDER BY e.AnnualOpportunity DESC;
END;
GO

/*
--------------------------------------------------------------------------------
usp_CategorySummary -- spend and erosion cut by one chosen attribute.

@GroupBy resolves through CASE rather than dynamic SQL. An unrecognised value
raises rather than silently falling back to a total, because a summary that
quietly answers a different question than the one asked is worse than an error.
--------------------------------------------------------------------------------
*/
CREATE PROCEDURE dbo.usp_CategorySummary
    @AsOf DATE = '2025-12-31',
    @GroupBy VARCHAR(20) = 'Category'  -- Category | Vendor | VendorTier | Buyer | Team | Criticality | Country
AS
BEGIN
    SET NOCOUNT ON;

    IF @GroupBy NOT IN ('Category','Vendor','VendorTier','Buyer','Team','Criticality','Country')
        THROW 51010, 'usp_CategorySummary: @GroupBy must be Category, Vendor, VendorTier, Buyer, Team, Criticality or Country.', 1;

    ;WITH C AS (
        SELECT c.*, p.Category, p.Criticality, p.QualifiedSupplierCount,
               v.VendorID, v.VendorName, v.VendorTier, v.Country,
               b.BuyerName, b.Team
        FROM dbo.fn_POLineCost(@AsOf) c
        JOIN dbo.Dim_Part   p ON p.PartKey   = c.PartKey
        JOIN dbo.Dim_Vendor v ON v.VendorKey = c.VendorKey
        JOIN dbo.Dim_Buyer  b ON b.BuyerKey  = c.BuyerKey
        WHERE c.PONumber NOT LIKE 'PO-D%' AND c.OrderDate > DATEADD(MONTH, -12, @AsOf)
    )
    SELECT
        AsOfDate = @AsOf, GroupedBy = @GroupBy,
        GroupValue = CASE @GroupBy
            WHEN 'Category'    THEN C.Category   WHEN 'Vendor'      THEN C.VendorName
            WHEN 'VendorTier'  THEN C.VendorTier WHEN 'Buyer'       THEN C.BuyerName
            WHEN 'Team'        THEN C.Team       WHEN 'Criticality' THEN C.Criticality
            ELSE C.Country END,
        Lines        = COUNT(*),
        TotalSpend   = CAST(SUM(C.ExtendedPrice) AS DECIMAL(16,2)),
        LandedSpend  = CAST(SUM(C.LandedCost) AS DECIMAL(16,2)),
        SpendSharePct= CAST(100.0 * SUM(C.ExtendedPrice) / SUM(SUM(C.ExtendedPrice)) OVER () AS DECIMAL(6,2)),
        MaverickPct  = CAST(100.0 * SUM(CASE WHEN C.IsOnContract = 0 THEN C.ExtendedPrice ELSE 0 END)
                          / NULLIF(SUM(C.ExtendedPrice), 0) AS DECIMAL(6,2)),
        AcceptancePct= CAST(100.0 * SUM(ISNULL(C.QtyAccepted,0)) / NULLIF(SUM(ISNULL(C.QtyReceived,0)),0) AS DECIMAL(6,2)),
        OnTimePct    = CAST(100.0 * SUM(CASE WHEN C.IsOnTime = 1 THEN 1 ELSE 0 END)
                          / NULLIF(SUM(CASE WHEN C.IsOnTime IS NULL THEN 0 ELSE 1 END), 0) AS DECIMAL(6,2)),
        ExpeditePct  = CAST(100.0 * SUM(C.ExpediteFee) / NULLIF(SUM(C.ExtendedPrice), 0) AS DECIMAL(6,2)),
        SingleSourcePct = CAST(100.0 * SUM(CASE WHEN C.QualifiedSupplierCount = 1 THEN C.ExtendedPrice ELSE 0 END)
                          / NULLIF(SUM(C.ExtendedPrice), 0) AS DECIMAL(6,2))
    FROM C
    GROUP BY CASE @GroupBy
            WHEN 'Category'    THEN C.Category   WHEN 'Vendor'      THEN C.VendorName
            WHEN 'VendorTier'  THEN C.VendorTier WHEN 'Buyer'       THEN C.BuyerName
            WHEN 'Team'        THEN C.Team       WHEN 'Criticality' THEN C.Criticality
            ELSE C.Country END
    ORDER BY TotalSpend DESC;
END;
GO

/*
--------------------------------------------------------------------------------
usp_PartPriceHistory -- every price paid for one part, by supplier, over time.

The answer to "are you sure?" in a renegotiation. Shows the expected curve
alongside the paid prices so the gap is visible transaction by transaction
rather than asserted as a total.
--------------------------------------------------------------------------------
*/
CREATE PROCEDURE dbo.usp_PartPriceHistory
    @PartNumber VARCHAR(20),
    @AsOf DATE = '2025-12-31'
AS
BEGIN
    SET NOCOUNT ON;

    IF NOT EXISTS (SELECT 1 FROM dbo.Dim_Part WHERE PartNumber = @PartNumber)
        THROW 51011, 'usp_PartPriceHistory: unknown @PartNumber.', 1;

    DECLARE @PartKey INT = (SELECT PartKey FROM dbo.Dim_Part WHERE PartNumber = @PartNumber);
    DECLARE @FirstBuy DATE = (SELECT FirstPurchaseDate FROM dbo.Dim_Part WHERE PartNumber = @PartNumber);
    DECLARE @Erosion DECIMAL(5,2) = (SELECT b.AnnualErosionPct FROM dbo.Dim_Part p
                                     JOIN dbo.Ref_PriceErosionBenchmark b ON b.Category = p.Category
                                     WHERE p.PartNumber = @PartNumber);
    DECLARE @BasePrice DECIMAL(12,4) = (
        SELECT TOP 1 c.UnitPrice FROM dbo.fn_POLineCost(@AsOf) c
        WHERE c.PartKey = @PartKey AND c.PONumber NOT LIKE 'PO-D%' ORDER BY c.OrderDate, c.POLineKey);

    SELECT
        c.OrderDate, c.PONumber, v.VendorID, v.VendorName,
        c.OrderQty, c.UnitPrice,
        c.ContractedUnitPrice, c.IsOnContract, c.AgreementsInForce,
        ExpectedPriceOnCurve = CAST(@BasePrice * POWER(1.0 - @Erosion/100.0,
                                    DATEDIFF(DAY, @FirstBuy, c.OrderDate)/365.25) AS DECIMAL(12,4)),
        GapToCurve = CAST(c.UnitPrice - @BasePrice * POWER(1.0 - @Erosion/100.0,
                                    DATEDIFF(DAY, @FirstBuy, c.OrderDate)/365.25) AS DECIMAL(12,4)),
        c.QtyReceived, c.QtyAccepted, c.AcceptanceRatePct, c.CostPerAcceptedUnit,
        c.DaysLate, c.IsOnTime
    FROM dbo.fn_POLineCost(@AsOf) c
    JOIN dbo.Dim_Vendor v ON v.VendorKey = c.VendorKey
    WHERE c.PartKey = @PartKey AND c.PONumber NOT LIKE 'PO-D%'
    ORDER BY c.OrderDate, v.VendorID;
END;
GO

DECLARE @missing VARCHAR(400) = '';
IF OBJECT_ID('dbo.usp_SpendScorecard', 'P')     IS NULL SET @missing += 'usp_SpendScorecard ';
IF OBJECT_ID('dbo.usp_RenegotiationQueue', 'P') IS NULL SET @missing += 'usp_RenegotiationQueue ';
IF OBJECT_ID('dbo.usp_VendorScorecard', 'P')    IS NULL SET @missing += 'usp_VendorScorecard ';
IF OBJECT_ID('dbo.usp_PriceErosionDetail', 'P') IS NULL SET @missing += 'usp_PriceErosionDetail ';
IF OBJECT_ID('dbo.usp_CategorySummary', 'P')    IS NULL SET @missing += 'usp_CategorySummary ';
IF OBJECT_ID('dbo.usp_PartPriceHistory', 'P')   IS NULL SET @missing += 'usp_PartPriceHistory ';
IF @missing <> '' THROW 51006, 'FAILED to create stored procedures -- scroll up for the compile error.', 1;
PRINT 'Stored procedures created and verified.';
GO
