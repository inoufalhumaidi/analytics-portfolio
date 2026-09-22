/*
================================================================================
Project 2 -- Vantage Wholesale Supply: Receivables Performance
Script:  07_stored_procedures.sql
Purpose: The reusable interface. Everything a person, a workbook or a refresh
         job asks of this database goes through one of these procedures, so
         that no consumer has to know how a balance is settled or a bucket is
         derived.

DESIGN RULES APPLIED HERE
  1. Every procedure takes @AsOf and defaults it to the project reporting date.
     A report that cannot be re-run for last month cannot be audited.
  2. No dynamic SQL. @GroupBy picks a column through a CASE expression rather
     than being concatenated into a string, so there is no injection surface
     and the plan stays cacheable.
  3. SET NOCOUNT ON everywhere: "(2468 rows affected)" arriving ahead of a
     result set breaks Power Query and ADO.NET consumers in ways that are
     tedious to diagnose.
  4. The scorecard returns LONG format -- one row per metric, carrying its own
     target, threshold and RAG status. Wide format would force the Excel layer
     to re-implement the thresholds, and a threshold implemented twice is a
     threshold that will eventually disagree with itself.

DATA DISCLOSURE: Vantage Wholesale Supply is fictional; all data is synthetic.
================================================================================
*/

USE VantageAR;
GO

IF OBJECT_ID('dbo.usp_ARScorecard', 'P')      IS NOT NULL DROP PROCEDURE dbo.usp_ARScorecard;
IF OBJECT_ID('dbo.usp_PriorityQueue', 'P')    IS NOT NULL DROP PROCEDURE dbo.usp_PriorityQueue;
IF OBJECT_ID('dbo.usp_AgingSummary', 'P')     IS NOT NULL DROP PROCEDURE dbo.usp_AgingSummary;
IF OBJECT_ID('dbo.usp_CustomerStatement', 'P') IS NOT NULL DROP PROCEDURE dbo.usp_CustomerStatement;
IF OBJECT_ID('dbo.usp_DSOBridgeReport', 'P')  IS NOT NULL DROP PROCEDURE dbo.usp_DSOBridgeReport;
IF OBJECT_ID('dbo.usp_CashApplicationWorklist','P') IS NOT NULL DROP PROCEDURE dbo.usp_CashApplicationWorklist;
IF OBJECT_ID('dbo.usp_PromiseReport', 'P')    IS NOT NULL DROP PROCEDURE dbo.usp_PromiseReport;
IF OBJECT_ID('dbo.usp_BillingLagReport', 'P') IS NOT NULL DROP PROCEDURE dbo.usp_BillingLagReport;
GO

/*
--------------------------------------------------------------------------------
usp_ARScorecard -- the month-end scorecard, one row per metric, RAG-graded
against Ref_ARTargets.

The RAG rule is stated once, here, and read by every consumer:
    LowerBetter   Green  value <= Target
                  Amber  value <= Warning
                  Red    otherwise
    HigherBetter  Green  value >= Target
                  Amber  value >= Warning
                  Red    otherwise

CEI is graded on the BOOK variant because that is what industry benchmarks and
lenders quote. CEI_Cash is returned ungraded alongside it, because the honest
figure and the comparable figure are not the same figure and pretending
otherwise is how a scorecard starts lying.
--------------------------------------------------------------------------------
*/
CREATE PROCEDURE dbo.usp_ARScorecard
    @AsOf DATE = '2025-12-31'
AS
BEGIN
    SET NOCOUNT ON;

    DECLARE @k INT = YEAR(@AsOf) * 10000 + MONTH(@AsOf) * 100 + DAY(@AsOf);

    -- portfolio credit utilisation: open AR against the total limit extended
    DECLARE @Utilisation DECIMAL(10,2) = (
        SELECT CAST(100.0 * SUM(ca.OpenBalance) / NULLIF(SUM(c.CreditLimit), 0) AS DECIMAL(10,2))
        FROM dbo.fn_CustomerAR(@AsOf) ca
        JOIN dbo.Dim_Customer c ON c.CustomerKey = ca.CustomerKey
    );

    -- unapplied cash as a share of the open book
    DECLARE @UnappliedPct DECIMAL(10,2) = (
        SELECT CAST(100.0 * SUM(u.UnappliedCash)
             / NULLIF((SELECT SUM(OpenBalance) FROM dbo.fn_ARBalance(@AsOf)), 0) AS DECIMAL(10,2))
        FROM dbo.fn_UnappliedCash(@AsOf) u
    );

    -- dollar-weighted days between despatch and invoice, trailing twelve months
    DECLARE @LagDays DECIMAL(10,2) = (
        SELECT CAST(BillingLagDays AS DECIMAL(10,2))
        FROM dbo.fn_BillingLag(DATEADD(MONTH, -12, @AsOf), @AsOf)
    );

    -- promised dollars actually received, trailing twelve months
    DECLARE @KeptRate DECIMAL(10,2) = (
        SELECT CAST(KeptRatePct AS DECIMAL(10,2)) FROM dbo.fn_PromiseKeptRate(@AsOf, 12)
    );

    -- average age of disputes still open at @AsOf
    DECLARE @DisputeAge DECIMAL(10,2) = (
        SELECT CAST(AVG(CAST(DATEDIFF(DAY, d.[Date], @AsOf) AS FLOAT)) AS DECIMAL(10,2))
        FROM dbo.Fact_Invoice f
        JOIN dbo.Dim_Date d ON d.DateKey = f.DisputeOpenedKey
        WHERE f.IsDisputed = 1
          AND f.DisputeOpenedKey <= @k
          AND (f.DisputeClosedKey IS NULL OR f.DisputeClosedKey > @k)
    );

    ;WITH K AS (SELECT * FROM dbo.fn_ARKPI(@AsOf)),
    Measured AS (
        SELECT MetricName = 'DSO',               MetricValue = CAST(DSO_Countback AS DECIMAL(10,2)),
               Commentary = 'Countback method. Classic method reads ' + CAST(DSO_Simple AS VARCHAR(20)) + ' days.' FROM K
        UNION ALL SELECT 'AvgDaysDelinquent',    CAST(AvgDaysDelinquent AS DECIMAL(10,2)),
               'DSO less Best Possible DSO: the days attributable to late payment alone.' FROM K
        UNION ALL SELECT 'CEI',                  CAST(CEI_Book AS DECIMAL(10,2)),
               'Book variant, for benchmark comparability. Cash variant reads '
               + CAST(CEI_Cash AS VARCHAR(20)) + '%, a gap of ' + CAST(PaperCollectionsGap AS VARCHAR(20)) + ' points.' FROM K
        UNION ALL SELECT 'PctPastDue',           CAST(PctPastDue AS DECIMAL(10,2)),
               'By value. By invoice count it is ' + CAST(PctPastDueByCount AS VARCHAR(20)) + '%.' FROM K
        UNION ALL SELECT 'Pct90Plus',            CAST(Pct90Plus AS DECIMAL(10,2)),
               'By value. By invoice count it is ' + CAST(Pct90PlusByCount AS VARCHAR(20))
               + '%; a large gap means the bucket is full of small items, not distressed debt.' FROM K
        UNION ALL SELECT 'CreditUtilization',    @Utilisation,
               'Total open AR against total credit extended.'
        UNION ALL SELECT 'DisputeAgeDays',       @DisputeAge,
               'Mean age of disputes still open at the reporting date.'
        UNION ALL SELECT 'UnappliedCashPct',      @UnappliedPct,
               'Cash banked against customer accounts that nobody has matched to an invoice. Every dollar of it makes an invoice look unpaid when it is not.'
        UNION ALL SELECT 'BillingLagDays',        @LagDays,
               'Despatch to invoice, dollar-weighted. The only part of the cash cycle Vantage can shorten without a customer conversation.'
        UNION ALL SELECT 'PromiseKeptRate',       @KeptRate,
               'Dollar-weighted. Read it by risk tier before drawing any conclusion: a healthy portfolio figure can sit on top of one cohort that has stopped paying.'
    )
    SELECT
        AsOfDate = @AsOf,
        m.MetricName,
        m.MetricValue,
        t.TargetValue,
        t.WarningValue,
        t.Direction,
        t.Unit,
        RAGStatus = CASE
            WHEN m.MetricValue IS NULL THEN 'No Data'
            WHEN t.Direction = 'LowerBetter' THEN
                 CASE WHEN m.MetricValue <= t.TargetValue  THEN 'Green'
                      WHEN m.MetricValue <= t.WarningValue THEN 'Amber'
                      ELSE 'Red' END
            ELSE CASE WHEN m.MetricValue >= t.TargetValue  THEN 'Green'
                      WHEN m.MetricValue >= t.WarningValue THEN 'Amber'
                      ELSE 'Red' END
        END,
        VarianceToTarget = CAST(CASE WHEN t.Direction = 'LowerBetter'
                                     THEN m.MetricValue - t.TargetValue
                                     ELSE t.TargetValue - m.MetricValue END AS DECIMAL(10,2)),
        t.Description,
        m.Commentary
    FROM Measured m
    JOIN dbo.Ref_ARTargets t ON t.MetricName = m.MetricName
    ORDER BY CASE m.MetricName WHEN 'DSO' THEN 1 WHEN 'AvgDaysDelinquent' THEN 2
                               WHEN 'CEI' THEN 3 WHEN 'PctPastDue' THEN 4
                               WHEN 'Pct90Plus' THEN 5 WHEN 'CreditUtilization' THEN 6
                               WHEN 'UnappliedCashPct' THEN 7 WHEN 'BillingLagDays' THEN 8
                               WHEN 'PromiseKeptRate' THEN 9 ELSE 10 END;
END;
GO

/*
--------------------------------------------------------------------------------
usp_PriorityQueue -- the worklist a collector actually opens.

@CollectorID NULL returns every collector, which is the manager's view.
@TodaysWorklistOnly = 1 applies the capacity rule; 0 returns the full ranked
ledger for audit and for anyone who wants to see what was left out.
--------------------------------------------------------------------------------
*/
CREATE PROCEDURE dbo.usp_PriorityQueue
    @AsOf DATE = '2025-12-31',
    @CollectorID VARCHAR(10) = NULL,
    @TodaysWorklistOnly BIT = 1,
    @ActionCode VARCHAR(20) = NULL
AS
BEGIN
    SET NOCOUNT ON;

    -- This is the procedure a collector actually opens, and it was the one
    -- with no argument validation: a typo in either filter returned zero rows
    -- and exit code 0, which on a worklist reads as "nothing to call today".
    -- UAT-14's own rationale names that harm -- "a procedure that silently
    -- falls back to a default answers a question nobody asked, and the caller
    -- has no way to tell" -- and three of the eight procedures implemented it.
    IF @CollectorID IS NOT NULL
       AND NOT EXISTS (SELECT 1 FROM dbo.Dim_Collector WHERE CollectorID = @CollectorID)
        THROW 50010, 'No such CollectorID. An empty worklist must not be the answer to a typo.', 1;

    IF @ActionCode IS NOT NULL
       AND NOT EXISTS (SELECT 1 FROM dbo.vw_PriorityActionQueue WHERE ActionCode = @ActionCode)
        THROW 50011, 'No such ActionCode in the current queue. Valid codes are visible in dbo.vw_PriorityActionQueue.', 1;

    SELECT
        q.PriorityRank, q.CollectorRank,
        q.CustomerID, q.CustomerName, q.Segment, q.Region, q.RiskTier,
        q.CollectorID, q.CollectorName, q.Team, q.TermsCode,
        q.OpenBalance, q.PastDueBalance, q.Balance90Plus, q.DisputedBalance,
        q.WeightedExposure, q.CollectableExposure,
        q.OldestDaysPastDue, q.CreditUtilizationPct, q.AvgDaysLateHistoric,
        q.UnappliedCash, q.NetExposure, q.BrokenPromises90d,
        q.ActionCode, q.CreditHoldFlag, q.IsTodaysWorklist,
        q.RecommendedAction,
        -- The number the collector opens the call with. Net of the disputed
        -- balance (a different team's job) and of cash already banked against
        -- the account (ours to apply, not theirs to pay again).
        DollarAsk = CAST(CASE WHEN q.PastDueBalance - q.DisputedBalance - q.UnappliedCash > 0
                              THEN q.PastDueBalance - q.DisputedBalance - q.UnappliedCash
                              ELSE 0 END AS DECIMAL(14,2))
    FROM dbo.fn_PriorityActionQueue(@AsOf) q
    WHERE (@CollectorID IS NULL OR q.CollectorID = @CollectorID)
      AND (@TodaysWorklistOnly = 0 OR q.IsTodaysWorklist = 1)
      AND (@ActionCode IS NULL OR q.ActionCode = @ActionCode)
    ORDER BY q.CollectorName, q.CollectorRank;
END;
GO

/*
--------------------------------------------------------------------------------
usp_AgingSummary -- open AR by ageing bucket, cut by one chosen attribute.

@GroupBy is resolved through a CASE expression rather than concatenated into
dynamic SQL: no injection surface, and one cached plan instead of one per
argument value. An unrecognised @GroupBy raises rather than silently falling
back to a total, because a summary that quietly answers a different question
than the one asked is worse than an error.
--------------------------------------------------------------------------------
*/
CREATE PROCEDURE dbo.usp_AgingSummary
    @AsOf DATE = '2025-12-31',
    @GroupBy VARCHAR(20) = 'Segment'   -- Segment | Region | RiskTier | Terms | Collector | Total
AS
BEGIN
    SET NOCOUNT ON;

    IF @GroupBy NOT IN ('Segment','Region','RiskTier','Terms','Collector','Total')
        THROW 50010, 'usp_AgingSummary: @GroupBy must be Segment, Region, RiskTier, Terms, Collector or Total.', 1;

    SELECT
        AsOfDate = @AsOf,
        GroupedBy = @GroupBy,
        GroupValue = CASE @GroupBy
                        WHEN 'Segment'   THEN c.Segment
                        WHEN 'Region'    THEN c.Region
                        WHEN 'RiskTier'  THEN c.RiskTier
                        WHEN 'Terms'     THEN t.TermsCode
                        WHEN 'Collector' THEN col.CollectorName
                        ELSE 'All accounts' END,
        b.BucketKey,
        b.BucketName,
        OpenBalance  = CAST(SUM(b.OpenBalance) AS DECIMAL(14,2)),
        InvoiceCount = COUNT(*),
        AvgInvoiceBalance = CAST(AVG(b.OpenBalance) AS DECIMAL(14,2)),
        DisputedBalance = CAST(SUM(CASE WHEN b.IsDisputedOpen = 1 THEN b.OpenBalance ELSE 0 END) AS DECIMAL(14,2)),
        WeightedExposure = CAST(SUM(b.OpenBalance * ISNULL(b.RiskWeight, 0)) AS DECIMAL(14,2))
    FROM dbo.fn_ARBalance(@AsOf) b
    JOIN dbo.Dim_Customer c     ON c.CustomerKey = b.CustomerKey
    JOIN dbo.Dim_Collector col  ON col.CollectorKey = c.CollectorKey
    JOIN dbo.Dim_PaymentTerms t ON t.TermsKey = b.TermsKey
    WHERE b.IsOpen = 1
    GROUP BY
        CASE @GroupBy
            WHEN 'Segment'   THEN c.Segment
            WHEN 'Region'    THEN c.Region
            WHEN 'RiskTier'  THEN c.RiskTier
            WHEN 'Terms'     THEN t.TermsCode
            WHEN 'Collector' THEN col.CollectorName
            ELSE 'All accounts' END,
        b.BucketKey, b.BucketName
    ORDER BY GroupValue, b.BucketKey;
END;
GO

/*
--------------------------------------------------------------------------------
usp_CustomerStatement -- everything a collector needs in front of them on the
call: the open items, what has been paid against each, and what is disputed.

Settled invoices are excluded by default but @IncludeSettled = 1 brings them
back, because "we paid that in October" is the most common thing a customer
says and the collector needs to be able to answer it in the same screen.
--------------------------------------------------------------------------------
*/
CREATE PROCEDURE dbo.usp_CustomerStatement
    @CustomerID VARCHAR(10),
    @AsOf DATE = '2025-12-31',
    @IncludeSettled BIT = 0
AS
BEGIN
    SET NOCOUNT ON;

    IF NOT EXISTS (SELECT 1 FROM dbo.Dim_Customer WHERE CustomerID = @CustomerID)
        THROW 50011, 'usp_CustomerStatement: unknown @CustomerID.', 1;

    SELECT
        AsOfDate = @AsOf,
        c.CustomerID, c.CustomerName, c.Segment, c.Region, c.RiskTier, c.CreditLimit,
        col.CollectorName, t.TermsCode, t.NetDays,
        b.InvoiceNo, b.InvoiceDate, b.DueDate, b.InvoiceAmount,
        b.PaidToDate, b.DiscountToDate, b.CreditMemoToDate, b.WriteOffToDate,
        b.OpenBalance, b.OverAppliedAmount,
        b.DaysPastDue, b.BucketName,
        b.IsDisputedOpen, b.DisputeReason,
        Status = CASE WHEN b.IsOpen = 0 THEN 'Settled'
                      WHEN b.IsDisputedOpen = 1 THEN 'Open - disputed'
                      WHEN b.DaysPastDue > 0 THEN 'Open - past due'
                      ELSE 'Open - within terms' END
    FROM dbo.fn_ARBalance(@AsOf) b
    JOIN dbo.Dim_Customer c     ON c.CustomerKey = b.CustomerKey
    JOIN dbo.Dim_Collector col  ON col.CollectorKey = c.CollectorKey
    JOIN dbo.Dim_PaymentTerms t ON t.TermsKey = b.TermsKey
    WHERE c.CustomerID = @CustomerID
      AND (@IncludeSettled = 1 OR b.IsOpen = 1)
    ORDER BY b.DaysPastDue DESC, b.InvoiceDate;
END;
GO

/*
--------------------------------------------------------------------------------
usp_DSOBridgeReport -- the granted-versus-taken answer, in the shape a slide
wants: one row per bridge component, with the cash each component represents.

The components sum to classic DSO exactly; the procedure returns the total as
its own row so the reader can verify the addition without trusting the label.
--------------------------------------------------------------------------------
*/
CREATE PROCEDURE dbo.usp_DSOBridgeReport
    @AsOf DATE = '2025-12-31'
AS
BEGIN
    SET NOCOUNT ON;

    ;WITH B AS (SELECT * FROM dbo.fn_DSOBridge(@AsOf))
    SELECT AsOfDate = @AsOf, Seq = 1, Component = 'Granted -- not yet due',
           Days = GrantedDays, Cash = CurrentAR,
           Owner = 'Sales / terms policy',
           Note = 'Terms Vantage sold on purpose. Weighted average terms sold: '
                  + CAST(WeightedAvgTermsDays AS VARCHAR(20)) + ' days.' FROM B
    UNION ALL
    SELECT @AsOf, 2, 'Taken -- disputed and past due', DisputeDays, DisputedPastDueAR,
           'Billing / sales', 'Late, but blocked behind a commercial argument. A collections call cannot clear this.' FROM B
    UNION ALL
    SELECT @AsOf, 3, 'Taken -- past due, undisputed', LatenessDays, UndisputedPastDueAR,
           'Collections', 'Late with nothing in the way. This is the only component a call list can move.' FROM B
    UNION ALL
    SELECT @AsOf, 4, 'TOTAL (classic DSO)', DSO_Classic, TotalAR,
           '', 'Components above sum to this figure exactly, by construction.' FROM B
    ORDER BY Seq;
END;
GO

/*
--------------------------------------------------------------------------------
usp_CashApplicationWorklist -- the other worklist.

Collections gets a call list; cash application gets this. Both are ranked by
money, but this one is ranked by money Vantage ALREADY HAS and cannot see,
which is cheaper to recover than anything a phone call can achieve: no
negotiation, no concession, no relationship cost.

Ordered oldest-first rather than largest-first. A receipt that has sat unmatched
for a year is the one most likely to have already caused a wrong call, and the
one whose paperwork will be hardest to reconstruct later.
--------------------------------------------------------------------------------
*/
CREATE PROCEDURE dbo.usp_CashApplicationWorklist
    @AsOf DATE = '2025-12-31',
    @MinAmount DECIMAL(12,2) = 0
AS
BEGIN
    SET NOCOUNT ON;

    SELECT
        u.CustomerKey, c.CustomerID, c.CustomerName, c.Segment, c.Region,
        col.CollectorName,
        u.UnappliedCash, u.UnappliedReceipts, u.FullyUnappliedReceipts,
        u.NoRemittanceAdvice, u.OldestUnappliedDays,
        OpenBalance = ca.OpenBalance,
        PastDueBalance = ca.PastDueBalance,
        -- what the account owes once the cash we hold is applied against it
        NetExposure = ca.NetExposure,
        -- the reason this matters more than its size: without applying the
        -- cash, this much past-due balance is chaseable that should not be
        WronglyChaseable = CAST(CASE WHEN ca.PastDueBalance < u.UnappliedCash
                                     THEN ca.PastDueBalance ELSE u.UnappliedCash END AS DECIMAL(14,2)),
        RootCause = CASE WHEN u.NoRemittanceAdvice > 0
                         THEN 'No remittance advice received -- request it from the customer, then apply.'
                         ELSE 'Remittance advice was received; the cash simply has not been matched.' END
    FROM dbo.fn_UnappliedCash(@AsOf) u
    JOIN dbo.Dim_Customer c    ON c.CustomerKey = u.CustomerKey
    JOIN dbo.Dim_Collector col ON col.CollectorKey = c.CollectorKey
    JOIN dbo.fn_CustomerAR(@AsOf) ca ON ca.CustomerKey = u.CustomerKey
    WHERE u.UnappliedCash > @MinAmount
    ORDER BY u.OldestUnappliedDays DESC, u.UnappliedCash DESC;
END;
GO

/*
--------------------------------------------------------------------------------
usp_PromiseReport -- promise-to-pay outcomes, always cut by cohort.

@GroupBy defaults to RiskTier for a reason. The portfolio kept-rate on this book
reads comfortably above target while one tier of it has stopped paying, and a
single headline figure would hide exactly the thing the metric exists to find.

Status is computed from cash by dbo.fn_PromiseStatus; nothing here reads a
stored outcome, because the register is maintained by the people being measured.
--------------------------------------------------------------------------------
*/
CREATE PROCEDURE dbo.usp_PromiseReport
    @AsOf DATE = '2025-12-31',
    @Months INT = 12,
    @GroupBy VARCHAR(20) = 'RiskTier'   -- RiskTier | Collector | Segment | Region | Total
AS
BEGIN
    SET NOCOUNT ON;

    IF @GroupBy NOT IN ('RiskTier','Collector','Segment','Region','Total')
        THROW 50012, 'usp_PromiseReport: @GroupBy must be RiskTier, Collector, Segment, Region or Total.', 1;

    SELECT
        AsOfDate = @AsOf,
        GroupedBy = @GroupBy,
        GroupValue = CASE @GroupBy
                        WHEN 'RiskTier'  THEN c.RiskTier
                        WHEN 'Collector' THEN col.CollectorName
                        WHEN 'Segment'   THEN c.Segment
                        WHEN 'Region'    THEN c.Region
                        ELSE 'All accounts' END,
        PromisesDue     = COUNT(*),
        PromisedDollars = CAST(SUM(s.PromisedAmount) AS DECIMAL(14,2)),
        KeptCount       = SUM(CASE WHEN s.PromiseStatus = 'Kept' THEN 1 ELSE 0 END),
        PartialCount    = SUM(CASE WHEN s.PromiseStatus = 'PartiallyKept' THEN 1 ELSE 0 END),
        BrokenCount     = SUM(CASE WHEN s.PromiseStatus = 'Broken' THEN 1 ELSE 0 END),
        BrokenDollars   = CAST(SUM(CASE WHEN s.PromiseStatus = 'Broken' THEN s.PromisedAmount ELSE 0 END) AS DECIMAL(14,2)),
        KeptRatePct     = CAST(100.0 * SUM(CASE WHEN s.PromiseStatus = 'Kept' THEN s.PromisedAmount ELSE 0 END)
                             / NULLIF(SUM(s.PromisedAmount), 0) AS DECIMAL(6,2)),
        KeptRateByCountPct = CAST(100.0 * SUM(CASE WHEN s.PromiseStatus = 'Kept' THEN 1 ELSE 0 END)
                             / NULLIF(COUNT(*), 0) AS DECIMAL(6,2))
    FROM dbo.fn_PromiseStatus(@AsOf, 3, 90.00) s
    JOIN dbo.Dim_Customer c    ON c.CustomerKey = s.CustomerKey
    JOIN dbo.Dim_Collector col ON col.CollectorKey = s.CollectorKey
    WHERE s.PromiseStatus <> 'Outstanding'      -- unripe promises have not failed
      AND s.PromisedPayDate > DATEADD(MONTH, -@Months, @AsOf)
    GROUP BY
        CASE @GroupBy
            WHEN 'RiskTier'  THEN c.RiskTier
            WHEN 'Collector' THEN col.CollectorName
            WHEN 'Segment'   THEN c.Segment
            WHEN 'Region'    THEN c.Region
            ELSE 'All accounts' END
    ORDER BY KeptRatePct;
END;
GO

/*
--------------------------------------------------------------------------------
usp_BillingLagReport -- despatch-to-invoice days by region and month, with the
cash each day of lag represents.

Reported by region because a company-wide average conceals the only thing worth
knowing: whether the lag is everywhere (a process) or somewhere (a site).
--------------------------------------------------------------------------------
*/
CREATE PROCEDURE dbo.usp_BillingLagReport
    @AsOf DATE = '2025-12-31',
    @Months INT = 12,
    @BaselineLagDays DECIMAL(6,2) = 1.50    -- what a well-run site achieves here
AS
BEGIN
    SET NOCOUNT ON;

    DECLARE @From DATE = DATEADD(MONTH, -@Months, @AsOf);
    DECLARE @CashPerDay DECIMAL(14,2) = (SELECT CashPerLagDay FROM dbo.fn_DSOBridge(@AsOf));

    SELECT
        AsOfDate = @AsOf,
        l.Region,
        InvoiceCount = SUM(l.InvoiceCount),
        CreditSales  = CAST(SUM(l.CreditSales) AS DECIMAL(14,2)),
        BillingLagDays = CAST(SUM(l.CreditSales * l.BillingLagDays) / NULLIF(SUM(l.CreditSales), 0) AS DECIMAL(6,2)),
        ExcessOverBaseline = CAST(SUM(l.CreditSales * l.BillingLagDays) / NULLIF(SUM(l.CreditSales), 0)
                                - @BaselineLagDays AS DECIMAL(6,2)),
        -- the region's own sales carry its own lag, so the cash at stake is
        -- scaled by that region's share rather than by company-wide sales
        CashTiedUp = CAST(
            (SUM(l.CreditSales * l.BillingLagDays) / NULLIF(SUM(l.CreditSales), 0) - @BaselineLagDays)
            * (SUM(l.CreditSales) / NULLIF(DATEDIFF(DAY, @From, @AsOf), 0)) AS DECIMAL(14,2)),
        Verdict = CASE
            WHEN SUM(l.CreditSales * l.BillingLagDays) / NULLIF(SUM(l.CreditSales), 0) > @BaselineLagDays + 1
                 THEN 'Investigate: invoice posting is running late at this site.'
            ELSE 'Within normal range.' END
    FROM dbo.vw_BillingLagMonthly l
    WHERE l.MonthEndDate > @From AND l.MonthEndDate <= @AsOf
    GROUP BY l.Region
    ORDER BY BillingLagDays DESC;
END;
GO

DECLARE @missing VARCHAR(400) = '';
IF OBJECT_ID('dbo.usp_ARScorecard', 'P')       IS NULL SET @missing += 'usp_ARScorecard ';
IF OBJECT_ID('dbo.usp_PriorityQueue', 'P')     IS NULL SET @missing += 'usp_PriorityQueue ';
IF OBJECT_ID('dbo.usp_AgingSummary', 'P')      IS NULL SET @missing += 'usp_AgingSummary ';
IF OBJECT_ID('dbo.usp_CustomerStatement', 'P') IS NULL SET @missing += 'usp_CustomerStatement ';
IF OBJECT_ID('dbo.usp_DSOBridgeReport', 'P')   IS NULL SET @missing += 'usp_DSOBridgeReport ';
IF OBJECT_ID('dbo.usp_CashApplicationWorklist','P') IS NULL SET @missing += 'usp_CashApplicationWorklist ';
IF OBJECT_ID('dbo.usp_PromiseReport', 'P')     IS NULL SET @missing += 'usp_PromiseReport ';
IF OBJECT_ID('dbo.usp_BillingLagReport', 'P')  IS NULL SET @missing += 'usp_BillingLagReport ';
IF @missing <> '' THROW 50005, 'FAILED to create stored procedures -- scroll up for the compile error.', 1;
PRINT 'Stored procedures created and verified.';
GO
