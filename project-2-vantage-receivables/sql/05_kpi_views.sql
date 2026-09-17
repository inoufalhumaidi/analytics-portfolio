/*
================================================================================
Project 2 -- Vantage Wholesale Supply: Receivables Performance
Script:  05_kpi_views.sql
Purpose: The receivables KPI layer. Everything here reads its balances from
         dbo.fn_ARBalance so there is exactly one definition of what is owed.

Business question:
   "Which accounts and terms are driving DSO up, and who should collections
    call first?"

WHY TWO DSOs
    Simple DSO      = open AR / credit sales over a fixed window x days.
    Countback DSO   = walk back month by month, consuming AR against each
                      month's sales until it is exhausted; add up the days.
    They disagree whenever sales are seasonal, which they are here (open AR
    peaks in September and falls by December). Reporting only one hides that,
    so both are published and the gap is itself a reported figure. The target
    in Ref_ARTargets is set against the countback method.

WHY ADD USES THE SAME METHOD ON BOTH SIDES
    Average Days Delinquent = DSO - Best Possible DSO. BPDSO is the DSO the
    company would post if every customer paid exactly on terms, so it is
    computed by countback over the NOT-YET-DUE balance only. Mixing a
    countback DSO with a simple-formula BPDSO would make ADD measure the
    change of method rather than customer behaviour.

DATA DISCLOSURE: Vantage Wholesale Supply is fictional; all data is synthetic.
================================================================================
*/

USE VantageAR;
GO

IF OBJECT_ID('dbo.vw_PriorityActionQueue', 'V') IS NOT NULL DROP VIEW dbo.vw_PriorityActionQueue;
IF OBJECT_ID('dbo.vw_PromiseStatus', 'V')       IS NOT NULL DROP VIEW dbo.vw_PromiseStatus;
IF OBJECT_ID('dbo.fn_PromiseKeptRate', 'IF')    IS NOT NULL DROP FUNCTION dbo.fn_PromiseKeptRate;
IF OBJECT_ID('dbo.fn_PromiseStatus', 'IF')      IS NOT NULL DROP FUNCTION dbo.fn_PromiseStatus;
IF OBJECT_ID('dbo.vw_CustomerAR', 'V')          IS NOT NULL DROP VIEW dbo.vw_CustomerAR;
IF OBJECT_ID('dbo.vw_ARKPIMonthly', 'V')        IS NOT NULL DROP VIEW dbo.vw_ARKPIMonthly;
IF OBJECT_ID('dbo.fn_ARKPI', 'IF')              IS NOT NULL DROP FUNCTION dbo.fn_ARKPI;
IF OBJECT_ID('dbo.fn_DSO', 'IF')                IS NOT NULL DROP FUNCTION dbo.fn_DSO;
IF OBJECT_ID('dbo.fn_PriorityActionQueue', 'IF') IS NOT NULL DROP FUNCTION dbo.fn_PriorityActionQueue;
IF OBJECT_ID('dbo.fn_CustomerAR', 'IF')          IS NOT NULL DROP FUNCTION dbo.fn_CustomerAR;
IF OBJECT_ID('dbo.vw_MonthlyCreditSales', 'V')  IS NOT NULL DROP VIEW dbo.vw_MonthlyCreditSales;
GO

/*
--------------------------------------------------------------------------------
Credit sales by month -- the denominator of every DSO variant.
Sales are taken at invoice face value: credit memos are a later correction and
netting them into the month of sale would flatter the ratio.
--------------------------------------------------------------------------------
*/
CREATE VIEW dbo.vw_MonthlyCreditSales AS
SELECT
    d.YearMonth,
    MonthEndDate = MAX(d.MonthEndDate),
    DaysInMonth  = DAY(MAX(d.MonthEndDate)),
    CreditSales  = SUM(f.InvoiceAmount),
    InvoiceCount = COUNT(*)
FROM dbo.Fact_Invoice f
JOIN dbo.Dim_Date d ON d.DateKey = f.InvoiceDateKey
GROUP BY d.YearMonth;
GO

/*
--------------------------------------------------------------------------------
fn_DSO(@AsOf) -- countback and simple DSO, plus Best Possible DSO and ADD.

Countback, set-based: order months backwards from @AsOf, accumulate sales, and
find the first month where cumulative sales cover the balance. Days = all whole
months before it, plus the fraction of that month needed to finish the balance.

If AR exceeds every sale in the look-back window, DSO is capped at the window
and IsCapped is raised rather than returning a quietly wrong number.
--------------------------------------------------------------------------------
*/
CREATE FUNCTION dbo.fn_DSO (@AsOf DATE)
RETURNS TABLE
AS RETURN
(
    WITH Bal AS (
        SELECT TotalAR   = SUM(OpenBalance),
               CurrentAR = SUM(CASE WHEN DaysPastDue <= 0 THEN OpenBalance ELSE 0 END)
        FROM dbo.fn_ARBalance(@AsOf)
    ),
    Months AS (              -- most recent month first
        SELECT s.YearMonth, s.CreditSales, s.DaysInMonth,
               Ord = ROW_NUMBER() OVER (ORDER BY s.MonthEndDate DESC),
               RunSales = SUM(s.CreditSales) OVER (ORDER BY s.MonthEndDate DESC ROWS UNBOUNDED PRECEDING),
               RunDays  = SUM(s.DaysInMonth) OVER (ORDER BY s.MonthEndDate DESC ROWS UNBOUNDED PRECEDING)
        FROM dbo.vw_MonthlyCreditSales s
        WHERE s.MonthEndDate <= EOMONTH(@AsOf)
    ),
    Window12 AS (            -- trailing 12 months, the simple-DSO denominator
        SELECT Sales12 = SUM(CreditSales), Days12 = SUM(DaysInMonth)
        FROM Months WHERE Ord <= 12
    ),
    -- the month in which the balance runs out, for each of the two balances
    CountbackTotal AS (
        SELECT TOP 1 m.Ord, m.CreditSales, m.DaysInMonth, m.RunSales, m.RunDays,
               PriorSales = m.RunSales - m.CreditSales,
               PriorDays  = m.RunDays  - m.DaysInMonth
        FROM Months m CROSS JOIN Bal b
        WHERE m.RunSales >= b.TotalAR
        ORDER BY m.Ord
    ),
    CountbackCurrent AS (
        SELECT TOP 1 m.Ord, m.CreditSales, m.DaysInMonth, m.RunSales, m.RunDays,
               PriorSales = m.RunSales - m.CreditSales,
               PriorDays  = m.RunDays  - m.DaysInMonth
        FROM Months m CROSS JOIN Bal b
        WHERE m.RunSales >= b.CurrentAR
        ORDER BY m.Ord
    ),
    Horizon AS (SELECT MaxDays = SUM(DaysInMonth), MaxSales = SUM(CreditSales) FROM Months)
    SELECT
        AsOfDate  = @AsOf,
        b.TotalAR,
        b.CurrentAR,
        PastDueAR = b.TotalAR - b.CurrentAR,
        w.Sales12,
        -- countback DSO: whole months + the fraction of the month that finishes the balance
        DSO_Countback = CAST(ISNULL(
              ct.PriorDays + (b.TotalAR - ct.PriorSales) / NULLIF(ct.CreditSales, 0) * ct.DaysInMonth
            , h.MaxDays) AS DECIMAL(9,2)),
        BPDSO_Countback = CAST(ISNULL(
              cc.PriorDays + (b.CurrentAR - cc.PriorSales) / NULLIF(cc.CreditSales, 0) * cc.DaysInMonth
            , h.MaxDays) AS DECIMAL(9,2)),
        -- the conventional formula, published alongside so the gap is visible
        DSO_Simple = CAST(b.TotalAR * w.Days12 / NULLIF(w.Sales12, 0) AS DECIMAL(9,2)),   -- multiply first: matches fn_DSOBridge exactly, so the two tie to the cent
        -- Days lost purely to late payment. Derived from the two figures as
        -- PUBLISHED, not from their unrounded intermediates: rounding three
        -- numbers independently to two decimals lets the difference disagree
        -- with the subtraction a reader can do on the page.
        AvgDaysDelinquent =
              CAST(ISNULL(ct.PriorDays + (b.TotalAR - ct.PriorSales) / NULLIF(ct.CreditSales, 0) * ct.DaysInMonth, h.MaxDays) AS DECIMAL(9,2))
            - CAST(ISNULL(cc.PriorDays + (b.CurrentAR - cc.PriorSales) / NULLIF(cc.CreditSales, 0) * cc.DaysInMonth, h.MaxDays) AS DECIMAL(9,2)),
        IsCapped = CAST(CASE WHEN ct.Ord IS NULL THEN 1 ELSE 0 END AS BIT)
    FROM Bal b
    CROSS JOIN Window12 w
    CROSS JOIN Horizon h
    LEFT JOIN CountbackTotal   ct ON 1 = 1
    LEFT JOIN CountbackCurrent cc ON 1 = 1
);
GO

/*
--------------------------------------------------------------------------------
fn_ARKPI(@AsOf) -- the month-end scorecard, one row.

CEI (Collection Effectiveness Index) measures how much of what COULD have been
collected actually was:

    collectable = Beginning AR + Credit Sales - Ending CURRENT AR

The denominator excludes the not-yet-due balance, because nobody is late for
money that was never due.

CEI IS REPORTED TWICE, AND THE GAP IS THE POINT
    CEI_Book = (Beginning AR + Credit Sales - Ending Total AR) / collectable
    CEI_Cash = (cash and settlement discount applied in the month) / collectable

    The textbook figure is CEI_Book, and it measures "AR went down". AR goes
    down identically whether a customer wired the money or a controller wrote
    the balance off, so CEI_Book can be improved by giving up. Vantage wrote off
    $702k over this period, and every dollar of it scores as a collection.

    CEI_Cash counts only money that reached the bank. PaperCollectionsGap is the
    difference, in points, and it is a metric in its own right: it quantifies
    how much of the reported collections performance was paperwork.

Past due is reported in DOLLARS and in INVOICE COUNTS. They diverge, and the
divergence is diagnostic: a 90+ bucket that is alarming by count but modest by
value is clogged with small deductions, which is an operations problem, not a
credit problem, and no amount of collector time will fix it.
--------------------------------------------------------------------------------
*/
CREATE FUNCTION dbo.fn_ARKPI (@AsOf DATE)
RETURNS TABLE
AS RETURN
(
    WITH PriorEnd AS (SELECT d = EOMONTH(@AsOf, -1)),
    Opening AS (
        SELECT BeginAR = SUM(OpenBalance) FROM dbo.fn_ARBalance((SELECT d FROM PriorEnd))
    ),
    Closing AS (
        SELECT
            EndAR        = SUM(OpenBalance),
            EndCurrentAR = SUM(CASE WHEN DaysPastDue <= 0 THEN OpenBalance ELSE 0 END),
            PastDueAR    = SUM(CASE WHEN DaysPastDue >  0 THEN OpenBalance ELSE 0 END),
            AR90Plus     = SUM(CASE WHEN DaysPastDue > 90 THEN OpenBalance ELSE 0 END),
            DisputedAR   = SUM(CASE WHEN IsDisputedOpen = 1 THEN OpenBalance ELSE 0 END),
            OpenInvoices = SUM(CAST(IsOpen AS INT)),
            PastDueInvoices = SUM(CASE WHEN IsOpen = 1 AND DaysPastDue >  0 THEN 1 ELSE 0 END),
            Invoices90Plus  = SUM(CASE WHEN IsOpen = 1 AND DaysPastDue > 90 THEN 1 ELSE 0 END)
        FROM dbo.fn_ARBalance(@AsOf)
    ),
    Sales AS (
        SELECT MonthSales = ISNULL(SUM(CreditSales), 0)
        FROM dbo.vw_MonthlyCreditSales
        WHERE MonthEndDate = EOMONTH(@AsOf)
    ),
    -- Cash APPLIED in the month. CEI measures AR movement, and AR only moves
    -- when cash is matched to an invoice -- so applied, not received. Cash that
    -- arrived and was never matched is reported on its own below, because it is
    -- a different problem with a different owner.
    CashIn AS (
        SELECT CashApplied = ISNULL(SUM(p.AppliedAmount + p.DiscountTaken), 0)
        FROM dbo.Fact_CashApplication p
        JOIN dbo.Dim_Date d ON d.DateKey = p.ApplicationDateKey
        WHERE d.[Date] BETWEEN DATEFROMPARTS(YEAR(@AsOf), MONTH(@AsOf), 1) AND EOMONTH(@AsOf)
    ),
    -- Cash RECEIVED in the month. The gap against CashApplied is how much the
    -- cash application backlog grew or shrank.
    CashBanked AS (
        SELECT CashReceived = ISNULL(SUM(r.ReceiptAmount), 0)
        FROM dbo.Fact_CashReceipt r
        JOIN dbo.Dim_Date d ON d.DateKey = r.ReceiptDateKey
        WHERE d.[Date] BETWEEN DATEFROMPARTS(YEAR(@AsOf), MONTH(@AsOf), 1) AND EOMONTH(@AsOf)
    ),
    Unapplied AS (
        SELECT UnappliedCash = ISNULL(SUM(UnappliedCash), 0) FROM dbo.fn_UnappliedCash(@AsOf)
    ),
    NonCash AS (
        SELECT WriteOffs   = ISNULL(SUM(CASE WHEN j.AdjustmentType = 'WriteOff'   THEN j.Amount END), 0),
               CreditMemos = ISNULL(SUM(CASE WHEN j.AdjustmentType = 'CreditMemo' THEN j.Amount END), 0)
        FROM dbo.Fact_Adjustment j
        JOIN dbo.Dim_Date d ON d.DateKey = j.AdjustmentDateKey
        WHERE d.[Date] BETWEEN DATEFROMPARTS(YEAR(@AsOf), MONTH(@AsOf), 1) AND EOMONTH(@AsOf)
    )
    SELECT
        AsOfDate = @AsOf,
        YearMonth = CONVERT(CHAR(7), @AsOf, 126),
        o.BeginAR, s.MonthSales,
        c.EndAR, c.EndCurrentAR, c.PastDueAR, c.AR90Plus, c.DisputedAR,
        c.OpenInvoices, c.PastDueInvoices, c.Invoices90Plus,
        x.CashApplied, cb.CashReceived, n.WriteOffs, n.CreditMemos,
        CashAppliedLessReceived = CAST(x.CashApplied - cb.CashReceived AS DECIMAL(14,2)),
        ua.UnappliedCash,
        UnappliedCashPct = CAST(100.0 * ua.UnappliedCash / NULLIF(c.EndAR, 0) AS DECIMAL(6,2)),
        d.DSO_Countback, d.DSO_Simple, d.BPDSO_Countback, d.AvgDaysDelinquent,
        DSO_MethodGap = CAST(d.DSO_Countback - d.DSO_Simple AS DECIMAL(9,2)),
        -- the collectable denominator, published so the reader can see the scale
        CollectableBase = CAST(o.BeginAR + s.MonthSales - c.EndCurrentAR AS DECIMAL(14,2)),
        CEI_Book = CAST(100.0 * (o.BeginAR + s.MonthSales - c.EndAR)
                 / NULLIF(o.BeginAR + s.MonthSales - c.EndCurrentAR, 0) AS DECIMAL(6,2)),
        CEI_Cash = CAST(100.0 * x.CashApplied
                 / NULLIF(o.BeginAR + s.MonthSales - c.EndCurrentAR, 0) AS DECIMAL(6,2)),
        PaperCollectionsGap = CAST(
                   100.0 * (o.BeginAR + s.MonthSales - c.EndAR)
                 / NULLIF(o.BeginAR + s.MonthSales - c.EndCurrentAR, 0)
                 - 100.0 * x.CashApplied
                 / NULLIF(o.BeginAR + s.MonthSales - c.EndCurrentAR, 0) AS DECIMAL(6,2)),
        PctPastDue = CAST(100.0 * c.PastDueAR / NULLIF(c.EndAR, 0) AS DECIMAL(6,2)),
        Pct90Plus  = CAST(100.0 * c.AR90Plus  / NULLIF(c.EndAR, 0) AS DECIMAL(6,2)),
        PctDisputed = CAST(100.0 * c.DisputedAR / NULLIF(c.EndAR, 0) AS DECIMAL(6,2)),
        -- the count twins: read against the dollar versions, never instead of them
        PctPastDueByCount = CAST(100.0 * c.PastDueInvoices / NULLIF(c.OpenInvoices, 0) AS DECIMAL(6,2)),
        Pct90PlusByCount  = CAST(100.0 * c.Invoices90Plus  / NULLIF(c.OpenInvoices, 0) AS DECIMAL(6,2))
    FROM Opening o
    CROSS JOIN Closing c
    CROSS JOIN Sales s
    CROSS JOIN CashIn x
    CROSS JOIN CashBanked cb
    CROSS JOIN Unapplied ua
    CROSS JOIN NonCash n
    CROSS APPLY dbo.fn_DSO(@AsOf) d
);
GO

/*
--------------------------------------------------------------------------------
vw_ARKPIMonthly -- the trend. Every month end with a full prior month behind it.

IsComparablePeriod exists because the ledger starts on 2024-01-01, and a new
book cannot show a representative aging profile straight away. The first
invoice falls due on 2024-01-22, so no balance can reach the 90+ bucket before
late April 2024, and an account on NET60 terms needs 60 + 90 = 150 days before
it can appear there at all. Until roughly 2024-06-30 the aging ratios are
therefore structurally flattered, and DSO is computed over fewer months of
sales than the countback wants.

Those months are published rather than hidden -- deleting inconvenient periods
is worse than labelling them -- but every trend comparison, target assessment
and management conclusion in this project uses IsComparablePeriod = 1.
--------------------------------------------------------------------------------
*/
CREATE VIEW dbo.vw_ARKPIMonthly AS
SELECT k.*,
       IsComparablePeriod = CAST(CASE WHEN m.MonthEndDate >= '2024-06-30' THEN 1 ELSE 0 END AS BIT)
FROM (SELECT DISTINCT MonthEndDate FROM dbo.Dim_Date
      WHERE MonthEndDate BETWEEN '2024-02-29' AND '2025-12-31') m
CROSS APPLY dbo.fn_ARKPI(m.MonthEndDate) k;
GO
/*
--------------------------------------------------------------------------------
fn_PromiseStatus(@AsOf, @GraceDays, @KeptThresholdPct)

Whether a promise to pay was kept is DERIVED FROM CASH, never stored. The
register records what a collector was told; the bank records what happened. A
stored status column would be written by the same person the kept-rate measures,
and a collector under pressure logs optimistic outcomes.

BOTH TOLERANCES ARE NECESSARY OR THE METRIC IS MEANINGLESS
    Without a grace window, a wire that lands one day late scores as broken and
    the rate is absurdly low. Without an amount threshold, a $50 payment "keeps"
    a $50,000 promise and the rate is absurdly high. Defaults: three days of
    grace, 90% of the promised amount.

    The window opens a day EARLY as well, because customers who pay by cheque
    frequently send it ahead of the date they committed to.

A promise whose date has not yet arrived is Outstanding, not Broken. Counting
unripe promises as failures would make the rate depend on when you asked.

Cash is matched at ACCOUNT level, not to specific invoices: a promise is made
for an amount, and the customer settles it however their remittance falls.
--------------------------------------------------------------------------------
*/
IF OBJECT_ID('dbo.vw_PromiseStatus', 'V')      IS NOT NULL DROP VIEW dbo.vw_PromiseStatus;
IF OBJECT_ID('dbo.fn_PromiseKeptRate', 'IF')   IS NOT NULL DROP FUNCTION dbo.fn_PromiseKeptRate;
IF OBJECT_ID('dbo.fn_PromiseStatus', 'IF')     IS NOT NULL DROP FUNCTION dbo.fn_PromiseStatus;
GO

CREATE FUNCTION dbo.fn_PromiseStatus
    (@AsOf DATE, @GraceDays INT, @KeptThresholdPct DECIMAL(5,2))
RETURNS TABLE
AS RETURN
(
    SELECT
        p.PromiseKey, p.PromiseNo, p.CustomerKey, p.CollectorKey,
        PromiseMadeDate = dm.[Date],
        PromisedPayDate = dp.[Date],
        p.PromisedAmount,
        CashInWindow = CAST(x.Paid AS DECIMAL(14,2)),
        PctOfPromise = CAST(100.0 * x.Paid / NULLIF(p.PromisedAmount, 0) AS DECIMAL(9,2)),
        DaysSincePromised = DATEDIFF(DAY, dp.[Date], @AsOf),
        PromiseStatus = CASE
            WHEN dp.[Date] > @AsOf                                            THEN 'Outstanding'
            WHEN 100.0 * x.Paid / NULLIF(p.PromisedAmount,0) >= @KeptThresholdPct THEN 'Kept'
            WHEN 100.0 * x.Paid / NULLIF(p.PromisedAmount,0) >= 25.0          THEN 'PartiallyKept'
            ELSE 'Broken' END,
        AsOfDate = @AsOf
    FROM dbo.Fact_PromiseToPay p
    JOIN dbo.Dim_Date dm ON dm.DateKey = p.PromiseMadeKey
    JOIN dbo.Dim_Date dp ON dp.DateKey = p.PromisedPayDateKey
    CROSS APPLY (
        SELECT Paid = ISNULL(SUM(a.AppliedAmount), 0)
        FROM dbo.Fact_CashApplication a
        JOIN dbo.Dim_Date da ON da.DateKey = a.ApplicationDateKey
        WHERE a.CustomerKey = p.CustomerKey
          AND da.[Date] BETWEEN DATEADD(DAY, -1, dp.[Date]) AND DATEADD(DAY, @GraceDays, dp.[Date])
          AND da.[Date] <= @AsOf
    ) x
    WHERE p.PromiseMadeKey <= YEAR(@AsOf)*10000 + MONTH(@AsOf)*100 + DAY(@AsOf)
);
GO

CREATE VIEW dbo.vw_PromiseStatus AS
SELECT s.*, c.CustomerID, c.CustomerName, c.RiskTier, c.Segment, col.CollectorName
FROM dbo.fn_PromiseStatus('2025-12-31', 3, 90.00) s
JOIN dbo.Dim_Customer c    ON c.CustomerKey = s.CustomerKey
JOIN dbo.Dim_Collector col ON col.CollectorKey = s.CollectorKey;
GO

/*
--------------------------------------------------------------------------------
fn_PromiseKeptRate(@AsOf, @Months) -- kept rate over a trailing window.

Reported dollar-weighted AND by count, because they diverge and the divergence
is the finding: a stream of small kept promises can hide one large broken one.

Outstanding promises are excluded from both numerator and denominator. They have
not failed; they have not happened yet.
--------------------------------------------------------------------------------
*/
CREATE FUNCTION dbo.fn_PromiseKeptRate (@AsOf DATE, @Months INT)
RETURNS TABLE
AS RETURN
(
    WITH S AS (
        SELECT * FROM dbo.fn_PromiseStatus(@AsOf, 3, 90.00)
        WHERE PromisedPayDate > DATEADD(MONTH, -@Months, @AsOf)
          AND PromisedPayDate <= @AsOf
    )
    SELECT
        AsOfDate = @AsOf,
        WindowMonths = @Months,
        PromisesDue     = COUNT(*),
        PromisedDollars = CAST(ISNULL(SUM(PromisedAmount), 0) AS DECIMAL(14,2)),
        KeptCount       = SUM(CASE WHEN PromiseStatus = 'Kept' THEN 1 ELSE 0 END),
        PartialCount    = SUM(CASE WHEN PromiseStatus = 'PartiallyKept' THEN 1 ELSE 0 END),
        BrokenCount     = SUM(CASE WHEN PromiseStatus = 'Broken' THEN 1 ELSE 0 END),
        BrokenDollars   = CAST(ISNULL(SUM(CASE WHEN PromiseStatus = 'Broken' THEN PromisedAmount ELSE 0 END), 0) AS DECIMAL(14,2)),
        KeptRatePct     = CAST(100.0 * SUM(CASE WHEN PromiseStatus = 'Kept' THEN PromisedAmount ELSE 0 END)
                             / NULLIF(SUM(PromisedAmount), 0) AS DECIMAL(6,2)),
        KeptRateByCountPct = CAST(100.0 * SUM(CASE WHEN PromiseStatus = 'Kept' THEN 1 ELSE 0 END)
                             / NULLIF(COUNT(*), 0) AS DECIMAL(6,2))
    FROM S
);
GO

/*
--------------------------------------------------------------------------------
vw_CustomerAR -- one row per customer, the account-level picture.

Two exposure figures are published, and the distinction matters:

  WeightedExposure          sum(balance x aging RiskWeight). A dollar 90+ days
                            late counts fully; a dollar still within terms counts
                            nothing, because it is not yet a collections problem.

  DisputedWeightedExposure  the same weighting applied to the disputed portion
                            only. It exists so the queue can discount disputes
                            on the SAME scale. Subtracting raw disputed dollars
                            from a weighted total mixes units and can drive the
                            result below zero -- a within-terms dispute carries
                            weight 0.00 but full face value.
--------------------------------------------------------------------------------
*/
CREATE FUNCTION dbo.fn_CustomerAR (@AsOf DATE)
RETURNS TABLE
AS RETURN
(
WITH B AS (SELECT * FROM dbo.fn_ARBalance(@AsOf)),
Hist AS (      -- how this customer has behaved on everything already settled
    SELECT b.CustomerKey,
           SettledInvoices = COUNT(*),
           AvgDaysLate = AVG(CAST(DATEDIFF(DAY, b.DueDate, lp.LastPayDate) AS FLOAT))
    FROM B b
    CROSS APPLY (SELECT LastPayDate = MAX(d.[Date])
                 FROM dbo.Fact_CashApplication p
                 JOIN dbo.Dim_Date d ON d.DateKey = p.ApplicationDateKey
                 WHERE p.InvoiceKey = b.InvoiceKey) lp
    WHERE b.IsOpen = 0 AND lp.LastPayDate IS NOT NULL
    GROUP BY b.CustomerKey
)
SELECT
    c.CustomerKey, c.CustomerID, c.CustomerName, c.Segment, c.Region, c.SalesRep,
    c.RiskTier, c.CreditLimit,
    col.CollectorID, col.CollectorName, col.Team,
    t.TermsCode, t.NetDays,
    OpenBalance     = CAST(ISNULL(SUM(b.OpenBalance), 0) AS DECIMAL(14,2)),
    PastDueBalance  = CAST(ISNULL(SUM(CASE WHEN b.DaysPastDue > 0  THEN b.OpenBalance ELSE 0 END), 0) AS DECIMAL(14,2)),
    Balance90Plus   = CAST(ISNULL(SUM(CASE WHEN b.DaysPastDue > 90 THEN b.OpenBalance ELSE 0 END), 0) AS DECIMAL(14,2)),
    DisputedBalance = CAST(ISNULL(SUM(CASE WHEN b.IsDisputedOpen = 1 THEN b.OpenBalance ELSE 0 END), 0) AS DECIMAL(14,2)),
    WeightedExposure = CAST(ISNULL(SUM(b.OpenBalance * ISNULL(b.RiskWeight, 0)), 0) AS DECIMAL(14,2)),
    DisputedWeightedExposure = CAST(ISNULL(SUM(CASE WHEN b.IsDisputedOpen = 1
                                   THEN b.OpenBalance * ISNULL(b.RiskWeight, 0) ELSE 0 END), 0) AS DECIMAL(14,2)),
    OpenInvoices    = ISNULL(SUM(CAST(b.IsOpen AS INT)), 0),
    OldestDaysPastDue = ISNULL(MAX(CASE WHEN b.IsOpen = 1 THEN b.DaysPastDue END), 0),
    CreditUtilizationPct = CAST(ISNULL(100.0 * SUM(b.OpenBalance) / NULLIF(c.CreditLimit, 0), 0) AS DECIMAL(6,2)),
    AvgDaysLateHistoric  = CAST(ISNULL(h.AvgDaysLate, 0) AS DECIMAL(6,1)),
    SettledInvoices      = ISNULL(h.SettledInvoices, 0),
    -- cash already banked against this account that nobody has matched
    UnappliedCash        = CAST(ISNULL(u.UnappliedCash, 0) AS DECIMAL(14,2)),
    UnappliedReceipts    = ISNULL(u.UnappliedReceipts, 0),
    OldestUnappliedDays  = ISNULL(u.OldestUnappliedDays, 0),
    -- what the customer owes NET of money we are already holding. The gross
    -- figure is what the ledger says; this is what a collector should quote.
    NetExposure          = CAST(ISNULL(SUM(b.OpenBalance), 0) - ISNULL(u.UnappliedCash, 0) AS DECIMAL(14,2)),
    BrokenPromises90d    = ISNULL(pr.BrokenCount, 0),
    BrokenPromiseDollars = CAST(ISNULL(pr.BrokenDollars, 0) AS DECIMAL(14,2))
FROM dbo.Dim_Customer c
JOIN dbo.Dim_Collector    col ON col.CollectorKey = c.CollectorKey
JOIN dbo.Dim_PaymentTerms t   ON t.TermsKey = c.CurrentTermsKey
LEFT JOIN B b ON b.CustomerKey = c.CustomerKey
LEFT JOIN Hist h ON h.CustomerKey = c.CustomerKey
LEFT JOIN dbo.fn_UnappliedCash(@AsOf) u ON u.CustomerKey = c.CustomerKey
OUTER APPLY (
    SELECT BrokenCount   = COUNT(*),
           BrokenDollars = SUM(s.PromisedAmount)
    FROM dbo.fn_PromiseStatus(@AsOf, 3, 90.00) s
    WHERE s.CustomerKey = c.CustomerKey
      AND s.PromiseStatus = 'Broken'
      AND s.PromisedPayDate > DATEADD(DAY, -90, @AsOf)
) pr
GROUP BY c.CustomerKey, c.CustomerID, c.CustomerName, c.Segment, c.Region, c.SalesRep,
         c.RiskTier, c.CreditLimit, col.CollectorID, col.CollectorName, col.Team,
         t.TermsCode, t.NetDays, h.AvgDaysLate, h.SettledInvoices,
         u.UnappliedCash, u.UnappliedReceipts, u.OldestUnappliedDays,
         pr.BrokenCount, pr.BrokenDollars
);
GO

-- The reporting-date snapshot. Power Query and the Excel workbook read this;
-- anything needing another date calls dbo.fn_CustomerAR(<date>) directly.
CREATE VIEW dbo.vw_CustomerAR AS SELECT * FROM dbo.fn_CustomerAR('2025-12-31');
GO

/*
--------------------------------------------------------------------------------
vw_PriorityActionQueue -- the operational control.

The unit of work is the CUSTOMER, because a collector rings an account, not an
invoice.

RANKING
    By collectable exposure: aging-weighted past-due dollars, less 75% of the
    aging-weighted DISPUTED dollars. Chasing cash on a disputed invoice does not
    work -- that balance needs the dispute resolving, which is a different job
    for a different team -- so it is discounted rather than driving the call
    list. Because both terms carry the same weighting, the figure cannot go
    negative; it is floored at zero anyway as a guard.

BOUNDING
    A ranked list of every account with a past-due balance is a ledger, not a
    worklist: it runs to roughly three quarters of the customer base, and six
    collectors cannot call that many accounts in a day. The view therefore stays
    complete for audit, and IsTodaysWorklist marks what is actionable today:
        - every ESCALATE or credit-hold account, whatever its rank; plus
        - each collector's top CallsPerCollectorPerDay accounts by exposure.
    Work is ranked WITHIN collector as well as overall, because the queue is
    handed out per collector: a single overall top-N would give one collector
    the whole day and the rest nothing.

CREDIT HOLD IS A FLAG, NOT AN ACTION
    Holding or releasing orders is an order-management decision that sits
    alongside a collections action rather than replacing it. As a competing
    branch in the action ladder it was also unreachable -- every over-limit
    account in this book already qualifies for ESCALATE, so the branch never
    fired and the control silently did nothing.

ActionCode is mutually exclusive and ordered by what should happen FIRST, so the
queue never offers a collector two conflicting next steps.
--------------------------------------------------------------------------------
*/
CREATE FUNCTION dbo.fn_PriorityActionQueue (@AsOf DATE)
RETURNS TABLE
AS RETURN
(
WITH T AS (
    SELECT
        CreditLimitWarn = MAX(CASE WHEN MetricName = 'CreditUtilization' THEN WarningValue END),
        -- Collections capacity assumption, stated here rather than buried in a
        -- TOP N. Retune if the team size or call rate changes.
        CallsPerCollectorPerDay = CAST(10 AS INT)
    FROM dbo.Ref_ARTargets
),
Scored AS (
    SELECT ca.*,
        t.CreditLimitWarn, t.CallsPerCollectorPerDay,
        -- money genuinely chaseable by a phone call today (same units on both terms)
        -- Money already sitting in the bank against this account is not
        -- chaseable by a phone call either: it is chaseable by matching it.
        -- Netting it here stops the queue ranking work that does not exist.
        CollectableExposure = CAST(
            CASE WHEN ca.WeightedExposure - ca.DisputedWeightedExposure * 0.75 - ca.UnappliedCash > 0
                 THEN ca.WeightedExposure - ca.DisputedWeightedExposure * 0.75 - ca.UnappliedCash
                 ELSE 0 END AS DECIMAL(14,2)),
        CreditHoldFlag = CAST(CASE WHEN ca.CreditUtilizationPct > t.CreditLimitWarn THEN 1 ELSE 0 END AS BIT)
    FROM dbo.fn_CustomerAR(@AsOf) ca CROSS JOIN T t
    WHERE ca.PastDueBalance > 0 OR ca.DisputedBalance > 0
       OR ca.CreditUtilizationPct > t.CreditLimitWarn
),
Ranked AS (
    SELECT s.*,
        PriorityRank  = ROW_NUMBER() OVER (ORDER BY s.CollectableExposure DESC, s.Balance90Plus DESC, s.CustomerID),
        CollectorRank = ROW_NUMBER() OVER (PARTITION BY s.CollectorID
                                           ORDER BY s.CollectableExposure DESC, s.Balance90Plus DESC, s.CustomerID),
        ActionCode = CASE
            -- Highest precedence, ahead of everything: this customer has
            -- already paid and nobody matched the cash. Ringing them is the one
            -- call that damages a paying relationship, and the work belongs to
            -- cash application, not collections.
            WHEN s.UnappliedCash > 0
             AND s.UnappliedCash >= s.PastDueBalance * 0.5        THEN 'APPLY_CASH'
            WHEN s.DisputedBalance > 0
             AND s.DisputedBalance >= s.PastDueBalance * 0.5      THEN 'RESOLVE_DISPUTE'
            -- Two broken promises in ninety days is the earliest reliable
            -- distress signal available, and it costs nothing to collect.
            WHEN s.BrokenPromises90d >= 2                         THEN 'ESCALATE'
            WHEN s.Balance90Plus > 0 AND s.RiskTier = 'High'      THEN 'ESCALATE'
            WHEN s.Balance90Plus > 0                              THEN 'FINAL_NOTICE'
            WHEN s.OldestDaysPastDue > 30                         THEN 'COLLECTION_CALL'
            WHEN s.PastDueBalance > 0
             AND s.AvgDaysLateHistoric <= 5                       THEN 'COURTESY_REMINDER'
            WHEN s.PastDueBalance > 0                             THEN 'STANDARD_DUNNING'
            ELSE                                                       'CREDIT_REVIEW'
        END
    FROM Scored s
)
SELECT
    r.PriorityRank, r.CollectorRank,
    r.CustomerID, r.CustomerName, r.Segment, r.Region, r.RiskTier,
    r.CollectorID, r.CollectorName, r.Team, r.TermsCode,
    r.OpenBalance, r.PastDueBalance, r.Balance90Plus, r.DisputedBalance,
    r.WeightedExposure, r.DisputedWeightedExposure, r.CollectableExposure,
    r.OldestDaysPastDue, r.CreditUtilizationPct, r.AvgDaysLateHistoric,
    r.UnappliedCash, r.NetExposure, r.BrokenPromises90d,
    r.ActionCode, r.CreditHoldFlag,
    IsTodaysWorklist = CAST(CASE
        WHEN r.ActionCode = 'ESCALATE' OR r.CreditHoldFlag = 1 THEN 1
        WHEN r.CollectorRank <= r.CallsPerCollectorPerDay      THEN 1
        ELSE 0 END AS BIT),
    RecommendedAction = CASE r.ActionCode
        WHEN 'APPLY_CASH'        THEN 'Do not call. Cash is already banked against this account and unmatched -- route to cash application to apply it before any collections contact.'
        WHEN 'RESOLVE_DISPUTE'   THEN 'Route to billing/sales: most of this balance is disputed and will not pay until the dispute closes.'
        WHEN 'ESCALATE'          THEN 'Escalate: high-risk account carrying a 90+ balance. Hold further shipments and agree a payment plan or refer for recovery.'
        WHEN 'FINAL_NOTICE'      THEN 'Issue a final notice on the 90+ balance and confirm a payment date in writing.'
        WHEN 'COLLECTION_CALL'   THEN 'Collection call: oldest item is more than 30 days past due.'
        WHEN 'COURTESY_REMINDER' THEN 'Courtesy reminder only: this account normally pays on time, so treat it as an oversight.'
        WHEN 'STANDARD_DUNNING'  THEN 'Standard dunning letter and follow-up call.'
        ELSE                          'Credit review: at or over the credit limit with nothing past due. Reassess the limit before releasing further orders.'
    END
    + CASE WHEN r.CreditHoldFlag = 1 AND r.ActionCode <> 'CREDIT_REVIEW'
           THEN ' Also place a credit hold: the account is at or over its credit limit.' ELSE '' END
FROM Ranked r
);
GO

-- The reporting-date snapshot of the worklist. Backtests and UAT call
-- dbo.fn_PriorityActionQueue(<date>) to rebuild the queue as it stood then.
CREATE VIEW dbo.vw_PriorityActionQueue AS SELECT * FROM dbo.fn_PriorityActionQueue('2025-12-31');
GO

DECLARE @missing VARCHAR(400) = '';
IF OBJECT_ID('dbo.vw_MonthlyCreditSales', 'V')  IS NULL SET @missing += 'vw_MonthlyCreditSales ';
IF OBJECT_ID('dbo.fn_DSO', 'IF')                IS NULL SET @missing += 'fn_DSO ';
IF OBJECT_ID('dbo.fn_ARKPI', 'IF')              IS NULL SET @missing += 'fn_ARKPI ';
IF OBJECT_ID('dbo.vw_ARKPIMonthly', 'V')        IS NULL SET @missing += 'vw_ARKPIMonthly ';
IF OBJECT_ID('dbo.fn_PromiseStatus', 'IF')      IS NULL SET @missing += 'fn_PromiseStatus ';
IF OBJECT_ID('dbo.fn_PromiseKeptRate', 'IF')    IS NULL SET @missing += 'fn_PromiseKeptRate ';
IF OBJECT_ID('dbo.vw_PromiseStatus', 'V')       IS NULL SET @missing += 'vw_PromiseStatus ';
IF OBJECT_ID('dbo.fn_CustomerAR', 'IF')         IS NULL SET @missing += 'fn_CustomerAR ';
IF OBJECT_ID('dbo.fn_PriorityActionQueue', 'IF') IS NULL SET @missing += 'fn_PriorityActionQueue ';
IF OBJECT_ID('dbo.vw_CustomerAR', 'V')          IS NULL SET @missing += 'vw_CustomerAR ';
IF OBJECT_ID('dbo.vw_PriorityActionQueue', 'V') IS NULL SET @missing += 'vw_PriorityActionQueue ';
IF @missing <> '' THROW 50003, 'FAILED to create KPI objects -- scroll up for the compile error.', 1;
PRINT 'KPI objects created and verified.';
GO
