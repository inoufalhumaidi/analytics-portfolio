/*
================================================================================
Project 2 -- Vantage Wholesale Supply: Receivables Performance
Script:  03_core_views.sql
Purpose: THE single definition of what an invoice still owes, and how old that
         balance is. Every downstream object -- data-quality checks, KPI views,
         stored procedures, the Excel workbook -- reads its balances from here
         so that one formula can never drift into several.

Why a function and not a view:
    Ageing only means something "as of" a date. A parameterless view would bake
    in one reporting date and quietly lie on every other one. dbo.fn_ARBalance
    (@AsOf) makes the date explicit and keeps the object composable.

THE SETTLEMENT IDENTITY (the formula everything else depends on):

    OpenBalance = InvoiceAmount
                - cash applied       (Fact_CashApplication.AppliedAmount)
                - discounts taken    (Fact_CashApplication.DiscountTaken)
                - credit memos       (Fact_Adjustment 'CreditMemo')
                - write-offs         (Fact_Adjustment 'WriteOff')

    The discount term is easy to miss and expensive to miss: it is stored
    separately from the cash, so omitting it leaves every 2/10 NET30 invoice
    looking permanently 2% unpaid, which inflates both ageing and DSO.

Cash that has arrived but has not been matched to an invoice does NOT reduce a
balance here. It is reported per customer by dbo.fn_UnappliedCash, because an
invoice whose payment sits unapplied looks exactly like an unpaid one, and the
difference decides whether a collector rings the customer or cash application
does the work.

Over-applied cash is kept, not hidden. Duplicate receipts (a real defect in
this dataset) push an invoice's balance below zero. Summing raw balances would
net that credit against genuinely open invoices and UNDERSTATE total AR, so
OpenBalance is floored at zero and the excess is surfaced separately as
OverAppliedAmount for the data-quality layer to report.

DATA DISCLOSURE: Vantage Wholesale Supply is fictional; all data is synthetic.
================================================================================
*/

USE VantageAR;
GO

IF OBJECT_ID('dbo.fn_ARBalance', 'IF') IS NOT NULL DROP FUNCTION dbo.fn_ARBalance;
GO

CREATE FUNCTION dbo.fn_ARBalance (@AsOf DATE)
RETURNS TABLE
AS RETURN
(
    WITH AsOfKey AS (
        SELECT k = YEAR(@AsOf) * 10000 + MONTH(@AsOf) * 100 + DAY(@AsOf)
    ),
    Cash AS (
        -- Cash APPLIED to the invoice on or before @AsOf.
        --
        -- The predicate is on the APPLICATION date, never the receipt date.
        -- Cash banked on 28 December and matched to an invoice on 4 January has
        -- not cured anything in December: the invoice was still open at the
        -- year end, and the ageing bucket it sat in was real. Using the receipt
        -- date here would quietly improve every month end by whatever the cash
        -- application team had not got to yet.
        --
        -- The money that has arrived but has not been matched is not lost --
        -- it is reported separately by dbo.fn_UnappliedCash, because "is this
        -- invoice open?" and "does this customer owe us money?" are different
        -- questions and only one of them is answered here.
        SELECT p.InvoiceKey,
               PaidToDate     = SUM(p.AppliedAmount),
               DiscountToDate = SUM(p.DiscountTaken)
        FROM dbo.Fact_CashApplication p
        CROSS JOIN AsOfKey a
        WHERE p.ApplicationDateKey <= a.k
        GROUP BY p.InvoiceKey
    ),
    Adj AS (                       -- credit memos and write-offs on or before @AsOf
        SELECT j.InvoiceKey,
               CreditMemoToDate = SUM(CASE WHEN j.AdjustmentType = 'CreditMemo' THEN j.Amount ELSE 0 END),
               WriteOffToDate   = SUM(CASE WHEN j.AdjustmentType = 'WriteOff'   THEN j.Amount ELSE 0 END)
        FROM dbo.Fact_Adjustment j
        CROSS JOIN AsOfKey a
        WHERE j.AdjustmentDateKey <= a.k
        GROUP BY j.InvoiceKey
    ),
    Base AS (
        SELECT
            f.InvoiceKey, f.InvoiceNo, f.CustomerKey, f.TermsKey,
            InvoiceDate = di.[Date],
            DueDate     = dd.[Date],
            f.InvoiceAmount,
            PaidToDate       = ISNULL(c.PaidToDate, 0),
            DiscountToDate   = ISNULL(c.DiscountToDate, 0),
            CreditMemoToDate = ISNULL(j.CreditMemoToDate, 0),
            WriteOffToDate   = ISNULL(j.WriteOffToDate, 0),
            RawBalance = f.InvoiceAmount
                       - ISNULL(c.PaidToDate, 0)
                       - ISNULL(c.DiscountToDate, 0)
                       - ISNULL(j.CreditMemoToDate, 0)
                       - ISNULL(j.WriteOffToDate, 0),
            DaysPastDue = DATEDIFF(DAY, dd.[Date], @AsOf),
            -- a dispute counts as open only if it was raised, and not yet closed, by @AsOf
            IsDisputedOpen = CAST(CASE
                WHEN f.IsDisputed = 1
                 AND f.DisputeOpenedKey <= (SELECT k FROM AsOfKey)
                 AND (f.DisputeClosedKey IS NULL OR f.DisputeClosedKey > (SELECT k FROM AsOfKey))
                THEN 1 ELSE 0 END AS BIT),
            f.DisputeReason
        FROM dbo.Fact_Invoice f
        JOIN dbo.Dim_Date di ON di.DateKey = f.InvoiceDateKey
        JOIN dbo.Dim_Date dd ON dd.DateKey = f.DueDateKey
        LEFT JOIN Cash c ON c.InvoiceKey = f.InvoiceKey   -- LEFT: an unpaid invoice has no cash rows
        LEFT JOIN Adj  j ON j.InvoiceKey = f.InvoiceKey   -- LEFT: most invoices have no adjustments
        CROSS JOIN AsOfKey a
        WHERE f.InvoiceDateKey <= a.k          -- an invoice raised after @AsOf does not exist yet
    )
    SELECT
        b.InvoiceKey, b.InvoiceNo, b.CustomerKey, b.TermsKey,
        b.InvoiceDate, b.DueDate, b.InvoiceAmount,
        b.PaidToDate, b.DiscountToDate, b.CreditMemoToDate, b.WriteOffToDate,
        b.RawBalance,
        -- floored balance: what the company can still expect to collect
        OpenBalance       = CASE WHEN b.RawBalance > 0.005 THEN b.RawBalance ELSE 0 END,
        -- cash applied beyond the invoice value (duplicate/over-application)
        OverAppliedAmount = CASE WHEN b.RawBalance < -0.005 THEN -b.RawBalance ELSE 0 END,
        IsOpen     = CAST(CASE WHEN b.RawBalance > 0.005 THEN 1 ELSE 0 END AS BIT),
        IsPastDue  = CAST(CASE WHEN b.RawBalance > 0.005 AND b.DaysPastDue > 0 THEN 1 ELSE 0 END AS BIT),
        DaysPastDue = b.DaysPastDue,
        -- ageing bucket only applies to something still open
        BucketKey  = CASE WHEN b.RawBalance > 0.005 THEN bk.BucketKey END,
        BucketName = CASE WHEN b.RawBalance > 0.005 THEN bk.BucketName END,
        RiskWeight = CASE WHEN b.RawBalance > 0.005 THEN bk.RiskWeight END,
        b.IsDisputedOpen, b.DisputeReason,
        AsOfDate = @AsOf
    FROM Base b
    LEFT JOIN dbo.Ref_AgingBucket bk
           ON b.DaysPastDue BETWEEN bk.MinDaysPastDue AND bk.MaxDaysPastDue
);
GO

/*
--------------------------------------------------------------------------------
vw_ARBalance -- the reporting-date snapshot (2025-12-31).
Power Query and the Excel workbook read this; anything needing another date
calls dbo.fn_ARBalance(<date>) directly.
--------------------------------------------------------------------------------
*/
IF OBJECT_ID('dbo.vw_ARBalance', 'V') IS NOT NULL DROP VIEW dbo.vw_ARBalance;
GO
CREATE VIEW dbo.vw_ARBalance AS
SELECT * FROM dbo.fn_ARBalance('2025-12-31');
GO

/*
--------------------------------------------------------------------------------
vw_ARBalanceDetail -- the same snapshot enriched with the customer, terms and
collector attributes the collections team needs on screen. This is the one
extract the Excel workbook imports through Power Query.
--------------------------------------------------------------------------------
*/
IF OBJECT_ID('dbo.vw_ARBalanceDetail', 'V') IS NOT NULL DROP VIEW dbo.vw_ARBalanceDetail;
GO
CREATE VIEW dbo.vw_ARBalanceDetail AS
SELECT
    b.InvoiceKey, b.InvoiceNo,
    c.CustomerID, c.CustomerName, c.Segment, c.Region, c.SalesRep,
    c.RiskTier, c.CreditLimit,
    col.CollectorID, col.CollectorName, col.Team,
    t.TermsCode, t.NetDays, t.DiscountPct,
    b.InvoiceDate, b.DueDate, b.InvoiceAmount,
    b.PaidToDate, b.DiscountToDate, b.CreditMemoToDate, b.WriteOffToDate,
    b.OpenBalance, b.OverAppliedAmount,
    b.IsOpen, b.IsPastDue, b.DaysPastDue,
    b.BucketKey, b.BucketName, b.RiskWeight,
    b.IsDisputedOpen, b.DisputeReason,
    b.AsOfDate
FROM dbo.vw_ARBalance b
JOIN dbo.Dim_Customer     c   ON c.CustomerKey = b.CustomerKey
JOIN dbo.Dim_Collector    col ON col.CollectorKey = c.CollectorKey
JOIN dbo.Dim_PaymentTerms t   ON t.TermsKey = b.TermsKey;
GO

/*
--------------------------------------------------------------------------------
Creation gate: report what actually exists, not what was attempted. A script
that prints success after its objects failed to compile is worse than one that
prints nothing, because it stops anyone from looking.
--------------------------------------------------------------------------------
*/

/*
--------------------------------------------------------------------------------
fn_UnappliedCash(@AsOf) -- money already in the bank that nobody has matched to
an invoice, per customer.

WHY THIS MATTERS MORE THAN ITS SIZE SUGGESTS
    An invoice whose payment is sitting unapplied looks exactly like an invoice
    that has not been paid. The ageing report shows it, the priority queue ranks
    it, and a collector rings a customer who settled weeks ago. That call costs
    more than the balance: it is the one interaction that makes a paying
    customer doubt whether Vantage knows what it is doing.

    So this is not a finance-tidiness metric. It is a routing rule: these
    accounts belong in cash application, not on a call list.

WHY EACH RECEIPT IS FLOORED SEPARATELY
    A duplicate receipt is OVER-applied -- more cash matched than the receipt
    carried. Summing raw remainders would let that negative offset genuine
    unapplied cash on other receipts and understate the backlog. Each receipt is
    therefore floored at zero on its own, and the over-application is surfaced
    separately for the data-quality layer, exactly as invoice balances are.

    Receipts dated after @AsOf do not exist yet. Applications dated after @AsOf
    have not happened yet, which is precisely how cash banked in December but
    applied in January shows up as unapplied at the year end -- which it was.
--------------------------------------------------------------------------------
*/
IF OBJECT_ID('dbo.vw_UnappliedCash', 'V') IS NOT NULL DROP VIEW dbo.vw_UnappliedCash;
IF OBJECT_ID('dbo.fn_UnappliedCash', 'IF') IS NOT NULL DROP FUNCTION dbo.fn_UnappliedCash;
GO

CREATE FUNCTION dbo.fn_UnappliedCash (@AsOf DATE)
RETURNS TABLE
AS RETURN
(
    WITH AsOfKey AS (SELECT k = YEAR(@AsOf) * 10000 + MONTH(@AsOf) * 100 + DAY(@AsOf)),
    Receipt AS (
        SELECT
            r.ReceiptKey, r.ReceiptNo, r.CustomerKey, r.ReceiptDateKey,
            r.ReceiptAmount, r.PaymentMethod, r.RemittanceAdviceReceived,
            AppliedToDate = ISNULL((
                SELECT SUM(a.AppliedAmount)
                FROM dbo.Fact_CashApplication a
                CROSS JOIN AsOfKey k2
                WHERE a.ReceiptKey = r.ReceiptKey AND a.ApplicationDateKey <= k2.k), 0),
            AgeDays = DATEDIFF(DAY, d.[Date], @AsOf)
        FROM dbo.Fact_CashReceipt r
        JOIN dbo.Dim_Date d ON d.DateKey = r.ReceiptDateKey
        CROSS JOIN AsOfKey a
        WHERE r.ReceiptDateKey <= a.k
    )
    SELECT
        CustomerKey,
        UnappliedCash = CAST(SUM(CASE WHEN ReceiptAmount - AppliedToDate > 0.005
                                      THEN ReceiptAmount - AppliedToDate ELSE 0 END) AS DECIMAL(14,2)),
        OverAppliedCash = CAST(SUM(CASE WHEN ReceiptAmount - AppliedToDate < -0.005
                                        THEN AppliedToDate - ReceiptAmount ELSE 0 END) AS DECIMAL(14,2)),
        UnappliedReceipts = SUM(CASE WHEN ReceiptAmount - AppliedToDate > 0.005 THEN 1 ELSE 0 END),
        FullyUnappliedReceipts = SUM(CASE WHEN AppliedToDate <= 0.005 THEN 1 ELSE 0 END),
        NoRemittanceAdvice = SUM(CASE WHEN ReceiptAmount - AppliedToDate > 0.005
                                       AND RemittanceAdviceReceived = 0 THEN 1 ELSE 0 END),
        OldestUnappliedDays = MAX(CASE WHEN ReceiptAmount - AppliedToDate > 0.005 THEN AgeDays END),
        AsOfDate = @AsOf
    FROM Receipt
    GROUP BY CustomerKey
);
GO

CREATE VIEW dbo.vw_UnappliedCash AS
SELECT u.*, c.CustomerID, c.CustomerName, c.Segment, c.Region, col.CollectorName
FROM dbo.fn_UnappliedCash('2025-12-31') u
JOIN dbo.Dim_Customer c    ON c.CustomerKey = u.CustomerKey
JOIN dbo.Dim_Collector col ON col.CollectorKey = c.CollectorKey;
GO
DECLARE @missing VARCHAR(400) = '';
IF OBJECT_ID('dbo.fn_ARBalance', 'IF')      IS NULL SET @missing += 'fn_ARBalance ';
IF OBJECT_ID('dbo.vw_ARBalance', 'V')       IS NULL SET @missing += 'vw_ARBalance ';
IF OBJECT_ID('dbo.vw_ARBalanceDetail', 'V') IS NULL SET @missing += 'vw_ARBalanceDetail ';
IF OBJECT_ID('dbo.fn_UnappliedCash', 'IF')  IS NULL SET @missing += 'fn_UnappliedCash ';
IF OBJECT_ID('dbo.vw_UnappliedCash', 'V')   IS NULL SET @missing += 'vw_UnappliedCash ';
IF @missing <> ''
    THROW 50001, 'FAILED to create core objects -- scroll up for the compile error.', 1;
PRINT 'Core objects created and verified: fn_ARBalance, vw_ARBalance, vw_ARBalanceDetail, fn_UnappliedCash, vw_UnappliedCash.';
GO
