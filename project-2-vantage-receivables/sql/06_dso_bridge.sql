/*
================================================================================
Project 2 -- Vantage Wholesale Supply: Receivables Performance
Script:  06_dso_bridge.sql
Purpose: Answer the question the scorecard raises but cannot settle --
         of the days DSO has risen, how many did Vantage GRANT itself through
         the terms it sold, and how many are customers TAKING beyond those
         terms?

WHY THIS EXISTS
    "DSO is up 14 days" starts an argument and settles nothing. Finance reads
    it as slack collections; collections reads it as Sales having sold thirty
    extra days on the contractor book. Both are testable, and they lead to
    completely different actions: one is a call-list problem, the other is a
    terms-policy problem that no amount of collector time will fix.

WHY THE BRIDGE IS BUILT ON CLASSIC DSO, NOT THE COUNTBACK HEADLINE
    Countback DSO is the better headline because it survives seasonality, but
    it does NOT decompose: consuming a balance month by month gives no clean
    way to attribute days to causes, and every published attempt leaves a
    residual bar that invites someone to plug it into the nearest component.

    Classic DSO decomposes EXACTLY. By definition

        DSO_classic = TotalAR / SalesPerDay

    and TotalAR partitions, with no overlap and nothing left over, into

        TotalAR = CurrentAR                (not yet due -- terms we granted)
                + DisputedPastDueAR        (late, but blocked by a dispute)
                + UndisputedPastDueAR      (late, and nothing is stopping it)

    Divide each part by the same SalesPerDay and the three day-figures sum to
    DSO_classic to the cent. No residual, no judgement, no plug. UAT asserts
    the identity rather than trusting it.

    Both DSO variants are published side by side, and the gap between them is
    reported as a figure in its own right, because quoting a countback DSO to a
    lender who benchmarks on classic invites an argument you will lose.

THE SECOND HALF: IS THE TARGET EVEN ACHIEVABLE?
    Weighted Average Terms is the dollar-weighted average of the NetDays
    actually sold, taken from the TermsKey stored on each invoice rather than
    the customer's current terms -- otherwise a later terms change is smeared
    backwards over history and the policy shift becomes invisible.

    WAT is the floor. It is the DSO the company would post if every single
    customer paid exactly on the due date. A DSO target set below WAT is not a
    stretch goal, it is arithmetic nobody can satisfy, and publishing one
    teaches a collections team that the scorecard is theatre.

DATA DISCLOSURE: Vantage Wholesale Supply is fictional; all data is synthetic.
================================================================================
*/

USE VantageAR;
GO

IF OBJECT_ID('dbo.vw_DSOBridgeMonthly', 'V')  IS NOT NULL DROP VIEW dbo.vw_DSOBridgeMonthly;
IF OBJECT_ID('dbo.vw_DSOBridge', 'V')         IS NOT NULL DROP VIEW dbo.vw_DSOBridge;
IF OBJECT_ID('dbo.vw_TermsMixMonthly', 'V')   IS NOT NULL DROP VIEW dbo.vw_TermsMixMonthly;
IF OBJECT_ID('dbo.fn_DSOBridge', 'IF')        IS NOT NULL DROP FUNCTION dbo.fn_DSOBridge;
IF OBJECT_ID('dbo.vw_BillingLagMonthly', 'V') IS NOT NULL DROP VIEW dbo.vw_BillingLagMonthly;
IF OBJECT_ID('dbo.fn_BillingLag', 'IF')       IS NOT NULL DROP FUNCTION dbo.fn_BillingLag;
IF OBJECT_ID('dbo.fn_DBT', 'IF')              IS NOT NULL DROP FUNCTION dbo.fn_DBT;
IF OBJECT_ID('dbo.fn_WATShiftShare', 'IF')    IS NOT NULL DROP FUNCTION dbo.fn_WATShiftShare;
IF OBJECT_ID('dbo.fn_WeightedAvgTerms', 'IF') IS NOT NULL DROP FUNCTION dbo.fn_WeightedAvgTerms;
GO

/*
--------------------------------------------------------------------------------
fn_WeightedAvgTerms(@From, @To) -- the terms Vantage actually sold in a window.

Dollar-weighted, not count-weighted: two hundred $300 MRO invoices on NET30 do
not offset one $60,000 OEM invoice on NET60 in cash terms, and a count-weighted
average would say they do.

TermsKey is read from Fact_Invoice, never from Dim_Customer. An invoice carries
the terms it was sold on; the customer record carries the terms in force today.
Reading the second would rewrite history every time a customer is re-papered.
--------------------------------------------------------------------------------
*/
CREATE FUNCTION dbo.fn_WeightedAvgTerms (@From DATE, @To DATE)
RETURNS TABLE
AS RETURN
(
    SELECT
        PeriodFrom = @From,
        PeriodTo   = @To,
        CreditSales = CAST(ISNULL(SUM(f.InvoiceAmount), 0) AS DECIMAL(14,2)),
        InvoiceCount = COUNT(*),
        WeightedAvgTermsDays = CAST(SUM(f.InvoiceAmount * t.NetDays)
                                  / NULLIF(SUM(f.InvoiceAmount), 0) AS DECIMAL(9,2)),
        -- shown alongside purely to demonstrate how far the two diverge
        UnweightedAvgTermsDays = CAST(AVG(CAST(t.NetDays AS FLOAT)) AS DECIMAL(9,2)),
        PctDollarsOnExtendedTerms = CAST(100.0 * SUM(CASE WHEN t.NetDays > 45 THEN f.InvoiceAmount ELSE 0 END)
                                  / NULLIF(SUM(f.InvoiceAmount), 0) AS DECIMAL(6,2))
    FROM dbo.Fact_Invoice f
    JOIN dbo.Dim_PaymentTerms t ON t.TermsKey = f.TermsKey
    JOIN dbo.Dim_Date d ON d.DateKey = f.InvoiceDateKey
    WHERE d.[Date] BETWEEN @From AND @To
);
GO

/*
--------------------------------------------------------------------------------
vw_TermsMixMonthly -- WAT by month, with the dollar share sitting on extended
terms. This is the series that shows a terms policy change as a step, not a
drift.
--------------------------------------------------------------------------------
*/
CREATE VIEW dbo.vw_TermsMixMonthly AS
SELECT
    d.YearMonth,
    MonthEndDate = MAX(d.MonthEndDate),
    CreditSales  = CAST(SUM(f.InvoiceAmount) AS DECIMAL(14,2)),
    WeightedAvgTermsDays = CAST(SUM(f.InvoiceAmount * t.NetDays)
                              / NULLIF(SUM(f.InvoiceAmount), 0) AS DECIMAL(9,2)),
    PctDollarsNet30 = CAST(100.0 * SUM(CASE WHEN t.NetDays = 30 THEN f.InvoiceAmount ELSE 0 END)
                         / NULLIF(SUM(f.InvoiceAmount), 0) AS DECIMAL(6,2)),
    PctDollarsNet45 = CAST(100.0 * SUM(CASE WHEN t.NetDays = 45 THEN f.InvoiceAmount ELSE 0 END)
                         / NULLIF(SUM(f.InvoiceAmount), 0) AS DECIMAL(6,2)),
    PctDollarsNet60 = CAST(100.0 * SUM(CASE WHEN t.NetDays = 60 THEN f.InvoiceAmount ELSE 0 END)
                         / NULLIF(SUM(f.InvoiceAmount), 0) AS DECIMAL(6,2)),
    PctDollarsExtended = CAST(100.0 * SUM(CASE WHEN t.NetDays > 45 THEN f.InvoiceAmount ELSE 0 END)
                            / NULLIF(SUM(f.InvoiceAmount), 0) AS DECIMAL(6,2))
FROM dbo.Fact_Invoice f
JOIN dbo.Dim_PaymentTerms t ON t.TermsKey = f.TermsKey
JOIN dbo.Dim_Date d ON d.DateKey = f.InvoiceDateKey
GROUP BY d.YearMonth;
GO

/*
--------------------------------------------------------------------------------
fn_WATShiftShare(...) -- why did Weighted Average Terms move?

WAT can rise for two quite different reasons, and only one of them is a
decision somebody made:

    RATE effect  the same customers were re-papered onto longer terms.
                 That is a policy change, and it has an owner.
    MIX effect   nobody's terms changed; customers who were already on long
                 terms simply bought more this period.

Standard shift-share, with the residual kept visible rather than folded in:

    dWAT = SUM_c (w1 - w0) * n0      MIX
         + SUM_c  w0 * (n1 - n0)     RATE
         + SUM_c (w1 - w0)*(n1 - n0) INTERACTION

where for customer c, w is its share of period dollars and n its dollar-
weighted NetDays.

Customers trading in only one of the two periods have no terms figure in the
other. Rather than invent one, their missing n is set to the value they do
have, which makes their RATE and INTERACTION contributions exactly zero and
lands their whole effect in MIX -- which is what arriving or leaving IS. The
three components still sum to dWAT exactly, and UAT asserts that.
--------------------------------------------------------------------------------
*/
CREATE FUNCTION dbo.fn_WATShiftShare
    (@BaseFrom DATE, @BaseTo DATE, @CompFrom DATE, @CompTo DATE)
RETURNS TABLE
AS RETURN
(
    WITH Base AS (
        SELECT f.CustomerKey,
               Amt = SUM(f.InvoiceAmount),
               N   = SUM(f.InvoiceAmount * t.NetDays) / SUM(f.InvoiceAmount)
        FROM dbo.Fact_Invoice f
        JOIN dbo.Dim_PaymentTerms t ON t.TermsKey = f.TermsKey
        JOIN dbo.Dim_Date d ON d.DateKey = f.InvoiceDateKey
        WHERE d.[Date] BETWEEN @BaseFrom AND @BaseTo
        GROUP BY f.CustomerKey
    ),
    Comp AS (
        SELECT f.CustomerKey,
               Amt = SUM(f.InvoiceAmount),
               N   = SUM(f.InvoiceAmount * t.NetDays) / SUM(f.InvoiceAmount)
        FROM dbo.Fact_Invoice f
        JOIN dbo.Dim_PaymentTerms t ON t.TermsKey = f.TermsKey
        JOIN dbo.Dim_Date d ON d.DateKey = f.InvoiceDateKey
        WHERE d.[Date] BETWEEN @CompFrom AND @CompTo
        GROUP BY f.CustomerKey
    ),
    Tot AS (
        SELECT BaseAmt = (SELECT SUM(Amt) FROM Base),
               CompAmt = (SELECT SUM(Amt) FROM Comp)
    ),
    Joined AS (
        SELECT
            CustomerKey = ISNULL(b.CustomerKey, c.CustomerKey),
            w0 = ISNULL(b.Amt, 0) / NULLIF(t.BaseAmt, 0),
            w1 = ISNULL(c.Amt, 0) / NULLIF(t.CompAmt, 0),
            -- a customer absent from one period keeps the terms it does have,
            -- so its rate and interaction terms vanish and MIX absorbs it
            n0 = COALESCE(b.N, c.N),
            n1 = COALESCE(c.N, b.N)
        FROM Base b
        FULL OUTER JOIN Comp c ON c.CustomerKey = b.CustomerKey
        CROSS JOIN Tot t
    )
    SELECT
        BaseWAT = CAST((SELECT SUM(w0 * n0) FROM Joined) AS DECIMAL(9,2)),
        CompWAT = CAST((SELECT SUM(w1 * n1) FROM Joined) AS DECIMAL(9,2)),
        DeltaWAT = CAST((SELECT SUM(w1 * n1) - SUM(w0 * n0) FROM Joined) AS DECIMAL(9,2)),
        MixEffectDays  = CAST((SELECT SUM((w1 - w0) * n0) FROM Joined) AS DECIMAL(9,2)),
        RateEffectDays = CAST((SELECT SUM(w0 * (n1 - n0)) FROM Joined) AS DECIMAL(9,2)),
        InteractionDays = CAST((SELECT SUM((w1 - w0) * (n1 - n0)) FROM Joined) AS DECIMAL(9,2)),
        CustomersInBoth = (SELECT COUNT(*) FROM Base b JOIN Comp c ON c.CustomerKey = b.CustomerKey),
        CustomersArrived = (SELECT COUNT(*) FROM Comp c WHERE NOT EXISTS (SELECT 1 FROM Base b WHERE b.CustomerKey = c.CustomerKey)),
        CustomersLeft    = (SELECT COUNT(*) FROM Base b WHERE NOT EXISTS (SELECT 1 FROM Comp c WHERE c.CustomerKey = b.CustomerKey))
);
GO

/*
--------------------------------------------------------------------------------
fn_DBT(@AsOf, @Months) -- Days Beyond Terms, reported settled AND at-risk.

Settled DBT weights each cash application by how many days after the due date
it landed. Early payment contributes negatively; the sign is kept, because
netting it away would hide the discount-takers.

Settled DBT alone has a survivorship bias that INVERTS its meaning exactly when
it matters most: an invoice that is never paid never enters the average. As a
cohort slides towards default, only its healthy invoices settle, so settled DBT
can IMPROVE while the cohort collapses.

At-risk DBT fixes that by marking the unpaid dollars to the as-of date -- an
invoice 200 days past due and still open contributes 200 days at its open
balance. Deterioration then reads as deterioration.
--------------------------------------------------------------------------------
*/
CREATE FUNCTION dbo.fn_DBT (@AsOf DATE, @Months INT)
RETURNS TABLE
AS RETURN
(
    WITH WindowStart AS (SELECT s = DATEADD(MONTH, -@Months, @AsOf)),
    Settled AS (      -- cash applied inside the window, weighted by lateness
        SELECT
            Wt   = SUM(p.AppliedAmount + p.DiscountTaken),
            WtDays = SUM((p.AppliedAmount + p.DiscountTaken)
                       * DATEDIFF(DAY, dd.[Date], dp.[Date]))
        FROM dbo.Fact_CashApplication p
        JOIN dbo.Fact_Invoice f ON f.InvoiceKey = p.InvoiceKey
        JOIN dbo.Dim_Date dp ON dp.DateKey = p.ApplicationDateKey
        JOIN dbo.Dim_Date dd ON dd.DateKey = f.DueDateKey
        CROSS JOIN WindowStart w
        WHERE dp.[Date] > w.s AND dp.[Date] <= @AsOf
    ),
    StillOpen AS (    -- unpaid past-due dollars, marked to @AsOf
        SELECT
            Wt = SUM(b.OpenBalance),
            WtDays = SUM(b.OpenBalance * b.DaysPastDue)
        FROM dbo.fn_ARBalance(@AsOf) b
        WHERE b.IsOpen = 1 AND b.DaysPastDue > 0
    )
    SELECT
        AsOfDate = @AsOf,
        WindowMonths = @Months,
        SettledDollars = CAST(ISNULL(s.Wt, 0) AS DECIMAL(14,2)),
        OpenPastDueDollars = CAST(ISNULL(o.Wt, 0) AS DECIMAL(14,2)),
        DBT_Settled = CAST(s.WtDays / NULLIF(s.Wt, 0) AS DECIMAL(9,2)),
        DBT_AtRisk  = CAST((ISNULL(s.WtDays, 0) + ISNULL(o.WtDays, 0))
                         / NULLIF(ISNULL(s.Wt, 0) + ISNULL(o.Wt, 0), 0) AS DECIMAL(9,2)),
        -- when this is large and positive, settled DBT is flattering the book
        SurvivorshipGapDays = CAST((ISNULL(s.WtDays, 0) + ISNULL(o.WtDays, 0))
                         / NULLIF(ISNULL(s.Wt, 0) + ISNULL(o.Wt, 0), 0)
                         - s.WtDays / NULLIF(s.Wt, 0) AS DECIMAL(9,2))
    FROM Settled s CROSS JOIN StillOpen o
);
GO

/*
--------------------------------------------------------------------------------
fn_BillingLag(@From, @To) -- dollar-weighted days between shipping the goods and
issuing the invoice.

WHY IT IS NOT PART OF THE DSO BRIDGE
    DSO counts from the invoice date. Billing lag happens BEFORE that clock
    starts, so it is not a component of DSO and folding it in would break the
    identity the bridge depends on. It EXTENDS the measure instead:

        CashCycleDays = BillingLagDays + DSO_Classic

    That is the honest figure -- days from despatching goods to having the cash
    -- and it is larger than the number anyone is currently reporting.

WHY IT IS THE CHEAPEST DAY TO RECOVER
    Every other component needs a customer to behave differently. This one needs
    a batch-posting schedule changed. No negotiation, no phone call, no terms
    concession.

WHY A NEGATIVE LAG IS NOT CLIPPED TO ZERO
    An invoice dated before the goods shipped is pre-billing, which is a
    revenue-recognition question rather than a rounding artefact. Taking
    MAX(0, lag) would erase it silently, so negatives stay in the weighted
    average and are counted separately for the data-quality layer to raise.
--------------------------------------------------------------------------------
*/
IF OBJECT_ID('dbo.vw_BillingLagMonthly', 'V') IS NOT NULL DROP VIEW dbo.vw_BillingLagMonthly;
IF OBJECT_ID('dbo.fn_BillingLag', 'IF')       IS NOT NULL DROP FUNCTION dbo.fn_BillingLag;
GO

CREATE FUNCTION dbo.fn_BillingLag (@From DATE, @To DATE)
RETURNS TABLE
AS RETURN
(
    SELECT
        PeriodFrom = @From,
        PeriodTo   = @To,
        InvoiceCount = COUNT(*),
        CreditSales  = CAST(SUM(f.InvoiceAmount) AS DECIMAL(14,2)),
        -- dollar-weighted for the same reason as weighted average terms: two
        -- hundred small invoices posted promptly do not offset one large one
        -- held for a week
        BillingLagDays = CAST(SUM(f.InvoiceAmount * DATEDIFF(DAY, ds.[Date], di.[Date]))
                            / NULLIF(SUM(f.InvoiceAmount), 0) AS DECIMAL(9,2)),
        UnweightedLagDays = CAST(AVG(CAST(DATEDIFF(DAY, ds.[Date], di.[Date]) AS FLOAT)) AS DECIMAL(9,2)),
        PreBilledInvoices = SUM(CASE WHEN ds.[Date] > di.[Date] THEN 1 ELSE 0 END),
        WorstLagDays      = MAX(DATEDIFF(DAY, ds.[Date], di.[Date]))
    FROM dbo.Fact_Invoice f
    JOIN dbo.Dim_Date ds ON ds.DateKey = f.ShipDateKey
    JOIN dbo.Dim_Date di ON di.DateKey = f.InvoiceDateKey
    WHERE di.[Date] BETWEEN @From AND @To
);
GO

/*
--------------------------------------------------------------------------------
vw_BillingLagMonthly -- lag by month AND region, because a company-wide average
of 2.2 days hides one distribution centre at 4.8 and everyone else at 1.5.
--------------------------------------------------------------------------------
*/
CREATE VIEW dbo.vw_BillingLagMonthly AS
SELECT
    d.YearMonth,
    MonthEndDate = MAX(d.MonthEndDate),
    c.Region,
    InvoiceCount = COUNT(*),
    CreditSales  = CAST(SUM(f.InvoiceAmount) AS DECIMAL(14,2)),
    BillingLagDays = CAST(SUM(f.InvoiceAmount * DATEDIFF(DAY, ds.[Date], d.[Date]))
                        / NULLIF(SUM(f.InvoiceAmount), 0) AS DECIMAL(9,2))
FROM dbo.Fact_Invoice f
JOIN dbo.Dim_Customer c ON c.CustomerKey = f.CustomerKey
JOIN dbo.Dim_Date ds ON ds.DateKey = f.ShipDateKey
JOIN dbo.Dim_Date d  ON d.DateKey  = f.InvoiceDateKey
GROUP BY d.YearMonth, c.Region;
GO
/*
--------------------------------------------------------------------------------
fn_DSOBridge(@AsOf) -- the granted-versus-taken decomposition.

    GrantedDays  + DisputeDays + LatenessDays = DSO_Classic,   exactly.

Read it as: of the days of sales sitting in receivables, GrantedDays are the
ones Vantage sold on purpose and is not owed yet; DisputeDays are late but
blocked behind a commercial argument that collections cannot win by calling;
LatenessDays are late with nothing standing in the way, and are the only days a
call list can recover.

TermsPositionDays compares GrantedDays with the terms actually sold (WAT). A
positive figure means the not-yet-due book is carrying more days than the terms
mix implies -- normally rising sales, since a growing book front-loads current
balances. It is reported because without it a reader will misread GrantedDays
as the terms policy itself.

The same trailing-12-month sales window is used as in fn_DSO, so DSO_Classic
here is the same number fn_DSO publishes as DSO_Simple. UAT asserts the tie.
--------------------------------------------------------------------------------
*/
CREATE FUNCTION dbo.fn_DSOBridge (@AsOf DATE)
RETURNS TABLE
AS RETURN
(
    WITH Bal AS (
        SELECT
            TotalAR   = SUM(OpenBalance),
            CurrentAR = SUM(CASE WHEN DaysPastDue <= 0 THEN OpenBalance ELSE 0 END),
            DisputedPastDueAR = SUM(CASE WHEN DaysPastDue > 0 AND IsDisputedOpen = 1
                                         THEN OpenBalance ELSE 0 END),
            UndisputedPastDueAR = SUM(CASE WHEN DaysPastDue > 0 AND IsDisputedOpen = 0
                                           THEN OpenBalance ELSE 0 END)
        FROM dbo.fn_ARBalance(@AsOf)
    ),
    Months AS (
        SELECT TOP 12 s.CreditSales, s.DaysInMonth
        FROM dbo.vw_MonthlyCreditSales s
        WHERE s.MonthEndDate <= EOMONTH(@AsOf)
        ORDER BY s.MonthEndDate DESC
    ),
    Win AS (
        SELECT Sales12 = SUM(CreditSales), Days12 = SUM(DaysInMonth) FROM Months
    ),
    Terms AS (
        SELECT WeightedAvgTermsDays
        FROM dbo.fn_WeightedAvgTerms(DATEADD(MONTH, -12, @AsOf), @AsOf)
    ),
    Lag AS (
        SELECT BillingLagDays
        FROM dbo.fn_BillingLag(DATEADD(MONTH, -12, @AsOf), @AsOf)
    )
    SELECT
        AsOfDate = @AsOf,
        b.TotalAR, b.CurrentAR, b.DisputedPastDueAR, b.UndisputedPastDueAR,
        SalesPerDay = CAST(w.Sales12 / NULLIF(w.Days12, 0) AS DECIMAL(14,2)),
        DSO_Classic   = CAST(b.TotalAR             * w.Days12 / NULLIF(w.Sales12, 0) AS DECIMAL(9,2)),
        GrantedDays   = CAST(b.CurrentAR           * w.Days12 / NULLIF(w.Sales12, 0) AS DECIMAL(9,2)),
        DisputeDays   = CAST(b.DisputedPastDueAR   * w.Days12 / NULLIF(w.Sales12, 0) AS DECIMAL(9,2)),
        LatenessDays  = CAST(b.UndisputedPastDueAR * w.Days12 / NULLIF(w.Sales12, 0) AS DECIMAL(9,2)),
        WeightedAvgTermsDays = t.WeightedAvgTermsDays,
        TermsPositionDays = CAST(b.CurrentAR * w.Days12 / NULLIF(w.Sales12, 0)
                               - t.WeightedAvgTermsDays AS DECIMAL(9,2)),
        -- the recoverable half: what a call list can actually move
        RecoverableDays = CAST(b.UndisputedPastDueAR * w.Days12 / NULLIF(w.Sales12, 0) AS DECIMAL(9,2)),
        RecoverableCash = CAST(b.UndisputedPastDueAR AS DECIMAL(14,2)),
        -- Billing lag sits BEFORE the DSO clock starts, so it extends the
        -- measure rather than decomposing it. This is the honest figure: days
        -- from despatching goods to holding the cash.
        BillingLagDays = l.BillingLagDays,
        CashCycleDays  = CAST(l.BillingLagDays
                            + b.TotalAR * w.Days12 / NULLIF(w.Sales12, 0) AS DECIMAL(9,2)),
        -- what a day of lag is worth, so the operations fix can be priced
        CashPerLagDay  = CAST(w.Sales12 / NULLIF(w.Days12, 0) AS DECIMAL(14,2))
    FROM Bal b CROSS JOIN Win w CROSS JOIN Terms t CROSS JOIN Lag l
);
GO

CREATE VIEW dbo.vw_DSOBridge AS SELECT * FROM dbo.fn_DSOBridge('2025-12-31');
GO

CREATE VIEW dbo.vw_DSOBridgeMonthly AS
SELECT br.*,
       IsComparablePeriod = CAST(CASE WHEN m.MonthEndDate >= '2024-06-30' THEN 1 ELSE 0 END AS BIT)
FROM (SELECT DISTINCT MonthEndDate FROM dbo.Dim_Date
      WHERE MonthEndDate BETWEEN '2024-02-29' AND '2025-12-31') m
CROSS APPLY dbo.fn_DSOBridge(m.MonthEndDate) br;
GO

DECLARE @missing VARCHAR(400) = '';
IF OBJECT_ID('dbo.fn_WeightedAvgTerms', 'IF') IS NULL SET @missing += 'fn_WeightedAvgTerms ';
IF OBJECT_ID('dbo.fn_BillingLag', 'IF')       IS NULL SET @missing += 'fn_BillingLag ';
IF OBJECT_ID('dbo.vw_BillingLagMonthly', 'V') IS NULL SET @missing += 'vw_BillingLagMonthly ';
IF OBJECT_ID('dbo.vw_TermsMixMonthly', 'V')   IS NULL SET @missing += 'vw_TermsMixMonthly ';
IF OBJECT_ID('dbo.fn_WATShiftShare', 'IF')    IS NULL SET @missing += 'fn_WATShiftShare ';
IF OBJECT_ID('dbo.fn_DBT', 'IF')              IS NULL SET @missing += 'fn_DBT ';
IF OBJECT_ID('dbo.fn_DSOBridge', 'IF')        IS NULL SET @missing += 'fn_DSOBridge ';
IF OBJECT_ID('dbo.vw_DSOBridge', 'V')         IS NULL SET @missing += 'vw_DSOBridge ';
IF OBJECT_ID('dbo.vw_DSOBridgeMonthly', 'V')  IS NULL SET @missing += 'vw_DSOBridgeMonthly ';
IF @missing <> '' THROW 50004, 'FAILED to create DSO bridge objects -- scroll up for the compile error.', 1;
PRINT 'DSO bridge objects created and verified.';
GO
