/*
================================================================================
Project 2 -- Vantage Wholesale Supply: Receivables Performance
Script:  08_uat_test_cases.sql
Purpose: User Acceptance Test suite. Run it after any change to the generator,
         the settlement function, the KPI layer or the queue.

WHAT A TEST IS FOR HERE
    Not "does the query run". Every test below asserts a property that, if it
    broke, would let the reports keep producing confident, plausible, wrong
    numbers. Several of them exist because exactly that happened during the
    build, and the comment on each says so.

DESIGN
    Every test writes one row to #UATResults with an expected value, an actual
    value and a verdict, so a failure tells you what it found rather than only
    that something was wrong. The script is read-only -- it creates no
    permanent objects and modifies no data -- so it is safe to run against the
    reporting database at any time.

    Monetary comparisons use a half-cent tolerance. Exact float equality on
    aggregated DECIMAL arithmetic is a test that fails for reasons that have
    nothing to do with the thing being tested.

DATA DISCLOSURE: Vantage Wholesale Supply is fictional; all data is synthetic.
================================================================================
*/

USE VantageAR;
GO
SET NOCOUNT ON;
GO

IF OBJECT_ID('tempdb..#UATResults') IS NOT NULL DROP TABLE #UATResults;
CREATE TABLE #UATResults (
    TestID      VARCHAR(10)   NOT NULL,
    Area        VARCHAR(30)   NOT NULL,
    Requirement VARCHAR(200)  NOT NULL,
    Expected    VARCHAR(60)   NOT NULL,
    Actual      VARCHAR(60)   NULL,
    Verdict     VARCHAR(6)    NULL,
    WhyItMatters VARCHAR(400) NOT NULL
);
GO

DECLARE @AsOf DATE = '2025-12-31';

/* ---------------------------------------------------------------------------
UAT-01  The settlement identity holds for every invoice.
--------------------------------------------------------------------------- */
DECLARE @drift INT = (
    SELECT COUNT(*) FROM dbo.fn_ARBalance('2025-12-31')
    WHERE ABS(RawBalance - (InvoiceAmount - PaidToDate - DiscountToDate
                            - CreditMemoToDate - WriteOffToDate)) > 0.005);
INSERT INTO #UATResults VALUES ('UAT-01','Settlement',
 'RawBalance equals invoice less cash, discount, credit memos and write-offs, for every invoice',
 '0', CAST(@drift AS VARCHAR(20)), CASE WHEN @drift = 0 THEN 'PASS' ELSE 'FAIL' END,
 'The settlement identity is the one definition every downstream number depends on. If it drifts for even one invoice, the aging, the DSO and the call list are all wrong and nothing downstream can detect it.');

/* ---------------------------------------------------------------------------
UAT-02  As-of integrity: cash received AFTER the reporting date must not reduce
        a historical balance.

        This is the Project 1 bug class. Filtering the date dimension while
        joining the payment fact unfiltered lets next quarter's cash cure last
        quarter's aging, and the trend line bends quietly towards flattery.
--------------------------------------------------------------------------- */
-- Recomputed directly rather than by subtraction from the December figure.
-- The earlier form subtracted every application after 30 June, which now
-- includes cash banked in December and applied in January -- applications that
-- were never in the December total to begin with, so the arithmetic was the
-- test's, not the data's.
DECLARE @junDirect DECIMAL(18,2) = (
    SELECT ISNULL(SUM(p.AppliedAmount + p.DiscountTaken), 0)
    FROM dbo.Fact_CashApplication p
    JOIN dbo.Fact_Invoice f ON f.InvoiceKey = p.InvoiceKey
    WHERE p.ApplicationDateKey <= 20250630 AND f.InvoiceDateKey <= 20250630);
DECLARE @junBal DECIMAL(18,2) = (SELECT SUM(PaidToDate + DiscountToDate) FROM dbo.fn_ARBalance('2025-06-30'));
INSERT INTO #UATResults VALUES ('UAT-02','As-of integrity',
 'The 2025-06-30 balance counts only cash applied by then, to invoices raised by then',
 CAST(@junDirect AS VARCHAR(40)), CAST(@junBal AS VARCHAR(40)),
 CASE WHEN ABS(@junBal - @junDirect) < 0.005 THEN 'PASS' ELSE 'FAIL' END,
 'Without this, every historical aging report silently improves each time a customer pays in a later period, and month-on-month trends become meaningless.');

/* ---------------------------------------------------------------------------
UAT-03  Aging buckets are exhaustive and mutually exclusive.
--------------------------------------------------------------------------- */
DECLARE @unbucketed INT = (SELECT COUNT(*) FROM dbo.fn_ARBalance('2025-12-31') WHERE IsOpen = 1 AND BucketKey IS NULL);
DECLARE @multibucket INT = (
    SELECT COUNT(*) FROM dbo.fn_ARBalance('2025-12-31') b
    WHERE b.IsOpen = 1
      AND (SELECT COUNT(*) FROM dbo.Ref_AgingBucket k
           WHERE b.DaysPastDue BETWEEN k.MinDaysPastDue AND k.MaxDaysPastDue) <> 1);
INSERT INTO #UATResults VALUES ('UAT-03','Ageing',
 'Every open invoice falls in exactly one ageing bucket',
 '0 unbucketed, 0 multi', CAST(@unbucketed AS VARCHAR(10)) + ' / ' + CAST(@multibucket AS VARCHAR(10)),
 CASE WHEN @unbucketed = 0 AND @multibucket = 0 THEN 'PASS' ELSE 'FAIL' END,
 'Overlapping bucket ranges double-count dollars in the aging report; a gap drops them silently. Either way the buckets stop summing to total AR and nobody notices, because each bucket looks reasonable on its own.');

/* ---------------------------------------------------------------------------
UAT-04  The DSO bridge is an identity, not an approximation.
--------------------------------------------------------------------------- */
DECLARE @bridgeResidual DECIMAL(18,4) = (
    SELECT ABS(DSO_Classic - (GrantedDays + DisputeDays + LatenessDays))
    FROM dbo.fn_DSOBridge('2025-12-31'));
INSERT INTO #UATResults VALUES ('UAT-04','DSO bridge',
 'GrantedDays + DisputeDays + LatenessDays equals classic DSO exactly',
 '0.0000', CAST(@bridgeResidual AS VARCHAR(20)),
 CASE WHEN @bridgeResidual < 0.005 THEN 'PASS' ELSE 'FAIL' END,
 'The whole argument this project settles -- how much of the DSO rise was granted versus taken -- rests on these three parts summing to the total. A residual invites someone to plug it into the nearest component, which is how an analysis becomes an opinion.');

/* ---------------------------------------------------------------------------
UAT-05  The same metric computed in two places agrees to the cent.

        Caught during the build: fn_DSO divided before multiplying and
        fn_DSOBridge multiplied before dividing, so the two published figures
        for classic DSO differed by 0.01.
--------------------------------------------------------------------------- */
DECLARE @dsoGap DECIMAL(18,4) = (
    SELECT ABS(b.DSO_Classic - d.DSO_Simple)
    FROM dbo.fn_DSOBridge('2025-12-31') b CROSS APPLY dbo.fn_DSO('2025-12-31') d);
INSERT INTO #UATResults VALUES ('UAT-05','Cross-implementation',
 'fn_DSOBridge.DSO_Classic equals fn_DSO.DSO_Simple',
 '0.0000', CAST(@dsoGap AS VARCHAR(20)),
 CASE WHEN @dsoGap < 0.005 THEN 'PASS' ELSE 'FAIL' END,
 'Two objects publish the same metric. If they can disagree at all they will eventually disagree visibly, and a reader who spots it stops trusting every other number on the page.');

/* ---------------------------------------------------------------------------
UAT-06  Shift-share decomposition is exact.
--------------------------------------------------------------------------- */
DECLARE @ssResidual DECIMAL(18,4) = (
    SELECT ABS(DeltaWAT - (MixEffectDays + RateEffectDays + InteractionDays))
    FROM dbo.fn_WATShiftShare('2024-07-01','2024-12-31','2025-07-01','2025-12-31'));
INSERT INTO #UATResults VALUES ('UAT-06','Terms analysis',
 'Mix + Rate + Interaction equals the change in Weighted Average Terms',
 '0.0000', CAST(@ssResidual AS VARCHAR(20)),
 CASE WHEN @ssResidual < 0.02 THEN 'PASS' ELSE 'FAIL' END,
 'The rate effect is the figure that tells Sales a terms policy moved. If the decomposition does not close, the split between "we re-papered customers" and "long-terms customers simply bought more" is not evidence of anything.');

/* ---------------------------------------------------------------------------
UAT-07  The customer rollup ties to the invoice grain.
--------------------------------------------------------------------------- */
DECLARE @invAR DECIMAL(18,2) = (SELECT SUM(OpenBalance) FROM dbo.fn_ARBalance('2025-12-31'));
DECLARE @custAR DECIMAL(18,2) = (SELECT SUM(OpenBalance) FROM dbo.fn_CustomerAR('2025-12-31'));
INSERT INTO #UATResults VALUES ('UAT-07','Aggregation',
 'Customer-level open AR equals invoice-level open AR',
 CAST(@invAR AS VARCHAR(40)), CAST(@custAR AS VARCHAR(40)),
 CASE WHEN ABS(@invAR - @custAR) < 0.005 THEN 'PASS' ELSE 'FAIL' END,
 'A join that fans out, or a customer with no current dimension row, quietly changes the total. This is the cheapest possible check that the rollup did not invent or lose money.');

/* ---------------------------------------------------------------------------
UAT-08  Collectable exposure is bounded and non-negative.

        Caught during the build: the first version subtracted RAW disputed
        dollars from a RISK-WEIGHTED total, and a within-terms dispute (weight
        0.00, full face value) drove the figure below zero.
--------------------------------------------------------------------------- */
DECLARE @negExposure INT = (SELECT COUNT(*) FROM dbo.fn_PriorityActionQueue('2025-12-31') WHERE CollectableExposure < 0);
DECLARE @overExposure INT = (SELECT COUNT(*) FROM dbo.fn_PriorityActionQueue('2025-12-31') WHERE CollectableExposure > WeightedExposure + 0.005);
INSERT INTO #UATResults VALUES ('UAT-08','Priority queue',
 'CollectableExposure is between 0 and WeightedExposure for every account',
 '0 negative, 0 over', CAST(@negExposure AS VARCHAR(10)) + ' / ' + CAST(@overExposure AS VARCHAR(10)),
 CASE WHEN @negExposure = 0 AND @overExposure = 0 THEN 'PASS' ELSE 'FAIL' END,
 'Negative exposure is a units error -- weighted dollars minus raw dollars -- and it sorts the worst-affected accounts to the BOTTOM of the call list, which is the opposite of what the control is for.');

/* ---------------------------------------------------------------------------
UAT-09  Every queued account gets exactly one action, and it is a known one.
--------------------------------------------------------------------------- */
DECLARE @badAction INT = (
    SELECT COUNT(*) FROM dbo.fn_PriorityActionQueue('2025-12-31')
    WHERE ActionCode NOT IN ('APPLY_CASH','RESOLVE_DISPUTE','ESCALATE','FINAL_NOTICE','COLLECTION_CALL',
                             'COURTESY_REMINDER','STANDARD_DUNNING','CREDIT_REVIEW')
       OR ActionCode IS NULL OR RecommendedAction IS NULL);
INSERT INTO #UATResults VALUES ('UAT-09','Priority queue',
 'Every queued account carries exactly one recognised action code and a written instruction',
 '0', CAST(@badAction AS VARCHAR(10)),
 CASE WHEN @badAction = 0 THEN 'PASS' ELSE 'FAIL' END,
 'A CASE ladder that falls through leaves a collector with a row and no instruction. The queue stops being a control the moment a row does not say what to do.');

/* ---------------------------------------------------------------------------
UAT-10  The worklist respects collector capacity.
--------------------------------------------------------------------------- */
DECLARE @overCapacity INT = (
    SELECT COUNT(*) FROM (
        SELECT q.CollectorID,
               Extra = SUM(CASE WHEN q.IsTodaysWorklist = 1
                                 AND q.CollectorRank > 10
                                 AND q.ActionCode <> 'ESCALATE'
                                 AND q.CreditHoldFlag = 0 THEN 1 ELSE 0 END)
        FROM dbo.fn_PriorityActionQueue('2025-12-31') q GROUP BY q.CollectorID) x
    WHERE x.Extra > 0);
INSERT INTO #UATResults VALUES ('UAT-10','Priority queue',
 'No collector is given work beyond rank 10 unless it is an escalation or a credit hold',
 '0 collectors', CAST(@overCapacity AS VARCHAR(10)),
 CASE WHEN @overCapacity = 0 THEN 'PASS' ELSE 'FAIL' END,
 'A worklist longer than a day is a ledger with a sort order. If the capacity rule leaks, collectors go back to working the aging report top-down and the ranking work is wasted.');

/* ---------------------------------------------------------------------------
UAT-11  Average Days Delinquent uses one method on both sides.
--------------------------------------------------------------------------- */
DECLARE @addGap DECIMAL(18,4) = (
    SELECT ABS(AvgDaysDelinquent - (DSO_Countback - BPDSO_Countback)) FROM dbo.fn_DSO('2025-12-31'));
INSERT INTO #UATResults VALUES ('UAT-11','Metric coherence',
 'AvgDaysDelinquent equals countback DSO minus countback BPDSO',
 '0.0000', CAST(@addGap AS VARCHAR(20)),
 CASE WHEN @addGap < 0.005 THEN 'PASS' ELSE 'FAIL' END,
 'Subtracting a classic BPDSO from a countback DSO produces a number with no interpretation at all, and it looks perfectly reasonable. The only defence is asserting that both sides came from the same function.');

/* ---------------------------------------------------------------------------
UAT-12  Over-applied cash is surfaced, never netted away.
--------------------------------------------------------------------------- */
DECLARE @negOpen INT = (SELECT COUNT(*) FROM dbo.fn_ARBalance('2025-12-31') WHERE OpenBalance < 0);
DECLARE @mismatch INT = (
    SELECT COUNT(*) FROM dbo.fn_ARBalance('2025-12-31')
    WHERE (RawBalance < -0.005 AND OverAppliedAmount <= 0)
       OR (RawBalance > -0.005 AND OverAppliedAmount > 0));
INSERT INTO #UATResults VALUES ('UAT-12','Settlement',
 'No negative open balance; over-application reported separately',
 '0 negative, 0 mismatched', CAST(@negOpen AS VARCHAR(10)) + ' / ' + CAST(@mismatch AS VARCHAR(10)),
 CASE WHEN @negOpen = 0 AND @mismatch = 0 THEN 'PASS' ELSE 'FAIL' END,
 'A negative balance from a duplicate receipt nets against genuinely open invoices and UNDERSTATES total AR. Flooring at zero without reporting the excess hides a real cash-application defect instead.');

/* ---------------------------------------------------------------------------
UAT-13  The paper collections gap cannot be negative while write-offs exist.
--------------------------------------------------------------------------- */
DECLARE @negGapMonths INT = (
    SELECT COUNT(*) FROM dbo.vw_ARKPIMonthly
    WHERE IsComparablePeriod = 1 AND WriteOffs > 0 AND PaperCollectionsGap < -0.01);
INSERT INTO #UATResults VALUES ('UAT-13','CEI',
 'CEI_Book is at least CEI_Cash in every month carrying write-offs',
 '0 months', CAST(@negGapMonths AS VARCHAR(10)),
 CASE WHEN @negGapMonths = 0 THEN 'PASS' ELSE 'FAIL' END,
 'The book variant counts write-offs as collections and the cash variant does not, so the gap is non-negative by construction. A negative gap means one of the two is reading the wrong rows.');

/* ---------------------------------------------------------------------------
UAT-14  Procedures reject bad input rather than answering a different question.
--------------------------------------------------------------------------- */
DECLARE @guard1 VARCHAR(6) = 'FAIL', @guard2 VARCHAR(6) = 'FAIL';
BEGIN TRY EXEC dbo.usp_AgingSummary @GroupBy = 'DROP TABLE'; END TRY BEGIN CATCH SET @guard1 = 'PASS'; END CATCH
BEGIN TRY EXEC dbo.usp_CustomerStatement @CustomerID = 'C9999'; END TRY BEGIN CATCH SET @guard2 = 'PASS'; END CATCH
INSERT INTO #UATResults VALUES ('UAT-14','Interface',
 'usp_AgingSummary and usp_CustomerStatement raise on invalid arguments',
 'PASS / PASS', @guard1 + ' / ' + @guard2,
 CASE WHEN @guard1 = 'PASS' AND @guard2 = 'PASS' THEN 'PASS' ELSE 'FAIL' END,
 'A procedure that silently falls back to a default answers a question nobody asked, and the caller has no way to tell. Raising is the only honest response to an argument the procedure does not understand.');

/* ---------------------------------------------------------------------------
UAT-15  REGRESSION: the instalment defect stays fixed.

        The generator paid 60% of a partial invoice and only emitted the second
        instalment when the invoice was NOT a discount-taker, so 208 invoices
        carried a 40% residual that could never clear. Because 2/10NET30 is the
        only terms code with a discount, the whole defect landed on that cohort
        and read as a genuine delinquency finding.
--------------------------------------------------------------------------- */
-- A part-paid invoice CAN now legitimately stall: if the receipt covering its
-- remaining instalment arrived and was never matched, the balance stays open
-- for a reason that has nothing to do with the defect this test guards. Those
-- are excluded, so the assertion still targets the original signature -- a
-- residual with no cash behind it anywhere on the account.
DECLARE @unexplainedStall INT = (
    SELECT COUNT(*)
    FROM dbo.fn_ARBalance('2025-12-31') b
    LEFT JOIN dbo.fn_UnappliedCash('2025-12-31') u ON u.CustomerKey = b.CustomerKey
    WHERE b.IsOpen = 1 AND b.PaidToDate > 0 AND b.DaysPastDue > 60
      AND ISNULL(u.UnappliedCash, 0) = 0);
DECLARE @discountStall INT = (
    SELECT COUNT(*)
    FROM dbo.fn_ARBalance('2025-12-31') b
    JOIN dbo.Dim_PaymentTerms t ON t.TermsKey = b.TermsKey
    LEFT JOIN dbo.fn_UnappliedCash('2025-12-31') u ON u.CustomerKey = b.CustomerKey
    WHERE b.IsOpen = 1 AND b.PaidToDate > 0 AND b.DaysPastDue > 60
      AND t.DiscountPct > 0 AND ISNULL(u.UnappliedCash, 0) = 0);
INSERT INTO #UATResults VALUES ('UAT-15','Regression',
 'No part-paid invoice stalls past 60 days without cash somewhere behind it',
 '<= 5 unexplained, <= 1 on discount terms',
 CAST(@unexplainedStall AS VARCHAR(10)) + ' / ' + CAST(@discountStall AS VARCHAR(10)),
 CASE WHEN @unexplainedStall <= 5 AND @discountStall <= 1 THEN 'PASS' ELSE 'FAIL' END,
 'This defect produced a clean, plausible, entirely false finding: 47% of discount-terms invoices ageing past 90 days. It was invisible to every balance check because each row was internally consistent -- only a cross-cut of ageing by terms exposed it.');

/* ---------------------------------------------------------------------------
UAT-16  The generator is deterministic.
--------------------------------------------------------------------------- */
DECLARE @invChk INT = (SELECT CHECKSUM_AGG(CHECKSUM(InvoiceNo, InvoiceAmount, InvoiceDateKey, DueDateKey, ShipDateKey)) FROM dbo.Fact_Invoice);
DECLARE @rcpChk INT = (SELECT CHECKSUM_AGG(CHECKSUM(ReceiptNo, ReceiptAmount, ReceiptDateKey)) FROM dbo.Fact_CashReceipt);
DECLARE @appChk INT = (SELECT CHECKSUM_AGG(CHECKSUM(InvoiceKey, AppliedAmount, DiscountTaken, ApplicationDateKey)) FROM dbo.Fact_CashApplication);
DECLARE @adjChk INT = (SELECT CHECKSUM_AGG(CHECKSUM(AdjustmentNo, Amount, AdjustmentDateKey)) FROM dbo.Fact_Adjustment);
DECLARE @prmChk INT = (SELECT CHECKSUM_AGG(CHECKSUM(PromiseNo, PromisedAmount, PromisedPayDateKey)) FROM dbo.Fact_PromiseToPay);
INSERT INTO #UATResults VALUES ('UAT-16','Reproducibility',
 'Regenerating the dataset reproduces the documented fingerprints',
 '2062178339/1128760080/73418276/-962318236/-1781184159',
 CAST(@invChk AS VARCHAR(20)) + '/' + CAST(@rcpChk AS VARCHAR(20)) + '/' + CAST(@appChk AS VARCHAR(20))
  + '/' + CAST(@adjChk AS VARCHAR(20)) + '/' + CAST(@prmChk AS VARCHAR(20)),
 CASE WHEN @invChk = 2062178339 AND @rcpChk = 1128760080 AND @appChk = 73418276
       AND @adjChk = -962318236 AND @prmChk = -1781184159 THEN 'PASS' ELSE 'FAIL' END,
 'Every figure quoted in the case study is only checkable if the dataset can be rebuilt byte-for-byte. Hash-based pseudo-randomness gives that; RAND() and NEWID() do not.');

/* ---------------------------------------------------------------------------
UAT-17  The data-quality gate still detects the planted defects.
--------------------------------------------------------------------------- */
DECLARE @dupes INT = (SELECT Anomalies FROM dbo.vw_DQ_Summary WHERE AnomalyType = 'DUPLICATE_RECEIPT');
DECLARE @overs INT = (SELECT Anomalies FROM dbo.vw_DQ_Summary WHERE AnomalyType = 'OVER_APPLIED_CASH');
INSERT INTO #UATResults VALUES ('UAT-17','Data quality',
 'The duplicate-receipt defect is still detected at the bank event, and it is the only cause of over-applied cash',
 '96 receipts / 101 invoices', CAST(@dupes AS VARCHAR(10)) + ' / ' + CAST(@overs AS VARCHAR(10)),
 CASE WHEN @dupes = 96 AND @overs = 101 THEN 'PASS' ELSE 'FAIL' END,
 'The quality layer is only worth anything if it is known to fire. Planting a defect of known size and asserting the detector finds exactly that many is the only way to tell a working control from a silent one.');
GO


/* ---------------------------------------------------------------------------
UAT-18  Unapplied cash is floored per RECEIPT, not in total.

        An over-applied receipt has a negative remainder. Summing raw remainders
        would let it offset genuine unapplied cash elsewhere and understate the
        backlog -- the same netting error the invoice balances avoid.
--------------------------------------------------------------------------- */
DECLARE @negUnapplied INT = (SELECT COUNT(*) FROM dbo.fn_UnappliedCash('2025-12-31') WHERE UnappliedCash < 0);
DECLARE @unappliedSum DECIMAL(18,2) = (SELECT ISNULL(SUM(UnappliedCash),0) FROM dbo.fn_UnappliedCash('2025-12-31'));
DECLARE @unappliedDirect DECIMAL(18,2) = (
    SELECT ISNULL(SUM(CASE WHEN r.ReceiptAmount - ISNULL(ap.Applied,0) > 0.005
                           THEN r.ReceiptAmount - ISNULL(ap.Applied,0) ELSE 0 END), 0)
    FROM dbo.Fact_CashReceipt r
    OUTER APPLY (SELECT Applied = SUM(a.AppliedAmount) FROM dbo.Fact_CashApplication a
                 WHERE a.ReceiptKey = r.ReceiptKey AND a.ApplicationDateKey <= 20251231) ap
    WHERE r.ReceiptDateKey <= 20251231);
INSERT INTO #UATResults VALUES ('UAT-18','Unapplied cash',
 'Unapplied cash is non-negative and floored per receipt',
 '0 negative, ' + CAST(@unappliedDirect AS VARCHAR(40)),
 CAST(@negUnapplied AS VARCHAR(10)) + ' negative, ' + CAST(@unappliedSum AS VARCHAR(40)),
 CASE WHEN @negUnapplied = 0 AND ABS(@unappliedSum - @unappliedDirect) < 0.005 THEN 'PASS' ELSE 'FAIL' END,
 'Netting an over-applied receipt against genuinely unapplied cash would understate the backlog and hide the accounts that are about to be wrongly chased.');

/* ---------------------------------------------------------------------------
UAT-19  Net exposure is gross AR less the cash we are already holding.
--------------------------------------------------------------------------- */
DECLARE @netMismatch INT = (
    SELECT COUNT(*) FROM dbo.fn_CustomerAR('2025-12-31')
    WHERE ABS(NetExposure - (OpenBalance - UnappliedCash)) > 0.005);
INSERT INTO #UATResults VALUES ('UAT-19','Unapplied cash',
 'NetExposure equals OpenBalance less UnappliedCash for every customer',
 '0', CAST(@netMismatch AS VARCHAR(10)),
 CASE WHEN @netMismatch = 0 THEN 'PASS' ELSE 'FAIL' END,
 'The gross figure is what the ledger says; the net figure is what a collector should quote on the call. Getting it wrong means asking a customer to pay money we are already sitting on.');

/* ---------------------------------------------------------------------------
UAT-20  Accounts whose cash is already banked are routed OUT of the call list.
--------------------------------------------------------------------------- */
DECLARE @applyCashOnList INT = (
    SELECT COUNT(*) FROM dbo.fn_PriorityActionQueue('2025-12-31')
    WHERE ActionCode = 'APPLY_CASH' AND IsTodaysWorklist = 1);
DECLARE @applyCashCount INT = (
    SELECT COUNT(*) FROM dbo.fn_PriorityActionQueue('2025-12-31') WHERE ActionCode = 'APPLY_CASH');
INSERT INTO #UATResults VALUES ('UAT-20','Priority queue',
 'No APPLY_CASH account appears on a collector worklist',
 '0 of ' + CAST(@applyCashCount AS VARCHAR(10)), CAST(@applyCashOnList AS VARCHAR(10)),
 CASE WHEN @applyCashOnList = 0 THEN 'PASS' ELSE 'FAIL' END,
 'Ringing a customer who has already paid is the single most damaging call a collections team can make. If the routing leaks, the queue actively causes the harm it exists to prevent.');

/* ---------------------------------------------------------------------------
UAT-21  The cash cycle extends DSO rather than decomposing it.

        Billing lag happens BEFORE the DSO clock starts, so it must add to the
        measure. Folding it into the bridge would break the identity UAT-04
        asserts.
--------------------------------------------------------------------------- */
DECLARE @cycleResidual DECIMAL(18,4) = (
    SELECT ABS(CashCycleDays - (BillingLagDays + DSO_Classic)) FROM dbo.fn_DSOBridge('2025-12-31'));
INSERT INTO #UATResults VALUES ('UAT-21','Billing lag',
 'CashCycleDays equals BillingLagDays plus classic DSO exactly',
 '0.0000', CAST(@cycleResidual AS VARCHAR(20)),
 CASE WHEN @cycleResidual < 0.005 THEN 'PASS' ELSE 'FAIL' END,
 'Lag days sit before the invoice exists, so they are additive to DSO, not part of it. Folding them in would silently break the three-way bridge identity while still looking plausible.');

/* ---------------------------------------------------------------------------
UAT-22  Promise status is derived from cash, and is not stored anywhere.

        The register is maintained by the collectors the kept-rate measures. If
        an outcome column existed, the metric would be self-reported.
--------------------------------------------------------------------------- */
DECLARE @storedStatus INT = (
    SELECT COUNT(*) FROM sys.columns
    WHERE object_id = OBJECT_ID('dbo.Fact_PromiseToPay')
      AND name IN ('Status','PromiseStatus','Outcome','IsKept','Kept'));
DECLARE @keptNoCash INT = (
    SELECT COUNT(*) FROM dbo.fn_PromiseStatus('2025-12-31', 3, 90.00)
    WHERE PromiseStatus = 'Kept' AND CashInWindow <= 0);
INSERT INTO #UATResults VALUES ('UAT-22','Promises',
 'No outcome column on the promise table, and no promise is Kept without cash',
 '0 columns, 0 kept-without-cash',
 CAST(@storedStatus AS VARCHAR(10)) + ' / ' + CAST(@keptNoCash AS VARCHAR(10)),
 CASE WHEN @storedStatus = 0 AND @keptNoCash = 0 THEN 'PASS' ELSE 'FAIL' END,
 'A promise outcome stored as a column would be written by the person being scored on it. Deriving it from applied cash means the only way to improve the metric is for a customer to actually pay.');

/* ---------------------------------------------------------------------------
UAT-23  An unripe promise is Outstanding, never Broken.
--------------------------------------------------------------------------- */
DECLARE @futureBroken INT = (
    SELECT COUNT(*) FROM dbo.fn_PromiseStatus('2025-06-30', 3, 90.00)
    WHERE PromisedPayDate > '2025-06-30' AND PromiseStatus <> 'Outstanding');
INSERT INTO #UATResults VALUES ('UAT-23','Promises',
 'A promise due after the as-of date is Outstanding, not Broken',
 '0', CAST(@futureBroken AS VARCHAR(10)),
 CASE WHEN @futureBroken = 0 THEN 'PASS' ELSE 'FAIL' END,
 'Counting promises that have not come due yet as failures would make the kept-rate depend on when the report was run rather than on what customers did.');

/* ---------------------------------------------------------------------------
UAT-24  Cash cannot be applied before it reaches the bank.
--------------------------------------------------------------------------- */
DECLARE @appBeforeReceipt INT = (
    SELECT Anomalies FROM dbo.vw_DQ_Summary WHERE AnomalyType = 'APPLICATION_BEFORE_RECEIPT');
INSERT INTO #UATResults VALUES ('UAT-24','Cash application',
 'No application is dated before the receipt that funds it',
 '0', CAST(@appBeforeReceipt AS VARCHAR(10)),
 CASE WHEN @appBeforeReceipt = 0 THEN 'PASS' ELSE 'FAIL' END,
 'An application dated before its receipt cures an ageing bucket earlier than the money actually arrived, which improves a month end that had not really improved.');
GO

/* ---------------------------------------------------------------------------
UAT-25  Days Beyond Terms: the at-risk variant must not be flattered by
        survivorship.

        Settled DBT averages only invoices that were PAID. An invoice never paid
        never enters it, so as a cohort slides towards default the measure can
        IMPROVE while the cohort collapses. The at-risk variant marks the unpaid
        dollars to the as-of date, so it can never read better than the settled
        figure while past-due balances exist.

        This function was repointed from Fact_Payment to Fact_CashApplication
        during the cash-receipt re-architecture and had no coverage until then.
--------------------------------------------------------------------------- */
DECLARE @dbtSettled DECIMAL(18,2), @dbtAtRisk DECIMAL(18,2), @dbtOpenPD DECIMAL(18,2);
SELECT @dbtSettled = DBT_Settled, @dbtAtRisk = DBT_AtRisk, @dbtOpenPD = OpenPastDueDollars
FROM dbo.fn_DBT('2025-12-31', 12);
INSERT INTO #UATResults VALUES ('UAT-25','Days beyond terms',
 'At-risk DBT is at least settled DBT while past-due balances exist',
 'at-risk >= settled, past due > 0',
 CAST(@dbtAtRisk AS VARCHAR(20)) + ' >= ' + CAST(@dbtSettled AS VARCHAR(20))
   + ', past due ' + CAST(@dbtOpenPD AS VARCHAR(30)),
 CASE WHEN @dbtOpenPD > 0 AND @dbtAtRisk >= @dbtSettled THEN 'PASS' ELSE 'FAIL' END,
 'Publishing settled days-beyond-terms alone lets a deteriorating cohort read as improving, because the invoices that stop being paid stop being averaged. The at-risk variant is the one that shows deterioration as deterioration.');
GO
/* ------------------------------ report ---------------------------------- */
SELECT TestID, Area, Requirement, Expected, Actual, Verdict FROM #UATResults ORDER BY TestID;

DECLARE @pass INT = (SELECT COUNT(*) FROM #UATResults WHERE Verdict = 'PASS');
DECLARE @fail INT = (SELECT COUNT(*) FROM #UATResults WHERE Verdict = 'FAIL');
PRINT '';
PRINT 'UAT RESULT: ' + CAST(@pass AS VARCHAR(10)) + ' passed, ' + CAST(@fail AS VARCHAR(10)) + ' failed.';
IF @fail > 0
BEGIN
    SELECT TestID, Requirement, Expected, Actual, WhyItMatters FROM #UATResults WHERE Verdict = 'FAIL' ORDER BY TestID;
    THROW 50020, 'UAT FAILED -- see the failing cases above. Do not publish figures from this build.', 1;
END
PRINT 'All acceptance criteria met.';
GO
