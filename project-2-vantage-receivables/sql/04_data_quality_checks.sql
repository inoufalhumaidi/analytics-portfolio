/*
================================================================================
Project 2 -- Vantage Wholesale Supply: Receivables Performance
Script:  04_data_quality_checks.sql
Purpose: Reusable data-quality control layer for the receivables ledger.
         Creates:
           dbo.fn_DQ_Anomalies(@AsOf)   row-level: one row per defect found
           dbo.vw_DQ_Anomalies          the reporting-date snapshot
           dbo.vw_DQ_Summary            counts + rates + money at risk
           dbo.usp_RunDataQualityChecks a pass/fail QA gate

These catch what CHECK constraints and foreign keys cannot: defects that are
only visible ACROSS rows or ACROSS tables -- a receipt posted twice, cash dated
before the invoice it settles, a due date that contradicts the customer's terms.

DETECTED BY BEHAVIOUR, NOT BY MARKER. The generator tags its planted duplicate
receipts with an 'RCT-D' prefix. These checks deliberately ignore that prefix
and match on the business signature instead (same customer, same date, same
amount), because a check that finds defects only by the label the generator
gave them proves nothing about a real feed.

EVERY CHECK REPORTS MONEY. A count answers "how many rows are wrong"; the
finance question is "how much cash is mis-stated", so each anomaly carries an
AmountAtRisk the summary can total.

DATA DISCLOSURE: Vantage Wholesale Supply is fictional; all data is synthetic.
================================================================================
*/

USE VantageAR;
GO

IF OBJECT_ID('dbo.usp_RunDataQualityChecks', 'P') IS NOT NULL DROP PROCEDURE dbo.usp_RunDataQualityChecks;
IF OBJECT_ID('dbo.vw_DQ_Summary', 'V')            IS NOT NULL DROP VIEW dbo.vw_DQ_Summary;
IF OBJECT_ID('dbo.vw_DQ_Anomalies', 'V')          IS NOT NULL DROP VIEW dbo.vw_DQ_Anomalies;
IF OBJECT_ID('dbo.fn_DQ_Anomalies', 'IF')         IS NOT NULL DROP FUNCTION dbo.fn_DQ_Anomalies;
IF OBJECT_ID('dbo.vw_DQ_CheckCatalog', 'V')       IS NOT NULL DROP VIEW dbo.vw_DQ_CheckCatalog;
GO

/*
--------------------------------------------------------------------------------
The catalog of checks that EXIST. The summary left-joins onto this so a check
that finds nothing still reports "0 found" rather than vanishing from the
report. Silence and success look identical otherwise.
--------------------------------------------------------------------------------
*/
/*
CountsTowardExposure exists to stop double counting. A duplicate receipt is the
CAUSE; the over-applied balance it produces is the EFFECT, and they are the same
money. Summing every check's AmountAtRisk would report that money twice and
overstate the exposure. Only OVER_APPLIED_CASH -- the net effect on the ledger --
counts toward the gate. The other checks remain fully reported: they mis-state
WHEN cash arrived or HOW OLD a balance is, not how much is owed.
*/
CREATE VIEW dbo.vw_DQ_CheckCatalog AS
-- EntityType is declared here, not inferred from the rows a check happens to
-- return. A check that finds nothing has no rows to infer from, and would
-- otherwise fall back to the invoice count and report a rate against the wrong
-- denominator -- correct only while it is failing.
SELECT AnomalyType, EntityType, Severity, ImpactClass, CountsTowardExposure, WhatItMeans
FROM (VALUES
 ('OVER_APPLIED_CASH','Invoice',        'High',   'Balance',   CAST(1 AS BIT), 'Cash, discounts and adjustments applied exceed the invoice value, leaving a credit balance. This is the net mis-statement of AR, so it alone counts toward exposure.'),
 ('DUPLICATE_RECEIPT','Receipt',        'High',   'Balance',   CAST(0 AS BIT), 'The same cash receipt was banked twice (same customer, date and amount). A cause of over-application; its money is counted there.'),
 ('ADJUSTMENT_EXCEEDS_INVOICE','Invoice','High',  'Balance',   CAST(0 AS BIT), 'Credit memos and write-offs together exceed the invoice value. Also a cause of over-application.'),
 ('RECEIPT_BEFORE_INVOICE','Receipt',   'High',   'Timing',    CAST(0 AS BIT), 'A receipt is dated before the invoice it settles. Corrupts days-to-pay and DSO without changing the balance.'),
 ('APPLICATION_BEFORE_RECEIPT','Application','High',  'Timing',    CAST(0 AS BIT), 'Cash was matched to an invoice before it reached the bank. Impossible in fact, so it means the application date was keyed wrong -- and it cures an ageing bucket early.'),
 ('UNAPPLIED_CASH_AGED','Receipt',      'High',   'Routing',   CAST(0 AS BIT), 'A receipt has sat unmatched for more than 90 days. The customer has paid, the invoice still shows open, and the queue will send a collector to chase them.'),
 ('PRE_BILLING','Invoice',              'Medium', 'Integrity', CAST(0 AS BIT), 'The invoice is dated before the goods shipped. A negative billing lag is a revenue-recognition question, not a rounding artefact, so it is raised rather than clipped to zero.'),
 ('DUE_DATE_TERMS_MISMATCH','Invoice',  'Medium', 'Ageing',    CAST(0 AS BIT), 'The due date does not equal invoice date + the terms'' net days, so the invoice lands in the wrong ageing bucket.'),
 ('DISCOUNT_WITHOUT_TERMS','Application',   'Medium', 'Balance',   CAST(0 AS BIT), 'An early-payment discount was taken on terms that offer none.'),
 ('DISPUTE_FLAG_INCONSISTENT','Invoice','Medium', 'Integrity', CAST(0 AS BIT), 'Dispute flag, reason and dates contradict each other.'),
 ('RECEIPT_AFTER_ASOF','Receipt',       'Medium', 'Timing',    CAST(0 AS BIT), 'A receipt is dated after the reporting date, so the snapshot includes future cash.'),
 ('ORPHAN_DIMENSION_KEY','Invoice',     'High',   'Integrity', CAST(0 AS BIT), 'A fact row points at a customer, terms or date that does not exist.')
) AS c(AnomalyType, EntityType, Severity, ImpactClass, CountsTowardExposure, WhatItMeans);
GO

/*
--------------------------------------------------------------------------------
fn_DQ_Anomalies(@AsOf) -- one row per defect.
Parameterized so the same checks can be run against any reporting date, e.g.
to prove a defect was already present last month.
--------------------------------------------------------------------------------
*/
CREATE FUNCTION dbo.fn_DQ_Anomalies (@AsOf DATE)
RETURNS TABLE
AS RETURN
(
    -- surplus copies of a receipt: keep the first, flag the rest
    WITH AsOfKey AS (SELECT k = YEAR(@AsOf)*10000 + MONTH(@AsOf)*100 + DAY(@AsOf)),
    -- A duplicate is now detected at the BANK EVENT, not the application. Two
    -- receipts from one customer, on one day, for one amount is a re-keyed
    -- cheque -- and because the re-key carries its matching with it, the
    -- applications duplicate too and the invoice gets paid twice.
    DupRank AS (
        SELECT r.ReceiptKey, r.ReceiptNo, r.CustomerKey, r.ReceiptAmount, r.ReceiptDateKey,
               CopyNo = ROW_NUMBER() OVER (
                          PARTITION BY r.CustomerKey, r.ReceiptDateKey, r.ReceiptAmount
                          ORDER BY r.ReceiptKey)
        FROM dbo.Fact_CashReceipt r
        CROSS JOIN AsOfKey a
        WHERE r.ReceiptDateKey <= a.k
    )
    SELECT AnomalyType = 'DUPLICATE_RECEIPT',
           EntityType  = 'Receipt',
           EntityRef   = d.ReceiptNo,
           InvoiceKey  = CAST(NULL AS BIGINT),
           AmountAtRisk = d.ReceiptAmount,
           Detail = CONCAT('Receipt duplicates an earlier banking of ', FORMAT(d.ReceiptAmount, 'C', 'en-US'),
                           ' from the same customer on the same day')
    FROM DupRank d
    WHERE d.CopyNo > 1

    UNION ALL
    SELECT 'OVER_APPLIED_CASH', 'Invoice', b.InvoiceNo, b.InvoiceKey,
           b.OverAppliedAmount,
           CONCAT('Applied ', FORMAT(b.PaidToDate + b.DiscountToDate + b.CreditMemoToDate + b.WriteOffToDate, 'C', 'en-US'),
                  ' against an invoice of ', FORMAT(b.InvoiceAmount, 'C', 'en-US'))
    FROM dbo.fn_ARBalance(@AsOf) b
    WHERE b.OverAppliedAmount > 0.005

    UNION ALL
    SELECT 'RECEIPT_BEFORE_INVOICE', 'Receipt', r.ReceiptNo, a2.InvoiceKey,
           a2.AppliedAmount,
           CONCAT('Receipt banked ', CONVERT(CHAR(10), dp.[Date], 120),
                  ' but the invoice it settles is dated ', CONVERT(CHAR(10), di.[Date], 120))
    FROM dbo.Fact_CashReceipt r
    JOIN dbo.Fact_CashApplication a2 ON a2.ReceiptKey = r.ReceiptKey
    JOIN dbo.Fact_Invoice f ON f.InvoiceKey = a2.InvoiceKey
    JOIN dbo.Dim_Date dp ON dp.DateKey = r.ReceiptDateKey
    JOIN dbo.Dim_Date di ON di.DateKey = f.InvoiceDateKey
    CROSS JOIN AsOfKey a
    WHERE r.ReceiptDateKey < f.InvoiceDateKey AND r.ReceiptDateKey <= a.k

    UNION ALL
    -- Cash cannot be matched to an invoice before it reaches the bank. When it
    -- appears to have been, the application date is wrong -- and a wrong
    -- application date cures an ageing bucket earlier than the money did.
    SELECT 'APPLICATION_BEFORE_RECEIPT', 'Application',
           CONCAT(r.ReceiptNo, '/', a2.ApplicationKey), a2.InvoiceKey,
           a2.AppliedAmount,
           CONCAT('Applied ', CONVERT(CHAR(10), da.[Date], 120),
                  ' but the cash was not banked until ', CONVERT(CHAR(10), dr.[Date], 120))
    FROM dbo.Fact_CashApplication a2
    JOIN dbo.Fact_CashReceipt r ON r.ReceiptKey = a2.ReceiptKey
    JOIN dbo.Dim_Date da ON da.DateKey = a2.ApplicationDateKey
    JOIN dbo.Dim_Date dr ON dr.DateKey = r.ReceiptDateKey
    CROSS JOIN AsOfKey a
    WHERE a2.ApplicationDateKey < r.ReceiptDateKey AND a2.ApplicationDateKey <= a.k

    UNION ALL
    -- The routing defect: the customer has paid, the invoice still reads open,
    -- and the call list will happily send someone to chase them for it.
    SELECT 'UNAPPLIED_CASH_AGED', 'Receipt', c2.CustomerID, CAST(NULL AS BIGINT),
           u.UnappliedCash,
           CONCAT(FORMAT(u.UnappliedCash, 'C', 'en-US'), ' banked and unmatched across ',
                  u.UnappliedReceipts, ' receipt(s); oldest ', u.OldestUnappliedDays, ' days')
    FROM dbo.fn_UnappliedCash(@AsOf) u
    JOIN dbo.Dim_Customer c2 ON c2.CustomerKey = u.CustomerKey
    WHERE u.UnappliedCash > 0.005 AND u.OldestUnappliedDays > 90

    UNION ALL
    SELECT 'PRE_BILLING', 'Invoice', f.InvoiceNo, f.InvoiceKey,
           f.InvoiceAmount,
           CONCAT('Invoiced ', CONVERT(CHAR(10), di.[Date], 120),
                  ' but shipped ', CONVERT(CHAR(10), ds.[Date], 120),
                  ' -- billed ', DATEDIFF(DAY, di.[Date], ds.[Date]), ' days before despatch')
    FROM dbo.Fact_Invoice f
    JOIN dbo.Dim_Date di ON di.DateKey = f.InvoiceDateKey
    JOIN dbo.Dim_Date ds ON ds.DateKey = f.ShipDateKey
    CROSS JOIN AsOfKey a
    WHERE f.ShipDateKey > f.InvoiceDateKey AND f.InvoiceDateKey <= a.k

    UNION ALL
    SELECT 'ADJUSTMENT_EXCEEDS_INVOICE', 'Invoice', f.InvoiceNo, f.InvoiceKey,
           x.AdjTotal - f.InvoiceAmount,
           CONCAT('Adjustments total ', FORMAT(x.AdjTotal, 'C', 'en-US'),
                  ' against an invoice of ', FORMAT(f.InvoiceAmount, 'C', 'en-US'))
    FROM dbo.Fact_Invoice f
    CROSS JOIN AsOfKey a
    CROSS APPLY (SELECT AdjTotal = SUM(j.Amount)
                 FROM dbo.Fact_Adjustment j
                 WHERE j.InvoiceKey = f.InvoiceKey AND j.AdjustmentDateKey <= a.k) x
    WHERE x.AdjTotal > f.InvoiceAmount + 0.005

    UNION ALL
    SELECT 'DUE_DATE_TERMS_MISMATCH', 'Invoice', f.InvoiceNo, f.InvoiceKey,
           b.OpenBalance,
           CONCAT('Terms ', t.TermsCode, ' imply ', t.NetDays, ' days but the due date is ',
                  DATEDIFF(DAY, di.[Date], dd.[Date]), ' days out')
    FROM dbo.Fact_Invoice f
    JOIN dbo.Dim_PaymentTerms t ON t.TermsKey = f.TermsKey
    JOIN dbo.Dim_Date di ON di.DateKey = f.InvoiceDateKey
    JOIN dbo.Dim_Date dd ON dd.DateKey = f.DueDateKey
    JOIN dbo.fn_ARBalance(@AsOf) b ON b.InvoiceKey = f.InvoiceKey
    WHERE DATEDIFF(DAY, di.[Date], dd.[Date]) <> t.NetDays

    UNION ALL
    SELECT 'DISCOUNT_WITHOUT_TERMS', 'Application', CONCAT('APP-', p.ApplicationKey), p.InvoiceKey,
           p.DiscountTaken,
           CONCAT('Discount of ', FORMAT(p.DiscountTaken, 'C', 'en-US'),
                  ' taken on terms ', t.TermsCode, ', which offers none')
    FROM dbo.Fact_CashApplication p
    JOIN dbo.Fact_Invoice f ON f.InvoiceKey = p.InvoiceKey
    JOIN dbo.Dim_PaymentTerms t ON t.TermsKey = f.TermsKey
    CROSS JOIN AsOfKey a
    WHERE p.DiscountTaken > 0.005 AND t.DiscountPct = 0 AND p.ApplicationDateKey <= a.k

    UNION ALL
    SELECT 'DISPUTE_FLAG_INCONSISTENT', 'Invoice', f.InvoiceNo, f.InvoiceKey,
           b.OpenBalance,
           CASE
             WHEN f.IsDisputed = 1 AND f.DisputeOpenedKey IS NULL THEN 'Flagged disputed with no opened date'
             WHEN f.IsDisputed = 0 AND f.DisputeReason IS NOT NULL THEN 'Not flagged disputed but carries a reason'
             WHEN f.DisputeClosedKey < f.DisputeOpenedKey THEN 'Dispute closed before it was opened'
             ELSE 'Dispute dates inconsistent'
           END
    FROM dbo.Fact_Invoice f
    JOIN dbo.fn_ARBalance(@AsOf) b ON b.InvoiceKey = f.InvoiceKey
    WHERE (f.IsDisputed = 1 AND f.DisputeOpenedKey IS NULL)
       OR (f.IsDisputed = 0 AND f.DisputeReason IS NOT NULL)
       OR (f.DisputeClosedKey < f.DisputeOpenedKey)

    UNION ALL
    SELECT 'RECEIPT_AFTER_ASOF', 'Receipt', r.ReceiptNo, CAST(NULL AS BIGINT),
           r.ReceiptAmount,
           CONCAT('Receipt banked ', CONVERT(CHAR(10), dp.[Date], 120),
                  ', after the reporting date ', CONVERT(CHAR(10), @AsOf, 120))
    FROM dbo.Fact_CashReceipt r
    JOIN dbo.Dim_Date dp ON dp.DateKey = r.ReceiptDateKey
    CROSS JOIN AsOfKey a
    WHERE r.ReceiptDateKey > a.k

    UNION ALL
    -- defence in depth: foreign keys already prevent this, so a hit means the
    -- constraints themselves were dropped or bypassed by a bulk load
    SELECT 'ORPHAN_DIMENSION_KEY', 'Invoice', f.InvoiceNo, f.InvoiceKey,
           f.InvoiceAmount, 'Invoice references a missing customer, terms or date row'
    FROM dbo.Fact_Invoice f
    WHERE NOT EXISTS (SELECT 1 FROM dbo.Dim_Customer     c WHERE c.CustomerKey = f.CustomerKey)
       OR NOT EXISTS (SELECT 1 FROM dbo.Dim_PaymentTerms t WHERE t.TermsKey    = f.TermsKey)
       OR NOT EXISTS (SELECT 1 FROM dbo.Dim_Date         d WHERE d.DateKey     = f.InvoiceDateKey)
       OR NOT EXISTS (SELECT 1 FROM dbo.Dim_Date         d WHERE d.DateKey     = f.DueDateKey)
);
GO

CREATE VIEW dbo.vw_DQ_Anomalies AS
SELECT * FROM dbo.fn_DQ_Anomalies('2025-12-31');
GO

/*
--------------------------------------------------------------------------------
vw_DQ_Summary -- every catalogued check, whether or not it fired.
Rates are expressed against the population each check actually examines, so a
payment-level defect is not diluted by the invoice count.
--------------------------------------------------------------------------------
*/
CREATE VIEW dbo.vw_DQ_Summary AS
WITH Pop AS (
    SELECT InvoiceCount     = (SELECT COUNT(*) FROM dbo.Fact_Invoice),
           ReceiptCount     = (SELECT COUNT(*) FROM dbo.Fact_CashReceipt),
           ApplicationCount = (SELECT COUNT(*) FROM dbo.Fact_CashApplication)
),
Found AS (
    SELECT AnomalyType, EntityType,
           Anomalies = COUNT(*),
           AmountAtRisk = SUM(AmountAtRisk)
    FROM dbo.vw_DQ_Anomalies
    GROUP BY AnomalyType, EntityType
)
SELECT
    cat.AnomalyType,
    cat.EntityType,
    cat.Severity,
    cat.ImpactClass,
    cat.CountsTowardExposure,
    Anomalies      = ISNULL(f.Anomalies, 0),
    AmountAtRisk   = CAST(ISNULL(f.AmountAtRisk, 0) AS DECIMAL(14,2)),
    PopulationScanned = CASE cat.EntityType WHEN 'Receipt' THEN p.ReceiptCount WHEN 'Application' THEN p.ApplicationCount ELSE p.InvoiceCount END,
    AnomalyRatePct = CAST(100.0 * ISNULL(f.Anomalies, 0)
                     / NULLIF(CASE cat.EntityType WHEN 'Receipt' THEN p.ReceiptCount WHEN 'Application' THEN p.ApplicationCount ELSE p.InvoiceCount END, 0)
                     AS DECIMAL(6,3)),
    cat.WhatItMeans
FROM dbo.vw_DQ_CheckCatalog cat
LEFT JOIN Found f ON f.AnomalyType = cat.AnomalyType
CROSS JOIN Pop p;
GO

/*
--------------------------------------------------------------------------------
usp_RunDataQualityChecks -- the gate. Fails loudly rather than returning a
report nobody reads.
--------------------------------------------------------------------------------
*/
CREATE PROCEDURE dbo.usp_RunDataQualityChecks
    @MaxAcceptableRatePct    DECIMAL(6,3) = 1.000,  -- any single check above this share of its population fails
    @MaxMisstatementPctOfAR  DECIMAL(6,3) = 1.000   -- ...as does mis-stating more than this share of open AR
AS
BEGIN
    SET NOCOUNT ON;

    PRINT '=== Vantage Wholesale Supply: AR Data Quality Report (as of 2025-12-31) ===';

    SELECT AnomalyType, Severity, ImpactClass, Anomalies, AmountAtRisk,
           PopulationScanned, AnomalyRatePct, CountsTowardExposure,
           Result = CASE WHEN AnomalyRatePct > @MaxAcceptableRatePct THEN 'FAIL' ELSE 'PASS' END
    FROM dbo.vw_DQ_Summary
    ORDER BY CASE Severity WHEN 'High' THEN 1 ELSE 2 END, AnomalyRatePct DESC;

    DECLARE @failed INT, @found INT, @exposure DECIMAL(14,2), @openAR DECIMAL(14,2), @exposurePct DECIMAL(6,3);
    SELECT @failed = COUNT(*) FROM dbo.vw_DQ_Summary WHERE AnomalyRatePct > @MaxAcceptableRatePct;
    SELECT @found  = SUM(Anomalies) FROM dbo.vw_DQ_Summary;
    -- exposure counts the net effect only, never the cause and the effect together
    SELECT @exposure = SUM(AmountAtRisk) FROM dbo.vw_DQ_Summary WHERE CountsTowardExposure = 1;
    SELECT @openAR = SUM(OpenBalance) FROM dbo.vw_ARBalance;
    SET @exposurePct = CAST(100.0 * @exposure / NULLIF(@openAR, 0) AS DECIMAL(6,3));

    PRINT CONCAT('Defects found: ', @found,
                 ' | AR mis-stated by: ', FORMAT(@exposure, 'C', 'en-US'),
                 ' (', @exposurePct, '% of ', FORMAT(@openAR, 'C', 'en-US'), ' open AR)');

    IF @failed > 0
        PRINT CONCAT('QA GATE: FAIL -- ', @failed, ' check(s) exceed ', @MaxAcceptableRatePct, '% of their population. See dbo.vw_DQ_Anomalies.');
    ELSE IF @exposurePct > @MaxMisstatementPctOfAR
        PRINT CONCAT('QA GATE: FAIL -- AR mis-stated by ', @exposurePct, '%, above the ', @MaxMisstatementPctOfAR, '% tolerance. See dbo.vw_DQ_Anomalies.');
    ELSE
        PRINT 'QA GATE: PASS -- all checks within tolerance.';
END
GO

DECLARE @missing VARCHAR(400) = '';
IF OBJECT_ID('dbo.fn_DQ_Anomalies', 'IF')         IS NULL SET @missing += 'fn_DQ_Anomalies ';
IF OBJECT_ID('dbo.vw_DQ_Anomalies', 'V')          IS NULL SET @missing += 'vw_DQ_Anomalies ';
IF OBJECT_ID('dbo.vw_DQ_Summary', 'V')            IS NULL SET @missing += 'vw_DQ_Summary ';
IF OBJECT_ID('dbo.usp_RunDataQualityChecks', 'P') IS NULL SET @missing += 'usp_RunDataQualityChecks ';
IF @missing <> '' THROW 50002, 'FAILED to create data-quality objects -- scroll up for the compile error.', 1;
PRINT 'Data-quality objects created and verified.';
GO
