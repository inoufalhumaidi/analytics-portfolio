/*
================================================================================
Project 3 -- Lumen Optics Manufacturing: Photonics Spend Scorecard
Script:  04_data_quality_checks.sql
Purpose: Reusable data-quality control layer for the procurement ledger.
         Creates:
           dbo.vw_DQ_CheckCatalog        the checks that EXIST, with their scope
           dbo.fn_DQ_Anomalies(@AsOf)    row-level: one row per defect found
           dbo.vw_DQ_Anomalies           the reporting-date snapshot
           dbo.vw_DQ_Summary             counts, rates and money at risk
           dbo.usp_RunDataQualityChecks  a pass/fail QA gate

These catch what CHECK constraints and foreign keys cannot: defects only
visible ACROSS rows or ACROSS tables -- a requisition raised twice under two PO
numbers, goods booked in before the order existed, two price agreements in
force at once.

DETECTED BY BEHAVIOUR, NOT BY MARKER. The generator tags its planted duplicates
with a 'PO-D' prefix and its overlapping agreements with 'PA-X'. These checks
deliberately ignore both and match on the business signature instead, because a
check that finds defects only by the label the generator gave them proves
nothing about a real feed.

EVERY CHECK REPORTS MONEY. A count answers "how many rows are wrong"; the
sourcing question is "how much spend is mis-stated", so each anomaly carries an
AmountAtRisk the summary can total.

COUNTING EACH DEFECT ONCE. A duplicated PO line is the CAUSE; the overstated
committed spend it produces is the EFFECT, and they are the same dollars.
Only checks flagged CountsTowardExposure contribute to the gate, so the
headline cannot double count.

DATA DISCLOSURE: Lumen Optics Manufacturing is fictional; all data is synthetic.
================================================================================
*/

USE LumenSpend;
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
report -- silence and success look identical otherwise.

EntityType is declared HERE rather than inferred from the rows a check returns,
because a check that finds nothing has no rows to infer from and would report
its rate against whatever population happened to be the fallback.
--------------------------------------------------------------------------------
*/
CREATE VIEW dbo.vw_DQ_CheckCatalog AS
SELECT AnomalyType, EntityType, Severity, ImpactClass, CountsTowardExposure, WhatItMeans
FROM (VALUES
 ('DUPLICATE_PO_LINE',    'POLine',    'High',   'Spend',     CAST(1 AS BIT),
  'The same requisition raised twice under two PO numbers (same part, vendor, date, quantity and price). Overstates committed spend and double-counts volume in every price average.'),
 ('AMBIGUOUS_CONTRACT',   'POLine',    'High',   'Contract',  CAST(0 AS BIT),
  'Two or more price agreements were in force for this vendor-part at the order date, so the contracted price -- and therefore the purchase price variance -- depends on which row the query picks.'),
 ('PRICE_OUTLIER',        'POLine',    'High',   'Spend',     CAST(0 AS BIT),
  'Unit price is more than five times the agreement in force. Almost always a decimal-point slip on entry: obvious in a list, invisible in a total.'),
 ('RECEIPT_BEFORE_ORDER', 'Receipt',   'High',   'Timing',    CAST(0 AS BIT),
  'Goods booked in before the purchase order was raised. Corrupts lead time and on-time delivery without changing spend.'),
 ('RECEIPT_EXCEEDS_ORDER','Receipt',   'High',   'Quantity',  CAST(0 AS BIT),
  'Cumulative receipts exceed the quantity ordered on the line. Either an over-delivery nobody recorded, or a receipt posted against the wrong line.'),
 ('ORPHAN_DIMENSION_KEY', 'POLine',    'High',   'Integrity', CAST(0 AS BIT),
  'A fact row points at a part, vendor, buyer or date that does not exist.'),
 ('AGREEMENT_INVERTED',   'Agreement', 'Medium', 'Contract',  CAST(0 AS BIT),
  'A price agreement whose validity window ends before it starts.'),
 ('EXPEDITE_FLAG_MISMATCH','POLine',   'Medium', 'Integrity', CAST(0 AS BIT),
  'An expedite fee with no expedite flag, or the flag with no fee. One of the two was keyed without the other.'),
 ('ORDER_AFTER_ASOF',     'POLine',    'Medium', 'Timing',    CAST(0 AS BIT),
  'A purchase order dated after the reporting date, so the snapshot includes future commitment.'),
 ('NEVER_RECEIVED',       'POLine',    'Medium', 'Quantity',  CAST(0 AS BIT),
  'An order placed more than 180 days ago with no receipt at all, EXCLUDING lines already reported as duplicates -- a duplicated requisition has no goods against it because it is a duplicate, and counting it here too would report one defect twice.')
) AS c(AnomalyType, EntityType, Severity, ImpactClass, CountsTowardExposure, WhatItMeans);
GO

/*
--------------------------------------------------------------------------------
fn_DQ_Anomalies(@AsOf) -- one row per defect.
Parameterized so the same checks can run against any reporting date, e.g. to
prove a defect was already present last quarter.
--------------------------------------------------------------------------------
*/
/*
--------------------------------------------------------------------------------
fn_DuplicatePOLines(@AsOf) -- the surplus copies of a repeated requisition.

ONE definition of "duplicate", reused everywhere. It was previously written out
inline here, while fn_VendorScorecard excluded duplicates with
`PONumber NOT LIKE 'PO-D%'` -- reading the generator's planted marker, which is
exactly what UAT-15 exists to prove the data-quality layer does NOT do. On this
dataset the two agree perfectly (26 lines either way), because the generator
happens to give every planted duplicate that prefix. On real data there is no
'PO-D' prefix and that exclusion would have silently matched nothing.

A check that only works against its own test fixture is not a reusable check.
--------------------------------------------------------------------------------
*/
IF OBJECT_ID('dbo.fn_DuplicatePOLines', 'IF') IS NOT NULL DROP FUNCTION dbo.fn_DuplicatePOLines;
GO
CREATE FUNCTION dbo.fn_DuplicatePOLines (@AsOf DATE)
RETURNS TABLE
AS RETURN
(
    WITH AsOfKey AS (SELECT k = YEAR(@AsOf)*10000 + MONTH(@AsOf)*100 + DAY(@AsOf)),
    -- surplus copies of a requisition: keep the first, flag the rest. Matched on
    -- the business signature, never on the generator's 'PO-D' prefix.
    DupRank AS (
        SELECT l.POLineKey, l.PONumber, l.PartKey, l.VendorKey, l.OrderQty, l.UnitPrice,
               ExtendedPrice = l.OrderQty * l.UnitPrice,
               CopyNo = ROW_NUMBER() OVER (
                          PARTITION BY l.PartKey, l.VendorKey, l.OrderDateKey, l.OrderQty, l.UnitPrice
                          ORDER BY l.POLineKey)
        FROM dbo.Fact_PurchaseOrderLine l
        CROSS JOIN AsOfKey a
        WHERE l.OrderDateKey <= a.k
    )
    SELECT POLineKey, PONumber, PartKey, VendorKey, OrderQty, UnitPrice, ExtendedPrice, CopyNo
    FROM DupRank
    WHERE CopyNo > 1
);
GO

CREATE FUNCTION dbo.fn_DQ_Anomalies (@AsOf DATE)
RETURNS TABLE
AS RETURN
(
    WITH AsOfKey AS (SELECT k = YEAR(@AsOf)*10000 + MONTH(@AsOf)*100 + DAY(@AsOf))
    SELECT AnomalyType = 'DUPLICATE_PO_LINE',
           EntityType  = 'POLine',
           EntityRef   = d.PONumber,
           AmountAtRisk = CAST(d.ExtendedPrice AS DECIMAL(16,2)),
           Detail = CONCAT('Duplicates an earlier line for the same part, vendor, date, quantity and price: ',
                           FORMAT(d.ExtendedPrice, 'C', 'en-US'))
    FROM dbo.fn_DuplicatePOLines(@AsOf) d

    UNION ALL
    SELECT 'AMBIGUOUS_CONTRACT', 'POLine', c.PONumber,
           CAST(c.ExtendedPrice AS DECIMAL(16,2)),
           CONCAT(c.AgreementsInForce, ' price agreements were in force at the order date; the contracted price is ambiguous')
    FROM dbo.fn_POLineCost(@AsOf) c
    WHERE c.AgreementsInForce > 1

    UNION ALL
    SELECT 'PRICE_OUTLIER', 'POLine', c.PONumber,
           CAST(c.ExtendedPrice - ISNULL(c.ContractedExtended, 0) AS DECIMAL(16,2)),
           CONCAT('Paid ', FORMAT(c.UnitPrice, 'C', 'en-US'),
                  ' against a contracted ', FORMAT(c.ContractedUnitPrice, 'C', 'en-US'))
    FROM dbo.fn_POLineCost(@AsOf) c
    WHERE c.ContractedUnitPrice IS NOT NULL
      AND c.UnitPrice > c.ContractedUnitPrice * 5

    UNION ALL
    SELECT 'RECEIPT_BEFORE_ORDER', 'Receipt', g.ReceiptNo,
           CAST(g.QtyReceived * l.UnitPrice AS DECIMAL(16,2)),
           CONCAT('Received ', CONVERT(CHAR(10), dg.[Date], 120),
                  ' but ordered ', CONVERT(CHAR(10), dl.[Date], 120))
    FROM dbo.Fact_GoodsReceipt g
    JOIN dbo.Fact_PurchaseOrderLine l ON l.POLineKey = g.POLineKey
    JOIN dbo.Dim_Date dg ON dg.DateKey = g.ReceiptDateKey
    JOIN dbo.Dim_Date dl ON dl.DateKey = l.OrderDateKey
    CROSS JOIN AsOfKey a
    WHERE g.ReceiptDateKey < l.OrderDateKey AND g.ReceiptDateKey <= a.k

    UNION ALL
    SELECT 'RECEIPT_EXCEEDS_ORDER', 'Receipt', CONCAT(l.PONumber, '/', l.POLineNo),
           CAST((x.TotalReceived - l.OrderQty) * l.UnitPrice AS DECIMAL(16,2)),
           CONCAT('Received ', x.TotalReceived, ' against an order of ', l.OrderQty)
    FROM dbo.Fact_PurchaseOrderLine l
    CROSS JOIN AsOfKey a
    CROSS APPLY (SELECT TotalReceived = SUM(g.QtyReceived)
                 FROM dbo.Fact_GoodsReceipt g
                 WHERE g.POLineKey = l.POLineKey AND g.ReceiptDateKey <= a.k) x
    WHERE x.TotalReceived > l.OrderQty AND l.OrderDateKey <= a.k

    UNION ALL
    -- defence in depth: foreign keys already prevent this, so a hit means the
    -- constraints were dropped or bypassed by a bulk load
    SELECT 'ORPHAN_DIMENSION_KEY', 'POLine', l.PONumber,
           CAST(l.OrderQty * l.UnitPrice AS DECIMAL(16,2)),
           'Line references a missing part, vendor, buyer or date row'
    FROM dbo.Fact_PurchaseOrderLine l
    WHERE NOT EXISTS (SELECT 1 FROM dbo.Dim_Part   p WHERE p.PartKey   = l.PartKey)
       OR NOT EXISTS (SELECT 1 FROM dbo.Dim_Vendor v WHERE v.VendorKey = l.VendorKey)
       OR NOT EXISTS (SELECT 1 FROM dbo.Dim_Buyer  b WHERE b.BuyerKey  = l.BuyerKey)
       OR NOT EXISTS (SELECT 1 FROM dbo.Dim_Date   d WHERE d.DateKey   = l.OrderDateKey)

    UNION ALL
    SELECT 'AGREEMENT_INVERTED', 'Agreement', pa.AgreementNo,
           CAST(0 AS DECIMAL(16,2)),
           'Validity window ends before it begins'
    FROM dbo.Fact_PriceAgreement pa
    WHERE pa.ValidToDateKey IS NOT NULL AND pa.ValidToDateKey < pa.ValidFromDateKey

    UNION ALL
    SELECT 'EXPEDITE_FLAG_MISMATCH', 'POLine', l.PONumber,
           CAST(l.ExpediteFee AS DECIMAL(16,2)),
           'Expedite fee and expedite flag disagree'
    FROM dbo.Fact_PurchaseOrderLine l
    CROSS JOIN AsOfKey a
    WHERE l.OrderDateKey <= a.k
      AND ((l.IsExpedited = 1 AND l.ExpediteFee <= 0) OR (l.IsExpedited = 0 AND l.ExpediteFee > 0))

    UNION ALL
    SELECT 'ORDER_AFTER_ASOF', 'POLine', l.PONumber,
           CAST(l.OrderQty * l.UnitPrice AS DECIMAL(16,2)),
           CONCAT('Ordered ', CONVERT(CHAR(10), d.[Date], 120), ', after the reporting date')
    FROM dbo.Fact_PurchaseOrderLine l
    JOIN dbo.Dim_Date d ON d.DateKey = l.OrderDateKey
    CROSS JOIN AsOfKey a
    WHERE l.OrderDateKey > a.k

    UNION ALL
    -- Lines already flagged as duplicates are excluded. A duplicated
    -- requisition never had goods against it BECAUSE it is a duplicate, so
    -- reporting it here as well counts one defect twice and pads the headline
    -- with a consequence of something already on the report.
    SELECT 'NEVER_RECEIVED', 'POLine', c.PONumber,
           CAST(c.ExtendedPrice AS DECIMAL(16,2)),
           CONCAT('Ordered ', CONVERT(CHAR(10), c.OrderDate, 120), ' with no receipt booked')
    FROM dbo.fn_POLineCost(@AsOf) c
    WHERE c.IsReceived = 0 AND DATEDIFF(DAY, c.OrderDate, @AsOf) > 180
      AND NOT EXISTS (
          SELECT 1 FROM dbo.Fact_PurchaseOrderLine l2
          WHERE l2.PartKey = c.PartKey AND l2.VendorKey = c.VendorKey
            AND l2.OrderQty = c.OrderQty AND l2.UnitPrice = c.UnitPrice
            AND l2.POLineKey < c.POLineKey)
);
GO

CREATE VIEW dbo.vw_DQ_Anomalies AS
SELECT * FROM dbo.fn_DQ_Anomalies('2025-12-31');
GO

/*
--------------------------------------------------------------------------------
vw_DQ_Summary -- every catalogued check, whether or not it fired. Rates are
expressed against the population each check actually examines, so a
receipt-level defect is not diluted by the purchase-order line count.
--------------------------------------------------------------------------------
*/
CREATE VIEW dbo.vw_DQ_Summary AS
WITH Pop AS (
    SELECT POLineCount    = (SELECT COUNT(*) FROM dbo.Fact_PurchaseOrderLine),
           ReceiptCount   = (SELECT COUNT(*) FROM dbo.Fact_GoodsReceipt),
           AgreementCount = (SELECT COUNT(*) FROM dbo.Fact_PriceAgreement)
),
Found AS (
    SELECT AnomalyType, Anomalies = COUNT(*), AmountAtRisk = SUM(AmountAtRisk)
    FROM dbo.vw_DQ_Anomalies GROUP BY AnomalyType
)
SELECT
    cat.AnomalyType, cat.EntityType, cat.Severity, cat.ImpactClass, cat.CountsTowardExposure,
    Anomalies    = ISNULL(f.Anomalies, 0),
    AmountAtRisk = CAST(ISNULL(f.AmountAtRisk, 0) AS DECIMAL(16,2)),
    PopulationScanned = CASE cat.EntityType WHEN 'Receipt' THEN p.ReceiptCount
                                            WHEN 'Agreement' THEN p.AgreementCount
                                            ELSE p.POLineCount END,
    AnomalyRatePct = CAST(100.0 * ISNULL(f.Anomalies, 0)
        / NULLIF(CASE cat.EntityType WHEN 'Receipt' THEN p.ReceiptCount
                                     WHEN 'Agreement' THEN p.AgreementCount
                                     ELSE p.POLineCount END, 0) AS DECIMAL(6,3)),
    cat.WhatItMeans
FROM dbo.vw_DQ_CheckCatalog cat
LEFT JOIN Found f ON f.AnomalyType = cat.AnomalyType
CROSS JOIN Pop p;
GO

/*
--------------------------------------------------------------------------------
usp_RunDataQualityChecks -- the gate.

The mis-statement tolerance is a PERCENTAGE of total spend rather than a fixed
dollar figure, because an absolute threshold silently becomes stricter as the
book grows and looser as it shrinks, without anyone deciding that it should.
--------------------------------------------------------------------------------
*/
CREATE PROCEDURE dbo.usp_RunDataQualityChecks
    @AsOf DATE = '2025-12-31',
    @MaxAcceptableRatePct    DECIMAL(6,3) = 1.000,
    @MaxMisstatementPctOfSpend DECIMAL(6,3) = 0.500
AS
BEGIN
    SET NOCOUNT ON;

    PRINT '=== Lumen Optics Manufacturing: Procurement Data Quality Report (as of '
          + CONVERT(CHAR(10), @AsOf, 120) + ') ===';

    SELECT AnomalyType, EntityType, Severity, ImpactClass, Anomalies, AmountAtRisk,
           PopulationScanned, AnomalyRatePct, CountsTowardExposure,
           Result = CASE WHEN AnomalyRatePct <= @MaxAcceptableRatePct THEN 'PASS' ELSE 'REVIEW' END
    FROM dbo.vw_DQ_Summary
    ORDER BY CASE Severity WHEN 'High' THEN 1 WHEN 'Medium' THEN 2 ELSE 3 END,
             Anomalies DESC, AnomalyType;

    DECLARE @defects INT       = (SELECT SUM(Anomalies) FROM dbo.vw_DQ_Summary);
    DECLARE @exposure DECIMAL(16,2) = (SELECT ISNULL(SUM(AmountAtRisk), 0)
                                       FROM dbo.vw_DQ_Summary WHERE CountsTowardExposure = 1);
    DECLARE @spend DECIMAL(16,2) = (SELECT SUM(LandedCost) FROM dbo.fn_POLineCost(@AsOf));
    DECLARE @pct DECIMAL(9,3) = CAST(100.0 * @exposure / NULLIF(@spend, 0) AS DECIMAL(9,3));

    PRINT '';
    PRINT 'Defects found: ' + CAST(@defects AS VARCHAR(20))
        + ' | Spend mis-stated by: ' + FORMAT(@exposure, 'C', 'en-US')
        + ' (' + CAST(@pct AS VARCHAR(20)) + '% of ' + FORMAT(@spend, 'C', 'en-US') + ' landed cost)';

    IF @pct > @MaxMisstatementPctOfSpend
    BEGIN
        PRINT 'QA GATE: FAIL -- spend mis-stated by ' + CAST(@pct AS VARCHAR(20))
            + '%, above the ' + CAST(@MaxMisstatementPctOfSpend AS VARCHAR(20))
            + '% tolerance. See dbo.vw_DQ_Anomalies.';
        -- A gate that only PRINTs is not a gate. This printed 'QA GATE: FAIL'
        -- and returned exit code 0, so sqlcmd -b reported success and nothing
        -- downstream could act on it -- while the README described it as "a
        -- gate that can fail the build". The identical defect was found in
        -- Project 1's gate on the same day; it is worth noting that both were
        -- written the same way and neither was noticed until the exit code was
        -- actually checked, because the printed word FAIL reads as a failure.
        DECLARE @msg VARCHAR(400) =
            'QA gate failed: spend mis-stated by ' + CAST(@pct AS VARCHAR(20))
            + '% of landed cost, above the ' + CAST(@MaxMisstatementPctOfSpend AS VARCHAR(20))
            + '% tolerance. See dbo.vw_DQ_Anomalies.';
        ;THROW 51005, @msg, 1;
    END
    ELSE
        PRINT 'QA GATE: PASS -- spend mis-statement within tolerance.';
END;
GO

DECLARE @missing VARCHAR(400) = '';
IF OBJECT_ID('dbo.vw_DQ_CheckCatalog', 'V')       IS NULL SET @missing += 'vw_DQ_CheckCatalog ';
IF OBJECT_ID('dbo.fn_DQ_Anomalies', 'IF')         IS NULL SET @missing += 'fn_DQ_Anomalies ';
IF OBJECT_ID('dbo.vw_DQ_Anomalies', 'V')          IS NULL SET @missing += 'vw_DQ_Anomalies ';
IF OBJECT_ID('dbo.vw_DQ_Summary', 'V')            IS NULL SET @missing += 'vw_DQ_Summary ';
IF OBJECT_ID('dbo.usp_RunDataQualityChecks', 'P') IS NULL SET @missing += 'usp_RunDataQualityChecks ';
IF @missing <> '' THROW 51004, 'FAILED to create data-quality objects -- scroll up for the compile error.', 1;
PRINT 'Data-quality objects created and verified.';
GO
