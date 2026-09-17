/*
================================================================================
Project 2 -- Vantage Wholesale Supply: Receivables Performance
Script:  02_generate_synthetic_data.sql
Purpose: Populates VantageAR with two years of synthetic invoices, cash
         receipts, credit memos and write-offs (Jan 2024 - Dec 2025, reporting
         date 2025-12-31), with documented behavioral patterns for the KPI
         layer to find, plus a small, known set of data-quality defects.

REPRODUCIBLE BY DESIGN: every random draw comes from dbo.fn_Rand(), a hash of
a stable text key (e.g. 'inv|1234|late'). Re-running this script produces
the identical dataset, so figures quoted in the case study never drift.

DATA DISCLOSURE: 100% synthetic. No real company, customer, or person.

Deliberate patterns (the "story"):
  P1 Terms extension  - From 2025-07-01, Contractor accounts in West and
                         Southwest on NET30 were moved to NET60 in a sales
                         push. Volume rose ~35%, and so did days-to-pay.
  P2 Chronic slow pay - ~12% of OEM Manufacturer accounts pay 50-70 days past
                         due; they carry most of the 90+ balance.
  P3 Dispute cluster  - Invoices sold by one Southeast rep are disputed at
                         ~4x the company rate (mostly Pricing / PO Mismatch),
                         and a disputed invoice is not paid until closed.
  P4 Early-pay lever  - Low-risk accounts on 2/10 NET30 mostly pay in 10
                         days and take the 2% discount.
  P5 Defaults         - A few high-risk Contractors stop paying; their
                         invoices are written off 150 days past due.
  P6 Slow but safe    - Institutional accounts pay ~2-3 weeks late but
                         always pay.
================================================================================
*/

USE VantageAR;
GO
SET NOCOUNT ON;
GO

-- =============================================================================
-- 0. Deterministic random-number helper
--    First 4 bytes of SHA2_256(key) as an unsigned integer / 2^32 -> [0,1)
-- =============================================================================
IF OBJECT_ID('dbo.fn_Rand', 'IF') IS NOT NULL DROP FUNCTION dbo.fn_Rand;
GO
CREATE FUNCTION dbo.fn_Rand (@Key NVARCHAR(200))
RETURNS TABLE
AS RETURN
    SELECT CAST(CONVERT(BIGINT, CONVERT(BINARY(4), HASHBYTES('SHA2_256', @Key))) AS FLOAT) / 4294967296.0 AS r;
GO

DECLARE @AsOf DATE = '2025-12-31';

-- =============================================================================
-- 1. Dim_Date  2023-11-01 .. 2026-06-30
--    Starts two months before the first invoice because SHIP dates precede
--    invoice dates: a date dimension has to cover every date any fact can
--    reference, or the first billing lag inserted violates a foreign key.
-- =============================================================================
;WITH d AS (
    SELECT CAST('2023-11-01' AS DATE) AS dt
    UNION ALL SELECT DATEADD(DAY, 1, dt) FROM d WHERE dt < '2026-06-30'
)
INSERT INTO dbo.Dim_Date (DateKey, [Date], [Year], [Quarter], [Month], MonthName, YearMonth, MonthEndDate, IsMonthEnd, IsWeekend)
SELECT YEAR(dt)*10000 + MONTH(dt)*100 + DAY(dt), dt, YEAR(dt), DATEPART(QUARTER, dt), MONTH(dt),
       DATENAME(MONTH, dt), CONVERT(CHAR(7), dt, 126), EOMONTH(dt),
       CASE WHEN dt = EOMONTH(dt) THEN 1 ELSE 0 END,
       CASE WHEN DATENAME(WEEKDAY, dt) IN ('Saturday','Sunday') THEN 1 ELSE 0 END
FROM d
OPTION (MAXRECURSION 0);   -- the range now exceeds the 1000 default

-- =============================================================================
-- 2. Dim_PaymentTerms, Dim_Collector
-- =============================================================================
INSERT INTO dbo.Dim_PaymentTerms (TermsCode, TermsName, NetDays, DiscountPct, DiscountDays) VALUES
('NET30',     'Net 30',                 30, 0, 0),
('2/10NET30', '2% 10, Net 30',          30, 2, 10),
('NET45',     'Net 45',                 45, 0, 0),
('NET60',     'Net 60',                 60, 0, 0),
('NET90',     'Net 90',                 90, 0, 0);

INSERT INTO dbo.Dim_Collector (CollectorID, CollectorName, Team) VALUES
('COL-01', 'Priya Natarajan',  'Strategic Accounts'),
('COL-02', 'Marcus Bell',      'Strategic Accounts'),
('COL-03', 'Elena Ruiz',       'Commercial'),
('COL-04', 'Tomas Okafor',     'Commercial'),
('COL-05', 'Grace Lindqvist',  'Small Business'),
('COL-06', 'Sam Horowitz',     'Small Business');

-- =============================================================================
-- 3. Customers (400) with behavioral attributes kept in #Cust
-- =============================================================================
IF OBJECT_ID('tempdb..#Cust') IS NOT NULL DROP TABLE #Cust;

;WITH n AS (
    SELECT TOP (400) ROW_NUMBER() OVER (ORDER BY (SELECT NULL)) AS CustN
    FROM sys.all_objects a CROSS JOIN sys.all_objects b
),
base AS (
    SELECT n.CustN,
        CASE WHEN r1.r < 0.40 THEN 'Contractor' WHEN r1.r < 0.60 THEN 'OEM Manufacturer'
             WHEN r1.r < 0.85 THEN 'MRO Reseller' ELSE 'Institutional' END AS Segment,
        CASE CAST(FLOOR(r2.r * 5) AS INT) WHEN 0 THEN 'Northeast' WHEN 1 THEN 'Southeast'
             WHEN 2 THEN 'Midwest' WHEN 3 THEN 'Southwest' ELSE 'West' END AS Region,
        r3.r AS rRisk, r4.r AS rTerms, r5.r AS rLimit, r6.r AS rFlag, r7.r AS rRep, r8.r AS rSince,
        r9.r AS rName, r10.r AS rColl, r11.r AS rDefault
    FROM n
    CROSS APPLY dbo.fn_Rand(CONCAT('cust|', n.CustN, '|segment')) r1
    CROSS APPLY dbo.fn_Rand(CONCAT('cust|', n.CustN, '|region'))  r2
    CROSS APPLY dbo.fn_Rand(CONCAT('cust|', n.CustN, '|risk'))    r3
    CROSS APPLY dbo.fn_Rand(CONCAT('cust|', n.CustN, '|terms'))   r4
    CROSS APPLY dbo.fn_Rand(CONCAT('cust|', n.CustN, '|limit'))   r5
    CROSS APPLY dbo.fn_Rand(CONCAT('cust|', n.CustN, '|flag'))    r6
    CROSS APPLY dbo.fn_Rand(CONCAT('cust|', n.CustN, '|rep'))     r7
    CROSS APPLY dbo.fn_Rand(CONCAT('cust|', n.CustN, '|since'))   r8
    CROSS APPLY dbo.fn_Rand(CONCAT('cust|', n.CustN, '|name'))    r9
    CROSS APPLY dbo.fn_Rand(CONCAT('cust|', n.CustN, '|coll'))    r10
    CROSS APPLY dbo.fn_Rand(CONCAT('cust|', n.CustN, '|default')) r11
),
attr AS (
    SELECT b.*,
        CASE
            WHEN Segment = 'Institutional'    THEN CASE WHEN rRisk < 0.90 THEN 'Low' ELSE 'Medium' END
            WHEN Segment = 'Contractor'       THEN CASE WHEN rRisk < 0.15 THEN 'High' WHEN rRisk < 0.55 THEN 'Medium' ELSE 'Low' END
            WHEN Segment = 'OEM Manufacturer' THEN CASE WHEN rRisk < 0.10 THEN 'High' WHEN rRisk < 0.40 THEN 'Medium' ELSE 'Low' END
            ELSE                                   CASE WHEN rRisk < 0.10 THEN 'High' WHEN rRisk < 0.50 THEN 'Medium' ELSE 'Low' END
        END AS RiskTier,
        CASE
            WHEN Segment = 'Contractor'       THEN CASE WHEN rTerms < 0.30 THEN '2/10NET30' ELSE 'NET30' END
            WHEN Segment = 'OEM Manufacturer' THEN CASE WHEN rTerms < 0.50 THEN 'NET45' ELSE 'NET60' END
            WHEN Segment = 'MRO Reseller'     THEN CASE WHEN rTerms < 0.45 THEN '2/10NET30' ELSE 'NET30' END
            ELSE                                   CASE WHEN rTerms < 0.50 THEN 'NET60' ELSE 'NET90' END
        END AS OrigTermsCode,
        ROUND(CASE
            WHEN Segment = 'OEM Manufacturer' THEN 150000 + rLimit * 450000
            WHEN Segment = 'MRO Reseller'     THEN  25000 + rLimit *  95000
            WHEN Segment = 'Contractor'       THEN  20000 + rLimit * 130000
            ELSE                                    50000 + rLimit * 200000
        END / 5000, 0) * 5000 AS CreditLimit
    FROM base b
)
SELECT a.*,
    CAST(CASE WHEN Segment = 'OEM Manufacturer' AND rFlag < 0.12 THEN 1 ELSE 0 END AS BIT) AS IsChronicSlow,
    CAST(CASE WHEN Segment = 'Contractor' AND RiskTier = 'High' AND rDefault < 0.30 THEN 1 ELSE 0 END AS BIT) AS IsDefaulter,
    CAST(CASE WHEN Segment = 'Contractor' AND Region IN ('West','Southwest') AND OrigTermsCode = 'NET30' THEN 1 ELSE 0 END AS BIT) AS InTermsPromo,
    DATEADD(DAY, CAST(FLOOR(rDefault * 240) AS INT), CAST('2024-10-01' AS DATE)) AS DefaultStart,
    CASE Region
        WHEN 'Northeast' THEN CASE WHEN rRep < 0.5 THEN 'Adrian Cole'    ELSE 'Nadia Farouk' END
        WHEN 'Southeast' THEN CASE WHEN rRep < 0.5 THEN 'Dana Whitfield' ELSE 'Luis Moreno'  END
        WHEN 'Midwest'   THEN CASE WHEN rRep < 0.5 THEN 'Ben Kowalczyk'  ELSE 'Hana Sato'    END
        WHEN 'Southwest' THEN CASE WHEN rRep < 0.5 THEN 'Rosa Delgado'   ELSE 'Kevin Ames'   END
        ELSE                  CASE WHEN rRep < 0.5 THEN 'Maya Brennan'   ELSE 'Owen Price'   END
    END AS SalesRep
INTO #Cust
FROM attr a;

-- Customer names: prefix x segment-specific suffix
;WITH pre AS (
    SELECT * FROM (VALUES (0,'Apex'),(1,'Summit'),(2,'Keystone'),(3,'Ironbridge'),(4,'Northfield'),(5,'Redline'),
        (6,'Granite'),(7,'Bluewater'),(8,'Cedar Ridge'),(9,'Harbor'),(10,'Pioneer'),(11,'Silverline'),(12,'Titan'),
        (13,'Frontier'),(14,'Crestview'),(15,'Riverbend'),(16,'Lakeshore'),(17,'Oakmont'),(18,'Stonegate'),(19,'Westbrook')) p(i, Prefix)
),
suf AS (
    SELECT * FROM (VALUES
        ('Contractor',0,'Construction'),('Contractor',1,'Builders'),('Contractor',2,'Mechanical'),('Contractor',3,'Electric'),('Contractor',4,'Contracting'),
        ('OEM Manufacturer',0,'Manufacturing'),('OEM Manufacturer',1,'Industries'),('OEM Manufacturer',2,'Fabrication'),('OEM Manufacturer',3,'Machine Works'),('OEM Manufacturer',4,'Components'),
        ('MRO Reseller',0,'Industrial Supply'),('MRO Reseller',1,'MRO Services'),('MRO Reseller',2,'Maintenance Supply'),('MRO Reseller',3,'Parts & Service'),('MRO Reseller',4,'Supply Co.'),
        ('Institutional',0,'School District'),('Institutional',1,'Medical Center'),('Institutional',2,'County Facilities'),('Institutional',3,'University Facilities'),('Institutional',4,'Water Authority')
    ) s(Segment, i, Suffix)
),
named AS (
    SELECT c.CustN, pre.Prefix + ' ' + suf.Suffix AS BaseName,
           ROW_NUMBER() OVER (PARTITION BY pre.Prefix, suf.Suffix ORDER BY c.CustN) AS dup, c.Region
    FROM #Cust c
    JOIN pre ON pre.i = CAST(FLOOR(c.rName * 20) AS INT)
    JOIN suf ON suf.Segment = c.Segment AND suf.i = CAST(FLOOR(c.rTerms * 1000) AS INT) % 5
)
INSERT INTO dbo.Dim_Customer (CustomerID, CustomerName, Segment, Region, SalesRep, CollectorKey, CurrentTermsKey, CreditLimit, RiskTier, CustomerSince)
SELECT
    'C' + RIGHT('0000' + CAST(c.CustN AS VARCHAR(4)), 4),
    nm.BaseName + CASE WHEN nm.dup = 1 THEN '' ELSE ' - ' + c.Region + CASE WHEN nm.dup > 2 THEN ' ' + CAST(nm.dup AS VARCHAR(3)) ELSE '' END END,
    c.Segment, c.Region, c.SalesRep,
    col.CollectorKey,
    ct.TermsKey,
    c.CreditLimit, c.RiskTier,
    DATEADD(DAY, -CAST(FLOOR(c.rSince * 3650) AS INT), CAST('2023-12-31' AS DATE))
FROM #Cust c
JOIN named nm ON nm.CustN = c.CustN
JOIN dbo.Dim_PaymentTerms ct ON ct.TermsCode = CASE WHEN c.InTermsPromo = 1 THEN 'NET60' ELSE c.OrigTermsCode END
JOIN dbo.Dim_Collector col ON col.CollectorID =
    CASE
        WHEN c.CreditLimit >= 150000 THEN CASE WHEN c.rColl < 0.5 THEN 'COL-01' ELSE 'COL-02' END
        WHEN c.Segment IN ('Contractor','Institutional') THEN CASE WHEN c.rColl < 0.5 THEN 'COL-03' ELSE 'COL-04' END
        ELSE CASE WHEN c.rColl < 0.5 THEN 'COL-05' ELSE 'COL-06' END
    END
ORDER BY c.CustN;

ALTER TABLE #Cust ADD CustomerKey INT NULL;
UPDATE c SET c.CustomerKey = dc.CustomerKey
FROM #Cust c JOIN dbo.Dim_Customer dc ON dc.CustomerID = 'C' + RIGHT('0000' + CAST(c.CustN AS VARCHAR(4)), 4);

-- =============================================================================
-- 4. Invoices: customer x business day, kept with segment-specific daily rates
-- =============================================================================
IF OBJECT_ID('tempdb..#Inv') IS NOT NULL DROP TABLE #Inv;

;WITH cand AS (
    SELECT c.CustomerKey, c.CustN, c.Segment, c.Region, c.SalesRep, c.RiskTier, c.OrigTermsCode,
           c.IsChronicSlow, c.IsDefaulter, c.InTermsPromo, c.DefaultStart,
           d.[Date] AS InvDate, d.[Month] AS Mo,
           CASE c.Segment WHEN 'OEM Manufacturer' THEN 0.22 WHEN 'MRO Reseller' THEN 0.16
                          WHEN 'Contractor' THEN 0.10 ELSE 0.07 END
         * (0.5 + c.rLimit)
         * CASE WHEN c.Segment = 'Contractor' AND d.[Month] IN (1,2) THEN 0.6
                WHEN c.Segment = 'Contractor' AND d.[Month] BETWEEN 4 AND 9 THEN 1.25 ELSE 1.0 END
         * CASE WHEN c.InTermsPromo = 1 AND d.[Date] >= '2025-07-01' THEN 1.35 ELSE 1.0 END AS KeepProb
    FROM #Cust c
    CROSS JOIN dbo.Dim_Date d
    WHERE d.[Date] BETWEEN '2024-01-01' AND '2025-12-31' AND d.IsWeekend = 0
)
SELECT cand.*, rk.r AS rKeep
INTO #InvCand
FROM cand
CROSS APPLY dbo.fn_Rand(CONCAT('cand|', cand.CustN, '|', CONVERT(CHAR(8), cand.InvDate, 112))) rk;

SELECT
    ROW_NUMBER() OVER (ORDER BY c.InvDate, c.CustN) AS InvN,
    c.CustomerKey, c.CustN, c.Segment, c.Region, c.SalesRep, c.RiskTier, c.IsChronicSlow, c.IsDefaulter, c.DefaultStart,
    c.InvDate,
    CASE WHEN c.InTermsPromo = 1 AND c.InvDate >= '2025-07-01' THEN 'NET60' ELSE c.OrigTermsCode END AS TermsCode
INTO #Inv
FROM #InvCand c
WHERE c.rKeep < c.KeepProb;

DROP TABLE #InvCand;

ALTER TABLE #Inv ADD
    NetDays SMALLINT, DiscountPct DECIMAL(5,2), TermsKey INT, DueDate DATE, Amount DECIMAL(12,2),
    rDisp FLOAT, rReason FLOAT, rDispOpen FLOAT, rDispDur FLOAT, rDisc FLOAT, rLate FLOAT, rPartial FLOAT,
    rPay2 FLOAT, rMethod FLOAT, rMemo FLOAT, rAfterClose FLOAT,
    IsDisputed BIT, DisputeReason VARCHAR(40), DisputeOpened DATE, DisputeClosed DATE,
    PlanPayDate DATE, TakesDiscount BIT, IsPartial BIT, MemoPct DECIMAL(5,2),
    ShipDate DATE;

UPDATE i SET
    i.NetDays = t.NetDays, i.DiscountPct = t.DiscountPct, i.TermsKey = t.TermsKey,
    i.DueDate = DATEADD(DAY, t.NetDays, i.InvDate),
    i.Amount = ROUND(
        CASE i.Segment WHEN 'OEM Manufacturer' THEN 6500 WHEN 'MRO Reseller' THEN 1800 WHEN 'Contractor' THEN 3200 ELSE 4200 END
        * (0.25 + 2.5 * ra.r * ra.r), 2),
    i.rDisp = r1.r, i.rReason = r2.r, i.rDispOpen = r3.r, i.rDispDur = r4.r, i.rDisc = r5.r,
    i.rLate = r6.r, i.rPartial = r7.r, i.rPay2 = r8.r, i.rMethod = r9.r, i.rMemo = r10.r, i.rAfterClose = r11.r
FROM #Inv i
JOIN dbo.Dim_PaymentTerms t ON t.TermsCode = i.TermsCode
CROSS APPLY dbo.fn_Rand(CONCAT('inv|', i.InvN, '|amount'))    ra
CROSS APPLY dbo.fn_Rand(CONCAT('inv|', i.InvN, '|dispute'))   r1
CROSS APPLY dbo.fn_Rand(CONCAT('inv|', i.InvN, '|reason'))    r2
CROSS APPLY dbo.fn_Rand(CONCAT('inv|', i.InvN, '|dispopen'))  r3
CROSS APPLY dbo.fn_Rand(CONCAT('inv|', i.InvN, '|dispdur'))   r4
CROSS APPLY dbo.fn_Rand(CONCAT('inv|', i.InvN, '|discount'))  r5
CROSS APPLY dbo.fn_Rand(CONCAT('inv|', i.InvN, '|late'))      r6
CROSS APPLY dbo.fn_Rand(CONCAT('inv|', i.InvN, '|partial'))   r7
CROSS APPLY dbo.fn_Rand(CONCAT('inv|', i.InvN, '|pay2'))      r8
CROSS APPLY dbo.fn_Rand(CONCAT('inv|', i.InvN, '|method'))    r9
CROSS APPLY dbo.fn_Rand(CONCAT('inv|', i.InvN, '|memo'))      r10
CROSS APPLY dbo.fn_Rand(CONCAT('inv|', i.InvN, '|afterclose')) r11;

-- Disputes (P3): ~4% base, ~16% for the Southeast rep's accounts, +2 pts for OEM
UPDATE #Inv SET
    IsDisputed = CASE WHEN rDisp < (CASE WHEN SalesRep = 'Dana Whitfield' THEN 0.16 ELSE 0.035 END
                                    + CASE WHEN Segment = 'OEM Manufacturer' THEN 0.02 ELSE 0 END) THEN 1 ELSE 0 END,
    DisputeOpened = DATEADD(DAY, 3 + CAST(FLOOR(rDispOpen * 18) AS INT), InvDate);

UPDATE #Inv SET
    DisputeReason = CASE
        WHEN SalesRep = 'Dana Whitfield' THEN CASE WHEN rReason < 0.50 THEN 'Pricing' WHEN rReason < 0.80 THEN 'PO Mismatch' WHEN rReason < 0.90 THEN 'Short Shipment' ELSE 'Damaged Goods' END
        ELSE CASE WHEN rReason < 0.25 THEN 'Pricing' WHEN rReason < 0.45 THEN 'Short Shipment' WHEN rReason < 0.65 THEN 'Damaged Goods' WHEN rReason < 0.85 THEN 'PO Mismatch' ELSE 'Duplicate Billing' END
    END,
    DisputeClosed = DATEADD(DAY, 10 + CAST(FLOOR(POWER(rDispDur, 1.5) * 80) AS INT), DisputeOpened)
WHERE IsDisputed = 1;

-- A dispute that would be raised after the reporting date has not happened yet
UPDATE #Inv SET IsDisputed = 0, DisputeReason = NULL, DisputeOpened = NULL, DisputeClosed = NULL
WHERE IsDisputed = 1 AND DisputeOpened > '2025-12-31';
UPDATE #Inv SET DisputeOpened = NULL WHERE IsDisputed = 0;
UPDATE #Inv SET DisputeClosed = NULL WHERE IsDisputed = 1 AND DisputeClosed > '2025-12-31';

-- Payment plan (P2, P4, P5, P6)
UPDATE i SET
    TakesDiscount = CASE WHEN i.DiscountPct > 0 AND i.IsDisputed = 0
                          AND i.rDisc < CASE i.RiskTier WHEN 'Low' THEN 0.80 WHEN 'Medium' THEN 0.55 ELSE 0.25 END
                         THEN 1 ELSE 0 END,
    IsPartial = CASE WHEN i.rPartial < 0.08 THEN 1 ELSE 0 END,
    MemoPct = CASE
        WHEN i.IsDisputed = 1 AND i.DisputeReason = 'Duplicate Billing' THEN 100
        WHEN i.IsDisputed = 1 AND i.DisputeReason IN ('Pricing','Short Shipment','Damaged Goods') THEN ROUND(5 + i.rMemo * 20, 0)
        ELSE 0 END
FROM #Inv i;

-- A settlement discount is earned by paying the invoice IN FULL inside the
-- discount window; paying 60% and promising the rest does not earn it. So a
-- discount-taker is never a partial payer. Without this the instalment filter
-- below emitted only the first instalment for these invoices, leaving 40% of
-- the value permanently unpayable and ageing into the 90+ bucket forever.
UPDATE #Inv SET IsPartial = 0 WHERE TakesDiscount = 1;

UPDATE i SET PlanPayDate =
    CASE
        WHEN i.IsDefaulter = 1 AND i.InvDate >= i.DefaultStart THEN NULL                           -- never pays
        WHEN i.TakesDiscount = 1 THEN DATEADD(DAY, 4 + CAST(FLOOR(i.rLate * 7) AS INT), i.InvDate)  -- within discount window
        ELSE DATEADD(DAY,
                CAST(ROUND(
                    x.MeanLate + (i.rLate - 0.5) * 2 * (8 + x.MeanLate * 0.4)
                , 0) AS INT)
                , i.DueDate)
    END
FROM #Inv i
CROSS APPLY (SELECT
    CASE i.RiskTier WHEN 'Low' THEN 2 WHEN 'Medium' THEN 9 ELSE 20 END
    + CASE WHEN i.Segment = 'Institutional' THEN 16 ELSE 0 END
    + CASE WHEN i.IsChronicSlow = 1 THEN 55 ELSE 0 END AS MeanLate) x;

-- Nobody pays earlier than 5 days after the invoice date
UPDATE #Inv SET PlanPayDate = DATEADD(DAY, 5, InvDate) WHERE PlanPayDate < DATEADD(DAY, 5, InvDate);

-- A disputed invoice is paid only after the dispute closes; open disputes stay unpaid
UPDATE #Inv SET PlanPayDate =
    CASE WHEN DisputeClosed IS NULL THEN NULL
         WHEN DATEADD(DAY, 3 + CAST(FLOOR(rAfterClose * 12) AS INT), DisputeClosed) > PlanPayDate
              THEN DATEADD(DAY, 3 + CAST(FLOOR(rAfterClose * 12) AS INT), DisputeClosed)
         ELSE PlanPayDate END
WHERE IsDisputed = 1 AND PlanPayDate IS NOT NULL;

-- =============================================================================
-- 5. Load Fact_Invoice
-- =============================================================================

-- -----------------------------------------------------------------------------
-- P7  BILLING LAG: the days between shipping the goods and issuing the invoice.
--
-- This is the one component of the order-to-cash cycle Vantage controls
-- unilaterally -- no customer conversation, no terms negotiation, no phone call.
-- The base lag is one to two days everywhere. The Southeast distribution
-- centre's invoice batch-posting schedule slips progressively through 2025,
-- which is the kind of drift nobody notices because no report shows it.
-- -----------------------------------------------------------------------------
UPDATE i SET i.ShipDate = DATEADD(DAY, -CAST(ROUND(
        1.0 + rs.r
      + CASE WHEN i.Region = 'Southeast' AND i.InvDate >= '2025-01-01'
             THEN 3.4 * (DATEDIFF(DAY, '2025-01-01', i.InvDate) / 364.0)
             ELSE 0 END, 0) AS INT), i.InvDate)
FROM #Inv i
CROSS APPLY dbo.fn_Rand(CONCAT('inv|', i.InvN, '|ship')) rs;
INSERT INTO dbo.Fact_Invoice (InvoiceNo, CustomerKey, TermsKey, ShipDateKey, InvoiceDateKey, DueDateKey, InvoiceAmount,
                              IsDisputed, DisputeReason, DisputeOpenedKey, DisputeClosedKey)
SELECT
    'INV-' + RIGHT('000000' + CAST(InvN AS VARCHAR(6)), 6),
    CustomerKey, TermsKey,
    YEAR(ShipDate)*10000 + MONTH(ShipDate)*100 + DAY(ShipDate),
    YEAR(InvDate)*10000 + MONTH(InvDate)*100 + DAY(InvDate),
    YEAR(DueDate)*10000 + MONTH(DueDate)*100 + DAY(DueDate),
    Amount, IsDisputed, DisputeReason,
    CASE WHEN DisputeOpened IS NULL THEN NULL ELSE YEAR(DisputeOpened)*10000 + MONTH(DisputeOpened)*100 + DAY(DisputeOpened) END,
    CASE WHEN DisputeClosed IS NULL THEN NULL ELSE YEAR(DisputeClosed)*10000 + MONTH(DisputeClosed)*100 + DAY(DisputeClosed) END
FROM #Inv
ORDER BY InvN;

ALTER TABLE #Inv ADD InvoiceKey BIGINT NULL;
UPDATE i SET i.InvoiceKey = f.InvoiceKey
FROM #Inv i JOIN dbo.Fact_Invoice f ON f.InvoiceNo = 'INV-' + RIGHT('000000' + CAST(i.InvN AS VARCHAR(6)), 6);

-- =============================================================================
-- 6. Credit memos (dispute resolutions) and write-offs (P5)
-- =============================================================================
IF OBJECT_ID('tempdb..#Adj') IS NOT NULL DROP TABLE #Adj;
SELECT InvoiceKey, CustomerKey, InvN, DisputeClosed AS AdjDate, 'CreditMemo' AS AdjType,
       ROUND(Amount * MemoPct / 100.0, 2) AS AdjAmount, 'Dispute resolved: ' + DisputeReason AS Reason
INTO #Adj
FROM #Inv
WHERE IsDisputed = 1 AND DisputeClosed IS NOT NULL AND MemoPct > 0;

-- A write-off clears what is still OWED, which is the invoice less any credit
-- memo already issued against it -- not the original face value. Writing off the
-- full amount on an invoice that also carried a dispute credit pushed total
-- adjustments above the invoice value and drove the balance negative.
INSERT INTO #Adj (InvoiceKey, CustomerKey, InvN, AdjDate, AdjType, AdjAmount, Reason)
SELECT InvoiceKey, CustomerKey, InvN, DATEADD(DAY, 150, DueDate), 'WriteOff',
       Amount - ROUND(Amount * MemoPct / 100.0, 2), 'Uncollectible: 150+ days past due'
FROM #Inv
WHERE PlanPayDate IS NULL AND IsDefaulter = 1 AND InvDate >= DefaultStart
  AND DATEADD(DAY, 150, DueDate) <= '2025-12-31'
  AND Amount - ROUND(Amount * MemoPct / 100.0, 2) > 0;

INSERT INTO dbo.Fact_Adjustment (AdjustmentNo, InvoiceKey, CustomerKey, AdjustmentDateKey, AdjustmentType, Amount, Reason)
SELECT 'ADJ-' + RIGHT('000000' + CAST(ROW_NUMBER() OVER (ORDER BY AdjDate, InvN) AS VARCHAR(6)), 6),
       InvoiceKey, CustomerKey, YEAR(AdjDate)*10000 + MONTH(AdjDate)*100 + DAY(AdjDate), AdjType, AdjAmount, Reason
FROM #Adj
WHERE AdjAmount > 0;

-- =============================================================================
-- 7. Cash receipts (only those dated on or before the reporting date exist)
-- =============================================================================
IF OBJECT_ID('tempdb..#Pay') IS NOT NULL DROP TABLE #Pay;

SELECT i.InvoiceKey, i.CustomerKey, i.InvN, i.Segment, i.rMethod, s.Seq,
    CASE WHEN s.Seq = 1 THEN i.PlanPayDate
         ELSE DATEADD(DAY, 20 + CAST(FLOOR(i.rPay2 * 26) AS INT), i.PlanPayDate) END AS PayDate,
    ROUND(
        (i.Amount - ROUND(i.Amount * i.MemoPct / 100.0, 2) - CASE WHEN i.TakesDiscount = 1 THEN ROUND(i.Amount * i.DiscountPct / 100.0, 2) ELSE 0 END)
        * CASE WHEN i.IsPartial = 0 THEN 1.0 WHEN s.Seq = 1 THEN 0.6 ELSE 0.4 END, 2) AS PayAmount,
    CASE WHEN s.Seq = 1 AND i.TakesDiscount = 1 THEN ROUND(i.Amount * i.DiscountPct / 100.0, 2) ELSE 0 END AS Discount
INTO #Pay
FROM #Inv i
CROSS APPLY (VALUES (1),(2)) s(Seq)
WHERE i.PlanPayDate IS NOT NULL
  AND (s.Seq = 1 OR i.IsPartial = 1)   -- IsPartial already excludes discount-takers
  AND i.MemoPct < 100;

-- Fix rounding so partial installments sum exactly to the net amount due
UPDATE p2 SET p2.PayAmount = ROUND(i.Amount - ROUND(i.Amount * i.MemoPct / 100.0, 2), 2) - p1.PayAmount
FROM #Pay p2
JOIN #Pay p1 ON p1.InvoiceKey = p2.InvoiceKey AND p1.Seq = 1
JOIN #Inv i ON i.InvoiceKey = p2.InvoiceKey
WHERE p2.Seq = 2;

DELETE FROM #Pay WHERE PayDate > '2025-12-31' OR PayAmount <= 0;

-- -----------------------------------------------------------------------------
-- Turn the per-invoice settlements into the receipts they actually arrived as.
--
-- A customer does not send one payment per invoice. They send one cheque, wire
-- or ACH covering everything they are settling that day, and somebody in cash
-- application matches it against invoices afterwards. Modelling that properly
-- is what makes unapplied cash representable at all.
--
-- P8  UNAPPLIED CASH: about 2.5% of receipts arrive with no remittance advice
--     and one line is never matched. The money is in the bank, the invoice
--     stays open, and the collections queue will happily send someone to chase
--     a customer who has already paid. That is the single most damaging call a
--     collections team can make, and it is invisible unless receipt and
--     application are separate facts.
--
-- P9  APPLICATION LAG: most cash is applied the day it lands, but a fifth of it
--     takes one to nine days. When that lag crosses a month end, cash banked in
--     December is applied in January -- and the December ageing must not be
--     cured by it.
-- -----------------------------------------------------------------------------
SELECT
    CustomerKey,
    PayDate,
    Segment      = MIN(Segment),
    rMethod      = MIN(rMethod),
    ReceiptAmount = SUM(PayAmount),        -- cash only; a discount is not cash
    Lines        = COUNT(*)
INTO #Rcpt
FROM #Pay
GROUP BY CustomerKey, PayDate;

ALTER TABLE #Rcpt ADD
    ReceiptNo VARCHAR(15), ApplyLagDays INT, HasUnapplied BIT, RemitAdvice BIT;

UPDATE r SET
    -- A single-line receipt left unmatched is the purest case: cash in the
    -- bank with no application at all against it. Excluding it would have
    -- removed the very situation the control exists to catch.
    r.HasUnapplied = CASE WHEN r.Lines >= 2 THEN CASE WHEN ru.r < 0.040 THEN 1 ELSE 0 END
                          ELSE CASE WHEN ru.r < 0.0008 THEN 1 ELSE 0 END END,
    -- the magnitude draws on its own hash: reusing the branch random makes
    -- every value inside a branch identical
    r.ApplyLagDays = CASE WHEN rl.r < 0.80 THEN 0
                          WHEN rl.r < 0.95 THEN 1 + CAST(FLOOR(rm.r * 3) AS INT)
                          ELSE 4 + CAST(FLOOR(rm.r * 6) AS INT) END
FROM #Rcpt r
CROSS APPLY dbo.fn_Rand(CONCAT('rcpt|', r.CustomerKey, '|', CONVERT(CHAR(8), r.PayDate, 112), '|unapp')) ru
CROSS APPLY dbo.fn_Rand(CONCAT('rcpt|', r.CustomerKey, '|', CONVERT(CHAR(8), r.PayDate, 112), '|lag'))   rl
CROSS APPLY dbo.fn_Rand(CONCAT('rcpt|', r.CustomerKey, '|', CONVERT(CHAR(8), r.PayDate, 112), '|lagmag')) rm;

-- A receipt with no remittance advice is the usual reason cash sits unapplied.
UPDATE #Rcpt SET RemitAdvice = CASE WHEN HasUnapplied = 1 THEN 0 ELSE 1 END;

-- Receipts land in the bank on or before the reporting date. Applications may
-- legitimately fall after it -- that is exactly the December-banked,
-- January-applied case -- so they are not filtered here.
DELETE FROM #Rcpt WHERE PayDate > '2025-12-31';
DELETE FROM #Pay  WHERE PayDate > '2025-12-31';

UPDATE r SET r.ReceiptNo = 'RCT-' + RIGHT('0000000' + CAST(x.rn AS VARCHAR(7)), 7)
FROM #Rcpt r
JOIN (SELECT CustomerKey, PayDate,
             rn = ROW_NUMBER() OVER (ORDER BY PayDate, CustomerKey)
      FROM #Rcpt) x
  ON x.CustomerKey = r.CustomerKey AND x.PayDate = r.PayDate;

INSERT INTO dbo.Fact_CashReceipt
    (ReceiptNo, CustomerKey, ReceiptDateKey, ReceiptAmount, PaymentMethod, RemittanceAdviceReceived, BankBatchID)
SELECT r.ReceiptNo, r.CustomerKey,
       YEAR(r.PayDate)*10000 + MONTH(r.PayDate)*100 + DAY(r.PayDate),
       r.ReceiptAmount,
       CASE WHEN r.Segment IN ('OEM Manufacturer','Institutional')
            THEN CASE WHEN r.rMethod < 0.30 THEN 'Wire' WHEN r.rMethod < 0.90 THEN 'ACH' ELSE 'Check' END
            ELSE CASE WHEN r.rMethod < 0.50 THEN 'ACH' WHEN r.rMethod < 0.85 THEN 'Check' ELSE 'Card' END END,
       r.RemitAdvice,
       'BB-' + CONVERT(CHAR(8), r.PayDate, 112)
FROM #Rcpt r
ORDER BY r.PayDate, r.CustomerKey;

-- Mark the one line on each unapplied receipt that never gets matched. The cash
-- is still in the receipt total; only the application is missing.
ALTER TABLE #Pay ADD IsUnapplied BIT;
UPDATE #Pay SET IsUnapplied = 0;

UPDATE p SET p.IsUnapplied = 1
FROM #Pay p
JOIN (SELECT CustomerKey, PayDate, InvoiceKey, Seq,
             rn = ROW_NUMBER() OVER (PARTITION BY CustomerKey, PayDate ORDER BY PayAmount, InvoiceKey)
      FROM #Pay) x
  ON  x.CustomerKey = p.CustomerKey AND x.PayDate = p.PayDate
  AND x.InvoiceKey  = p.InvoiceKey  AND x.Seq     = p.Seq
JOIN #Rcpt r ON r.CustomerKey = p.CustomerKey AND r.PayDate = p.PayDate
WHERE r.HasUnapplied = 1 AND x.rn = 1;

INSERT INTO dbo.Fact_CashApplication
    (ReceiptKey, InvoiceKey, CustomerKey, ApplicationDateKey, AppliedAmount, DiscountTaken)
SELECT fr.ReceiptKey, p.InvoiceKey, p.CustomerKey,
       YEAR(DATEADD(DAY, r.ApplyLagDays, p.PayDate))*10000
     + MONTH(DATEADD(DAY, r.ApplyLagDays, p.PayDate))*100
     + DAY(DATEADD(DAY, r.ApplyLagDays, p.PayDate)),
       p.PayAmount, p.Discount
FROM #Pay p
JOIN #Rcpt r ON r.CustomerKey = p.CustomerKey AND r.PayDate = p.PayDate
JOIN dbo.Fact_CashReceipt fr ON fr.ReceiptNo = r.ReceiptNo
WHERE p.IsUnapplied = 0
ORDER BY p.PayDate, p.InvN, p.Seq;

-- -----------------------------------------------------------------------------
-- P10  PROMISES TO PAY.
--
-- Two populations, and the generator plants only the INTENT -- never the
-- outcome. Status is derived from cash in dbo.fn_PromiseStatus, so a promise
-- counts as kept only because money actually arrived.
--
--   Kept-shaped:   anchored on cash that really did land, promised a few days
--                  before it arrived.
--   Broken-shaped: made by accounts that were past due and then paid nothing,
--                  weighted towards the distressed cohort, whose promise
--                  behaviour collapses before they default.
-- -----------------------------------------------------------------------------
IF OBJECT_ID('tempdb..#Prom') IS NOT NULL DROP TABLE #Prom;
CREATE TABLE #Prom (
    CustomerKey INT, CollectorKey INT, MadeDate DATE, PayDate DATE, Amount DECIMAL(12,2));

-- (a) promises anchored on cash that actually arrived
INSERT INTO #Prom (CustomerKey, CollectorKey, MadeDate, PayDate, Amount)
SELECT a.CustomerKey, c.CollectorKey,
       DATEADD(DAY, -(4 + CAST(FLOOR(r2.r * 11) AS INT)), d.[Date]),
       DATEADD(DAY, -CAST(FLOOR(r3.r * 3) AS INT), d.[Date]),
       ROUND(a.AppliedAmount * (0.90 + r4.r * 0.10), 2)
FROM dbo.Fact_CashApplication a
JOIN dbo.Dim_Date d  ON d.DateKey = a.ApplicationDateKey
JOIN dbo.Dim_Customer c ON c.CustomerKey = a.CustomerKey
JOIN dbo.Fact_Invoice f ON f.InvoiceKey = a.InvoiceKey
JOIN dbo.Dim_Date fd ON fd.DateKey = f.DueDateKey
CROSS APPLY dbo.fn_Rand(CONCAT('prom|keep|', a.ApplicationKey))  r1
CROSS APPLY dbo.fn_Rand(CONCAT('prom|made|', a.ApplicationKey))  r2
CROSS APPLY dbo.fn_Rand(CONCAT('prom|shift|', a.ApplicationKey)) r3
CROSS APPLY dbo.fn_Rand(CONCAT('prom|amt|', a.ApplicationKey))   r4
WHERE d.[Date] BETWEEN '2025-01-01' AND '2025-12-20'
  AND d.[Date] > fd.[Date]                 -- only late invoices get chased
  AND r1.r < 0.10
  AND a.AppliedAmount > 500;

-- (b) promises that nothing came in against
INSERT INTO #Prom (CustomerKey, CollectorKey, MadeDate, PayDate, Amount)
SELECT f.CustomerKey, c.CollectorKey,
       -- the promised date is derived FROM the made date, so it can never fall
       -- before it and the table's own CHECK constraint cannot be tripped
       DATEADD(DAY, 20 + CAST(FLOOR(r2.r * 40) AS INT), fd.[Date]),
       DATEADD(DAY, 20 + CAST(FLOOR(r2.r * 40) AS INT)
                       + 8 + CAST(FLOOR(r3.r * 18) AS INT), fd.[Date]),
       ROUND(f.InvoiceAmount * (0.45 + r4.r * 0.5), 2)
FROM dbo.Fact_Invoice f
JOIN dbo.Dim_Customer c ON c.CustomerKey = f.CustomerKey
JOIN dbo.Dim_Date fd ON fd.DateKey = f.DueDateKey
CROSS APPLY dbo.fn_Rand(CONCAT('prom|break|', f.InvoiceKey))  r1
CROSS APPLY dbo.fn_Rand(CONCAT('prom|bmade|', f.InvoiceKey))  r2
CROSS APPLY dbo.fn_Rand(CONCAT('prom|bpay|', f.InvoiceKey))   r3
CROSS APPLY dbo.fn_Rand(CONCAT('prom|bamt|', f.InvoiceKey))   r4
WHERE NOT EXISTS (SELECT 1 FROM dbo.Fact_CashApplication a WHERE a.InvoiceKey = f.InvoiceKey)
  AND fd.[Date] BETWEEN '2025-01-01' AND '2025-11-15'
  AND r1.r < CASE c.RiskTier WHEN 'High' THEN 0.90 WHEN 'Medium' THEN 0.60 ELSE 0.35 END
  AND f.InvoiceAmount > 500;

DELETE FROM #Prom WHERE PayDate > '2025-12-31' OR MadeDate < '2024-01-01' OR MadeDate > PayDate;

INSERT INTO dbo.Fact_PromiseToPay
    (PromiseNo, CustomerKey, CollectorKey, PromiseMadeKey, PromisedPayDateKey, PromisedAmount)
SELECT 'PTP-' + RIGHT('000000' + CAST(ROW_NUMBER() OVER (ORDER BY PayDate, CustomerKey, Amount) AS VARCHAR(6)), 6),
       CustomerKey, CollectorKey,
       YEAR(MadeDate)*10000 + MONTH(MadeDate)*100 + DAY(MadeDate),
       YEAR(PayDate)*10000  + MONTH(PayDate)*100  + DAY(PayDate),
       Amount
FROM #Prom;

-- =============================================================================
-- 8. Known data-quality defects (small, counted, documented) for 04 to catch
--    Selected deterministically by hashing the key, never randomly.
-- =============================================================================

-- DQ-A  DUPLICATE_RECEIPT: the same receipt posted twice under a new number
--       (~0.3%). Both the bank event and its applications are duplicated,
--       because a re-keyed receipt carries its matching with it.
IF OBJECT_ID('tempdb..#Dup') IS NOT NULL DROP TABLE #Dup;
SELECT r.ReceiptKey,
       NewNo = 'RCT-D' + RIGHT('000000' + CAST(ROW_NUMBER() OVER (ORDER BY r.ReceiptKey) AS VARCHAR(6)), 6)
INTO #Dup
FROM dbo.Fact_CashReceipt r
CROSS APPLY dbo.fn_Rand(CONCAT('dq|dup|', r.ReceiptKey)) x
WHERE x.r < 0.003;

INSERT INTO dbo.Fact_CashReceipt
    (ReceiptNo, CustomerKey, ReceiptDateKey, ReceiptAmount, PaymentMethod, RemittanceAdviceReceived, BankBatchID)
SELECT d.NewNo, r.CustomerKey, r.ReceiptDateKey, r.ReceiptAmount, r.PaymentMethod, r.RemittanceAdviceReceived, r.BankBatchID
FROM #Dup d JOIN dbo.Fact_CashReceipt r ON r.ReceiptKey = d.ReceiptKey;

INSERT INTO dbo.Fact_CashApplication
    (ReceiptKey, InvoiceKey, CustomerKey, ApplicationDateKey, AppliedAmount, DiscountTaken)
SELECT nr.ReceiptKey, a.InvoiceKey, a.CustomerKey, a.ApplicationDateKey, a.AppliedAmount, 0
FROM #Dup d
JOIN dbo.Fact_CashApplication a ON a.ReceiptKey = d.ReceiptKey
JOIN dbo.Fact_CashReceipt nr ON nr.ReceiptNo = d.NewNo;

-- DQ-B  RECEIPT_BEFORE_INVOICE: receipt dated before the invoice it settles (~0.1%)
UPDATE r SET r.ReceiptDateKey =
       YEAR(DATEADD(DAY, -7, d.[Date]))*10000 + MONTH(DATEADD(DAY, -7, d.[Date]))*100 + DAY(DATEADD(DAY, -7, d.[Date]))
FROM dbo.Fact_CashReceipt r
JOIN dbo.Fact_CashApplication a ON a.ReceiptKey = r.ReceiptKey
JOIN dbo.Fact_Invoice f ON f.InvoiceKey = a.InvoiceKey
JOIN dbo.Dim_Date d ON d.DateKey = f.InvoiceDateKey
CROSS APPLY dbo.fn_Rand(CONCAT('dq|early|', r.ReceiptKey)) x
WHERE x.r < 0.001 AND r.ReceiptNo NOT LIKE 'RCT-D%' AND d.[Date] >= '2024-01-08';

-- DQ-C  DUE_DATE_TERMS_MISMATCH: due date keyed as invoice + 10 days regardless of terms (~0.2%)
UPDATE f SET f.DueDateKey = YEAR(DATEADD(DAY, 10, d.[Date]))*10000 + MONTH(DATEADD(DAY, 10, d.[Date]))*100 + DAY(DATEADD(DAY, 10, d.[Date]))
FROM dbo.Fact_Invoice f
JOIN dbo.Dim_Date d ON d.DateKey = f.InvoiceDateKey
JOIN dbo.Dim_PaymentTerms t ON t.TermsKey = f.TermsKey
CROSS APPLY dbo.fn_Rand(CONCAT('dq|due|', f.InvoiceKey)) r
WHERE r.r < 0.002 AND t.NetDays > 10;

-- DQ-D  PRE_BILLING: invoice issued before the goods shipped (~0.1%).
--       Not clipped to a zero lag: a negative lag is a revenue-recognition
--       question, and silently taking MAX(0, lag) would erase it.
UPDATE f SET f.ShipDateKey = YEAR(DATEADD(DAY, 2, d.[Date]))*10000 + MONTH(DATEADD(DAY, 2, d.[Date]))*100 + DAY(DATEADD(DAY, 2, d.[Date]))
FROM dbo.Fact_Invoice f
JOIN dbo.Dim_Date d ON d.DateKey = f.InvoiceDateKey
CROSS APPLY dbo.fn_Rand(CONCAT('dq|prebill|', f.InvoiceKey)) r
WHERE r.r < 0.001;

DROP TABLE #Cust; DROP TABLE #Inv; DROP TABLE #Adj; DROP TABLE #Pay;
DROP TABLE #Rcpt; DROP TABLE #Prom; DROP TABLE #Dup;
GO

PRINT '=== VantageAR synthetic data generation complete ===';
SELECT 'Dim_Date' AS TableName, COUNT(*) AS [Rows] FROM dbo.Dim_Date
UNION ALL SELECT 'Dim_PaymentTerms',    COUNT(*) FROM dbo.Dim_PaymentTerms
UNION ALL SELECT 'Dim_Collector',       COUNT(*) FROM dbo.Dim_Collector
UNION ALL SELECT 'Dim_Customer',        COUNT(*) FROM dbo.Dim_Customer
UNION ALL SELECT 'Fact_Invoice',        COUNT(*) FROM dbo.Fact_Invoice
UNION ALL SELECT 'Fact_CashReceipt',    COUNT(*) FROM dbo.Fact_CashReceipt
UNION ALL SELECT 'Fact_CashApplication',COUNT(*) FROM dbo.Fact_CashApplication
UNION ALL SELECT 'Fact_Adjustment',     COUNT(*) FROM dbo.Fact_Adjustment
UNION ALL SELECT 'Fact_PromiseToPay',   COUNT(*) FROM dbo.Fact_PromiseToPay;
GO
