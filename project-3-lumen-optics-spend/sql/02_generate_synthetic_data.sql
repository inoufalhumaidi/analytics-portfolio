/*
================================================================================
Project 3 -- Lumen Optics Manufacturing: Photonics Spend Scorecard
Script:  02_generate_synthetic_data.sql
Purpose: Generate a reproducible three-year direct-materials procurement
         dataset with deliberate, documented behaviour for the analysis to
         find -- and a small number of deliberate data defects for the quality
         layer to catch.

DETERMINISM
    Every random draw comes from dbo.fn_Rand(<stable text key>), which hashes
    the key with SHA2_256 and scales the result into [0,1). It is therefore a
    pure function of the key: dropping and regenerating the database reproduces
    the dataset byte for byte, and every figure quoted in the case study stays
    checkable by a reader. RAND() and NEWID() cannot do this.

THE PRICE MODEL (the analytical core -- worth reading before the code)

    Photonic components follow learning curves. For each part:

        ExpectedPrice(t) = BasePrice x (1 - CategoryErosion) ^ years
        ActualPrice(t)   = BasePrice x (1 - VendorErosion)   ^ years x (1 + noise)

    CategoryErosion comes from Ref_PriceErosionBenchmark and is what the market
    should deliver. VendorErosion is how each supplier actually behaves, and it
    is assigned per VENDOR-PART pair, because the same supplier can track the
    curve on a commodity and hold firm on a sole-source critical part.

    The gap between the two is the finding: a supplier holding price flat while
    its category erodes 8% a year is extracting margin every quarter, and no
    variance report catches it, because the price never went up.

DELIBERATE BEHAVIOUR PLANTED HERE
    P1  Erosion laggards -- vendor-part pairs whose price barely moves while
        the category declines. Weighted towards sole-source critical parts,
        because that is where a supplier can afford to hold firm.
    P2  Cheap-but-rejected -- one laser diode vendor consistently undercuts on
        unit price and consistently fails beam-profile inspection, so it is the
        most expensive supplier per ACCEPTED unit.
    P3  Maverick spend -- lines placed with no agreement in force, concentrated
        in one buying team and in expedited situations.
    P4  Expedite creep -- expedite fees rising through 2025 on long-lead parts.
    P5  Single-source concentration -- critical parts with one qualified
        supplier and long requalification, where the opportunity is largest and
        the leverage is smallest.
    P6  Lapsed agreements -- contracts that expired and were never renewed, so
        buying continued off-contract at drifting prices.
    P7  Delivery decay -- one vendor's on-time performance falls through 2025.

DATA DISCLOSURE: Lumen Optics Manufacturing is fictional; all data is
synthetic. No confidential data and no production system is involved.
================================================================================
*/

USE LumenSpend;
GO
SET NOCOUNT ON;
GO

-- =============================================================================
-- 0. Deterministic pseudo-random helper
-- =============================================================================
IF OBJECT_ID('dbo.fn_Rand', 'IF') IS NOT NULL DROP FUNCTION dbo.fn_Rand;
GO
CREATE FUNCTION dbo.fn_Rand (@Key VARCHAR(200))
RETURNS TABLE
AS RETURN
(
    -- top 4 bytes of the hash, scaled into [0,1)
    SELECT r = CAST(CONVERT(BIGINT, CONVERT(BINARY(4), HASHBYTES('SHA2_256', @Key))) AS FLOAT)
               / 4294967296.0
);
GO

DELETE FROM dbo.Fact_GoodsReceipt;
DELETE FROM dbo.Fact_PurchaseOrderLine;
DELETE FROM dbo.Fact_PriceAgreement;
DELETE FROM dbo.Dim_Part;
DELETE FROM dbo.Dim_Vendor;
DELETE FROM dbo.Dim_Buyer;
DELETE FROM dbo.Dim_Date;
GO

-- =============================================================================
-- 1. Dim_Date  2022-10-01 .. 2026-06-30
--    Starts before the first order so agreement validity windows that open
--    ahead of the first purchase still resolve, and runs past the reporting
--    date so promised dates in the future have a row to point at.
-- =============================================================================
;WITH d AS (
    SELECT CAST('2022-10-01' AS DATE) AS dt
    UNION ALL SELECT DATEADD(DAY, 1, dt) FROM d WHERE dt < '2026-06-30'
)
INSERT INTO dbo.Dim_Date (DateKey, [Date], [Year], [Quarter], [Month], MonthName, YearMonth, YearQuarter, MonthEndDate, IsMonthEnd, IsWeekend)
SELECT YEAR(dt)*10000 + MONTH(dt)*100 + DAY(dt), dt, YEAR(dt), DATEPART(QUARTER, dt), MONTH(dt),
       DATENAME(MONTH, dt), CONVERT(CHAR(7), dt, 126),
       CAST(YEAR(dt) AS CHAR(4)) + '-Q' + CAST(DATEPART(QUARTER, dt) AS CHAR(1)),
       EOMONTH(dt),
       CASE WHEN dt = EOMONTH(dt) THEN 1 ELSE 0 END,
       CASE WHEN DATENAME(WEEKDAY, dt) IN ('Saturday','Sunday') THEN 1 ELSE 0 END
FROM d
OPTION (MAXRECURSION 0);

-- =============================================================================
-- 2. Dim_Buyer  (six-person sourcing team, organised by category family)
-- =============================================================================
INSERT INTO dbo.Dim_Buyer (BuyerID, BuyerName, Team, IsActive) VALUES
('BUY-01','Adaeze Okonkwo',   'Optoelectronics', 1),
('BUY-02','Rafael Duarte',    'Optoelectronics', 1),
('BUY-03','Mei-Ling Chow',    'Optics & Coatings',1),
('BUY-04','Johan Lindberg',   'Optics & Coatings',1),
('BUY-05','Priya Raghunathan','Mechanical & Thermal',1),
('BUY-06','Tomas Varga',      'Mechanical & Thermal',1);

-- =============================================================================
-- 3. Dim_Vendor  (36 suppliers)
--    LumenRevenueSharePct is our estimate of how much of the supplier's revenue
--    we represent. It drives leverage in the priority queue: a 2% share means
--    walking away is a threat they can absorb.
-- =============================================================================
;WITH v AS (
    SELECT n FROM (VALUES (1),(2),(3),(4),(5),(6),(7),(8),(9),(10),(11),(12),
                          (13),(14),(15),(16),(17),(18),(19),(20),(21),(22),(23),(24),
                          (25),(26),(27),(28),(29),(30),(31),(32),(33),(34),(35),(36)) x(n)
),
vn AS (
    SELECT v.n,
           BaseName = CASE v.n % 12
             WHEN 0 THEN 'Aperture' WHEN 1 THEN 'Lumitek'  WHEN 2 THEN 'Photonica'
             WHEN 3 THEN 'Crystalline' WHEN 4 THEN 'Spectra' WHEN 5 THEN 'Helix Optics'
             WHEN 6 THEN 'Northlight' WHEN 7 THEN 'Quantex' WHEN 8 THEN 'Meridian Photonics'
             WHEN 9 THEN 'Oribel'   WHEN 10 THEN 'Vectra'   ELSE 'Kestrel Optical' END,
           Suffix = CASE (v.n / 12) WHEN 0 THEN ' GmbH' WHEN 1 THEN ' Industries' ELSE ' Technologies' END,
           Country = CASE v.n % 6 WHEN 0 THEN 'Germany' WHEN 1 THEN 'United States' WHEN 2 THEN 'Japan'
                                  WHEN 3 THEN 'Taiwan'  WHEN 4 THEN 'United Kingdom' ELSE 'Switzerland' END
    FROM v
)
INSERT INTO dbo.Dim_Vendor (VendorID, VendorName, Country, VendorTier, RelationshipStartDate, LumenRevenueSharePct, IsActive)
SELECT
    'VEN-' + RIGHT('000' + CAST(vn.n AS VARCHAR(3)), 3),
    vn.BaseName + vn.Suffix,
    vn.Country,
    CASE WHEN vn.n <= 6 THEN 'Strategic'
         WHEN vn.n <= 16 THEN 'Preferred'
         WHEN vn.n <= 28 THEN 'Approved'
         ELSE 'Transactional' END,
    DATEADD(MONTH, -CAST(6 + FLOOR(r1.r * 110) AS INT), '2026-01-01'),
    -- strategic suppliers are large and we are a small part of them; the small
    -- transactional shops depend on us far more, which is where leverage lives
    CAST(CASE WHEN vn.n <= 6 THEN 1.0 + r2.r * 4.0
              WHEN vn.n <= 16 THEN 3.0 + r2.r * 9.0
              WHEN vn.n <= 28 THEN 6.0 + r2.r * 16.0
              ELSE 10.0 + r2.r * 28.0 END AS DECIMAL(5,2)),
    1
FROM vn
CROSS APPLY dbo.fn_Rand(CONCAT('ven|', vn.n, '|start')) r1
CROSS APPLY dbo.fn_Rand(CONCAT('ven|', vn.n, '|share')) r2;

-- =============================================================================
-- 4. Dim_Part  (140 parts across eight photonic categories)
--
-- P5  Single-source concentration: critical parts are far more likely to have
--     one qualified supplier and a long requalification, because qualifying a
--     second source for a laser diode means re-running reliability testing.
-- =============================================================================
IF OBJECT_ID('tempdb..#PartSeed') IS NOT NULL DROP TABLE #PartSeed;
;WITH p AS (
    SELECT TOP (140) n = ROW_NUMBER() OVER (ORDER BY (SELECT NULL))
    FROM sys.all_objects
)
SELECT
    p.n,
    Category = CASE p.n % 8
      WHEN 0 THEN 'Laser Diode'     WHEN 1 THEN 'Photodetector'  WHEN 2 THEN 'Optical Fibre'
      WHEN 3 THEN 'Lens Assembly'   WHEN 4 THEN 'Optical Coating' WHEN 5 THEN 'Precision Mount'
      WHEN 6 THEN 'Fibre Connector' ELSE 'Thermal Control' END,
    rCrit = rc.r, rQual = rq.r, rPrice = rp.r, rFirst = rf.r
INTO #PartSeed
FROM p
CROSS APPLY dbo.fn_Rand(CONCAT('part|', p.n, '|crit'))  rc
CROSS APPLY dbo.fn_Rand(CONCAT('part|', p.n, '|qual'))  rq
CROSS APPLY dbo.fn_Rand(CONCAT('part|', p.n, '|price')) rp
CROSS APPLY dbo.fn_Rand(CONCAT('part|', p.n, '|first')) rf;

ALTER TABLE #PartSeed ADD Criticality VARCHAR(10), QualCount TINYINT, QualMonths TINYINT,
                          BasePrice DECIMAL(12,4), FirstBuy DATE;

UPDATE #PartSeed SET
    Criticality = CASE WHEN rCrit < 0.22 THEN 'Critical'
                       WHEN rCrit < 0.68 THEN 'Standard' ELSE 'Commodity' END;

UPDATE #PartSeed SET
    -- a critical part usually has one or two qualified sources; a commodity has several
    QualCount  = CASE Criticality
                   WHEN 'Critical'  THEN CASE WHEN rQual < 0.55 THEN 1 WHEN rQual < 0.90 THEN 2 ELSE 3 END
                   WHEN 'Standard'  THEN CASE WHEN rQual < 0.18 THEN 1 WHEN rQual < 0.62 THEN 2 ELSE 3 END
                   ELSE                  CASE WHEN rQual < 0.05 THEN 1 WHEN rQual < 0.35 THEN 2 ELSE 3 END END,
    QualMonths = CASE Criticality WHEN 'Critical' THEN CAST(6 + rQual * 6 AS TINYINT)
                                  WHEN 'Standard' THEN CAST(2 + rQual * 4 AS TINYINT)
                                  ELSE CAST(rQual * 2 AS TINYINT) END,
    BasePrice  = ROUND(CASE Category
                   WHEN 'Laser Diode'     THEN  180 + rPrice * 900
                   WHEN 'Photodetector'   THEN   95 + rPrice * 420
                   WHEN 'Optical Fibre'   THEN   12 + rPrice *  48
                   WHEN 'Lens Assembly'   THEN   60 + rPrice * 340
                   WHEN 'Optical Coating' THEN   28 + rPrice * 130
                   WHEN 'Precision Mount' THEN   44 + rPrice * 210
                   WHEN 'Fibre Connector' THEN    8 + rPrice *  34
                   ELSE                          70 + rPrice * 260 END, 4),
    FirstBuy   = DATEADD(DAY, CAST(FLOOR(rFirst * 120) AS INT), '2023-01-02');

INSERT INTO dbo.Dim_Part (PartNumber, PartName, Category, UnitOfMeasure, Criticality,
                          QualifiedSupplierCount, QualificationMonths, FirstPurchaseDate)
SELECT
    'LOM-' + RIGHT('0000' + CAST(s.n AS VARCHAR(4)), 4),
    s.Category + ' '
      + CASE s.Criticality WHEN 'Critical' THEN 'Assy' WHEN 'Standard' THEN 'Module' ELSE 'Component' END
      + ' ' + RIGHT('0000' + CAST(s.n AS VARCHAR(4)), 4),
    s.Category,
    CASE WHEN s.Category = 'Optical Fibre' THEN 'metre' ELSE 'each' END,
    s.Criticality, s.QualCount, s.QualMonths, s.FirstBuy
FROM #PartSeed s
ORDER BY s.n;

-- =============================================================================
-- 5. Which vendors supply which parts, and how each behaves on price
--
-- P1  Erosion laggards. A vendor's behaviour is assigned per VENDOR-PART pair:
--     the same supplier can track the learning curve on a commodity and hold
--     firm on a sole-source critical part, because leverage differs by part,
--     not by supplier.
-- =============================================================================
IF OBJECT_ID('tempdb..#PartVendor') IS NOT NULL DROP TABLE #PartVendor;

SELECT
    p.PartKey, p.PartNumber, p.Category, p.Criticality,
    p.QualifiedSupplierCount, p.FirstPurchaseDate,
    v.VendorKey, v.VendorID, v.VendorTier,
    SupplierRank = rn.rn,
    b.AnnualErosionPct,
    rBehav = rb.r, rShare = rs.r
INTO #PartVendor
FROM dbo.Dim_Part p
JOIN dbo.Ref_PriceErosionBenchmark b ON b.Category = p.Category
CROSS APPLY (
    -- pick QualifiedSupplierCount vendors for this part, deterministically
    SELECT TOP (p.QualifiedSupplierCount) v2.VendorKey, v2.VendorID, v2.VendorTier,
           rn = ROW_NUMBER() OVER (ORDER BY rr.r)
    FROM dbo.Dim_Vendor v2
    CROSS APPLY dbo.fn_Rand(CONCAT('pv|', p.PartKey, '|', v2.VendorKey)) rr
    ORDER BY rr.r
) v
CROSS APPLY (SELECT rn = v.rn) rn
CROSS APPLY dbo.fn_Rand(CONCAT('behav|', p.PartKey, '|', v.VendorKey)) rb
CROSS APPLY dbo.fn_Rand(CONCAT('share|', p.PartKey, '|', v.VendorKey)) rs;

ALTER TABLE #PartVendor ADD VendorErosionPct DECIMAL(6,3), BehaviourLabel VARCHAR(20);

UPDATE #PartVendor SET
    -- A supplier can hold price when the buyer has nowhere else to go. Sole
    -- source raises the chance of "flat"; a competitive part lowers it.
    BehaviourLabel = CASE
        WHEN QualifiedSupplierCount = 1 AND rBehav < 0.62 THEN 'Flat'
        WHEN QualifiedSupplierCount = 1 AND rBehav < 0.88 THEN 'Partial'
        WHEN QualifiedSupplierCount = 2 AND rBehav < 0.26 THEN 'Flat'
        WHEN QualifiedSupplierCount = 2 AND rBehav < 0.62 THEN 'Partial'
        WHEN QualifiedSupplierCount >= 3 AND rBehav < 0.10 THEN 'Flat'
        WHEN QualifiedSupplierCount >= 3 AND rBehav < 0.40 THEN 'Partial'
        WHEN rBehav < 0.90 THEN 'Tracking'
        ELSE 'Aggressive' END;

UPDATE #PartVendor SET
    VendorErosionPct = CASE BehaviourLabel
        WHEN 'Flat'       THEN AnnualErosionPct * (0.00 + rBehav * 0.18)
        WHEN 'Partial'    THEN AnnualErosionPct * (0.35 + rBehav * 0.25)
        WHEN 'Tracking'   THEN AnnualErosionPct * (0.90 + rBehav * 0.20)
        ELSE                   AnnualErosionPct * (1.15 + rBehav * 0.25) END;

-- P2  Cheap-but-rejected: one vendor undercuts on laser diodes and fails
--     inspection often enough to be the most expensive per accepted unit.
--
--     The vendor is DERIVED from the data, not named. Vendor-part assignment
--     is random, so naming a vendor up front risks picking one that supplies
--     none of the category -- and an UPDATE that matches nothing does not
--     error, it silently does nothing, which is exactly how this pattern
--     failed to appear on the first run.
IF OBJECT_ID('tempdb..#Flags') IS NOT NULL DROP TABLE #Flags;
SELECT TOP 1 CheapVendorKey = pv.VendorKey, CheapVendorID = v.VendorID
INTO #Flags
FROM #PartVendor pv
JOIN dbo.Dim_Vendor v ON v.VendorKey = pv.VendorKey
WHERE pv.Category = 'Laser Diode' AND pv.SupplierRank = 1
GROUP BY pv.VendorKey, v.VendorID
ORDER BY COUNT(*) DESC, pv.VendorKey;

UPDATE pv SET pv.VendorErosionPct = pv.AnnualErosionPct * 1.9, pv.BehaviourLabel = 'Aggressive'
FROM #PartVendor pv
WHERE pv.VendorKey = (SELECT CheapVendorKey FROM #Flags) AND pv.Category = 'Laser Diode';
GO

-- =============================================================================
-- 6. Fact_PriceAgreement
--
-- Contracts run for twelve months and hold a FIXED price for the term, which is
-- how photonic components are actually bought. That matters analytically: the
-- learning-curve erosion does not arrive continuously, it arrives at renewal.
-- A supplier that renews flat keeps a year of margin that the category gave up.
--
-- P6  Lapsed agreements: roughly one in seven contracts is not renewed, and
--     buying simply continues off-contract at drifting spot prices. That is how
--     maverick spend and price erosion failure turn out to be the same story.
-- =============================================================================
IF OBJECT_ID('tempdb..#Agree') IS NOT NULL DROP TABLE #Agree;

SELECT
    pv.PartKey, pv.VendorKey, pv.SupplierRank, pv.BasePrice, pv.VendorErosionPct,
    TermNo = t.TermNo,
    ValidFrom = DATEADD(YEAR, t.TermNo - 1, pv.FirstPurchaseDate),
    rLapse = rl.r
INTO #Agree
FROM (SELECT pv.*, p.BasePrice
      FROM #PartVendor pv
      JOIN (SELECT p.PartKey, s.BasePrice FROM dbo.Dim_Part p JOIN #PartSeed s
            ON p.PartNumber = 'LOM-' + RIGHT('0000' + CAST(s.n AS VARCHAR(4)), 4)) p
        ON p.PartKey = pv.PartKey
      WHERE pv.SupplierRank <= 2) pv
CROSS JOIN (VALUES (1),(2),(3)) t(TermNo)
CROSS APPLY dbo.fn_Rand(CONCAT('agr|', pv.PartKey, '|', pv.VendorKey, '|', t.TermNo)) rl
WHERE DATEADD(YEAR, t.TermNo - 1, pv.FirstPurchaseDate) <= '2025-12-31';

-- A lapsed term takes every later term with it: nobody renews a contract that
-- was already allowed to expire two years ago. Record how many terms WERE
-- planned before deleting, because a truncated chain must end with a real
-- expiry date -- if the surviving last term is left open-ended, the contract
-- that lapsed becomes a contract that never expires, which is the opposite of
-- the pattern and leaves coverage looking perfect.
ALTER TABLE #Agree ADD PlannedTerms TINYINT, Truncated BIT;

UPDATE a SET a.PlannedTerms = x.MaxTerm
FROM #Agree a
JOIN (SELECT PartKey, VendorKey, MaxTerm = MAX(TermNo) FROM #Agree GROUP BY PartKey, VendorKey) x
  ON x.PartKey = a.PartKey AND x.VendorKey = a.VendorKey;

DELETE a
FROM #Agree a
WHERE EXISTS (SELECT 1 FROM #Agree a2
              WHERE a2.PartKey = a.PartKey AND a2.VendorKey = a.VendorKey
                AND a2.TermNo <= a.TermNo AND a2.TermNo > 1 AND a2.rLapse < 0.15);

UPDATE a SET a.Truncated = CASE WHEN y.MaxSurviving < a.PlannedTerms THEN 1 ELSE 0 END
FROM #Agree a
JOIN (SELECT PartKey, VendorKey, MaxSurviving = MAX(TermNo) FROM #Agree GROUP BY PartKey, VendorKey) y
  ON y.PartKey = a.PartKey AND y.VendorKey = a.VendorKey;

INSERT INTO dbo.Fact_PriceAgreement
    (AgreementNo, VendorKey, PartKey, ValidFromDateKey, ValidToDateKey, AgreedUnitPrice, MinOrderQty)
SELECT
    'PA-' + RIGHT('0000000' + CAST(ROW_NUMBER() OVER (ORDER BY a.PartKey, a.VendorKey, a.TermNo) AS VARCHAR(7)), 7),
    a.VendorKey, a.PartKey,
    YEAR(a.ValidFrom)*10000 + MONTH(a.ValidFrom)*100 + DAY(a.ValidFrom),
    -- the final surviving term is left open-ended, which is normal and is why
    -- the lookup must treat NULL as "still in force"
    CASE WHEN a.Truncated = 0
           AND a.TermNo = (SELECT MAX(a2.TermNo) FROM #Agree a2
                           WHERE a2.PartKey = a.PartKey AND a2.VendorKey = a.VendorKey)
         THEN NULL
         ELSE (SELECT YEAR(DATEADD(DAY,-1,DATEADD(YEAR,1,a.ValidFrom)))*10000
                    + MONTH(DATEADD(DAY,-1,DATEADD(YEAR,1,a.ValidFrom)))*100
                    + DAY(DATEADD(DAY,-1,DATEADD(YEAR,1,a.ValidFrom)))) END,
    -- the negotiated price is the point on the vendor's own erosion path at the
    -- moment the term opens, then held flat for the year
    ROUND(a.BasePrice * POWER(1.0 - a.VendorErosionPct/100.0,
          DATEDIFF(DAY, DATEADD(YEAR, -(a.TermNo-1), a.ValidFrom), a.ValidFrom) / 365.25), 4),
    CASE WHEN a.SupplierRank = 1 THEN 25 ELSE 10 END
FROM #Agree a
WHERE a.ValidFrom >= '2022-10-01';

-- =============================================================================
-- 7. Fact_PurchaseOrderLine
--
-- P3  Maverick spend: a line placed when no agreement is in force. Concentrated
--     in the Mechanical & Thermal team and in expedited situations, because
--     that is how it actually happens -- somebody needs a part this week and
--     there is no contract, so they raise a one-off PO at whatever price the
--     supplier quotes.
-- P4  Expedite creep: expedite frequency and fee both rise through 2025.
-- =============================================================================
IF OBJECT_ID('tempdb..#POCand') IS NOT NULL DROP TABLE #POCand;

SELECT
    p.PartKey, p.Category, p.Criticality, ps.BasePrice, p.FirstPurchaseDate,
    OrderDate = d.[Date],
    Seq = s.Seq,
    rKeep = rk.r, rVend = rv.r, rQty = rq.r, rExp = re.r, rNoise = rn.r, rPPV = rp.r
INTO #POCand
FROM dbo.Dim_Part p
JOIN (SELECT p2.PartKey, s.BasePrice FROM dbo.Dim_Part p2 JOIN #PartSeed s
      ON p2.PartNumber = 'LOM-' + RIGHT('0000' + CAST(s.n AS VARCHAR(4)), 4)) ps
  ON ps.PartKey = p.PartKey
JOIN dbo.Dim_Date d ON d.[Date] BETWEEN p.FirstPurchaseDate AND '2025-12-31' AND d.IsWeekend = 0
CROSS JOIN (VALUES (1),(2)) s(Seq)
CROSS APPLY dbo.fn_Rand(CONCAT('po|', p.PartKey, '|', CONVERT(CHAR(8), d.[Date], 112), '|', s.Seq, '|keep'))  rk
CROSS APPLY dbo.fn_Rand(CONCAT('po|', p.PartKey, '|', CONVERT(CHAR(8), d.[Date], 112), '|', s.Seq, '|vend'))  rv
CROSS APPLY dbo.fn_Rand(CONCAT('po|', p.PartKey, '|', CONVERT(CHAR(8), d.[Date], 112), '|', s.Seq, '|qty'))   rq
CROSS APPLY dbo.fn_Rand(CONCAT('po|', p.PartKey, '|', CONVERT(CHAR(8), d.[Date], 112), '|', s.Seq, '|exp'))   re
CROSS APPLY dbo.fn_Rand(CONCAT('po|', p.PartKey, '|', CONVERT(CHAR(8), d.[Date], 112), '|', s.Seq, '|noise')) rn
CROSS APPLY dbo.fn_Rand(CONCAT('po|', p.PartKey, '|', CONVERT(CHAR(8), d.[Date], 112), '|', s.Seq, '|ppv'))   rp
WHERE rk.r < 0.030;      -- about one order per part every three weeks

IF OBJECT_ID('tempdb..#POLine') IS NOT NULL DROP TABLE #POLine;

SELECT
    c.*,
    -- the primary supplier takes most of the volume; the alternates share the rest
    VendorKey = COALESCE(pick.VendorKey, prim.VendorKey),
    VendorErosionPct = COALESCE(pick.VendorErosionPct, prim.VendorErosionPct),
    YearsSinceFirst = DATEDIFF(DAY, c.FirstPurchaseDate, c.OrderDate) / 365.25
INTO #POLine
FROM #POCand c
OUTER APPLY (SELECT TOP 1 pv.VendorKey, pv.VendorErosionPct FROM #PartVendor pv
             WHERE pv.PartKey = c.PartKey AND pv.SupplierRank = 1) prim
OUTER APPLY (SELECT TOP 1 pv.VendorKey, pv.VendorErosionPct FROM #PartVendor pv
             WHERE pv.PartKey = c.PartKey
               AND pv.SupplierRank = CASE WHEN c.rVend < 0.72 THEN 1
                                          WHEN c.rVend < 0.92 THEN 2 ELSE 3 END) pick;

ALTER TABLE #POLine ADD SpotPrice DECIMAL(12,4), AgreedPrice DECIMAL(12,4),
                        UnitPrice DECIMAL(12,4), OrderQty INT, BuyerKey INT,
                        IsExpedited BIT, ExpediteFee DECIMAL(12,2), FreightAmount DECIMAL(12,2),
                        PromisedDate DATE;

-- the price the market would quote today on this vendor's own erosion path
UPDATE #POLine SET
    SpotPrice = ROUND(BasePrice * POWER(1.0 - VendorErosionPct/100.0, YearsSinceFirst)
                      * (0.985 + rNoise * 0.03), 4);

-- the agreement in force at the ORDER date, if any (NULL ValidTo = still open)
UPDATE l SET l.AgreedPrice = a.AgreedUnitPrice
FROM #POLine l
CROSS APPLY (
    SELECT TOP 1 pa.AgreedUnitPrice
    FROM dbo.Fact_PriceAgreement pa
    WHERE pa.PartKey = l.PartKey AND pa.VendorKey = l.VendorKey
      AND pa.ValidFromDateKey <= YEAR(l.OrderDate)*10000 + MONTH(l.OrderDate)*100 + DAY(l.OrderDate)
      AND (pa.ValidToDateKey IS NULL
           OR pa.ValidToDateKey >= YEAR(l.OrderDate)*10000 + MONTH(l.OrderDate)*100 + DAY(l.OrderDate))
    ORDER BY pa.ValidFromDateKey DESC
) a;

UPDATE #POLine SET
    -- On contract the agreed price governs, except for the minority of lines
    -- that drift: volume breaks downward, rush buys upward. Off contract there
    -- is nothing to govern it, and the supplier quotes spot plus a premium.
    UnitPrice = CASE
        WHEN AgreedPrice IS NULL THEN ROUND(SpotPrice * (1.04 + rPPV * 0.14), 4)
        WHEN rPPV < 0.06         THEN ROUND(AgreedPrice * (0.94 + rPPV * 0.05), 4)
        WHEN rPPV > 0.93         THEN ROUND(AgreedPrice * (1.02 + (rPPV-0.93) * 0.60), 4)
        ELSE AgreedPrice END,
    OrderQty = CASE Category
        WHEN 'Optical Fibre'   THEN 200 + CAST(FLOOR(rQty * 1800) AS INT)
        WHEN 'Fibre Connector' THEN 100 + CAST(FLOOR(rQty *  900) AS INT)
        WHEN 'Optical Coating' THEN  40 + CAST(FLOOR(rQty *  260) AS INT)
        ELSE                          10 + CAST(FLOOR(rQty *  140) AS INT) END,
    -- P4 expedite creep: more frequent, and dearer, as 2025 wears on
    IsExpedited = CASE WHEN rExp < 0.035 + 0.055 * CASE WHEN OrderDate >= '2025-01-01'
                            THEN DATEDIFF(DAY, '2025-01-01', OrderDate) / 364.0 ELSE 0 END
                       THEN 1 ELSE 0 END,
    PromisedDate = DATEADD(DAY, CASE Criticality WHEN 'Critical' THEN 35 WHEN 'Standard' THEN 21 ELSE 12 END
                                + CAST(FLOOR(rQty * 18) AS INT), OrderDate);

UPDATE #POLine SET
    ExpediteFee = CASE WHEN IsExpedited = 1
                       THEN ROUND(UnitPrice * OrderQty * (0.030 + rExp * 0.075), 2) ELSE 0 END,
    FreightAmount = ROUND(UnitPrice * OrderQty * (0.006 + rNoise * 0.020), 2);

-- buyers own category families
UPDATE l SET l.BuyerKey = b.BuyerKey
FROM #POLine l
CROSS APPLY (
    SELECT TOP 1 b2.BuyerKey FROM dbo.Dim_Buyer b2
    WHERE b2.Team = CASE WHEN l.Category IN ('Laser Diode','Photodetector') THEN 'Optoelectronics'
                         WHEN l.Category IN ('Optical Fibre','Lens Assembly','Optical Coating','Fibre Connector') THEN 'Optics & Coatings'
                         ELSE 'Mechanical & Thermal' END
    ORDER BY CASE WHEN l.rVend < 0.5 THEN b2.BuyerKey ELSE -b2.BuyerKey END
) b;

INSERT INTO dbo.Fact_PurchaseOrderLine
    (PONumber, POLineNo, PartKey, VendorKey, BuyerKey, OrderDateKey, PromisedDateKey,
     OrderQty, UnitPrice, FreightAmount, ExpediteFee, IsExpedited)
SELECT
    'PO-' + RIGHT('000000' + CAST(x.POSeq AS VARCHAR(6)), 6),
    1,
    x.PartKey, x.VendorKey, x.BuyerKey,
    YEAR(x.OrderDate)*10000 + MONTH(x.OrderDate)*100 + DAY(x.OrderDate),
    YEAR(x.PromisedDate)*10000 + MONTH(x.PromisedDate)*100 + DAY(x.PromisedDate),
    x.OrderQty, x.UnitPrice, x.FreightAmount, x.ExpediteFee, x.IsExpedited
FROM (SELECT l.*, POSeq = ROW_NUMBER() OVER (ORDER BY l.OrderDate, l.PartKey, l.Seq) FROM #POLine l) x;

-- =============================================================================
-- 8. Fact_GoodsReceipt
--
-- P2  Cheap-but-rejected. VEN-021 undercuts everyone on laser diodes and fails
--     beam-profile inspection often enough that its cost per ACCEPTED unit is
--     the worst on the category. Unit price alone cannot see this.
-- P7  Delivery decay: VEN-009 slips progressively later through 2025.
-- =============================================================================
IF OBJECT_ID('tempdb..#Rcpt') IS NOT NULL DROP TABLE #Rcpt;

SELECT
    l.POLineKey, l.PartKey, l.VendorKey, l.OrderQty, l.PromisedDateKey,
    p.Category, v.VendorID,
    OrderDate = do.[Date], PromisedDate = dp.[Date],
    rLate = r1.r, rRej = r2.r, rReason = r3.r, rSplit = r4.r
INTO #Rcpt
FROM dbo.Fact_PurchaseOrderLine l
JOIN dbo.Dim_Part p   ON p.PartKey = l.PartKey
JOIN dbo.Dim_Vendor v ON v.VendorKey = l.VendorKey
JOIN dbo.Dim_Date do  ON do.DateKey = l.OrderDateKey
JOIN dbo.Dim_Date dp  ON dp.DateKey = l.PromisedDateKey
CROSS APPLY dbo.fn_Rand(CONCAT('rc|', l.POLineKey, '|late'))   r1
CROSS APPLY dbo.fn_Rand(CONCAT('rc|', l.POLineKey, '|rej'))    r2
CROSS APPLY dbo.fn_Rand(CONCAT('rc|', l.POLineKey, '|reason')) r3
CROSS APPLY dbo.fn_Rand(CONCAT('rc|', l.POLineKey, '|split'))  r4;

ALTER TABLE #Rcpt ADD ReceiptDate DATE, RejectRate FLOAT, QtyRejected INT;

UPDATE #Rcpt SET
    ReceiptDate = DATEADD(DAY,
        CASE
          -- P7 one vendor's delivery decays through 2025
          WHEN VendorID = 'VEN-009' AND PromisedDate >= '2025-01-01'
               THEN CAST(FLOOR(rLate * (4 + 18 * DATEDIFF(DAY,'2025-01-01',PromisedDate)/364.0)) AS INT) - 2
          WHEN rLate < 0.82 THEN -CAST(FLOOR(rLate * 5) AS INT)
          ELSE CAST(FLOOR((rLate - 0.82) * 60) AS INT)
        END, PromisedDate);

UPDATE #Rcpt SET
    RejectRate = CASE
        -- P2 the cheap laser diode supplier, derived in section 5
        WHEN VendorID = (SELECT CheapVendorID FROM #Flags) AND Category = 'Laser Diode'
             THEN 0.055 + rRej * 0.055
        WHEN Category IN ('Optical Coating','Lens Assembly')   THEN 0.004 + rRej * 0.022
        WHEN Category = 'Optical Fibre'                        THEN 0.002 + rRej * 0.010
        ELSE                                                        0.002 + rRej * 0.014 END;

UPDATE #Rcpt SET QtyRejected = CAST(ROUND(OrderQty * RejectRate, 0) AS INT);
UPDATE #Rcpt SET QtyRejected = 0 WHERE QtyRejected < 0;
UPDATE #Rcpt SET QtyRejected = OrderQty WHERE QtyRejected > OrderQty;

-- receipts dated after the reporting horizon have not happened yet
DELETE FROM #Rcpt WHERE ReceiptDate > '2025-12-31';

INSERT INTO dbo.Fact_GoodsReceipt
    (ReceiptNo, POLineKey, ReceiptDateKey, QtyReceived, QtyAccepted, QtyRejected, RejectReasonCode)
SELECT
    'GR-' + RIGHT('0000000' + CAST(ROW_NUMBER() OVER (ORDER BY r.ReceiptDate, r.POLineKey) AS VARCHAR(7)), 7),
    r.POLineKey,
    YEAR(r.ReceiptDate)*10000 + MONTH(r.ReceiptDate)*100 + DAY(r.ReceiptDate),
    r.OrderQty, r.OrderQty - r.QtyRejected, r.QtyRejected,
    CASE WHEN r.QtyRejected = 0 THEN NULL
         WHEN r.Category = 'Laser Diode'     THEN CASE WHEN r.rReason < 0.70 THEN 'BEAM_PROFILE' ELSE 'DIMENSIONAL' END
         WHEN r.Category = 'Optical Coating' THEN CASE WHEN r.rReason < 0.75 THEN 'COATING_DEFECT' ELSE 'CONTAMINATION' END
         WHEN r.Category = 'Lens Assembly'   THEN CASE WHEN r.rReason < 0.65 THEN 'SURFACE_FIGURE' ELSE 'CONTAMINATION' END
         WHEN r.Category = 'Optical Fibre'   THEN CASE WHEN r.rReason < 0.80 THEN 'ATTENUATION' ELSE 'TRANSIT_DAMAGE' END
         WHEN r.rReason < 0.45 THEN 'DIMENSIONAL'
         WHEN r.rReason < 0.75 THEN 'TRANSIT_DAMAGE'
         ELSE 'SPEC_CHANGE' END
FROM #Rcpt r;

-- =============================================================================
-- 9. Known data-quality defects (small, counted, documented) for 04 to catch.
--    Selected deterministically by hashing the key, never randomly.
-- =============================================================================

-- DQ-A  DUPLICATE_PO_LINE: the same requisition raised twice under a new PO
--       number (~0.3%). Inflates committed spend and double-counts volume.
INSERT INTO dbo.Fact_PurchaseOrderLine
    (PONumber, POLineNo, PartKey, VendorKey, BuyerKey, OrderDateKey, PromisedDateKey,
     OrderQty, UnitPrice, FreightAmount, ExpediteFee, IsExpedited)
SELECT 'PO-D' + RIGHT('00000' + CAST(ROW_NUMBER() OVER (ORDER BY l.POLineKey) AS VARCHAR(5)), 5),
       1, l.PartKey, l.VendorKey, l.BuyerKey, l.OrderDateKey, l.PromisedDateKey,
       l.OrderQty, l.UnitPrice, l.FreightAmount, l.ExpediteFee, l.IsExpedited
FROM dbo.Fact_PurchaseOrderLine l
CROSS APPLY dbo.fn_Rand(CONCAT('dq|dup|', l.POLineKey)) r
WHERE r.r < 0.003;

-- DQ-B  RECEIPT_BEFORE_ORDER: goods booked in before the order was raised
--       (~0.1%). Usually a back-dated PO covering an informal delivery.
UPDATE g SET g.ReceiptDateKey =
       YEAR(DATEADD(DAY,-9,d.[Date]))*10000 + MONTH(DATEADD(DAY,-9,d.[Date]))*100 + DAY(DATEADD(DAY,-9,d.[Date]))
FROM dbo.Fact_GoodsReceipt g
JOIN dbo.Fact_PurchaseOrderLine l ON l.POLineKey = g.POLineKey
JOIN dbo.Dim_Date d ON d.DateKey = l.OrderDateKey
CROSS APPLY dbo.fn_Rand(CONCAT('dq|early|', g.ReceiptKey)) r
WHERE r.r < 0.001 AND d.[Date] >= '2023-02-01';

-- DQ-C  OVERLAPPING_AGREEMENT: two price agreements in force for the same
--       vendor-part at the same time (~1% of parts). Makes the contracted price
--       ambiguous, so purchase price variance silently depends on which row the
--       query happens to pick.
INSERT INTO dbo.Fact_PriceAgreement
    (AgreementNo, VendorKey, PartKey, ValidFromDateKey, ValidToDateKey, AgreedUnitPrice, MinOrderQty)
SELECT 'PA-X' + RIGHT('000000' + CAST(ROW_NUMBER() OVER (ORDER BY pa.AgreementKey) AS VARCHAR(6)), 6),
       pa.VendorKey, pa.PartKey,
       pa.ValidFromDateKey, pa.ValidToDateKey,
       ROUND(pa.AgreedUnitPrice * 1.035, 4), pa.MinOrderQty
FROM dbo.Fact_PriceAgreement pa
CROSS APPLY dbo.fn_Rand(CONCAT('dq|ovl|', pa.AgreementKey)) r
WHERE r.r < 0.010 AND pa.AgreementNo NOT LIKE 'PA-X%';

-- DQ-D  PRICE_OUTLIER: a decimal-point slip on entry (~0.05%). Ten times the
--       contracted price, which is obvious in a list and invisible in a total.
UPDATE l SET l.UnitPrice = l.UnitPrice * 10
FROM dbo.Fact_PurchaseOrderLine l
CROSS APPLY dbo.fn_Rand(CONCAT('dq|fat|', l.POLineKey)) r
WHERE r.r < 0.0005 AND l.PONumber NOT LIKE 'PO-D%';

DROP TABLE #PartSeed; DROP TABLE #PartVendor; DROP TABLE #Agree;
DROP TABLE #POCand;  DROP TABLE #POLine;      DROP TABLE #Rcpt;
DROP TABLE #Flags;
GO

PRINT '=== LumenSpend synthetic data generation complete ===';
SELECT 'Dim_Date' AS TableName, COUNT(*) AS [Rows] FROM dbo.Dim_Date
UNION ALL SELECT 'Dim_Buyer',              COUNT(*) FROM dbo.Dim_Buyer
UNION ALL SELECT 'Dim_Vendor',             COUNT(*) FROM dbo.Dim_Vendor
UNION ALL SELECT 'Dim_Part',               COUNT(*) FROM dbo.Dim_Part
UNION ALL SELECT 'Fact_PriceAgreement',    COUNT(*) FROM dbo.Fact_PriceAgreement
UNION ALL SELECT 'Fact_PurchaseOrderLine', COUNT(*) FROM dbo.Fact_PurchaseOrderLine
UNION ALL SELECT 'Fact_GoodsReceipt',      COUNT(*) FROM dbo.Fact_GoodsReceipt;
GO
