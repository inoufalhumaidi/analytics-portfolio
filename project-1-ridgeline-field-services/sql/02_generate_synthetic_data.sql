/*
================================================================================
Project 1 — Ridgeline Field Services: Operational Efficiency
Script:  02_generate_synthetic_data.sql
Purpose: Populates all dimension tables and generates ~2 years of synthetic
         HVAC/Electrical field-service job data with deliberate, documented
         operational patterns (regional workload imbalance, skill-driven
         first-time-fix variance, travel-driven SLA risk, seasonality) and a
         small, deliberate rate of data-quality anomalies for
         03_data_quality_checks.sql to detect.

DATA DISCLOSURE: 100% synthetic. Generation parameters (intensity factors,
travel baselines, probability weights) are documented inline and in
/docs/data_validation_report.md. No real company, technician, or client data
is used.
================================================================================
*/

USE RidgelineFieldOps;
GO

SET NOCOUNT ON;

-- =============================================================================
-- 1. Dim_Date  (2024-01-01 .. 2025-12-31)
-- =============================================================================
;WITH DateSeq AS (
    SELECT CAST('2024-01-01' AS DATE) AS [Date]
    UNION ALL
    SELECT DATEADD(DAY, 1, [Date]) FROM DateSeq WHERE [Date] < '2025-12-31'
)
INSERT INTO dbo.Dim_Date (DateKey, [Date], [Year], [Quarter], [Month], MonthName, [Week], DayOfWeek, DayName, IsWeekend)
SELECT
    CONVERT(INT, FORMAT([Date], 'yyyyMMdd'))                              AS DateKey,
    [Date],
    YEAR([Date])                                                          AS [Year],
    DATEPART(QUARTER, [Date])                                             AS [Quarter],
    MONTH([Date])                                                         AS [Month],
    DATENAME(MONTH, [Date])                                               AS MonthName,
    DATEPART(WEEK, [Date])                                                AS [Week],
    DATEPART(WEEKDAY, [Date])                                             AS DayOfWeek,
    DATENAME(WEEKDAY, [Date])                                             AS DayName,
    CASE WHEN DATEPART(WEEKDAY, [Date]) IN (1,7) THEN 1 ELSE 0 END        AS IsWeekend
FROM DateSeq
OPTION (MAXRECURSION 800);
GO

-- =============================================================================
-- 2. Dim_Region  (5 regions; intensity/travel assumptions live in this script
--    only — they are generation parameters, not warehouse attributes)
-- =============================================================================
INSERT INTO dbo.Dim_Region (RegionCode, RegionName, TechnicianTarget) VALUES
('NORTH',   'North Metro',      10),
('SOUTH',   'South Metro',       8),
('EAST',    'East Valley',       8),
('WEST',    'West Hills',        7),
('CENTRAL', 'Central District',  7);
GO

-- =============================================================================
-- 3. Dim_ServiceType (10 types: 5 HVAC, 5 Electrical)
-- =============================================================================
-- NOTE on SLAHours semantics: this is an ARRIVAL-WINDOW commitment (max
-- acceptable minutes/hours between ScheduledStart and ActualStart), which is
-- how most residential/commercial field-service SLAs are actually written
-- (e.g. "technician arrives within a 1-hour window"), not a multi-day
-- completion SLA. This is what makes dispatch delay/travel time an
-- SLA-relevant KPI driver.
INSERT INTO dbo.Dim_ServiceType (ServiceTypeID, ServiceCategory, ServiceName, StandardDurationMin, SLAHours, BaseRevenue, BaseCost) VALUES
('SVC-01','HVAC',      'AC Repair',                    90,  1.0,  220.00,  90.00),
('SVC-02','HVAC',      'Furnace Repair',                90,  1.0,  240.00,  95.00),
('SVC-03','HVAC',      'HVAC Installation',            240,  4.0, 1800.00, 900.00),
('SVC-04','HVAC',      'Duct Cleaning',                120,  4.0,  300.00, 120.00),
('SVC-05','HVAC',      'Preventive Maintenance (HVAC)', 60,  4.0,  150.00,  60.00),
('SVC-06','Electrical','Panel Repair',                  90,  1.0,  260.00, 100.00),
('SVC-07','Electrical','Wiring Installation',          180,  4.0,  900.00, 400.00),
('SVC-08','Electrical','Outlet/Switch Repair',          45,  2.0,  120.00,  45.00),
('SVC-09','Electrical','Generator Install',            300,  4.0, 2200.00,1100.00),
('SVC-10','Electrical','Electrical Safety Inspection',  60,  2.0,  180.00,  70.00);
GO

-- =============================================================================
-- 4. Dim_Technician (40 technicians across 5 regions)
-- =============================================================================
;WITH Names AS (
    SELECT * FROM (VALUES
    ('Alex'),('Jordan'),('Taylor'),('Morgan'),('Casey'),('Riley'),('Jamie'),('Cameron'),
    ('Drew'),('Skyler'),('Avery'),('Reese'),('Rowan'),('Quinn'),('Hayden'),('Emerson'),
    ('Dakota'),('Finley'),('Peyton'),('Sawyer'),('Elliot'),('Marlowe'),('Kendall'),('Blake'),
    ('Shawn'),('Terry'),('Robin'),('Lee'),('Devon'),('Jesse'),('Kai'),('Micah'),
    ('Noel'),('Sage'),('Tatum'),('Wren'),('Amari'),('Lior'),('Remy'),('Shay')
    ) AS n(FirstName)
),
LastNames AS (
    SELECT * FROM (VALUES
    ('Rivera'),('Chen'),('Kowalski'),('Nguyen'),('Osei'),('Petrov'),('Alvarez'),('Singh'),
    ('Novak'),('Dubois'),('Hartman'),('Okafor'),('Reyes'),('Sato'),('Weber'),('Larsen'),
    ('Mendez'),('Kim'),('Fischer'),('Brennan'),('Costa'),('Haddad'),('Volkov'),('Suzuki'),
    ('Bianchi'),('Muller'),('Park'),('Silva'),('Andersen'),('Duarte'),('Kaur'),('Yamamoto'),
    ('Moreau'),('Adeyemi'),('Craig'),('Lindqvist'),('Barros'),('Marchetti'),('Solis'),('Tanaka')
    ) AS n(LastName)
),
Numbered AS (
    SELECT ROW_NUMBER() OVER (ORDER BY (SELECT NULL)) AS rn, FirstName FROM Names
),
NumberedLast AS (
    SELECT ROW_NUMBER() OVER (ORDER BY (SELECT NULL)) AS rn, LastName FROM LastNames
),
RegionAlloc AS (
    -- Explicit headcount allocation matching each region's TechnicianTarget
    SELECT RegionKey, RegionCode, n AS slot
    FROM dbo.Dim_Region
    CROSS APPLY (SELECT TOP (TechnicianTarget) ROW_NUMBER() OVER (ORDER BY (SELECT NULL)) AS n FROM sys.all_objects) x
),
Ordered AS (
    SELECT ROW_NUMBER() OVER (ORDER BY RegionKey, slot) AS rn, RegionKey
    FROM RegionAlloc
)
INSERT INTO dbo.Dim_Technician (TechnicianID, TechnicianName, RegionKey, SkillLevel, HireDate, EmploymentStatus, DailyCapacityHours)
SELECT
    'TECH-' + RIGHT('000' + CAST(o.rn AS VARCHAR(3)), 3)                         AS TechnicianID,
    nl.FirstName + ' ' + nll.LastName                                             AS TechnicianName,
    o.RegionKey,
    CASE
        WHEN r.p <= 0.25 THEN 'Apprentice'
        WHEN r.p <= 0.75 THEN 'Journeyman'
        ELSE 'Master'
    END                                                                          AS SkillLevel,
    DATEADD(DAY, -ABS(CHECKSUM(NEWID())) % 2555, '2025-12-31')                   AS HireDate,   -- up to 7 yrs tenure
    CASE WHEN r.p2 <= 0.90 THEN 'Active' ELSE 'Inactive' END                     AS EmploymentStatus,
    8.0                                                                          AS DailyCapacityHours
FROM Ordered o
JOIN Numbered nl      ON nl.rn = o.rn
JOIN NumberedLast nll ON nll.rn = o.rn
CROSS APPLY (SELECT RAND(CHECKSUM(NEWID())) AS p, RAND(CHECKSUM(NEWID())) AS p2) r
JOIN dbo.Dim_Region r2 ON r2.RegionKey = o.RegionKey;
GO

-- =============================================================================
-- 5. Dim_Client (300 clients across regions)
-- =============================================================================
;WITH Tally AS (
    SELECT TOP (300) ROW_NUMBER() OVER (ORDER BY (SELECT NULL)) AS ClientSeq
    FROM sys.all_objects a CROSS JOIN sys.all_objects b
),
RegionCycle AS (
    SELECT t.ClientSeq,
           ((t.ClientSeq - 1) % 5) + 1 AS RegionKey
    FROM Tally t
)
INSERT INTO dbo.Dim_Client (ClientID, ClientName, ClientType, RegionKey, ContractTier)
SELECT
    'CLI-' + RIGHT('0000' + CAST(rc.ClientSeq AS VARCHAR(4)), 4)                   AS ClientID,
    CASE WHEN rr.p < 0.70
         THEN 'Residence #' + CAST(rc.ClientSeq AS VARCHAR(4))
         ELSE nm.v + ' - Site ' + CAST(rc.ClientSeq AS VARCHAR(4))
    END                                                                            AS ClientName,
    CASE WHEN rr.p < 0.70 THEN 'Residential' ELSE 'Commercial' END                 AS ClientType,
    rc.RegionKey,
    CASE WHEN rr.p2 <= 0.60 THEN 'Standard' WHEN rr.p2 <= 0.90 THEN 'Priority' ELSE 'Premium' END AS ContractTier
FROM RegionCycle rc
CROSS APPLY (SELECT RAND(CHECKSUM(NEWID())) AS p, RAND(CHECKSUM(NEWID())) AS p2) rr
CROSS APPLY (SELECT TOP 1 v FROM (VALUES ('Summit Retail Group'),('Harborview Offices'),
       ('Cascade Manufacturing'),('Blue Oak Medical Center'),('Union Square Hospitality'),
       ('Fairground Logistics'),('Riverside School District'),('Pinnacle Data Centers'),
       ('Greenway Property Mgmt'),('Anchor Point Bank')) AS x(v) ORDER BY NEWID()) nm;
GO

PRINT 'Dimension tables populated.';
GO

-- =============================================================================
-- 6. Fact_ServiceJobs
--    Generation model (documented for transparency / reproducibility):
--      - Candidate grain = Active Technician x Date x Slot(0..2)
--      - A candidate becomes a job if RAND() < RegionIntensity * DayOfWeekFactor
--      - RegionIntensity is a deliberate imbalance: SOUTH is underutilized
--        (0.55) vs NORTH near capacity (0.85) -> the core "wasted capacity"
--        story for the case study.
--      - Travel time baseline is highest in EAST (long geographic spread) ->
--        drives that region's SLA-compliance shortfall.
--      - First-time-fix probability is skill-driven (Apprentice < Journeyman
--        < Master) -> callback / rework story.
--      - Service-type mix shifts seasonally: HVAC weighted up in
--        Jun-Aug/Dec-Feb, Electrical the rest of the year.
-- =============================================================================

IF OBJECT_ID('tempdb..#Candidates') IS NOT NULL DROP TABLE #Candidates;

;WITH Slots AS (SELECT n FROM (VALUES (0),(1),(2)) s(n)),
RegionParams AS (
    SELECT RegionKey, RegionCode,
           CASE RegionCode
                WHEN 'NORTH'   THEN 0.85
                WHEN 'SOUTH'   THEN 0.55   -- deliberately underutilized
                WHEN 'EAST'    THEN 0.78
                WHEN 'WEST'    THEN 0.72
                WHEN 'CENTRAL' THEN 0.80
           END AS RegionIntensity,
           CASE RegionCode
                WHEN 'NORTH'   THEN 15
                WHEN 'SOUTH'   THEN 18
                WHEN 'EAST'    THEN 35     -- deliberately high travel -> SLA risk
                WHEN 'WEST'    THEN 22
                WHEN 'CENTRAL' THEN 12
           END AS TravelBaseMin
    FROM dbo.Dim_Region
)
SELECT
    t.TechnicianKey, t.RegionKey, t.SkillLevel,
    d.DateKey, d.[Date], d.[Month], d.DayOfWeek,
    rp.RegionCode, rp.RegionIntensity, rp.TravelBaseMin,
    sl.n AS SlotNum,
    rnd.keep_p, rnd.hour_p, rnd.min_p
INTO #Candidates
FROM dbo.Dim_Technician t
JOIN RegionParams rp ON rp.RegionKey = t.RegionKey
CROSS JOIN dbo.Dim_Date d
CROSS JOIN Slots sl
CROSS APPLY (SELECT RAND(CHECKSUM(NEWID())) AS keep_p,
                    RAND(CHECKSUM(NEWID())) AS hour_p,
                    RAND(CHECKSUM(NEWID())) AS min_p) rnd
WHERE t.EmploymentStatus = 'Active'
  AND d.[Date] >= t.HireDate;   -- a technician cannot be dispatched before being hired

-- Filter to kept slots using day-of-week factor (1=Sun,7=Sat)
IF OBJECT_ID('tempdb..#Kept') IS NOT NULL DROP TABLE #Kept;

SELECT *,
    CASE DayOfWeek WHEN 1 THEN 0.05 WHEN 7 THEN 0.30 ELSE 1.00 END AS DayFactor
INTO #Kept
FROM #Candidates c
WHERE c.keep_p < c.RegionIntensity *
      (CASE c.DayOfWeek WHEN 1 THEN 0.05 WHEN 7 THEN 0.30 ELSE 1.00 END);

DECLARE @KeptCount INT = (SELECT COUNT(*) FROM #Kept);
PRINT 'Candidate job slots kept: ' + CAST(@KeptCount AS VARCHAR(20));

-- Build final job rows with all derived attributes
IF OBJECT_ID('tempdb..#Jobs') IS NOT NULL DROP TABLE #Jobs;

;WITH Enriched AS (
    SELECT
        k.*,
        ROW_NUMBER() OVER (ORDER BY k.[Date], k.TechnicianKey, k.SlotNum) AS JobSeq,
        -- scheduled hour: business hours 7-16 driven by hour_p
        7 + FLOOR(k.hour_p * 10)                                         AS SchedHour,
        CASE FLOOR(k.min_p * 4) WHEN 0 THEN 0 WHEN 1 THEN 15 WHEN 2 THEN 30 ELSE 45 END AS SchedMin
    FROM #Kept k
),
WithPicks AS (
    SELECT
        e.*,
        st.ServiceTypeKey, st.ServiceCategory, st.StandardDurationMin, st.SLAHours, st.BaseRevenue, st.BaseCost,
        cl.ClientKey,
        rr.status_p, rr.ftf_p, rr.travel_noise, rr.delay_p, rr.dur_noise, rr.cost_noise
    FROM Enriched e
    CROSS APPLY (SELECT RAND(CHECKSUM(NEWID())) AS seasonal_p) sp
    CROSS APPLY (
        SELECT TOP 1 ServiceTypeKey, ServiceCategory, StandardDurationMin, SLAHours, BaseRevenue, BaseCost
        FROM dbo.Dim_ServiceType st
        WHERE st.ServiceCategory = CASE
                WHEN e.[Month] IN (6,7,8,12,1,2) THEN (CASE WHEN sp.seasonal_p < 0.70 THEN 'HVAC' ELSE 'Electrical' END)
                ELSE (CASE WHEN sp.seasonal_p < 0.40 THEN 'HVAC' ELSE 'Electrical' END)
             END
        ORDER BY NEWID()
    ) st
    CROSS APPLY (
        SELECT TOP 1 ClientKey
        FROM dbo.Dim_Client cl2
        WHERE (sp.seasonal_p < 0.85 AND cl2.RegionKey = e.RegionKey) OR (sp.seasonal_p >= 0.85)
        ORDER BY NEWID()
    ) cl
    CROSS APPLY (SELECT
        RAND(CHECKSUM(NEWID())) AS status_p,
        RAND(CHECKSUM(NEWID())) AS ftf_p,
        (RAND(CHECKSUM(NEWID())) * 16 - 8)  AS travel_noise,      -- +/-8 min
        RAND(CHECKSUM(NEWID())) AS delay_p,
        (RAND(CHECKSUM(NEWID())) * 0.30 - 0.15) AS dur_noise,     -- +/-15%
        (RAND(CHECKSUM(NEWID())) * 0.20 - 0.10) AS cost_noise     -- +/-10%
    ) rr
)
SELECT
    'JOB-' + RIGHT('000000' + CAST(JobSeq AS VARCHAR(6)), 6)                      AS JobID,
    DateKey,
    TechnicianKey,
    ClientKey,
    ServiceTypeKey,
    RegionKey,
    DATETIMEFROMPARTS(YEAR([Date]), MONTH([Date]), DAY([Date]), SchedHour, SchedMin, 0, 0) AS ScheduledStart,

    GREATEST(5, TravelBaseMin
        + travel_noise
        + CASE SkillLevel WHEN 'Master' THEN -5 WHEN 'Apprentice' THEN 5 ELSE 0 END)        AS RawTravelTimeMin,

    CASE
        WHEN delay_p < 0.70 THEN delay_p / 0.70 * 10
        WHEN delay_p < 0.90 THEN 10 + (delay_p - 0.70) / 0.20 * 20
        ELSE 30 + (delay_p - 0.90) / 0.10 * 60
    END + CASE WHEN RegionCode = 'EAST' THEN 15 ELSE 0 END AS DelayMin,

    StandardDurationMin * (1 + CASE SkillLevel WHEN 'Apprentice' THEN 0.20 WHEN 'Master' THEN -0.10 ELSE 0 END + dur_noise) AS RawDurationMin,

    CASE WHEN status_p < 0.90 THEN 'Completed' WHEN status_p < 0.95 THEN 'Cancelled' ELSE 'Rescheduled' END AS JobStatus,

    CASE SkillLevel
        WHEN 'Apprentice' THEN 0.78
        WHEN 'Journeyman' THEN 0.88
        ELSE 0.95
    END AS FTFProbability,
    ftf_p,

    SLAHours, BaseRevenue, BaseCost, cost_noise,
    CASE SkillLevel WHEN 'Master' THEN 1.15 WHEN 'Apprentice' THEN 0.90 ELSE 1.00 END AS SkillCostFactor
INTO #Jobs
FROM WithPicks Enriched;

DECLARE @JobsStaged INT = (SELECT COUNT(*) FROM #Jobs);
PRINT 'Job rows staged: ' + CAST(@JobsStaged AS VARCHAR(20));

-- Final insert with all business rules resolved
INSERT INTO dbo.Fact_ServiceJobs
    (JobID, DateKey, TechnicianKey, ClientKey, ServiceTypeKey, RegionKey,
     ScheduledStart, ActualStart, ActualEnd, TravelTimeMin, JobDurationMin,
     FirstTimeFix, CallbackFlag, SLAMet, JobCost, JobRevenue, JobStatus)
SELECT
    JobID, DateKey, TechnicianKey, ClientKey, ServiceTypeKey, RegionKey,
    ScheduledStart,
    CASE WHEN JobStatus = 'Completed' THEN DATEADD(MINUTE, CAST(ROUND(DelayMin,0) AS INT), ScheduledStart) ELSE NULL END AS ActualStart,
    CASE WHEN JobStatus = 'Completed' THEN DATEADD(MINUTE, CAST(ROUND(RawDurationMin,0) AS INT),
                DATEADD(MINUTE, CAST(ROUND(DelayMin,0) AS INT), ScheduledStart)) ELSE NULL END AS ActualEnd,
    CASE WHEN JobStatus = 'Completed' THEN CAST(ROUND(RawTravelTimeMin,0) AS INT) ELSE NULL END AS TravelTimeMin,
    CASE WHEN JobStatus = 'Completed' THEN CAST(ROUND(RawDurationMin,0) AS INT) ELSE NULL END AS JobDurationMin,
    CASE WHEN JobStatus = 'Completed' THEN CASE WHEN ftf_p < FTFProbability THEN 1 ELSE 0 END ELSE NULL END AS FirstTimeFix,
    CASE WHEN JobStatus = 'Completed' AND ftf_p >= FTFProbability THEN 1 ELSE 0 END AS CallbackFlag,
    CASE WHEN JobStatus = 'Completed' THEN
        CASE WHEN CAST(ROUND(DelayMin,0) AS INT) <= SLAHours * 60 THEN 1 ELSE 0 END
    ELSE NULL END AS SLAMet,
    CASE
        WHEN JobStatus = 'Completed' THEN ROUND(BaseCost * (1 + cost_noise) * SkillCostFactor, 2)
        WHEN JobStatus = 'Cancelled' THEN 25.00
        ELSE NULL
    END AS JobCost,
    CASE
        WHEN JobStatus = 'Completed' THEN ROUND(BaseRevenue * (1 + cost_noise), 2)
        ELSE 0.00
    END AS JobRevenue,
    JobStatus
FROM #Jobs;

PRINT 'Fact_ServiceJobs rows inserted: ' + CAST(@@ROWCOUNT AS VARCHAR(20));
GO

-- =============================================================================
-- 7. Deliberate, documented data-quality anomalies (~1.5% of Completed jobs)
--    Injected AFTER the clean load so the clean generation logic above stays
--    auditable. These exist solely so 03_data_quality_checks.sql has real,
--    known anomalies to detect -- this is a synthetic QA fixture, not noise.
-- =============================================================================

-- Anomaly A (~0.5%): ActualEnd before ActualStart (data entry swap)
;WITH Target AS (
    SELECT TOP (0.5) PERCENT JobKey, ActualStart, ActualEnd
    FROM dbo.Fact_ServiceJobs
    WHERE JobStatus = 'Completed'
    ORDER BY NEWID()
)
UPDATE f
SET f.ActualStart = t.ActualEnd,
    f.ActualEnd   = t.ActualStart
FROM dbo.Fact_ServiceJobs f
JOIN Target t ON t.JobKey = f.JobKey;

-- Anomaly B (~0.5%): Completed jobs missing actual start/end timestamps
;WITH Target AS (
    SELECT TOP (0.5) PERCENT JobKey
    FROM dbo.Fact_ServiceJobs
    WHERE JobStatus = 'Completed'
    ORDER BY NEWID()
)
UPDATE f
SET f.ActualStart = NULL,
    f.ActualEnd = NULL
FROM dbo.Fact_ServiceJobs f
JOIN Target t ON t.JobKey = f.JobKey;

-- Anomaly C (~0.5%): extreme outlier duration (minutes/hours data-entry error)
;WITH Target AS (
    SELECT TOP (0.5) PERCENT JobKey
    FROM dbo.Fact_ServiceJobs
    WHERE JobStatus = 'Completed'
    ORDER BY NEWID()
)
UPDATE f
SET f.JobDurationMin = 1200 + (ABS(CHECKSUM(NEWID())) % 400)
FROM dbo.Fact_ServiceJobs f
JOIN Target t ON t.JobKey = f.JobKey;

PRINT 'Deliberate data-quality anomalies injected for QA-check validation.';

DROP TABLE #Candidates;
DROP TABLE #Kept;
DROP TABLE #Jobs;
GO

PRINT '=== Synthetic data generation complete ===';
SELECT 'Dim_Date' AS TableName, COUNT(*) AS Rows FROM dbo.Dim_Date
UNION ALL SELECT 'Dim_Region', COUNT(*) FROM dbo.Dim_Region
UNION ALL SELECT 'Dim_Technician', COUNT(*) FROM dbo.Dim_Technician
UNION ALL SELECT 'Dim_Client', COUNT(*) FROM dbo.Dim_Client
UNION ALL SELECT 'Dim_ServiceType', COUNT(*) FROM dbo.Dim_ServiceType
UNION ALL SELECT 'Fact_ServiceJobs', COUNT(*) FROM dbo.Fact_ServiceJobs;
GO
