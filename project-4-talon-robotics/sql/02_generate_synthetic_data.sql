/*
================================================================================
Project 4 — Talon Robotics: Payload Deployment Delivery
Script:  02_generate_synthetic_data.sql
Purpose: Generate the synthetic programme dataset, deterministically.

DETERMINISM
    Every random draw comes from dbo.fn_Rand(<stable text key>), which hashes
    the key with SHA2_256 and scales it into [0,1). It is a pure function of
    the key, so dropping and regenerating reproduces the dataset byte for byte
    and every figure quoted in the case study stays checkable by a reader.
    RAND() and NEWID() cannot do this.

THE MECHANIC THIS DATA EXISTS TO EXPRESS

    A requirement's verification is CURRENT if its latest passing test ran
    against a build that nothing has invalidated since. It is STALE if, after
    that build:
        (a) any later build changed the same subsystem, or
        (b) the requirement itself was Modified (not merely Clarified).

    The generator therefore has to produce a plausible history in which tests
    are run early and code keeps moving underneath them -- which is what
    actually happens on a hardware programme as integration pressure rises.

DELIBERATE BEHAVIOUR PLANTED HERE
    P1  Late churn in Release Mechanism and Flight Control. Both are verified
        early and then changed repeatedly in the last quarter, so their
        verifications go stale in bulk. This is the headline finding.
    P2  Work items close steadily to ~94% while ship readiness lags far
        behind. Effort completes; evidence does not.
    P3  Policy breaches: some Safety requirements verified only at Integration
        level (below the HIL minimum), and some verified by their own owner in
        breach of the independent-tester rule.
    P4  RAID items past due, concentrated on the same churning subsystems, and
        a handful in the Critical exposure band left open.
    P5  Blocked runs on the HIL rig, because rig time is the scarce resource
        and contention is what actually stops verification closing out.
    P6  A cohort effect: Safety requirements look acceptable in aggregate while
        the MustShip Safety subset is the worst group in the programme.

PLANTED DATA DEFECTS -- for the data-quality layer to find BY BEHAVIOUR
    D1  Duplicate test runs (same case, same build, same day, same result).
    D2  Test runs dated before the build they ran against existed.
    D3  Work items closed before they were opened.
    D4  RAID items due before they were raised.
    D5  Test cases still being run against withdrawn requirements.

    None of these is findable by reading an ID prefix; each is found by the
    business signature, and 04_data_quality_checks.sql does exactly that.

DATA DISCLOSURE
    Talon Robotics is fictional. All data here is synthetic. No confidential
    data is used and no claim is made about any production system.
================================================================================
*/

USE TalonDelivery;
GO
SET NOCOUNT ON;
GO

-- =============================================================================
-- 0. Deterministic pseudo-randomness
-- =============================================================================
IF OBJECT_ID('dbo.fn_Rand', 'FN') IS NOT NULL DROP FUNCTION dbo.fn_Rand;
GO
CREATE FUNCTION dbo.fn_Rand (@Key VARCHAR(200))
RETURNS DECIMAL(18,15)
AS
BEGIN
    -- Top 6 bytes of the hash, scaled into [0,1). Six bytes is ample spread
    -- and keeps the arithmetic inside DECIMAL without overflow.
    DECLARE @h BINARY(32) = HASHBYTES('SHA2_256', @Key);
    DECLARE @n BIGINT =
          CAST(SUBSTRING(@h,1,1) AS BIGINT) * 1099511627776
        + CAST(SUBSTRING(@h,2,1) AS BIGINT) * 4294967296
        + CAST(SUBSTRING(@h,3,1) AS BIGINT) * 16777216
        + CAST(SUBSTRING(@h,4,1) AS BIGINT) * 65536
        + CAST(SUBSTRING(@h,5,1) AS BIGINT) * 256
        + CAST(SUBSTRING(@h,6,1) AS BIGINT);
    RETURN CAST(@n AS DECIMAL(38,15)) / CAST(281474976710656 AS DECIMAL(38,15));
END;
GO

-- =============================================================================
-- 1. Dim_Date  (2024-12-01 .. 2027-06-30)
-- =============================================================================
;WITH N AS (
    SELECT TOP (950) ROW_NUMBER() OVER (ORDER BY (SELECT NULL)) - 1 AS n
    FROM sys.all_objects a CROSS JOIN sys.all_objects b
),
D AS (SELECT [Date] = DATEADD(DAY, n, CAST('2024-12-01' AS DATE)) FROM N)
INSERT INTO dbo.Dim_Date (DateKey, [Date], [Year], [Quarter], [Month], MonthName,
                          YearMonth, YearQuarter, MonthEndDate, IsMonthEnd, IsWeekend, IsWorkday)
SELECT
    DateKey    = YEAR([Date])*10000 + MONTH([Date])*100 + DAY([Date]),
    [Date],
    [Year]     = YEAR([Date]),
    [Quarter]  = DATEPART(QUARTER, [Date]),
    [Month]    = MONTH([Date]),
    MonthName  = DATENAME(MONTH, [Date]),
    YearMonth  = CONVERT(CHAR(7), [Date], 126),
    YearQuarter= CAST(YEAR([Date]) AS CHAR(4)) + '-Q' + CAST(DATEPART(QUARTER,[Date]) AS CHAR(1)),
    MonthEndDate = EOMONTH([Date]),
    IsMonthEnd = CASE WHEN [Date] = EOMONTH([Date]) THEN 1 ELSE 0 END,
    IsWeekend  = CASE WHEN DATEPART(WEEKDAY, [Date]) IN (1,7) THEN 1 ELSE 0 END,
    IsWorkday  = CASE WHEN DATEPART(WEEKDAY, [Date]) IN (1,7) THEN 0 ELSE 1 END
FROM D
WHERE [Date] <= '2027-06-30';
GO

-- =============================================================================
-- 2. Dim_Subsystem
--
-- Eight subsystems. Release Mechanism and Flight Control are the ones that
-- churn late -- they are the payload's moving parts and its control loop, and
-- on a real programme those are exactly what keeps changing while everything
-- around them stabilises.
-- =============================================================================
INSERT INTO dbo.Dim_Subsystem (SubsystemCode, SubsystemName, Criticality, RequiresHILRig, LeadEngineerID) VALUES
 ('REL-MECH', 'Payload Release Mechanism',   'Safety',  1, 'P-004'),
 ('FLT-CTRL', 'Flight Control Interface',    'Safety',  1, 'P-006'),
 ('PWR-MGMT', 'Power Management',            'Mission', 1, 'P-008'),
 ('NAV-SENS', 'Navigation & Sensor Fusion',  'Mission', 1, 'P-010'),
 ('COMM-LNK', 'Command & Telemetry Link',    'Mission', 1, 'P-012'),
 ('GND-CTRL', 'Ground Control Software',     'Support', 0, 'P-014'),
 ('DIAG-LOG', 'Diagnostics & Logging',       'Support', 0, 'P-016'),
 ('STRUCT-I', 'Structural Interface',        'Safety',  1, 'P-018');
GO

-- =============================================================================
-- 3. Dim_Person
-- =============================================================================
;WITH N AS (SELECT TOP (24) ROW_NUMBER() OVER (ORDER BY (SELECT NULL)) AS n FROM sys.all_objects),
First AS (SELECT v, rn = ROW_NUMBER() OVER (ORDER BY v) FROM (VALUES
 ('Aditi'),('Bruno'),('Camille'),('Dmitri'),('Elena'),('Farid'),('Grace'),('Hiroshi'),
 ('Imani'),('Jonas'),('Katya'),('Liam'),('Mei'),('Nils'),('Olabisi'),('Priya'),
 ('Quentin'),('Rafael'),('Sana'),('Tomas'),('Ulrike'),('Viktor'),('Wen'),('Yusuf')) x(v)),
Last AS (SELECT v, rn = ROW_NUMBER() OVER (ORDER BY v) FROM (VALUES
 ('Arkwright'),('Baptiste'),('Castellan'),('Dalgleish'),('Eriksen'),('Fairweather'),
 ('Grigorescu'),('Halvorsen'),('Ibarra'),('Jindal'),('Kowalczyk'),('Lindqvist'),
 ('Moreau'),('Nakamura'),('Oyelaran'),('Petrovic'),('Quintana'),('Rasmussen'),
 ('Saltzman'),('Thackeray'),('Ustinov'),('Villalobos'),('Weatherby'),('Zielinski')) x(v))
INSERT INTO dbo.Dim_Person (PersonID, PersonName, [Role], Team, WeeklyCapacityHours)
SELECT
    PersonID = 'P-' + RIGHT('000' + CAST(n.n AS VARCHAR(3)), 3),
    PersonName = f.v + ' ' + l.v,
    [Role] = CASE
        WHEN n.n <= 4  THEN 'Systems Engineer'
        WHEN n.n <= 10 THEN 'Software Engineer'
        WHEN n.n <= 16 THEN 'Test Engineer'
        WHEN n.n <= 19 THEN 'Quality Engineer'
        WHEN n.n <= 21 THEN 'Hardware Engineer'
        ELSE 'Project Manager' END,
    Team = CASE WHEN n.n % 4 = 0 THEN 'Payload'
                WHEN n.n % 4 = 1 THEN 'Avionics'
                WHEN n.n % 4 = 2 THEN 'Ground Segment'
                ELSE 'Integration & Test' END,
    -- Test engineers carry the verification load; everyone else contributes
    -- some. The queue is bounded by the sum of this column.
    WeeklyCapacityHours = CASE
        WHEN n.n BETWEEN 11 AND 16 THEN 28.0
        WHEN n.n BETWEEN 17 AND 19 THEN 16.0
        ELSE 8.0 END
FROM N n
JOIN First f ON f.rn = n.n
JOIN Last  l ON l.rn = ((n.n * 7) % 24) + 1;
GO

-- =============================================================================
-- 4. Dim_Build
--
-- 88 builds, roughly weekly from 2025-01-06 to 2026-09-28. Ordering is by
-- BuildNumber, never by date: two builds can share a date and a date can be
-- corrected, but the sequence is what "superseded by" means.
-- =============================================================================
;WITH N AS (
    SELECT TOP (88) ROW_NUMBER() OVER (ORDER BY (SELECT NULL)) AS n
    FROM sys.all_objects a CROSS JOIN sys.all_objects b
)
INSERT INTO dbo.Dim_Build (BuildID, BuildNumber, BuildDateKey, BuildLabel, IsReleaseCandidate)
SELECT
    BuildID = 'B-' + RIGHT('000' + CAST(n AS VARCHAR(3)), 3),
    BuildNumber = n,
    BuildDateKey = YEAR(d)*10000 + MONTH(d)*100 + DAY(d),
    BuildLabel = CASE
        WHEN n = 88 THEN 'RC1 - payload delivery candidate'
        WHEN n % 13 = 0 THEN 'Integration drop ' + CAST(n/13 AS VARCHAR(3))
        ELSE 'Weekly build ' + CAST(n AS VARCHAR(3)) END,
    IsReleaseCandidate = CASE WHEN n = 88 THEN 1 ELSE 0 END
FROM (SELECT n, d = DATEADD(DAY, (n-1)*7, CAST('2025-01-06' AS DATE)) FROM N) x;
GO

-- =============================================================================
-- 5. Dim_Requirement  (420 requirements)
-- =============================================================================
DECLARE @ReqTarget INT = 540;

;WITH N AS (
    SELECT TOP (540) ROW_NUMBER() OVER (ORDER BY (SELECT NULL)) AS n
    FROM sys.all_objects a CROSS JOIN sys.all_objects b
),
R AS (
    SELECT n,
           rSub  = dbo.fn_Rand(CONCAT('req|sub|',  n)),
           rType = dbo.fn_Rand(CONCAT('req|type|', n)),
           rPri  = dbo.fn_Rand(CONCAT('req|pri|',  n)),
           rOwn  = dbo.fn_Rand(CONCAT('req|own|',  n)),
           rBase = dbo.fn_Rand(CONCAT('req|base|', n)),
           rWdn  = dbo.fn_Rand(CONCAT('req|wdn|',  n))
    FROM N
),
Assigned AS (
    SELECT r.n,
           -- Subsystem weights: the safety-critical moving parts carry the most
           -- requirements, which is why their churn hurts so much later.
           SubCode = CASE
               WHEN r.rSub < 0.18 THEN 'REL-MECH'
               WHEN r.rSub < 0.34 THEN 'FLT-CTRL'
               WHEN r.rSub < 0.48 THEN 'PWR-MGMT'
               WHEN r.rSub < 0.62 THEN 'NAV-SENS'
               WHEN r.rSub < 0.75 THEN 'COMM-LNK'
               WHEN r.rSub < 0.86 THEN 'GND-CTRL'
               WHEN r.rSub < 0.94 THEN 'DIAG-LOG'
               ELSE                    'STRUCT-I' END,
           r.rType, r.rPri, r.rOwn, r.rBase, r.rWdn
    FROM R r
),
Typed AS (
    SELECT a.*,
           s.SubsystemKey, s.Criticality,
           ReqType = CASE
               -- Safety-critical subsystems attract Safety and Regulatory
               -- requirements; support subsystems almost never do.
               WHEN s.Criticality = 'Safety' AND a.rType < 0.34 THEN 'Safety'
               WHEN s.Criticality = 'Safety' AND a.rType < 0.44 THEN 'Regulatory'
               WHEN s.Criticality = 'Safety' AND a.rType < 0.62 THEN 'Interface'
               WHEN s.Criticality = 'Safety' AND a.rType < 0.80 THEN 'Performance'
               WHEN s.Criticality = 'Safety'                    THEN 'Functional'
               WHEN s.Criticality = 'Mission' AND a.rType < 0.10 THEN 'Safety'
               WHEN s.Criticality = 'Mission' AND a.rType < 0.16 THEN 'Regulatory'
               WHEN s.Criticality = 'Mission' AND a.rType < 0.40 THEN 'Interface'
               WHEN s.Criticality = 'Mission' AND a.rType < 0.66 THEN 'Performance'
               WHEN s.Criticality = 'Mission'                    THEN 'Functional'
               WHEN a.rType < 0.22 THEN 'Interface'
               WHEN a.rType < 0.40 THEN 'Performance'
               ELSE 'Functional' END
    FROM Assigned a
    JOIN dbo.Dim_Subsystem s ON s.SubsystemCode = a.SubCode
)
INSERT INTO dbo.Dim_Requirement (RequirementID, Title, ReqType, Priority, SubsystemKey,
                                 OwnerKey, BaselinedDateKey, ReqStatus)
SELECT
    RequirementID = 'REQ-' + RIGHT('0000' + CAST(t.n AS VARCHAR(4)), 4),
    Title = CONCAT(
        CASE t.ReqType
            WHEN 'Safety'      THEN 'The payload shall inhibit release when '
            WHEN 'Regulatory'  THEN 'The system shall record, for certification, '
            WHEN 'Performance' THEN 'The subsystem shall complete, within budget, '
            WHEN 'Interface'   THEN 'The interface shall exchange, without loss, '
            ELSE                    'The system shall provide ' END,
        LOWER(s2.SubsystemName), ' behaviour case ', CAST(t.n AS VARCHAR(4))),
    t.ReqType,
    Priority = CASE
        -- Safety and Regulatory are overwhelmingly MustShip; that is what makes
        -- the MustShip cohort the one that matters.
        WHEN t.ReqType IN ('Safety','Regulatory') AND t.rPri < 0.88 THEN 'MustShip'
        WHEN t.ReqType IN ('Safety','Regulatory')                    THEN 'ShouldShip'
        WHEN t.rPri < 0.42 THEN 'MustShip'
        WHEN t.rPri < 0.80 THEN 'ShouldShip'
        ELSE 'Nice' END,
    t.SubsystemKey,
    OwnerKey = (SELECT PersonKey FROM dbo.Dim_Person
                WHERE PersonID = 'P-' + RIGHT('000' + CAST(1 + CAST(t.rOwn * 21 AS INT) AS VARCHAR(3)), 3)),
    BaselinedDateKey = (SELECT DateKey FROM dbo.Dim_Date
                        WHERE [Date] = DATEADD(DAY, CAST(t.rBase * 150 AS INT), CAST('2025-01-06' AS DATE))),
    -- A small number of requirements are withdrawn during the programme. They
    -- are not verification failures, and the KPI layer must not count them.
    ReqStatus = CASE WHEN t.rWdn < 0.035 THEN 'Withdrawn' ELSE 'Baselined' END
FROM Typed t
JOIN dbo.Dim_Subsystem s2 ON s2.SubsystemKey = t.SubsystemKey;
GO

-- =============================================================================
-- 6. Dim_TestCase
--
-- One to four cases per requirement. The TEST LEVEL is drawn against the
-- policy minimum, but deliberately falls short for a slice of Safety
-- requirements -- planted pattern P3.
-- =============================================================================
;WITH Req AS (
    SELECT r.RequirementKey, r.RequirementID, r.ReqType, r.SubsystemKey,
           s.RequiresHILRig, s.Criticality,
           p.MinTestLevelRank,
           nCases = 1 + CAST(dbo.fn_Rand(CONCAT('tc|n|', r.RequirementKey)) * 3.9 AS INT)
    FROM dbo.Dim_Requirement r
    JOIN dbo.Dim_Subsystem s ON s.SubsystemKey = r.SubsystemKey
    JOIN dbo.Ref_VerificationPolicy p ON p.ReqType = r.ReqType
),
Tally AS (SELECT TOP (4) i = ROW_NUMBER() OVER (ORDER BY (SELECT NULL)) FROM sys.all_objects),
Expanded AS (
    SELECT q.RequirementKey, q.RequirementID, q.ReqType, q.MinTestLevelRank,
           q.RequiresHILRig, t.i,
           seq = ROW_NUMBER() OVER (ORDER BY q.RequirementKey, t.i),
           rLvl = dbo.fn_Rand(CONCAT('tc|lvl|', q.RequirementKey, '|', t.i)),
           rAut = dbo.fn_Rand(CONCAT('tc|aut|', q.RequirementKey, '|', t.i)),
           rShort = dbo.fn_Rand(CONCAT('tc|short|', q.RequirementKey))
    FROM Req q JOIN Tally t ON t.i <= q.nCases
)
INSERT INTO dbo.Dim_TestCase (TestCaseID, RequirementKey, TestLevel, AutomationLevel, ExpectedDurationMin)
SELECT
    TestCaseID = 'TC-' + RIGHT('0000' + CAST(e.seq AS VARCHAR(4)), 4),
    e.RequirementKey,
    TestLevel = CASE
        -- PLANTED P3: roughly one in nine requirements whose policy demands HIL
        -- is covered only at Integration level. The requirement still shows a
        -- passing test; the test is not sufficient evidence for its type.
        WHEN e.MinTestLevelRank >= 3 AND e.rShort < 0.11 THEN 'Integration'
        WHEN e.MinTestLevelRank = 4 THEN CASE WHEN e.rLvl < 0.55 THEN 'Field' ELSE 'HIL' END
        WHEN e.MinTestLevelRank = 3 THEN CASE WHEN e.rLvl < 0.62 THEN 'HIL'
                                              WHEN e.rLvl < 0.88 THEN 'Integration' ELSE 'Unit' END
        ELSE CASE WHEN e.rLvl < 0.50 THEN 'Integration'
                  WHEN e.rLvl < 0.85 THEN 'Unit' ELSE 'HIL' END END,
    AutomationLevel = CASE WHEN e.rAut < 0.58 THEN 'Automated' ELSE 'Manual' END,
    ExpectedDurationMin = 15 + CAST(dbo.fn_Rand(CONCAT('tc|dur|', e.seq)) * 210 AS INT)
FROM Expanded e;
GO

-- =============================================================================
-- 7. Fact_BuildSubsystemChange
--
-- PLANTED P1. Early builds spread changes broadly as everything is written.
-- From roughly build 62 onward, change concentrates hard in REL-MECH and
-- FLT-CTRL -- late integration churn in the payload's moving parts. Every one
-- of those late changes invalidates any earlier verification of that subsystem.
-- =============================================================================
INSERT INTO dbo.Fact_BuildSubsystemChange (BuildKey, SubsystemKey, LinesChanged)
SELECT b.BuildKey, s.SubsystemKey,
       LinesChanged = 20 + CAST(dbo.fn_Rand(CONCAT('bsc|lines|', b.BuildNumber, '|', s.SubsystemCode)) * 900 AS INT)
FROM dbo.Dim_Build b
CROSS JOIN dbo.Dim_Subsystem s
CROSS APPLY (SELECT r = dbo.fn_Rand(CONCAT('bsc|hit|', b.BuildNumber, '|', s.SubsystemCode))) x
WHERE
    -- early programme: broad, steady change everywhere
    (b.BuildNumber <= 61 AND x.r < 0.30)
    -- late programme: the two churning subsystems move in most builds, the
    -- rest have largely stabilised
 OR (b.BuildNumber >  61 AND s.SubsystemCode IN ('REL-MECH','FLT-CTRL') AND x.r < 0.72)
 OR (b.BuildNumber >  61 AND s.SubsystemCode NOT IN ('REL-MECH','FLT-CTRL') AND x.r < 0.09);
GO

-- =============================================================================
-- 8. Fact_TestRun
--
-- Each test case is attempted on one to three builds. The build it was LAST
-- run against is what decides staleness, so the distribution of that build
-- across the programme is the heart of the dataset.
-- =============================================================================
IF OBJECT_ID('tempdb..#RunPlan') IS NOT NULL DROP TABLE #RunPlan;

;WITH TC AS (
    SELECT tc.TestCaseKey, tc.TestCaseID, tc.TestLevel, tc.RequirementKey, tc.ExpectedDurationMin,
           r.SubsystemKey, r.OwnerKey, r.ReqType, r.ReqStatus,
           s.RequiresHILRig, s.SubsystemCode,
           rWhen  = dbo.fn_Rand(CONCAT('run|when|',  tc.TestCaseKey)),
           rCount = dbo.fn_Rand(CONCAT('run|count|', tc.TestCaseKey)),
           rReg   = dbo.fn_Rand(CONCAT('run|reg|',   tc.TestCaseKey))
    FROM dbo.Dim_TestCase tc
    JOIN dbo.Dim_Requirement r ON r.RequirementKey = tc.RequirementKey
    JOIN dbo.Dim_Subsystem s   ON s.SubsystemKey   = r.SubsystemKey
),
Planned AS (
    SELECT t.*,
           nRuns = 1 + CAST(t.rCount * 3.4 AS INT),
           -- WHETHER THIS CASE MADE THE RELEASE-CANDIDATE REGRESSION.
           --
           -- This is the causal spine of the dataset, and the reason the
           -- programme's problem is one problem rather than two.
           --
           -- Approaching a release candidate a team re-runs its regression
           -- suite, and anything re-run on the RC is current by definition.
           -- What does NOT get re-run is whatever needs the scarce resource:
           -- the HIL rig, and above it the airframe. Automated software tests
           -- are nearly free and go back through every time; rig time and
           -- flight time are rationed.
           --
           -- So rig contention is not a separate finding sitting beside stale
           -- evidence -- it is the CAUSE of it, and it lands hardest on
           -- exactly the Safety and Regulatory requirements whose policy
           -- demands hardware evidence in the first place.
           --
           -- Teams on the churning subsystems re-run less again, because they
           -- spend the endgame changing code rather than testing it.
           InRCRegression = CASE WHEN t.rReg <
                   CASE t.TestLevel
                        WHEN 'Unit'        THEN 0.90
                        WHEN 'Integration' THEN 0.88
                        WHEN 'HIL'         THEN 0.74
                        ELSE                    0.54   -- Field: airframe-limited
                   END
                 * CASE WHEN t.SubsystemCode IN ('REL-MECH','FLT-CTRL') THEN 0.42 ELSE 1.00 END
                THEN 1 ELSE 0 END
    FROM TC t
),
Placed AS (
    SELECT p.*,
           LastBuildNo = CASE WHEN p.InRCRegression = 1
                                THEN 88
                -- otherwise the evidence is whatever was last captured, and
                -- the programme has moved on underneath it
                ELSE 28 + CAST(p.rWhen * 44 AS INT) END
    FROM Planned p
),
Tally AS (SELECT TOP (4) i = ROW_NUMBER() OVER (ORDER BY (SELECT NULL)) FROM sys.all_objects)
SELECT p.TestCaseKey, p.TestCaseID, p.TestLevel, p.RequirementKey, p.ReqType, p.ReqStatus,
       p.OwnerKey, p.SubsystemCode, p.RequiresHILRig, p.ExpectedDurationMin,
       RunIndex = t.i,
       -- earlier attempts sit before the final one, six builds apart
       BuildNo = p.LastBuildNo - (p.nRuns - t.i) * 6,
       IsFinalRun = CASE WHEN t.i = p.nRuns THEN 1 ELSE 0 END
INTO #RunPlan
FROM Placed p
JOIN Tally t ON t.i <= p.nRuns
WHERE p.LastBuildNo - (p.nRuns - t.i) * 6 >= 1;

INSERT INTO dbo.Fact_TestRun (TestCaseKey, BuildKey, RunDateKey, RunByKey, Result, DurationMin)
SELECT
    rp.TestCaseKey,
    b.BuildKey,
    -- Runs land one to five days after the build, on a workday.
    RunDateKey = (SELECT TOP 1 d2.DateKey FROM dbo.Dim_Date d2
                  WHERE d2.[Date] >= DATEADD(DAY, 1 + CAST(dbo.fn_Rand(CONCAT('run|lag|', rp.TestCaseKey, '|', rp.RunIndex)) * 5 AS INT), bd.[Date])
                    AND d2.IsWorkday = 1 ORDER BY d2.[Date]),
    RunByKey = CASE
        -- PLANTED P3 (second half): on a slice of Safety requirements the
        -- owner verifies their own work, breaching the independent-tester
        -- rule. The run passes; the evidence does not count.
        WHEN rp.ReqType IN ('Safety','Regulatory')
             AND dbo.fn_Rand(CONCAT('run|self|', rp.TestCaseKey, '|', rp.RunIndex)) < 0.14
        THEN rp.OwnerKey
        ELSE (SELECT PersonKey FROM dbo.Dim_Person
              WHERE PersonID = 'P-' + RIGHT('000' + CAST(11 + CAST(dbo.fn_Rand(CONCAT('run|by|', rp.TestCaseKey, '|', rp.RunIndex)) * 8 AS INT) AS VARCHAR(3)), 3))
        END,
    Result = CASE
        -- PLANTED P5: HIL rig contention shows up as Blocked runs, not
        -- failures. A blocked test is not evidence of anything, and a
        -- readiness metric that treats "not run" as "not failing" is the
        -- reason rig contention stays invisible until the ship date.
        WHEN rp.RequiresHILRig = 1 AND rp.TestLevel IN ('HIL','Field')
             AND dbo.fn_Rand(CONCAT('run|blk|', rp.TestCaseKey, '|', rp.RunIndex)) < 0.10 THEN 'Blocked'
        WHEN rp.IsFinalRun = 0
             AND dbo.fn_Rand(CONCAT('run|earlyfail|', rp.TestCaseKey, '|', rp.RunIndex)) < 0.38 THEN 'Fail'
        WHEN rp.IsFinalRun = 1
             AND dbo.fn_Rand(CONCAT('run|finfail|', rp.TestCaseKey, '|', rp.RunIndex)) < 0.07 THEN 'Fail'
        ELSE 'Pass' END,
    DurationMin = CAST(rp.ExpectedDurationMin
                  * (0.7 + dbo.fn_Rand(CONCAT('run|durf|', rp.TestCaseKey, '|', rp.RunIndex)) * 0.8) AS INT)
FROM #RunPlan rp
JOIN dbo.Dim_Build b ON b.BuildNumber = rp.BuildNo
JOIN dbo.Dim_Date bd ON bd.DateKey = b.BuildDateKey
-- Withdrawn requirements stop being tested once withdrawn; D5 below puts a
-- handful back deliberately so the data-quality layer has something to find.
WHERE rp.ReqStatus = 'Baselined';
GO

-- =============================================================================
-- 9. Fact_RequirementChange
--
-- PLANTED P6. Some requirements are Modified after their verification ran,
-- which invalidates it just as surely as a code change. Clarifications are
-- also recorded and deliberately do NOT invalidate anything -- the distinction
-- is the point, and a model that lumped them together would overstate the
-- problem and lose credibility for the part that is real.
-- =============================================================================
;WITH R AS (
    SELECT r.RequirementKey, r.RequirementID, r.OwnerKey,
           rN   = dbo.fn_Rand(CONCAT('rc|n|',    r.RequirementKey)),
           rWhen= dbo.fn_Rand(CONCAT('rc|when|', r.RequirementKey)),
           rType= dbo.fn_Rand(CONCAT('rc|type|', r.RequirementKey))
    FROM dbo.Dim_Requirement r
    WHERE r.ReqStatus = 'Baselined'
),
Changed AS (
    SELECT R.*, BuildNo = 40 + CAST(R.rWhen * 46 AS INT)
    FROM R WHERE R.rN < 0.22
)
INSERT INTO dbo.Fact_RequirementChange (RequirementKey, ChangeDateKey, BuildKey, ChangeType, ChangedByKey, ChangeNote)
SELECT
    c.RequirementKey,
    ChangeDateKey = b.BuildDateKey,
    b.BuildKey,
    ChangeType = CASE WHEN c.rType < 0.55 THEN 'Modified' ELSE 'Clarified' END,
    ChangedByKey = c.OwnerKey,
    ChangeNote = CASE WHEN c.rType < 0.55
        THEN 'Acceptance threshold revised after integration findings; prior verification no longer applies.'
        ELSE 'Wording clarified for certification review; behaviour unchanged, verification still valid.' END
FROM Changed c
JOIN dbo.Dim_Build b ON b.BuildNumber = c.BuildNo;
GO

-- =============================================================================
-- 10. Fact_WorkItem
--
-- PLANTED P2. Work items close steadily and reach roughly 94% complete, which
-- is the number the programme reports. It counts effort, not evidence.
-- =============================================================================
IF OBJECT_ID('tempdb..#WI') IS NOT NULL DROP TABLE #WI;

;WITH Req AS (
    SELECT r.RequirementKey, r.SubsystemKey, r.Priority,
           nItems = 1 + CAST(dbo.fn_Rand(CONCAT('wi|n|', r.RequirementKey)) * 3.2 AS INT)
    FROM dbo.Dim_Requirement r WHERE r.ReqStatus = 'Baselined'
),
Tally AS (SELECT TOP (4) i = ROW_NUMBER() OVER (ORDER BY (SELECT NULL)) FROM sys.all_objects)
SELECT q.RequirementKey, q.SubsystemKey, q.Priority, t.i,
       seq   = ROW_NUMBER() OVER (ORDER BY q.RequirementKey, t.i),
       rOpen = dbo.fn_Rand(CONCAT('wi|open|',  q.RequirementKey, '|', t.i)),
       rDur  = dbo.fn_Rand(CONCAT('wi|dur|',   q.RequirementKey, '|', t.i)),
       rDone = dbo.fn_Rand(CONCAT('wi|done|',  q.RequirementKey, '|', t.i)),
       rEst  = dbo.fn_Rand(CONCAT('wi|est|',   q.RequirementKey, '|', t.i)),
       rWho  = dbo.fn_Rand(CONCAT('wi|who|',   q.RequirementKey, '|', t.i))
INTO #WI
FROM Req q JOIN Tally t ON t.i <= q.nItems;

INSERT INTO dbo.Fact_WorkItem (WorkItemID, RequirementKey, AssignedToKey, OpenedDateKey,
                               ClosedDateKey, EstimateHours, ActualHours, WorkItemStatus)
SELECT
    WorkItemID = 'WI-' + RIGHT('0000' + CAST(w.seq AS VARCHAR(4)), 4),
    w.RequirementKey,
    AssignedToKey = (SELECT PersonKey FROM dbo.Dim_Person
                     WHERE PersonID = 'P-' + RIGHT('000' + CAST(1 + CAST(w.rWho * 21 AS INT) AS VARCHAR(3)), 3)),
    OpenedDateKey = od.DateKey,
    ClosedDateKey = CASE WHEN st.Status = 'Closed' THEN cd.DateKey END,
    EstimateHours = CAST(2 + w.rEst * 34 AS DECIMAL(6,1)),
    ActualHours   = CASE WHEN st.Status = 'Closed'
                         THEN CAST((2 + w.rEst * 34) * (0.75 + w.rDur * 0.7) AS DECIMAL(6,1)) END,
    WorkItemStatus = st.Status
FROM #WI w
CROSS APPLY (SELECT Status = CASE WHEN w.rDone < 0.940 THEN 'Closed'
                                  WHEN w.rDone < 0.978 THEN 'InProgress'
                                  ELSE 'Open' END) st
CROSS APPLY (SELECT OpenDate = DATEADD(DAY, CAST(w.rOpen * 540 AS INT), CAST('2025-01-06' AS DATE))) o
CROSS APPLY (SELECT CloseDate = DATEADD(DAY, 3 + CAST(w.rDur * 55 AS INT), o.OpenDate)) c
JOIN dbo.Dim_Date od ON od.[Date] = o.OpenDate
LEFT JOIN dbo.Dim_Date cd ON cd.[Date] = CASE WHEN c.CloseDate <= '2026-09-30' THEN c.CloseDate ELSE '2026-09-30' END;
GO

-- =============================================================================
-- 11. Fact_RAID
--
-- PLANTED P4. Exposure concentrates on the churning subsystems, several
-- Critical items stay open, and a meaningful share are past their due date.
-- =============================================================================
;WITH N AS (
    SELECT TOP (96) ROW_NUMBER() OVER (ORDER BY (SELECT NULL)) AS n
    FROM sys.all_objects a CROSS JOIN sys.all_objects b
),
R AS (
    SELECT n,
           rSub   = dbo.fn_Rand(CONCAT('raid|sub|',   n)),
           rType  = dbo.fn_Rand(CONCAT('raid|type|',  n)),
           rProb  = dbo.fn_Rand(CONCAT('raid|prob|',  n)),
           rImp   = dbo.fn_Rand(CONCAT('raid|imp|',   n)),
           rStat  = dbo.fn_Rand(CONCAT('raid|stat|',  n)),
           rRaise = dbo.fn_Rand(CONCAT('raid|raise|', n)),
           rDue   = dbo.fn_Rand(CONCAT('raid|due|',   n)),
           rOwn   = dbo.fn_Rand(CONCAT('raid|own|',   n))
    FROM N
),
Built AS (
    SELECT r.*,
           SubCode = CASE WHEN r.rSub < 0.30 THEN 'REL-MECH'
                          WHEN r.rSub < 0.54 THEN 'FLT-CTRL'
                          WHEN r.rSub < 0.66 THEN 'PWR-MGMT'
                          WHEN r.rSub < 0.77 THEN 'NAV-SENS'
                          WHEN r.rSub < 0.86 THEN 'COMM-LNK'
                          WHEN r.rSub < 0.93 THEN 'GND-CTRL'
                          WHEN r.rSub < 0.97 THEN 'DIAG-LOG'
                          ELSE                    'STRUCT-I' END,
           RAIDType = CASE WHEN r.rType < 0.44 THEN 'Risk'
                           WHEN r.rType < 0.68 THEN 'Issue'
                           WHEN r.rType < 0.86 THEN 'Dependency'
                           ELSE 'Assumption' END,
           Probability = 1 + CAST(r.rProb * 4.999 AS INT),
           Impact      = 1 + CAST(r.rImp  * 4.999 AS INT),
           RaisedDate  = DATEADD(DAY, CAST(r.rRaise * 560 AS INT), CAST('2025-01-06' AS DATE))
    FROM R r
)
INSERT INTO dbo.Fact_RAID (RAIDID, RAIDType, Title, SubsystemKey, OwnerKey, RaisedDateKey,
                           DueDateKey, ClosedDateKey, Probability, Impact, RAIDStatus, MitigationNote)
SELECT
    RAIDID = 'RAID-' + RIGHT('000' + CAST(b.n AS VARCHAR(3)), 3),
    b.RAIDType,
    Title = CONCAT(b.RAIDType, ' on ', s.SubsystemName, ': ',
        CASE b.RAIDType
            WHEN 'Risk'       THEN 'late design change may invalidate completed qualification evidence'
            WHEN 'Issue'      THEN 'rig availability is below the rate needed to clear outstanding verification'
            WHEN 'Dependency' THEN 'awaiting supplier confirmation before verification can close'
            ELSE                   'assumed unchanged from the previous airframe, not yet confirmed' END),
    s.SubsystemKey,
    OwnerKey = (SELECT PersonKey FROM dbo.Dim_Person
                WHERE PersonID = 'P-' + RIGHT('000' + CAST(1 + CAST(b.rOwn * 21 AS INT) AS VARCHAR(3)), 3)),
    RaisedDateKey = rd.DateKey,
    DueDateKey    = dd.DateKey,
    ClosedDateKey = CASE WHEN st.S = 'Closed' THEN cd.DateKey END,
    b.Probability, b.Impact,
    RAIDStatus = st.S,
    MitigationNote = CASE st.S
        WHEN 'Closed'     THEN 'Closed: mitigation completed and evidence attached to the verification record.'
        WHEN 'Mitigating' THEN 'Mitigation in progress; owner reporting weekly to the programme board.'
        ELSE                   'Open: no mitigation started. Requires a decision at the next programme board.' END
FROM Built b
JOIN dbo.Dim_Subsystem s ON s.SubsystemCode = b.SubCode
CROSS APPLY (SELECT S = CASE
        -- Items on the churning subsystems stay open far more often: the
        -- change that caused them has not stopped happening.
        WHEN b.SubCode IN ('REL-MECH','FLT-CTRL') AND b.rStat < 0.46 THEN 'Open'
        WHEN b.SubCode IN ('REL-MECH','FLT-CTRL') AND b.rStat < 0.68 THEN 'Mitigating'
        WHEN b.rStat < 0.22 THEN 'Open'
        WHEN b.rStat < 0.38 THEN 'Mitigating'
        ELSE 'Closed' END) st
CROSS APPLY (SELECT DueDate = DATEADD(DAY, 45 + CAST(b.rDue * 330 AS INT), b.RaisedDate)) du
CROSS APPLY (SELECT CloseDate = DATEADD(DAY, 10 + CAST(b.rDue * 90 AS INT), b.RaisedDate)) cl
JOIN dbo.Dim_Date rd ON rd.[Date] = b.RaisedDate
JOIN dbo.Dim_Date dd ON dd.[Date] = du.DueDate
LEFT JOIN dbo.Dim_Date cd ON cd.[Date] = cl.CloseDate;
GO

-- =============================================================================
-- 12. PLANTED DATA DEFECTS
--
-- Each is created by BEHAVIOUR, not by a marker in an ID, so the data-quality
-- layer has to find it the way it would have to in real data.
-- =============================================================================

-- D1: duplicate test runs -- the same case, build, date and result recorded twice.
INSERT INTO dbo.Fact_TestRun (TestCaseKey, BuildKey, RunDateKey, RunByKey, Result, DurationMin)
SELECT TOP (22) r.TestCaseKey, r.BuildKey, r.RunDateKey, r.RunByKey, r.Result, r.DurationMin
FROM dbo.Fact_TestRun r
WHERE r.Result = 'Pass'
ORDER BY dbo.fn_Rand(CONCAT('defect|dup|', r.TestRunKey));
GO

-- D2: test runs dated BEFORE the build they ran against existed.
UPDATE TOP (14) r
SET r.RunDateKey = (SELECT d2.DateKey FROM dbo.Dim_Date d2
                    WHERE d2.[Date] = DATEADD(DAY, -9, bd.[Date]))
FROM dbo.Fact_TestRun r
JOIN dbo.Dim_Build b  ON b.BuildKey = r.BuildKey
JOIN dbo.Dim_Date bd  ON bd.DateKey = b.BuildDateKey
WHERE b.BuildNumber > 12
  AND dbo.fn_Rand(CONCAT('defect|early|', r.TestRunKey)) < 0.05;
GO

-- D3: work items closed before they were opened.
UPDATE TOP (9) w
SET w.ClosedDateKey = (SELECT d2.DateKey FROM dbo.Dim_Date d2
                       WHERE d2.[Date] = DATEADD(DAY, -6, od.[Date]))
FROM dbo.Fact_WorkItem w
JOIN dbo.Dim_Date od ON od.DateKey = w.OpenedDateKey
WHERE w.WorkItemStatus = 'Closed'
  AND od.[Date] > '2025-03-01'
  AND dbo.fn_Rand(CONCAT('defect|wiclose|', w.WorkItemKey)) < 0.05;
GO

-- D4: RAID items due before they were raised.
UPDATE TOP (7) x
SET x.DueDateKey = (SELECT d2.DateKey FROM dbo.Dim_Date d2
                    WHERE d2.[Date] = DATEADD(DAY, -11, rd.[Date]))
FROM dbo.Fact_RAID x
JOIN dbo.Dim_Date rd ON rd.DateKey = x.RaisedDateKey
WHERE rd.[Date] > '2025-04-01'
  AND dbo.fn_Rand(CONCAT('defect|raiddue|', x.RAIDKey)) < 0.12;
GO

-- D5: test runs against WITHDRAWN requirements -- effort spent verifying
-- something the programme has already decided not to ship.
INSERT INTO dbo.Fact_TestRun (TestCaseKey, BuildKey, RunDateKey, RunByKey, Result, DurationMin)
SELECT TOP (16)
       tc.TestCaseKey,
       b.BuildKey,
       b.BuildDateKey,
       (SELECT TOP 1 PersonKey FROM dbo.Dim_Person WHERE [Role] = 'Test Engineer' ORDER BY PersonKey),
       'Pass',
       tc.ExpectedDurationMin
FROM dbo.Dim_TestCase tc
JOIN dbo.Dim_Requirement r ON r.RequirementKey = tc.RequirementKey
CROSS JOIN (SELECT TOP 1 BuildKey, BuildDateKey FROM dbo.Dim_Build ORDER BY BuildNumber DESC) b
WHERE r.ReqStatus = 'Withdrawn'
ORDER BY dbo.fn_Rand(CONCAT('defect|withdrawn|', tc.TestCaseKey));
GO

-- =============================================================================
-- 13. Reproducibility fingerprints
-- =============================================================================
PRINT '';
PRINT '=== Talon Robotics: generation summary ===';
SELECT
    Subsystems   = (SELECT COUNT(*) FROM dbo.Dim_Subsystem),
    People       = (SELECT COUNT(*) FROM dbo.Dim_Person),
    Builds       = (SELECT COUNT(*) FROM dbo.Dim_Build),
    Requirements = (SELECT COUNT(*) FROM dbo.Dim_Requirement),
    TestCases    = (SELECT COUNT(*) FROM dbo.Dim_TestCase),
    TestRuns     = (SELECT COUNT(*) FROM dbo.Fact_TestRun),
    BuildChanges = (SELECT COUNT(*) FROM dbo.Fact_BuildSubsystemChange),
    ReqChanges   = (SELECT COUNT(*) FROM dbo.Fact_RequirementChange),
    WorkItems    = (SELECT COUNT(*) FROM dbo.Fact_WorkItem),
    RAIDItems    = (SELECT COUNT(*) FROM dbo.Fact_RAID);

PRINT '';
PRINT 'Fingerprints (asserted by the UAT suite -- a rebuild must reproduce these):';
SELECT
    TestRunFingerprint  = CHECKSUM_AGG(CAST(TestCaseKey * 31 + BuildKey * 7
                                        + CASE Result WHEN 'Pass' THEN 1 WHEN 'Fail' THEN 2 ELSE 3 END AS INT))
FROM dbo.Fact_TestRun;
SELECT
    WorkItemFingerprint = CHECKSUM_AGG(CAST(RequirementKey * 17 + ISNULL(ClosedDateKey,0) % 9973 AS INT))
FROM dbo.Fact_WorkItem;
SELECT
    RAIDFingerprint     = CHECKSUM_AGG(CAST(SubsystemKey * 13 + Probability * 5 + Impact AS INT))
FROM dbo.Fact_RAID;
GO

DECLARE @runs INT = (SELECT COUNT(*) FROM dbo.Fact_TestRun);
IF @runs < 3000 THROW 52010, 'Generator produced too few test runs -- the model will not exercise staleness.', 1;
PRINT '';
PRINT 'Synthetic programme data generated.';
GO
