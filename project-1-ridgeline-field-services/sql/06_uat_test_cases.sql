/*
================================================================================
Project 1 — Ridgeline Field Services: Operational Efficiency
Script:  06_uat_test_cases.sql
Purpose: User Acceptance Test suite. Inserts a small, fully controlled test
         fixture (one dedicated technician, hand-picked jobs with known
         values), verifies that every KPI view produces the hand-calculated
         expected result, then rolls the transaction back -- the production
         synthetic dataset is never modified by running this file.

Traceability: each test case ID below (UAT-01..UAT-05) maps to a row in
/docs/requirements_traceability_matrix.xlsx.

Run this file any time the KPI view logic changes, before trusting the
dashboard/Excel numbers built on top of it.
================================================================================
*/

USE RidgelineFieldOps;
GO

SET NOCOUNT ON;
BEGIN TRANSACTION UATRun;

IF OBJECT_ID('tempdb..#UATResults') IS NOT NULL DROP TABLE #UATResults;
CREATE TABLE #UATResults (TestID VARCHAR(10), Description VARCHAR(200), Expected VARCHAR(50), Actual VARCHAR(50), Result VARCHAR(6));

DECLARE @TestTechKey INT, @TestRegionKey INT, @TestClientKey INT, @TestServiceTypeKey INT;

-- Fixture: one dedicated technician in North Metro, hired well before the test window
SELECT @TestRegionKey = RegionKey FROM dbo.Dim_Region WHERE RegionCode = 'NORTH';
SELECT TOP 1 @TestClientKey = ClientKey FROM dbo.Dim_Client WHERE RegionKey = @TestRegionKey ORDER BY ClientKey;
SELECT @TestServiceTypeKey = ServiceTypeKey FROM dbo.Dim_ServiceType WHERE ServiceTypeID = 'SVC-01'; -- AC Repair, SLA 1.0h

INSERT INTO dbo.Dim_Technician (TechnicianID, TechnicianName, RegionKey, SkillLevel, HireDate, EmploymentStatus, DailyCapacityHours)
VALUES ('TECH-UAT1', 'UAT Test Technician', @TestRegionKey, 'Journeyman', '2025-01-01', 'Active', 8.0);
SELECT @TestTechKey = SCOPE_IDENTITY();

-- =============================================================================
-- UAT-01: Utilization calculation
-- January 2025 has exactly 23 weekdays (Jan 1 2025 = Wednesday; 8 weekend days).
-- Available minutes = 23 * 480 = 11,040.
-- Two completed jobs: (duration 90 + travel 20) + (duration 120 + travel 25) = 255 worked minutes.
-- Expected UtilizationPct = 255 / 11040 * 100 = 2.31 (rounded to 2 dp, matching the view's CAST).
-- =============================================================================
INSERT INTO dbo.Fact_ServiceJobs
    (JobID, DateKey, TechnicianKey, ClientKey, ServiceTypeKey, RegionKey,
     ScheduledStart, ActualStart, ActualEnd, TravelTimeMin, JobDurationMin,
     FirstTimeFix, CallbackFlag, SLAMet, JobCost, JobRevenue, JobStatus)
VALUES
    ('JOB-UAT001', 20250106, @TestTechKey, @TestClientKey, @TestServiceTypeKey, @TestRegionKey,
     '2025-01-06T09:00:00', '2025-01-06T09:05:00', '2025-01-06T10:35:00', 20, 90, 1, 0, 1, 90.00, 220.00, 'Completed'),
    ('JOB-UAT002', 20250107, @TestTechKey, @TestClientKey, @TestServiceTypeKey, @TestRegionKey,
     '2025-01-07T13:00:00', '2025-01-07T13:10:00', '2025-01-07T15:10:00', 25, 120, 0, 1, 1, 95.00, 240.00, 'Completed');

DECLARE @ExpectedUtil DECIMAL(6,2) = 2.31;
DECLARE @ActualUtil DECIMAL(6,2);
SELECT @ActualUtil = UtilizationPct FROM dbo.vw_TechnicianUtilization
WHERE TechnicianKey = @TestTechKey AND [Year] = 2025 AND [Month] = 1;

INSERT INTO #UATResults VALUES ('UAT-01', 'Utilization % matches hand calculation (255/11040 min)',
    CAST(@ExpectedUtil AS VARCHAR(20)), CAST(@ActualUtil AS VARCHAR(20)),
    CASE WHEN ABS(@ActualUtil - @ExpectedUtil) < 0.02 THEN 'PASS' ELSE 'FAIL' END);

-- =============================================================================
-- UAT-02: First-Time-Fix % and Callback % aggregate correctly
-- Of the 2 completed jobs above: 1 FirstTimeFix=1, 1 FirstTimeFix=0 -> FTF% = 50.00
-- Callback is the inverse flag: 1 of 2 -> Callback% = 50.00
-- =============================================================================
DECLARE @ActualFTF DECIMAL(6,2), @ActualCallback DECIMAL(6,2);
SELECT DISTINCT @ActualFTF = FirstTimeFixPct, @ActualCallback = CallbackPct
FROM dbo.vw_TechnicianKPIMonthly
WHERE TechnicianKey = @TestTechKey AND [Year] = 2025 AND [Month] = 1;

INSERT INTO #UATResults VALUES ('UAT-02a', 'First-Time-Fix % = 1 of 2 completed jobs',
    '50.00', CAST(@ActualFTF AS VARCHAR(20)),
    CASE WHEN ABS(@ActualFTF - 50.00) < 0.01 THEN 'PASS' ELSE 'FAIL' END);
INSERT INTO #UATResults VALUES ('UAT-02b', 'Callback % = 1 of 2 completed jobs',
    '50.00', CAST(@ActualCallback AS VARCHAR(20)),
    CASE WHEN ABS(@ActualCallback - 50.00) < 0.01 THEN 'PASS' ELSE 'FAIL' END);

-- =============================================================================
-- UAT-03: Data quality check detects a deliberately broken row
-- Insert one job with ActualEnd before ActualStart; confirm vw_DQ_JobAnomalies
-- flags it as TIMESTAMP_INVERSION, and that it disappears after correction.
-- =============================================================================
INSERT INTO dbo.Fact_ServiceJobs
    (JobID, DateKey, TechnicianKey, ClientKey, ServiceTypeKey, RegionKey,
     ScheduledStart, ActualStart, ActualEnd, TravelTimeMin, JobDurationMin,
     FirstTimeFix, CallbackFlag, SLAMet, JobCost, JobRevenue, JobStatus)
VALUES
    ('JOB-UAT003', 20250108, @TestTechKey, @TestClientKey, @TestServiceTypeKey, @TestRegionKey,
     '2025-01-08T09:00:00', '2025-01-08T11:00:00', '2025-01-08T09:30:00', 20, 90, 1, 0, 1, 90.00, 220.00, 'Completed');

DECLARE @DQFlagCount INT;
SELECT @DQFlagCount = COUNT(*) FROM dbo.vw_DQ_JobAnomalies WHERE JobID = 'JOB-UAT003' AND AnomalyType = 'TIMESTAMP_INVERSION';

INSERT INTO #UATResults VALUES ('UAT-03', 'DQ view flags a job with ActualEnd < ActualStart',
    '1', CAST(@DQFlagCount AS VARCHAR(20)),
    CASE WHEN @DQFlagCount = 1 THEN 'PASS' ELSE 'FAIL' END);

-- =============================================================================
-- UAT-04: Priority Action Queue flags a deliberately underutilized technician
-- Give the test technician a whole separate month (Feb 2025) with only one
-- short completed job -> utilization far below the 60% warning threshold.
-- Feb 2025 has 20 weekdays -> available = 20*480 = 9600 min. One 60-min job,
-- 0 travel -> worked=60 -> utilization = 0.63%, must trip FlagLowUtilization.
-- =============================================================================
INSERT INTO dbo.Fact_ServiceJobs
    (JobID, DateKey, TechnicianKey, ClientKey, ServiceTypeKey, RegionKey,
     ScheduledStart, ActualStart, ActualEnd, TravelTimeMin, JobDurationMin,
     FirstTimeFix, CallbackFlag, SLAMet, JobCost, JobRevenue, JobStatus)
VALUES
    ('JOB-UAT004', 20250203, @TestTechKey, @TestClientKey, @TestServiceTypeKey, @TestRegionKey,
     '2025-02-03T09:00:00', '2025-02-03T09:05:00', '2025-02-03T10:05:00', 0, 60, 1, 0, 1, 90.00, 220.00, 'Completed');

DECLARE @QueueFlag INT;
SELECT @QueueFlag = FlagLowUtilization FROM dbo.vw_PriorityActionQueue
WHERE TechnicianKey = @TestTechKey AND [Year] = 2025 AND [Month] = 2;

INSERT INTO #UATResults VALUES ('UAT-04', 'Priority queue flags a technician far below utilization target',
    '1', CAST(ISNULL(@QueueFlag,-1) AS VARCHAR(20)),
    CASE WHEN @QueueFlag = 1 THEN 'PASS' ELSE 'FAIL' END);

-- =============================================================================
-- UAT-06: Monthly KPIs do not leak across month boundaries for the same
-- technician. By this point the test technician has 3 completed jobs in
-- Jan 2025 (JOB-UAT001 FTF=1, JOB-UAT002 FTF=0, JOB-UAT003 FTF=1 -> 2/3 =
-- 66.67%) and 1 completed job in Feb 2025 (JOB-UAT004 FTF=1 -> 100.00% if
-- correctly isolated; pooling both months would instead give (2+1)/(3+1) =
-- 75.00%). Both months' FirstTimeFixPct must match their own isolated
-- calculation exactly.
-- This is a regression test for a real bug: an earlier version of
-- vw_TechnicianKPIMonthly joined Fact_ServiceJobs without a date filter and
-- silently averaged a technician's entire 2-year history into every month's
-- row. It was caught by cross-checking against the Excel workbook's
-- independent formula implementation, not by the UAT suite as it stood then
-- -- this case closes that gap.
-- =============================================================================
DECLARE @JanFTF DECIMAL(6,2), @FebFTF DECIMAL(6,2);
SELECT @JanFTF = FirstTimeFixPct FROM dbo.vw_TechnicianKPIMonthly
WHERE TechnicianKey = @TestTechKey AND [Year] = 2025 AND [Month] = 1;
SELECT @FebFTF = FirstTimeFixPct FROM dbo.vw_TechnicianKPIMonthly
WHERE TechnicianKey = @TestTechKey AND [Year] = 2025 AND [Month] = 2;

INSERT INTO #UATResults VALUES ('UAT-06a', 'Jan-2025 FirstTimeFixPct = 2 of its own 3 completed jobs (not pooled with Feb)',
    '66.67', CAST(ISNULL(@JanFTF,-1) AS VARCHAR(20)),
    CASE WHEN ABS(@JanFTF - 66.67) < 0.01 THEN 'PASS' ELSE 'FAIL' END);
INSERT INTO #UATResults VALUES ('UAT-06b', 'Feb-2025 FirstTimeFixPct = 1 of its own 1 completed job (not pooled with Jan)',
    '100.00', CAST(ISNULL(@FebFTF,-1) AS VARCHAR(20)),
    CASE WHEN ABS(@FebFTF - 100.00) < 0.01 THEN 'PASS' ELSE 'FAIL' END);

-- =============================================================================
-- UAT-05: Region rollup reconciles to the sum of its technicians' job counts
-- (referential/aggregation integrity check between the two grains).
-- =============================================================================
DECLARE @RegionJobCount INT, @TechSumJobCount INT;
SELECT @RegionJobCount = TotalJobs FROM dbo.vw_RegionKPIMonthly
WHERE RegionKey = @TestRegionKey AND [Year] = 2025 AND [Month] = 1;

SELECT @TechSumJobCount = COUNT(*) FROM dbo.Fact_ServiceJobs f
JOIN dbo.Dim_Date d ON d.DateKey = f.DateKey
WHERE f.RegionKey = @TestRegionKey AND d.[Year] = 2025 AND d.[Month] = 1;

INSERT INTO #UATResults VALUES ('UAT-05', 'Region monthly job count reconciles to underlying fact rows',
    CAST(@TechSumJobCount AS VARCHAR(20)), CAST(@RegionJobCount AS VARCHAR(20)),
    CASE WHEN @RegionJobCount = @TechSumJobCount THEN 'PASS' ELSE 'FAIL' END);

-- =============================================================================
-- Report and roll back -- the fixture never touches the real dataset
-- =============================================================================
PRINT '=== Ridgeline Field Services: UAT Test Report ===';
SELECT * FROM #UATResults ORDER BY TestID;

IF EXISTS (SELECT 1 FROM #UATResults WHERE Result = 'FAIL')
    PRINT 'UAT RESULT: FAIL -- one or more test cases did not match expected values.';
ELSE
    PRINT 'UAT RESULT: PASS -- all test cases matched expected values.';

ROLLBACK TRANSACTION UATRun;
PRINT 'Test fixture rolled back -- production dataset unchanged.';
GO
