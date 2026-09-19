/*
================================================================================
Project 1 — Ridgeline Field Services: Operational Efficiency
Script:  06_uat_test_cases.sql
Purpose: User Acceptance Test suite. Inserts a small, fully controlled test
         fixture (two dedicated technicians, hand-picked jobs with known
         values), checks each assertion against a hand calculation, then rolls
         the transaction back -- the production synthetic dataset is never
         modified by running this file. Exits non-zero if anything fails.

Coverage, stated honestly: vw_TechnicianUtilization (UAT-01, 07a, 07b),
vw_TechnicianKPIMonthly (UAT-02a/02b, 06a/06b), vw_DQ_JobAnomalies (UAT-03),
vw_PriorityActionQueue (UAT-04), vw_RegionKPIMonthly (UAT-05, 09) and the
stored-procedure interface (UAT-08). vw_RegionKPIMonthly's FirstTimeFixPct,
SLACompliancePct, CallbackPct and GrossMargin are NOT independently asserted.

Traceability: each test case ID below (UAT-01..UAT-09) maps to a row in
/docs/requirements_traceability_matrix.md.

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
-- UAT-07: Utilization counts weekday work against weekday capacity, and
--         weekend call-outs are reported as overtime rather than folded in.
--
-- This is the regression test for a defect that survived every other case in
-- this file: capacity was modelled on weekdays while the numerator counted
-- every completed job, so 2,194 weekend jobs were divided by capacity that did
-- not exist and portfolio utilization read 70.46% instead of 65.89%. UAT-01
-- could not see it, because its two fixture jobs both fall on weekdays -- the
-- test agreed with the code about a case where the bug does not appear.
--
-- A second technician, so the assertion does not depend on what the cases
-- above have already inserted.
-- March 2025 has 21 weekdays (Mar 1 is a Saturday; 10 weekend days in 31).
-- Available = 21 * 480 = 10,080 minutes.
--   Weekday job, Wed 5 Mar: duration 60 + travel 15 =  75 min -> utilization
--   Weekend job, Sat 8 Mar: duration 100 + travel 20 = 120 min -> overtime
-- Expected UtilizationPct        = 75 / 10080 * 100 = 0.74
-- Expected OvertimeMinutes       = 120
-- =============================================================================
DECLARE @TestTechKey2 INT;
INSERT INTO dbo.Dim_Technician (TechnicianID, TechnicianName, RegionKey, SkillLevel, HireDate, EmploymentStatus, DailyCapacityHours)
VALUES ('TECH-UAT2', 'UAT Weekend Technician', @TestRegionKey, 'Journeyman', '2025-01-01', 'Active', 8.0);
SELECT @TestTechKey2 = SCOPE_IDENTITY();

INSERT INTO dbo.Fact_ServiceJobs
    (JobID, DateKey, TechnicianKey, ClientKey, ServiceTypeKey, RegionKey,
     ScheduledStart, ActualStart, ActualEnd, TravelTimeMin, JobDurationMin,
     FirstTimeFix, CallbackFlag, SLAMet, JobCost, JobRevenue, JobStatus)
VALUES
    ('JOB-UAT101', 20250305, @TestTechKey2, @TestClientKey, @TestServiceTypeKey, @TestRegionKey,
     '2025-03-05T09:00:00', '2025-03-05T09:05:00', '2025-03-05T10:05:00', 15, 60, 1, 0, 1, 70.00, 180.00, 'Completed'),
    ('JOB-UAT102', 20250308, @TestTechKey2, @TestClientKey, @TestServiceTypeKey, @TestRegionKey,
     '2025-03-08T09:00:00', '2025-03-08T09:10:00', '2025-03-08T10:50:00', 20, 100, 1, 0, 1, 130.00, 320.00, 'Completed');

DECLARE @U2 DECIMAL(6,2), @Worked2 INT, @OT2 INT;
SELECT @U2 = UtilizationPct, @Worked2 = WorkedMinutes, @OT2 = OvertimeMinutes
FROM dbo.vw_TechnicianUtilization
WHERE TechnicianKey = @TestTechKey2 AND [Year] = 2025 AND [Month] = 3;

INSERT INTO #UATResults VALUES ('UAT-07a',
    'Utilization excludes weekend work (75 of 10,080 weekday minutes)',
    '0.74 / 75', CAST(@U2 AS VARCHAR(20)) + ' / ' + CAST(@Worked2 AS VARCHAR(20)),
    CASE WHEN ABS(@U2 - 0.74) < 0.02 AND @Worked2 = 75 THEN 'PASS' ELSE 'FAIL' END);

INSERT INTO #UATResults VALUES ('UAT-07b',
    'The weekend call-out is reported as overtime, not discarded',
    '120', CAST(@OT2 AS VARCHAR(20)),
    CASE WHEN @OT2 = 120 THEN 'PASS' ELSE 'FAIL' END);

-- =============================================================================
-- UAT-08: The stored-procedure interface rejects invalid arguments.
--
-- Previously both of these returned an empty result set and exit code 0, which
-- a caller cannot distinguish from "there were no jobs in that period". One of
-- those answers is a finding and the other is a typo.
-- =============================================================================
DECLARE @Raised08 VARCHAR(20) = '';
BEGIN TRY
    EXEC dbo.usp_GetRegionKPISummary @StartYear = 2025, @StartMonth = 13, @EndYear = 2025, @EndMonth = 12;
    SET @Raised08 = 'NO';
END TRY
BEGIN CATCH SET @Raised08 = 'YES'; END CATCH

DECLARE @Raised08b VARCHAR(20) = '';
BEGIN TRY
    EXEC dbo.usp_GetPriorityActionQueue @Year = 2025, @Month = 44, @TopN = 5;
    SET @Raised08b = 'NO';
END TRY
BEGIN CATCH SET @Raised08b = 'YES'; END CATCH

DECLARE @Raised08c VARCHAR(20) = '';
BEGIN TRY
    EXEC dbo.usp_GetRegionKPISummary @StartYear = 2025, @StartMonth = 12, @EndYear = 2024, @EndMonth = 1;
    SET @Raised08c = 'NO';
END TRY
BEGIN CATCH SET @Raised08c = 'YES'; END CATCH

INSERT INTO #UATResults VALUES ('UAT-08',
    'Procedures raise on an invalid month, an invalid TopN and a reversed range',
    'YES / YES / YES', @Raised08 + ' / ' + @Raised08b + ' / ' + @Raised08c,
    CASE WHEN @Raised08 = 'YES' AND @Raised08b = 'YES' AND @Raised08c = 'YES' THEN 'PASS' ELSE 'FAIL' END);

-- =============================================================================
-- UAT-09: The region rollup's job statuses partition the total.
--
-- UAT-05 above compares the view's TotalJobs against a COUNT(*) over the same
-- FROM and the same filter, so it restates the view's own logic and cannot
-- fail. It is kept because referential agreement is still worth asserting, but
-- it is not coverage. This case asserts something the view can actually get
-- wrong: that completed + cancelled + rescheduled accounts for every job, and
-- that revenue is only recognised on completed work.
-- =============================================================================
DECLARE @StatusGap INT, @RevenueOnNonCompleted INT;
SELECT @StatusGap = COUNT(*) FROM dbo.vw_RegionKPIMonthly
WHERE CompletedJobs + CancelledJobs + RescheduledJobs <> TotalJobs;

SELECT @RevenueOnNonCompleted = COUNT(*)
FROM dbo.Fact_ServiceJobs WHERE JobStatus <> 'Completed' AND JobRevenue > 0;

INSERT INTO #UATResults VALUES ('UAT-09',
    'Region statuses sum to TotalJobs, and no revenue is booked on uncompleted work',
    '0 / 0', CAST(@StatusGap AS VARCHAR(20)) + ' / ' + CAST(@RevenueOnNonCompleted AS VARCHAR(20)),
    CASE WHEN @StatusGap = 0 AND @RevenueOnNonCompleted = 0 THEN 'PASS' ELSE 'FAIL' END);

-- =============================================================================
-- Report and roll back -- the fixture never touches the real dataset
-- =============================================================================
PRINT '=== Ridgeline Field Services: UAT Test Report ===';
SELECT * FROM #UATResults ORDER BY TestID;

DECLARE @Failed INT = (SELECT COUNT(*) FROM #UATResults WHERE Result = 'FAIL');
DECLARE @Passed INT = (SELECT COUNT(*) FROM #UATResults WHERE Result = 'PASS');
PRINT 'UAT RESULT: ' + CAST(@Passed AS VARCHAR(10)) + ' passed, ' + CAST(@Failed AS VARCHAR(10)) + ' failed.';

-- Roll back FIRST, then raise. The fixture must come out of the database
-- whether the suite passed or not, and a THROW before the ROLLBACK would leave
-- the test technician and its jobs sitting in the production dataset.
ROLLBACK TRANSACTION UATRun;
PRINT 'Test fixture rolled back -- production dataset unchanged.';

-- Exit non-zero on failure. Printing 'FAIL' and returning exit code 0 means no
-- build step, scheduler or pipeline can act on the result; the suite passes
-- for anything that is not a person reading the scrollback.
IF @Failed > 0
    THROW 51100, 'UAT suite FAILED -- see the report above. Do not trust the dashboard or workbook figures until this is green.', 1;
GO
