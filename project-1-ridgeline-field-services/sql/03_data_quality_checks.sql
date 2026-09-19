/*
================================================================================
Project 1 — Ridgeline Field Services: Operational Efficiency
Script:  03_data_quality_checks.sql
Purpose: Reusable data-quality control layer. Creates:
   - dbo.vw_DQ_JobAnomalies   : row-level anomaly detection (one row per
                                 anomalous job, tagged with anomaly type)
   - dbo.vw_DQ_Summary        : rollup of anomaly counts/rates for monitoring
   - dbo.usp_RunDataQualityChecks : prints a pass/fail QA report, used both
                                 standalone and by the UAT test suite

These checks catch issues that CHECK constraints cannot: logical
inconsistencies across columns (e.g. ActualEnd before ActualStart),
completeness gaps conditional on business state (e.g. a Completed job with
no actual timestamps), and statistical outliers.
================================================================================
*/

USE RidgelineFieldOps;
GO

IF OBJECT_ID('dbo.vw_DQ_JobAnomalies', 'V') IS NOT NULL DROP VIEW dbo.vw_DQ_JobAnomalies;
GO
CREATE VIEW dbo.vw_DQ_JobAnomalies AS
SELECT
    f.JobKey, f.JobID, f.JobStatus, f.DateKey, f.TechnicianKey, f.RegionKey,
    'TIMESTAMP_INVERSION' AS AnomalyType,
    'ActualEnd is earlier than ActualStart' AS AnomalyDescription
FROM dbo.Fact_ServiceJobs f
WHERE f.ActualStart IS NOT NULL AND f.ActualEnd IS NOT NULL AND f.ActualEnd < f.ActualStart

UNION ALL
SELECT
    f.JobKey, f.JobID, f.JobStatus, f.DateKey, f.TechnicianKey, f.RegionKey,
    'MISSING_ACTUALS',
    'JobStatus = Completed but ActualStart/ActualEnd is NULL'
FROM dbo.Fact_ServiceJobs f
WHERE f.JobStatus = 'Completed' AND (f.ActualStart IS NULL OR f.ActualEnd IS NULL)

UNION ALL
SELECT
    f.JobKey, f.JobID, f.JobStatus, f.DateKey, f.TechnicianKey, f.RegionKey,
    'EXTREME_DURATION',
    'JobDurationMin exceeds 600 minutes (10 hours) for a single job -- likely a unit/entry error'
FROM dbo.Fact_ServiceJobs f
WHERE f.JobDurationMin > 600

UNION ALL
SELECT
    f.JobKey, f.JobID, f.JobStatus, f.DateKey, f.TechnicianKey, f.RegionKey,
    'ORPHAN_DIMENSION_KEY',
    'Fact row references a key not present in one or more dimension tables'
FROM dbo.Fact_ServiceJobs f
WHERE NOT EXISTS (SELECT 1 FROM dbo.Dim_Date d WHERE d.DateKey = f.DateKey)
   OR NOT EXISTS (SELECT 1 FROM dbo.Dim_Technician t WHERE t.TechnicianKey = f.TechnicianKey)
   OR NOT EXISTS (SELECT 1 FROM dbo.Dim_Client c WHERE c.ClientKey = f.ClientKey)
   OR NOT EXISTS (SELECT 1 FROM dbo.Dim_ServiceType s WHERE s.ServiceTypeKey = f.ServiceTypeKey)
   OR NOT EXISTS (SELECT 1 FROM dbo.Dim_Region r WHERE r.RegionKey = f.RegionKey)

UNION ALL
SELECT
    f.JobKey, f.JobID, f.JobStatus, f.DateKey, f.TechnicianKey, f.RegionKey,
    'INCONSISTENT_COMPLETION_FLAGS',
    'Completed job has NULL FirstTimeFix or NULL SLAMet (should always be populated when Completed)'
FROM dbo.Fact_ServiceJobs f
WHERE f.JobStatus = 'Completed' AND (f.FirstTimeFix IS NULL OR f.SLAMet IS NULL);
GO

IF OBJECT_ID('dbo.vw_DQ_Summary', 'V') IS NOT NULL DROP VIEW dbo.vw_DQ_Summary;
GO
CREATE VIEW dbo.vw_DQ_Summary AS
-- Every check that exists in vw_DQ_JobAnomalies is enumerated here so a
-- check that finds zero issues still reports "0 found", not silence.
WITH KnownChecks AS (
    SELECT AnomalyType FROM (VALUES
        ('TIMESTAMP_INVERSION'), ('MISSING_ACTUALS'), ('EXTREME_DURATION'),
        ('ORPHAN_DIMENSION_KEY'), ('INCONSISTENT_COMPLETION_FLAGS')
    ) c(AnomalyType)
)
SELECT
    kc.AnomalyType,
    COUNT(a.JobKey) AS AnomalyCount,
    (SELECT COUNT(*) FROM dbo.Fact_ServiceJobs) AS TotalFactRows,
    CAST(COUNT(a.JobKey) * 100.0 / NULLIF((SELECT COUNT(*) FROM dbo.Fact_ServiceJobs), 0) AS DECIMAL(6,3)) AS AnomalyRatePct
FROM KnownChecks kc
LEFT JOIN dbo.vw_DQ_JobAnomalies a ON a.AnomalyType = kc.AnomalyType
GROUP BY kc.AnomalyType;
GO

IF OBJECT_ID('dbo.usp_RunDataQualityChecks', 'P') IS NOT NULL DROP PROCEDURE dbo.usp_RunDataQualityChecks;
GO
CREATE PROCEDURE dbo.usp_RunDataQualityChecks
    @MaxAcceptableRatePct DECIMAL(6,3) = 2.000   -- QA gate: fail if any single anomaly type exceeds this rate
AS
BEGIN
    SET NOCOUNT ON;

    DECLARE @FailCount INT;

    PRINT '=== Ridgeline Field Services: Data Quality Report ===';

    SELECT AnomalyType, AnomalyCount, TotalFactRows, AnomalyRatePct,
           CASE WHEN AnomalyRatePct > @MaxAcceptableRatePct THEN 'FAIL' ELSE 'PASS' END AS QAResult
    FROM dbo.vw_DQ_Summary
    ORDER BY AnomalyRatePct DESC;

    SELECT @FailCount = COUNT(*) FROM dbo.vw_DQ_Summary WHERE AnomalyRatePct > @MaxAcceptableRatePct;

    IF @FailCount = 0
        PRINT 'QA GATE: PASS -- all anomaly rates within the ' + CAST(@MaxAcceptableRatePct AS VARCHAR(10)) + '% threshold.';
    ELSE
    BEGIN
        PRINT 'QA GATE: FAIL -- ' + CAST(@FailCount AS VARCHAR(10)) + ' anomaly type(s) exceed the ' + CAST(@MaxAcceptableRatePct AS VARCHAR(10)) + '% threshold. Review dbo.vw_DQ_JobAnomalies.';
        -- A gate that only PRINTs is not a gate. sqlcmd returned exit code 0
        -- whether this passed or failed, so no build step, scheduler or
        -- pipeline could ever act on the result -- the word "gate" was doing
        -- work the code was not. THROW makes the failure observable to
        -- something other than a human reading the scrollback.
        ;THROW 51000, 'QA gate failed: one or more anomaly rates exceed the acceptable threshold. See the report above and dbo.vw_DQ_JobAnomalies.', 1;
    END
END
GO

PRINT 'Data quality objects created: vw_DQ_JobAnomalies, vw_DQ_Summary, usp_RunDataQualityChecks.';
GO

EXEC dbo.usp_RunDataQualityChecks;
GO
