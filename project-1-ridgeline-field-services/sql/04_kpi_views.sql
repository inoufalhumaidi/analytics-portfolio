/*
================================================================================
Project 1 — Ridgeline Field Services: Operational Efficiency
Script:  04_kpi_views.sql
Purpose: Reusable KPI views. These are the single source of truth consumed
         by the Excel workbook, the Power BI model (via DAX measures built
         on the same star schema), and the live HTML dashboard. Keeping the
         business logic here (not duplicated per-consumer) is what makes the
         SQL layer "reusable" rather than one-off.

Business question this layer answers:
   "Where is technician capacity being wasted, and which dispatch decisions
    should change this week?"
================================================================================
*/

USE RidgelineFieldOps;
GO

-- =============================================================================
-- vw_TechnicianDailyCapacity
-- One row per (Technician, Date) they were Active -- the denominator for
-- utilization. A technician "counts" as available on a date if their
-- HireDate <= Date and they are Active (this synthetic model does not track
-- termination dates separately -- documented limitation).
-- =============================================================================
IF OBJECT_ID('dbo.vw_TechnicianDailyCapacity', 'V') IS NOT NULL DROP VIEW dbo.vw_TechnicianDailyCapacity;
GO
CREATE VIEW dbo.vw_TechnicianDailyCapacity AS
SELECT
    t.TechnicianKey, t.TechnicianID, t.TechnicianName, t.RegionKey, t.SkillLevel,
    d.DateKey, d.[Date],
    t.DailyCapacityHours * 60 AS CapacityMinutes
FROM dbo.Dim_Technician t
CROSS JOIN dbo.Dim_Date d
WHERE t.EmploymentStatus = 'Active'
  AND d.[Date] >= t.HireDate
  AND d.IsWeekend = 0;      -- capacity is modeled on scheduled weekdays
GO

-- =============================================================================
-- vw_TechnicianUtilization
-- Grain: one row per Technician per Month.
-- UtilizationPct = (job duration + travel time actually worked) / available capacity
-- =============================================================================
IF OBJECT_ID('dbo.vw_TechnicianUtilization', 'V') IS NOT NULL DROP VIEW dbo.vw_TechnicianUtilization;
GO
CREATE VIEW dbo.vw_TechnicianUtilization AS
WITH Capacity AS (
    SELECT TechnicianKey, TechnicianID, TechnicianName, RegionKey, SkillLevel,
           YEAR([Date]) AS [Year], MONTH([Date]) AS [Month],
           SUM(CapacityMinutes) AS AvailableMinutes
    FROM dbo.vw_TechnicianDailyCapacity
    GROUP BY TechnicianKey, TechnicianID, TechnicianName, RegionKey, SkillLevel, YEAR([Date]), MONTH([Date])
),
Worked AS (
    SELECT f.TechnicianKey, d.[Year], d.[Month],
           SUM(ISNULL(f.JobDurationMin,0) + ISNULL(f.TravelTimeMin,0)) AS WorkedMinutes,
           COUNT(*) AS JobCount
    FROM dbo.Fact_ServiceJobs f
    JOIN dbo.Dim_Date d ON d.DateKey = f.DateKey
    WHERE f.JobStatus = 'Completed'
    GROUP BY f.TechnicianKey, d.[Year], d.[Month]
)
SELECT
    c.TechnicianKey, c.TechnicianID, c.TechnicianName, c.RegionKey, c.SkillLevel,
    c.[Year], c.[Month],
    c.AvailableMinutes,
    ISNULL(w.WorkedMinutes, 0) AS WorkedMinutes,
    ISNULL(w.JobCount, 0) AS JobCount,
    CAST(ISNULL(w.WorkedMinutes, 0) * 100.0 / NULLIF(c.AvailableMinutes, 0) AS DECIMAL(6,2)) AS UtilizationPct
FROM Capacity c
LEFT JOIN Worked w ON w.TechnicianKey = c.TechnicianKey AND w.[Year] = c.[Year] AND w.[Month] = c.[Month];
GO

-- =============================================================================
-- vw_RegionKPIMonthly
-- Grain: one row per Region per Month -- the primary dispatch-manager view.
-- =============================================================================
IF OBJECT_ID('dbo.vw_RegionKPIMonthly', 'V') IS NOT NULL DROP VIEW dbo.vw_RegionKPIMonthly;
GO
CREATE VIEW dbo.vw_RegionKPIMonthly AS
WITH JobFacts AS (
    SELECT
        f.RegionKey, d.[Year], d.[Month],
        f.JobStatus, f.FirstTimeFix, f.CallbackFlag, f.SLAMet, f.JobCost, f.JobRevenue
    FROM dbo.Fact_ServiceJobs f
    JOIN dbo.Dim_Date d ON d.DateKey = f.DateKey
),
Util AS (
    SELECT RegionKey, [Year], [Month],
           CAST(AVG(UtilizationPct) AS DECIMAL(6,2)) AS AvgUtilizationPct
    FROM dbo.vw_TechnicianUtilization
    GROUP BY RegionKey, [Year], [Month]
)
SELECT
    r.RegionKey, r.RegionName, jf.[Year], jf.[Month],
    COUNT(*) AS TotalJobs,
    SUM(CASE WHEN jf.JobStatus = 'Completed' THEN 1 ELSE 0 END) AS CompletedJobs,
    SUM(CASE WHEN jf.JobStatus = 'Cancelled' THEN 1 ELSE 0 END) AS CancelledJobs,
    SUM(CASE WHEN jf.JobStatus = 'Rescheduled' THEN 1 ELSE 0 END) AS RescheduledJobs,
    CAST(AVG(CASE WHEN jf.JobStatus = 'Completed' THEN CAST(jf.FirstTimeFix AS FLOAT) END) * 100.0 AS DECIMAL(6,2)) AS FirstTimeFixPct,
    CAST(AVG(CASE WHEN jf.JobStatus = 'Completed' THEN CAST(jf.SLAMet AS FLOAT) END) * 100.0 AS DECIMAL(6,2)) AS SLACompliancePct,
    CAST(AVG(CASE WHEN jf.JobStatus = 'Completed' THEN CAST(jf.CallbackFlag AS FLOAT) END) * 100.0 AS DECIMAL(6,2)) AS CallbackPct,
    ISNULL(u.AvgUtilizationPct, 0) AS AvgUtilizationPct,
    SUM(jf.JobRevenue) AS TotalRevenue,
    SUM(jf.JobCost) AS TotalCost,
    SUM(jf.JobRevenue) - SUM(jf.JobCost) AS GrossMargin
FROM JobFacts jf
JOIN dbo.Dim_Region r ON r.RegionKey = jf.RegionKey
LEFT JOIN Util u ON u.RegionKey = jf.RegionKey AND u.[Year] = jf.[Year] AND u.[Month] = jf.[Month]
GROUP BY r.RegionKey, r.RegionName, jf.[Year], jf.[Month], u.AvgUtilizationPct;
GO

-- =============================================================================
-- vw_TechnicianKPIMonthly
-- Grain: one row per Technician per Month -- feeds the priority action queue.
-- =============================================================================
IF OBJECT_ID('dbo.vw_TechnicianKPIMonthly', 'V') IS NOT NULL DROP VIEW dbo.vw_TechnicianKPIMonthly;
GO
CREATE VIEW dbo.vw_TechnicianKPIMonthly AS
SELECT
    u.TechnicianKey, u.TechnicianID, u.TechnicianName, u.RegionKey, r.RegionName, u.SkillLevel,
    u.[Year], u.[Month],
    u.UtilizationPct, u.JobCount,
    CAST(AVG(CASE WHEN f.JobStatus = 'Completed' THEN CAST(f.FirstTimeFix AS FLOAT) END) * 100.0 AS DECIMAL(6,2)) AS FirstTimeFixPct,
    CAST(AVG(CASE WHEN f.JobStatus = 'Completed' THEN CAST(f.SLAMet AS FLOAT) END) * 100.0 AS DECIMAL(6,2)) AS SLACompliancePct,
    CAST(AVG(CASE WHEN f.JobStatus = 'Completed' THEN CAST(f.CallbackFlag AS FLOAT) END) * 100.0 AS DECIMAL(6,2)) AS CallbackPct
FROM dbo.vw_TechnicianUtilization u
JOIN dbo.Dim_Region r ON r.RegionKey = u.RegionKey
LEFT JOIN dbo.Dim_Date d ON d.[Year] = u.[Year] AND d.[Month] = u.[Month]
LEFT JOIN dbo.Fact_ServiceJobs f ON f.TechnicianKey = u.TechnicianKey AND f.DateKey = d.DateKey
GROUP BY u.TechnicianKey, u.TechnicianID, u.TechnicianName, u.RegionKey, r.RegionName, u.SkillLevel,
         u.[Year], u.[Month], u.UtilizationPct, u.JobCount;
GO

-- Fixed 2026-09-06: the original version LEFT JOINed Fact_ServiceJobs on
-- TechnicianKey alone (no date filter) and only restricted Dim_Date to the
-- target Year/Month -- so a technician's FirstTimeFix/SLAMet/CallbackFlag
-- averages silently included every job in their 2-year history, not just the
-- selected month. Caught by cross-checking against the Excel workbook's
-- independent formula re-implementation, which disagreed and was right.
-- The current version joins Dim_Date to the target month FIRST, then joins
-- jobs on (TechnicianKey, DateKey) to one of those dates -- this also
-- eliminates the fan-out/duplicate-row issue the old window-function version
-- had (no DISTINCT needed downstream anymore, though it's left in place in
-- vw_PriorityActionQueue and usp_GetTechnicianScorecard as harmless defense).

-- =============================================================================
-- vw_PriorityActionQueue
-- The operational control output: one row per Technician per Month that
-- breaches at least one target in dbo.Ref_KPITargets, ranked by a composite
-- RiskScore, with a plain-language RecommendedAction. This is what a
-- dispatch manager reads on Monday morning.
-- =============================================================================
IF OBJECT_ID('dbo.vw_PriorityActionQueue', 'V') IS NOT NULL DROP VIEW dbo.vw_PriorityActionQueue;
GO
CREATE VIEW dbo.vw_PriorityActionQueue AS
WITH Base AS (
    SELECT DISTINCT
        TechnicianKey, TechnicianID, TechnicianName, RegionKey, RegionName, SkillLevel,
        [Year], [Month], UtilizationPct, JobCount, FirstTimeFixPct, SLACompliancePct, CallbackPct
    FROM dbo.vw_TechnicianKPIMonthly
),
Targets AS (
    SELECT
        MAX(CASE WHEN MetricName = 'UtilizationPct' THEN WarningValue END) AS UtilWarn,
        MAX(CASE WHEN MetricName = 'UtilizationCeilingPct' THEN WarningValue END) AS UtilCeilingWarn,
        MAX(CASE WHEN MetricName = 'FirstTimeFixPct' THEN WarningValue END) AS FTFWarn,
        MAX(CASE WHEN MetricName = 'SLACompliancePct' THEN WarningValue END) AS SLAWarn,
        MAX(CASE WHEN MetricName = 'CallbackPct' THEN WarningValue END) AS CallbackWarn
    FROM dbo.Ref_KPITargets
),
Flagged AS (
    SELECT b.*,
        CASE WHEN b.UtilizationPct < t.UtilWarn THEN 1 ELSE 0 END AS FlagLowUtilization,
        CASE WHEN b.UtilizationPct > t.UtilCeilingWarn THEN 1 ELSE 0 END AS FlagOverCapacity,
        CASE WHEN b.FirstTimeFixPct < t.FTFWarn THEN 1 ELSE 0 END AS FlagLowFTF,
        CASE WHEN b.SLACompliancePct < t.SLAWarn THEN 1 ELSE 0 END AS FlagLowSLA,
        CASE WHEN b.CallbackPct > t.CallbackWarn THEN 1 ELSE 0 END AS FlagHighCallback
    FROM Base b
    CROSS JOIN Targets t
)
SELECT
    TechnicianKey, TechnicianID, TechnicianName, RegionKey, RegionName, SkillLevel,
    [Year], [Month], UtilizationPct, JobCount, FirstTimeFixPct, SLACompliancePct, CallbackPct,
    FlagLowUtilization, FlagOverCapacity, FlagLowFTF, FlagLowSLA, FlagHighCallback,
    (FlagLowUtilization + FlagOverCapacity + FlagLowFTF + FlagLowSLA + FlagHighCallback) AS RiskScore,
    CONCAT_WS('; ',
        CASE WHEN FlagLowUtilization = 1 THEN 'Underutilized -- reassign jobs from over-capacity peers or reduce headcount buffer' END,
        CASE WHEN FlagOverCapacity  = 1 THEN 'Over capacity -- redistribute jobs to underutilized technicians in region' END,
        CASE WHEN FlagLowFTF        = 1 THEN 'Low first-time-fix -- pair with a Master technician or schedule refresher training' END,
        CASE WHEN FlagLowSLA        = 1 THEN 'SLA compliance below target -- review dispatch routing / travel radius' END,
        CASE WHEN FlagHighCallback  = 1 THEN 'High callback rate -- audit recent completed jobs for rework root cause' END
    ) AS RecommendedAction
FROM Flagged
WHERE (FlagLowUtilization + FlagOverCapacity + FlagLowFTF + FlagLowSLA + FlagHighCallback) > 0;
GO

PRINT 'KPI views created: vw_TechnicianDailyCapacity, vw_TechnicianUtilization, vw_RegionKPIMonthly, vw_TechnicianKPIMonthly, vw_PriorityActionQueue.';
GO
