/*
================================================================================
Project 1 — Ridgeline Field Services: Operational Efficiency
Script:  05_stored_procedures.sql
Purpose: Parameterized, reusable procedures on top of the KPI views. These
         are what a BI tool, a scheduled job, or an analyst would actually
         call -- not the raw views -- so the calling contract stays stable
         even if the underlying views are refactored later.
================================================================================
*/

USE RidgelineFieldOps;
GO

-- =============================================================================
-- usp_GetRegionKPISummary
-- Region-level KPI rollup for a date range, optionally filtered to one region.
-- =============================================================================
IF OBJECT_ID('dbo.usp_GetRegionKPISummary', 'P') IS NOT NULL DROP PROCEDURE dbo.usp_GetRegionKPISummary;
GO
CREATE PROCEDURE dbo.usp_GetRegionKPISummary
    @StartYear  INT,
    @StartMonth INT,
    @EndYear    INT,
    @EndMonth   INT,
    @RegionCode VARCHAR(10) = NULL      -- NULL = all regions
AS
BEGIN
    SET NOCOUNT ON;

    -- Reject bad arguments instead of returning an empty result set. An empty
    -- grid reads as "no jobs in that period", which is a finding; it should
    -- read as "you asked for month 13", which is a mistake. The two look
    -- identical to the caller, and only one of them is safe to act on.
    IF @StartMonth NOT BETWEEN 1 AND 12 OR @EndMonth NOT BETWEEN 1 AND 12
        THROW 51001, 'StartMonth and EndMonth must each be between 1 and 12.', 1;
    IF (@EndYear * 100 + @EndMonth) < (@StartYear * 100 + @StartMonth)
        THROW 51002, 'The end of the range falls before its start.', 1;
    IF @RegionCode IS NOT NULL AND NOT EXISTS (SELECT 1 FROM dbo.Dim_Region WHERE RegionCode = @RegionCode)
        THROW 51003, 'RegionCode not found in Dim_Region.', 1;

    DECLARE @StartKey INT = @StartYear * 100 + @StartMonth;
    DECLARE @EndKey   INT = @EndYear   * 100 + @EndMonth;

    SELECT k.*
    FROM dbo.vw_RegionKPIMonthly k
    JOIN dbo.Dim_Region r ON r.RegionKey = k.RegionKey
    WHERE (k.[Year] * 100 + k.[Month]) BETWEEN @StartKey AND @EndKey
      AND (@RegionCode IS NULL OR r.RegionCode = @RegionCode)
    ORDER BY k.[Year], k.[Month], k.RegionName;
END
GO

-- =============================================================================
-- usp_GetPriorityActionQueue
-- Top-N technicians needing dispatch review for a given month, ranked by
-- RiskScore.
--
-- This is the SQL-side interface. It is deliberately NOT what the other two
-- artefacts call, and an earlier version of this comment said it was:
--   * the Excel Priority_Action_Queue sheet re-implements the ranking in
--     worksheet formulas (SUMIFS/AVERAGEIFS over tblJobs), and that
--     independence is the point -- it is what makes the Excel-versus-SQL
--     reconciliation evidence of anything at all;
--   * the dashboard exporter reads dbo.vw_PriorityActionQueue directly,
--     because it needs OvertimePctOfCapacity, which this procedure does not
--     return.
-- Claiming one calling contract where there are three is how a reader comes to
-- believe a change here propagates everywhere. It does not.
-- =============================================================================
IF OBJECT_ID('dbo.usp_GetPriorityActionQueue', 'P') IS NOT NULL DROP PROCEDURE dbo.usp_GetPriorityActionQueue;
GO
CREATE PROCEDURE dbo.usp_GetPriorityActionQueue
    @Year   INT,
    @Month  INT,
    @TopN   INT = 20
AS
BEGIN
    SET NOCOUNT ON;

    IF @Month NOT BETWEEN 1 AND 12
        THROW 51004, 'Month must be between 1 and 12.', 1;
    IF @TopN IS NULL OR @TopN < 1
        THROW 51005, 'TopN must be at least 1.', 1;

    SELECT TOP (@TopN)
        TechnicianID, TechnicianName, RegionName, SkillLevel,
        UtilizationPct, FirstTimeFixPct, SLACompliancePct, CallbackPct,
        RiskScore, RecommendedAction
    FROM dbo.vw_PriorityActionQueue
    WHERE [Year] = @Year AND [Month] = @Month
    ORDER BY RiskScore DESC, UtilizationPct ASC;
END
GO

-- =============================================================================
-- usp_GetTechnicianScorecard
-- Full monthly trend for a single technician -- used for the 1:1 coaching
-- conversation a dispatch manager has after the priority queue flags someone.
-- =============================================================================
IF OBJECT_ID('dbo.usp_GetTechnicianScorecard', 'P') IS NOT NULL DROP PROCEDURE dbo.usp_GetTechnicianScorecard;
GO
CREATE PROCEDURE dbo.usp_GetTechnicianScorecard
    @TechnicianID VARCHAR(10)
AS
BEGIN
    SET NOCOUNT ON;

    IF NOT EXISTS (SELECT 1 FROM dbo.Dim_Technician WHERE TechnicianID = @TechnicianID)
    BEGIN
        RAISERROR('TechnicianID %s not found in Dim_Technician.', 16, 1, @TechnicianID);
        RETURN;
    END

    -- An INACTIVE technician exists, so the guard above lets it through -- and
    -- then vw_TechnicianKPIMonthly returns nothing, because the capacity view
    -- beneath it filters EmploymentStatus = 'Active'. The caller gets an empty
    -- grid and exit code 0.
    --
    -- That is the failure mode this file already warns about on
    -- usp_GetRegionKPISummary: an empty grid reads as "this technician had no
    -- jobs", when it actually means "this technician is not in the capacity
    -- model at all". The two look identical to the caller and only one of them
    -- is safe to act on.
    IF EXISTS (SELECT 1 FROM dbo.Dim_Technician
               WHERE TechnicianID = @TechnicianID AND EmploymentStatus <> 'Active')
    BEGIN
        RAISERROR('TechnicianID %s is not Active. The capacity model covers active technicians only, so this scorecard would be empty -- which reads as "no jobs" rather than "not modelled".', 16, 1, @TechnicianID);
        RETURN;
    END

    SELECT DISTINCT k.[Year], k.[Month], k.UtilizationPct, k.JobCount,
           k.FirstTimeFixPct, k.SLACompliancePct, k.CallbackPct
    FROM dbo.vw_TechnicianKPIMonthly k
    WHERE k.TechnicianID = @TechnicianID
    ORDER BY k.[Year], k.[Month];
END
GO

PRINT 'Stored procedures created: usp_GetRegionKPISummary, usp_GetPriorityActionQueue, usp_GetTechnicianScorecard.';
GO
