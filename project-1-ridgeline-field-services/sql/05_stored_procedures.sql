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
-- RiskScore. This is the exact query the Excel Priority_Action_Queue sheet
-- and the live dashboard's "This Week's Actions" panel both call.
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

    SELECT DISTINCT k.[Year], k.[Month], k.UtilizationPct, k.JobCount,
           k.FirstTimeFixPct, k.SLACompliancePct, k.CallbackPct
    FROM dbo.vw_TechnicianKPIMonthly k
    WHERE k.TechnicianID = @TechnicianID
    ORDER BY k.[Year], k.[Month];
END
GO

PRINT 'Stored procedures created: usp_GetRegionKPISummary, usp_GetPriorityActionQueue, usp_GetTechnicianScorecard.';
GO
