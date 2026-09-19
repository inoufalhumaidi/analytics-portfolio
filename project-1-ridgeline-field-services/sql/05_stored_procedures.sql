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

    SELECT DISTINCT k.[Year], k.[Month], k.UtilizationPct, k.JobCount,
           k.FirstTimeFixPct, k.SLACompliancePct, k.CallbackPct
    FROM dbo.vw_TechnicianKPIMonthly k
    WHERE k.TechnicianID = @TechnicianID
    ORDER BY k.[Year], k.[Month];
END
GO

PRINT 'Stored procedures created: usp_GetRegionKPISummary, usp_GetPriorityActionQueue, usp_GetTechnicianScorecard.';
GO
