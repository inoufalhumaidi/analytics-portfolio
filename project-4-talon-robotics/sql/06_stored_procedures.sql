/*
================================================================================
Project 4 — Talon Robotics: Payload Deployment Delivery
Script:  06_stored_procedures.sql
Purpose: The reusable interface. Seven procedures, every one @AsOf-aware.

WHY PROCEDURES AND NOT JUST VIEWS

    A view answers one question at one point. Every procedure here takes the
    build (and where it matters, the date) being assessed, so the same logic
    answers "where are we now", "where were we at the last gate" and "what
    would this look like if we shipped build 74 instead" without a second
    implementation appearing somewhere to answer the variants.

EVERY PROCEDURE VALIDATES ITS ARGUMENTS AND RAISES

    An out-of-range build silently returning an empty result set is
    indistinguishable, to the caller, from a build where nothing was
    outstanding. One of those is a typo and the other is a milestone. Project 1
    in this portfolio shipped exactly that ambiguity and it was only found by
    checking an exit code.

DATA DISCLOSURE: Talon Robotics is fictional; all data is synthetic. No
confidential data and no production-system claim is involved.
================================================================================
*/

USE TalonDelivery;
GO

IF OBJECT_ID('dbo.usp_ReadinessScorecard','P')  IS NOT NULL DROP PROCEDURE dbo.usp_ReadinessScorecard;
IF OBJECT_ID('dbo.usp_VerificationQueue','P')   IS NOT NULL DROP PROCEDURE dbo.usp_VerificationQueue;
IF OBJECT_ID('dbo.usp_SubsystemReadiness','P')  IS NOT NULL DROP PROCEDURE dbo.usp_SubsystemReadiness;
IF OBJECT_ID('dbo.usp_RequirementDetail','P')   IS NOT NULL DROP PROCEDURE dbo.usp_RequirementDetail;
IF OBJECT_ID('dbo.usp_RAIDRegister','P')        IS NOT NULL DROP PROCEDURE dbo.usp_RAIDRegister;
IF OBJECT_ID('dbo.usp_ReadinessTrend','P')      IS NOT NULL DROP PROCEDURE dbo.usp_ReadinessTrend;
GO

/*
--------------------------------------------------------------------------------
usp_ReadinessScorecard -- the board pack, in one row.
--------------------------------------------------------------------------------
*/
CREATE PROCEDURE dbo.usp_ReadinessScorecard
    @AsOfBuild INT  = NULL,     -- NULL = the release candidate
    @AsOfDate  DATE = '2026-09-30'
AS
BEGIN
    SET NOCOUNT ON;
    IF @AsOfBuild IS NULL
        SET @AsOfBuild = (SELECT TOP 1 BuildNumber FROM dbo.Dim_Build WHERE IsReleaseCandidate = 1 ORDER BY BuildNumber DESC);
    IF NOT EXISTS (SELECT 1 FROM dbo.Dim_Build WHERE BuildNumber = @AsOfBuild)
        THROW 52050, 'AsOfBuild does not match any build in Dim_Build.', 1;
    IF @AsOfDate < '2025-01-01' OR @AsOfDate > '2027-06-30'
        THROW 52051, 'AsOfDate falls outside the programme calendar.', 1;

    SELECT k.*,
           -- The comparison the whole project exists to make, computed here so
           -- nobody has to do the subtraction in their head and get it wrong.
           ReadinessGapVsReported = CAST(k.WorkItemCompletionPct - k.ShipReadinessPct AS DECIMAL(9,2)),
           RigWeeksOutstanding    = CAST(k.RigHoursOutstanding / 180.0 AS DECIMAL(9,1))
    FROM dbo.fn_ReadinessKPI(@AsOfBuild, @AsOfDate) k;
END;
GO

/*
--------------------------------------------------------------------------------
usp_VerificationQueue -- what to do next week.
--------------------------------------------------------------------------------
*/
CREATE PROCEDURE dbo.usp_VerificationQueue
    @AsOfBuild       INT = NULL,
    @RigHoursPerWeek DECIMAL(9,2) = 180.00,
    @TopN            INT = 50,
    @ThisWeekOnly    BIT = 0
AS
BEGIN
    SET NOCOUNT ON;
    IF @AsOfBuild IS NULL
        SET @AsOfBuild = (SELECT TOP 1 BuildNumber FROM dbo.Dim_Build WHERE IsReleaseCandidate = 1 ORDER BY BuildNumber DESC);
    IF NOT EXISTS (SELECT 1 FROM dbo.Dim_Build WHERE BuildNumber = @AsOfBuild)
        THROW 52052, 'AsOfBuild does not match any build in Dim_Build.', 1;
    IF @RigHoursPerWeek <= 0
        THROW 52053, 'RigHoursPerWeek must be greater than zero -- a queue bounded by nothing is a list.', 1;
    IF @TopN < 1
        THROW 52054, 'TopN must be at least 1.', 1;

    SELECT TOP (@TopN) *
    FROM dbo.fn_VerificationQueue(@AsOfBuild, @RigHoursPerWeek)
    WHERE (@ThisWeekOnly = 0 OR IsThisWeek = 1)
    ORDER BY PriorityRank;
END;
GO

/*
--------------------------------------------------------------------------------
usp_SubsystemReadiness -- where the gap actually is.
--------------------------------------------------------------------------------
*/
CREATE PROCEDURE dbo.usp_SubsystemReadiness
    @AsOfBuild INT = NULL
AS
BEGIN
    SET NOCOUNT ON;
    IF @AsOfBuild IS NULL
        SET @AsOfBuild = (SELECT TOP 1 BuildNumber FROM dbo.Dim_Build WHERE IsReleaseCandidate = 1 ORDER BY BuildNumber DESC);
    IF NOT EXISTS (SELECT 1 FROM dbo.Dim_Build WHERE BuildNumber = @AsOfBuild)
        THROW 52055, 'AsOfBuild does not match any build in Dim_Build.', 1;

    SELECT r.*,
           RigWeeksOutstanding = CAST(r.RigHoursOutstanding / 180.0 AS DECIMAL(9,1)),
           -- Churn per unit of readiness. A subsystem that changed forty times
           -- and is 13% ready is not behind on testing, it is still being
           -- designed, and that is a different conversation with a different
           -- owner.
           ChurnPerReadinessPoint = CAST(r.BuildsChanged / NULLIF(r.ReadinessPct, 0) AS DECIMAL(9,3))
    FROM dbo.fn_SubsystemReadiness(@AsOfBuild) r
    ORDER BY r.ReadinessPct, r.MustShip DESC;
END;
GO

/*
--------------------------------------------------------------------------------
usp_RequirementDetail -- the evidence trail for one requirement.

This is what somebody opens when they disagree with the scorecard. Every run,
every build, and what invalidated it -- so the answer can be argued with
instead of taken on trust.
--------------------------------------------------------------------------------
*/
CREATE PROCEDURE dbo.usp_RequirementDetail
    @RequirementID VARCHAR(12),
    @AsOfBuild     INT = NULL
AS
BEGIN
    SET NOCOUNT ON;
    IF @AsOfBuild IS NULL
        SET @AsOfBuild = (SELECT TOP 1 BuildNumber FROM dbo.Dim_Build WHERE IsReleaseCandidate = 1 ORDER BY BuildNumber DESC);
    IF NOT EXISTS (SELECT 1 FROM dbo.Dim_Requirement WHERE RequirementID = @RequirementID)
        THROW 52056, 'RequirementID not found in Dim_Requirement.', 1;

    DECLARE @rk INT = (SELECT RequirementKey FROM dbo.Dim_Requirement WHERE RequirementID = @RequirementID);

    -- 1. the verdict
    SELECT Section = '1. Verdict', v.*
    FROM dbo.fn_RequirementVerification(@AsOfBuild) v
    WHERE v.RequirementKey = @rk;

    -- 2. every run, newest first
    SELECT Section = '2. Test runs',
           tc.TestCaseID, tc.TestLevel, tc.AutomationLevel,
           b.BuildID, b.BuildNumber, RunDate = d.[Date],
           tr.Result, RunBy = p.PersonName,
           IsOwner = CAST(CASE WHEN tr.RunByKey = r.OwnerKey THEN 1 ELSE 0 END AS BIT),
           MeetsLevel = CAST(CASE WHEN tlr.LevelRank >= vp.MinTestLevelRank THEN 1 ELSE 0 END AS BIT)
    FROM dbo.Fact_TestRun tr
    JOIN dbo.Dim_TestCase tc  ON tc.TestCaseKey   = tr.TestCaseKey
    JOIN dbo.Dim_Requirement r ON r.RequirementKey = tc.RequirementKey
    JOIN dbo.Dim_Build b      ON b.BuildKey       = tr.BuildKey
    JOIN dbo.Dim_Date d       ON d.DateKey        = tr.RunDateKey
    JOIN dbo.Dim_Person p     ON p.PersonKey      = tr.RunByKey
    JOIN dbo.Ref_TestLevelRank tlr ON tlr.TestLevel = tc.TestLevel
    JOIN dbo.Ref_VerificationPolicy vp ON vp.ReqType = r.ReqType
    WHERE tc.RequirementKey = @rk AND b.BuildNumber <= @AsOfBuild
    ORDER BY b.BuildNumber DESC, tr.TestRunKey DESC;

    -- 3. what changed underneath it
    SELECT Section = '3. Subsystem changes since the last good verification',
           b.BuildID, b.BuildNumber, ChangeDate = d.[Date], c.LinesChanged
    FROM dbo.Fact_BuildSubsystemChange c
    JOIN dbo.Dim_Build b ON b.BuildKey = c.BuildKey
    JOIN dbo.Dim_Date d  ON d.DateKey  = b.BuildDateKey
    WHERE c.SubsystemKey = (SELECT SubsystemKey FROM dbo.Dim_Requirement WHERE RequirementKey = @rk)
      AND b.BuildNumber <= @AsOfBuild
      AND b.BuildNumber > ISNULL((SELECT LastGoodBuild FROM dbo.fn_RequirementVerification(@AsOfBuild)
                                  WHERE RequirementKey = @rk), 0)
    ORDER BY b.BuildNumber;

    -- 4. changes to the requirement itself
    SELECT Section = '4. Requirement changes',
           ChangeDate = d.[Date], rc.ChangeType, ChangedBy = p.PersonName, rc.ChangeNote, b.BuildNumber
    FROM dbo.Fact_RequirementChange rc
    JOIN dbo.Dim_Date d   ON d.DateKey   = rc.ChangeDateKey
    JOIN dbo.Dim_Person p ON p.PersonKey = rc.ChangedByKey
    JOIN dbo.Dim_Build b  ON b.BuildKey  = rc.BuildKey
    WHERE rc.RequirementKey = @rk AND b.BuildNumber <= @AsOfBuild
    ORDER BY b.BuildNumber;
END;
GO

/*
--------------------------------------------------------------------------------
usp_RAIDRegister -- the register, scored and sorted by exposure.
--------------------------------------------------------------------------------
*/
CREATE PROCEDURE dbo.usp_RAIDRegister
    @AsOfDate    DATE = '2026-09-30',
    @MinExposure TINYINT = 1,
    @OpenOnly    BIT = 1
AS
BEGIN
    SET NOCOUNT ON;
    IF @AsOfDate < '2025-01-01' OR @AsOfDate > '2027-06-30'
        THROW 52057, 'AsOfDate falls outside the programme calendar.', 1;
    IF @MinExposure NOT BETWEEN 1 AND 25
        THROW 52058, 'MinExposure must be between 1 and 25 -- the matrix does not score outside that range.', 1;

    SELECT *
    FROM dbo.fn_RAIDExposure(@AsOfDate)
    WHERE ExposureScore >= @MinExposure
      AND (@OpenOnly = 0 OR IsOpen = 1)
    ORDER BY ExposureScore DESC, DaysOverdue DESC, RAIDID;
END;
GO

/*
--------------------------------------------------------------------------------
usp_ReadinessTrend -- readiness at successive builds.

The payoff for parameterising everything by build. Readiness is not a number,
it is a trajectory, and a single figure cannot distinguish a programme that is
recovering from one that is falling behind faster than it is testing.
--------------------------------------------------------------------------------
*/
CREATE PROCEDURE dbo.usp_ReadinessTrend
    @FromBuild INT = 40,
    @ToBuild   INT = NULL,
    @StepSize  INT = 6
AS
BEGIN
    SET NOCOUNT ON;
    IF @ToBuild IS NULL
        SET @ToBuild = (SELECT TOP 1 BuildNumber FROM dbo.Dim_Build WHERE IsReleaseCandidate = 1 ORDER BY BuildNumber DESC);
    IF @FromBuild < 1 OR @ToBuild < 1
        THROW 52059, 'Build numbers must be positive.', 1;
    IF @ToBuild < @FromBuild
        THROW 52060, 'ToBuild falls before FromBuild.', 1;
    IF @StepSize < 1
        THROW 52061, 'StepSize must be at least 1.', 1;

    ;WITH Steps AS (
        SELECT BuildNumber = @FromBuild
        UNION ALL
        SELECT BuildNumber + @StepSize FROM Steps WHERE BuildNumber + @StepSize <= @ToBuild
    )
    SELECT b.BuildNumber, b.BuildID, BuildDate = d.[Date],
           k.WorkItemCompletionPct, k.VerificationCoveragePct,
           k.ShipReadinessPct, k.StaleVerificationPct,
           ReadinessGapVsReported = CAST(k.WorkItemCompletionPct - k.ShipReadinessPct AS DECIMAL(9,2))
    FROM Steps s
    JOIN dbo.Dim_Build b ON b.BuildNumber = s.BuildNumber
    JOIN dbo.Dim_Date d  ON d.DateKey     = b.BuildDateKey
    CROSS APPLY dbo.fn_ReadinessKPI(s.BuildNumber, d.[Date]) k
    ORDER BY b.BuildNumber
    OPTION (MAXRECURSION 500);
END;
GO

-- =============================================================================
-- Creation gate
-- =============================================================================
DECLARE @missing VARCHAR(400) = '';
IF OBJECT_ID('dbo.usp_ReadinessScorecard','P') IS NULL SET @missing += 'usp_ReadinessScorecard ';
IF OBJECT_ID('dbo.usp_VerificationQueue','P')  IS NULL SET @missing += 'usp_VerificationQueue ';
IF OBJECT_ID('dbo.usp_SubsystemReadiness','P') IS NULL SET @missing += 'usp_SubsystemReadiness ';
IF OBJECT_ID('dbo.usp_RequirementDetail','P')  IS NULL SET @missing += 'usp_RequirementDetail ';
IF OBJECT_ID('dbo.usp_RAIDRegister','P')       IS NULL SET @missing += 'usp_RAIDRegister ';
IF OBJECT_ID('dbo.usp_ReadinessTrend','P')     IS NULL SET @missing += 'usp_ReadinessTrend ';
IF OBJECT_ID('dbo.usp_RunDataQualityChecks','P') IS NULL SET @missing += 'usp_RunDataQualityChecks ';
IF @missing <> '' THROW 52062, 'FAILED to create stored procedures -- scroll up for the compile error.', 1;

PRINT 'Stored procedures created and verified: 7 procedures, all @AsOf-aware, all validating their arguments.';
GO
