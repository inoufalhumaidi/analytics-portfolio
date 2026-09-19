/*
================================================================================
Project 5 — Meridian UAV Services: Predictive Maintenance
Script:  06_stored_procedures.sql
Purpose: The reusable interface. Eight procedures, every one as-of aware.

WHY PROCEDURES AND NOT JUST VIEWS

    Excel, Power BI, Streamlit and the Python model all need the same figures.
    If each writes its own query, each owns its own definition of 'overdue', and
    the first time they disagree nobody can tell which one is wrong -- they are
    four opinions, not one answer checked four ways.

    So the definitions live in 03 and 05, and everything downstream calls these.

DEFAULTS COME FROM Ref_Reporting, NOT FROM THE SIGNATURE

    Every procedure takes @AsOfDate = NULL and resolves NULL from the reference
    table. A default baked into the signature is invisible to the analyst
    reading the output and has to be changed in eight places; a default in a
    table is one row, with the reason for it written beside it.

    NULL specifically, and not 'the latest date in the data'. Anchoring on the
    data means the published answer moves the moment someone loads one more
    sortie, and a figure in a report silently stops matching the system that
    produced it.

DATA DISCLOSURE
    Meridian UAV Services is fictional and all data is synthetic. No
    confidential data is used and no claim is made about any production system.
================================================================================
*/

USE MeridianUAV;
GO
SET NOCOUNT ON;
GO

-- =============================================================================
-- fn_ReportingDate / fn_ReportingDecimal -- resolve a setting, or fail loudly.
--
-- A missing setting must NOT silently fall back to GETDATE(). That would make
-- every figure in the project depend on when it was run, and the failure would
-- be invisible because the output would still look like a date.
-- =============================================================================
IF OBJECT_ID('dbo.fn_ReportingDate', 'FN') IS NOT NULL DROP FUNCTION dbo.fn_ReportingDate;
GO
CREATE FUNCTION dbo.fn_ReportingDate (@Supplied DATE)
RETURNS DATE
AS
BEGIN
    IF @Supplied IS NOT NULL RETURN @Supplied;
    RETURN (SELECT CAST(SettingValue AS DATE) FROM dbo.Ref_Reporting
            WHERE SettingName = 'ReportingAsOfDate');
END;
GO

IF OBJECT_ID('dbo.fn_ReportingDecimal', 'FN') IS NOT NULL DROP FUNCTION dbo.fn_ReportingDecimal;
GO
CREATE FUNCTION dbo.fn_ReportingDecimal (@SettingName VARCHAR(40), @Supplied DECIMAL(9,2))
RETURNS DECIMAL(9,2)
AS
BEGIN
    IF @Supplied IS NOT NULL RETURN @Supplied;
    RETURN (SELECT CAST(SettingValue AS DECIMAL(9,2)) FROM dbo.Ref_Reporting
            WHERE SettingName = @SettingName);
END;
GO

-- =============================================================================
-- usp_FleetScorecard -- the board.
-- =============================================================================
IF OBJECT_ID('dbo.usp_FleetScorecard', 'P') IS NOT NULL DROP PROCEDURE dbo.usp_FleetScorecard;
GO
CREATE PROCEDURE dbo.usp_FleetScorecard @AsOfDate DATE = NULL
AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @d DATE = dbo.fn_ReportingDate(@AsOfDate);
    IF @d IS NULL THROW 53050, 'No reporting date supplied and Ref_Reporting has no ReportingAsOfDate row.', 1;

    DECLARE @t DECIMAL(9,2) = dbo.fn_ReportingDecimal('AlertVibThreshold', NULL);

    -- Both boards, in one result set. They were separate functions because the
    -- prediction metrics need a threshold as well as a date; they are one
    -- SCORECARD because a reader deciding what to do about this fleet needs to
    -- see the compliance figures and the prediction figures together.
    -- The sort key is SELECTED, not written into the ORDER BY. A CASE
    -- expression cannot order a UNION directly -- the items must appear in the
    -- select list -- so the board's reading order becomes a column.
    SELECT AsOfDate, MetricName, MetricValue, Numerator, Denominator,
           TargetValue, WarningValue, Direction, Unit, RAGStatus, [Description]
    FROM (
        SELECT AsOfDate = @d, MetricName, MetricValue, Numerator, Denominator,
               TargetValue, WarningValue, Direction, Unit, RAGStatus, [Description],
               SortOrder = CASE MetricName
                   WHEN 'StressCompliancePct'      THEN 1
                   WHEN 'HourCompliancePct'        THEN 2
                   WHEN 'OverdueFlightCritical'    THEN 3
                   WHEN 'UnscheduledRatePct'       THEN 4
                   WHEN 'AirframeAvailabilityPct'  THEN 5 ELSE 9 END
        FROM dbo.fn_FleetKPI(@d)
        UNION ALL
        SELECT @d, MetricName, MetricValue, Numerator, Denominator,
               TargetValue, WarningValue, Direction, Unit, RAGStatus, [Description],
               CASE MetricName
                   WHEN 'ActionableLeadTimePct'    THEN 6
                   WHEN 'PredictionPrecisionPct'   THEN 7
                   WHEN 'PredictionRecallPct'      THEN 8 ELSE 9 END
        FROM dbo.fn_PredictionKPI(@d, CAST(@t AS DECIMAL(6,3)))
    ) z
    ORDER BY z.SortOrder, z.MetricName;
END;
GO

-- =============================================================================
-- usp_BaseScorecard -- the same board, cut where the variation actually is.
-- =============================================================================
IF OBJECT_ID('dbo.usp_BaseScorecard', 'P') IS NOT NULL DROP PROCEDURE dbo.usp_BaseScorecard;
GO
CREATE PROCEDURE dbo.usp_BaseScorecard @AsOfDate DATE = NULL
AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @d DATE = dbo.fn_ReportingDate(@AsOfDate);
    IF @d IS NULL THROW 53050, 'No reporting date supplied and Ref_Reporting has no ReportingAsOfDate row.', 1;
    SELECT * FROM dbo.fn_BaseScorecard(@d) ORDER BY StressCompliancePct;
END;
GO

-- =============================================================================
-- usp_AirframeDetail -- one aircraft, every position.
--
-- The screen an engineer opens when the queue tells them to ground something,
-- and the one that makes the finding concrete: an airframe reading 60% on the
-- hour meter and 140% on stress, position by position.
-- =============================================================================
IF OBJECT_ID('dbo.usp_AirframeDetail', 'P') IS NOT NULL DROP PROCEDURE dbo.usp_AirframeDetail;
GO
CREATE PROCEDURE dbo.usp_AirframeDetail
    @TailNumber VARCHAR(12),
    @AsOfDate   DATE = NULL
AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @d DATE = dbo.fn_ReportingDate(@AsOfDate);
    IF @d IS NULL THROW 53050, 'No reporting date supplied and Ref_Reporting has no ReportingAsOfDate row.', 1;

    IF NOT EXISTS (SELECT 1 FROM dbo.Dim_Airframe WHERE TailNumber = @TailNumber)
        THROW 53051, 'No such tail number. A typo must not return an empty result set that reads as a clean aircraft.', 1;

    SELECT AsOfDate = @d, a.TailNumber, a.ModelCode, a.AirframeStatus,
           b.BaseCode, b.BaseName,
           s.Sorties, s.FlightHours, s.StressHours, s.StressRatio,
           s.PctStressFromHours, s.PctStressFromCycles, s.PctStressFromPayload,
           s.DominantProfileName
    FROM dbo.Dim_Airframe a
    JOIN dbo.Dim_Base b ON b.BaseKey = a.BaseKey
    JOIN dbo.fn_AirframeStress(@d) s ON s.AirframeKey = a.AirframeKey
    WHERE a.TailNumber = @TailNumber;

    SELECT w.ComponentCode, w.ComponentName, w.PositionNo, w.ComponentSerial, w.Criticality,
           w.InstalledDate, w.Sorties, w.FlightHours, w.StressHours,
           w.IntervalFlightHours, w.IntervalStressHours,
           w.PctOfHourInterval, w.PctOfStressInterval,
           w.IsOverdueByHours, w.IsOverdueByStress, w.IsHiddenOverdue,
           w.RemainingStressHours, w.ProjectedFlightHoursToStressLimit, w.LeadTimeFlightHours
    FROM dbo.fn_ComponentWear(@d) w
    WHERE w.TailNumber = @TailNumber
    ORDER BY w.PctOfStressInterval DESC, w.ComponentCode, w.PositionNo;
END;
GO

-- =============================================================================
-- usp_ActionQueue -- the operational output.
-- =============================================================================
IF OBJECT_ID('dbo.usp_ActionQueue', 'P') IS NOT NULL DROP PROCEDURE dbo.usp_ActionQueue;
GO
CREATE PROCEDURE dbo.usp_ActionQueue
    @AsOfDate           DATE = NULL,
    @HangarHoursPerWeek DECIMAL(9,2) = NULL,
    @BaseCode           VARCHAR(10) = NULL,   -- NULL = the whole fleet
    @ThisWeekOnly       BIT = 0
AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @d DATE = dbo.fn_ReportingDate(@AsOfDate);
    DECLARE @h DECIMAL(9,2) = dbo.fn_ReportingDecimal('HangarHoursPerWeek', @HangarHoursPerWeek);

    -- A NULL capacity makes every IsThisWeek comparison UNKNOWN, so the week's
    -- plan comes back empty and reads as 'nothing to do'. Only one of the eight
    -- procedures guarded its resolved setting; a missing reference row must
    -- raise, not quietly empty the plan.
    IF @d IS NULL THROW 53050, 'No reporting date supplied and Ref_Reporting has no ReportingAsOfDate row.', 1;
    IF @h IS NULL OR @h <= 0
        THROW 53057, 'Hangar capacity resolved to NULL or zero. The week plan would be empty and would read as nothing to do.', 1;

    IF @BaseCode IS NOT NULL AND NOT EXISTS (SELECT 1 FROM dbo.Dim_Base WHERE BaseCode = @BaseCode)
        THROW 53052, 'No such base code.', 1;

    /*
    THE CAPACITY CUT IS APPLIED TO THE WHOLE FLEET AND THEN FILTERED, NOT THE
    OTHER WAY AROUND.

    Hangar capacity is a fleet-level constraint. Ranking one base in isolation
    and cutting at the fleet's weekly hours would tell every base it can do its
    top 26 jobs this week -- four bases each planning against the same 120
    hours. The cut has to happen before the filter or it means nothing.
    */
    SELECT *
    FROM dbo.fn_ActionQueue(@d, @h)
    WHERE (@BaseCode IS NULL OR BaseCode = @BaseCode)
      AND (@ThisWeekOnly = 0 OR IsThisWeek = 1)
    ORDER BY PriorityRank;
END;
GO

-- =============================================================================
-- usp_AlertPerformance -- accuracy AND usefulness, side by side.
-- =============================================================================
IF OBJECT_ID('dbo.usp_AlertPerformance', 'P') IS NOT NULL DROP PROCEDURE dbo.usp_AlertPerformance;
GO
CREATE PROCEDURE dbo.usp_AlertPerformance
    @AsOfDate     DATE = NULL,
    @VibThreshold DECIMAL(9,2) = NULL
AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @d DATE = dbo.fn_ReportingDate(@AsOfDate);
    DECLARE @t DECIMAL(9,2) = dbo.fn_ReportingDecimal('AlertVibThreshold', @VibThreshold);
    IF @d IS NULL THROW 53050, 'No reporting date supplied and Ref_Reporting has no ReportingAsOfDate row.', 1;
    IF @t IS NULL OR @t <= 0
        THROW 53058, 'Alert threshold resolved to NULL. Every alert comparison would be UNKNOWN and the rule would appear to raise no alerts at all.', 1;

    SELECT Scope = 'Fleet', ComponentCode = 'ALL', f.*
    FROM dbo.fn_AlertSummaryFleet(@d, CAST(@t AS DECIMAL(6,3))) f;

    SELECT Scope = 'Component', s.*
    FROM dbo.fn_AlertSummary(@d, CAST(@t AS DECIMAL(6,3))) s
    ORDER BY s.ComponentCode;
END;
GO

-- =============================================================================
-- usp_ThresholdSweep -- the trade-off, as a table.
--
-- The single most useful output in the project, because it shows that the
-- conflict is structural rather than a tuning mistake. Precision rises with the
-- threshold and lead time falls with it, monotonically, and there is no setting
-- where both are acceptable for the motor controller.
--
-- A programme shown only its chosen operating point will conclude it needs
-- better tuning. Shown the curve, it concludes it needs a different supply
-- chain -- which is the actual recommendation.
-- =============================================================================
IF OBJECT_ID('dbo.usp_ThresholdSweep', 'P') IS NOT NULL DROP PROCEDURE dbo.usp_ThresholdSweep;
GO
CREATE PROCEDURE dbo.usp_ThresholdSweep
    @AsOfDate DATE = NULL,
    @From     DECIMAL(6,3) = 1.200,
    @To       DECIMAL(6,3) = 4.000,
    @Step     DECIMAL(6,3) = 0.200
AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @d DATE = dbo.fn_ReportingDate(@AsOfDate);
    IF @d IS NULL THROW 53050, 'No reporting date supplied and Ref_Reporting has no ReportingAsOfDate row.', 1;

    IF @Step <= 0 THROW 53053, 'Step must be positive; a zero or negative step would loop forever.', 1;
    IF @To < @From THROW 53054, 'To must not be below From.', 1;

    DECLARE @Steps TABLE (v DECIMAL(6,3) PRIMARY KEY);
    DECLARE @v DECIMAL(6,3) = @From;
    WHILE @v <= @To
    BEGIN
        INSERT INTO @Steps (v) VALUES (@v);
        SET @v = @v + @Step;
    END

    SELECT s.VibThreshold, s.ComponentCode, s.Population, s.Failures, s.AlertsRaised,
           s.TP, s.FP, s.FN, s.ActionableTP, s.LateTP,
           s.PrecisionPct, s.RecallPct, s.ActionableLeadTimePct, s.PreventableFailurePct,
           s.MedianLeadTimeFlightHours, s.LeadTimeFlightHours,
           -- the single column that makes the point: is the median warning
           -- longer than the time it takes to get the part?
           MedianLeadTimeCoversSupply = CAST(CASE WHEN s.MedianLeadTimeFlightHours >= s.LeadTimeFlightHours
                                                  THEN 1 ELSE 0 END AS BIT)
    FROM @Steps t
    CROSS APPLY dbo.fn_AlertSummary(@d, t.v) s
    ORDER BY s.ComponentCode, s.VibThreshold;
END;
GO

-- =============================================================================
-- usp_ComplianceTrend -- the two lines, month by month.
--
-- A trend, not a snapshot, because the gap between the two measures is not
-- noise: it widens as the fleet accumulates hours on the wrong unit. A single
-- month cannot show that, and a programme that sees only a snapshot will treat
-- the gap as a one-off to be cleared rather than a policy that keeps producing
-- it.
-- =============================================================================
IF OBJECT_ID('dbo.usp_ComplianceTrend', 'P') IS NOT NULL DROP PROCEDURE dbo.usp_ComplianceTrend;
GO
CREATE PROCEDURE dbo.usp_ComplianceTrend
    @AsOfDate DATE = NULL,
    @Months   INT  = 18
AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @d DATE = dbo.fn_ReportingDate(@AsOfDate);
    IF @d IS NULL THROW 53050, 'No reporting date supplied and Ref_Reporting has no ReportingAsOfDate row.', 1;
    IF @Months < 2 THROW 53055, 'A trend needs at least two points.', 1;

    ;WITH MonthEnds AS (
        SELECT TOP (@Months) d.MonthEndDate
        FROM dbo.Dim_Date d
        WHERE d.IsMonthEnd = 1 AND d.MonthEndDate <= @d
        GROUP BY d.MonthEndDate
        ORDER BY d.MonthEndDate DESC
    )
    SELECT m.MonthEndDate,
           k.MetricName,
           k.MetricValue
    FROM MonthEnds m
    CROSS APPLY dbo.fn_FleetKPI(m.MonthEndDate) k
    WHERE k.MetricName IN ('StressCompliancePct','HourCompliancePct','UnscheduledRatePct')
    ORDER BY m.MonthEndDate, k.MetricName;
END;
GO

-- =============================================================================
-- usp_DataQualitySummary -- every check, including the clean ones.
-- =============================================================================
IF OBJECT_ID('dbo.usp_DataQualitySummary', 'P') IS NOT NULL DROP PROCEDURE dbo.usp_DataQualitySummary;
GO
CREATE PROCEDURE dbo.usp_DataQualitySummary @IncludeDetail BIT = 0
AS
BEGIN
    SET NOCOUNT ON;

    IF OBJECT_ID('dbo.DQ_Findings', 'U') IS NULL
        THROW 53056, 'DQ_Findings does not exist. Run 04_data_quality_checks.sql before asking for its summary -- an empty summary would read as a clean dataset.', 1;

    SELECT CheckCode, CheckName, Severity,
           Findings = COUNT(*),
           AffectingCurrentWear = SUM(CAST(AffectsCurrentWear AS INT))
    FROM dbo.DQ_Findings
    GROUP BY CheckCode, CheckName, Severity
    ORDER BY CheckCode;

    IF @IncludeDetail = 1
        SELECT CheckCode, CheckName, Severity, EntityType, EntityRef, [Detail], AffectsCurrentWear
        FROM dbo.DQ_Findings
        ORDER BY CheckCode, FindingKey;
END;
GO

-- =============================================================================
-- usp_RecommendedThresholds -- the recommendation the second finding produces.
--
-- One threshold per component, and the fleet effect of adopting them.
--
-- The fleet rollup SUMS THE COUNTS at each component's own threshold. It does
-- not average the two percentages: an average weights a population of 130 the
-- same as one of 400 and produces a third number describing neither.
-- =============================================================================
IF OBJECT_ID('dbo.usp_RecommendedThresholds', 'P') IS NOT NULL DROP PROCEDURE dbo.usp_RecommendedThresholds;
GO
CREATE PROCEDURE dbo.usp_RecommendedThresholds
    @AsOfDate        DATE = NULL,
    @MinPrecisionPct DECIMAL(6,2) = 75.00
AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @d DATE = dbo.fn_ReportingDate(@AsOfDate);
    IF @d IS NULL THROW 53050, 'No reporting date supplied and Ref_Reporting has no ReportingAsOfDate row.', 1;
    IF @MinPrecisionPct IS NULL OR @MinPrecisionPct < 0 OR @MinPrecisionPct > 100
        THROW 53059, 'The precision floor must be a percentage. It is the most arguable number in the recommendation and it must be stated.', 1;

    SELECT * FROM dbo.fn_ThresholdRecommendation(@d, @MinPrecisionPct)
    ORDER BY ActionablePointsGained DESC;

    DECLARE @cur DECIMAL(6,3) = (SELECT CAST(SettingValue AS DECIMAL(6,3))
                                 FROM dbo.Ref_Reporting WHERE SettingName = 'AlertVibThreshold');

    SELECT Scope = 'Today: one threshold for the whole fleet',
           Threshold = CAST(@cur AS VARCHAR(10)),
           TP = SUM(TP), ActionableTP = SUM(ActionableTP), LateTP = SUM(LateTP),
           PrecisionPct = CAST(100.0 * SUM(TP) / NULLIF(SUM(AlertsRaised), 0) AS DECIMAL(6,2)),
           ActionablePct = CAST(100.0 * SUM(ActionableTP) / NULLIF(SUM(TP), 0) AS DECIMAL(6,2))
    FROM dbo.fn_AlertSummary(@d, @cur)
    UNION ALL
    SELECT 'Recommended: one threshold per component', 'per component',
           SUM(TP), SUM(ActionableTP), SUM(LateTP),
           CAST(100.0 * SUM(TP) / NULLIF(SUM(AlertsRaised), 0) AS DECIMAL(6,2)),
           CAST(100.0 * SUM(ActionableTP) / NULLIF(SUM(TP), 0) AS DECIMAL(6,2))
    FROM dbo.fn_ThresholdRecommendation(@d, @MinPrecisionPct);
END;
GO

PRINT 'Stored procedures created: 9 procedures, 2 settings resolvers.';
GO

-- =============================================================================
-- Creation gate -- every procedure is EXECUTED, not merely compiled.
--
-- A procedure that compiles proves its syntax. It proves nothing about whether
-- it runs: a bad column reference inside a deferred-resolution body, a type
-- mismatch on a CAST, an infinite loop. All eight are executed here.
-- =============================================================================
DECLARE @rc INT;

EXEC dbo.usp_FleetScorecard;
EXEC dbo.usp_BaseScorecard;
EXEC dbo.usp_ActionQueue;
EXEC dbo.usp_AlertPerformance;
EXEC dbo.usp_ThresholdSweep @From = 2.000, @To = 2.400, @Step = 0.200;
EXEC dbo.usp_ComplianceTrend @Months = 3;
EXEC dbo.usp_DataQualitySummary;
EXEC dbo.usp_AirframeDetail @TailNumber = 'MU-001';
EXEC dbo.usp_RecommendedThresholds;

PRINT 'Procedure gate passed: all nine executed.';
GO
