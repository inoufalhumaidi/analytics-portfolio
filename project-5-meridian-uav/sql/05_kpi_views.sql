/*
================================================================================
Project 5 — Meridian UAV Services: Predictive Maintenance
Script:  05_kpi_views.sql
Purpose: The scorecard, the alert evaluation, and the action queue.

WHAT IS HERE

  fn_FleetKPI          the five fleet-condition metrics, as-of a date
  fn_PredictionKPI     the three prediction metrics, at the pinned threshold
  fn_BaseScorecard     the same cut by base, because the fleet average hides it
  fn_AlertEvaluation   what a condition-monitoring alert rule is actually worth
  fn_AlertSummary      that evaluation rolled up, per component and fleet-wide
  fn_ThresholdRecommendation  one threshold per component, not one per fleet
  fn_ActionQueue       what to do on Monday, bounded by hangar capacity

  Between them, fn_FleetKPI and fn_PredictionKPI evaluate ALL EIGHT rows of
  Ref_FleetTargets. The first version returned five and left three targets in
  the table with nothing comparing anything to them -- including
  ActionableLeadTimePct, which is the target the second finding is about. A
  target nothing evaluates is a target nobody can fail.

THE SECOND FINDING LIVES IN fn_AlertEvaluation

    A failure prediction is not worth its accuracy. It is worth its LEAD TIME.

    A model that flags a rotor bearing four flight hours before it fails is
    correct, and the aircraft is grounded either way, because the assembly takes
    eighteen flight hours to obtain and fit. Precision and recall cannot tell
    those two alerts apart -- both are true positives -- so a scorecard built on
    them will report a successful programme that has changed nothing.

    So this file reports precision and recall (because people ask for them) and
    then reports ActionableLeadTimePct alongside, which is the share of caught
    failures caught EARLY ENOUGH TO ACT. Dim_ComponentType.LeadTimeFlightHours
    is what 'early enough' means, per component, and it lives in a table so a
    planner can change it when the supply chain changes.

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
-- fn_FleetKPI -- the headline board, tall rather than wide.
--
-- One row per metric rather than one row of many columns, so that Excel and
-- Power BI can both bind to Ref_FleetTargets by name. A wide row forces every
-- consumer to hard-code the target beside each column, and hard-coded targets
-- are how a dashboard comes to show green against a threshold nobody agreed.
-- =============================================================================
IF OBJECT_ID('dbo.fn_FleetKPI', 'IF') IS NOT NULL DROP FUNCTION dbo.fn_FleetKPI;
GO
CREATE FUNCTION dbo.fn_FleetKPI (@AsOfDate DATE)
RETURNS TABLE
AS
RETURN
/*
POPULATION: the fleet AS IT WAS on @AsOfDate.

Not AirframeStatus = 'Active', which describes the fleet today. Three aircraft
retired inside the reporting window; filtering on their current status removed
them from every historical figure and froze the availability denominator at 28
across an eighteen-month trend.
*/
WITH Fitted AS (
    SELECT * FROM dbo.fn_ComponentWear(@AsOfDate) WHERE IsInFleetAsOf = 1
),
FleetAsOf AS (
    SELECT AirframeKey FROM dbo.fn_AirframeStress(@AsOfDate) WHERE IsInFleetAsOf = 1
),
/*
Events in the TRAILING 12 MONTHS ending @AsOfDate.

Not since-inception. An unscheduled rate computed over the whole history of the
fleet includes the period before any of the current aircraft reached their first
interval, so it improves every month for reasons that have nothing to do with
maintenance. A trailing window answers 'how are we doing now'.
*/
Events AS (
    -- joined to FleetAsOf, not to Dim_Airframe with no filter. The first
    -- version joined the dimension and filtered nothing, so this rate was
    -- computed over a different population from every metric beside it.
    SELECT ev.EventType, ev.DowntimeHours
    FROM dbo.Fact_MaintenanceEvent ev
    JOIN dbo.Dim_Date d   ON d.DateKey = ev.EventDateKey
    JOIN FleetAsOf f      ON f.AirframeKey = ev.AirframeKey
    WHERE d.[Date] <= @AsOfDate
      AND d.[Date] >  DATEADD(YEAR, -1, @AsOfDate)
      AND ev.EventType IN ('Scheduled','Unscheduled')
),
/*
Availability. An airframe is unavailable on @AsOfDate if it is inside the
downtime window of a maintenance event -- the event date plus its downtime
converted to working days at 8 hours a day.
*/
Grounded AS (
    SELECT DISTINCT ev.AirframeKey
    FROM dbo.Fact_MaintenanceEvent ev
    JOIN dbo.Dim_Date d ON d.DateKey = ev.EventDateKey
    JOIN FleetAsOf f    ON f.AirframeKey = ev.AirframeKey
    WHERE d.[Date] <= @AsOfDate
      AND DATEADD(DAY, CEILING(ev.DowntimeHours / 8.0), d.[Date]) > @AsOfDate
),
Agg AS (
    SELECT
        FittedCount      = (SELECT COUNT(*) FROM Fitted),
        HourCompliant    = (SELECT COUNT(*) FROM Fitted WHERE IsOverdueByHours = 0),
        StressCompliant  = (SELECT COUNT(*) FROM Fitted WHERE IsOverdueByStress = 0),
        HiddenOverdue    = (SELECT COUNT(*) FROM Fitted WHERE IsHiddenOverdue = 1),
        OverdueFltCrit   = (SELECT COUNT(*) FROM Fitted WHERE IsOverdueByStress = 1 AND Criticality = 'FlightCritical'),
        EventsTotal      = (SELECT COUNT(*) FROM Events),
        EventsUnsched    = (SELECT COUNT(*) FROM Events WHERE EventType = 'Unscheduled'),
        DowntimeUnsched  = (SELECT ISNULL(SUM(DowntimeHours), 0) FROM Events WHERE EventType = 'Unscheduled'),
        DowntimeTotal    = (SELECT ISNULL(SUM(DowntimeHours), 0) FROM Events),
        -- the fleet ON @AsOfDate, not the fleet today. This denominator read
        -- 28.00 at every date in the trend while the numerator moved, which is
        -- exactly why a test asserting 'the metric changes' passed against it.
        ActiveAirframes  = (SELECT COUNT(*) FROM FleetAsOf),
        GroundedNow      = (SELECT COUNT(*) FROM Grounded)
),
Metrics AS (
    SELECT MetricName = 'StressCompliancePct',
           MetricValue = CAST(100.0 * StressCompliant / NULLIF(FittedCount, 0) AS DECIMAL(10,2)),
           Numerator = CAST(StressCompliant AS DECIMAL(12,2)), Denominator = CAST(FittedCount AS DECIMAL(12,2)) FROM Agg
    UNION ALL
    SELECT 'HourCompliancePct',
           CAST(100.0 * HourCompliant / NULLIF(FittedCount, 0) AS DECIMAL(10,2)),
           CAST(HourCompliant AS DECIMAL(12,2)), CAST(FittedCount AS DECIMAL(12,2)) FROM Agg
    UNION ALL
    SELECT 'UnscheduledRatePct',
           CAST(100.0 * EventsUnsched / NULLIF(EventsTotal, 0) AS DECIMAL(10,2)),
           CAST(EventsUnsched AS DECIMAL(12,2)), CAST(EventsTotal AS DECIMAL(12,2)) FROM Agg
    UNION ALL
    SELECT 'AirframeAvailabilityPct',
           CAST(100.0 * (ActiveAirframes - GroundedNow) / NULLIF(ActiveAirframes, 0) AS DECIMAL(10,2)),
           CAST(ActiveAirframes - GroundedNow AS DECIMAL(12,2)), CAST(ActiveAirframes AS DECIMAL(12,2)) FROM Agg
    UNION ALL
    SELECT 'OverdueFlightCritical',
           CAST(OverdueFltCrit AS DECIMAL(10,2)),
           CAST(OverdueFltCrit AS DECIMAL(12,2)), NULL FROM Agg
)
SELECT
    AsOfDate = @AsOfDate,
    m.MetricName,
    m.MetricValue,
    m.Numerator,
    m.Denominator,
    t.TargetValue,
    t.WarningValue,
    t.Direction,
    t.Unit,
    t.[Description],
    /*
    RAG. The comparison DEPENDS ON DIRECTION, which is why Direction is a column
    in Ref_FleetTargets rather than an assumption in this expression. Half these
    metrics are better when higher and half when lower, and a single hard-coded
    comparison would silently paint one half of the board the wrong colour.
    */
    RAGStatus = CASE
        WHEN m.MetricValue IS NULL THEN 'Grey'
        WHEN t.Direction = 'HigherBetter' THEN
            CASE WHEN m.MetricValue >= t.TargetValue  THEN 'Green'
                 WHEN m.MetricValue >= t.WarningValue THEN 'Amber'
                 ELSE 'Red' END
        ELSE
            CASE WHEN m.MetricValue <= t.TargetValue  THEN 'Green'
                 WHEN m.MetricValue <= t.WarningValue THEN 'Amber'
                 ELSE 'Red' END
    END
FROM Metrics m
JOIN dbo.Ref_FleetTargets t ON t.MetricName = m.MetricName;
GO

-- =============================================================================
-- fn_BaseScorecard -- the cut that matters.
--
-- The fleet figure is an average over four bases doing different work, and it
-- is the wrong number to manage by: it is dragged up by the benign base and
-- down by the hard one, and describes neither. Project 3 shipped a portfolio
-- figure that was fine while a cohort inside it was failing; this is that
-- lesson made structural.
-- =============================================================================
IF OBJECT_ID('dbo.fn_BaseScorecard', 'IF') IS NOT NULL DROP FUNCTION dbo.fn_BaseScorecard;
GO
CREATE FUNCTION dbo.fn_BaseScorecard (@AsOfDate DATE)
RETURNS TABLE
AS
RETURN
/*
ONE POPULATION FOR THE WHOLE ROW.

The first version filtered the component and airframe columns on
AirframeStatus = 'Active' and left the maintenance-event columns unfiltered, so
a single row mixed two fleets: Coastal's unscheduled rate included nine events
from an aircraft its compliance columns excluded. A row whose columns come from
different populations cannot be read across, and nothing about it looks wrong.
*/
WITH FleetAsOf AS (
    SELECT AirframeKey, BaseKey, BaseCode, FlightHours, StressHours
    FROM dbo.fn_AirframeStress(@AsOfDate) WHERE IsInFleetAsOf = 1
),
Fitted AS (
    SELECT * FROM dbo.fn_ComponentWear(@AsOfDate) WHERE IsInFleetAsOf = 1
),
Stress AS (
    SELECT BaseKey, BaseCode,
           Airframes   = COUNT(*),
           FlightHours = SUM(FlightHours),
           StressHours = SUM(StressHours)
    FROM FleetAsOf
    GROUP BY BaseKey, BaseCode
),
Comp AS (
    SELECT BaseKey,
           FittedCount     = COUNT(*),
           HourCompliant   = SUM(CASE WHEN IsOverdueByHours  = 0 THEN 1 ELSE 0 END),
           StressCompliant = SUM(CASE WHEN IsOverdueByStress = 0 THEN 1 ELSE 0 END),
           HiddenOverdue   = SUM(CASE WHEN IsHiddenOverdue   = 1 THEN 1 ELSE 0 END),
           HiddenFltCrit   = SUM(CASE WHEN IsHiddenOverdue   = 1 AND Criticality = 'FlightCritical' THEN 1 ELSE 0 END)
    FROM Fitted GROUP BY BaseKey
),
Unsched AS (
    SELECT f.BaseKey,
           EventsTotal   = COUNT(*),
           EventsUnsched = SUM(CASE WHEN ev.EventType = 'Unscheduled' THEN 1 ELSE 0 END),
           DowntimeHours = SUM(ev.DowntimeHours)
    FROM dbo.Fact_MaintenanceEvent ev
    JOIN FleetAsOf f    ON f.AirframeKey = ev.AirframeKey
    JOIN dbo.Dim_Date d ON d.DateKey = ev.EventDateKey
    WHERE d.[Date] <= @AsOfDate AND d.[Date] > DATEADD(YEAR, -1, @AsOfDate)
      AND ev.EventType IN ('Scheduled','Unscheduled')
    GROUP BY f.BaseKey
)
SELECT
    AsOfDate = @AsOfDate,
    b.BaseKey, b.BaseCode, b.BaseName, b.Region, b.HangarBays,
    Airframes   = ISNULL(s.Airframes, 0),
    FlightHours = CAST(ISNULL(s.FlightHours, 0) AS DECIMAL(12,2)),
    StressHours = CAST(ISNULL(s.StressHours, 0) AS DECIMAL(12,2)),
    -- the single number that explains the rest of the row
    StressRatio = CAST(s.StressHours / NULLIF(s.FlightHours, 0) AS DECIMAL(9,4)),
    FittedComponents  = ISNULL(c.FittedCount, 0),
    HourCompliancePct   = CAST(100.0 * c.HourCompliant   / NULLIF(c.FittedCount, 0) AS DECIMAL(6,2)),
    StressCompliancePct = CAST(100.0 * c.StressCompliant / NULLIF(c.FittedCount, 0) AS DECIMAL(6,2)),
    -- the gap between what the maintenance system reports and what is true
    ComplianceGapPoints = CAST(100.0 * (c.HourCompliant - c.StressCompliant) / NULLIF(c.FittedCount, 0) AS DECIMAL(6,2)),
    HiddenOverdue       = ISNULL(c.HiddenOverdue, 0),
    HiddenOverdueFlightCritical = ISNULL(c.HiddenFltCrit, 0),
    UnscheduledRatePct  = CAST(100.0 * u.EventsUnsched / NULLIF(u.EventsTotal, 0) AS DECIMAL(6,2)),
    MaintenanceEvents   = ISNULL(u.EventsTotal, 0),
    DowntimeHours       = CAST(ISNULL(u.DowntimeHours, 0) AS DECIMAL(10,1))
FROM dbo.Dim_Base b
LEFT JOIN Stress s  ON s.BaseKey = b.BaseKey
LEFT JOIN Comp c    ON c.BaseKey = b.BaseKey
LEFT JOIN Unsched u ON u.BaseKey = b.BaseKey;
GO

/*
==============================================================================
 AlertFrontier -- the one materialised table in this project, and why.

 THE PROBLEM

 Evaluating an alarm at one threshold takes about 1.7 seconds: vw_SensorFeatures
 computes rolling windows and per-component baselines over 83,520 readings.
 That is fine once. A threshold SWEEP asks the same question fifty-seven times,
 and the view is recomputed for every one -- the extract build spent over ten
 minutes on a question that has a single answer.

 THE OBSERVATION THAT MAKES IT CHEAP

 'The first reading at or above threshold T' is the same row as 'the first
 reading whose RUNNING MAXIMUM is at or above T'. So the answer for EVERY
 threshold is determined by the points where a component's running maximum
 increases -- its frontier. Everything between two increases can never be a
 first crossing for any threshold.

 That collapses the 73,900 readings past the baseline window to 22,006 frontier
 points -- under a third as many -- and, more importantly, turns each threshold
 from a full scan with window functions into an index seek. The row reduction is
 the smaller half of the win; not recomputing the windows is the larger.

 WHY A TABLE AND NOT A VIEW

 A view would recompute the window functions on every reference, which is the
 problem. This is a derived table rebuilt from scratch every time 05 is
 deployed; it holds no information of its own and no one may edit it.

 THE RISK, AND WHAT GUARDS IT

 A materialised optimisation that disagrees with the thing it replaced is worse
 than the slow version, because it is fast and wrong. UAT-35 recomputes the
 confusion matrix directly from vw_SensorFeatures at the pinned threshold and
 requires it to match the frontier-based answer exactly.
==============================================================================
*/
IF OBJECT_ID('dbo.AlertFrontier', 'U') IS NOT NULL DROP TABLE dbo.AlertFrontier;
GO
WITH Scoped AS (
    SELECT f.InstallKey, f.ComponentSerial, f.ComponentCode, f.SortieSeq, f.SortieDate,
           f.VibVsBaseline, f.FlightHoursToRemoval, f.StressHoursToRemoval,
           -- Every alarm rule in this project ignores the baseline window: inside
           -- its first ten sorties a component is being compared against a mean
           -- that includes the reading itself, so the ratio means nothing.
           RunningMax = MAX(f.VibVsBaseline) OVER (
               PARTITION BY f.InstallKey ORDER BY f.SortieSeq ROWS UNBOUNDED PRECEDING)
    FROM dbo.vw_SensorFeatures f
    WHERE f.SortieSeq > 10 AND f.VibVsBaseline IS NOT NULL
),
Frontier AS (
    -- keep only the rows where the running maximum actually increases
    SELECT sc.*,
           PrevMax = LAG(sc.RunningMax) OVER (PARTITION BY sc.InstallKey ORDER BY sc.SortieSeq)
    FROM Scoped sc
)
SELECT
    InstallKey, ComponentSerial, ComponentCode, SortieSeq, SortieDate,
    ThresholdReached = RunningMax,
    VibVsBaseline, FlightHoursToRemoval, StressHoursToRemoval
INTO dbo.AlertFrontier
FROM Frontier
WHERE PrevMax IS NULL OR RunningMax > PrevMax;
GO

-- The seek that makes the sweep cheap: for a given install, find the earliest
-- frontier point at or above the threshold.
CREATE CLUSTERED INDEX IX_AlertFrontier ON dbo.AlertFrontier (InstallKey, ThresholdReached, SortieSeq);
GO

DECLARE @frontierRows INT = (SELECT COUNT(*) FROM dbo.AlertFrontier);
IF @frontierRows = 0
    THROW 53049, 'AlertFrontier is empty. Every alarm evaluation would report zero alerts raised, which reads as a fleet with nothing wrong.', 1;
PRINT CONCAT('AlertFrontier built: ', @frontierRows, ' frontier points from the monitored readings.');
GO

-- =============================================================================
-- fn_AlertEvaluation -- what the alert rule is worth.
--
-- The rule is deliberately simple: raise an alert the first time a component's
-- vibration reaches @VibThreshold times ITS OWN baseline. Simple because the
-- point of this function is not the rule -- it is the evaluation. The Python
-- model in python/ is far more capable and, as the case study shows, arrives at
-- the same operational conclusion, which is the useful result.
--
-- EVALUATED ON COMPLETED LIVES ONLY. A component still fitted has not yet had
-- the chance to fail, so scoring it as a false positive would penalise the rule
-- for alerts whose outcome is not yet known. That is survivorship bias pointed
-- the other way, and it is easy to do by accident.
-- =============================================================================
IF OBJECT_ID('dbo.fn_AlertEvaluation', 'IF') IS NOT NULL DROP FUNCTION dbo.fn_AlertEvaluation;
GO
CREATE FUNCTION dbo.fn_AlertEvaluation (@AsOfDate DATE, @VibThreshold DECIMAL(6,3))
RETURNS TABLE
AS
RETURN
/*
READ FROM THE BASE TABLES, NOT FROM vw_ComponentLifeHistory.

The population this function needs is identity and outcome: which component,
on which aircraft, did it fail, and when did it come off. None of that requires
accrued hours.

vw_ComponentLifeHistory computes accrued flight and stress hours by aggregating
every sortie for every install. Reading it here meant that work happened once
per THRESHOLD -- and a threshold sweep evaluates fifty-seven of them, so the
extract build spent more than ten minutes recomputing a figure this function
never looks at.

The predicate is deliberately identical to the view's, restated: monitored
components, a real outcome, a removal on or before the as-of date, and dates
that are physically possible. UAT-33 asserts the two populations agree.
*/
WITH Completed AS (
    SELECT
        ci.InstallKey,
        ct.ComponentCode,
        ct.LeadTimeFlightHours,
        b.BaseCode,
        a.TailNumber,
        ci.ComponentSerial,
        IsFailure = CAST(CASE WHEN ci.RemovalReason = 'Failure' THEN 1 ELSE 0 END AS BIT),
        RemovedDate = dr.[Date]
    FROM dbo.Fact_ComponentInstall ci
    JOIN dbo.Dim_ComponentType ct ON ct.ComponentTypeKey = ci.ComponentTypeKey
    JOIN dbo.Dim_Airframe a       ON a.AirframeKey = ci.AirframeKey
    JOIN dbo.Dim_Base b           ON b.BaseKey = a.BaseKey
    JOIN dbo.Dim_Date di          ON di.DateKey = ci.InstalledDateKey
    JOIN dbo.Dim_Date dr          ON dr.DateKey = ci.RemovedDateKey
    WHERE dr.[Date] <= @AsOfDate
      AND ct.ComponentCode IN ('ROTOR-ASSY','MOTOR-ESC')
      AND ci.RemovalReason IN ('Failure','Scheduled')
      -- Six installs were removed before they were installed. Their life has no
      -- measurable length, and five of them sat inside the published true
      -- positives before this predicate existed. They stay in the life history
      -- and in the data-quality report; they do not score a prediction.
      AND dr.[Date] >= di.[Date]
),
FirstAlert AS (
    -- The first reading that breached, per install, read from the frontier.
    -- Identical to scanning vw_SensorFeatures for the first row at or above
    -- the threshold -- see the note on AlertFrontier above, and UAT-35, which
    -- asserts the two agree rather than assuming it.
    SELECT InstallKey, SortieDate, SortieSeq, VibVsBaseline,
           FlightHoursToRemoval, StressHoursToRemoval
    FROM (
        SELECT af.InstallKey, af.SortieDate, af.SortieSeq, af.VibVsBaseline,
               af.FlightHoursToRemoval, af.StressHoursToRemoval,
               rn = ROW_NUMBER() OVER (PARTITION BY af.InstallKey ORDER BY af.SortieSeq)
        FROM dbo.AlertFrontier af
        WHERE af.SortieDate <= @AsOfDate
          AND af.ThresholdReached >= @VibThreshold
    ) z
    WHERE rn = 1
)
SELECT
    AsOfDate = @AsOfDate,
    VibThreshold = @VibThreshold,
    c.InstallKey, c.ComponentSerial, c.ComponentCode, c.TailNumber, c.BaseCode,
    c.LeadTimeFlightHours,
    c.IsFailure,
    c.RemovedDate,
    Alerted = CAST(CASE WHEN fa.InstallKey IS NOT NULL THEN 1 ELSE 0 END AS BIT),
    AlertDate = fa.SortieDate,
    AlertVibRatio = fa.VibVsBaseline,
    -- flight hours between the alert and the removal: the warning the planner got
    LeadTimeGivenFlightHours = fa.FlightHoursToRemoval,
    LeadTimeGivenStressHours = fa.StressHoursToRemoval,
    /*
    The four-way classification the whole finding rests on.

    TruePositive          alerted, and it failed -- the textbook success
    ActionableTruePositive alerted, failed, AND with more warning than the part
                          takes to obtain. This is the only one that prevented
                          anything.
    LateTruePositive      alerted, failed, too late to matter. Counted as a
                          success by precision and recall; changed nothing.
    FalsePositive         alerted, then came off serviceable at its interval
    FalseNegative         failed with no alert
    */
    TruePositive = CAST(CASE WHEN fa.InstallKey IS NOT NULL AND c.IsFailure = 1 THEN 1 ELSE 0 END AS BIT),
    ActionableTruePositive = CAST(CASE
        WHEN fa.InstallKey IS NOT NULL AND c.IsFailure = 1
         AND fa.FlightHoursToRemoval >= c.LeadTimeFlightHours THEN 1 ELSE 0 END AS BIT),
    LateTruePositive = CAST(CASE
        WHEN fa.InstallKey IS NOT NULL AND c.IsFailure = 1
         AND fa.FlightHoursToRemoval <  c.LeadTimeFlightHours THEN 1 ELSE 0 END AS BIT),
    FalsePositive = CAST(CASE WHEN fa.InstallKey IS NOT NULL AND c.IsFailure = 0 THEN 1 ELSE 0 END AS BIT),
    FalseNegative = CAST(CASE WHEN fa.InstallKey IS NULL     AND c.IsFailure = 1 THEN 1 ELSE 0 END AS BIT),
    TrueNegative  = CAST(CASE WHEN fa.InstallKey IS NULL     AND c.IsFailure = 0 THEN 1 ELSE 0 END AS BIT)
FROM Completed c
LEFT JOIN FirstAlert fa ON fa.InstallKey = c.InstallKey;
GO

-- =============================================================================
-- fn_AlertSummary -- the three numbers, and the fourth one that matters.
-- =============================================================================
IF OBJECT_ID('dbo.fn_AlertSummary', 'IF') IS NOT NULL DROP FUNCTION dbo.fn_AlertSummary;
GO
CREATE FUNCTION dbo.fn_AlertSummary (@AsOfDate DATE, @VibThreshold DECIMAL(6,3))
RETURNS TABLE
AS
RETURN
/*
THE MEDIAN IS COMPUTED SEPARATELY, AND IT HAS TO BE.

PERCENTILE_CONT is a window function with no aggregate form. Written inside
this GROUP BY it forces its arguments into the grouping key, and the first
version of this function did exactly that -- LeadTimeGivenFlightHours and
TruePositive went into the GROUP BY so the statement would compile.

It compiled, and that was all it did. Instead of one row per component it
returned one row per distinct lead-time value: seventy-five motor controllers
came back as dozens of rows reading Population 1, Precision 0.00. Every
individual row was well formed. Nothing was NULL that should not have been.
Nothing raised.

The creation gate at the foot of this file tested four of the five functions in
it, and this was the fifth.
*/
WITH Median AS (
    SELECT DISTINCT
        e.ComponentCode,
        MedianLeadTimeFlightHours = CAST(PERCENTILE_CONT(0.5) WITHIN GROUP (
            ORDER BY CASE WHEN e.TruePositive = 1 THEN e.LeadTimeGivenFlightHours END)
            OVER (PARTITION BY e.ComponentCode) AS DECIMAL(9,2))
    FROM dbo.fn_AlertEvaluation(@AsOfDate, @VibThreshold) e
),
Agg AS (
SELECT
    e.ComponentCode,
    LeadTimeFlightHours = MAX(e.LeadTimeFlightHours),
    Population   = COUNT(*),
    Failures     = SUM(CAST(e.IsFailure AS INT)),
    AlertsRaised = SUM(CAST(e.Alerted AS INT)),
    TP = SUM(CAST(e.TruePositive AS INT)),
    FP = SUM(CAST(e.FalsePositive AS INT)),
    FN = SUM(CAST(e.FalseNegative AS INT)),
    TN = SUM(CAST(e.TrueNegative AS INT)),
    ActionableTP = SUM(CAST(e.ActionableTruePositive AS INT)),
    LateTP       = SUM(CAST(e.LateTruePositive AS INT)),

    PrecisionPct = CAST(100.0 * SUM(CAST(e.TruePositive AS INT))
                        / NULLIF(SUM(CAST(e.Alerted AS INT)), 0) AS DECIMAL(6,2)),
    RecallPct    = CAST(100.0 * SUM(CAST(e.TruePositive AS INT))
                        / NULLIF(SUM(CAST(e.IsFailure AS INT)), 0) AS DECIMAL(6,2)),

    /*
    ActionableLeadTimePct -- of the failures the rule CAUGHT, the share it
    caught early enough to do something about.

    Deliberately measured over caught failures rather than over all failures, so
    it is not just recall wearing a different hat. It answers a separate
    question: when this thing fires, is the warning useful? A programme can fix
    recall by lowering the threshold. It cannot fix this one that way, because
    the component's degradation curve decides how much warning exists to be had.
    */
    ActionableLeadTimePct = CAST(100.0 * SUM(CAST(e.ActionableTruePositive AS INT))
                                 / NULLIF(SUM(CAST(e.TruePositive AS INT)), 0) AS DECIMAL(6,2)),

    -- Of ALL failures, the share both caught and caught in time. This is the
    -- number that describes what the programme actually prevented.
    PreventableFailurePct = CAST(100.0 * SUM(CAST(e.ActionableTruePositive AS INT))
                                 / NULLIF(SUM(CAST(e.IsFailure AS INT)), 0) AS DECIMAL(6,2))
FROM dbo.fn_AlertEvaluation(@AsOfDate, @VibThreshold) e
GROUP BY e.ComponentCode
)
SELECT
    AsOfDate = @AsOfDate,
    VibThreshold = @VibThreshold,
    a.*,
    m.MedianLeadTimeFlightHours
FROM Agg a
JOIN Median m ON m.ComponentCode = a.ComponentCode;
GO

-- =============================================================================
-- fn_AlertSummaryFleet -- the headline, across both monitored components.
--
-- Rolled up from the EVALUATION rows, not averaged from fn_AlertSummary. An
-- average of two percentages weights a population of 75 the same as one of 416
-- and quietly invents a third number that describes neither. Project 3 shipped
-- a cohort figure built that way.
-- =============================================================================
IF OBJECT_ID('dbo.fn_AlertSummaryFleet', 'IF') IS NOT NULL DROP FUNCTION dbo.fn_AlertSummaryFleet;
GO
CREATE FUNCTION dbo.fn_AlertSummaryFleet (@AsOfDate DATE, @VibThreshold DECIMAL(6,3))
RETURNS TABLE
AS
RETURN
SELECT
    AsOfDate = @AsOfDate,
    VibThreshold = @VibThreshold,
    Population   = COUNT(*),
    Failures     = SUM(CAST(e.IsFailure AS INT)),
    AlertsRaised = SUM(CAST(e.Alerted AS INT)),
    TP = SUM(CAST(e.TruePositive AS INT)),
    FP = SUM(CAST(e.FalsePositive AS INT)),
    FN = SUM(CAST(e.FalseNegative AS INT)),
    TN = SUM(CAST(e.TrueNegative AS INT)),
    ActionableTP = SUM(CAST(e.ActionableTruePositive AS INT)),
    LateTP       = SUM(CAST(e.LateTruePositive AS INT)),
    PredictionPrecisionPct = CAST(100.0 * SUM(CAST(e.TruePositive AS INT))
                                  / NULLIF(SUM(CAST(e.Alerted AS INT)), 0) AS DECIMAL(6,2)),
    PredictionRecallPct    = CAST(100.0 * SUM(CAST(e.TruePositive AS INT))
                                  / NULLIF(SUM(CAST(e.IsFailure AS INT)), 0) AS DECIMAL(6,2)),
    ActionableLeadTimePct  = CAST(100.0 * SUM(CAST(e.ActionableTruePositive AS INT))
                                  / NULLIF(SUM(CAST(e.TruePositive AS INT)), 0) AS DECIMAL(6,2)),
    PreventableFailurePct  = CAST(100.0 * SUM(CAST(e.ActionableTruePositive AS INT))
                                  / NULLIF(SUM(CAST(e.IsFailure AS INT)), 0) AS DECIMAL(6,2))
FROM dbo.fn_AlertEvaluation(@AsOfDate, @VibThreshold) e;
GO

-- =============================================================================
-- fn_PredictionKPI -- the three prediction targets, evaluated.
--
-- Separate from fn_FleetKPI for one reason: these depend on the alert threshold
-- as well as the date, and merging them would give fn_FleetKPI a second
-- parameter that four of its five metrics ignore. A function whose parameters
-- apply to only part of its output is the shape of defect this project has
-- already found twice.
-- =============================================================================
IF OBJECT_ID('dbo.fn_PredictionKPI', 'IF') IS NOT NULL DROP FUNCTION dbo.fn_PredictionKPI;
GO
CREATE FUNCTION dbo.fn_PredictionKPI (@AsOfDate DATE, @VibThreshold DECIMAL(6,3))
RETURNS TABLE
AS
RETURN
WITH F AS (
    SELECT * FROM dbo.fn_AlertSummaryFleet(@AsOfDate, @VibThreshold)
),
Metrics AS (
    SELECT MetricName = 'PredictionPrecisionPct', MetricValue = PredictionPrecisionPct,
           Numerator = CAST(TP AS DECIMAL(12,2)), Denominator = CAST(AlertsRaised AS DECIMAL(12,2)) FROM F
    UNION ALL
    SELECT 'PredictionRecallPct', PredictionRecallPct,
           CAST(TP AS DECIMAL(12,2)), CAST(Failures AS DECIMAL(12,2)) FROM F
    UNION ALL
    SELECT 'ActionableLeadTimePct', ActionableLeadTimePct,
           CAST(ActionableTP AS DECIMAL(12,2)), CAST(TP AS DECIMAL(12,2)) FROM F
)
SELECT
    AsOfDate = @AsOfDate,
    VibThreshold = @VibThreshold,
    m.MetricName,
    MetricValue = CAST(m.MetricValue AS DECIMAL(10,2)),
    m.Numerator,
    m.Denominator,
    t.TargetValue, t.WarningValue, t.Direction, t.Unit, t.[Description],
    RAGStatus = CASE
        WHEN m.MetricValue IS NULL THEN 'Grey'
        WHEN t.Direction = 'HigherBetter' THEN
            CASE WHEN m.MetricValue >= t.TargetValue  THEN 'Green'
                 WHEN m.MetricValue >= t.WarningValue THEN 'Amber'
                 ELSE 'Red' END
        ELSE
            CASE WHEN m.MetricValue <= t.TargetValue  THEN 'Green'
                 WHEN m.MetricValue <= t.WarningValue THEN 'Amber'
                 ELSE 'Red' END
    END
FROM Metrics m
JOIN dbo.Ref_FleetTargets t ON t.MetricName = m.MetricName;
GO

-- =============================================================================
-- fn_ThresholdRecommendation -- one threshold per component, not one per fleet.
--
-- THE DEFECT THIS EXISTS TO FIX
--
-- The fleet runs a single alarm threshold across two components whose
-- degradation curves have nothing in common. A rotor bearing's vibration rises
-- gradually from about half its life; a motor controller's stays flat and then
-- turns up sharply past 85%. One number cannot suit both, and at the setting
-- chosen to make the ROTOR's precision acceptable the motor controller alarm
-- becomes almost perfectly accurate and almost entirely useless.
--
-- So the search objective here is ACTIONABLE LEAD TIME, subject to a precision
-- floor -- not precision, and not F1. Precision is a constraint because crews
-- stop acting on alarms that are wrong too often; lead time is the objective
-- because it is the only one of the two that prevents anything.
--
-- The floor is a parameter. It is the single most arguable number in the
-- recommendation and it belongs in the caller's hands.
-- =============================================================================
IF OBJECT_ID('dbo.fn_ThresholdRecommendation', 'IF') IS NOT NULL DROP FUNCTION dbo.fn_ThresholdRecommendation;
GO
CREATE FUNCTION dbo.fn_ThresholdRecommendation (@AsOfDate DATE, @MinPrecisionPct DECIMAL(6,2))
RETURNS TABLE
AS
RETURN
WITH Steps AS (
    -- 1.20 to 4.00 in steps of 0.05
    SELECT v = CAST(1.20 + (n * 0.05) AS DECIMAL(6,3))
    FROM (SELECT TOP (57) n = ROW_NUMBER() OVER (ORDER BY (SELECT NULL)) - 1
          FROM sys.all_objects) x
),
Scored AS (
    SELECT s.*
    FROM Steps t
    CROSS APPLY dbo.fn_AlertSummary(@AsOfDate, t.v) s
),
Ranked AS (
    SELECT sc.*,
           rn = ROW_NUMBER() OVER (
                PARTITION BY sc.ComponentCode
                ORDER BY sc.ActionableLeadTimePct DESC,   -- the objective
                         sc.PrecisionPct DESC,            -- then the safer alarm
                         sc.VibThreshold ASC)             -- then stable
    FROM Scored sc
    WHERE sc.PrecisionPct >= @MinPrecisionPct
      AND sc.AlertsRaised > 0
)
SELECT
    AsOfDate = @AsOfDate,
    MinPrecisionPct = @MinPrecisionPct,
    r.ComponentCode,
    r.LeadTimeFlightHours,
    RecommendedThreshold = r.VibThreshold,
    r.Population, r.Failures, r.AlertsRaised,
    r.TP, r.FP, r.FN, r.ActionableTP, r.LateTP,
    r.PrecisionPct, r.RecallPct, r.ActionableLeadTimePct, r.PreventableFailurePct,
    r.MedianLeadTimeFlightHours,
    -- what the current single fleet threshold delivers for this component,
    -- so the recommendation can be read as a change rather than as a number
    CurrentThreshold        = cur.VibThreshold,
    CurrentPrecisionPct     = cur.PrecisionPct,
    CurrentActionablePct    = cur.ActionableLeadTimePct,
    CurrentMedianLeadHours  = cur.MedianLeadTimeFlightHours,
    ActionablePointsGained  = CAST(r.ActionableLeadTimePct - cur.ActionableLeadTimePct AS DECIMAL(7,2)),
    PrecisionPointsGiven    = CAST(cur.PrecisionPct - r.PrecisionPct AS DECIMAL(7,2)),
    ExtraFailuresPrevented  = r.ActionableTP - cur.ActionableTP
FROM Ranked r
CROSS APPLY dbo.fn_AlertSummary(@AsOfDate,
        (SELECT CAST(SettingValue AS DECIMAL(6,3)) FROM dbo.Ref_Reporting
         WHERE SettingName = 'AlertVibThreshold')) cur
WHERE r.rn = 1 AND cur.ComponentCode = r.ComponentCode;
GO

-- =============================================================================
-- fn_ActionQueue -- what to do on Monday.
--
-- A list of 131 overdue components is not a plan; it is the same problem
-- restated. The queue ranks by consequence, converts each job into hangar
-- hours, and cuts the line where the week's capacity runs out.
--
-- THE ACTION FOLLOWS THE STATE, NOT THE RANK. A component past its stress
-- interval needs a slot. One whose PROJECTION crosses the interval inside its
-- part lead time needs a part ORDERED, which is a different person on a
-- different day, and lumping both into 'overdue' loses that.
-- =============================================================================
IF OBJECT_ID('dbo.fn_ActionQueue', 'IF') IS NOT NULL DROP FUNCTION dbo.fn_ActionQueue;
GO
CREATE FUNCTION dbo.fn_ActionQueue (@AsOfDate DATE, @HangarHoursPerWeek DECIMAL(9,2))
RETURNS TABLE
AS
RETURN
WITH Candidates AS (
    SELECT
        w.*,
        ActionCode = CASE
            WHEN w.IsOverdueByStress = 1 AND w.Criticality = 'FlightCritical' THEN 'GROUND_AND_REPLACE'
            WHEN w.IsOverdueByStress = 1                                      THEN 'REPLACE_AT_NEXT_SLOT'
            -- A NULL projection means the airframe has not flown, so there is
            -- no rate to project at. It falls through to MONITOR rather than
            -- being silently treated as either urgent or safe.
            WHEN w.ProjectedFlightHoursToStressLimit <= w.LeadTimeFlightHours THEN 'ORDER_PART_NOW'
            WHEN w.ProjectedFlightHoursToStressLimit <= 40.0                  THEN 'SCHEDULE_THIS_MONTH'
            ELSE 'MONITOR' END,
        /*
        Hangar hours per job, by criticality, matching the downtime the fleet
        records for a SCHEDULED removal. An unscheduled one costs roughly three
        times as much, which is the argument for doing any of this.

        A procurement action costs ZERO hangar hours. ORDER_PART_NOW is a
        purchase order, not a bay and a technician, and charging it bay time
        consumed capacity the week never spent. It does not move the published
        cut today, because the cut falls inside the GROUND_AND_REPLACE block,
        but it would the moment capacity rose -- which is the one question this
        queue exists to answer.
        */
        JobHours = CASE
            WHEN w.IsOverdueByStress = 0
             AND w.ProjectedFlightHoursToStressLimit <= w.LeadTimeFlightHours THEN 0.0
            ELSE CASE w.Criticality
                WHEN 'FlightCritical'  THEN 4.5
                WHEN 'MissionCritical' THEN 2.5
                ELSE 1.5 END
        END
    FROM dbo.fn_ComponentWear(@AsOfDate) w
    WHERE w.IsInFleetAsOf = 1
),
Ranked AS (
    SELECT c.*,
        PriorityRank = ROW_NUMBER() OVER (ORDER BY
            /*
            The tie-break chain, most consequential first:
              1. flight-critical before anything else
              2. already overdue before merely projected
              3. furthest past its limit first
              4. soonest to reach the limit first
              5. InstallKey, so the order is stable between runs -- an
                 unstable queue is one nobody can work through, because the
                 rows move under them between refreshes
            */
            CASE c.Criticality WHEN 'FlightCritical' THEN 0 WHEN 'MissionCritical' THEN 1 ELSE 2 END,
            CASE WHEN c.IsOverdueByStress = 1 THEN 0 ELSE 1 END,
            CASE WHEN c.IsOverdueByStress = 1 THEN c.RemainingStressHours END ASC,
            c.ProjectedFlightHoursToStressLimit ASC,
            c.InstallKey)
    FROM Candidates c
    WHERE c.ActionCode <> 'MONITOR'
)
SELECT
    AsOfDate = @AsOfDate,
    r.PriorityRank,
    r.InstallKey, r.ComponentSerial, r.ComponentCode, r.ComponentName, r.Criticality,
    r.TailNumber, r.BaseCode, r.BaseName, r.PositionNo,
    r.InstalledDate,
    r.FlightHours, r.StressHours,
    r.IntervalFlightHours, r.IntervalStressHours,
    r.PctOfHourInterval, r.PctOfStressInterval,
    r.IsOverdueByHours, r.IsOverdueByStress, r.IsHiddenOverdue,
    r.StressPerFlightHour,
    r.RemainingStressHours,
    r.ProjectedFlightHoursToStressLimit,
    r.LeadTimeFlightHours,
    r.ActionCode,
    r.JobHours,
    CumulativeJobHours = CAST(SUM(r.JobHours) OVER (ORDER BY r.PriorityRank ROWS UNBOUNDED PRECEDING) AS DECIMAL(10,2)),
    -- the line where the week's capacity runs out
    IsThisWeek = CAST(CASE WHEN SUM(r.JobHours) OVER (ORDER BY r.PriorityRank ROWS UNBOUNDED PRECEDING)
                                <= @HangarHoursPerWeek THEN 1 ELSE 0 END AS BIT),
    RecommendedAction = CASE r.ActionCode
        WHEN 'GROUND_AND_REPLACE'   THEN CONCAT('Ground ', r.TailNumber, '. ', r.ComponentName,
             ' at position ', r.PositionNo, ' is at ', r.PctOfStressInterval,
             '% of its stress interval while reading ', r.PctOfHourInterval, '% on the hour meter.')
        WHEN 'REPLACE_AT_NEXT_SLOT' THEN CONCAT('Book ', r.TailNumber, ' for ', r.ComponentName,
             ' replacement. Past its stress interval; not flight-critical, so it can wait for a slot.')
        WHEN 'ORDER_PART_NOW'       THEN CONCAT('Order a ', r.ComponentName, ' for ', r.TailNumber,
             '. Projected to reach its stress limit in ', r.ProjectedFlightHoursToStressLimit,
             ' flight hours against a ', r.LeadTimeFlightHours, ' hour lead time.')
        ELSE CONCAT('Schedule ', r.TailNumber, ' within the month. ', r.ComponentName,
             ' projected to reach its stress limit in ', r.ProjectedFlightHoursToStressLimit, ' flight hours.')
    END
FROM Ranked r;
GO

PRINT 'KPI layer created: fn_FleetKPI, fn_BaseScorecard, fn_AlertEvaluation, fn_AlertSummary, fn_ActionQueue.';
GO

-- =============================================================================
-- Creation gate
-- =============================================================================
DECLARE @AsOf DATE = '2026-09-30';
IF (SELECT COUNT(*) FROM dbo.fn_FleetKPI(@AsOf)) < 5
    THROW 53040, 'fn_FleetKPI did not return all five metrics. A metric name does not match Ref_FleetTargets.', 1;
IF (SELECT COUNT(*) FROM dbo.fn_BaseScorecard(@AsOf)) <> (SELECT COUNT(*) FROM dbo.Dim_Base)
    THROW 53041, 'fn_BaseScorecard did not return one row per base.', 1;
IF (SELECT COUNT(*) FROM dbo.fn_ActionQueue(@AsOf, 120.0)) = 0
    THROW 53042, 'fn_ActionQueue is empty. Either nothing is overdue, or the projection is not computing.', 1;
IF (SELECT COUNT(*) FROM dbo.fn_AlertEvaluation(@AsOf, 1.35)) = 0
    THROW 53043, 'fn_AlertEvaluation returned no completed lives to score.', 1;

/*
GRAIN, not just presence.

fn_AlertSummary must return exactly one row per monitored component code. The
first version returned dozens per code because a window function had dragged
its arguments into the GROUP BY, and every row looked reasonable on its own.
'Returns rows' would have passed it; only asserting the GRAIN catches it.

Every function above is now checked for the shape it promises, not merely for
being non-empty.
*/
IF (SELECT COUNT(*) FROM dbo.fn_AlertSummary(@AsOf, 1.35)) <>
   (SELECT COUNT(DISTINCT ComponentCode) FROM dbo.fn_AlertEvaluation(@AsOf, 1.35))
    THROW 53045, 'fn_AlertSummary is not at one row per component code. An aggregate has been split by an unintended grouping column.', 1;

IF (SELECT COUNT(*) FROM dbo.fn_ComponentWear(@AsOf)) <>
   (SELECT COUNT(DISTINCT InstallKey) FROM dbo.fn_ComponentWear(@AsOf))
    THROW 53046, 'fn_ComponentWear is returning duplicate InstallKeys -- a join has fanned out and every percentage it feeds is wrong.', 1;

-- fn_ThresholdRecommendation must return exactly one row per monitored
-- component. Returning none means the precision floor is unreachable and the
-- recommendation silently becomes 'do nothing'.
IF (SELECT COUNT(*) FROM dbo.fn_ThresholdRecommendation(@AsOf, 75.00)) <> 2
    THROW 53047, 'fn_ThresholdRecommendation did not return one row per monitored component. Either the precision floor is unreachable for one of them, or the ranking has fanned out.', 1;

/*
EVERY target must be evaluated by something.

Ref_FleetTargets carries eight rows. Three of them -- including
ActionableLeadTimePct, the target the second finding is about -- were compared
to nothing at all, and a target nothing evaluates cannot be failed. This gate
makes adding a target without an evaluation a build error.
*/
IF EXISTS (
    SELECT 1 FROM dbo.Ref_FleetTargets t
    WHERE NOT EXISTS (SELECT 1 FROM dbo.fn_FleetKPI(@AsOf) k WHERE k.MetricName = t.MetricName)
      AND NOT EXISTS (SELECT 1 FROM dbo.fn_PredictionKPI(@AsOf, 3.400) k WHERE k.MetricName = t.MetricName)
)
    THROW 53048, 'A row in Ref_FleetTargets is not evaluated by any KPI function. A target nothing compares against is a target nothing can fail.', 1;

/*
The gate that would have caught Project 4's flat trend line.

fn_FleetKPI takes an as-of date, so its answer must MOVE when the date moves.
A function that accepts a point-in-time parameter and quietly reports 'now' for
one of its outputs produces a trend chart that is a flat line for twenty months
and looks entirely plausible. Asserting that two different dates disagree is a
cheap test for an expensive class of defect.
*/
DECLARE @Now DECIMAL(10,2), @Then DECIMAL(10,2);
SELECT @Now  = MetricValue FROM dbo.fn_FleetKPI('2026-09-30') WHERE MetricName = 'StressCompliancePct';
SELECT @Then = MetricValue FROM dbo.fn_FleetKPI('2025-03-31') WHERE MetricName = 'StressCompliancePct';
IF @Now = @Then
    THROW 53044, 'fn_FleetKPI returns the same stress compliance eighteen months apart. The @AsOfDate parameter is not reaching every metric.', 1;

PRINT 'KPI gate passed.';
GO
