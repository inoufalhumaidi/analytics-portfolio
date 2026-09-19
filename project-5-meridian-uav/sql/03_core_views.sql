/*
================================================================================
Project 5 — Meridian UAV Services: Predictive Maintenance
Script:  03_core_views.sql
Purpose: The wear model. Everything downstream reads from here.

THE ONE IDEA IN THIS FILE

    A flight hour is not a unit of wear.

    The maintenance system counts flight hours because that is what the hour
    meter produces. But an hour of mountain survey at 22kg and an hour of
    coastal survey at 12kg damage a rotor bearing by very different amounts, and
    an hour of circuit training produces twelve landings while an hour of cargo
    relay produces one. Ref_StressModel converts a sortie into STRESS HOURS:

        stress hours = flight hours x profile multiplier
                     + landings     x cycle penalty
                     + flight hours x (payload - 12kg) x payload penalty

    Every view here measures component life in stress hours and reports flight
    hours alongside, so the two can be compared rather than one silently
    standing in for the other.

WHY THIS IS A DATED FACT AND NOT AN ATTRIBUTE

    'Hours on this component' is not a property of the component. It is a
    question with a date in it, and the answer changes every time the airframe
    flies. So it is computed by summing sorties inside a window, not stored on
    Fact_ComponentInstall where it would have to be maintained and would quietly
    go stale.

    This is the same modelling decision as Project 3's dated price agreements
    and Project 4's build-scoped verifications. Storing a time-dependent
    quantity as an attribute lets today's answer govern yesterday's question.

FLEET MEMBERSHIP IS A DATE QUESTION, NOT A STATUS FLAG

    Dim_Airframe.AirframeStatus says whether an aircraft is in the fleet TODAY.
    It is an attribute, and using it to choose the population for a
    point-in-time question is the same mistake as storing accrued hours on the
    component row -- today's answer governing yesterday's question.

    MU-003 retired in September 2025. It flew throughout 2024 and most of 2025,
    and every compliance figure computed for those dates must include it. The
    first version of this file filtered on AirframeStatus = 'Active', so the
    entire eighteen-month compliance trend was computed over the 28 aircraft
    that happen to be flying today, and the availability denominator was 28 at
    every date in the series.

    fn_AirframeStress and fn_ComponentWear therefore expose IsInFleetAsOf, which
    is computed from the in-service and retirement dates, and every caller
    filters on that.

THE @AsOfDate DISCIPLINE

    Every function here takes @AsOfDate and must use it for EVERY output it
    returns. Project 4 shipped a function that accepted the parameter and
    answered about 'now' for one of its metrics; the trend chart drawn from it
    was a flat line for twenty months and looked plausible. A function that
    offers a point-in-time parameter and ignores it for one column is worse than
    one that never offered it, because the caller has no way to tell.

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
-- vw_SortieStress -- the conversion, in one place.
--
-- Every stress figure in this project resolves to this view. It is deliberately
-- a view and not a column on Fact_Sortie: the stress model is reference data an
-- engineer is expected to argue with, and baking its output into the fact table
-- would mean a changed multiplier silently disagreeing with history.
-- =============================================================================
IF OBJECT_ID('dbo.vw_SortieStress', 'V') IS NOT NULL DROP VIEW dbo.vw_SortieStress;
GO
CREATE VIEW dbo.vw_SortieStress
AS
SELECT
    s.SortieKey,
    s.SortieID,
    s.AirframeKey,
    a.TailNumber,
    a.BaseKey,
    s.SortieDateKey,
    d.[Date]          AS SortieDate,
    p.ProfileCode,
    p.ProfileName,
    s.FlightMinutes,
    FlightHours = CAST(s.FlightMinutes / 60.0 AS DECIMAL(9,4)),
    s.Landings,
    s.PayloadKg,
    sm.StressMultiplier,
    -- The three terms are exposed separately, not just their sum. An engineer
    -- disputing the model needs to see whether a figure is driven by the hourly
    -- multiplier, the landing count or the payload -- they imply different
    -- fixes, and a single number implies none of them.
    HourlyStress  = CAST((s.FlightMinutes / 60.0) * sm.StressMultiplier AS DECIMAL(9,4)),
    CycleStress   = CAST(s.Landings * sm.CyclePenaltyHours AS DECIMAL(9,4)),
    PayloadStress = CAST((s.FlightMinutes / 60.0) * (s.PayloadKg - 12.0) * sm.PayloadPenaltyPerKg AS DECIMAL(9,4)),
    StressHours   = CAST((s.FlightMinutes / 60.0) * sm.StressMultiplier
                       + s.Landings * sm.CyclePenaltyHours
                       + (s.FlightMinutes / 60.0) * (s.PayloadKg - 12.0) * sm.PayloadPenaltyPerKg AS DECIMAL(9,4))
FROM dbo.Fact_Sortie s
JOIN dbo.Dim_Airframe a        ON a.AirframeKey = s.AirframeKey
JOIN dbo.Dim_Date d            ON d.DateKey = s.SortieDateKey
JOIN dbo.Dim_MissionProfile p  ON p.MissionProfileKey = s.MissionProfileKey
JOIN dbo.Ref_StressModel sm    ON sm.ProfileCode = p.ProfileCode;
GO

-- =============================================================================
-- fn_ComponentWear -- what is fitted right now, and how worn is it.
--
-- Returns one row per component instance FITTED AT @AsOfDate, with its life
-- measured both ways.
--
-- ON FAN-OUT. The sortie totals are aggregated in a CTE keyed on InstallKey and
-- only then joined back. Joining Fact_ComponentInstall directly to the sortie
-- rows and aggregating afterwards would multiply each install by its sortie
-- count -- and because both the numerator and the denominator inflate, the
-- resulting percentage still looks like a percentage. Project 3 shipped exactly
-- that defect: 36 of 36 vendors wrong, every figure plausible, nothing to see
-- unless you reconciled a single row by hand.
-- =============================================================================
IF OBJECT_ID('dbo.fn_ComponentWear', 'IF') IS NOT NULL DROP FUNCTION dbo.fn_ComponentWear;
GO
CREATE FUNCTION dbo.fn_ComponentWear (@AsOfDate DATE)
RETURNS TABLE
AS
RETURN
WITH FittedInstalls AS (
    -- 'Fitted as of' is a date question, not a status flag. A component removed
    -- after @AsOfDate was still on the aircraft on @AsOfDate.
    SELECT
        ci.InstallKey, ci.ComponentSerial, ci.ComponentTypeKey, ci.AirframeKey, ci.PositionNo,
        InstalledDate = di.[Date],
        /*
        THE HANDOVER SORTIE BELONGS TO THE PART THAT FLEW IT.

        A part comes off ON a sortie and its replacement goes on afterwards, so
        both installs carry that date. Counting it at both ends charged 1,144
        sorties in this dataset to two components at once -- inflating the
        accrued hours of every replacement and, at the margin, moving components
        across their overdue threshold.

        The outgoing part keeps the sortie. The incoming part starts the day
        after, unless nothing preceded it on that position, in which case the
        component was fitted before the aircraft ever flew and keeps its first
        day.
        */
        EffectiveStartDate = CASE WHEN prev.InstallKey IS NULL
                                  THEN di.[Date]
                                  ELSE DATEADD(DAY, 1, di.[Date]) END
    FROM dbo.Fact_ComponentInstall ci
    JOIN dbo.Dim_Date di       ON di.DateKey = ci.InstalledDateKey
    LEFT JOIN dbo.Dim_Date dr  ON dr.DateKey = ci.RemovedDateKey
    LEFT JOIN dbo.Fact_ComponentInstall prev
           ON prev.AirframeKey      = ci.AirframeKey
          AND prev.ComponentTypeKey = ci.ComponentTypeKey
          AND prev.PositionNo       = ci.PositionNo
          AND prev.RemovedDateKey   = ci.InstalledDateKey
    WHERE di.[Date] <= @AsOfDate
      AND (dr.[Date] IS NULL OR dr.[Date] > @AsOfDate)
),
Accrued AS (
    -- one row per install; aggregation happens HERE, before any join back
    SELECT
        fi.InstallKey,
        Sorties      = COUNT(*),
        FlightHours  = CAST(SUM(ss.FlightHours) AS DECIMAL(11,4)),
        StressHours  = CAST(SUM(ss.StressHours) AS DECIMAL(11,4)),
        Landings     = SUM(ss.Landings),
        LastSortieDate = MAX(ss.SortieDate)
    FROM FittedInstalls fi
    JOIN dbo.vw_SortieStress ss
      ON ss.AirframeKey = fi.AirframeKey
     AND ss.SortieDate >= fi.EffectiveStartDate
     AND ss.SortieDate <= @AsOfDate
    GROUP BY fi.InstallKey
),
RecentRate AS (
    /*
    The airframe's stress-per-flight-hour over its last 30 sorties before
    @AsOfDate.

    Deliberately RECENT rather than lifetime. An airframe moved from coastal
    survey to mountain work six weeks ago has a lifetime ratio that describes a
    job it no longer does, and the whole point of the forward projection is to
    answer 'at the rate it is being flown NOW'. Thirty sorties is roughly two
    months of flying -- long enough to be stable, short enough to follow a
    change of task.
    */
    SELECT r.AirframeKey,
           StressPerFlightHour = CAST(SUM(r.StressHours) / NULLIF(SUM(r.FlightHours), 0) AS DECIMAL(9,4)),
           SortiesInWindow = COUNT(*)
    FROM (
        SELECT ss.AirframeKey, ss.StressHours, ss.FlightHours,
               rn = ROW_NUMBER() OVER (PARTITION BY ss.AirframeKey ORDER BY ss.SortieDate DESC, ss.SortieKey DESC)
        FROM dbo.vw_SortieStress ss
        WHERE ss.SortieDate <= @AsOfDate
    ) r
    WHERE r.rn <= 30
    GROUP BY r.AirframeKey
)
SELECT
    fi.InstallKey,
    fi.ComponentSerial,
    fi.ComponentTypeKey,
    ct.ComponentCode,
    ct.ComponentName,
    ct.Criticality,
    ct.LeadTimeFlightHours,
    fi.AirframeKey,
    a.TailNumber,
    a.ModelCode,
    a.AirframeStatus,
    -- Was this aircraft in the fleet ON @AsOfDate? Computed from dates, never
    -- from AirframeStatus, which describes today. Every caller filters on this.
    IsInFleetAsOf = CAST(CASE WHEN isd.[Date] <= @AsOfDate
                               AND (rtd.[Date] IS NULL OR rtd.[Date] > @AsOfDate)
                              THEN 1 ELSE 0 END AS BIT),
    b.BaseKey,
    b.BaseCode,
    b.BaseName,
    fi.PositionNo,
    fi.InstalledDate,
    AsOfDate = @AsOfDate,
    Sorties     = ISNULL(ac.Sorties, 0),
    Landings    = ISNULL(ac.Landings, 0),
    -- A component fitted to an airframe that has not flown since has accrued
    -- nothing. ISNULL here is a genuine zero, not a missing value.
    FlightHours = CAST(ISNULL(ac.FlightHours, 0) AS DECIMAL(11,4)),
    StressHours = CAST(ISNULL(ac.StressHours, 0) AS DECIMAL(11,4)),
    si.IntervalFlightHours,
    si.IntervalStressHours,
    si.WeibullShape,

    -- THE TWO ANSWERS, side by side. This pair is the project.
    PctOfHourInterval   = CAST(100.0 * ISNULL(ac.FlightHours, 0) / si.IntervalFlightHours  AS DECIMAL(9,2)),
    PctOfStressInterval = CAST(100.0 * ISNULL(ac.StressHours, 0) / si.IntervalStressHours  AS DECIMAL(9,2)),

    IsOverdueByHours  = CAST(CASE WHEN ISNULL(ac.FlightHours, 0) >= si.IntervalFlightHours  THEN 1 ELSE 0 END AS BIT),
    IsOverdueByStress = CAST(CASE WHEN ISNULL(ac.StressHours, 0) >= si.IntervalStressHours THEN 1 ELSE 0 END AS BIT),

    /*
    HiddenOverdue -- past the stress interval, inside the hour interval.

    This is the population the maintenance system cannot see. It is not a
    prediction and it involves no model: it is the same parts the operator is
    already tracking, measured in the unit their wear actually accrues in.
    */
    IsHiddenOverdue = CAST(CASE WHEN ISNULL(ac.StressHours, 0) >= si.IntervalStressHours
                                 AND ISNULL(ac.FlightHours, 0) <  si.IntervalFlightHours
                                THEN 1 ELSE 0 END AS BIT),

    -- How much harder this airframe is worked than its hour meter suggests.
    -- NULL for an airframe that has not flown, not 1.0. A default of 1.0 is a
    -- claim that the aircraft flies benign work, which reads as reassuring and
    -- is unsupported -- the same substitution fn_AirframeStress refuses below.
    StressPerFlightHour = rr.StressPerFlightHour,

    -- Remaining life in each unit. Negative means past due; NOT clamped to
    -- zero, because 'how far past due' is the whole scheduling signal and a
    -- floor at zero would throw it away.
    RemainingStressHours = CAST(si.IntervalStressHours - ISNULL(ac.StressHours, 0) AS DECIMAL(11,4)),
    RemainingFlightHours = CAST(si.IntervalFlightHours - ISNULL(ac.FlightHours, 0) AS DECIMAL(11,4)),

    /*
    ProjectedFlightHoursToStressLimit -- the planner's number.

    Remaining stress hours converted into the unit crews actually schedule in,
    at the rate THIS airframe is currently being flown. A part with 40 stress
    hours left has 40 flight hours of life on a coastal airframe and 17 on a
    mountain one, and the difference decides whether it needs a slot this month.
    */
    -- NULL rather than a guess when the rate is unknown: a projection built on
    -- an assumed rate is a number with no evidence behind it, and the action
    -- queue treats a NULL projection as 'cannot schedule this yet'.
    ProjectedFlightHoursToStressLimit =
        CAST((si.IntervalStressHours - ISNULL(ac.StressHours, 0))
             / NULLIF(rr.StressPerFlightHour, 0) AS DECIMAL(11,2)),

    LastSortieDate = ac.LastSortieDate
FROM FittedInstalls fi
JOIN dbo.Dim_ComponentType ct   ON ct.ComponentTypeKey = fi.ComponentTypeKey
JOIN dbo.Ref_ServiceInterval si ON si.ComponentCode = ct.ComponentCode
JOIN dbo.Dim_Airframe a         ON a.AirframeKey = fi.AirframeKey
JOIN dbo.Dim_Date isd           ON isd.DateKey = a.InServiceDateKey
LEFT JOIN dbo.Dim_Date rtd      ON rtd.DateKey = a.RetiredDateKey
JOIN dbo.Dim_Base b             ON b.BaseKey = a.BaseKey
LEFT JOIN Accrued ac            ON ac.InstallKey = fi.InstallKey
LEFT JOIN RecentRate rr         ON rr.AirframeKey = fi.AirframeKey;
GO

-- =============================================================================
-- fn_AirframeStress -- the fleet, one row per airframe.
--
-- StressRatio is the single number that explains the whole dataset: how many
-- stress hours the airframe accrues per logged flight hour. An operator that
-- knew only this column would already be most of the way to the finding.
-- =============================================================================
IF OBJECT_ID('dbo.fn_AirframeStress', 'IF') IS NOT NULL DROP FUNCTION dbo.fn_AirframeStress;
GO
CREATE FUNCTION dbo.fn_AirframeStress (@AsOfDate DATE)
RETURNS TABLE
AS
RETURN
WITH Totals AS (
    SELECT ss.AirframeKey,
           Sorties     = COUNT(*),
           FlightHours = CAST(SUM(ss.FlightHours) AS DECIMAL(11,4)),
           StressHours = CAST(SUM(ss.StressHours) AS DECIMAL(11,4)),
           Landings    = SUM(ss.Landings),
           HourlyStress  = CAST(SUM(ss.HourlyStress)  AS DECIMAL(11,4)),
           CycleStress   = CAST(SUM(ss.CycleStress)   AS DECIMAL(11,4)),
           PayloadStress = CAST(SUM(ss.PayloadStress) AS DECIMAL(11,4)),
           FirstSortie = MIN(ss.SortieDate),
           LastSortie  = MAX(ss.SortieDate)
    FROM dbo.vw_SortieStress ss
    WHERE ss.SortieDate <= @AsOfDate
    GROUP BY ss.AirframeKey
),
DominantProfile AS (
    -- the profile the airframe spends the most STRESS on, not the most sorties;
    -- eight benign flights and two mountain ones is a mountain airframe
    SELECT AirframeKey, ProfileCode, ProfileName, ProfileStress
    FROM (
        SELECT ss.AirframeKey, ss.ProfileCode, ss.ProfileName,
               ProfileStress = SUM(ss.StressHours),
               rn = ROW_NUMBER() OVER (PARTITION BY ss.AirframeKey
                                       ORDER BY SUM(ss.StressHours) DESC, ss.ProfileCode)
        FROM dbo.vw_SortieStress ss
        WHERE ss.SortieDate <= @AsOfDate
        GROUP BY ss.AirframeKey, ss.ProfileCode, ss.ProfileName
    ) z
    WHERE rn = 1
)
SELECT
    a.AirframeKey,
    a.TailNumber,
    a.ModelCode,
    a.AirframeStatus,
    -- in the fleet ON @AsOfDate, computed from dates. AirframeStatus is kept
    -- alongside because a reader will ask for it, but nothing filters on it.
    IsInFleetAsOf = CAST(CASE WHEN isd.[Date] <= @AsOfDate
                               AND (rtd.[Date] IS NULL OR rtd.[Date] > @AsOfDate)
                              THEN 1 ELSE 0 END AS BIT),
    InServiceDate = isd.[Date],
    RetiredDate   = rtd.[Date],
    b.BaseKey, b.BaseCode, b.BaseName, b.Region,
    AsOfDate = @AsOfDate,
    Sorties     = ISNULL(t.Sorties, 0),
    Landings    = ISNULL(t.Landings, 0),
    FlightHours = CAST(ISNULL(t.FlightHours, 0) AS DECIMAL(11,4)),
    StressHours = CAST(ISNULL(t.StressHours, 0) AS DECIMAL(11,4)),
    -- An airframe with no sorties has no ratio. 1.0 would be a lie that reads
    -- as 'benign'; NULL forces the caller to decide what to do about it.
    StressRatio = CAST(ISNULL(t.StressHours, 0) / NULLIF(t.FlightHours, 0) AS DECIMAL(9,4)),
    -- the decomposition, so a high ratio can be attributed
    PctStressFromHours   = CAST(100.0 * ISNULL(t.HourlyStress, 0)  / NULLIF(t.StressHours, 0) AS DECIMAL(6,2)),
    PctStressFromCycles  = CAST(100.0 * ISNULL(t.CycleStress, 0)   / NULLIF(t.StressHours, 0) AS DECIMAL(6,2)),
    PctStressFromPayload = CAST(100.0 * ISNULL(t.PayloadStress, 0) / NULLIF(t.StressHours, 0) AS DECIMAL(6,2)),
    dp.ProfileCode AS DominantProfileCode,
    dp.ProfileName AS DominantProfileName,
    t.FirstSortie, t.LastSortie
FROM dbo.Dim_Airframe a
JOIN dbo.Dim_Date isd      ON isd.DateKey = a.InServiceDateKey
LEFT JOIN dbo.Dim_Date rtd ON rtd.DateKey = a.RetiredDateKey
JOIN dbo.Dim_Base b        ON b.BaseKey = a.BaseKey
LEFT JOIN Totals t         ON t.AirframeKey = a.AirframeKey
LEFT JOIN DominantProfile dp ON dp.AirframeKey = a.AirframeKey;
GO

-- =============================================================================
-- vw_ComponentLifeHistory -- completed lives, for reliability analysis.
--
-- CENSORING. A component replaced on schedule did not fail; it was removed
-- while still serviceable, and all we know is that its life was AT LEAST that
-- long. Treating those as failures understates component life badly -- here it
-- would drag the apparent mean life down by roughly the share of scheduled
-- removals, which is most of them at the benign bases.
--
-- IsCensored marks them so the Python survival model can handle them properly
-- rather than silently averaging over two different kinds of event.
-- =============================================================================
IF OBJECT_ID('dbo.vw_ComponentLifeHistory', 'V') IS NOT NULL DROP VIEW dbo.vw_ComponentLifeHistory;
GO
CREATE VIEW dbo.vw_ComponentLifeHistory
AS
WITH Windows AS (
    -- same handover rule as fn_ComponentWear: the sortie a part came off on
    -- belongs to the part that was fitted for it, not to its replacement
    SELECT ci.InstallKey,
           EffectiveStartDate = CASE WHEN prev.InstallKey IS NULL
                                     THEN di.[Date] ELSE DATEADD(DAY, 1, di.[Date]) END,
           EndDate = dr.[Date]
    FROM dbo.Fact_ComponentInstall ci
    JOIN dbo.Dim_Date di ON di.DateKey = ci.InstalledDateKey
    JOIN dbo.Dim_Date dr ON dr.DateKey = ci.RemovedDateKey
    LEFT JOIN dbo.Fact_ComponentInstall prev
           ON prev.AirframeKey      = ci.AirframeKey
          AND prev.ComponentTypeKey = ci.ComponentTypeKey
          AND prev.PositionNo       = ci.PositionNo
          AND prev.RemovedDateKey   = ci.InstalledDateKey
),
Accrued AS (
    SELECT
        ci.InstallKey,
        Sorties     = COUNT(*),
        FlightHours = CAST(SUM(ss.FlightHours) AS DECIMAL(11,4)),
        StressHours = CAST(SUM(ss.StressHours) AS DECIMAL(11,4)),
        Landings    = SUM(ss.Landings)
    FROM dbo.Fact_ComponentInstall ci
    JOIN Windows w ON w.InstallKey = ci.InstallKey
    JOIN dbo.vw_SortieStress ss
      ON ss.AirframeKey = ci.AirframeKey
     AND ss.SortieDate >= w.EffectiveStartDate
     AND ss.SortieDate <= w.EndDate
    GROUP BY ci.InstallKey
)
SELECT
    ci.InstallKey,
    ci.ComponentSerial,
    ct.ComponentCode,
    ct.ComponentName,
    ct.Criticality,
    ct.LeadTimeFlightHours,
    a.AirframeKey, a.TailNumber, a.ModelCode,
    b.BaseCode, b.BaseName,
    ci.PositionNo,
    InstalledDate = di.[Date],
    RemovedDate   = dr.[Date],
    ci.RemovalReason,
    DaysFitted  = DATEDIFF(DAY, di.[Date], dr.[Date]),
    Sorties     = ISNULL(ac.Sorties, 0),
    Landings    = ISNULL(ac.Landings, 0),
    FlightHours = CAST(ISNULL(ac.FlightHours, 0) AS DECIMAL(11,4)),
    StressHours = CAST(ISNULL(ac.StressHours, 0) AS DECIMAL(11,4)),
    StressRatio = CAST(ISNULL(ac.StressHours, 0) / NULLIF(ac.FlightHours, 0) AS DECIMAL(9,4)),
    si.IntervalFlightHours,
    si.IntervalStressHours,
    si.WeibullShape,
    PctOfHourInterval   = CAST(100.0 * ISNULL(ac.FlightHours, 0) / si.IntervalFlightHours AS DECIMAL(9,2)),
    PctOfStressInterval = CAST(100.0 * ISNULL(ac.StressHours, 0) / si.IntervalStressHours AS DECIMAL(9,2)),
    -- 1 = the component failed and the life is observed exactly.
    -- 0 = removed serviceable; the true life is longer than what we recorded.
    IsFailure  = CAST(CASE WHEN ci.RemovalReason = 'Failure' THEN 1 ELSE 0 END AS BIT),
    IsCensored = CAST(CASE WHEN ci.RemovalReason = 'Failure' THEN 0 ELSE 1 END AS BIT),
    /*
    IsDateValid -- exposed, not filtered out.

    Six installs in this dataset were removed before they were installed
    (DQ-05). Their life is not short, it is NEGATIVE: DaysFitted -9, zero flight
    hours, NULL stress ratio. The view keeps them, because a reliability view
    that quietly drops rows is a view whose totals nobody can reconcile against
    the fact table.

    But nothing may SCORE them. Five of the six were inside the published alert
    evaluation, contributing a life with no measurable length to a prediction
    quality figure -- fn_AlertEvaluation now excludes them on this flag.
    */
    IsDateValid = CAST(CASE WHEN dr.[Date] >= di.[Date] THEN 1 ELSE 0 END AS BIT),
    ev.EventType,
    ev.DowntimeHours
FROM dbo.Fact_ComponentInstall ci
JOIN dbo.Dim_ComponentType ct   ON ct.ComponentTypeKey = ci.ComponentTypeKey
JOIN dbo.Ref_ServiceInterval si ON si.ComponentCode = ct.ComponentCode
JOIN dbo.Dim_Airframe a         ON a.AirframeKey = ci.AirframeKey
JOIN dbo.Dim_Base b             ON b.BaseKey = a.BaseKey
JOIN dbo.Dim_Date di            ON di.DateKey = ci.InstalledDateKey
JOIN dbo.Dim_Date dr            ON dr.DateKey = ci.RemovedDateKey
LEFT JOIN Accrued ac            ON ac.InstallKey = ci.InstallKey
LEFT JOIN dbo.Fact_MaintenanceEvent ev ON ev.InstallKey = ci.InstallKey
WHERE ci.RemovedDateKey IS NOT NULL;
GO

-- =============================================================================
-- vw_SensorFeatures -- the per-sortie condition-monitoring record.
--
-- One row per monitored component per sortie, with the rolling features a
-- degradation model actually uses. Rolling windows are computed HERE rather
-- than in Python so that the SQL and the model cannot disagree about what a
-- feature means -- the reconciliation problem Project 4 hit between Excel and
-- Power BI, moved one layer earlier.
--
-- StressHoursSoFar is the exposure variable: it is what the component has
-- actually lived through, and it is the x-axis every degradation curve here is
-- drawn against.
-- =============================================================================
IF OBJECT_ID('dbo.vw_SensorFeatures', 'V') IS NOT NULL DROP VIEW dbo.vw_SensorFeatures;
GO
CREATE VIEW dbo.vw_SensorFeatures
AS
WITH Base AS (
    SELECT
        sr.ReadingKey,
        sr.InstallKey,
        sr.SortieKey,
        ci.ComponentSerial,
        ct.ComponentCode,
        ct.LeadTimeFlightHours,
        a.AirframeKey, a.TailNumber,
        b.BaseCode,
        ss.SortieDate,
        ss.FlightHours,
        ss.StressHours,
        sr.VibrationRms,
        sr.TempRiseC,
        sr.CurrentDrawA,
        ci.RemovalReason,
        RemovedDate = dr.[Date],
        SortieSeq = ROW_NUMBER() OVER (PARTITION BY sr.InstallKey ORDER BY ss.SortieDate, sr.SortieKey)
    FROM dbo.Fact_SensorReading sr
    JOIN dbo.Fact_ComponentInstall ci ON ci.InstallKey = sr.InstallKey
    JOIN dbo.Dim_ComponentType ct     ON ct.ComponentTypeKey = ci.ComponentTypeKey
    JOIN dbo.vw_SortieStress ss       ON ss.SortieKey = sr.SortieKey
    JOIN dbo.Dim_Airframe a           ON a.AirframeKey = ci.AirframeKey
    JOIN dbo.Dim_Base b               ON b.BaseKey = a.BaseKey
    JOIN dbo.Dim_Date di              ON di.DateKey = ci.InstalledDateKey
    LEFT JOIN dbo.Dim_Date dr         ON dr.DateKey = ci.RemovedDateKey
    /*
    THE WINDOW PREDICATE THAT WAS MISSING.

    This view joins readings to installs on InstallKey alone. Without a date
    predicate it emitted 362 readings taken outside the component's fitted
    window -- the 13 orphans planted for DQ-03 plus 349 belonging to the six
    installs whose dates run backwards.

    Those rows reached the degradation model as training data and reached
    fn_AlertEvaluation as alerts, where three of the published 288 true
    positives were alerts raised against a part already off the aircraft. The
    data-quality layer had flagged every one of them; nothing downstream was
    reading the flags.

    Rows excluded here are NOT lost: DQ-03 and DQ-05 report them by count and by
    serial. They are excluded from the feature set because a reading taken while
    a component was in a crate is not evidence about that component.
    */
    WHERE ss.SortieDate >= di.[Date]
      AND (dr.[Date] IS NULL OR ss.SortieDate <= dr.[Date])
      AND (dr.[Date] IS NULL OR dr.[Date] >= di.[Date])
)
SELECT
    b.ReadingKey, b.InstallKey, b.SortieKey, b.ComponentSerial, b.ComponentCode,
    b.LeadTimeFlightHours, b.AirframeKey, b.TailNumber, b.BaseCode,
    b.SortieDate, b.SortieSeq,
    b.FlightHours, b.StressHours,
    b.VibrationRms, b.TempRiseC, b.CurrentDrawA,

    -- exposure to date on this component instance
    FlightHoursSoFar = CAST(SUM(b.FlightHours) OVER (PARTITION BY b.InstallKey
                            ORDER BY b.SortieSeq ROWS UNBOUNDED PRECEDING) AS DECIMAL(11,4)),
    StressHoursSoFar = CAST(SUM(b.StressHours) OVER (PARTITION BY b.InstallKey
                            ORDER BY b.SortieSeq ROWS UNBOUNDED PRECEDING) AS DECIMAL(11,4)),

    -- Rolling means over the last 10 sorties. A single reading is mostly noise;
    -- the trend is the signal, and 10 sorties is roughly a fortnight.
    VibRoll10  = CAST(AVG(b.VibrationRms) OVER (PARTITION BY b.InstallKey
                      ORDER BY b.SortieSeq ROWS BETWEEN 9 PRECEDING AND CURRENT ROW) AS DECIMAL(9,4)),
    TempRoll10 = CAST(AVG(b.TempRiseC) OVER (PARTITION BY b.InstallKey
                      ORDER BY b.SortieSeq ROWS BETWEEN 9 PRECEDING AND CURRENT ROW) AS DECIMAL(9,4)),
    CurrRoll10 = CAST(AVG(b.CurrentDrawA) OVER (PARTITION BY b.InstallKey
                      ORDER BY b.SortieSeq ROWS BETWEEN 9 PRECEDING AND CURRENT ROW) AS DECIMAL(9,4)),

    /*
    Ratio to the component's OWN baseline -- its first ten sorties.

    Absolute vibration varies between individual units for reasons that have
    nothing to do with wear: build tolerance, sensor mounting, the airframe it
    sits on. A threshold on the absolute value therefore flags the noisy units
    and misses the quiet ones. Every unit is its own control here, which is
    what makes one threshold work across the fleet.
    */
    VibVsBaseline = CAST(b.VibrationRms / NULLIF(
        AVG(CASE WHEN b.SortieSeq <= 10 THEN b.VibrationRms END) OVER (PARTITION BY b.InstallKey), 0) AS DECIMAL(9,4)),
    TempVsBaseline = CAST(b.TempRiseC / NULLIF(
        AVG(CASE WHEN b.SortieSeq <= 10 THEN b.TempRiseC END) OVER (PARTITION BY b.InstallKey), 0) AS DECIMAL(9,4)),

    -- the label, and the horizon it is measured over
    b.RemovalReason,
    b.RemovedDate,
    IsFailureInstall = CAST(CASE WHEN b.RemovalReason = 'Failure' THEN 1 ELSE 0 END AS BIT),
    /*
    FlightHoursToRemoval -- how much flying was left after this reading.

    This is the column the second finding turns on. A model that fires when this
    is 4 is accurate and useless: a rotor assembly takes 18 flight hours to
    obtain and fit, so the aircraft is grounded either way. Accuracy is measured
    against the label; usefulness is measured against this.
    */
    FlightHoursToRemoval = CAST(
        SUM(b.FlightHours) OVER (PARTITION BY b.InstallKey)
      - SUM(b.FlightHours) OVER (PARTITION BY b.InstallKey ORDER BY b.SortieSeq ROWS UNBOUNDED PRECEDING)
        AS DECIMAL(11,4)),
    StressHoursToRemoval = CAST(
        SUM(b.StressHours) OVER (PARTITION BY b.InstallKey)
      - SUM(b.StressHours) OVER (PARTITION BY b.InstallKey ORDER BY b.SortieSeq ROWS UNBOUNDED PRECEDING)
        AS DECIMAL(11,4))
FROM Base b;
GO

PRINT 'Core views created: vw_SortieStress, fn_ComponentWear, fn_AirframeStress, vw_ComponentLifeHistory, vw_SensorFeatures.';
GO

-- =============================================================================
-- Creation gate. A view that compiles is not a view that works.
-- =============================================================================
DECLARE @AsOf DATE = '2026-09-30';
DECLARE @Fitted INT, @Airframes INT, @Lives INT, @Features INT;

SELECT @Fitted    = COUNT(*) FROM dbo.fn_ComponentWear(@AsOf);
SELECT @Airframes = COUNT(*) FROM dbo.fn_AirframeStress(@AsOf);
SELECT @Lives     = COUNT(*) FROM dbo.vw_ComponentLifeHistory;
SELECT @Features  = COUNT(*) FROM dbo.vw_SensorFeatures;

IF @Fitted = 0    THROW 53020, 'fn_ComponentWear returned no fitted components. The as-of window or the install dates are wrong.', 1;
IF @Airframes = 0 THROW 53021, 'fn_AirframeStress returned no airframes.', 1;
IF @Lives = 0     THROW 53022, 'vw_ComponentLifeHistory returned no completed lives.', 1;
IF @Features = 0  THROW 53023, 'vw_SensorFeatures returned no rows. Sensor readings are not joining to installs.', 1;

/*
The gate that matters most. Every active airframe must carry a full set of
13 components -- 4 rotors, 4 motors, 2 batteries, 1 gimbal, 1 air data module,
1 landing gear. An airframe missing a position is an airframe whose compliance
percentage is computed over the wrong denominator, and it would read as
BETTER than a complete one because the missing part cannot be overdue.

A count that is merely non-zero would pass while half the fleet was missing.
*/
IF EXISTS (
    SELECT 1
    FROM dbo.fn_ComponentWear('2026-09-30') w
    CROSS JOIN (SELECT ExpectedPositions = SUM(PositionsPerAirframe) FROM dbo.Dim_ComponentType) e
    WHERE w.IsInFleetAsOf = 1
    GROUP BY w.AirframeKey, e.ExpectedPositions
    HAVING COUNT(*) <> MAX(e.ExpectedPositions)
)
    THROW 53024, 'An airframe in the fleet at the as-of date is not carrying a full set of components. Compliance percentages would be computed over an incomplete denominator.', 1;

/*
The gate for the handover rule.

No sortie may be charged to two components on the same position. This is the
defect that inflated 1,144 sorties, and it is invisible in any total -- the
accrued hours simply come out slightly high on every replacement, which reads
as an aircraft that has flown a little more than it has.
*/
IF EXISTS (
    SELECT 1
    FROM dbo.Fact_ComponentInstall a
    JOIN dbo.Fact_ComponentInstall b
      ON b.AirframeKey      = a.AirframeKey
     AND b.ComponentTypeKey = a.ComponentTypeKey
     AND b.PositionNo       = a.PositionNo
     AND b.InstallKey       <> a.InstallKey
    JOIN dbo.fn_ComponentWear('2026-09-30') wa ON wa.InstallKey = a.InstallKey
    JOIN dbo.fn_ComponentWear('2026-09-30') wb ON wb.InstallKey = b.InstallKey
    -- two installs on one position both claiming sorties on the same day
    WHERE wa.Sorties > 0 AND wb.Sorties > 0
      AND a.RemovedDateKey = b.InstalledDateKey
      AND EXISTS (SELECT 1 FROM dbo.vw_SortieStress ss
                  WHERE ss.AirframeKey = a.AirframeKey
                    AND ss.SortieDateKey = a.RemovedDateKey)
      AND wb.InstalledDate = wa.InstalledDate
)
    THROW 53025, 'A handover sortie is being charged to two components at once.', 1;

/*
The gate for the sensor window.

Every row in vw_SensorFeatures must fall inside its component's fitted window.
362 did not, and they reached both the model and the published prediction
figures.
*/
IF EXISTS (
    SELECT 1
    FROM dbo.vw_SensorFeatures f
    JOIN dbo.Fact_ComponentInstall ci ON ci.InstallKey = f.InstallKey
    JOIN dbo.Dim_Date di              ON di.DateKey = ci.InstalledDateKey
    LEFT JOIN dbo.Dim_Date dr         ON dr.DateKey = ci.RemovedDateKey
    WHERE f.SortieDate < di.[Date]
       OR (dr.[Date] IS NOT NULL AND f.SortieDate > dr.[Date])
)
    THROW 53026, 'vw_SensorFeatures contains readings taken outside the component fitted window. They would train the model and raise alerts against parts that were off the aircraft.', 1;

PRINT 'Core view gate passed.';
GO
