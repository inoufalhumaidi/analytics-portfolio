/*
================================================================================
Project 5 — Meridian UAV Services: Predictive Maintenance
Script:  04_data_quality_checks.sql
Purpose: Find the defects by BEHAVIOUR, then decide whether the answer survives.

TWO PRINCIPLES

1.  NO CHECK LOOKS FOR A MARKER.
    The generator plants its defects by making the data behave wrongly -- a
    sortie recorded twice, a reading against a component that was in a crate at
    the time. Nothing is tagged. A check that searched for a flag would prove
    only that the flag exists, and would find nothing at all in real extracts
    where no one labelled anything.

2.  THE GATE MEASURES DECISION IMPACT, NOT UNTIDINESS.
    A defect count answers 'how messy is this data'. Nobody makes a decision
    with that. The gate here asks a different question:

        What share of FLIGHT-CRITICAL components fitted right now have a wear
        figure that rests on flagged data?

    A duplicated sortie from March 2024 is a real defect and it is also
    completely irrelevant, because the component it inflated was replaced twice
    since. Counting it would inflate the alarm and, worse, would make a genuine
    problem look like the same size as a harmless one.

CHECKS THAT FIND NOTHING ARE STILL REPORTED.
    A summary that lists only what it found cannot distinguish 'we looked and it
    was clean' from 'we never looked'. Every check below reports its count
    whether or not that count is zero.

DATA DISCLOSURE
    Meridian UAV Services is fictional and all data is synthetic. No
    confidential data is used and no claim is made about any production system.
================================================================================
*/

USE MeridianUAV;
GO
SET NOCOUNT ON;
GO

IF OBJECT_ID('dbo.DQ_Findings', 'U') IS NOT NULL DROP TABLE dbo.DQ_Findings;
GO
CREATE TABLE dbo.DQ_Findings (
    FindingKey  INT IDENTITY(1,1) PRIMARY KEY,
    CheckCode   VARCHAR(10)  NOT NULL,
    CheckName   VARCHAR(80)  NOT NULL,
    Severity    VARCHAR(10)  NOT NULL CHECK (Severity IN ('High','Medium','Low')),
    EntityType  VARCHAR(30)  NOT NULL,
    EntityKey   INT          NULL,
    EntityRef   VARCHAR(40)  NULL,
    AirframeKey INT          NULL,
    [Detail]    VARCHAR(300) NOT NULL,
    -- Does this defect change a number someone acts on today? Set per finding,
    -- not per check: the same defect class can be decision-relevant on one row
    -- and irrelevant on another.
    AffectsCurrentWear BIT NOT NULL DEFAULT 0
);
GO

-- =============================================================================
-- DQ-01  Duplicate sorties
--
-- Same airframe, same date, same profile, same duration, same landings. A
-- second flight identical to the first to the minute is a double-entry, not a
-- coincidence -- and it inflates BOTH the hour meter and the stress total,
-- which is why it matters here rather than being merely untidy.
-- =============================================================================
INSERT INTO dbo.DQ_Findings (CheckCode, CheckName, Severity, EntityType, EntityKey, EntityRef, AirframeKey, [Detail])
SELECT 'DQ-01', 'Duplicate sortie record', 'High', 'Sortie', s.SortieKey, s.SortieID, s.AirframeKey,
       CONCAT('Airframe ', a.TailNumber, ' on ', CONVERT(CHAR(10), d.[Date], 126),
              ': ', dup.n, ' identical sorties (', s.FlightMinutes, ' min, ', s.Landings,
              ' landings, same profile). Flight and stress hours are double counted.')
FROM dbo.Fact_Sortie s
JOIN dbo.Dim_Airframe a ON a.AirframeKey = s.AirframeKey
JOIN dbo.Dim_Date d     ON d.DateKey = s.SortieDateKey
JOIN (
    SELECT AirframeKey, SortieDateKey, MissionProfileKey, FlightMinutes, Landings, n = COUNT(*)
    FROM dbo.Fact_Sortie
    GROUP BY AirframeKey, SortieDateKey, MissionProfileKey, FlightMinutes, Landings
    HAVING COUNT(*) > 1
) dup ON dup.AirframeKey = s.AirframeKey AND dup.SortieDateKey = s.SortieDateKey
     AND dup.MissionProfileKey = s.MissionProfileKey AND dup.FlightMinutes = s.FlightMinutes
     AND dup.Landings = s.Landings;
GO

-- =============================================================================
-- DQ-02  Flight time with no landings
--
-- An aircraft that took off has landed, one way or another. Zero landings
-- against real flight minutes is a recording failure, and it silently removes
-- the cycle term from the stress calculation -- which is the DOMINANT term for
-- training and urban work.
-- =============================================================================
INSERT INTO dbo.DQ_Findings (CheckCode, CheckName, Severity, EntityType, EntityKey, EntityRef, AirframeKey, [Detail])
SELECT 'DQ-02', 'Flight time recorded with zero landings', 'High', 'Sortie', s.SortieKey, s.SortieID, s.AirframeKey,
       CONCAT('Airframe ', a.TailNumber, ' on ', CONVERT(CHAR(10), d.[Date], 126), ': ',
              s.FlightMinutes, ' flight minutes and 0 landings. Cycle stress is understated for this sortie.')
FROM dbo.Fact_Sortie s
JOIN dbo.Dim_Airframe a ON a.AirframeKey = s.AirframeKey
JOIN dbo.Dim_Date d     ON d.DateKey = s.SortieDateKey
WHERE s.Landings = 0 AND s.FlightMinutes > 0;
GO

-- =============================================================================
-- DQ-03  Sensor reading against a component that was not fitted
--
-- Found by comparing the reading's sortie date against the install's fitted
-- window -- no flag, no marker. These readings carry a plausible vibration
-- figure, so nothing about the value itself gives them away; only the
-- relationship between two tables does.
--
-- CASCADE SUPPRESSION. Installs already flagged by DQ-05 are excluded.
--
-- When an install's own dates run backwards, its fitted window is not a window
-- at all, so EVERY reading on it falls outside and the check fires once per
-- reading. Six malformed date fields produced 357 extra findings here -- the
-- check was right each time and the total was still nonsense, because it
-- reported one root cause 358 times and buried the 13 genuine orphan readings
-- inside it.
--
-- Reporting a consequence as though it were an independent finding is how a
-- data-quality summary loses the reader: the list looks catastrophic, the fix
-- list has six items on it, and nobody can tell which from the report.
-- =============================================================================
INSERT INTO dbo.DQ_Findings (CheckCode, CheckName, Severity, EntityType, EntityKey, EntityRef, AirframeKey, [Detail])
SELECT 'DQ-03', 'Sensor reading outside the component fitted window', 'High', 'SensorReading',
       sr.ReadingKey, ci.ComponentSerial, ci.AirframeKey,
       CONCAT('Reading on ', CONVERT(CHAR(10), ds.[Date], 126), ' against ', ci.ComponentSerial,
              ', fitted ', CONVERT(CHAR(10), di.[Date], 126),
              ISNULL(' and removed ' + CONVERT(CHAR(10), dr.[Date], 126), ' and still fitted'),
              '. The component was not on the aircraft.')
FROM dbo.Fact_SensorReading sr
JOIN dbo.Fact_ComponentInstall ci ON ci.InstallKey = sr.InstallKey
JOIN dbo.Fact_Sortie s            ON s.SortieKey = sr.SortieKey
JOIN dbo.Dim_Date ds              ON ds.DateKey = s.SortieDateKey
JOIN dbo.Dim_Date di              ON di.DateKey = ci.InstalledDateKey
LEFT JOIN dbo.Dim_Date dr         ON dr.DateKey = ci.RemovedDateKey
WHERE (ds.[Date] < di.[Date]
    OR (dr.[Date] IS NOT NULL AND ds.[Date] > dr.[Date]))
  -- the cascade suppression described above
  AND NOT (dr.[Date] IS NOT NULL AND dr.[Date] < di.[Date]);
GO

-- =============================================================================
-- DQ-04  Maintenance event dated before the component was installed
-- =============================================================================
INSERT INTO dbo.DQ_Findings (CheckCode, CheckName, Severity, EntityType, EntityKey, EntityRef, AirframeKey, [Detail])
SELECT 'DQ-04', 'Maintenance event precedes component installation', 'Medium', 'MaintenanceEvent',
       e.EventKey, e.EventID, e.AirframeKey,
       CONCAT('Event ', e.EventID, ' dated ', CONVERT(CHAR(10), de.[Date], 126),
              ' against ', ci.ComponentSerial, ' installed ', CONVERT(CHAR(10), di.[Date], 126),
              '. Downtime is attributed to a component that was not yet fitted.')
FROM dbo.Fact_MaintenanceEvent e
JOIN dbo.Fact_ComponentInstall ci ON ci.InstallKey = e.InstallKey
JOIN dbo.Dim_Date de              ON de.DateKey = e.EventDateKey
JOIN dbo.Dim_Date di              ON di.DateKey = ci.InstalledDateKey
WHERE de.[Date] < di.[Date];
GO

-- =============================================================================
-- DQ-05  Component removed before it was installed
-- =============================================================================
INSERT INTO dbo.DQ_Findings (CheckCode, CheckName, Severity, EntityType, EntityKey, EntityRef, AirframeKey, [Detail])
SELECT 'DQ-05', 'Component removed before it was installed', 'High', 'ComponentInstall',
       ci.InstallKey, ci.ComponentSerial, ci.AirframeKey,
       CONCAT(ci.ComponentSerial, ' installed ', CONVERT(CHAR(10), di.[Date], 126),
              ', removed ', CONVERT(CHAR(10), dr.[Date], 126),
              '. Accrued life for this instance is meaningless.')
FROM dbo.Fact_ComponentInstall ci
JOIN dbo.Dim_Date di ON di.DateKey = ci.InstalledDateKey
JOIN dbo.Dim_Date dr ON dr.DateKey = ci.RemovedDateKey
WHERE dr.[Date] < di.[Date];
GO

/*
=============================================================================
 DQ-06  Two components fitted to the same position at the same time

 Not planted. It is here because it is the defect that would do the most damage
 if it occurred: every sortie would be counted against both components, and
 BOTH would look like they were wearing normally.

 THE PREDICATE LIVES IN A VIEW, AND THAT IS NOT TIDINESS.

 The first version wrote this self-join here and then wrote a second copy of it
 inside UAT-27 to prove the check could fire. The two were textually equivalent
 on the day they were written and had no dependency on each other at all: change
 the predicate here to a2.PositionNo = a1.PositionNo + 1 and DQ-06 silently
 finds nothing forever while UAT-27 goes on reporting PASS, because it is
 testing its own copy.

 That mattered more than usual here, because DQ-06 has never produced a row on
 the real dataset -- so a silently broken DQ-06 and a working one look identical
 in every artefact the suite inspects. One definition, two callers.
=============================================================================
*/
IF OBJECT_ID('dbo.vw_DQ_OverlappingInstalls', 'V') IS NOT NULL DROP VIEW dbo.vw_DQ_OverlappingInstalls;
GO
CREATE VIEW dbo.vw_DQ_OverlappingInstalls
AS
SELECT
    InstallKey      = a1.InstallKey,
    ComponentSerial = a1.ComponentSerial,
    OtherSerial     = a2.ComponentSerial,
    AirframeKey     = a1.AirframeKey,
    PositionNo      = a1.PositionNo
FROM dbo.Fact_ComponentInstall a1
JOIN dbo.Fact_ComponentInstall a2
  ON a2.AirframeKey = a1.AirframeKey AND a2.ComponentTypeKey = a1.ComponentTypeKey
 AND a2.PositionNo = a1.PositionNo   AND a2.InstallKey > a1.InstallKey
JOIN dbo.Dim_Date d1i ON d1i.DateKey = a1.InstalledDateKey
JOIN dbo.Dim_Date d2i ON d2i.DateKey = a2.InstalledDateKey
LEFT JOIN dbo.Dim_Date d1r ON d1r.DateKey = a1.RemovedDateKey
LEFT JOIN dbo.Dim_Date d2r ON d2r.DateKey = a2.RemovedDateKey
-- Strict overlap: each starts before the other ends. A handover on a shared
-- date is NOT an overlap -- one part comes off and the next goes on.
WHERE d1i.[Date] < ISNULL(d2r.[Date], '9999-12-31')
  AND d2i.[Date] < ISNULL(d1r.[Date], '9999-12-31');
GO

INSERT INTO dbo.DQ_Findings (CheckCode, CheckName, Severity, EntityType, EntityKey, EntityRef, AirframeKey, [Detail])
SELECT 'DQ-06', 'Overlapping installs on one position', 'High', 'ComponentInstall',
       o.InstallKey, o.ComponentSerial, o.AirframeKey,
       CONCAT(o.ComponentSerial, ' and ', o.OtherSerial, ' both fitted to ',
              af.TailNumber, ' position ', o.PositionNo, ' over an overlapping period.')
FROM dbo.vw_DQ_OverlappingInstalls o
JOIN dbo.Dim_Airframe af ON af.AirframeKey = o.AirframeKey;
GO

-- =============================================================================
-- DQ-07  Sortie flown by a retired airframe after its retirement date
-- =============================================================================
INSERT INTO dbo.DQ_Findings (CheckCode, CheckName, Severity, EntityType, EntityKey, EntityRef, AirframeKey, [Detail])
SELECT 'DQ-07', 'Sortie flown after airframe retirement', 'Medium', 'Sortie', s.SortieKey, s.SortieID, s.AirframeKey,
       CONCAT('Airframe ', a.TailNumber, ' retired ', CONVERT(CHAR(10), drt.[Date], 126),
              ' but flew on ', CONVERT(CHAR(10), ds.[Date], 126), '.')
FROM dbo.Fact_Sortie s
JOIN dbo.Dim_Airframe a ON a.AirframeKey = s.AirframeKey
JOIN dbo.Dim_Date ds    ON ds.DateKey = s.SortieDateKey
JOIN dbo.Dim_Date drt   ON drt.DateKey = a.RetiredDateKey
WHERE a.AirframeStatus = 'Retired' AND ds.[Date] > drt.[Date];
GO

-- =============================================================================
-- DQ-08  Component fitted before its airframe entered service
-- =============================================================================
INSERT INTO dbo.DQ_Findings (CheckCode, CheckName, Severity, EntityType, EntityKey, EntityRef, AirframeKey, [Detail])
SELECT 'DQ-08', 'Component fitted before airframe entered service', 'Medium', 'ComponentInstall',
       ci.InstallKey, ci.ComponentSerial, ci.AirframeKey,
       CONCAT(ci.ComponentSerial, ' fitted ', CONVERT(CHAR(10), di.[Date], 126), ' to ',
              a.TailNumber, ', in service from ', CONVERT(CHAR(10), dis.[Date], 126), '.')
FROM dbo.Fact_ComponentInstall ci
JOIN dbo.Dim_Airframe a ON a.AirframeKey = ci.AirframeKey
JOIN dbo.Dim_Date di    ON di.DateKey = ci.InstalledDateKey
JOIN dbo.Dim_Date dis   ON dis.DateKey = a.InServiceDateKey
WHERE di.[Date] < dis.[Date];
GO

-- =============================================================================
-- DQ-09  Implausible sensor values
--
-- Bounds come from the physical envelope, not from the observed distribution.
-- A bound fitted to the data cannot flag a sensor that has been wrong since the
-- day it was installed, because that sensor's readings ARE the distribution.
-- =============================================================================
INSERT INTO dbo.DQ_Findings (CheckCode, CheckName, Severity, EntityType, EntityKey, EntityRef, AirframeKey, [Detail])
SELECT 'DQ-09', 'Sensor reading outside the physical envelope', 'Medium', 'SensorReading',
       sr.ReadingKey, ci.ComponentSerial, ci.AirframeKey,
       CONCAT('Reading ', sr.ReadingKey, ': vibration ', sr.VibrationRms,
              ' g, temp rise ', sr.TempRiseC, ' C, current ', sr.CurrentDrawA, ' A.')
FROM dbo.Fact_SensorReading sr
JOIN dbo.Fact_ComponentInstall ci ON ci.InstallKey = sr.InstallKey
WHERE sr.VibrationRms <= 0 OR sr.VibrationRms > 25.0
   OR sr.TempRiseC    <  0 OR sr.TempRiseC    > 200.0
   OR sr.CurrentDrawA <= 0 OR sr.CurrentDrawA > 120.0;
GO

-- =============================================================================
-- DQ-10  Monitored component with no condition-monitoring coverage
--
-- A rotor assembly or motor that flew sorties but produced no readings is
-- invisible to the prediction model. Absence of alerts on it is not evidence of
-- health, and this is the check that stops a silent gap being read as a pass.
-- =============================================================================
INSERT INTO dbo.DQ_Findings (CheckCode, CheckName, Severity, EntityType, EntityKey, EntityRef, AirframeKey, [Detail])
SELECT 'DQ-10', 'Monitored component with no sensor coverage', 'Medium', 'ComponentInstall',
       ci.InstallKey, ci.ComponentSerial, ci.AirframeKey,
       CONCAT(ci.ComponentSerial, ' on ', a.TailNumber, ' flew ', sc.Sorties,
              ' sorties with no condition-monitoring readings. The model cannot see it.')
FROM dbo.Fact_ComponentInstall ci
JOIN dbo.Dim_ComponentType ct ON ct.ComponentTypeKey = ci.ComponentTypeKey
JOIN dbo.Dim_Airframe a       ON a.AirframeKey = ci.AirframeKey
JOIN dbo.Dim_Date di          ON di.DateKey = ci.InstalledDateKey
LEFT JOIN dbo.Dim_Date dr     ON dr.DateKey = ci.RemovedDateKey
CROSS APPLY (
    SELECT Sorties = COUNT(*)
    FROM dbo.vw_SortieStress ss
    WHERE ss.AirframeKey = ci.AirframeKey
      AND ss.SortieDate >= di.[Date]
      AND (dr.[Date] IS NULL OR ss.SortieDate <= dr.[Date])
) sc
WHERE ct.ComponentCode IN ('ROTOR-ASSY','MOTOR-ESC')
  AND sc.Sorties >= 10
  AND NOT EXISTS (SELECT 1 FROM dbo.Fact_SensorReading sr WHERE sr.InstallKey = ci.InstallKey);
GO

-- =============================================================================
-- DQ-11  Mission profile with no entry in the stress model
--
-- Constrained against at schema creation, so it should be impossible. Reported
-- anyway, because a profile with no stress row would accrue ZERO wear and every
-- airframe flying it would look immaculate -- the failure mode is invisible in
-- exactly the direction that flatters the fleet.
-- =============================================================================
INSERT INTO dbo.DQ_Findings (CheckCode, CheckName, Severity, EntityType, EntityKey, EntityRef, AirframeKey, [Detail])
SELECT 'DQ-11', 'Mission profile missing from the stress model', 'High', 'MissionProfile',
       p.MissionProfileKey, p.ProfileCode, NULL,
       CONCAT('Profile ', p.ProfileCode, ' has no row in Ref_StressModel. Sorties flown on it accrue no wear.')
FROM dbo.Dim_MissionProfile p
WHERE NOT EXISTS (SELECT 1 FROM dbo.Ref_StressModel sm WHERE sm.ProfileCode = p.ProfileCode);
GO

-- =============================================================================
-- DQ-12  Active airframe dormant for 90 days
--
-- Not a data error. It is a data-INTERPRETATION error waiting to happen: a
-- dormant airframe accrues no stress, so it reports as perfectly compliant and
-- drags every fleet average upward. Counting it as a healthy aircraft is how a
-- fleet average comes to describe aircraft that are not flying.
-- =============================================================================
INSERT INTO dbo.DQ_Findings (CheckCode, CheckName, Severity, EntityType, EntityKey, EntityRef, AirframeKey, [Detail])
SELECT 'DQ-12', 'Active airframe with no recent flying', 'Low', 'Airframe',
       a.AirframeKey, a.TailNumber, a.AirframeKey,
       CONCAT(a.TailNumber, ' is Active but last flew ',
              ISNULL(CONVERT(CHAR(10), ls.LastSortie, 126), 'never'),
              '. It accrues no stress and inflates fleet compliance.')
FROM dbo.Dim_Airframe a
OUTER APPLY (SELECT LastSortie = MAX(ss.SortieDate) FROM dbo.vw_SortieStress ss
             WHERE ss.AirframeKey = a.AirframeKey AND ss.SortieDate <= '2026-09-30') ls
JOIN dbo.Dim_Date isd ON isd.DateKey = a.InServiceDateKey
LEFT JOIN dbo.Dim_Date rtd ON rtd.DateKey = a.RetiredDateKey
-- in the fleet at the reporting date, from dates rather than from status
WHERE isd.[Date] <= '2026-09-30'
  AND (rtd.[Date] IS NULL OR rtd.[Date] > '2026-09-30')
  AND (ls.LastSortie IS NULL OR ls.LastSortie < DATEADD(DAY, -90, CAST('2026-09-30' AS DATE)));
GO

/*
==============================================================================
 Decide which findings actually change today's answer.

 THE DEFINITION THIS GATE USES, AND THE ONE IT REJECTED

 The first version of this gate marked a component affected if ANY flagged
 sortie was counted into its hours. It failed at 22.50% against a 3% tolerance
 and the number was meaningless: a duplicated sortie adds about 2.5 stress
 hours to a 150-hour interval, which is a 1.7% error, and it only changes an
 answer if the component happens to be sitting within 2.5 hours of its
 threshold. 'Touched by bad data' and 'would answer differently without the bad
 data' are different questions, and only the second is worth stopping a
 pipeline over.

 So the gate now compares an UNCERTAINTY BAND against a MARGIN.

 The uncertainty band is how much of a component's accrued stress we cannot
 stand behind:
   * DQ-01 duplicates -- the excess copies. Which copy is genuine is unknowable,
     so all but one in each group are treated as possibly not having happened.
   * DQ-02 zero-landing sorties -- the MISSING cycle stress, estimated at the
     average landings its profile actually records. This one points the other
     way: the recorded figure is too LOW.
   * DQ-07 post-retirement sorties -- the whole sortie.

 The margin is how far the component sits from its stress threshold.

 A verdict is UNSAFE when the band is wider than the margin -- when correcting
 the data could move the component across the line. That is a two-sided test:
 it does not matter whether the error flatters or maligns the component, only
 whether the conclusion depends on it.

 Install-level defects are unconditionally unsafe. An install whose dates run
 backwards has no computable life at all, so there is no margin to compare.
==============================================================================
*/
IF OBJECT_ID('tempdb..#SortieUncertainty') IS NOT NULL DROP TABLE #SortieUncertainty;

-- average landings per profile, computed over the sorties that are NOT flagged,
-- so the estimate for a missing value is not dragged down by the missing values
;WITH ProfileLandings AS (
    SELECT ss.ProfileCode, AvgLandings = AVG(CAST(ss.Landings AS DECIMAL(9,4)))
    FROM dbo.vw_SortieStress ss
    WHERE ss.Landings > 0
    GROUP BY ss.ProfileCode
),
DupExcess AS (
    -- all but one row in each duplicate group; which one is arbitrary and that
    -- is precisely the point -- the excess is what we cannot stand behind
    SELECT ss.SortieKey, UncertainStress = ss.StressHours
    FROM (
        SELECT s.SortieKey,
               rn = ROW_NUMBER() OVER (PARTITION BY s.AirframeKey, s.SortieDateKey, s.MissionProfileKey,
                                                    s.FlightMinutes, s.Landings
                                       ORDER BY s.SortieKey)
        FROM dbo.Fact_Sortie s
        WHERE EXISTS (SELECT 1 FROM dbo.DQ_Findings f
                      WHERE f.CheckCode = 'DQ-01' AND f.EntityKey = s.SortieKey)
    ) r
    JOIN dbo.vw_SortieStress ss ON ss.SortieKey = r.SortieKey
    WHERE r.rn > 1
),
MissingCycles AS (
    -- the cycle stress that SHOULD have been recorded on a zero-landing sortie
    SELECT ss.SortieKey,
           UncertainStress = CAST(pl.AvgLandings * sm.CyclePenaltyHours AS DECIMAL(9,4))
    FROM dbo.vw_SortieStress ss
    JOIN dbo.Ref_StressModel sm ON sm.ProfileCode = ss.ProfileCode
    JOIN ProfileLandings pl     ON pl.ProfileCode = ss.ProfileCode
    WHERE EXISTS (SELECT 1 FROM dbo.DQ_Findings f
                  WHERE f.CheckCode = 'DQ-02' AND f.EntityKey = ss.SortieKey)
),
PostRetirement AS (
    SELECT ss.SortieKey, UncertainStress = ss.StressHours
    FROM dbo.vw_SortieStress ss
    WHERE EXISTS (SELECT 1 FROM dbo.DQ_Findings f
                  WHERE f.CheckCode = 'DQ-07' AND f.EntityKey = ss.SortieKey)
)
SELECT SortieKey, UncertainStress = SUM(UncertainStress)
INTO #SortieUncertainty
FROM (
    SELECT * FROM DupExcess
    UNION ALL SELECT * FROM MissingCycles
    UNION ALL SELECT * FROM PostRetirement
) u
GROUP BY SortieKey;

IF OBJECT_ID('tempdb..#Affected') IS NOT NULL DROP TABLE #Affected;

;WITH FittedNow AS (
    SELECT w.InstallKey, w.AirframeKey, w.InstalledDate, w.Criticality, w.ComponentCode,
           w.StressHours, w.IntervalStressHours
    FROM dbo.fn_ComponentWear('2026-09-30') w
    WHERE w.IsInFleetAsOf = 1
),
Band AS (
    -- aggregate to InstallKey FIRST; joining the detail rows through and
    -- summing afterwards would multiply each component by its sortie count
    SELECT fn.InstallKey, UncertainStress = SUM(su.UncertainStress)
    FROM FittedNow fn
    JOIN dbo.vw_SortieStress ss ON ss.AirframeKey = fn.AirframeKey
                               AND ss.SortieDate >= fn.InstalledDate
                               AND ss.SortieDate <= '2026-09-30'
    JOIN #SortieUncertainty su  ON su.SortieKey = ss.SortieKey
    GROUP BY fn.InstallKey
),
FlaggedInstalls AS (
    SELECT DISTINCT f.EntityKey AS InstallKey
    FROM dbo.DQ_Findings f
    WHERE f.CheckCode IN ('DQ-05','DQ-08','DQ-10') AND f.EntityType = 'ComponentInstall'
    UNION
    SELECT DISTINCT sr.InstallKey
    FROM dbo.DQ_Findings f
    JOIN dbo.Fact_SensorReading sr ON sr.ReadingKey = f.EntityKey
    WHERE f.CheckCode = 'DQ-03' AND f.EntityType = 'SensorReading'
)
SELECT
    fn.InstallKey, fn.Criticality, fn.ComponentCode,
    UncertainStress = CAST(ISNULL(b.UncertainStress, 0) AS DECIMAL(11,4)),
    MarginStress    = CAST(ABS(fn.IntervalStressHours - fn.StressHours) AS DECIMAL(11,4)),
    [Reason] = CASE WHEN fi.InstallKey IS NOT NULL THEN 'Install record unusable'
                    ELSE 'Verdict within the uncertainty band' END
INTO #Affected
FROM FittedNow fn
LEFT JOIN Band b            ON b.InstallKey = fn.InstallKey
LEFT JOIN FlaggedInstalls fi ON fi.InstallKey = fn.InstallKey
WHERE fi.InstallKey IS NOT NULL
   OR ISNULL(b.UncertainStress, 0) >= ABS(fn.IntervalStressHours - fn.StressHours);

/*
AffectsCurrentWear -- a column that could never be set.

The first version updated only rows with EntityType = 'ComponentInstall' whose
EntityKey appeared in #Affected. #Affected holds components FITTED at the
reporting date; the install-level checks (DQ-05, DQ-08, DQ-10) fire on installs
that have been REMOVED. The two sets are disjoint by construction, so the flag
was 0 on all 73 findings, in every run, and nothing downstream ever noticed --
because a column that is always 0 reads exactly like good news.

It now marks a finding when the ENTITY IT NAMES touches a component whose wear
figure is in today's answer, whichever kind of entity that is:
  * an install-level finding on a currently fitted component;
  * a sortie-level finding on a sortie inside a currently fitted component's
    accrual window;
  * a reading-level finding on a currently fitted component.
*/
UPDATE f SET AffectsCurrentWear = 1
FROM dbo.DQ_Findings f
WHERE EXISTS (SELECT 1 FROM #Affected a WHERE a.InstallKey = f.EntityKey)
  AND f.EntityType IN ('ComponentInstall');

UPDATE f SET AffectsCurrentWear = 1
FROM dbo.DQ_Findings f
WHERE f.EntityType = 'SensorReading'
  AND EXISTS (SELECT 1 FROM dbo.Fact_SensorReading sr
              JOIN #Affected a ON a.InstallKey = sr.InstallKey
              WHERE sr.ReadingKey = f.EntityKey);

UPDATE f SET AffectsCurrentWear = 1
FROM dbo.DQ_Findings f
WHERE f.EntityType = 'Sortie'
  AND EXISTS (
      SELECT 1
      FROM dbo.fn_ComponentWear('2026-09-30') w
      JOIN dbo.vw_SortieStress ss ON ss.AirframeKey = w.AirframeKey
      WHERE w.IsInFleetAsOf = 1
        AND ss.SortieKey = f.EntityKey
        AND ss.SortieDate >= w.InstalledDate
        AND ss.SortieDate <= '2026-09-30');
GO

-- =============================================================================
-- Report
-- =============================================================================
PRINT '';
PRINT '=== Data quality: all checks, including those that found nothing ===';

WITH AllChecks AS (
    SELECT * FROM (VALUES
        ('DQ-01','Duplicate sortie record','High'),
        ('DQ-02','Flight time recorded with zero landings','High'),
        ('DQ-03','Sensor reading outside the component fitted window','High'),
        ('DQ-04','Maintenance event precedes component installation','Medium'),
        ('DQ-05','Component removed before it was installed','High'),
        ('DQ-06','Overlapping installs on one position','High'),
        ('DQ-07','Sortie flown after airframe retirement','Medium'),
        ('DQ-08','Component fitted before airframe entered service','Medium'),
        ('DQ-09','Sensor reading outside the physical envelope','Medium'),
        ('DQ-10','Monitored component with no sensor coverage','Medium'),
        ('DQ-11','Mission profile missing from the stress model','High'),
        ('DQ-12','Active airframe with no recent flying','Low')
    ) v(CheckCode, CheckName, Severity)
)
SELECT c.CheckCode, c.CheckName, c.Severity,
       Findings = ISNULL(f.n, 0),
       Result   = CASE WHEN ISNULL(f.n, 0) = 0 THEN 'clean' ELSE 'found' END
FROM AllChecks c
LEFT JOIN (SELECT CheckCode, n = COUNT(*) FROM dbo.DQ_Findings GROUP BY CheckCode) f
       ON f.CheckCode = c.CheckCode
ORDER BY c.CheckCode;
GO

-- =============================================================================
-- THE GATE
--
-- It RAISES. It does not print.
--
-- Every earlier project in this portfolio shipped a gate that printed 'FAIL'
-- and returned exit code 0, so `sqlcmd -b` carried on to the next script and
-- the pipeline went green on a dataset that had just failed its own check. A
-- gate that cannot stop anything is decoration.
-- =============================================================================
DECLARE @FlightCriticalFitted INT, @FlightCriticalAffected INT, @Pct DECIMAL(9,3);
DECLARE @Tolerance DECIMAL(9,3) = 3.000;

SELECT @FlightCriticalFitted = COUNT(*)
FROM dbo.fn_ComponentWear('2026-09-30')
WHERE IsInFleetAsOf = 1 AND Criticality = 'FlightCritical';

SELECT @FlightCriticalAffected = COUNT(*)
FROM #Affected WHERE Criticality = 'FlightCritical';

/*
An empty population is a FAILURE, not a pass.

NULLIF guards the division, and then NULL > 3.000 evaluates to UNKNOWN, the
THROW does not fire, and the gate prints 'Gate PASSED' over a fleet it could not
see. Every earlier project in this portfolio shipped a gate that could not stop
anything; this one could not even start.
*/
IF @FlightCriticalFitted IS NULL OR @FlightCriticalFitted = 0
    THROW 53031, 'QA gate found no flight-critical components fitted at the reporting date. There is nothing to gate, which is itself the failure: the wear model returned an empty fleet.', 1;

SET @Pct = CAST(100.0 * @FlightCriticalAffected / @FlightCriticalFitted AS DECIMAL(9,3));

PRINT '';
PRINT '=== QA gate: flight-critical wear figures resting on flagged data ===';
PRINT CONCAT('  Flight-critical components fitted : ', @FlightCriticalFitted);
PRINT CONCAT('  ...whose wear rests on flagged data: ', @FlightCriticalAffected);
PRINT CONCAT('  Share                              : ', CAST(@Pct AS VARCHAR(20)), '%');
PRINT CONCAT('  Tolerance                          : ', CAST(@Tolerance AS VARCHAR(20)), '%');

IF @Pct > @Tolerance
    THROW 53030, 'QA gate failed: too many flight-critical wear figures rest on flagged data. The compliance finding cannot be published at this tolerance.', 1;

-- PRINT takes scalar expressions only. A subquery inside CONCAT here is a
-- SYNTAX error, not a runtime one, so it aborts the batch before the gate ever
-- runs -- and with sqlcmd -b the pipeline stops on a data-quality script that
-- never reached its data-quality check.
DECLARE @AffectingNow INT = (SELECT COUNT(*) FROM dbo.DQ_Findings WHERE AffectsCurrentWear = 1);
DECLARE @AllFindings  INT = (SELECT COUNT(*) FROM dbo.DQ_Findings);
PRINT '';
PRINT CONCAT('  Findings whose entity is in today''s answer: ', @AffectingNow, ' of ', @AllFindings);
PRINT '  Gate PASSED.';
PRINT '';
GO
