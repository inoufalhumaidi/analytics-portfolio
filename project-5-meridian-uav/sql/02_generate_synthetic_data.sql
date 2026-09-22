/*
================================================================================
Project 5 — Meridian UAV Services: Predictive Maintenance
Script:  02_generate_synthetic_data.sql
Purpose: Generate the synthetic fleet history, deterministically.

DETERMINISM
    Every random draw comes from dbo.fn_Rand(<stable text key>), which hashes
    the key with SHA2_256 and scales it into [0,1). It is a pure function of the
    key, so dropping and regenerating reproduces the dataset byte for byte and
    every figure quoted in the case study stays checkable. RAND() and NEWID()
    cannot do this.

    THE KEYS ARE BUSINESS KEYS -- TailNumber, SortieID, ComponentSerial -- AND
    NEVER SURROGATE IDENTITY VALUES. This is not a style preference; the first
    version of this script keyed on AirframeKey and was not reproducible.

    INSERT ... SELECT assigns IDENTITY values in UNSPECIFIED order. The airframe
    rows are generated from sys.all_objects, a catalog view whose row order
    changes when the database's object list changes -- so adding one reference
    table to 01_create_schema.sql was enough to hand MU-017 a different
    AirframeKey. Because every draw was keyed on that surrogate, permuting the
    keys permuted which airframe flew on which day: the sortie count moved from
    10,346 to 10,288 and every fingerprint changed, with no edit to any
    generation logic at all.

    A business key cannot do that. 'MU-017' is 'MU-017' whatever the insert
    order was, so the draws it seeds are fixed before the database has decided
    anything.

HOW WEAR IS SIMULATED

    Wear is NOT drawn at random and it is NOT a function of calendar time. Each
    component instance is given a life in STRESS HOURS, drawn from a Weibull
    distribution with the shape its type actually exhibits, and it then accrues
    stress from the sorties its airframe happens to fly:

        stress hours = flight hours x profile multiplier
                     + landings     x cycle penalty
                     + flight hours x (payload - 12kg) x payload penalty

    Meanwhile the OPERATOR is running a different clock: it pulls the part at
    IntervalFlightHours, counted in logged flight hours. Every fitted component
    is therefore in a race between the two, and whichever clock finishes first
    decides whether the removal is 'Scheduled' or a 'Failure'. See section 6.

    This matters because it means the FINDING IS NOT PLANTED -- it emerges. An
    airframe flying mountain and cargo profiles accrues stress 2-2.4x faster
    than the hour meter suggests, so its parts reach the end of their physical
    life at around 60 flight hours against a 150-hour interval. They fail,
    on aircraft that every compliance report shows as inside its interval.
    Nobody wrote that in; it falls out of flying the fleet.

DELIBERATE BEHAVIOUR PLANTED HERE
    P1  Bases fly different profile mixes. Highland flies mountain and cargo;
        Coastal flies survey. Their airframes wear at very different rates per
        logged hour, and the hour meter cannot see it.
    P2  Sensor degradation is real but its LEAD TIME varies by component. Rotor
        bearings signal gradually and early; motor controllers stay quiet and
        then turn sharply. This is what makes prediction accuracy and
        prediction usefulness different numbers.
    P3  AIRDATA has a Weibull shape of ~1.05 -- effectively random failure. It
        is in the fleet deliberately as an honest negative: condition
        monitoring and scheduled replacement both earn almost nothing on it.
    P4  TRAINING accrues most of its damage from LANDINGS rather than airborne
        time, so the academy airframes look like the lightest-used fleet on the
        hour meter and are among the hardest worn once cycles are counted.
        NOTE: the cycle penalty is per PROFILE, not per component -- see the
        comment on Ref_StressModel in 01. Differences between component types
        come from their Weibull shapes, not from a landing-gear-specific term.
    P5  A handful of airframes carry components well past their stress interval
        while remaining inside the hour interval. These are the queue.

PLANTED DATA DEFECTS -- for the data-quality layer to find BY BEHAVIOUR
    D1  Duplicate sorties (same airframe, date, profile, duration, landings).
    D2  Sensor readings against a component not fitted on that date.
    D3  Maintenance events dated before the component was installed.
    D4  Sorties with flight time but zero landings.
    D5  Component installs removed before they were installed.

DATA DISCLOSURE
    Meridian UAV Services is fictional. Degradation behaviour is modelled on
    publicly described characteristics of rotating machinery; no proprietary
    reliability data is used and no claim is made about any real aircraft,
    operator or manufacturer.
================================================================================
*/

USE MeridianUAV;
GO
SET NOCOUNT ON;
GO

/*
=============================================================================
 0a. Make this script re-runnable.

 A generator that only works against an empty database is one you can never
 re-validate: the first failed run leaves a partial dataset behind and every
 later attempt fails on a primary-key violation instead of on whatever was
 actually wrong. Clearing first is what lets determinism be TESTED -- run it
 twice and the fingerprints in section 10 must match.

 DELETE in foreign-key order, then reseed the identity columns. The reseed is a
 convenience, not a guarantee: nothing in this project may DEPEND on a surrogate
 key having a particular value, because INSERT ... SELECT does not promise one.
 It simply makes two consecutive builds easier to diff by hand. What determinism
 actually rests on is that every draw is keyed on a BUSINESS key -- see the
 header.

 Dim_MissionProfile, Dim_ComponentType and the three Ref_ tables are seeded by
 01_create_schema.sql and are deliberately NOT touched -- they are reference
 data an engineer is meant to edit, not generated output.
=============================================================================
*/
DELETE FROM dbo.Fact_SensorReading;
DELETE FROM dbo.Fact_MaintenanceEvent;
DELETE FROM dbo.Fact_ComponentInstall;
DELETE FROM dbo.Fact_Sortie;
DELETE FROM dbo.Dim_Airframe;
DELETE FROM dbo.Dim_Base;
DELETE FROM dbo.Dim_Date;
GO
DBCC CHECKIDENT ('dbo.Fact_SensorReading',    RESEED, 0) WITH NO_INFOMSGS;
DBCC CHECKIDENT ('dbo.Fact_MaintenanceEvent', RESEED, 0) WITH NO_INFOMSGS;
DBCC CHECKIDENT ('dbo.Fact_ComponentInstall', RESEED, 0) WITH NO_INFOMSGS;
DBCC CHECKIDENT ('dbo.Fact_Sortie',           RESEED, 0) WITH NO_INFOMSGS;
DBCC CHECKIDENT ('dbo.Dim_Airframe',          RESEED, 0) WITH NO_INFOMSGS;
DBCC CHECKIDENT ('dbo.Dim_Base',              RESEED, 0) WITH NO_INFOMSGS;
GO

-- =============================================================================
-- 0. Deterministic pseudo-randomness
-- =============================================================================
IF OBJECT_ID('dbo.fn_Rand', 'FN') IS NOT NULL DROP FUNCTION dbo.fn_Rand;
GO
CREATE FUNCTION dbo.fn_Rand (@Key VARCHAR(200))
RETURNS DECIMAL(18,15)
AS
BEGIN
    DECLARE @h BINARY(32) = HASHBYTES('SHA2_256', @Key);
    DECLARE @n BIGINT =
          CAST(SUBSTRING(@h,1,1) AS BIGINT) * 1099511627776
        + CAST(SUBSTRING(@h,2,1) AS BIGINT) * 4294967296
        + CAST(SUBSTRING(@h,3,1) AS BIGINT) * 16777216
        + CAST(SUBSTRING(@h,4,1) AS BIGINT) * 65536
        + CAST(SUBSTRING(@h,5,1) AS BIGINT) * 256
        + CAST(SUBSTRING(@h,6,1) AS BIGINT);
    RETURN CAST(@n AS DECIMAL(38,15)) / CAST(281474976710656 AS DECIMAL(38,15));
END;
GO

/*
fn_WeibullLife -- an inverse-CDF Weibull draw.

    life = scale * (-ln(1-u)) ^ (1/shape)

Shape is what makes this worth doing rather than drawing uniformly. Above 1 the
hazard rises with age (wear-out) and lives cluster near the scale; at 1 it is
memoryless and a component is no more likely to fail at 300 hours than at 30.
The fleet deliberately contains both, because the right maintenance policy for
each is opposite and a single policy applied to both wastes money at one end
and causes failures at the other.
*/
IF OBJECT_ID('dbo.fn_WeibullLife', 'FN') IS NOT NULL DROP FUNCTION dbo.fn_WeibullLife;
GO
CREATE FUNCTION dbo.fn_WeibullLife (@Scale DECIMAL(9,3), @Shape DECIMAL(4,2), @U DECIMAL(18,15))
RETURNS DECIMAL(9,3)
AS
BEGIN
    -- guard the tails: u=0 gives life 0, u=1 is undefined
    -- Note: T-SQL variable names are case-insensitive, so this local cannot be
    -- called @u while the parameter is @U.
    DECLARE @p FLOAT = CASE WHEN @U < 0.0005 THEN 0.0005 WHEN @U > 0.9995 THEN 0.9995 ELSE CAST(@U AS FLOAT) END;
    RETURN CAST(@Scale * POWER(-LOG(1.0 - @p), 1.0 / @Shape) AS DECIMAL(9,3));
END;
GO

-- =============================================================================
-- 1. Dim_Date  (2023-10-01 .. 2027-03-31)
-- =============================================================================
;WITH N AS (
    SELECT TOP (1280) ROW_NUMBER() OVER (ORDER BY (SELECT NULL)) - 1 AS n
    FROM sys.all_objects a CROSS JOIN sys.all_objects b
),
D AS (SELECT [Date] = DATEADD(DAY, n, CAST('2023-10-01' AS DATE)) FROM N)
INSERT INTO dbo.Dim_Date (DateKey, [Date], [Year], [Quarter], [Month], MonthName,
                          YearMonth, YearQuarter, MonthEndDate, IsMonthEnd, IsWeekend)
SELECT
    YEAR([Date])*10000 + MONTH([Date])*100 + DAY([Date]),
    [Date], YEAR([Date]), DATEPART(QUARTER,[Date]), MONTH([Date]),
    -- NOT DATENAME(MONTH, ...): that returns the session language's month name,
    -- so the same script run by a German-locale login produces 'Marz' and every
    -- downstream label silently changes. The one session setting left
    -- unnormalised after DATEFIRST.
    CHOOSE(MONTH([Date]), 'January','February','March','April','May','June',
                          'July','August','September','October','November','December'),
    CONVERT(CHAR(7), [Date], 126),
    CAST(YEAR([Date]) AS CHAR(4)) + '-Q' + CAST(DATEPART(QUARTER,[Date]) AS CHAR(1)),
    EOMONTH([Date]),
    CASE WHEN [Date] = EOMONTH([Date]) THEN 1 ELSE 0 END,
    -- normalised so 6 = Saturday and 7 = Sunday whatever DATEFIRST is set to
    CASE WHEN ((DATEPART(WEEKDAY,[Date]) + @@DATEFIRST - 2) % 7) + 1 IN (6,7) THEN 1 ELSE 0 END
FROM D WHERE [Date] <= '2027-03-31';
GO

-- =============================================================================
-- 2. Dim_Base
--
-- The profile mix each base flies is the engine of the whole finding. Highland
-- and Port are not worse-run bases; they fly harder work.
-- =============================================================================
INSERT INTO dbo.Dim_Base (BaseCode, BaseName, Region, HangarBays) VALUES
 ('HIGHLND', 'Highland Operations',  'North Highlands', 3),
 ('COASTAL', 'Coastal Survey Base',  'West Coast',      4),
 ('PORT',    'Port Logistics Hub',   'East Estuary',    3),
 ('ACADEMY', 'Training Academy',     'Central Plain',   2);
GO

-- =============================================================================
-- 3. Dim_Airframe  (31 airframes, 3 of them retired during the window)
-- =============================================================================
/*
ORDER BY n on the insert below is belt-and-braces. It makes the AirframeKey
assignment follow tail-number order, which is convenient to read -- but nothing
DEPENDS on it, because every draw downstream is keyed on TailNumber. That is
the actual guarantee; the ORDER BY is a courtesy to whoever reads the table.
*/
;WITH N AS (SELECT TOP (31) ROW_NUMBER() OVER (ORDER BY (SELECT NULL)) AS n
            FROM sys.all_objects a CROSS JOIN sys.all_objects b)
INSERT INTO dbo.Dim_Airframe (TailNumber, ModelCode, BaseKey, InServiceDateKey, AirframeStatus, RetiredDateKey)
SELECT
    TailNumber = 'MU-' + RIGHT('000' + CAST(n AS VARCHAR(3)), 3),
    ModelCode  = CASE WHEN r.rModel < 0.55 THEN 'MRD-8X' WHEN r.rModel < 0.85 THEN 'MRD-12H' ELSE 'MRD-6L' END,
    BaseKey    = b.BaseKey,
    InServiceDateKey = isd.DateKey,
    AirframeStatus = CASE WHEN r.rRetire < 0.097 THEN 'Retired' ELSE 'Active' END,
    RetiredDateKey = CASE WHEN r.rRetire < 0.097 THEN rtd.DateKey END
FROM N
CROSS APPLY (SELECT rModel  = dbo.fn_Rand(CONCAT('af|model|',  n)),
                    rBase   = dbo.fn_Rand(CONCAT('af|base|',   n)),
                    rInSvc  = dbo.fn_Rand(CONCAT('af|insvc|',  n)),
                    rRetire = dbo.fn_Rand(CONCAT('af|retire|', n))) r
CROSS APPLY (SELECT BaseCode = CASE WHEN r.rBase < 0.29 THEN 'HIGHLND'
                                    WHEN r.rBase < 0.58 THEN 'COASTAL'
                                    WHEN r.rBase < 0.81 THEN 'PORT'
                                    ELSE 'ACADEMY' END) bc
JOIN dbo.Dim_Base b ON b.BaseCode = bc.BaseCode
-- Airframes enter service staggered across the first eight months, so the
-- fleet does not all reach its intervals on the same day.
CROSS APPLY (SELECT d = DATEADD(DAY, CAST(r.rInSvc * 240 AS INT), CAST('2024-01-01' AS DATE))) i
JOIN dbo.Dim_Date isd ON isd.[Date] = i.d
LEFT JOIN dbo.Dim_Date rtd ON rtd.[Date] = DATEADD(DAY, 430 + CAST(r.rRetire * 2000 AS INT), i.d)
ORDER BY n;
GO

-- =============================================================================
-- 4. Fact_Sortie
--
-- Each airframe flies on roughly 45% of days from entering service to the
-- earlier of retirement and the reporting date. The profile is drawn from its
-- BASE's mix, which is what makes the wear rates diverge.
-- =============================================================================
IF OBJECT_ID('tempdb..#Flyable') IS NOT NULL DROP TABLE #Flyable;

SELECT a.AirframeKey, a.TailNumber, b.BaseCode, d.DateKey, d.[Date],
       DayNo = ROW_NUMBER() OVER (PARTITION BY a.AirframeKey ORDER BY d.[Date])
INTO #Flyable
FROM dbo.Dim_Airframe a
JOIN dbo.Dim_Base b   ON b.BaseKey = a.BaseKey
JOIN dbo.Dim_Date isd ON isd.DateKey = a.InServiceDateKey
LEFT JOIN dbo.Dim_Date rtd ON rtd.DateKey = a.RetiredDateKey
JOIN dbo.Dim_Date d ON d.[Date] >= isd.[Date]
                   AND d.[Date] <= CASE WHEN rtd.[Date] IS NULL OR rtd.[Date] > '2026-09-30'
                                        THEN '2026-09-30' ELSE rtd.[Date] END;

CREATE INDEX IX_Flyable ON #Flyable (AirframeKey, DateKey);

INSERT INTO dbo.Fact_Sortie (SortieID, AirframeKey, MissionProfileKey, SortieDateKey,
                             FlightMinutes, PayloadKg, Landings, AmbientTempC, WindGustKts)
SELECT
    -- SortieID is ordered by TAIL NUMBER, not AirframeKey, so the same sortie
    -- carries the same ID no matter what surrogate keys the database handed out.
    SortieID = 'S-' + RIGHT('00000000' + CAST(ROW_NUMBER() OVER (ORDER BY f.TailNumber, f.DateKey) AS VARCHAR(8)), 8),
    f.AirframeKey,
    p.MissionProfileKey,
    f.DateKey,
    FlightMinutes = CASE pc.ProfileCode
        WHEN 'TRAINING' THEN 35 + CAST(dbo.fn_Rand(CONCAT('so|min|', f.TailNumber, '|', f.DateKey)) * 40 AS INT)
        WHEN 'URBAN'    THEN 30 + CAST(dbo.fn_Rand(CONCAT('so|min|', f.TailNumber, '|', f.DateKey)) * 55 AS INT)
        WHEN 'CARGO'    THEN 80 + CAST(dbo.fn_Rand(CONCAT('so|min|', f.TailNumber, '|', f.DateKey)) * 95 AS INT)
        ELSE                 55 + CAST(dbo.fn_Rand(CONCAT('so|min|', f.TailNumber, '|', f.DateKey)) * 85 AS INT) END,
    PayloadKg = CAST(mp.TypicalPayloadKg
                   + (dbo.fn_Rand(CONCAT('so|pay|', f.TailNumber, '|', f.DateKey)) - 0.5) * 6.0 AS DECIMAL(5,1)),
    -- Training flies circuits: far more landings per hour than anything else.
    Landings = CASE pc.ProfileCode
        WHEN 'TRAINING' THEN 6 + CAST(dbo.fn_Rand(CONCAT('so|ldg|', f.TailNumber, '|', f.DateKey)) * 9 AS INT)
        WHEN 'URBAN'    THEN 3 + CAST(dbo.fn_Rand(CONCAT('so|ldg|', f.TailNumber, '|', f.DateKey)) * 5 AS INT)
        WHEN 'SAR'      THEN 2 + CAST(dbo.fn_Rand(CONCAT('so|ldg|', f.TailNumber, '|', f.DateKey)) * 4 AS INT)
        ELSE                 1 + CAST(dbo.fn_Rand(CONCAT('so|ldg|', f.TailNumber, '|', f.DateKey)) * 2 AS INT) END,
    AmbientTempC = CAST(4.0 + 16.0 * SIN((DATEPART(DAYOFYEAR, f.[Date]) - 100) * 0.0172)
                      + (dbo.fn_Rand(CONCAT('so|tmp|', f.TailNumber, '|', f.DateKey)) - 0.5) * 9.0 AS DECIMAL(4,1)),
    WindGustKts = CAST(5.0 + dbo.fn_Rand(CONCAT('so|wnd|', f.TailNumber, '|', f.DateKey)) * 26.0 AS DECIMAL(4,1))
FROM #Flyable f
CROSS APPLY (SELECT rFly = dbo.fn_Rand(CONCAT('so|fly|', f.TailNumber, '|', f.DateKey)),
                    rPro = dbo.fn_Rand(CONCAT('so|pro|', f.TailNumber, '|', f.DateKey))) r
-- Profile mix BY BASE. This is planted pattern P1 and the engine of the finding.
CROSS APPLY (SELECT ProfileCode = CASE f.BaseCode
    WHEN 'HIGHLND' THEN CASE WHEN r.rPro < 0.46 THEN 'MOUNTAIN' WHEN r.rPro < 0.66 THEN 'SAR'
                             WHEN r.rPro < 0.84 THEN 'CARGO'    ELSE 'COASTAL' END
    -- Deliberately the benign end of the fleet. Coastal exists in the dataset
    -- as the control: a base whose work the 150-hour interval genuinely suits,
    -- so that the Highland figure has something to be compared against. Without
    -- it the finding degenerates into 'the whole fleet fails a lot'.
    WHEN 'COASTAL' THEN CASE WHEN r.rPro < 0.78 THEN 'COASTAL'  WHEN r.rPro < 0.92 THEN 'URBAN'
                             WHEN r.rPro < 0.98 THEN 'SAR'      ELSE 'MOUNTAIN' END
    WHEN 'PORT'    THEN CASE WHEN r.rPro < 0.52 THEN 'CARGO'    WHEN r.rPro < 0.80 THEN 'URBAN'
                             WHEN r.rPro < 0.93 THEN 'COASTAL'  ELSE 'SAR' END
    ELSE                CASE WHEN r.rPro < 0.72 THEN 'TRAINING' WHEN r.rPro < 0.88 THEN 'URBAN'
                             ELSE 'COASTAL' END END) pc
JOIN dbo.Dim_MissionProfile mp ON mp.ProfileCode = pc.ProfileCode
JOIN dbo.Dim_MissionProfile p  ON p.MissionProfileKey = mp.MissionProfileKey
WHERE r.rFly < 0.45
  -- DATEPART(WEEKDAY) depends on the session's DATEFIRST, which would make the
  -- flying calendar depend on the client's language settings. @@DATEFIRST is
  -- normalised out so 'Sunday' means Sunday on every connection.
  AND ((DATEPART(WEEKDAY, f.[Date]) + @@DATEFIRST - 2) % 7) + 1 <> 7;   -- 7 = Sunday
GO

-- =============================================================================
-- 5. The stress timeline
--
-- Cumulative stress hours per airframe, sortie by sortie. Everything about
-- component life is a lookup against this.
-- =============================================================================
IF OBJECT_ID('tempdb..#Stress') IS NOT NULL DROP TABLE #Stress;

SELECT
    s.SortieKey, s.AirframeKey, s.SortieDateKey,
    FlightHours = CAST(s.FlightMinutes / 60.0 AS DECIMAL(9,4)),
    s.Landings,
    StressHours = CAST(
          (s.FlightMinutes / 60.0) * sm.StressMultiplier
        + s.Landings * sm.CyclePenaltyHours
        + (s.FlightMinutes / 60.0) * (s.PayloadKg - 12.0) * sm.PayloadPenaltyPerKg
        AS DECIMAL(9,4)),
    SortieSeq = ROW_NUMBER() OVER (PARTITION BY s.AirframeKey ORDER BY s.SortieDateKey, s.SortieKey)
INTO #Stress
FROM dbo.Fact_Sortie s
JOIN dbo.Dim_MissionProfile p ON p.MissionProfileKey = s.MissionProfileKey
JOIN dbo.Ref_StressModel sm   ON sm.ProfileCode = p.ProfileCode;

ALTER TABLE #Stress ADD CumStress DECIMAL(12,4), CumFlight DECIMAL(12,4), CumLandings INT;

;WITH C AS (
    SELECT SortieKey,
           cs = SUM(StressHours) OVER (PARTITION BY AirframeKey ORDER BY SortieSeq ROWS UNBOUNDED PRECEDING),
           cf = SUM(FlightHours) OVER (PARTITION BY AirframeKey ORDER BY SortieSeq ROWS UNBOUNDED PRECEDING),
           cl = SUM(Landings)    OVER (PARTITION BY AirframeKey ORDER BY SortieSeq ROWS UNBOUNDED PRECEDING)
    FROM #Stress
)
UPDATE t SET CumStress = c.cs, CumFlight = c.cf, CumLandings = c.cl
FROM #Stress t JOIN C c ON c.SortieKey = t.SortieKey;

CREATE INDEX IX_Stress ON #Stress (AirframeKey, SortieSeq) INCLUDE (CumStress, CumFlight, SortieDateKey);
GO

-- =============================================================================
-- 6. Fact_ComponentInstall
--
-- THE RACE. Every fitted component is running two clocks at once:
--
--   the MAINTENANCE clock -- flight hours since fitting. The operator pulls the
--                            part when it reaches IntervalFlightHours.
--   the PHYSICS clock     -- stress hours since fitting. The part fails when it
--                            passes its Weibull-drawn life.
--
-- Whichever clock finishes first decides what happens, and every figure this
-- project reports falls out of that race:
--
--   * On a coastal airframe the two clocks run at nearly the same rate, so
--     scheduled replacement usually wins and unscheduled removals are rare.
--   * On a Highland airframe the physics clock runs at 2.15 stress hours per
--     flight hour, so 150 stress hours arrives at about 69 -- long before anyone
--     is booked to touch it. The physics clock wins, and the removal is a
--     failure. (No airframe reaches the MOUNTAIN profile's own 2.4 multiplier;
--     every airframe flies a MIX, and the base ratio is what wear responds to.)
--
-- Nothing here says "make Highland worse". Highland is worse because of the
-- work it flies, and an hour-based interval cannot see the difference.
--
-- The loop advances every position in the fleet by ONE GENERATION per pass, so
-- six passes cover the reporting window regardless of fleet size.
-- =============================================================================
IF OBJECT_ID('tempdb..#Pos') IS NOT NULL DROP TABLE #Pos;

SELECT a.AirframeKey, a.TailNumber, ct.ComponentTypeKey, ct.ComponentCode, p.PositionNo,
       si.IntervalFlightHours, si.IntervalStressHours, si.WeibullShape,
       /*
       The Weibull SCALE is derived, not chosen, and the derivation is the most
       contestable modelling decision in this script.

       An interval is not set at the average life -- setting it there means half
       the population fails before anyone touches it. It is set at the B10 life:
       the point by which 10% of the population has failed. Solving
       life = scale * (-ln(1-U))^(1/shape) for life = interval at U = 0.10 gives

           scale = interval / (-ln 0.90)^(1/shape)

       So IntervalStressHours means something precise: it is the stress-hour
       figure at which one part in ten has already failed.

       On a Coastal airframe stress accrues at 1.17x flight hours, so the 150
       FLIGHT hour interval is about 176 stress hours -- modestly past B10 --
       and 21.4% of its rotors fail before replacement. On a Highland airframe
       stress accrues at 2.15x, so the same 150 flight hours is 323 stress
       hours, far past B10, and 56.3% fail first.

       Same part, same interval, same policy; 2.6 times the failure rate. Only
       the work is different.

       Note these are BASE ratios, not profile multipliers. The MOUNTAIN profile
       coefficient is 2.400, but no airframe reaches it because every airframe
       flies a mix -- the hardest-worked in the fleet sits at 2.18.
       */
       WeibullScale = CAST(si.IntervalStressHours / POWER(-LOG(0.90), 1.0 / si.WeibullShape) AS DECIMAL(9,3))
INTO #Pos
FROM dbo.Dim_Airframe a
CROSS JOIN dbo.Dim_ComponentType ct
JOIN dbo.Ref_ServiceInterval si ON si.ComponentCode = ct.ComponentCode
-- TOP over a catalog view with no ORDER BY does not guarantee WHICH rows come
-- back, so the position numbers could in principle be {1,4,7,9} rather than
-- {1,2,3,4}. Ordering the numbers before taking the top n removes the doubt --
-- the same construct section 9 refuses to rely on.
CROSS APPLY (SELECT TOP (ct.PositionsPerAirframe) z.PositionNo
             FROM (SELECT PositionNo = ROW_NUMBER() OVER (ORDER BY (SELECT NULL))
                   FROM sys.all_objects) z
             ORDER BY z.PositionNo) p;

IF OBJECT_ID('tempdb..#Gen') IS NOT NULL DROP TABLE #Gen;
CREATE TABLE #Gen (
    GenKey INT IDENTITY(1,1) PRIMARY KEY,
    AirframeKey INT, TailNumber VARCHAR(12), ComponentTypeKey INT,
    ComponentCode VARCHAR(12), PositionNo TINYINT,
    GenNo INT,
    StartStress DECIMAL(12,4), StartFlight DECIMAL(12,4),
    LifeStress  DECIMAL(9,3),                  -- the physics clock's finish line
    EndStress   DECIMAL(12,4),                 -- StartStress + LifeStress
    SchedFlight DECIMAL(12,4),                 -- the maintenance clock's finish line
    EndStressActual DECIMAL(12,4), EndFlightActual DECIMAL(12,4),
    InstalledDateKey INT, RemovedDateKey INT, RemovalReason VARCHAR(20),
    InstallKey INT NULL);

-- generation 1: everything fitted at entry into service
INSERT INTO #Gen (AirframeKey, TailNumber, ComponentTypeKey, ComponentCode, PositionNo, GenNo,
                  StartStress, StartFlight, LifeStress, EndStress, SchedFlight, InstalledDateKey)
SELECT p.AirframeKey, p.TailNumber, p.ComponentTypeKey, p.ComponentCode, p.PositionNo, 1,
       0, 0,
       -- B10-derived scale; see the comment on #Pos above.
       dbo.fn_WeibullLife(p.WeibullScale, p.WeibullShape,
            dbo.fn_Rand(CONCAT('life|', p.TailNumber, '|', p.ComponentCode, '|', p.PositionNo, '|1'))),
       0, p.IntervalFlightHours, a.InServiceDateKey
FROM #Pos p JOIN dbo.Dim_Airframe a ON a.AirframeKey = p.AirframeKey;

UPDATE #Gen SET EndStress = StartStress + LifeStress;

/*
The loop runs until a pass closes nothing, NOT for a fixed number of passes.

This was originally capped at six, which looked generous -- a 150-hour interval
against roughly 560 flight hours per airframe is under four replacements. It is
wrong for exactly the airframes this project is about: a mountain airframe
accrues stress at 2.15x, burns a rotor set every ~69 flight hours, and needs
seven generations -- one more than the cap allowed. Capping at six left generation seven created but never
closed, so those components sat permanently fitted with their stress life long
exhausted and never failed.

The consequence was not a missing row here and there. It suppressed failures
PRECISELY ON THE HIGH-STRESS AIRFRAMES, which is the effect the dataset exists
to demonstrate -- the bug would have quietly argued against the finding.

The cap below is a runaway guard, not a model parameter. The gate after the
loop is what actually proves the loop ran long enough.
*/
DECLARE @gen INT = 1, @closed INT = 1;
WHILE @gen <= 40 AND @closed > 0
BEGIN
    -- Run the race for this generation. The two OUTER APPLYs find where each
    -- clock finishes; the third picks the winner by SORTIE rather than by date,
    -- because both clocks can finish on the same day.
    UPDATE g
    SET RemovedDateKey  = w.SortieDateKey,
        RemovalReason   = w.Reason,
        EndStressActual = w.CumStress,
        EndFlightActual = w.CumFlight
    FROM #Gen g
    OUTER APPLY (SELECT TOP 1 s.SortieSeq, s.SortieDateKey, s.CumFlight, s.CumStress
                 FROM #Stress s
                 WHERE s.AirframeKey = g.AirframeKey AND s.CumStress >= g.EndStress
                 ORDER BY s.SortieSeq) fl
    OUTER APPLY (SELECT TOP 1 s.SortieSeq, s.SortieDateKey, s.CumFlight, s.CumStress
                 FROM #Stress s
                 WHERE s.AirframeKey = g.AirframeKey AND s.CumFlight >= g.SchedFlight
                 ORDER BY s.SortieSeq) sc
    CROSS APPLY (SELECT TOP 1 u.Reason, u.SortieSeq, u.SortieDateKey, u.CumFlight, u.CumStress
                 FROM (SELECT Reason = 'Failure',   fl.SortieSeq, fl.SortieDateKey, fl.CumFlight, fl.CumStress
                       UNION ALL
                       SELECT Reason = 'Scheduled', sc.SortieSeq, sc.SortieDateKey, sc.CumFlight, sc.CumStress) u
                 WHERE u.SortieSeq IS NOT NULL
                 -- Ties go to Failure. If the part was already past its life on
                 -- the sortie the inspection came due, it failed first.
                 ORDER BY u.SortieSeq, CASE u.Reason WHEN 'Failure' THEN 0 ELSE 1 END) w
    WHERE g.GenNo = @gen AND g.RemovedDateKey IS NULL;

    SET @closed = @@ROWCOUNT;   -- must be the very next statement

    -- fit a replacement on every position that just came off
    INSERT INTO #Gen (AirframeKey, TailNumber, ComponentTypeKey, ComponentCode, PositionNo, GenNo,
                      StartStress, StartFlight, LifeStress, EndStress, SchedFlight, InstalledDateKey)
    SELECT g.AirframeKey, g.TailNumber, g.ComponentTypeKey, g.ComponentCode, g.PositionNo, @gen + 1,
           g.EndStressActual, g.EndFlightActual,
           dbo.fn_WeibullLife(p.WeibullScale, p.WeibullShape,
                dbo.fn_Rand(CONCAT('life|', g.TailNumber, '|', g.ComponentCode, '|', g.PositionNo, '|', @gen + 1))),
           0, g.EndFlightActual + p.IntervalFlightHours, g.RemovedDateKey
    FROM #Gen g
    JOIN #Pos p ON p.AirframeKey = g.AirframeKey AND p.ComponentCode = g.ComponentCode AND p.PositionNo = g.PositionNo
    WHERE g.GenNo = @gen AND g.RemovedDateKey IS NOT NULL;

    UPDATE #Gen SET EndStress = StartStress + LifeStress WHERE GenNo = @gen + 1;
    SET @gen += 1;
END

-- an airframe retired while a part was fitted takes the part off with it
UPDATE g SET RemovedDateKey = a.RetiredDateKey, RemovalReason = 'AirframeRetired'
FROM #Gen g JOIN dbo.Dim_Airframe a ON a.AirframeKey = g.AirframeKey
WHERE g.RemovedDateKey IS NULL AND a.AirframeStatus = 'Retired';

/*
The gate that proves the loop ran long enough.

Every component still shown as fitted must have stress life REMAINING. One that
is still fitted with its life already exhausted did not survive -- it is a
component the loop stopped simulating. Asserting the property directly is what
makes the earlier cap of six impossible to reintroduce unnoticed; raising the
cap alone would have fixed today's data and left the next person to rediscover
it when the fleet grows or the window lengthens.
*/
IF EXISTS (
    SELECT 1 FROM #Gen g
    WHERE g.RemovedDateKey IS NULL
      AND EXISTS (SELECT 1 FROM #Stress s
                  WHERE s.AirframeKey = g.AirframeKey AND s.CumStress >= g.EndStress)
)
    THROW 53012, 'Generation loop stopped while components still had exhausted stress life fitted. Raise the runaway cap in section 6 -- the fleet now cycles components more often than the loop allows.', 1;

/*
BOTH clocks, not one.

The gate above checks only the physics clock. A truncated loop that stranded
components past their SCHEDULED replacement point would sail through it, and the
symptom would be an unscheduled rate that looked better than the fleet deserved
-- the direction that flatters the finding.
*/
IF EXISTS (
    SELECT 1 FROM #Gen g
    WHERE g.RemovedDateKey IS NULL
      AND EXISTS (SELECT 1 FROM #Stress s
                  WHERE s.AirframeKey = g.AirframeKey AND s.CumFlight >= g.SchedFlight)
)
    THROW 53013, 'Generation loop stopped while components were still fitted past their scheduled replacement point. The scheduled-removal count is understated and the unscheduled rate is flattered.', 1;

INSERT INTO dbo.Fact_ComponentInstall (ComponentSerial, ComponentTypeKey, AirframeKey, PositionNo,
                                       InstalledDateKey, RemovedDateKey, RemovalReason)
SELECT
    -- ordered by TAIL NUMBER so the serial a component carries does not depend
    -- on what surrogate key its airframe happened to be handed
    ComponentSerial = g.ComponentCode + '-' + RIGHT('00000' + CAST(
        ROW_NUMBER() OVER (PARTITION BY g.ComponentCode ORDER BY g.TailNumber, g.PositionNo, g.GenNo) AS VARCHAR(5)), 5),
    g.ComponentTypeKey, g.AirframeKey, g.PositionNo,
    g.InstalledDateKey, g.RemovedDateKey, g.RemovalReason
FROM #Gen g
-- a generation fitted after the reporting window never existed
JOIN dbo.Dim_Date d ON d.DateKey = g.InstalledDateKey
WHERE d.[Date] <= '2026-09-30'
ORDER BY g.TailNumber, g.ComponentCode, g.PositionNo, g.GenNo;

-- Carry the install keys back onto #Gen so the sensor and event steps can find
-- the life fraction without re-deriving it.
UPDATE g SET InstallKey = ci.InstallKey
FROM #Gen g
JOIN dbo.Fact_ComponentInstall ci
  ON ci.AirframeKey = g.AirframeKey AND ci.ComponentTypeKey = g.ComponentTypeKey
 AND ci.PositionNo = g.PositionNo   AND ci.InstalledDateKey = g.InstalledDateKey;
GO

-- =============================================================================
-- 7. Fact_SensorReading
--
-- One row per sortie per MONITORED component instance. Rotor assemblies and
-- motors carry condition monitoring; the rest do not -- which becomes a finding
-- in its own right when an unmonitored component turns out to drive the queue.
--
-- PLANTED P2: the degradation signal is real, but its SHAPE differs by
-- component. The exponent on the life fraction sets how early it is visible:
--
--   ROTOR-ASSY  exponent 2.3 -- rises gradually, detectable from roughly 55% of
--                               life, which is many flying hours of warning.
--   MOTOR-ESC   exponent 3.6 -- stays flat, then turns up sharply past ~85% of
--                               life, by which point the remaining life can be
--                               shorter than the lead time to get a part.
--
-- A model trained on both will score well on both. Only one of them gives the
-- planner enough notice to act, and that gap is the second finding.
-- =============================================================================
/*
This runs in TWO statements, and the reason is worth stating because the
one-statement version does not work.

#Gen joins #Stress on AirframeKey alone, so before filtering it produces a row
for every sortie the airframe ever flew -- including sorties flown two years
after the component came off. For those rows the life fraction is not 0.8, it
is 180, and TempRiseC is DECIMAL(5,2). The WHERE clause removes them, but
SQL Server does not promise to apply a filter before it evaluates the SELECT
list, and here it did not: the insert failed with an arithmetic overflow.

The fix is NOT to clamp the life fraction. A clamp would make the expression
incapable of overflowing and would therefore also make it incapable of telling
anyone the join was wrong. Materialising the filtered rows first is what
actually removes the hazard: after this first statement, every remaining row is
one where the component was genuinely fitted, so the narrow casts below cannot
see a value they were not designed for.
*/
IF OBJECT_ID('tempdb..#Fitted') IS NOT NULL DROP TABLE #Fitted;

/*
THE HANDOVER SORTIE BELONGS TO THE PART THAT WAS ON THE AIRCRAFT FOR IT.

A component is removed ON a sortie and its replacement is fitted after that
sortie, so both installs carry the same date. A window that is inclusive at both
ends therefore charges that day's flying to BOTH parts -- 1,144 sorties in this
dataset -- and generates a sensor reading for a component that was still in its
crate.

The outgoing part keeps the sortie (it is what finished it off). The incoming
part starts from the next one, which is what the wear simulation in section 6
already assumes: a replacement's StartStress is the cumulative stress AT the
handover sortie, so its own accrued stress is zero until the following flight.
*/
SELECT
    s.SortieKey,
    g.InstallKey,
    g.ComponentCode,
    ci.ComponentSerial,
    fs.SortieID,
    -- Wide on purpose. This column is allowed to hold a nonsense value; the
    -- narrow ones downstream are not.
    LifeFraction = CAST((s.CumStress - g.StartStress) / g.LifeStress AS DECIMAL(12,6))
INTO #Fitted
FROM #Gen g
JOIN dbo.Fact_ComponentInstall ci ON ci.InstallKey = g.InstallKey
JOIN #Stress s       ON s.AirframeKey = g.AirframeKey
JOIN dbo.Fact_Sortie fs ON fs.SortieKey = s.SortieKey
JOIN dbo.Dim_Date ds ON ds.DateKey = s.SortieDateKey
JOIN dbo.Dim_Date di ON di.DateKey = g.InstalledDateKey
LEFT JOIN dbo.Dim_Date dr ON dr.DateKey = g.RemovedDateKey
WHERE g.InstallKey IS NOT NULL
  AND g.ComponentCode IN ('ROTOR-ASSY','MOTOR-ESC')
  AND g.LifeStress > 0
  -- generation 1 is fitted before the aircraft ever flies, so it keeps its
  -- first day; every later generation starts after the handover sortie
  AND (CASE WHEN g.GenNo = 1 THEN ds.[Date] ELSE DATEADD(DAY, -1, ds.[Date]) END) >= di.[Date]
  AND (dr.[Date] IS NULL OR ds.[Date] <= dr.[Date])
  AND s.CumStress >= g.StartStress;

INSERT INTO dbo.Fact_SensorReading (SortieKey, InstallKey, VibrationRms, TempRiseC, CurrentDrawA)
SELECT
    f.SortieKey,
    f.InstallKey,
    VibrationRms = CAST(
        CASE f.ComponentCode
            WHEN 'ROTOR-ASSY' THEN 0.42 + 1.55 * POWER(f.LifeFraction, 2.3)
            ELSE                   0.30 + 0.85 * POWER(f.LifeFraction, 3.6)
        END
        * (1.0 + (dbo.fn_Rand(CONCAT('sn|v|', f.SortieID, '|', f.ComponentSerial)) - 0.5) * 0.16)
        AS DECIMAL(7,4)),
    TempRiseC = CAST(
        (14.0 + 17.0 * POWER(f.LifeFraction, 2.0))
        * (1.0 + (dbo.fn_Rand(CONCAT('sn|t|', f.SortieID, '|', f.ComponentSerial)) - 0.5) * 0.20)
        AS DECIMAL(5,2)),
    CurrentDrawA = CAST(
        (9.0 + 4.6 * POWER(f.LifeFraction, 2.8))
        * (1.0 + (dbo.fn_Rand(CONCAT('sn|c|', f.SortieID, '|', f.ComponentSerial)) - 0.5) * 0.12)
        AS DECIMAL(6,3))
FROM #Fitted f;
GO

-- =============================================================================
-- 8. Fact_MaintenanceEvent
--
-- Downtime separates the two removal reasons, because that is where the money
-- is. A scheduled replacement is booked into a hangar bay with the part already
-- on the shelf. An unscheduled one grounds an aircraft until a part arrives.
-- =============================================================================
INSERT INTO dbo.Fact_MaintenanceEvent (EventID, AirframeKey, InstallKey, EventDateKey, EventType,
                                       ComponentHoursAtEvent, DowntimeHours, Finding)
SELECT
    EventID = 'ME-' + RIGHT('000000' + CAST(ROW_NUMBER() OVER (ORDER BY g.RemovedDateKey, ci.ComponentSerial) AS VARCHAR(6)), 6),
    g.AirframeKey, g.InstallKey, g.RemovedDateKey,
    EventType = CASE g.RemovalReason
        WHEN 'Failure'   THEN 'Unscheduled'
        WHEN 'Scheduled' THEN 'Scheduled'
        ELSE                  'Inspection' END,
    /*
    EndFlightActual is only populated by the race in section 6; a component
    taken off with a retired airframe never raced, so it is NULL and the
    original ISNULL(..., 0) quietly reported 0.00 hours on all 39 retirement
    events. Zero is a measurement, not a missing value, and it dragged the
    average hours-at-removal down for anyone who did not split by event type.
    The fallback now computes the hours actually flown.
    */
    ComponentHoursAtEvent = CAST(
        COALESCE(g.EndFlightActual - g.StartFlight, ret.FlightHours, 0) AS DECIMAL(9,2)),
    DowntimeHours = CAST(
        CASE WHEN g.RemovalReason = 'Failure'
             -- AOG: the aircraft is grounded until a part arrives, and the
             -- flight-critical parts carry the longest lead times
             THEN CASE ct.Criticality WHEN 'FlightCritical'  THEN 14.0
                                      WHEN 'MissionCritical' THEN  7.0 ELSE 4.0 END
             ELSE CASE ct.Criticality WHEN 'FlightCritical'  THEN  4.5
                                      WHEN 'MissionCritical' THEN  2.5 ELSE 1.5 END
        END * (0.6 + dbo.fn_Rand(CONCAT('me|dt|', ci.ComponentSerial)) * 1.3) AS DECIMAL(6,1)),
    Finding = CASE g.RemovalReason
        WHEN 'Failure' THEN CONCAT(ct.ComponentName, ' removed on failure at ',
             CAST(CAST(ISNULL(g.EndFlightActual - g.StartFlight, 0) AS DECIMAL(9,1)) AS VARCHAR(12)),
             ' flight hours against a ', CAST(CAST(si.IntervalFlightHours AS DECIMAL(9,1)) AS VARCHAR(12)),
             ' hour interval. ',
             CASE g.ComponentCode
                WHEN 'ROTOR-ASSY' THEN 'Bearing roughness beyond limits; spectral peak at outer-race frequency.'
                WHEN 'MOTOR-ESC'  THEN 'Winding temperature excursion under load; controller derating observed.'
                WHEN 'BATT-PACK'  THEN 'Capacity below dispatch minimum; cell imbalance on charge.'
                WHEN 'GIMBAL'     THEN 'Servo backlash beyond limits; slip-ring intermittent.'
                WHEN 'AIRDATA'    THEN 'Static port disagreement between channels.'
                ELSE                   'Fatigue cracking at the leg attachment.' END)
        WHEN 'Scheduled' THEN CONCAT(ct.ComponentName, ' replaced at the scheduled ',
             CAST(CAST(si.IntervalFlightHours AS DECIMAL(9,1)) AS VARCHAR(12)),
             ' flight hour interval. Part serviceable on removal.')
        ELSE CONCAT(ct.ComponentName, ' removed with airframe retirement; serviceable, returned to stores.') END
FROM #Gen g
JOIN dbo.Fact_ComponentInstall ci ON ci.InstallKey = g.InstallKey
JOIN dbo.Dim_ComponentType ct   ON ct.ComponentTypeKey = g.ComponentTypeKey
JOIN dbo.Ref_ServiceInterval si ON si.ComponentCode = g.ComponentCode
OUTER APPLY (SELECT FlightHours = MAX(st.CumFlight) - g.StartFlight
             FROM #Stress st
             WHERE st.AirframeKey = g.AirframeKey
               AND st.SortieDateKey <= g.RemovedDateKey) ret
WHERE g.InstallKey IS NOT NULL AND g.RemovedDateKey IS NOT NULL
ORDER BY g.RemovedDateKey, ci.ComponentSerial;
GO


-- =============================================================================
-- 9. PLANTED DATA DEFECTS -- created by behaviour, never by a marker
--
-- EVERY SELECTION HERE IS A DETERMINISTIC TOP-N OVER A TOTAL ORDER.
--
-- The first version used `UPDATE TOP (11) ... WHERE fn_Rand(...) < 0.004`,
-- which reads like a 0.4% sample and is not one. About forty rows satisfy that
-- predicate, TOP takes eleven of them, and WHICH eleven is whatever the plan
-- produces -- no ORDER BY, no guarantee, and a different answer whenever the
-- plan changes. The defects moved between runs while everything else held.
--
-- Each block below instead orders the whole candidate set by a hash of its
-- BUSINESS key and takes the first n. That is a total order over distinct
-- values, so the same n rows are chosen every time, on any server.
-- =============================================================================

-- D1: duplicate sorties
INSERT INTO dbo.Fact_Sortie (SortieID, AirframeKey, MissionProfileKey, SortieDateKey,
                             FlightMinutes, PayloadKg, Landings, AmbientTempC, WindGustKts)
SELECT TOP (18)
       -- derived from the ORIGINAL SortieID, not from its surrogate key
       'D' + SUBSTRING(s.SortieID, 2, 20),
       s.AirframeKey, s.MissionProfileKey, s.SortieDateKey,
       s.FlightMinutes, s.PayloadKg, s.Landings, s.AmbientTempC, s.WindGustKts
FROM dbo.Fact_Sortie s
WHERE s.SortieID LIKE 'S-%'
ORDER BY dbo.fn_Rand(CONCAT('dfct|dup|', s.SortieID)), s.SortieID;
GO

-- D4: flight time recorded with no landings -- physically impossible
;WITH Pick AS (
    SELECT TOP (11) s.SortieKey, s.Landings
    FROM dbo.Fact_Sortie s
    WHERE s.FlightMinutes > 20
    ORDER BY dbo.fn_Rand(CONCAT('dfct|ldg|', s.SortieID)), s.SortieID
)
UPDATE Pick SET Landings = 0;
GO

-- D3: maintenance events dated before the component was installed
;WITH Pick AS (
    SELECT TOP (7) e.EventDateKey, NewDateKey = d2.DateKey
    FROM dbo.Fact_MaintenanceEvent e
    JOIN dbo.Fact_ComponentInstall ci ON ci.InstallKey = e.InstallKey
    JOIN dbo.Dim_Date di ON di.DateKey = ci.InstalledDateKey
    JOIN dbo.Dim_Date d2 ON d2.[Date] = DATEADD(DAY, -14, di.[Date])
    WHERE di.[Date] > '2024-06-01'
    ORDER BY dbo.fn_Rand(CONCAT('dfct|me|', e.EventID)), e.EventID
)
UPDATE Pick SET EventDateKey = NewDateKey;
GO

-- D5: a handful of installs removed before they were installed
;WITH Pick AS (
    SELECT TOP (6) ci.RemovedDateKey, NewDateKey = d2.DateKey
    FROM dbo.Fact_ComponentInstall ci
    JOIN dbo.Dim_Date di ON di.DateKey = ci.InstalledDateKey
    JOIN dbo.Dim_Date d2 ON d2.[Date] = DATEADD(DAY, -9, di.[Date])
    WHERE ci.RemovedDateKey IS NOT NULL AND di.[Date] > '2024-06-01'
    ORDER BY dbo.fn_Rand(CONCAT('dfct|ci|', ci.ComponentSerial)), ci.ComponentSerial
)
UPDATE Pick SET RemovedDateKey = NewDateKey;
GO

-- D2: sensor readings against a component that was not fitted that day
INSERT INTO dbo.Fact_SensorReading (SortieKey, InstallKey, VibrationRms, TempRiseC, CurrentDrawA)
SELECT TOP (13) s.SortieKey, ci.InstallKey, 0.55, 18.0, 9.4
FROM dbo.Fact_ComponentInstall ci
JOIN dbo.Dim_Date dr ON dr.DateKey = ci.RemovedDateKey
JOIN dbo.Fact_Sortie s ON s.AirframeKey = ci.AirframeKey
JOIN dbo.Dim_Date ds ON ds.DateKey = s.SortieDateKey AND ds.[Date] > DATEADD(DAY, 30, dr.[Date])
ORDER BY dbo.fn_Rand(CONCAT('dfct|sn|', ci.ComponentSerial, '|', s.SortieID)),
         ci.ComponentSerial, s.SortieID;
GO

-- =============================================================================
-- 10. Summary and fingerprints
-- =============================================================================
PRINT '';
PRINT '=== Meridian UAV: generation summary ===';
SELECT Bases=(SELECT COUNT(*) FROM dbo.Dim_Base),
       Airframes=(SELECT COUNT(*) FROM dbo.Dim_Airframe),
       Active=(SELECT COUNT(*) FROM dbo.Dim_Airframe WHERE AirframeStatus='Active'),
       Sorties=(SELECT COUNT(*) FROM dbo.Fact_Sortie),
       Installs=(SELECT COUNT(*) FROM dbo.Fact_ComponentInstall),
       Failures=(SELECT COUNT(*) FROM dbo.Fact_ComponentInstall WHERE RemovalReason='Failure'),
       StillFitted=(SELECT COUNT(*) FROM dbo.Fact_ComponentInstall WHERE RemovedDateKey IS NULL),
       Readings=(SELECT COUNT(*) FROM dbo.Fact_SensorReading),
       Events=(SELECT COUNT(*) FROM dbo.Fact_MaintenanceEvent);

PRINT '';
/*
THE FINGERPRINTS, AND TWO THINGS THE FIRST VERSION GOT WRONG.

They are computed over BUSINESS keys and measured values, never over surrogate
keys: a fingerprint built from IDENTITY values tests that the database handed out
the same numbers, which is the one thing that is not guaranteed and not what
determinism means here.

1. NOT CHECKSUM_AGG. It aggregates by XOR, so any two rows with an equal
   CHECKSUM cancel each other out completely and become invisible to the
   fingerprint. 142 sensor readings were in exactly that position. SUM over
   BIGINT has no such cancellation.

2. THE COLUMNS THAT DRIVE THE FINDINGS ARE INCLUDED. The first version hashed
   SortieID, FlightMinutes, Landings and PayloadKg -- and omitted the airframe,
   the mission profile and the date. Those three decide which stress multiplier
   applies, which base it lands on, and which component instance accrues it. A
   rebuild could have reassigned the entire fleet's workload to different
   aircraft on different profiles and reproduced the fingerprint exactly.
*/
PRINT 'Fingerprints (asserted by the UAT suite -- a rebuild must reproduce these):';
SELECT SortieFingerprint = SUM(CAST(CHECKSUM(
           s.SortieID, a.TailNumber, p.ProfileCode, s.SortieDateKey,
           s.FlightMinutes, s.Landings, s.PayloadKg) AS BIGINT))
FROM dbo.Fact_Sortie s
JOIN dbo.Dim_Airframe a       ON a.AirframeKey = s.AirframeKey
JOIN dbo.Dim_MissionProfile p ON p.MissionProfileKey = s.MissionProfileKey;

SELECT InstallFingerprint = SUM(CAST(CHECKSUM(
           ci.ComponentSerial, a.TailNumber, ci.PositionNo,
           ci.InstalledDateKey, ci.RemovedDateKey, ci.RemovalReason) AS BIGINT))
FROM dbo.Fact_ComponentInstall ci
JOIN dbo.Dim_Airframe a ON a.AirframeKey = ci.AirframeKey;

SELECT ReadingFingerprint = SUM(CAST(CHECKSUM(
           ci.ComponentSerial, s.SortieID,
           sr.VibrationRms, sr.TempRiseC, sr.CurrentDrawA) AS BIGINT))
FROM dbo.Fact_SensorReading sr
JOIN dbo.Fact_ComponentInstall ci ON ci.InstallKey = sr.InstallKey
JOIN dbo.Fact_Sortie s ON s.SortieKey = sr.SortieKey;

SELECT EventFingerprint = SUM(CAST(CHECKSUM(
           e.EventID, ci.ComponentSerial, e.EventDateKey, e.EventType,
           e.ComponentHoursAtEvent, e.DowntimeHours) AS BIGINT))
FROM dbo.Fact_MaintenanceEvent e
JOIN dbo.Fact_ComponentInstall ci ON ci.InstallKey = e.InstallKey;
GO

DECLARE @sorties INT = (SELECT COUNT(*) FROM dbo.Fact_Sortie);
DECLARE @failures INT = (SELECT COUNT(*) FROM dbo.Fact_ComponentInstall WHERE RemovalReason = 'Failure');
IF @sorties < 3000 THROW 53010, 'Generator produced too few sorties to accumulate meaningful wear.', 1;
IF @failures < 60  THROW 53011, 'Generator produced too few component failures to train or evaluate a model.', 1;
PRINT '';
PRINT 'Synthetic fleet history generated.';
GO
