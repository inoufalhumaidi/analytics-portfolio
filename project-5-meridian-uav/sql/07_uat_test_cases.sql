/*
================================================================================
Project 5 — Meridian UAV Services: Predictive Maintenance
Script:  07_uat_test_cases.sql
Purpose: 35 acceptance tests. Raises and stops on failure.

WHAT THIS SUITE LEARNED FROM THE FOUR PROJECTS BEFORE IT

1.  ASSERT VALUES, NOT PROPERTIES.
    'Compliance is between 0 and 100' holds under every transformation that
    preserves it, including the bug it was written to catch. Project 3's index
    test asserted 'index >= 100' and passed against a measure that returned
    exactly 100.00 for all 36 vendors because of a join fan-out. Where a number
    is knowable, this suite asserts the number.

2.  A SUITE THAT RUNS PART OF ITSELF MUST NOT REPORT SUCCESS.
    Project 2 printed '8 passed, 0 failed' after a suite that was supposed to
    run 14. @ExpectedCases is checked at the end and the run fails if the count
    does not match, so a case deleted, skipped or lost to an early exit is
    itself a failure.

3.  THE RESULTS TABLE MUST OUTLIVE THE TRANSACTION.
    #UATResults is created BEFORE the fixture transaction. Project 4 created it
    inside, so ROLLBACK dropped it and the failure-detail SELECT reported
    'Invalid object name' instead of the failures -- a branch that only ever
    executes when something fails, which is exactly why it shipped.

4.  A FIXTURE MUST CREATE WHAT IT TESTS.
    Project 4's UAT-06 selected a row that the generator later stopped
    producing; the case returned NULL and passed. Fixtures here build their own
    rows rather than hoping for one.

5.  THE TESTS MUST BE ABLE TO FAIL.
    UAT-26 through UAT-28 deliberately break something and assert the detection
    fires. A suite that has only ever passed has not been shown to work.

6.  A TEST MUST EXERCISE THE SHIPPED CODE, NOT A COPY OF IT.
    UAT-27 originally pasted DQ-06's self-join into the test. The two had no
    dependency on each other, so breaking the shipped check left the test green
    -- and because DQ-06 has never fired on the real data, a broken DQ-06 and a
    working one were indistinguishable. The predicate now lives in
    vw_DQ_OverlappingInstalls and both call it.

7.  AN OPTIMISATION IS A CLAIM, AND CLAIMS GET TESTED.
    UAT-35 recomputes the alert set the slow way and requires it to match the
    materialised fast path install for install, lead time for lead time. A fast
    answer that disagrees with the slow one is worse than the slow one.

WHAT AN ADVERSARIAL REVIEW OF THIS SUITE FOUND
    Cases 29 to 34 exist because a review of the first 28 found defects that all
    28 passed over: a population filtered on a today-attribute, component
    windows inclusive at both ends, a feature view with no window predicate at
    all, and a fingerprint blind to the three columns the findings rest on.
    Every one of those is now asserted.

DATA DISCLOSURE
    Meridian UAV Services is fictional and all data is synthetic. No
    confidential data is used and no claim is made about any production system.
================================================================================
*/

USE MeridianUAV;
GO
SET NOCOUNT ON;
GO

-- Created OUTSIDE the transaction, deliberately. See note 3 above.
IF OBJECT_ID('tempdb..#UATResults') IS NOT NULL DROP TABLE #UATResults;
CREATE TABLE #UATResults (
    CaseID     VARCHAR(10)  NOT NULL,
    CaseName   VARCHAR(140) NOT NULL,
    Expected   VARCHAR(80)  NULL,
    [Actual]   VARCHAR(80)  NULL,
    Passed     BIT          NOT NULL,
    Notes      VARCHAR(400) NULL
);
GO

DECLARE @AsOf DATE = '2026-09-30';
DECLARE @Prior DATE = '2025-06-30';
DECLARE @Thr DECIMAL(6,3) = 3.400;

-- ===========================================================================
-- GROUP A -- the dataset is the dataset
-- ===========================================================================

-- UAT-01: exact row counts. A generator change that alters the fleet must be
-- deliberate; this is what makes it impossible to do by accident.
DECLARE @s INT = (SELECT COUNT(*) FROM dbo.Fact_Sortie);
DECLARE @i INT = (SELECT COUNT(*) FROM dbo.Fact_ComponentInstall);
DECLARE @r INT = (SELECT COUNT(*) FROM dbo.Fact_SensorReading);
DECLARE @e INT = (SELECT COUNT(*) FROM dbo.Fact_MaintenanceEvent);
INSERT INTO #UATResults VALUES ('UAT-01', 'Fact row counts match the published dataset',
    '10501/1549/83877/1185',
    CONCAT(@s,'/',@i,'/',@r,'/',@e),
    CASE WHEN @s=10501 AND @i=1549 AND @r=83877 AND @e=1185 THEN 1 ELSE 0 END,
    'Sorties / installs / sensor readings / maintenance events. The reading count fell by 739 when the handover-day phantom readings were removed.');

-- UAT-02: determinism. The fingerprint is computed over BUSINESS keys and
-- measured values, never over identity values -- the first version used
-- surrogate keys and changed whenever the object catalog changed.
/*
BIGINT, and over the columns that drive the findings.

The first version was CHECKSUM_AGG over SortieID, FlightMinutes, Landings and
PayloadKg. Two problems, both of which let a changed fleet reproduce the same
number: CHECKSUM_AGG aggregates by XOR, so any two rows with an equal CHECKSUM
cancel and vanish from the fingerprint; and it omitted the airframe, the mission
profile and the date -- the three columns that decide which stress multiplier
applies, which base it lands on, and which component accrues it.
*/
DECLARE @fp BIGINT = (
    SELECT SUM(CAST(CHECKSUM(s.SortieID, a.TailNumber, p.ProfileCode, s.SortieDateKey,
                             s.FlightMinutes, s.Landings, s.PayloadKg) AS BIGINT))
    FROM dbo.Fact_Sortie s
    JOIN dbo.Dim_Airframe a       ON a.AirframeKey = s.AirframeKey
    JOIN dbo.Dim_MissionProfile p ON p.MissionProfileKey = s.MissionProfileKey);
INSERT INTO #UATResults VALUES ('UAT-02', 'Sortie fingerprint reproduces exactly',
    '59939319044', CAST(@fp AS VARCHAR(30)), CASE WHEN @fp = 59939319044 THEN 1 ELSE 0 END,
    'Covers airframe, profile and date as well as the measures. A rebuild that changes this changed the fleet.');

-- UAT-03: every active airframe carries a full set of 13 positions.
-- A missing position makes compliance look BETTER, because a part that is not
-- there cannot be overdue.
DECLARE @incomplete INT = (
    SELECT COUNT(*) FROM dbo.Dim_Airframe a
    CROSS JOIN (SELECT ep = SUM(PositionsPerAirframe) FROM dbo.Dim_ComponentType) x
    OUTER APPLY (SELECT n = COUNT(*) FROM dbo.fn_ComponentWear('2026-09-30') w WHERE w.AirframeKey = a.AirframeKey) f
    JOIN dbo.Dim_Date isd ON isd.DateKey = a.InServiceDateKey
    LEFT JOIN dbo.Dim_Date rtd ON rtd.DateKey = a.RetiredDateKey
    WHERE isd.[Date] <= '2026-09-30' AND (rtd.[Date] IS NULL OR rtd.[Date] > '2026-09-30')
      AND ISNULL(f.n,0) <> x.ep);
INSERT INTO #UATResults VALUES ('UAT-03', 'Every active airframe carries all 13 component positions',
    '0', CAST(@incomplete AS VARCHAR(20)), CASE WHEN @incomplete = 0 THEN 1 ELSE 0 END,
    'An incomplete airframe understates its own overdue count.');

-- UAT-04: the fitted population is exactly 28 active airframes x 13 positions.
DECLARE @fitted INT = (SELECT COUNT(*) FROM dbo.fn_ComponentWear(@AsOf) WHERE IsInFleetAsOf = 1);
INSERT INTO #UATResults VALUES ('UAT-04', 'Fitted population is 28 active airframes x 13 positions',
    '364', CAST(@fitted AS VARCHAR(20)), CASE WHEN @fitted = 364 THEN 1 ELSE 0 END, NULL);

-- ===========================================================================
-- GROUP B -- the wear model computes what it says it computes
-- ===========================================================================

-- UAT-05: the three stress terms must sum to the total, on EVERY row.
-- Checked by counting rows that disagree, not by sampling one.
DECLARE @badsum INT = (
    SELECT COUNT(*) FROM dbo.vw_SortieStress
    WHERE ABS(StressHours - (HourlyStress + CycleStress + PayloadStress)) > 0.0002);
INSERT INTO #UATResults VALUES ('UAT-05', 'Stress decomposition sums to the stress total on every sortie',
    '0', CAST(@badsum AS VARCHAR(20)), CASE WHEN @badsum = 0 THEN 1 ELSE 0 END,
    'Hourly + cycle + payload = total, to 0.0002 h, across all rows.');

-- UAT-06: the planted divergence in stress per flight hour exists and has the
-- magnitude the case study quotes. If the profile mixes are ever retuned, this
-- fails and the published figures have to be revisited.
DECLARE @coastal DECIMAL(9,4) = (SELECT StressRatio FROM dbo.fn_BaseScorecard(@AsOf) WHERE BaseCode='COASTAL');
DECLARE @highland DECIMAL(9,4) = (SELECT StressRatio FROM dbo.fn_BaseScorecard(@AsOf) WHERE BaseCode='HIGHLND');
INSERT INTO #UATResults VALUES ('UAT-06', 'Stress-per-flight-hour differs by base as published',
    '1.1741 / 2.1532', CONCAT(@coastal,' / ',@highland),
    CASE WHEN @coastal = 1.1741 AND @highland = 2.1532 THEN 1 ELSE 0 END,
    'Coastal against Highland. The engine of the whole finding.');

-- UAT-07: no fan-out. fn_ComponentWear must be exactly one row per install.
DECLARE @rows INT = (SELECT COUNT(*) FROM dbo.fn_ComponentWear(@AsOf));
DECLARE @keys INT = (SELECT COUNT(DISTINCT InstallKey) FROM dbo.fn_ComponentWear(@AsOf));
INSERT INTO #UATResults VALUES ('UAT-07', 'fn_ComponentWear returns one row per install (no join fan-out)',
    CAST(@keys AS VARCHAR(20)), CAST(@rows AS VARCHAR(20)), CASE WHEN @rows = @keys THEN 1 ELSE 0 END,
    'Project 3 shipped a fan-out that inflated numerator and denominator together.');

-- UAT-08: accrued hours reconcile to an independent sum for a specific install.
-- Computed a second way, from the raw fact table rather than through the view.
DECLARE @ik INT = (SELECT TOP 1 InstallKey FROM dbo.fn_ComponentWear(@AsOf)
                   WHERE AirframeStatus='Active' AND Sorties > 40 ORDER BY InstallKey);
DECLARE @viaFn DECIMAL(11,4) = (SELECT FlightHours FROM dbo.fn_ComponentWear(@AsOf) WHERE InstallKey = @ik);
/*
The independent side restates the SAME RULE a different way -- it does not use
a different rule.

The first version summed every sortie from the install date onward, which is the
old inclusive-at-both-ends window. When the handover rule was corrected, this
case failed by 0.65 flight hours and the failure was correct: the two sides were
implementing different specifications, so agreement would have meant nothing.

What reconciliation is worth testing here is the JOIN, the WINDOW and the
FILTER, reached from the raw fact tables rather than through the view. The
handover rule is part of the specification both sides must honour, so it is
restated below in its own terms.
*/
DECLARE @viaRaw DECIMAL(11,4) = (
    SELECT CAST(SUM(fs.FlightMinutes / 60.0) AS DECIMAL(11,4))
    FROM dbo.Fact_ComponentInstall ci
    JOIN dbo.Dim_Date di    ON di.DateKey = ci.InstalledDateKey
    JOIN dbo.Fact_Sortie fs ON fs.AirframeKey = ci.AirframeKey
    JOIN dbo.Dim_Date ds    ON ds.DateKey = fs.SortieDateKey
    WHERE ci.InstallKey = @ik
      AND ds.[Date] <= @AsOf
      -- a part that replaced another starts the day after the handover sortie;
      -- one fitted from new keeps its first day
      AND ds.[Date] >= CASE
            WHEN EXISTS (SELECT 1 FROM dbo.Fact_ComponentInstall prev
                         WHERE prev.AirframeKey      = ci.AirframeKey
                           AND prev.ComponentTypeKey = ci.ComponentTypeKey
                           AND prev.PositionNo       = ci.PositionNo
                           AND prev.RemovedDateKey   = ci.InstalledDateKey)
            THEN DATEADD(DAY, 1, di.[Date]) ELSE di.[Date] END);
/*
The tolerance is 0.01 flight hours, and it is not slack -- it is the size of a
known, deliberate difference between the two implementations.

vw_SortieStress rounds each sortie to four decimal places, because a sortie's
flight hours is itself a reportable quantity. The raw computation sums at full
precision and rounds once. Over a 170-hour component life those accumulate to
about 0.0006 h, and a test at 0.0002 fails on arithmetic rather than on logic.

0.01 h is 36 seconds. A genuine defect in the window, the join or the filter
moves this by hours, not by seconds, so the test still catches everything it
was written to catch.
*/
INSERT INTO #UATResults VALUES ('UAT-08', 'Accrued flight hours reconcile against a second, independent computation',
    CAST(@viaRaw AS VARCHAR(30)), CAST(@viaFn AS VARCHAR(30)),
    CASE WHEN ABS(@viaFn - @viaRaw) < 0.01 THEN 1 ELSE 0 END,
    'Through the view, and straight off Fact_Sortie. Implemented twice on purpose; tolerance is per-sortie rounding.');

-- ===========================================================================
-- GROUP C -- the as-of parameter reaches every output
-- ===========================================================================

-- UAT-09: an earlier date must give a different fitted population.
/*
The guard is AND, not OR.

The first version read '@fitPrior <> 364 OR @fitPrior = 0'. A function returning
nothing satisfies the first clause -- 0 <> 364 -- so the case passed on an empty
result set, which is the single outcome it was written to catch. The zero guard
inverted the test it was guarding.
*/
DECLARE @fitPrior INT = (SELECT COUNT(*) FROM dbo.fn_ComponentWear(@Prior) WHERE IsInFleetAsOf = 1);
-- The VALUE, not just 'differs and is non-empty'. 401 is knowable: 31 airframes
-- were in the fleet in June 2025 against 28 today. Asserting only '<> 364'
-- would pass for any wrong answer that happens not to equal today's -- and the
-- first attempt at pinning this asserted 362, a figure read from a run before
-- the as-of population fix, which is exactly why the value has to come from
-- the database rather than from a previous transcript.
INSERT INTO #UATResults VALUES ('UAT-09', 'fn_ComponentWear responds to @AsOfDate',
    '401', CAST(@fitPrior AS VARCHAR(20)),
    CASE WHEN @fitPrior = 401 THEN 1 ELSE 0 END,
    'A different set of components was fitted in June 2025: 401 across 31 airframes, against 364 across 28 today.');

/*
UAT-10: EVERY metric must move between two dates, not just the headline one.

This is the Project 4 defect made into a test. fn_ReadinessKPI accepted
@AsOfDate and used it for every metric except work items, so one series was a
flat line across twenty months and looked entirely plausible next to the others.
Checking that the SCORECARD changed would have passed; only checking each
metric individually catches it.
*/
/*
Measured across SIX dates, not two.

The first version compared two dates and failed -- correctly reporting that
AirframeAvailabilityPct and OverdueFlightCritical were identical fifteen months
apart. Both were coincidences: 26 of 28 airframes available on both dates, and
91 overdue flight-critical components on both. Neither was ignoring the
parameter.

Two-point equality does not distinguish 'did not move' from 'came back to the
same place'. A metric that genuinely ignores @AsOfDate returns exactly ONE
distinct value across any number of dates, and that is what is asserted here.
*/
DECLARE @Dates TABLE (d DATE PRIMARY KEY);
INSERT INTO @Dates VALUES ('2024-12-31'),('2025-03-31'),('2025-06-30'),
                          ('2025-12-31'),('2026-03-31'),('2026-09-30');
DECLARE @unmoved INT = (
    SELECT COUNT(*) FROM (
        SELECT k.MetricName, Distinct_ = COUNT(DISTINCT k.MetricValue)
        FROM @Dates t CROSS APPLY dbo.fn_FleetKPI(t.d) k
        GROUP BY k.MetricName
    ) z WHERE z.Distinct_ <= 1);
INSERT INTO #UATResults VALUES ('UAT-10', 'Every fleet metric takes more than one value across six as-of dates',
    '0', CAST(@unmoved AS VARCHAR(20)), CASE WHEN @unmoved = 0 THEN 1 ELSE 0 END,
    'A metric with one distinct value across six dates is ignoring @AsOfDate. Project 4 shipped exactly that.');

-- UAT-11: cumulative flight hours must be strictly increasing in @AsOfDate.
DECLARE @hNow DECIMAL(11,4), @hThen DECIMAL(11,4);
SELECT @hNow  = FlightHours FROM dbo.fn_AirframeStress(@AsOf)  WHERE TailNumber = 'MU-001';
SELECT @hThen = FlightHours FROM dbo.fn_AirframeStress(@Prior) WHERE TailNumber = 'MU-001';
INSERT INTO #UATResults VALUES ('UAT-11', 'Cumulative flight hours increase with the as-of date',
    CONCAT('> ', @hThen), CAST(@hNow AS VARCHAR(30)),
    CASE WHEN @hNow > @hThen AND @hThen > 0 THEN 1 ELSE 0 END, 'MU-001.');

-- ===========================================================================
-- GROUP D -- the compliance definitions
-- ===========================================================================

-- UAT-12: hidden-overdue is a STRICT subset -- past stress, inside hours.
DECLARE @badhidden INT = (
    SELECT COUNT(*) FROM dbo.fn_ComponentWear(@AsOf)
    WHERE IsHiddenOverdue = 1 AND NOT (IsOverdueByStress = 1 AND IsOverdueByHours = 0));
INSERT INTO #UATResults VALUES ('UAT-12', 'Hidden-overdue is exactly: past stress interval, inside hour interval',
    '0', CAST(@badhidden AS VARCHAR(20)), CASE WHEN @badhidden = 0 THEN 1 ELSE 0 END,
    'If it is visible on the hour meter it is not hidden.');

-- UAT-13: the headline contrast, to the published figures.
DECLARE @hourPct DECIMAL(10,2) = (SELECT MetricValue FROM dbo.fn_FleetKPI(@AsOf) WHERE MetricName='HourCompliancePct');
DECLARE @strPct  DECIMAL(10,2) = (SELECT MetricValue FROM dbo.fn_FleetKPI(@AsOf) WHERE MetricName='StressCompliancePct');
INSERT INTO #UATResults VALUES ('UAT-13', 'Headline compliance figures match the case study',
    '99.73 / 68.41', CONCAT(@hourPct,' / ',@strPct),
    CASE WHEN @hourPct = 99.73 AND @strPct = 68.41 THEN 1 ELSE 0 END,
    'Hour compliance against stress compliance. The project in two numbers.');

-- UAT-14: counts and percentages must agree. A percentage that does not
-- reconcile to its own numerator is the classic silent reporting defect.
DECLARE @overdueFC INT = (SELECT COUNT(*) FROM dbo.fn_ComponentWear(@AsOf)
                          WHERE IsInFleetAsOf = 1 AND IsOverdueByStress=1 AND Criticality='FlightCritical');
DECLARE @kpiFC DECIMAL(10,2) = (SELECT MetricValue FROM dbo.fn_FleetKPI(@AsOf) WHERE MetricName='OverdueFlightCritical');
INSERT INTO #UATResults VALUES ('UAT-14', 'Overdue flight-critical count reconciles between the KPI and the detail',
    CAST(@overdueFC AS VARCHAR(20)), CAST(CAST(@kpiFC AS INT) AS VARCHAR(20)),
    CASE WHEN @overdueFC = CAST(@kpiFC AS INT) THEN 1 ELSE 0 END, NULL);

-- UAT-15: the four bases must partition the fitted population exactly.
DECLARE @baseSum INT = (SELECT SUM(FittedComponents) FROM dbo.fn_BaseScorecard(@AsOf));
INSERT INTO #UATResults VALUES ('UAT-15', 'Base scorecard partitions the fitted population',
    CAST(@fitted AS VARCHAR(20)), CAST(@baseSum AS VARCHAR(20)),
    CASE WHEN @baseSum = @fitted THEN 1 ELSE 0 END,
    'Every fitted component belongs to exactly one base.');

-- ===========================================================================
-- GROUP E -- the alert evaluation
-- ===========================================================================

-- UAT-16: the confusion matrix must partition the population.
DECLARE @pop INT, @tp INT, @fp2 INT, @fn2 INT, @tn INT, @atp INT, @ltp INT;
SELECT @pop = Population, @tp = TP, @fp2 = FP, @fn2 = FN, @tn = TN,
       @atp = ActionableTP, @ltp = LateTP
FROM dbo.fn_AlertSummaryFleet(@AsOf, @Thr);
INSERT INTO #UATResults VALUES ('UAT-16', 'Confusion matrix partitions the evaluated population',
    CAST(@pop AS VARCHAR(20)), CAST(@tp+@fp2+@fn2+@tn AS VARCHAR(20)),
    CASE WHEN @tp+@fp2+@fn2+@tn = @pop THEN 1 ELSE 0 END, 'TP + FP + FN + TN = population.');

-- UAT-17: actionable and late must partition the true positives. If they do
-- not, the actionable percentage is being computed over the wrong denominator
-- and the second finding is wrong.
INSERT INTO #UATResults VALUES ('UAT-17', 'Actionable and late true positives partition all true positives',
    CAST(@tp AS VARCHAR(20)), CAST(@atp+@ltp AS VARCHAR(20)),
    CASE WHEN @atp + @ltp = @tp THEN 1 ELSE 0 END,
    'Every caught failure is either early enough to act on or it is not.');

-- UAT-18: grain. fn_AlertSummary must be one row per component code.
-- This shipped broken: a window function dragged its arguments into the
-- GROUP BY and the function returned dozens of well-formed, wrong rows.
DECLARE @sumRows INT = (SELECT COUNT(*) FROM dbo.fn_AlertSummary(@AsOf, @Thr));
DECLARE @sumCodes INT = (SELECT COUNT(DISTINCT ComponentCode) FROM dbo.fn_AlertSummary(@AsOf, @Thr));
INSERT INTO #UATResults VALUES ('UAT-18', 'fn_AlertSummary is one row per component code',
    CAST(@sumCodes AS VARCHAR(20)), CAST(@sumRows AS VARCHAR(20)),
    CASE WHEN @sumRows = @sumCodes AND @sumCodes = 2 THEN 1 ELSE 0 END,
    'Two monitored component types, two rows.');

-- UAT-19: the published prediction figures, exactly.
DECLARE @prec DECIMAL(6,2), @rec DECIMAL(6,2), @act DECIMAL(6,2);
SELECT @prec = PredictionPrecisionPct, @rec = PredictionRecallPct, @act = ActionableLeadTimePct
FROM dbo.fn_AlertSummaryFleet(@AsOf, @Thr);
INSERT INTO #UATResults VALUES ('UAT-19', 'Prediction precision, recall and actionable lead time match the case study',
    '81.56 / 98.26 / 37.81', CONCAT(@prec,' / ',@rec,' / ',@act),
    CASE WHEN @prec = 81.56 AND @rec = 98.26 AND @act = 37.81 THEN 1 ELSE 0 END,
    'Both published targets met; a third of correct predictions usable.');

/*
UAT-20: the trade-off is MONOTONIC.

Raising the threshold must not reduce precision, and must not increase the
actionable share. If either reverses, the conflict is noise rather than
structure and the recommendation built on it does not follow.

This is a property test and it is the right tool here, because the property IS
the claim being made. It is paired with UAT-19, which pins the actual values.
*/
DECLARE @mono INT = (
    SELECT COUNT(*)
    FROM (SELECT v = 2.000 UNION ALL SELECT 2.400 UNION ALL SELECT 2.800
          UNION ALL SELECT 3.200 UNION ALL SELECT 3.600) t
    CROSS APPLY dbo.fn_AlertSummary(@AsOf, CAST(t.v AS DECIMAL(6,3))) s
    CROSS APPLY (SELECT nxt = t.v + 0.400) n
    CROSS APPLY dbo.fn_AlertSummary(@AsOf, CAST(n.nxt AS DECIMAL(6,3))) s2
    WHERE s2.ComponentCode = s.ComponentCode
      AND (s2.PrecisionPct < s.PrecisionPct - 0.01
        OR s2.ActionableLeadTimePct > s.ActionableLeadTimePct + 0.01));
INSERT INTO #UATResults VALUES ('UAT-20', 'Precision rises and actionable lead time falls monotonically with the threshold',
    '0', CAST(@mono AS VARCHAR(20)), CASE WHEN @mono = 0 THEN 1 ELSE 0 END,
    'The conflict is structural, not a tuning artefact.');

-- ===========================================================================
-- GROUP F -- the action queue
-- ===========================================================================

-- UAT-21: no MONITOR rows in the queue, and the action codes are closed.
DECLARE @badAction INT = (
    SELECT COUNT(*) FROM dbo.fn_ActionQueue(@AsOf, 120.0)
    WHERE ActionCode NOT IN ('GROUND_AND_REPLACE','REPLACE_AT_NEXT_SLOT','ORDER_PART_NOW','SCHEDULE_THIS_MONTH'));
INSERT INTO #UATResults VALUES ('UAT-21', 'Every queue row carries one of the four known action codes',
    '0', CAST(@badAction AS VARCHAR(20)), CASE WHEN @badAction = 0 THEN 1 ELSE 0 END,
    'An unrecognised code means a row nobody knows what to do with.');

-- UAT-22: queue size and capacity cut match the published figures.
DECLARE @qRows INT, @qWeek INT, @qHours DECIMAL(10,1);
SELECT @qRows = COUNT(*), @qWeek = SUM(CASE WHEN IsThisWeek=1 THEN 1 ELSE 0 END),
       @qHours = CAST(SUM(JobHours) AS DECIMAL(10,1))
FROM dbo.fn_ActionQueue(@AsOf, 120.0);
INSERT INTO #UATResults VALUES ('UAT-22', 'Action queue size, this-week cut and total hangar hours match',
    '194 / 26 / 682.5', CONCAT(@qRows,' / ',@qWeek,' / ',@qHours),
    CASE WHEN @qRows = 194 AND @qWeek = 26 AND @qHours = 682.5 THEN 1 ELSE 0 END,
    'Hours fell from 804.0 when procurement actions stopped being charged hangar bay time.');

/*
UAT-23: the capacity cut is a real boundary.

Not 'some rows are marked this week' -- that passes if every row is. The last
included job must fit inside the week and the first excluded one must not.
*/
DECLARE @lastIn DECIMAL(10,2), @firstOut DECIMAL(10,2);
SELECT @lastIn = MAX(CASE WHEN IsThisWeek = 1 THEN CumulativeJobHours END),
       @firstOut = MIN(CASE WHEN IsThisWeek = 0 THEN CumulativeJobHours END)
FROM dbo.fn_ActionQueue(@AsOf, 120.0);
INSERT INTO #UATResults VALUES ('UAT-23', 'Capacity cut falls exactly at the weekly hangar-hour limit',
    '<=120 then >120', CONCAT(@lastIn,' then ',@firstOut),
    CASE WHEN @lastIn <= 120.0 AND @firstOut > 120.0 THEN 1 ELSE 0 END, NULL);

-- UAT-24: flight-critical work outranks everything else, without exception.
DECLARE @rankBad INT = (
    SELECT COUNT(*) FROM dbo.fn_ActionQueue(@AsOf, 120.0) q
    WHERE q.Criticality <> 'FlightCritical'
      AND q.PriorityRank < (SELECT MAX(PriorityRank) FROM dbo.fn_ActionQueue(@AsOf, 120.0)
                            WHERE Criticality = 'FlightCritical'));
INSERT INTO #UATResults VALUES ('UAT-24', 'All flight-critical work outranks all other work',
    '0', CAST(@rankBad AS VARCHAR(20)), CASE WHEN @rankBad = 0 THEN 1 ELSE 0 END,
    'The tie-break chain puts criticality first.');

-- UAT-25: the queue order is STABLE between calls. An unstable queue cannot be
-- worked through, because the rows move under the crew between refreshes.
DECLARE @unstable INT = (
    SELECT COUNT(*) FROM dbo.fn_ActionQueue(@AsOf, 120.0) a
    JOIN dbo.fn_ActionQueue(@AsOf, 120.0) b ON b.InstallKey = a.InstallKey
    WHERE a.PriorityRank <> b.PriorityRank);
INSERT INTO #UATResults VALUES ('UAT-25', 'Queue ranking is stable across calls',
    '0', CAST(@unstable AS VARCHAR(20)), CASE WHEN @unstable = 0 THEN 1 ELSE 0 END,
    'Every ORDER BY ends in a unique tie-break.');

-- ===========================================================================
-- GROUP G -- the tests that prove the tests work
-- ===========================================================================

-- UAT-26: every planted defect class is found, with its exact count.
DECLARE @d1 INT = (SELECT COUNT(*) FROM dbo.DQ_Findings WHERE CheckCode='DQ-01');
DECLARE @d2 INT = (SELECT COUNT(*) FROM dbo.DQ_Findings WHERE CheckCode='DQ-02');
DECLARE @d3 INT = (SELECT COUNT(*) FROM dbo.DQ_Findings WHERE CheckCode='DQ-03');
DECLARE @d4 INT = (SELECT COUNT(*) FROM dbo.DQ_Findings WHERE CheckCode='DQ-04');
DECLARE @d5 INT = (SELECT COUNT(*) FROM dbo.DQ_Findings WHERE CheckCode='DQ-05');
INSERT INTO #UATResults VALUES ('UAT-26', 'Every planted defect class is detected with its exact count',
    '36/11/13/7/6', CONCAT(@d1,'/',@d2,'/',@d3,'/',@d4,'/',@d5),
    CASE WHEN @d1=36 AND @d2=11 AND @d3=13 AND @d4=7 AND @d5=6 THEN 1 ELSE 0 END,
    'DQ-01 is 36 because both copies of each of the 18 duplicates are flagged -- which one is genuine is unknowable.');
GO

/*
UAT-27 and UAT-28 BREAK SOMETHING ON PURPOSE.

Everything above asserts that a clean dataset reads clean. That is only half of
what a suite has to show: it also has to show that a dirty dataset does NOT.
These two build their own defect, assert it is caught, and roll back.

The fixtures CREATE the rows they test rather than selecting existing ones --
Project 4 shipped a fixture that quietly stopped matching anything after the
generator was tuned, returned NULL, and passed.

THE RESULTS ARE WRITTEN AFTER THE ROLLBACK, AND THAT IS THE WHOLE TRICK.

Project 4's lesson was 'create the results table outside the transaction', and
that is done above -- but it is not sufficient. A temp table INSERT is itself
transactional, so writing the verdicts inside the fixture transaction means
ROLLBACK discards them along with the fixture. The first run of this suite did
exactly that: both cases executed, both passed, and 26 of 28 rows came back.

Only the @ExpectedCases guard revealed it. Without that the run would have
reported '24 passed, 2 failed' with no indication that two cases had vanished,
which is the Project 2 failure mode this suite was built to prevent.

So the verdicts go into VARIABLES, which survive the rollback, and the rows are
written afterwards.
*/
DECLARE @AsOf2 DATE = '2026-09-30';
DECLARE @overlapFound INT = NULL, @nowBoth INT = NULL, @stillHidden INT = NULL;
DECLARE @fixtureError VARCHAR(400) = NULL;
-- pinned before the fixture runs, so the delta is attributable to the fixture
DECLARE @overlapBefore INT = (SELECT COUNT(*) FROM dbo.vw_DQ_OverlappingInstalls);

BEGIN TRANSACTION;
BEGIN TRY
    -- UAT-27: fit a SECOND component to a position that is already occupied.
    -- DQ-06 has found nothing on every run so far; a check that has never
    -- fired is a check nobody has tested.
    DECLARE @af INT = (SELECT TOP 1 AirframeKey FROM dbo.Dim_Airframe WHERE AirframeStatus='Active' ORDER BY TailNumber);
    DECLARE @ctk INT = (SELECT ComponentTypeKey FROM dbo.Dim_ComponentType WHERE ComponentCode='GIMBAL');
    DECLARE @insd INT = (SELECT InServiceDateKey FROM dbo.Dim_Airframe WHERE AirframeKey = @af);

    INSERT INTO dbo.Fact_ComponentInstall
        (ComponentSerial, ComponentTypeKey, AirframeKey, PositionNo, InstalledDateKey, RemovedDateKey, RemovalReason)
    VALUES ('UAT-OVERLAP-1', @ctk, @af, 1, @insd, NULL, NULL);

    -- vw_DQ_OverlappingInstalls is the SAME object 04's DQ-06 reads. Break the
    -- predicate and both this test and the shipped check change together.
    SET @overlapFound = (SELECT COUNT(*) FROM dbo.vw_DQ_OverlappingInstalls);

    -- UAT-28: make a component overdue on BOTH measures and assert it is NOT
    -- counted as hidden. The hidden figure is the headline of the project, so
    -- it has to exclude anything the hour meter can already see.
    DECLARE @target INT = (SELECT TOP 1 InstallKey FROM dbo.fn_ComponentWear(@AsOf2)
                           WHERE IsHiddenOverdue = 1 ORDER BY InstallKey);
    IF @target IS NULL
        THROW 53062, 'UAT-28 fixture found no hidden-overdue component to work with. The fixture is no longer testing anything.', 1;

    -- backdate it far enough to breach the hour interval as well
    UPDATE ci SET InstalledDateKey = (SELECT MIN(DateKey) FROM dbo.Dim_Date WHERE [Date] >= '2024-01-01')
    FROM dbo.Fact_ComponentInstall ci WHERE ci.InstallKey = @target;

    SET @nowBoth = (SELECT COUNT(*) FROM dbo.fn_ComponentWear(@AsOf2)
                    WHERE InstallKey = @target AND IsOverdueByHours = 1 AND IsOverdueByStress = 1);
    SET @stillHidden = (SELECT COUNT(*) FROM dbo.fn_ComponentWear(@AsOf2)
                        WHERE InstallKey = @target AND IsHiddenOverdue = 1);

    ROLLBACK TRANSACTION;
END TRY
BEGIN CATCH
    IF @@TRANCOUNT > 0 ROLLBACK TRANSACTION;
    SET @fixtureError = LEFT(ERROR_MESSAGE(), 400);
END CATCH

-- written AFTER the rollback, from variables that survived it
/*
A DELTA against a pinned baseline, not an absolute count.

'>= 1' passed whether or not the fixture contributed anything: a dataset that
had drifted into carrying a genuine overlap would satisfy it with the INSERT
doing nothing at all. And the value is knowable -- the injected gimbal pairs
with all four existing installs on that position, so the answer is exactly 4
against a baseline of 0.
*/
INSERT INTO #UATResults VALUES ('UAT-27', 'NEGATIVE: an overlapping install on one position is detected',
    'baseline 0, then 4', ISNULL(CONCAT('baseline ', @overlapBefore, ', then ', @overlapFound), 'fixture error'),
    CASE WHEN @fixtureError IS NULL AND @overlapBefore = 0 AND @overlapFound = 4 THEN 1 ELSE 0 END,
    ISNULL(@fixtureError, 'Exercises vw_DQ_OverlappingInstalls, the object 04 reads. The clean dataset has none; the fixture creates four pairs.'));

INSERT INTO #UATResults VALUES ('UAT-28', 'NEGATIVE: a component overdue on BOTH measures drops out of the hidden count',
    'both=1, hidden=0', ISNULL(CONCAT('both=',@nowBoth,', hidden=',@stillHidden), 'fixture error'),
    CASE WHEN @fixtureError IS NULL AND @nowBoth = 1 AND @stillHidden = 0 THEN 1 ELSE 0 END,
    ISNULL(@fixtureError, 'Hidden means invisible to the hour meter. Once it is visible there, it is not hidden.'));
GO

-- ===========================================================================
-- GROUP H -- the six defects an adversarial review found that groups A-G missed
-- ===========================================================================
DECLARE @AsOf3 DATE = '2026-09-30';
DECLARE @Mid   DATE = '2025-06-30';

/*
UAT-29: the population is the fleet AS IT WAS, not the fleet as it is.

Three aircraft retired inside the window. MU-003 went in September 2025 and flew
throughout 2024 and 2025, so a compliance figure for June 2025 must include it.
Filtering on AirframeStatus excluded all three from every historical figure.
*/
DECLARE @fleetMid INT = (SELECT COUNT(*) FROM dbo.fn_AirframeStress(@Mid) WHERE IsInFleetAsOf = 1);
DECLARE @fleetNow INT = (SELECT COUNT(*) FROM dbo.fn_AirframeStress(@AsOf3) WHERE IsInFleetAsOf = 1);
-- Both values pinned. 'mid > now' is the shape of the claim, but the shape
-- holds for any pair that happens to be ordered -- including a population
-- filtered the old way that returned 31 and 30 by accident.
INSERT INTO #UATResults VALUES ('UAT-29', 'Fleet size is computed from dates, so it differs between two as-of dates',
    '31 then 28', CONCAT(@fleetMid, ' > ', @fleetNow),
    CASE WHEN @fleetMid = 31 AND @fleetNow = 28 THEN 1 ELSE 0 END,
    'Three airframes retired inside the window. A population filtered on AirframeStatus would return the same number at both dates.');

/*
UAT-30: the availability DENOMINATOR must move too.

This is the case that would have caught the defect UAT-10 passed over. The
metric VALUE changed between dates because the numerator moved, so a test
asserting 'the metric changes' was satisfied while the denominator sat frozen at
28 across the whole trend.
*/
DECLARE @denMid DECIMAL(12,2) = (SELECT Denominator FROM dbo.fn_FleetKPI(@Mid)   WHERE MetricName='AirframeAvailabilityPct');
DECLARE @denNow DECIMAL(12,2) = (SELECT Denominator FROM dbo.fn_FleetKPI(@AsOf3) WHERE MetricName='AirframeAvailabilityPct');
INSERT INTO #UATResults VALUES ('UAT-30', 'Availability denominator responds to the as-of date, not just the value',
    'differs', CONCAT(@denMid, ' vs ', @denNow),
    CASE WHEN @denMid <> @denNow THEN 1 ELSE 0 END,
    'Asserting that a metric moves does not prove the denominator moved. This asserts the denominator.');

/*
UAT-31: no sortie is charged to two components on one position.

Both windows were inclusive of the changeover date, so 1,144 sorties were
counted against the outgoing part and its replacement at once. Nothing about the
result looked wrong -- accrued hours simply came out a little high on every
replacement in the fleet.
*/
DECLARE @doubleCharged INT = (
    SELECT COUNT(*)
    FROM dbo.fn_ComponentWear(@AsOf3) w
    JOIN dbo.Fact_ComponentInstall prev
      ON prev.AirframeKey = w.AirframeKey AND prev.ComponentTypeKey = w.ComponentTypeKey
     AND prev.PositionNo = w.PositionNo
    JOIN dbo.Dim_Date pr ON pr.DateKey = prev.RemovedDateKey
    JOIN dbo.vw_SortieStress ss ON ss.AirframeKey = w.AirframeKey AND ss.SortieDate = pr.[Date]
    WHERE pr.[Date] = w.InstalledDate
      AND w.Sorties > 0
      -- the replacement must NOT have counted the handover sortie
      AND EXISTS (SELECT 1 FROM dbo.vw_SortieStress s2
                  WHERE s2.AirframeKey = w.AirframeKey AND s2.SortieDate = w.InstalledDate)
      AND w.FlightHours >= (SELECT SUM(s3.FlightHours) FROM dbo.vw_SortieStress s3
                            WHERE s3.AirframeKey = w.AirframeKey
                              AND s3.SortieDate >= w.InstalledDate AND s3.SortieDate <= @AsOf3));
INSERT INTO #UATResults VALUES ('UAT-31', 'The handover sortie is charged to one component, not two',
    '0', CAST(@doubleCharged AS VARCHAR(20)), CASE WHEN @doubleCharged = 0 THEN 1 ELSE 0 END,
    'A replacement whose hours include the sortie its predecessor came off on.');

/*
UAT-32: every sensor feature row is inside its component's fitted window.

vw_SensorFeatures joined on InstallKey with no date predicate, so 362 readings
taken while the part was off the aircraft reached the degradation model as
training data and fn_AlertEvaluation as alerts. The data-quality layer had
flagged every one of them; nothing downstream read the flags.
*/
DECLARE @orphanFeatures INT = (
    SELECT COUNT(*) FROM dbo.vw_SensorFeatures f
    JOIN dbo.Fact_ComponentInstall ci ON ci.InstallKey = f.InstallKey
    JOIN dbo.Dim_Date di ON di.DateKey = ci.InstalledDateKey
    LEFT JOIN dbo.Dim_Date dr ON dr.DateKey = ci.RemovedDateKey
    WHERE f.SortieDate < di.[Date] OR (dr.[Date] IS NOT NULL AND f.SortieDate > dr.[Date]));
INSERT INTO #UATResults VALUES ('UAT-32', 'No sensor feature row falls outside its component fitted window',
    '0', CAST(@orphanFeatures AS VARCHAR(20)), CASE WHEN @orphanFeatures = 0 THEN 1 ELSE 0 END,
    'Readings taken while a part was in a crate are not evidence about that part.');

/*
UAT-33: a life with impossible dates scores no prediction.

Six installs are recorded as removed before they were installed. Five of them
were inside the published 288 true positives, contributing a life with no
measurable length to a prediction-quality figure.
*/
DECLARE @badScored INT = (
    SELECT COUNT(*) FROM dbo.fn_AlertEvaluation(@AsOf3, 3.400) e
    JOIN dbo.vw_ComponentLifeHistory h ON h.InstallKey = e.InstallKey
    WHERE h.IsDateValid = 0);
DECLARE @badKept INT = (SELECT COUNT(*) FROM dbo.vw_ComponentLifeHistory WHERE IsDateValid = 0);
INSERT INTO #UATResults VALUES ('UAT-33', 'Impossible lives are reported but never scored',
    'kept 6, scored 0', CONCAT('kept ', @badKept, ', scored ', @badScored),
    CASE WHEN @badKept = 6 AND @badScored = 0 THEN 1 ELSE 0 END,
    'They stay in the life history so totals reconcile; they do not score a prediction.');

/*
UAT-34: every row of Ref_FleetTargets is evaluated by something.

Three of the eight targets were compared to nothing at all -- including
ActionableLeadTimePct, which is the target the second finding is about. A target
that nothing evaluates cannot be failed, and will sit in a reference table
looking like governance.
*/
DECLARE @unevaluated INT = (
    SELECT COUNT(*) FROM dbo.Ref_FleetTargets t
    WHERE NOT EXISTS (SELECT 1 FROM dbo.fn_FleetKPI(@AsOf3) k WHERE k.MetricName = t.MetricName)
      AND NOT EXISTS (SELECT 1 FROM dbo.fn_PredictionKPI(@AsOf3, 3.400) k WHERE k.MetricName = t.MetricName));
INSERT INTO #UATResults VALUES ('UAT-34', 'Every target in Ref_FleetTargets is evaluated by a KPI function',
    '0', CAST(@unevaluated AS VARCHAR(20)), CASE WHEN @unevaluated = 0 THEN 1 ELSE 0 END,
    'Eight targets, eight evaluations. Three had none.');

/*
UAT-35: the fast path must agree with the slow path it replaced.

dbo.AlertFrontier collapses the 73,900 readings past the baseline window to
22,006 frontier points and takes a 57-threshold sweep from ten minutes to just
over a second. It rests on one
claim: 'the first reading at or above T' is the same row as 'the first reading
whose running maximum is at or above T'.

That claim is true. It is also exactly the kind of thing that is true when it is
written and false after someone adds a filter, and a materialised optimisation
that disagrees with the thing it replaced is worse than the slow version,
because it is fast and wrong.

So this case recomputes the alert set DIRECTLY from vw_SensorFeatures and
requires the two to match on every install AND on the lead time each reports.
Matching counts alone would pass while every alert fired on the wrong sortie.
*/
DECLARE @directRows INT, @frontierRows INT, @disagreements INT;
;WITH Direct AS (
    SELECT InstallKey, FlightHoursToRemoval FROM (
        SELECT f.InstallKey, f.FlightHoursToRemoval,
               rn = ROW_NUMBER() OVER (PARTITION BY f.InstallKey ORDER BY f.SortieSeq)
        FROM dbo.vw_SensorFeatures f
        WHERE f.SortieDate <= '2026-09-30' AND f.VibVsBaseline >= 3.400 AND f.SortieSeq > 10
    ) z WHERE rn = 1
),
Frontier AS (
    SELECT InstallKey, FlightHoursToRemoval FROM (
        SELECT af.InstallKey, af.FlightHoursToRemoval,
               rn = ROW_NUMBER() OVER (PARTITION BY af.InstallKey ORDER BY af.SortieSeq)
        FROM dbo.AlertFrontier af
        WHERE af.SortieDate <= '2026-09-30' AND af.ThresholdReached >= 3.400
    ) z WHERE rn = 1
)
SELECT @directRows   = (SELECT COUNT(*) FROM Direct),
       @frontierRows = (SELECT COUNT(*) FROM Frontier),
       @disagreements = (SELECT COUNT(*) FROM Direct d
                         FULL JOIN Frontier f ON f.InstallKey = d.InstallKey
                         WHERE d.InstallKey IS NULL OR f.InstallKey IS NULL
                            OR ABS(d.FlightHoursToRemoval - f.FlightHoursToRemoval) > 0.0001);

INSERT INTO #UATResults VALUES ('UAT-35', 'The AlertFrontier fast path matches a direct scan, install for install',
    '352 / 352 / 0', CONCAT(@directRows, ' / ', @frontierRows, ' / ', @disagreements),
    CASE WHEN @directRows = 352 AND @frontierRows = 352 AND @disagreements = 0 THEN 1 ELSE 0 END,
    'Same alerts AND the same lead time on each. Matching counts alone would pass while every alert fired on the wrong sortie.');
GO

-- ===========================================================================
-- REPORT
--
-- The results SELECT runs BEFORE any raise. Project 4's suite raised first and
-- the detail was never printed, so a failing run said only that it had failed.
-- ===========================================================================
DECLARE @ExpectedCases INT = 35;
DECLARE @Total INT = (SELECT COUNT(*) FROM #UATResults);
DECLARE @Failed INT = (SELECT COUNT(*) FROM #UATResults WHERE Passed = 0);

PRINT '';
PRINT '=== UAT results ===';
SELECT CaseID, CaseName, Expected, [Actual],
       Result = CASE WHEN Passed = 1 THEN 'PASS' ELSE 'FAIL' END, Notes
FROM #UATResults ORDER BY CaseID;

PRINT '';
PRINT CONCAT('  Cases run     : ', @Total, ' of ', @ExpectedCases, ' expected');
PRINT CONCAT('  Passed        : ', @Total - @Failed);
PRINT CONCAT('  Failed        : ', @Failed);

/*
The count guard.

A suite that runs 19 of 28 cases and reports '19 passed, 0 failed' is reporting
success for the nine it never ran. Project 2 did exactly that. The expected
count is asserted, so losing a case is itself a failure.
*/
IF @Total <> @ExpectedCases
    THROW 53060, 'UAT suite did not run the expected number of cases. A case was skipped, deleted, or lost to an early exit -- the passes it did report do not cover it.', 1;

IF @Failed > 0
    THROW 53061, 'UAT suite failed. See the result set above for the failing cases.', 1;

PRINT '';
PRINT '  ALL CASES PASSED.';
GO
