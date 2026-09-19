/*
================================================================================
Project 4 — Talon Robotics: Payload Deployment Delivery
Script:  07_uat_test_cases.sql
Purpose: User acceptance tests. Runs inside a transaction and rolls back, so
         the programme dataset is never modified. Exits non-zero on failure.

HOW THESE CASES ARE WRITTEN, AND WHY IT MATTERS

    Three projects in this portfolio shipped defects past a green suite, and
    every one got through the same way: the test asserted a PROPERTY of the
    answer rather than its VALUE.

        "the index is at least 100"        -- true under any positive weighting,
                                              including the join fan-out it was
                                              written to catch
        "exposure is between 0 and the weighted total"
                                           -- guaranteed by an explicit clamp

    So the rule here is: wherever a case can recompute the value independently
    -- ideally at a DIFFERENT GRAIN from the implementation, because the grain
    is usually what the implementation got wrong -- it does. Bounds checks are
    kept where they are genuinely informative, but never alone.

    Several cases also plant a fixture designed to make the code fail if a
    specific rule were removed. A regression test that cannot fail against the
    old code is not a regression test.

DATA DISCLOSURE: Talon Robotics is fictional; all data is synthetic. No
confidential data and no production-system claim is involved.
================================================================================
*/

USE TalonDelivery;
GO
SET NOCOUNT ON;

BEGIN TRANSACTION UATRun;

IF OBJECT_ID('tempdb..#UATResults') IS NOT NULL DROP TABLE #UATResults;
CREATE TABLE #UATResults (
    TestID VARCHAR(10), Area VARCHAR(30), Requirement VARCHAR(220),
    Expected VARCHAR(90), Actual VARCHAR(90), Verdict VARCHAR(6), WhyItMatters VARCHAR(500));

DECLARE @RC INT = (SELECT TOP 1 BuildNumber FROM dbo.Dim_Build WHERE IsReleaseCandidate = 1 ORDER BY BuildNumber DESC);

-- Reproducibility fingerprints are captured HERE, before any fixture exists.
--
-- The alternative -- computing them at the end and subtracting the fixture
-- rows -- couples a reproducibility assertion to how many fixtures happen to
-- be inserted above it, so adding an unrelated test case silently breaks it.
-- Capturing first makes the assertion independent of the rest of the suite.
DECLARE @fpRun INT = (
    SELECT CHECKSUM_AGG(CAST(TestCaseKey * 31 + BuildKey * 7
         + CASE Result WHEN 'Pass' THEN 1 WHEN 'Fail' THEN 2 ELSE 3 END AS INT))
    FROM dbo.Fact_TestRun);
DECLARE @fpRAID INT = (
    SELECT CHECKSUM_AGG(CAST(SubsystemKey * 13 + Probability * 5 + Impact AS INT))
    FROM dbo.Fact_RAID);

/* ---------------------------------------------------------------------------
UAT-01  The four verification states partition the population exactly.
--------------------------------------------------------------------------- */
DECLARE @reqs INT = (SELECT COUNT(*) FROM dbo.Dim_Requirement WHERE ReqStatus = 'Baselined');
DECLARE @stated INT = (SELECT COUNT(*) FROM dbo.vw_RequirementVerification
                       WHERE VerificationState IN ('No evidence','Insufficient evidence',
                             'Stale - requirement changed','Stale - subsystem changed','Current'));
INSERT INTO #UATResults VALUES ('UAT-01','Verification model',
 'Every baselined requirement lands in exactly one of the five verification states',
 CAST(@reqs AS VARCHAR(20)), CAST(@stated AS VARCHAR(20)),
 CASE WHEN @reqs = @stated THEN 'PASS' ELSE 'FAIL' END,
 'An unclassified requirement is invisible to both the scorecard and the queue, so it is neither reported as a risk nor scheduled for work.');

/* ---------------------------------------------------------------------------
UAT-02  IsCurrent equals the recomputed definition, requirement by requirement.

        Not a bounds check. The three flags are recombined here from the raw
        facts rather than from the view's own columns, so an error in how the
        view assembles them is visible.
--------------------------------------------------------------------------- */
DECLARE @currentMismatch INT;
;WITH Recomputed AS (
    SELECT r.RequirementKey,
           ShouldBeCurrent = CASE
               WHEN lg.LastGoodBuild IS NULL THEN 0
               WHEN EXISTS (SELECT 1 FROM dbo.Fact_BuildSubsystemChange c
                            JOIN dbo.Dim_Build b ON b.BuildKey = c.BuildKey
                            WHERE c.SubsystemKey = r.SubsystemKey
                              AND b.BuildNumber > lg.LastGoodBuild AND b.BuildNumber <= @RC) THEN 0
               WHEN EXISTS (SELECT 1 FROM dbo.Fact_RequirementChange rc
                            JOIN dbo.Dim_Build b2 ON b2.BuildKey = rc.BuildKey
                            WHERE rc.RequirementKey = r.RequirementKey AND rc.ChangeType = 'Modified'
                              AND b2.BuildNumber > lg.LastGoodBuild AND b2.BuildNumber <= @RC) THEN 0
               ELSE 1 END
    FROM dbo.Dim_Requirement r
    OUTER APPLY (
        SELECT LastGoodBuild = MAX(b3.BuildNumber)
        FROM dbo.Fact_TestRun tr
        JOIN dbo.Dim_TestCase tc  ON tc.TestCaseKey = tr.TestCaseKey
        JOIN dbo.Dim_Build b3     ON b3.BuildKey    = tr.BuildKey
        JOIN dbo.Ref_TestLevelRank tlr ON tlr.TestLevel = tc.TestLevel
        JOIN dbo.Ref_VerificationPolicy vp ON vp.ReqType = r.ReqType
        WHERE tc.RequirementKey = r.RequirementKey
          AND tr.Result = 'Pass'
          AND b3.BuildNumber <= @RC
          AND tlr.LevelRank >= vp.MinTestLevelRank
          AND (vp.RequiresIndependentTester = 0 OR tr.RunByKey <> r.OwnerKey)
          -- only the LATEST run of each case counts as standing evidence
          AND tr.TestRunKey = (SELECT TOP 1 tr2.TestRunKey FROM dbo.Fact_TestRun tr2
                               JOIN dbo.Dim_Build b4 ON b4.BuildKey = tr2.BuildKey
                               WHERE tr2.TestCaseKey = tc.TestCaseKey AND b4.BuildNumber <= @RC
                               ORDER BY b4.BuildNumber DESC, tr2.RunDateKey DESC, tr2.TestRunKey DESC)
    ) lg
    WHERE r.ReqStatus = 'Baselined'
)
SELECT @currentMismatch = COUNT(*)
FROM dbo.vw_RequirementVerification v
JOIN Recomputed x ON x.RequirementKey = v.RequirementKey
WHERE CAST(v.IsCurrent AS INT) <> x.ShouldBeCurrent;

INSERT INTO #UATResults VALUES ('UAT-02','Verification model',
 'IsCurrent equals the definition recomputed from the raw facts, for every requirement',
 '0 mismatches', CAST(@currentMismatch AS VARCHAR(20)),
 CASE WHEN @currentMismatch = 0 THEN 'PASS' ELSE 'FAIL' END,
 'This is the number the ship decision rests on. Recomputing it from Fact_TestRun rather than from the view''s own flags is what makes the case capable of failing.');

/* ---------------------------------------------------------------------------
UAT-03  fn_LatestRunPerCase returns exactly one row per case, and it is the
        highest build -- with a deterministic tie-break.

        A fixture plants two runs of the same case against the same build on
        the same date. Without a total ordering the "latest" is whichever row
        the engine returns, which is the class of bug that gives a different
        answer on a different day with no code change.
--------------------------------------------------------------------------- */
DECLARE @tcFix INT = (SELECT TOP 1 TestCaseKey FROM dbo.Dim_TestCase ORDER BY TestCaseKey);
DECLARE @bFix  INT = (SELECT TOP 1 BuildKey FROM dbo.Dim_Build WHERE BuildNumber = 10);
DECLARE @dFix  INT = (SELECT TOP 1 BuildDateKey FROM dbo.Dim_Build WHERE BuildNumber = 10);
DECLARE @pFix  INT = (SELECT TOP 1 PersonKey FROM dbo.Dim_Person ORDER BY PersonKey);

INSERT INTO dbo.Fact_TestRun (TestCaseKey, BuildKey, RunDateKey, RunByKey, Result, DurationMin)
VALUES (@tcFix, @bFix, @dFix, @pFix, 'Fail', 30), (@tcFix, @bFix, @dFix, @pFix, 'Fail', 30);

DECLARE @dupRows INT = (SELECT COUNT(*) FROM dbo.fn_LatestRunPerCase(@RC) WHERE TestCaseKey = @tcFix);
DECLARE @caseRows INT, @casesWithRuns INT;
SELECT @caseRows = COUNT(*) FROM dbo.fn_LatestRunPerCase(@RC);
SELECT @casesWithRuns = COUNT(DISTINCT tr.TestCaseKey)
FROM dbo.Fact_TestRun tr JOIN dbo.Dim_Build b ON b.BuildKey = tr.BuildKey WHERE b.BuildNumber <= @RC;

INSERT INTO #UATResults VALUES ('UAT-03','Evidence selection',
 'Exactly one latest run per test case, even when two runs share a build and a date',
 '1 for the fixture, ' + CAST(@casesWithRuns AS VARCHAR(20)) + ' overall',
 CAST(@dupRows AS VARCHAR(10)) + ' / ' + CAST(@caseRows AS VARCHAR(20)),
 CASE WHEN @dupRows = 1 AND @caseRows = @casesWithRuns THEN 'PASS' ELSE 'FAIL' END,
 'Without a total ordering the latest run is whichever row the engine happens to return, and the readiness answer changes between runs with no code change.');

/* ---------------------------------------------------------------------------
UAT-04  As-of integrity: no evidence from after the assessed build leaks in.
--------------------------------------------------------------------------- */
DECLARE @leak INT = (
    SELECT COUNT(*) FROM dbo.fn_LatestRunPerCase(40) lr
    JOIN dbo.Dim_Build b ON b.BuildKey = lr.BuildKey WHERE b.BuildNumber > 40);
DECLARE @leakGood INT = (
    SELECT COUNT(*) FROM dbo.fn_RequirementVerification(40) WHERE LastGoodBuild > 40);
INSERT INTO #UATResults VALUES ('UAT-04','As-of integrity',
 'Assessing build 40 uses no run and no verification from a later build',
 '0 / 0', CAST(@leak AS VARCHAR(10)) + ' / ' + CAST(@leakGood AS VARCHAR(10)),
 CASE WHEN @leak = 0 AND @leakGood = 0 THEN 'PASS' ELSE 'FAIL' END,
 'A review board asks what the position was at the last gate. A model that leaks later evidence answers a flattering question nobody asked.');

/* ---------------------------------------------------------------------------
UAT-05  As-of integrity for WORK ITEMS -- regression.

        fn_ReadinessKPI originally counted every work item regardless of date,
        so WorkItemCompletionPct was identical at every build in the trend: a
        flat line across twenty months of a programme that was demonstrably
        closing items throughout. A function that takes a point-in-time
        parameter and answers about "now" for one of its outputs is worse than
        one that never offered the parameter.

        Against the old code the two figures below are equal and this fails.
--------------------------------------------------------------------------- */
DECLARE @wiEarly DECIMAL(9,2) = (SELECT WorkItemCompletionPct FROM dbo.fn_ReadinessKPI(40, '2025-10-01'));
DECLARE @wiLate  DECIMAL(9,2) = (SELECT WorkItemCompletionPct FROM dbo.fn_ReadinessKPI(@RC, '2026-09-30'));
INSERT INTO #UATResults VALUES ('UAT-05','As-of integrity',
 'Work-item completion differs between an early date and the reporting date',
 'early < late', CAST(@wiEarly AS VARCHAR(20)) + ' < ' + CAST(@wiLate AS VARCHAR(20)),
 CASE WHEN @wiEarly < @wiLate THEN 'PASS' ELSE 'FAIL' END,
 'The metric is reported on a trend. If it ignores its own as-of date it draws a flat line and hides the only thing a trend is for.');

/* ---------------------------------------------------------------------------
UAT-06  A Clarified requirement change does NOT invalidate evidence; a
        Modified one does.

        Two fixtures on two currently-Current requirements. This is the
        distinction the whole staleness claim rests on: treating a reworded
        sentence as invalidated evidence would inflate the finding and cost it
        credibility on the part that is real.
--------------------------------------------------------------------------- */
-- The fixture builds its own build, and depends on nothing the generator
-- happened to produce.
--
-- An earlier version picked two requirements that were Current with evidence
-- predating the release candidate. That worked until the regression rates were
-- tuned, after which EVERY Current requirement drew its evidence from the RC
-- regression at build 88 -- anything verified earlier has had its subsystem
-- move since -- so the selection returned nothing and the case reported NULL
-- rather than failing on its subject. A fixture that depends on the shape of
-- the data is one parameter change away from testing nothing.
--
-- Build 89 changes NO subsystem. That is the point: at build 89 the only thing
-- that can invalidate a verification is a requirement change, which is exactly
-- the rule under test, with every other cause held out by construction.
DECLARE @bLate INT, @dLate INT, @pAny INT;
SELECT @bLate = BuildKey, @dLate = BuildDateKey FROM dbo.Dim_Build WHERE BuildNumber = @RC;
SELECT TOP 1 @pAny = PersonKey FROM dbo.Dim_Person ORDER BY PersonKey;

DECLARE @dFuture INT = (SELECT DateKey FROM dbo.Dim_Date
    WHERE [Date] = DATEADD(DAY, 7, (SELECT [Date] FROM dbo.Dim_Date WHERE DateKey = @dLate)));
INSERT INTO dbo.Dim_Build (BuildID, BuildNumber, BuildDateKey, BuildLabel, IsReleaseCandidate)
VALUES ('B-089', 89, @dFuture, 'UAT fixture build - no subsystem changes', 0);
DECLARE @bNext INT = SCOPE_IDENTITY();

DECLARE @rkClar INT, @rkMod INT;
SELECT TOP 1 @rkClar = RequirementKey FROM dbo.vw_RequirementVerification
 WHERE IsCurrent = 1 ORDER BY RequirementKey;
SELECT TOP 1 @rkMod = RequirementKey FROM dbo.vw_RequirementVerification
 WHERE IsCurrent = 1 AND RequirementKey <> @rkClar ORDER BY RequirementKey DESC;

INSERT INTO dbo.Fact_RequirementChange (RequirementKey, ChangeDateKey, BuildKey, ChangeType, ChangedByKey, ChangeNote)
VALUES (@rkClar, @dFuture, @bNext, 'Clarified', @pAny, 'UAT fixture: wording only.'),
       (@rkMod,  @dFuture, @bNext, 'Modified',  @pAny, 'UAT fixture: acceptance threshold changed.');

DECLARE @clarStill INT = (
    SELECT CASE WHEN MeetsPolicy = 1 AND StaleByCode = 0 AND StaleByRequirement = 0 THEN 1 ELSE 0 END
    FROM dbo.fn_RequirementVerification(89) WHERE RequirementKey = @rkClar);
DECLARE @modNow INT = (
    SELECT CASE WHEN MeetsPolicy = 1 AND StaleByCode = 0 AND StaleByRequirement = 0 THEN 1 ELSE 0 END
    FROM dbo.fn_RequirementVerification(89) WHERE RequirementKey = @rkMod);
INSERT INTO #UATResults VALUES ('UAT-06','Staleness rules',
 'Clarified leaves a verification current; Modified invalidates it',
 'clarified 1, modified 0', CAST(@clarStill AS VARCHAR(5)) + ' / ' + CAST(@modNow AS VARCHAR(5)),
 CASE WHEN @clarStill = 1 AND @modNow = 0 THEN 'PASS' ELSE 'FAIL' END,
 'If a reworded sentence counted as invalidated evidence the staleness figure would be inflated, and the programme would rightly stop believing the part that is real.');

/* ---------------------------------------------------------------------------
UAT-07  Verification policy: a test below the required level is not evidence.
--------------------------------------------------------------------------- */
DECLARE @underLevelCurrent INT = (
    SELECT COUNT(*) FROM dbo.vw_RequirementVerification v
    JOIN dbo.Ref_VerificationPolicy p ON p.ReqType = v.ReqType
    WHERE v.IsCurrent = 1 AND v.MeetsPolicy = 1
      AND NOT EXISTS (
          SELECT 1 FROM dbo.Dim_TestCase tc
          JOIN dbo.Ref_TestLevelRank tlr ON tlr.TestLevel = tc.TestLevel
          WHERE tc.RequirementKey = v.RequirementKey AND tlr.LevelRank >= p.MinTestLevelRank));
INSERT INTO #UATResults VALUES ('UAT-07','Verification policy',
 'No requirement is counted current without at least one test case at or above its required level',
 '0', CAST(@underLevelCurrent AS VARCHAR(10)),
 CASE WHEN @underLevelCurrent = 0 THEN 'PASS' ELSE 'FAIL' END,
 'A safety behaviour signed off by a unit test is the failure mode the policy table exists to prevent, and it looks exactly like a pass on every dashboard.');

/* ---------------------------------------------------------------------------
UAT-08  Verification policy: the independent-tester rule is enforced.

        Fixture: give a Safety requirement a passing run performed by its own
        owner at the release candidate, and confirm that alone does not make it
        current.
--------------------------------------------------------------------------- */
DECLARE @rkSafe INT, @tcSafe INT, @ownerSafe INT;
SELECT TOP 1 @rkSafe = v.RequirementKey, @ownerSafe = v.OwnerKey
FROM dbo.vw_RequirementVerification v
JOIN dbo.Dim_TestCase tc ON tc.RequirementKey = v.RequirementKey
JOIN dbo.Ref_TestLevelRank tlr ON tlr.TestLevel = tc.TestLevel
WHERE v.ReqType = 'Safety' AND v.IsCurrent = 0 AND tlr.LevelRank >= 3
ORDER BY v.RequirementKey;
SELECT TOP 1 @tcSafe = tc.TestCaseKey
FROM dbo.Dim_TestCase tc JOIN dbo.Ref_TestLevelRank tlr ON tlr.TestLevel = tc.TestLevel
WHERE tc.RequirementKey = @rkSafe AND tlr.LevelRank >= 3 ORDER BY tc.TestCaseKey;

INSERT INTO dbo.Fact_TestRun (TestCaseKey, BuildKey, RunDateKey, RunByKey, Result, DurationMin)
VALUES (@tcSafe, @bLate, @dLate, @ownerSafe, 'Pass', 60);

DECLARE @selfVerifiedCurrent INT = (SELECT CAST(IsCurrent AS INT) FROM dbo.vw_RequirementVerification WHERE RequirementKey = @rkSafe);
INSERT INTO #UATResults VALUES ('UAT-08','Verification policy',
 'A Safety requirement passed at the right level by its OWN OWNER is still not current',
 '0', CAST(@selfVerifiedCurrent AS VARCHAR(5)),
 CASE WHEN @selfVerifiedCurrent = 0 THEN 'PASS' ELSE 'FAIL' END,
 'Independence is half the policy. A rule enforced only on test level would pass this fixture, which is why the fixture is at the right level and fails only on the author.');

/* ---------------------------------------------------------------------------
UAT-09  Withdrawn requirements are out of scope, not unverified.
--------------------------------------------------------------------------- */
DECLARE @withdrawnInModel INT = (
    SELECT COUNT(*) FROM dbo.vw_RequirementVerification v
    JOIN dbo.Dim_Requirement r ON r.RequirementKey = v.RequirementKey
    WHERE r.ReqStatus = 'Withdrawn');
DECLARE @withdrawnInQueue INT = (
    SELECT COUNT(*) FROM dbo.vw_VerificationQueue q
    JOIN dbo.Dim_Requirement r ON r.RequirementID = q.RequirementID
    WHERE r.ReqStatus = 'Withdrawn');
INSERT INTO #UATResults VALUES ('UAT-09','Population',
 'Withdrawn requirements appear in neither the verification model nor the queue',
 '0 / 0', CAST(@withdrawnInModel AS VARCHAR(10)) + ' / ' + CAST(@withdrawnInQueue AS VARCHAR(10)),
 CASE WHEN @withdrawnInModel = 0 AND @withdrawnInQueue = 0 THEN 'PASS' ELSE 'FAIL' END,
 'Counting a withdrawn requirement as unverified understates readiness and puts work on the queue that nobody should do.');

/* ---------------------------------------------------------------------------
UAT-10  The queue holds exactly the non-current requirements.
--------------------------------------------------------------------------- */
DECLARE @notCurrent INT = (SELECT COUNT(*) FROM dbo.vw_RequirementVerification WHERE IsCurrent = 0);
DECLARE @queued     INT = (SELECT COUNT(*) FROM dbo.vw_VerificationQueue);
DECLARE @queueMismatch INT = (
    SELECT COUNT(*) FROM dbo.vw_RequirementVerification v
    FULL JOIN dbo.vw_VerificationQueue q ON q.RequirementID = v.RequirementID AND v.IsCurrent = 0
    WHERE (v.IsCurrent = 0 AND q.RequirementID IS NULL) OR (q.RequirementID IS NOT NULL AND v.RequirementID IS NULL));
INSERT INTO #UATResults VALUES ('UAT-10','Priority queue',
 'The queue contains every non-current requirement and nothing else',
 CAST(@notCurrent AS VARCHAR(10)) + ' / 0 mismatched',
 CAST(@queued AS VARCHAR(10)) + ' / ' + CAST(@queueMismatch AS VARCHAR(10)),
 CASE WHEN @notCurrent = @queued AND @queueMismatch = 0 THEN 'PASS' ELSE 'FAIL' END,
 'A queue that silently drops work is worse than no queue: the gap still exists but nobody is looking for it.');

/* ---------------------------------------------------------------------------
UAT-11  Every queued requirement carries one recognised action and an
        instruction written for a person.
--------------------------------------------------------------------------- */
DECLARE @badAction INT = (
    SELECT COUNT(*) FROM dbo.vw_VerificationQueue
    WHERE ActionCode NOT IN ('VERIFY','RERUN','REVIEW_THEN_RERUN','RAISE_TEST_LEVEL','INDEPENDENT_WITNESS')
       OR RecommendedAction IS NULL OR LEN(RecommendedAction) < 40);
INSERT INTO #UATResults VALUES ('UAT-11','Priority queue',
 'Every queued item has a recognised action code and a written instruction',
 '0', CAST(@badAction AS VARCHAR(10)),
 CASE WHEN @badAction = 0 THEN 'PASS' ELSE 'FAIL' END,
 'An action code without an instruction makes the reader guess what RERUN means for a requirement whose wording changed, and the guess is usually wrong.');

/* ---------------------------------------------------------------------------
UAT-12  The queue's action matches the requirement's actual state.

        Recomputed from the verification flags rather than trusting the CASE
        expression in the view.
--------------------------------------------------------------------------- */
DECLARE @actionMismatch INT = (
    SELECT COUNT(*)
    FROM dbo.vw_VerificationQueue q
    JOIN dbo.vw_RequirementVerification v ON v.RequirementID = q.RequirementID
    WHERE q.ActionCode <> CASE
            WHEN v.HasEvidence = 0 THEN 'VERIFY'
            WHEN v.MeetsPolicy = 0 AND v.PassingButSelfVerified > 0
                                   AND v.PassingButUnderLevelled = 0 THEN 'INDEPENDENT_WITNESS'
            WHEN v.MeetsPolicy = 0 THEN 'RAISE_TEST_LEVEL'
            WHEN v.StaleByRequirement = 1 THEN 'REVIEW_THEN_RERUN'
            ELSE 'RERUN' END);
INSERT INTO #UATResults VALUES ('UAT-12','Priority queue',
 'Each action code matches the requirement''s verification state, recomputed',
 '0 mismatches', CAST(@actionMismatch AS VARCHAR(10)),
 CASE WHEN @actionMismatch = 0 THEN 'PASS' ELSE 'FAIL' END,
 'Sending a team to re-run a test when the requirement itself was reworded certifies against the old wording. The action has to follow the state, not the other way round.');

/* ---------------------------------------------------------------------------
UAT-13  The week's work respects the rig-hour bound, and the bound bites.

        Asserts the VALUE of the boundary, not just that a flag exists: the
        last included item fits and the first excluded one does not.
--------------------------------------------------------------------------- */
DECLARE @overBound INT = (SELECT COUNT(*) FROM dbo.fn_VerificationQueue(@RC, 180.00)
                          WHERE IsThisWeek = 1 AND CumulativeRigHours > 180.00);
DECLARE @thisWeek INT = (SELECT COUNT(*) FROM dbo.fn_VerificationQueue(@RC, 180.00) WHERE IsThisWeek = 1);
DECLARE @thisWeekDouble INT = (SELECT COUNT(*) FROM dbo.fn_VerificationQueue(@RC, 360.00) WHERE IsThisWeek = 1);
INSERT INTO #UATResults VALUES ('UAT-13','Priority queue',
 'Nothing in the week exceeds the rig bound, and doubling the rig widens the week',
 '0 over, more work at 360h',
 CAST(@overBound AS VARCHAR(10)) + ' over, ' + CAST(@thisWeek AS VARCHAR(10)) + ' -> ' + CAST(@thisWeekDouble AS VARCHAR(10)),
 CASE WHEN @overBound = 0 AND @thisWeekDouble > @thisWeek THEN 'PASS' ELSE 'FAIL' END,
 'A capacity bound that does not respond to capacity is a hardcoded row count wearing a parameter, and it cannot answer "what if we add a rig".');

/* ---------------------------------------------------------------------------
UAT-14  Subsystem readiness reconciles to the overall figure.

        Aggregation integrity at a DIFFERENT GRAIN from the scorecard.
--------------------------------------------------------------------------- */
DECLARE @overallCurrent INT = (SELECT COUNT(*) FROM dbo.vw_RequirementVerification WHERE Priority='MustShip' AND IsCurrent=1);
DECLARE @subsystemCurrent INT = (SELECT SUM(CurrentMustShip) FROM dbo.vw_SubsystemReadiness);
DECLARE @overallMustShip INT = (SELECT COUNT(*) FROM dbo.vw_RequirementVerification WHERE Priority='MustShip');
DECLARE @subsystemMustShip INT = (SELECT SUM(MustShip) FROM dbo.vw_SubsystemReadiness);
INSERT INTO #UATResults VALUES ('UAT-14','Aggregation',
 'Summing the subsystem cut reproduces the programme totals',
 CAST(@overallCurrent AS VARCHAR(10)) + ' / ' + CAST(@overallMustShip AS VARCHAR(10)),
 CAST(@subsystemCurrent AS VARCHAR(10)) + ' / ' + CAST(@subsystemMustShip AS VARCHAR(10)),
 CASE WHEN @overallCurrent = @subsystemCurrent AND @overallMustShip = @subsystemMustShip THEN 'PASS' ELSE 'FAIL' END,
 'The subsystem cut is what turns one number into a decision about where to send people. If it does not add up, the decision is being made on a different population from the headline.');

/* ---------------------------------------------------------------------------
UAT-15  ShipReadinessPct equals the figure recomputed from its components
        AS PUBLISHED, not from unrounded intermediates.
--------------------------------------------------------------------------- */
DECLARE @publishedReadiness DECIMAL(9,2) = (SELECT ShipReadinessPct FROM dbo.vw_ReadinessKPI);
DECLARE @recomputedReadiness DECIMAL(9,2) = (
    SELECT CAST(100.0 * SUM(CAST(IsCurrent AS INT)) / NULLIF(COUNT(*),0) AS DECIMAL(9,2))
    FROM dbo.vw_RequirementVerification WHERE Priority = 'MustShip');
INSERT INTO #UATResults VALUES ('UAT-15','Scorecard',
 'ShipReadinessPct equals the ratio recomputed at requirement grain',
 CAST(@publishedReadiness AS VARCHAR(20)), CAST(@recomputedReadiness AS VARCHAR(20)),
 CASE WHEN ABS(@publishedReadiness - @recomputedReadiness) < 0.005 THEN 'PASS' ELSE 'FAIL' END,
 'The ship gate is the one figure in this project that must be right to two decimal places, because 99.99% and 100% are different decisions.');

/* ---------------------------------------------------------------------------
UAT-16  Data quality: duplicates are found by BEHAVIOUR.

        The fixture inserted for UAT-03 is a genuine duplicate by the business
        signature, and the detector must find it without any marker to read.
--------------------------------------------------------------------------- */
DECLARE @dupFound INT = (SELECT COUNT(*) FROM dbo.fn_DuplicateTestRuns()
                         WHERE TestCaseKey = @tcFix AND BuildKey = @bFix);
INSERT INTO #UATResults VALUES ('UAT-16','Data quality',
 'The duplicate detector finds a planted duplicate from its business signature alone',
 '1', CAST(@dupFound AS VARCHAR(10)),
 CASE WHEN @dupFound = 1 THEN 'PASS' ELSE 'FAIL' END,
 'A detector that reads an identifier prefix works only against its own test fixture. Real data has no prefix, and the check would silently match nothing.');

/* ---------------------------------------------------------------------------
UAT-17  Data quality: every catalogued check is reported, including the clean
        ones.
--------------------------------------------------------------------------- */
DECLARE @catalogued INT = (SELECT COUNT(*) FROM dbo.vw_DQ_CheckCatalog);
DECLARE @reported   INT = (SELECT COUNT(*) FROM dbo.vw_DQ_Summary);
DECLARE @cleanChecks INT = (SELECT COUNT(*) FROM dbo.vw_DQ_Summary WHERE Anomalies = 0);
INSERT INTO #UATResults VALUES ('UAT-17','Data quality',
 'Every catalogued check appears in the summary, including those that found nothing',
 CAST(@catalogued AS VARCHAR(10)) + ' with some clean',
 CAST(@reported AS VARCHAR(10)) + ' with ' + CAST(@cleanChecks AS VARCHAR(10)) + ' clean',
 CASE WHEN @catalogued = @reported AND @cleanChecks > 0 THEN 'PASS' ELSE 'FAIL' END,
 'A summary listing only what was found cannot distinguish "we looked and it was clean" from "we never looked", and those are very different things to tell a review board.');

/* ---------------------------------------------------------------------------
UAT-18  Data quality: policy breaches are not counted as data defects.
--------------------------------------------------------------------------- */
DECLARE @policyCounted INT = (
    SELECT COUNT(*) FROM dbo.vw_DQ_CheckCatalog
    WHERE ImpactClass = 'Policy' AND CountsTowardExposure = 1);
INSERT INTO #UATResults VALUES ('UAT-18','Data quality',
 'No Policy-class check contributes to the data-quality exposure gate',
 '0', CAST(@policyCounted AS VARCHAR(10)),
 CASE WHEN @policyCounted = 0 THEN 'PASS' ELSE 'FAIL' END,
 'A safety requirement verified by its own owner is accurate data recording a bad practice. Counting it as a data defect makes both problems harder to fix and lets the gate be dismissed as noise.');

/* ---------------------------------------------------------------------------
UAT-19  RAID exposure comes from the matrix, not from probability x impact.

        The matrix deliberately distorts the corners -- 1x5 scores 10 while
        5x1 scores 5, though the product is identical -- so a formula would
        disagree with it and this case detects that.
--------------------------------------------------------------------------- */
DECLARE @matrixMismatch INT = (
    SELECT COUNT(*) FROM dbo.vw_RAIDExposure r
    JOIN dbo.Ref_RAIDMatrix m ON m.Probability = r.Probability AND m.Impact = r.Impact
    WHERE r.ExposureScore <> m.ExposureScore OR r.ExposureBand <> m.ExposureBand);
DECLARE @cornerDiffers INT = (
    SELECT CASE WHEN (SELECT ExposureScore FROM dbo.Ref_RAIDMatrix WHERE Probability=1 AND Impact=5)
               <> (SELECT ExposureScore FROM dbo.Ref_RAIDMatrix WHERE Probability=5 AND Impact=1)
           THEN 1 ELSE 0 END);
INSERT INTO #UATResults VALUES ('UAT-19','RAID',
 'Exposure is read from the matrix, and the matrix is not a multiplication',
 '0 mismatches, corners differ',
 CAST(@matrixMismatch AS VARCHAR(10)) + ' / ' + CAST(@cornerDiffers AS VARCHAR(5)),
 CASE WHEN @matrixMismatch = 0 AND @cornerDiffers = 1 THEN 'PASS' ELSE 'FAIL' END,
 'A board treats "unlikely but catastrophic" differently from "certain but trivial" even though the product is the same. A formula cannot express that, which is why the matrix is a table.');

/* ---------------------------------------------------------------------------
UAT-20  RAID overdue excludes items due before they were raised.
--------------------------------------------------------------------------- */
DECLARE @badOverdue INT = (
    SELECT COUNT(*) FROM dbo.vw_RAIDExposure WHERE IsOverdue = 1 AND DueDate < RaisedDate);
DECLARE @dueBeforeRaised INT = (
    SELECT COUNT(*) FROM dbo.vw_RAIDExposure WHERE DueDate < RaisedDate);
INSERT INTO #UATResults VALUES ('UAT-20','RAID',
 'Items due before they were raised are a data defect, not an overdue item',
 '0 counted overdue, > 0 present',
 CAST(@badOverdue AS VARCHAR(10)) + ' / ' + CAST(@dueBeforeRaised AS VARCHAR(10)),
 CASE WHEN @badOverdue = 0 AND @dueBeforeRaised > 0 THEN 'PASS' ELSE 'FAIL' END,
 'Such an item is overdue on the day it is created. Counting it inflates the overdue rate with a typing error and blames the programme for it.');

/* ---------------------------------------------------------------------------
UAT-21  The stored-procedure interface raises on invalid arguments.
--------------------------------------------------------------------------- */
DECLARE @r1 VARCHAR(5)='', @r2 VARCHAR(5)='', @r3 VARCHAR(5)='';
BEGIN TRY EXEC dbo.usp_ReadinessScorecard @AsOfBuild = 99999; SET @r1='NO'; END TRY BEGIN CATCH SET @r1='YES'; END CATCH
BEGIN TRY EXEC dbo.usp_VerificationQueue @RigHoursPerWeek = 0; SET @r2='NO'; END TRY BEGIN CATCH SET @r2='YES'; END CATCH
BEGIN TRY EXEC dbo.usp_ReadinessTrend @FromBuild = 80, @ToBuild = 20; SET @r3='NO'; END TRY BEGIN CATCH SET @r3='YES'; END CATCH
INSERT INTO #UATResults VALUES ('UAT-21','Interface',
 'Procedures raise on an unknown build, a zero rig capacity and a reversed range',
 'YES / YES / YES', @r1 + ' / ' + @r2 + ' / ' + @r3,
 CASE WHEN @r1='YES' AND @r2='YES' AND @r3='YES' THEN 'PASS' ELSE 'FAIL' END,
 'An out-of-range build returning an empty result set is indistinguishable from a build where nothing was outstanding. One is a typo, the other is a milestone.');

/* ---------------------------------------------------------------------------
UAT-22  Reproducibility: regenerating reproduces the documented fingerprints.
--------------------------------------------------------------------------- */
INSERT INTO #UATResults VALUES ('UAT-22','Reproducibility',
 'Fingerprints match the documented values for a hash-keyed rebuild',
 '8806 / 135', CAST(@fpRun AS VARCHAR(20)) + ' / ' + CAST(@fpRAID AS VARCHAR(20)),
 CASE WHEN @fpRun = 8806 AND @fpRAID = 135 THEN 'PASS' ELSE 'FAIL' END,
 'Every figure quoted in the case study is only checkable if the dataset can be rebuilt byte for byte. Hash-based pseudo-randomness gives that; RAND() and NEWID() do not.');

/* ---------------------------------------------------------------------------
UAT-23  RAID openness is as-of correct -- regression.

        fn_RAIDExposure derived IsOpen from the stored RAIDStatus, so it
        reported TODAY's position whatever @AsOfDate said. An item closed last
        month showed as closed in a register run for last quarter, when it was
        demonstrably open at the time: at 2025-06-30 it reported 12 open
        against 18 genuinely open.

        It was harmless at the reporting date, because every close in this
        dataset predates it -- which is exactly why it survived. The same
        defect shape as the work-item leak in UAT-05, in a different function.

        The check is run at a HISTORICAL date on purpose. Run at the reporting
        date it passes against the broken code too, and would be a test that
        cannot fail for the reason it was written.
--------------------------------------------------------------------------- */
DECLARE @histDate DATE = '2025-06-30';
DECLARE @raidReported INT = (SELECT COUNT(*) FROM dbo.fn_RAIDExposure(@histDate) WHERE IsOpen = 1);
DECLARE @raidTruly INT = (
    SELECT COUNT(*)
    FROM dbo.Fact_RAID x
    JOIN dbo.Dim_Date rd      ON rd.DateKey = x.RaisedDateKey
    LEFT JOIN dbo.Dim_Date cd ON cd.DateKey = x.ClosedDateKey
    WHERE rd.[Date] <= @histDate
      AND (cd.[Date] IS NULL OR cd.[Date] > @histDate));
-- and the overdue flag must agree with the same openness, not with the status
DECLARE @raidOverdueClosed INT = (
    SELECT COUNT(*) FROM dbo.fn_RAIDExposure(@histDate) WHERE IsOverdue = 1 AND IsOpen = 0);

INSERT INTO #UATResults VALUES ('UAT-23','RAID',
 'At a historical date the register reports the items open THEN, not the ones open now',
 CAST(@raidTruly AS VARCHAR(10)) + ' open, 0 overdue-but-closed',
 CAST(@raidReported AS VARCHAR(10)) + ' / ' + CAST(@raidOverdueClosed AS VARCHAR(10)),
 CASE WHEN @raidReported = @raidTruly AND @raidOverdueClosed = 0 THEN 'PASS' ELSE 'FAIL' END,
 'A procedure advertising an as-of date that silently answers about today is worse than one that never offered the parameter: the caller has no way to know the answer is not the one they asked for.');

/* ------------------------------ report ---------------------------------- */
SELECT TestID, Area, Requirement, Expected, Actual, Verdict FROM #UATResults ORDER BY TestID;

-- How many cases this suite is supposed to contain. A run that dies partway
-- through still reaches this report and prints a pass count covering only the
-- cases that executed, which is indistinguishable from a clean run of a
-- shorter suite. Counting what ran is not the same as counting what should
-- have run.
DECLARE @ExpectedCases INT = 23;
DECLARE @ran INT = (SELECT COUNT(*) FROM #UATResults);
IF @ran <> @ExpectedCases
BEGIN
    PRINT CONCAT('UAT HARNESS: only ', @ran, ' of ', @ExpectedCases,
                 ' cases recorded a result -- the suite did not run to completion.');
    ROLLBACK TRANSACTION UATRun;
    ;THROW 52070, 'UAT suite did not run to completion. Scroll up for the error that stopped it.', 1;
END

DECLARE @pass INT = (SELECT COUNT(*) FROM #UATResults WHERE Verdict = 'PASS');
DECLARE @fail INT = (SELECT COUNT(*) FROM #UATResults WHERE Verdict = 'FAIL');

-- The failure detail is selected BEFORE the rollback, not after.
--
-- #UATResults is created inside this transaction, so ROLLBACK drops it and
-- anything that reads it afterwards fails with "Invalid object name
-- '#UATResults'" -- masking the real failures behind a second, unrelated
-- error. The branch only executes when a case fails, which is exactly why a
-- suite that has always passed would never reveal it.
IF @fail > 0
    SELECT TestID, Requirement, Expected, Actual, WhyItMatters
    FROM #UATResults WHERE Verdict = 'FAIL' ORDER BY TestID;

PRINT '';
PRINT CONCAT('UAT RESULT: ', @pass, ' passed, ', @fail, ' failed.');

-- Roll back FIRST. The fixtures must leave the dataset whether the suite
-- passed or not, and a THROW before the ROLLBACK would leave them behind.
ROLLBACK TRANSACTION UATRun;
PRINT 'Test fixtures rolled back -- the programme dataset is unchanged.';

IF @fail > 0
    THROW 52071, 'UAT FAILED -- see the failing cases above. Do not publish readiness figures from this build.', 1;
PRINT 'All acceptance criteria met.';
GO
