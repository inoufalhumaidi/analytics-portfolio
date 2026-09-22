/*
================================================================================
Project 4 — Talon Robotics: Payload Deployment Delivery
Script:  03_core_views.sql
Purpose: The verification-currency model. Everything downstream reads from here.

WHAT THIS LAYER DECIDES

    For one requirement, at one point in the programme:
      - is there any passing evidence at all?
      - does that evidence meet the verification policy for its type?
      - is that evidence still CURRENT, or has the ground moved under it?

    "Current" is the whole project. A verification is current when nothing it
    depended on has changed since it ran:

        STALE BY CODE  a later build changed the same subsystem
        STALE BY REQ   the requirement was Modified (not merely Clarified)
                       after the run

    Note what is NOT a staleness rule: elapsed time. A verification from
    fourteen months ago is perfectly good if nothing underneath it has moved,
    and one from last Tuesday is worthless if the subsystem was rewritten on
    Wednesday. Age is a proxy people reach for because it is easy to compute;
    intervening change is the actual question.

WHY EVERYTHING IS @AsOf-PARAMETERISED

    The same question has a different answer at every programme milestone, and
    a review board will ask "what did this look like at the last gate?". A view
    hardcoded to today can only answer once. Every function here takes @AsOfBuild
    -- the build being assessed -- so the whole model can be run against the
    release candidate, against last quarter's drop, or against any build in
    between, and the answers are comparable.

    @AsOfBuild is a BUILD NUMBER, not a date. Two builds can share a date, and
    a build date can be corrected after the fact; the sequence is what
    "superseded by" means.

DATA DISCLOSURE: Talon Robotics is fictional; all data is synthetic. No
confidential data and no production-system claim is involved.
================================================================================
*/

USE TalonDelivery;
GO

IF OBJECT_ID('dbo.vw_RequirementVerification', 'V')   IS NOT NULL DROP VIEW dbo.vw_RequirementVerification;
IF OBJECT_ID('dbo.fn_RequirementVerification', 'IF')  IS NOT NULL DROP FUNCTION dbo.fn_RequirementVerification;
IF OBJECT_ID('dbo.fn_LatestRunPerCase', 'IF')         IS NOT NULL DROP FUNCTION dbo.fn_LatestRunPerCase;
GO

/*
--------------------------------------------------------------------------------
fn_LatestRunPerCase(@AsOfBuild)

One row per test case: its most recent run at or before @AsOfBuild.

"Most recent" is by BUILD NUMBER, then by run date, then by key. The tie-break
chain is not decoration -- a case can be run twice against the same build on
the same day (the generator plants exactly that as a data defect), and without
a total ordering the "latest" run is whichever row the engine happened to
return. That is the class of bug that produces a different answer on a
different day with no code change.
--------------------------------------------------------------------------------
*/
CREATE FUNCTION dbo.fn_LatestRunPerCase (@AsOfBuild INT)
RETURNS TABLE
AS RETURN
(
    WITH Ranked AS (
        SELECT tr.TestRunKey, tr.TestCaseKey, tr.BuildKey, tr.RunDateKey,
               tr.RunByKey, tr.Result, tr.DurationMin,
               b.BuildNumber,
               rn = ROW_NUMBER() OVER (
                        PARTITION BY tr.TestCaseKey
                        ORDER BY b.BuildNumber DESC, tr.RunDateKey DESC, tr.TestRunKey DESC)
        FROM dbo.Fact_TestRun tr
        JOIN dbo.Dim_Build b ON b.BuildKey = tr.BuildKey
        WHERE b.BuildNumber <= @AsOfBuild
    )
    SELECT TestRunKey, TestCaseKey, BuildKey, BuildNumber, RunDateKey,
           RunByKey, Result, DurationMin
    FROM Ranked WHERE rn = 1
);
GO

/*
--------------------------------------------------------------------------------
fn_RequirementVerification(@AsOfBuild)

The heart of the project. One row per baselined requirement.

Withdrawn requirements are excluded here rather than filtered downstream. A
withdrawn requirement is not an unverified requirement -- counting it as one
understates readiness and, worse, puts work on a queue that nobody should do.
Excluding it once, at the point where the population is defined, means no
downstream view can forget to.
--------------------------------------------------------------------------------
*/
CREATE FUNCTION dbo.fn_RequirementVerification (@AsOfBuild INT)
RETURNS TABLE
AS RETURN
(
    WITH Req AS (
        SELECT r.RequirementKey, r.RequirementID, r.Title, r.ReqType, r.Priority,
               r.SubsystemKey, r.OwnerKey, r.BaselinedDateKey,
               s.SubsystemCode, s.SubsystemName, s.Criticality, s.RequiresHILRig,
               p.MinTestLevel, p.MinTestLevelRank, p.RequiresIndependentTester
        FROM dbo.Dim_Requirement r
        JOIN dbo.Dim_Subsystem s          ON s.SubsystemKey = r.SubsystemKey
        JOIN dbo.Ref_VerificationPolicy p ON p.ReqType      = r.ReqType
        WHERE r.ReqStatus = 'Baselined'
    ),
    -- Every latest-run-per-case, joined back to its requirement and graded
    -- against the policy. A run is SUFFICIENT evidence when it passed, at or
    -- above the required test level, and -- where the policy demands it -- by
    -- somebody other than the requirement's owner.
    CaseEvidence AS (
        SELECT q.RequirementKey,
               lr.BuildNumber,
               lr.Result,
               lr.RunByKey,
               tlr.LevelRank,
               IsPass    = CASE WHEN lr.Result = 'Pass' THEN 1 ELSE 0 END,
               IsBlocked = CASE WHEN lr.Result = 'Blocked' THEN 1 ELSE 0 END,
               MeetsLevel = CASE WHEN tlr.LevelRank >= q.MinTestLevelRank THEN 1 ELSE 0 END,
               MeetsIndependence = CASE
                    WHEN q.RequiresIndependentTester = 0 THEN 1
                    WHEN lr.RunByKey <> q.OwnerKey       THEN 1
                    ELSE 0 END
        FROM Req q
        JOIN dbo.Dim_TestCase tc ON tc.RequirementKey = q.RequirementKey
        JOIN dbo.fn_LatestRunPerCase(@AsOfBuild) lr ON lr.TestCaseKey = tc.TestCaseKey
        JOIN dbo.Ref_TestLevelRank tlr ON tlr.TestLevel = tc.TestLevel
    ),
    Agg AS (
        SELECT RequirementKey,
               CasesWithRuns   = COUNT(*),
               PassingCases    = SUM(IsPass),
               BlockedCases    = SUM(IsBlocked),
               -- the newest build on which ANY case passed -- evidence exists
               LastPassBuild   = MAX(CASE WHEN IsPass = 1 THEN BuildNumber END),
               -- the newest build on which a SUFFICIENT case passed -- evidence
               -- that also satisfies the policy. These two differ exactly where
               -- a requirement is "verified" by something that does not count.
               LastGoodBuild   = MAX(CASE WHEN IsPass = 1 AND MeetsLevel = 1
                                           AND MeetsIndependence = 1 THEN BuildNumber END),
               PassingButUnderLevelled = SUM(CASE WHEN IsPass = 1 AND MeetsLevel = 0 THEN 1 ELSE 0 END),
               PassingButSelfVerified  = SUM(CASE WHEN IsPass = 1 AND MeetsIndependence = 0 THEN 1 ELSE 0 END)
        FROM CaseEvidence
        GROUP BY RequirementKey
    )
    SELECT
        q.RequirementKey, q.RequirementID, q.Title, q.ReqType, q.Priority,
        q.SubsystemKey, q.SubsystemCode, q.SubsystemName, q.Criticality, q.RequiresHILRig,
        q.OwnerKey,
        q.MinTestLevel, q.RequiresIndependentTester,

        TotalCases   = ISNULL(tc.TotalCases, 0),
        CasesWithRuns= ISNULL(a.CasesWithRuns, 0),
        PassingCases = ISNULL(a.PassingCases, 0),
        BlockedCases = ISNULL(a.BlockedCases, 0),
        a.LastPassBuild,
        a.LastGoodBuild,
        PassingButUnderLevelled = ISNULL(a.PassingButUnderLevelled, 0),
        PassingButSelfVerified  = ISNULL(a.PassingButSelfVerified, 0),

        -- 1. Is there evidence at all?
        HasEvidence = CAST(CASE WHEN a.LastPassBuild IS NOT NULL THEN 1 ELSE 0 END AS BIT),

        -- 2. Does the evidence satisfy the verification policy?
        MeetsPolicy = CAST(CASE WHEN a.LastGoodBuild IS NOT NULL THEN 1 ELSE 0 END AS BIT),

        -- 3. Has anything moved since the evidence was captured?
        --
        -- Evaluated against LastGoodBuild, not LastPassBuild: the question is
        -- whether the evidence THAT COUNTS is still current. Measuring
        -- staleness against a run that never satisfied the policy would let an
        -- insufficient test mask an out-of-date sufficient one.
        StaleByCode = CAST(CASE WHEN a.LastGoodBuild IS NOT NULL AND EXISTS (
                            SELECT 1
                            FROM dbo.Fact_BuildSubsystemChange c
                            JOIN dbo.Dim_Build b2 ON b2.BuildKey = c.BuildKey
                            WHERE c.SubsystemKey = q.SubsystemKey
                              AND b2.BuildNumber > a.LastGoodBuild
                              AND b2.BuildNumber <= @AsOfBuild)
                        THEN 1 ELSE 0 END AS BIT),

        StaleByRequirement = CAST(CASE WHEN a.LastGoodBuild IS NOT NULL AND EXISTS (
                            SELECT 1
                            FROM dbo.Fact_RequirementChange rc
                            JOIN dbo.Dim_Build b3 ON b3.BuildKey = rc.BuildKey
                            WHERE rc.RequirementKey = q.RequirementKey
                              -- Clarified deliberately does NOT invalidate.
                              -- Rewording for a certification reviewer does not
                              -- change behaviour, and treating it as staleness
                              -- would inflate the finding and cost it its
                              -- credibility on the one thing it must be right
                              -- about.
                              AND rc.ChangeType = 'Modified'
                              AND b3.BuildNumber > a.LastGoodBuild
                              AND b3.BuildNumber <= @AsOfBuild)
                        THEN 1 ELSE 0 END AS BIT),

        -- How much has moved, for the queue to rank by. A requirement whose
        -- subsystem changed once is a re-run; one whose subsystem changed
        -- eleven times is a re-think.
        InterveningBuilds = ISNULL((
            SELECT COUNT(*)
            FROM dbo.Fact_BuildSubsystemChange c
            JOIN dbo.Dim_Build b4 ON b4.BuildKey = c.BuildKey
            WHERE c.SubsystemKey = q.SubsystemKey
              AND b4.BuildNumber > a.LastGoodBuild
              AND b4.BuildNumber <= @AsOfBuild), 0),

        AsOfBuild = @AsOfBuild
    FROM Req q
    LEFT JOIN Agg a ON a.RequirementKey = q.RequirementKey
    OUTER APPLY (SELECT TotalCases = COUNT(*)
                 FROM dbo.Dim_TestCase t2 WHERE t2.RequirementKey = q.RequirementKey) tc
);
GO

/*
--------------------------------------------------------------------------------
vw_RequirementVerification -- the model at the release candidate.

Pinned to the RC rather than to MAX(BuildNumber): the release candidate is a
decision somebody made, and readiness is always assessed against the thing
being shipped. If a later experimental build appears, the answer to "can we
ship RC1" must not silently change.
--------------------------------------------------------------------------------
*/
CREATE VIEW dbo.vw_RequirementVerification AS
SELECT v.*,
       -- Published here so every consumer applies the same definition rather
       -- than each re-deriving "ready" from the three flags and eventually
       -- disagreeing about it.
       IsCurrent = CAST(CASE WHEN v.MeetsPolicy = 1
                              AND v.StaleByCode = 0
                              AND v.StaleByRequirement = 0
                         THEN 1 ELSE 0 END AS BIT),
       VerificationState = CASE
            WHEN v.HasEvidence = 0                        THEN 'No evidence'
            WHEN v.MeetsPolicy = 0                        THEN 'Insufficient evidence'
            WHEN v.StaleByRequirement = 1                 THEN 'Stale - requirement changed'
            WHEN v.StaleByCode = 1                        THEN 'Stale - subsystem changed'
            ELSE                                               'Current' END
FROM dbo.fn_RequirementVerification(
        (SELECT TOP 1 BuildNumber FROM dbo.Dim_Build WHERE IsReleaseCandidate = 1 ORDER BY BuildNumber DESC)
     ) v;
GO

-- =============================================================================
-- Creation gate
-- =============================================================================
DECLARE @missing VARCHAR(300) = '';
IF OBJECT_ID('dbo.fn_LatestRunPerCase','IF')        IS NULL SET @missing += 'fn_LatestRunPerCase ';
IF OBJECT_ID('dbo.fn_RequirementVerification','IF') IS NULL SET @missing += 'fn_RequirementVerification ';
IF OBJECT_ID('dbo.vw_RequirementVerification','V')  IS NULL SET @missing += 'vw_RequirementVerification ';
IF @missing <> '' THROW 52020, 'FAILED to create core objects -- scroll up for the compile error.', 1;

-- Sanity: every baselined requirement appears exactly once, and the four
-- verification states partition the population with nothing left over.
--
-- FIVE, not four -- 'No evidence', 'Insufficient evidence', 'Stale -
-- requirement changed', 'Stale - subsystem changed', 'Current'. And note what
-- the gate below actually does: it compares two ROW COUNTS. It never looks at
-- VerificationState, so it would pass unchanged if the CASE fell through to
-- NULL for an entire state. UAT-01 is what asserts the partition.
DECLARE @reqs INT = (SELECT COUNT(*) FROM dbo.Dim_Requirement WHERE ReqStatus = 'Baselined');
DECLARE @rows INT = (SELECT COUNT(*) FROM dbo.vw_RequirementVerification);
IF @reqs <> @rows
    THROW 52021, 'vw_RequirementVerification does not return exactly one row per baselined requirement.', 1;

PRINT 'Core objects created and verified: fn_LatestRunPerCase, fn_RequirementVerification(@AsOfBuild), vw_RequirementVerification.';
GO
