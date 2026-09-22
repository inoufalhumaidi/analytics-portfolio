/*
================================================================================
Project 4 — Talon Robotics: Payload Deployment Delivery
Script:  05_kpi_views.sql
Purpose: The readiness scorecard, the RAID exposure model, and the priority
         action queue that turns both into a week of work.

THE TWO NUMBERS THIS FILE EXISTS TO PUT SIDE BY SIDE

    WorkItemCompletionPct  -- what the programme reports
    ShipReadinessPct       -- whether it can ship

    Neither is wrong. They answer different questions, and only one of them is
    the question on the agenda. Publishing them adjacent, from the same script,
    is the entire argument: a reader who sees 93.86% and 51.96% together does
    not need the rest of the analysis explained to them.

WHY THE QUEUE RANKS THE WAY IT DOES

    Ranking outstanding verification by requirement count sends a team to
    whatever is most numerous, which is Functional requirements on ground
    software. Ranking by criticality alone sends them to Safety work that may
    need a rig they cannot book this week.

    So the queue ranks by the VALUE of closing the gap -- criticality weighted
    by whether the requirement is must-ship -- and then bounds the result by
    the scarce resource that actually governs throughput: rig hours. A list of
    everything outstanding is not a plan, and a plan that ignores the rig
    booking is a wish.

DATA DISCLOSURE: Talon Robotics is fictional; all data is synthetic. No
confidential data and no production-system claim is involved.
================================================================================
*/

USE TalonDelivery;
GO

IF OBJECT_ID('dbo.vw_VerificationQueue','V')     IS NOT NULL DROP VIEW dbo.vw_VerificationQueue;
IF OBJECT_ID('dbo.fn_VerificationQueue','IF')    IS NOT NULL DROP FUNCTION dbo.fn_VerificationQueue;
IF OBJECT_ID('dbo.vw_SubsystemReadiness','V')    IS NOT NULL DROP VIEW dbo.vw_SubsystemReadiness;
IF OBJECT_ID('dbo.fn_SubsystemReadiness','IF')   IS NOT NULL DROP FUNCTION dbo.fn_SubsystemReadiness;
IF OBJECT_ID('dbo.vw_RAIDExposure','V')          IS NOT NULL DROP VIEW dbo.vw_RAIDExposure;
IF OBJECT_ID('dbo.fn_RAIDExposure','IF')         IS NOT NULL DROP FUNCTION dbo.fn_RAIDExposure;
IF OBJECT_ID('dbo.vw_ReadinessKPI','V')          IS NOT NULL DROP VIEW dbo.vw_ReadinessKPI;
IF OBJECT_ID('dbo.fn_ReadinessKPI','IF')         IS NOT NULL DROP FUNCTION dbo.fn_ReadinessKPI;
GO

/*
--------------------------------------------------------------------------------
fn_RAIDExposure(@AsOfDate) -- RAID scored against the matrix, not a formula.

@AsOfDate is a DATE here rather than a build number, because RAID items live on
the calendar: "overdue" is a statement about today, not about a build.
--------------------------------------------------------------------------------
*/
CREATE FUNCTION dbo.fn_RAIDExposure (@AsOfDate DATE)
RETURNS TABLE
AS RETURN
(
    SELECT
        x.RAIDKey, x.RAIDID, x.RAIDType, x.Title, x.RAIDStatus,
        s.SubsystemKey, s.SubsystemCode, s.SubsystemName, s.Criticality,
        p.PersonID AS OwnerID, p.PersonName AS OwnerName, p.Team AS OwnerTeam,
        RaisedDate = rd.[Date],
        DueDate    = dd.[Date],
        -- A close that has not happened yet at @AsOfDate has not happened.
        ClosedDate = CASE WHEN cd.[Date] <= @AsOfDate THEN cd.[Date] END,
        x.Probability, x.Impact,
        m.ExposureScore, m.ExposureBand,

        -- IsOpen is derived from the DATES, not from the stored status.
        --
        -- Reading RAIDStatus directly makes this function report today's
        -- position whatever @AsOfDate says: an item closed last month shows as
        -- closed in a register run for last quarter, when it was demonstrably
        -- open at the time. At 2025-06-30 that understated the open register by
        -- six items out of eighteen.
        --
        -- It happens to be harmless at the reporting date, because every close
        -- in this dataset predates it -- which is exactly why it survived. A
        -- function that takes a point-in-time parameter and answers about "now"
        -- for one of its outputs is the same defect that made
        -- WorkItemCompletionPct a flat line in the readiness trend.
        IsOpen = CAST(CASE WHEN cd.[Date] IS NULL OR cd.[Date] > @AsOfDate THEN 1 ELSE 0 END AS BIT),

        -- The status as it stood at @AsOfDate. An item closed AFTER the as-of
        -- date was open then; the stored status cannot say whether it was
        -- 'Open' or 'Mitigating', so it is reported as 'Open' -- the
        -- conservative reading, and the one that does not invent a mitigation
        -- nobody had started.
        RAIDStatusAsOf = CASE
            WHEN cd.[Date] IS NOT NULL AND cd.[Date] <= @AsOfDate THEN 'Closed'
            WHEN x.RAIDStatus = 'Closed'                          THEN 'Open'
            ELSE x.RAIDStatus END,
        -- An item due before it was raised is overdue on the day it is
        -- created. That is a data defect, reported by the data-quality layer,
        -- and counting it here as a management failure would blame the
        -- programme for a typing error.
        -- Overdue uses the same date-derived openness, not the stored status,
        -- so the two cannot disagree about whether an item was open.
        IsOverdue = CAST(CASE
            WHEN (cd.[Date] IS NULL OR cd.[Date] > @AsOfDate)
             AND dd.[Date] IS NOT NULL
             AND dd.[Date] < @AsOfDate
             AND dd.[Date] >= rd.[Date]
            THEN 1 ELSE 0 END AS BIT),
        DaysOverdue = CASE
            WHEN (cd.[Date] IS NULL OR cd.[Date] > @AsOfDate)
             AND dd.[Date] IS NOT NULL AND dd.[Date] < @AsOfDate AND dd.[Date] >= rd.[Date]
            THEN DATEDIFF(DAY, dd.[Date], @AsOfDate) ELSE 0 END,
        AgeDays = DATEDIFF(DAY, rd.[Date], @AsOfDate),
        x.MitigationNote,
        AsOfDate = @AsOfDate
    FROM dbo.Fact_RAID x
    JOIN dbo.Dim_Subsystem s  ON s.SubsystemKey = x.SubsystemKey
    JOIN dbo.Dim_Person p     ON p.PersonKey    = x.OwnerKey
    JOIN dbo.Dim_Date rd      ON rd.DateKey     = x.RaisedDateKey
    JOIN dbo.Ref_RAIDMatrix m ON m.Probability  = x.Probability AND m.Impact = x.Impact
    LEFT JOIN dbo.Dim_Date dd ON dd.DateKey     = x.DueDateKey
    LEFT JOIN dbo.Dim_Date cd ON cd.DateKey     = x.ClosedDateKey
    WHERE rd.[Date] <= @AsOfDate
);
GO

CREATE VIEW dbo.vw_RAIDExposure AS
SELECT * FROM dbo.fn_RAIDExposure('2026-09-30');
GO

/*
--------------------------------------------------------------------------------
fn_SubsystemReadiness(@AsOfBuild) -- readiness cut by subsystem.

This is where the headline stops being a single number and starts being
actionable: the programme is not 33% ready everywhere, it is nearly ready in
five subsystems and nowhere near it in two.
--------------------------------------------------------------------------------
*/
CREATE FUNCTION dbo.fn_SubsystemReadiness (@AsOfBuild INT)
RETURNS TABLE
AS RETURN
(
    WITH V AS (
        SELECT v.*,
               IsCurrent = CASE WHEN v.MeetsPolicy = 1 AND v.StaleByCode = 0
                                 AND v.StaleByRequirement = 0 THEN 1 ELSE 0 END
        FROM dbo.fn_RequirementVerification(@AsOfBuild) v
    ),
    Churn AS (
        SELECT c.SubsystemKey,
               BuildsChanged = COUNT(*),
               LinesChanged  = SUM(c.LinesChanged),
               LastChangeBuild = MAX(b.BuildNumber)
        FROM dbo.Fact_BuildSubsystemChange c
        JOIN dbo.Dim_Build b ON b.BuildKey = c.BuildKey
        WHERE b.BuildNumber <= @AsOfBuild
        GROUP BY c.SubsystemKey
    )
    SELECT
        s.SubsystemKey, s.SubsystemCode, s.SubsystemName, s.Criticality, s.RequiresHILRig,
        Requirements   = COUNT(*),
        MustShip       = SUM(CASE WHEN v.Priority = 'MustShip' THEN 1 ELSE 0 END),
        CurrentAll     = SUM(v.IsCurrent),
        CurrentMustShip= SUM(CASE WHEN v.Priority = 'MustShip' THEN v.IsCurrent ELSE 0 END),
        NoEvidence     = SUM(CASE WHEN v.HasEvidence = 0 THEN 1 ELSE 0 END),
        Insufficient   = SUM(CASE WHEN v.HasEvidence = 1 AND v.MeetsPolicy = 0 THEN 1 ELSE 0 END),
        /*
        THE PRECEDENCE HERE MUST MATCH vw_RequirementVerification, AND IT DID NOT.

        03_core_views.sql resolves StaleByRequirement FIRST, so a requirement
        stale on both counts reports as 'requirement changed' -- deliberately,
        because that one needs a systems engineer before anyone re-runs
        anything, and re-running against changed wording risks certifying the
        wrong thing.

        This cut resolved StaleByCode first and excluded it from StaleByReq, so
        all 7 requirements stale on both counts landed in StaleByCode and
        StaleByReq was 0 in every subsystem. The subsystem table therefore said
        95 / 0 while the requirement view, the verification queue
        (88 RERUN / 7 REVIEW_THEN_RERUN), the case study, the validation report
        and the Power BI measures all said 88 / 7.

        Both totals were 95, so nothing failed to reconcile at the top -- the
        disagreement was entirely in how the 95 split, which is the only part
        that decides who picks the work up.
        */
        StaleByReq     = SUM(CASE WHEN v.MeetsPolicy = 1 AND v.StaleByRequirement = 1 THEN 1 ELSE 0 END),
        StaleByCode    = SUM(CASE WHEN v.MeetsPolicy = 1 AND v.StaleByRequirement = 0
                                   AND v.StaleByCode = 1 THEN 1 ELSE 0 END),
        ReadinessPct   = CAST(100.0 * SUM(CASE WHEN v.Priority = 'MustShip' THEN v.IsCurrent ELSE 0 END)
                            / NULLIF(SUM(CASE WHEN v.Priority = 'MustShip' THEN 1 ELSE 0 END), 0) AS DECIMAL(6,2)),
        BuildsChanged   = ISNULL(ch.BuildsChanged, 0),
        LinesChanged    = ISNULL(ch.LinesChanged, 0),
        LastChangeBuild = ch.LastChangeBuild,
        -- Rig hours needed to re-verify everything not current in this
        -- subsystem. This is the number that turns "we are behind" into "we
        -- are behind by eleven rig-weeks", which is a different conversation.
        RigHoursOutstanding = ISNULL((
            SELECT CAST(SUM(tlr.RigHours) AS DECIMAL(9,2))
            FROM V v2
            JOIN dbo.Dim_TestCase tc ON tc.RequirementKey = v2.RequirementKey
            JOIN dbo.Ref_TestLevelRank tlr ON tlr.TestLevel = tc.TestLevel
            WHERE v2.SubsystemKey = s.SubsystemKey AND v2.IsCurrent = 0), 0),
        AsOfBuild = @AsOfBuild
    FROM V v
    JOIN dbo.Dim_Subsystem s ON s.SubsystemKey = v.SubsystemKey
    LEFT JOIN Churn ch ON ch.SubsystemKey = s.SubsystemKey
    GROUP BY s.SubsystemKey, s.SubsystemCode, s.SubsystemName, s.Criticality, s.RequiresHILRig,
             ch.BuildsChanged, ch.LinesChanged, ch.LastChangeBuild
);
GO

CREATE VIEW dbo.vw_SubsystemReadiness AS
SELECT * FROM dbo.fn_SubsystemReadiness(
    (SELECT TOP 1 BuildNumber FROM dbo.Dim_Build WHERE IsReleaseCandidate = 1 ORDER BY BuildNumber DESC));
GO

/*
--------------------------------------------------------------------------------
fn_ReadinessKPI(@AsOfBuild, @AsOfDate) -- the scorecard.
--------------------------------------------------------------------------------
*/
CREATE FUNCTION dbo.fn_ReadinessKPI (@AsOfBuild INT, @AsOfDate DATE)
RETURNS TABLE
AS RETURN
(
    WITH V AS (
        SELECT v.*,
               IsCurrent = CASE WHEN v.MeetsPolicy = 1 AND v.StaleByCode = 0
                                 AND v.StaleByRequirement = 0 THEN 1 ELSE 0 END
        FROM dbo.fn_RequirementVerification(@AsOfBuild) v
    ),
    R AS (SELECT * FROM dbo.fn_RAIDExposure(@AsOfDate)),
    -- As-of correct, and it has to be said explicitly because the obvious
    -- version is not. Counting every work item regardless of date makes
    -- WorkItemCompletionPct constant at every build in the trend -- a flat
    -- line across twenty months of a programme that was demonstrably closing
    -- items throughout. A function that takes a point-in-time parameter and
    -- then answers about "now" for one of its outputs is worse than one that
    -- never offered the parameter.
    --
    -- An item not yet opened does not exist. An item closed AFTER the as-of
    -- date was still open at the as-of date.
    W AS (
        SELECT Total = COUNT(*),
               Closed = SUM(CASE WHEN w.WorkItemStatus = 'Closed'
                                  AND cd.[Date] IS NOT NULL
                                  AND cd.[Date] <= @AsOfDate THEN 1 ELSE 0 END)
        FROM dbo.Fact_WorkItem w
        JOIN dbo.Dim_Date od      ON od.DateKey = w.OpenedDateKey
        LEFT JOIN dbo.Dim_Date cd ON cd.DateKey = w.ClosedDateKey
        WHERE od.[Date] <= @AsOfDate
    )
    SELECT
        AsOfBuild = @AsOfBuild,
        AsOfDate  = @AsOfDate,

        Requirements      = (SELECT COUNT(*) FROM V),
        MustShipCount     = (SELECT COUNT(*) FROM V WHERE Priority = 'MustShip'),

        -- The number the programme reports.
        WorkItemCompletionPct = (SELECT CAST(100.0 * Closed / NULLIF(Total,0) AS DECIMAL(9,2)) FROM W),

        -- "Is there evidence at all" -- deliberately generous, and it is the
        -- generosity that makes the contrast with readiness meaningful.
        VerificationCoveragePct = (SELECT CAST(100.0 * SUM(CAST(HasEvidence AS INT)) / NULLIF(COUNT(*),0) AS DECIMAL(9,2)) FROM V),

        -- The ship gate.
        ShipReadinessPct = (SELECT CAST(100.0 * SUM(IsCurrent) / NULLIF(COUNT(*),0) AS DECIMAL(9,2))
                            FROM V WHERE Priority = 'MustShip'),

        -- Of the requirements that HAVE policy-compliant evidence, the share
        -- whose evidence has been overtaken. Denominator is deliberately not
        -- all requirements: mixing "never verified properly" into a staleness
        -- rate makes staleness look like a documentation problem instead of
        -- an evidence one.
        StaleVerificationPct = (SELECT CAST(100.0 * SUM(CASE WHEN StaleByCode = 1 OR StaleByRequirement = 1 THEN 1 ELSE 0 END)
                                    / NULLIF(SUM(CAST(MeetsPolicy AS INT)),0) AS DECIMAL(9,2))
                                FROM V WHERE MeetsPolicy = 1),

        PolicyCompliancePct = (SELECT CAST(100.0 * SUM(CAST(MeetsPolicy AS INT)) / NULLIF(SUM(CAST(HasEvidence AS INT)),0) AS DECIMAL(9,2))
                               FROM V WHERE HasEvidence = 1),

        NoEvidenceCount     = (SELECT COUNT(*) FROM V WHERE HasEvidence = 0),
        InsufficientCount   = (SELECT COUNT(*) FROM V WHERE HasEvidence = 1 AND MeetsPolicy = 0),
        StaleCount          = (SELECT COUNT(*) FROM V WHERE MeetsPolicy = 1 AND (StaleByCode = 1 OR StaleByRequirement = 1)),
        CurrentCount        = (SELECT SUM(IsCurrent) FROM V),

        BlockedTestPct = (SELECT CAST(100.0 * SUM(CASE WHEN lr.Result = 'Blocked' THEN 1 ELSE 0 END)
                                   / NULLIF(COUNT(*),0) AS DECIMAL(9,2))
                          FROM dbo.fn_LatestRunPerCase(@AsOfBuild) lr),

        CriticalRAIDOpen = (SELECT COUNT(*) FROM R WHERE ExposureBand = 'Critical' AND IsOpen = 1),
        OverdueRAIDPct   = (SELECT CAST(100.0 * SUM(CAST(IsOverdue AS INT)) / NULLIF(SUM(CAST(IsOpen AS INT)),0) AS DECIMAL(9,2))
                            FROM R),
        TotalRAIDExposure = (SELECT SUM(ExposureScore) FROM R WHERE IsOpen = 1),

        -- Total rig hours to bring every non-current requirement back to
        -- current. Divided by weekly rig capacity this is the honest schedule
        -- answer, and it is the only figure here a programme board can act on
        -- without a further study.
        RigHoursOutstanding = (SELECT CAST(SUM(tlr.RigHours) AS DECIMAL(11,2))
                               FROM V v2
                               JOIN dbo.Dim_TestCase tc ON tc.RequirementKey = v2.RequirementKey
                               JOIN dbo.Ref_TestLevelRank tlr ON tlr.TestLevel = tc.TestLevel
                               WHERE v2.IsCurrent = 0)
);
GO

CREATE VIEW dbo.vw_ReadinessKPI AS
SELECT * FROM dbo.fn_ReadinessKPI(
    (SELECT TOP 1 BuildNumber FROM dbo.Dim_Build WHERE IsReleaseCandidate = 1 ORDER BY BuildNumber DESC),
    '2026-09-30');
GO

/*
--------------------------------------------------------------------------------
fn_VerificationQueue(@AsOfBuild, @RigHoursPerWeek) -- the operational control.

One row per requirement that is not current, ranked, with the action named and
the rig cost attached. Bounded by rig capacity, because that is the constraint
that decides what can actually happen next week.
--------------------------------------------------------------------------------
*/
CREATE FUNCTION dbo.fn_VerificationQueue (@AsOfBuild INT, @RigHoursPerWeek DECIMAL(9,2))
RETURNS TABLE
AS RETURN
(
    WITH V AS (
        SELECT v.*,
               IsCurrent = CASE WHEN v.MeetsPolicy = 1 AND v.StaleByCode = 0
                                 AND v.StaleByRequirement = 0 THEN 1 ELSE 0 END
        FROM dbo.fn_RequirementVerification(@AsOfBuild) v
    ),
    Cost AS (
        SELECT v.RequirementKey,
               RigHours = CAST(ISNULL(SUM(tlr.RigHours), 0) AS DECIMAL(9,2)),
               TestMinutes = ISNULL(SUM(tc.ExpectedDurationMin), 0),
               Cases = COUNT(*)
        FROM V v
        JOIN dbo.Dim_TestCase tc      ON tc.RequirementKey = v.RequirementKey
        JOIN dbo.Ref_TestLevelRank tlr ON tlr.TestLevel     = tc.TestLevel
        GROUP BY v.RequirementKey
    ),
    Scored AS (
        SELECT
            v.RequirementKey, v.RequirementID, v.Title, v.ReqType, v.Priority,
            v.SubsystemCode, v.SubsystemName, v.Criticality, v.RequiresHILRig,
            v.HasEvidence, v.MeetsPolicy, v.StaleByCode, v.StaleByRequirement,
            v.PassingButUnderLevelled, v.PassingButSelfVerified,
            v.LastGoodBuild, v.InterveningBuilds, v.MinTestLevel,
            p.PersonID AS OwnerID, p.PersonName AS OwnerName, p.Team AS OwnerTeam,
            c.RigHours, c.TestMinutes, c.Cases,

            -- Named action. The state decides the work, and the four states
            -- need genuinely different work -- which is the argument for
            -- separating them rather than reporting one "not done" bucket.
            ActionCode = CASE
                WHEN v.HasEvidence = 0                             THEN 'VERIFY'
                WHEN v.MeetsPolicy = 0 AND v.PassingButSelfVerified > 0
                                        AND v.PassingButUnderLevelled = 0 THEN 'INDEPENDENT_WITNESS'
                WHEN v.MeetsPolicy = 0                             THEN 'RAISE_TEST_LEVEL'
                WHEN v.StaleByRequirement = 1                      THEN 'REVIEW_THEN_RERUN'
                ELSE                                                    'RERUN' END,

            -- Priority score, 0-100, four weighted terms so a reviewer can see
            -- which one is driving a given rank rather than trusting a single
            -- opaque number.
            PriorityScore = CAST(
                  CASE v.Priority WHEN 'MustShip' THEN 45 WHEN 'ShouldShip' THEN 18 ELSE 4 END
                + CASE v.Criticality WHEN 'Safety' THEN 25 WHEN 'Mission' THEN 14 ELSE 5 END
                + CASE
                    WHEN v.HasEvidence = 0          THEN 20   -- nothing at all
                    WHEN v.MeetsPolicy = 0          THEN 16   -- evidence that does not count
                    WHEN v.StaleByRequirement = 1   THEN 12   -- the requirement itself moved
                    ELSE 8 END                                -- code moved underneath it
                + CASE
                    WHEN v.InterveningBuilds >= 12 THEN 10
                    WHEN v.InterveningBuilds >= 6  THEN 6
                    WHEN v.InterveningBuilds >= 1  THEN 3
                    ELSE 0 END
                AS DECIMAL(6,2))
        FROM V v
        JOIN dbo.Dim_Person p ON p.PersonKey = v.OwnerKey
        LEFT JOIN Cost c      ON c.RequirementKey = v.RequirementKey
        WHERE v.IsCurrent = 0
    ),
    Ranked AS (
        SELECT s.*,
               PriorityRank = ROW_NUMBER() OVER (
                    ORDER BY s.PriorityScore DESC, s.RigHours ASC, s.RequirementID),
               -- Cumulative rig hours in priority order. The bound is applied
               -- to this, not to a row count: ten requirements needing eight
               -- rig hours each is not the same week of work as ten needing
               -- none, and a queue that cannot tell them apart will be
               -- abandoned by the people it is for.
               CumulativeRigHours = SUM(s.RigHours) OVER (
                    ORDER BY s.PriorityScore DESC, s.RigHours ASC, s.RequirementID
                    ROWS BETWEEN UNBOUNDED PRECEDING AND CURRENT ROW)
        FROM Scored s
    )
    SELECT
        r.PriorityRank, r.RequirementID, r.Title, r.ReqType, r.Priority,
        r.SubsystemCode, r.SubsystemName, r.Criticality,
        r.OwnerID, r.OwnerName, r.OwnerTeam,
        r.MinTestLevel, r.LastGoodBuild, r.InterveningBuilds,
        r.Cases, r.RigHours, r.TestMinutes, r.CumulativeRigHours,
        r.PriorityScore, r.ActionCode,
        IsThisWeek = CAST(CASE WHEN r.CumulativeRigHours <= @RigHoursPerWeek THEN 1 ELSE 0 END AS BIT),
        RecommendedAction = CASE r.ActionCode
            WHEN 'VERIFY'  THEN 'No passing evidence exists. Write and run the test before the gate review; this is not a re-run, it is first-time verification.'
            WHEN 'RERUN'   THEN CONCAT('Evidence was valid at build ', r.LastGoodBuild, ' but the subsystem has changed in ',
                                       r.InterveningBuilds, ' builds since. Re-run on the release candidate.')
            WHEN 'REVIEW_THEN_RERUN' THEN 'The requirement itself was modified after it was verified. Confirm the test still checks the right thing, then re-run -- re-running first risks certifying against the old wording.'
            WHEN 'RAISE_TEST_LEVEL'  THEN CONCAT('Passing evidence sits below the ', r.MinTestLevel,
                                       ' level this requirement type demands. Book rig time; a lower-level pass is not admissible.')
            ELSE 'Verified by its own owner against a policy requiring independence. Re-witness with a second engineer; the test need not be re-executed if nothing has changed.'
        END,
        AsOfBuild = @AsOfBuild
    FROM Ranked r
);
GO

/*
vw_VerificationQueue -- one rig-week.

180 hours is six rigs at 30 bookable hours. It is a parameter rather than a
constant in the view precisely so a programme manager can ask "what if we add
a rig" without editing SQL.
*/
CREATE VIEW dbo.vw_VerificationQueue AS
SELECT * FROM dbo.fn_VerificationQueue(
    (SELECT TOP 1 BuildNumber FROM dbo.Dim_Build WHERE IsReleaseCandidate = 1 ORDER BY BuildNumber DESC),
    180.00);
GO

-- =============================================================================
-- Creation gate
-- =============================================================================
DECLARE @missing VARCHAR(400) = '';
IF OBJECT_ID('dbo.fn_RAIDExposure','IF')       IS NULL SET @missing += 'fn_RAIDExposure ';
IF OBJECT_ID('dbo.fn_SubsystemReadiness','IF') IS NULL SET @missing += 'fn_SubsystemReadiness ';
IF OBJECT_ID('dbo.fn_ReadinessKPI','IF')       IS NULL SET @missing += 'fn_ReadinessKPI ';
IF OBJECT_ID('dbo.fn_VerificationQueue','IF')  IS NULL SET @missing += 'fn_VerificationQueue ';
IF OBJECT_ID('dbo.vw_ReadinessKPI','V')        IS NULL SET @missing += 'vw_ReadinessKPI ';
IF OBJECT_ID('dbo.vw_VerificationQueue','V')   IS NULL SET @missing += 'vw_VerificationQueue ';
IF OBJECT_ID('dbo.vw_SubsystemReadiness','V')  IS NULL SET @missing += 'vw_SubsystemReadiness ';
IF OBJECT_ID('dbo.vw_RAIDExposure','V')        IS NULL SET @missing += 'vw_RAIDExposure ';
IF @missing <> '' THROW 52040, 'FAILED to create KPI objects -- scroll up for the compile error.', 1;

-- The queue must hold exactly the requirements that are not current: one row
-- each, none missing, none invented.
DECLARE @notCurrent INT = (SELECT COUNT(*) FROM dbo.vw_RequirementVerification WHERE IsCurrent = 0);
DECLARE @queued     INT = (SELECT COUNT(*) FROM dbo.vw_VerificationQueue);
IF @notCurrent <> @queued
    THROW 52041, 'The verification queue does not contain exactly the non-current requirements.', 1;

PRINT 'KPI objects created and verified: readiness scorecard, subsystem cut, RAID exposure, verification queue.';
GO
