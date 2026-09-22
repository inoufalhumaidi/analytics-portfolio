/*
================================================================================
Project 4 — Talon Robotics: Payload Deployment Delivery
Script:  04_data_quality_checks.sql
Purpose: Behavioural data-quality checks, and a gate that can stop the build.

WHAT THE GATE MEASURES, AND WHY IT IS NOT A DEFECT COUNT

    Counting defects tells you how untidy the data is. It does not tell you
    whether to trust the answer. The question this programme actually needs
    answering is narrower and harder:

        of the MustShip requirements, how many have a readiness answer that
        rests on data we know to be wrong?

    So the gate is a share of the MustShip population, not a share of rows, and
    only checks that could CHANGE a readiness answer count toward it. A
    duplicate pass and a genuine pass produce the same verification state, so
    the duplicate is worth reporting and worth fixing -- but it cannot move the
    ship decision, and inflating the gate with it would make the gate easier to
    dismiss.

CAUSE AND EFFECT ARE COUNTED ONCE

    Every check declares its ImpactClass and whether it CountsTowardExposure,
    in the catalog, ahead of the data. A duplicated run is also, mechanically,
    a run that will look like a repeat execution; if both checks counted, the
    same defect would be paid for twice. Declaring it in the catalog rather
    than inferring it from the rows found is the difference between a number
    that is auditable and a number that happens to be right today.

POLICY BREACHES ARE NOT DATA DEFECTS

    A Safety requirement verified by its own owner is not bad data -- the row
    is accurate, and it is accurately recording something that should not have
    happened. It appears here because this is where a reviewer looks, but its
    ImpactClass is 'Policy' and it is excluded from the data-quality gate. The
    KPI layer holds it to account instead. Conflating "the record is wrong"
    with "the practice is wrong" makes both harder to fix.

DATA DISCLOSURE: Talon Robotics is fictional; all data is synthetic. No
confidential data and no production-system claim is involved.
================================================================================
*/

USE TalonDelivery;
GO

IF OBJECT_ID('dbo.usp_RunDataQualityChecks','P') IS NOT NULL DROP PROCEDURE dbo.usp_RunDataQualityChecks;
IF OBJECT_ID('dbo.vw_DQ_Summary','V')            IS NOT NULL DROP VIEW dbo.vw_DQ_Summary;
IF OBJECT_ID('dbo.vw_DQ_Anomalies','V')          IS NOT NULL DROP VIEW dbo.vw_DQ_Anomalies;
IF OBJECT_ID('dbo.fn_DQ_Anomalies','IF')         IS NOT NULL DROP FUNCTION dbo.fn_DQ_Anomalies;
IF OBJECT_ID('dbo.vw_DQ_CheckCatalog','V')       IS NOT NULL DROP VIEW dbo.vw_DQ_CheckCatalog;
IF OBJECT_ID('dbo.fn_DuplicateTestRuns','IF')    IS NOT NULL DROP FUNCTION dbo.fn_DuplicateTestRuns;
GO

/*
--------------------------------------------------------------------------------
fn_DuplicateTestRuns -- ONE definition of "duplicate", reused everywhere.

Matched on the business signature -- same test case, same build, same day, same
result -- never on anything the generator planted. There is no marker to read:
in real data a duplicated run looks exactly like this and nothing else.

It is a function rather than repeated inline SQL because the vendor-scorecard
equivalent in Project 3 was written twice, drifted, and one copy ended up
reading a synthetic prefix that would have matched nothing in production.
--------------------------------------------------------------------------------
*/
CREATE FUNCTION dbo.fn_DuplicateTestRuns ()
RETURNS TABLE
AS RETURN
(
    WITH Ranked AS (
        SELECT tr.TestRunKey, tr.TestCaseKey, tr.BuildKey, tr.RunDateKey, tr.Result,
               CopyNo = ROW_NUMBER() OVER (
                          PARTITION BY tr.TestCaseKey, tr.BuildKey, tr.RunDateKey, tr.Result
                          ORDER BY tr.TestRunKey)
        FROM dbo.Fact_TestRun tr
    )
    SELECT TestRunKey, TestCaseKey, BuildKey, RunDateKey, Result, CopyNo
    FROM Ranked WHERE CopyNo > 1
);
GO

/*
--------------------------------------------------------------------------------
vw_DQ_CheckCatalog -- every check declared before any row is counted.
--------------------------------------------------------------------------------
*/
CREATE VIEW dbo.vw_DQ_CheckCatalog AS
SELECT AnomalyType, EntityType, Severity, ImpactClass, CountsTowardExposure, WhatItMeans
FROM (VALUES
 ('DUPLICATE_TEST_RUN',    'TestRun',     'Medium', 'Evidence',  CAST(0 AS BIT),
  'The same test case recorded twice against the same build on the same day with the same result. Overstates how much testing was done; does not change any requirement''s verification state, because a duplicate pass and a pass are the same evidence.'),
 ('RUN_BEFORE_BUILD',      'TestRun',     'High',   'Timing',    CAST(1 AS BIT),
  'A test run dated before the build it ran against existed. Either the date or the build is wrong, and which one decides whether the evidence is current -- so it can change a readiness answer.'),
 ('RUN_ON_WITHDRAWN_REQ',  'TestRun',     'Medium', 'Waste',     CAST(0 AS BIT),
  'Test effort spent verifying a requirement the programme has already withdrawn. Costs rig time that the outstanding queue needs; does not affect readiness, because withdrawn requirements are out of scope by definition.'),
 ('REQ_WITHOUT_TESTCASE',  'Requirement', 'High',   'Coverage',  CAST(1 AS BIT),
  'A baselined requirement with no test case at all. It cannot be verified, so its readiness answer is not merely unknown -- it is unobtainable until somebody writes a test.'),
 ('TESTCASE_NEVER_RUN',    'TestCase',    'Medium', 'Coverage',  CAST(0 AS BIT),
  'A test case that exists but has never been executed. Counted separately from REQ_WITHOUT_TESTCASE so that "nobody wrote a test" and "nobody ran the test" are not reported as the same failure.'),
 ('WORKITEM_CLOSED_BEFORE_OPEN','WorkItem','High',  'Timing',    CAST(0 AS BIT),
  'A work item closed before it was opened. Corrupts cycle-time and burn-down figures; does not touch verification evidence.'),
 ('RAID_DUE_BEFORE_RAISED','RAID',        'Medium', 'Timing',    CAST(0 AS BIT),
  'A RAID item due before it was raised. Makes the item permanently overdue on the day it is created, which quietly inflates the overdue rate.'),
 ('ORPHAN_DIMENSION_KEY',  'TestRun',     'High',   'Integrity', CAST(1 AS BIT),
  'A fact row pointing at a requirement, build, person or date that does not exist. Enforced structurally by foreign keys; checked anyway, because a check that never fires is how you find out the constraint was dropped.'),
 ('SELF_VERIFIED_SAFETY',  'Requirement', 'High',   'Policy',    CAST(0 AS BIT),
  'A Safety or Regulatory requirement whose only passing evidence was produced by its own owner, against a policy that requires an independent tester. The record is accurate; the practice is not. Held to account by the KPI layer, not by the data-quality gate.'),
 ('UNDER_LEVELLED_VERIFICATION','Requirement','High','Policy',   CAST(0 AS BIT),
  'A requirement whose only passing evidence sits below the test level its type demands -- a safety behaviour signed off by a unit test. Again accurate data recording an inadequate practice.'),
 ('BLOCKED_AT_RELEASE_CANDIDATE','TestCase','Medium','Evidence', CAST(0 AS BIT),
  'A test case whose most recent run against the release candidate was Blocked. A blocked test is not evidence of anything, and a readiness metric that treats "could not run" as "did not fail" is how rig contention stays invisible until the ship date.')
) AS c(AnomalyType, EntityType, Severity, ImpactClass, CountsTowardExposure, WhatItMeans);
GO

/*
--------------------------------------------------------------------------------
fn_DQ_Anomalies(@AsOfBuild) -- one row per defect found.

Parameterised so the same checks can be run against any build, which is what
makes it possible to say "this was already true at the last gate" rather than
"this appeared some time in the last six months".
--------------------------------------------------------------------------------
*/
CREATE FUNCTION dbo.fn_DQ_Anomalies (@AsOfBuild INT)
RETURNS TABLE
AS RETURN
(
    -- D1
    SELECT AnomalyType = 'DUPLICATE_TEST_RUN',
           EntityType  = 'TestRun',
           EntityRef   = tc.TestCaseID,
           Detail      = CONCAT('Copy ', d.CopyNo, ' of the same result for ', tc.TestCaseID,
                                ' on build ', b.BuildID, '.')
    FROM dbo.fn_DuplicateTestRuns() d
    JOIN dbo.Dim_TestCase tc ON tc.TestCaseKey = d.TestCaseKey
    JOIN dbo.Dim_Build b     ON b.BuildKey     = d.BuildKey
    WHERE b.BuildNumber <= @AsOfBuild

    UNION ALL
    -- D2
    SELECT 'RUN_BEFORE_BUILD', 'TestRun', tc.TestCaseID,
           CONCAT('Run dated ', CONVERT(CHAR(10), rd.[Date], 120), ' against build ', b.BuildID,
                  ' which was not cut until ', CONVERT(CHAR(10), bd.[Date], 120), '.')
    FROM dbo.Fact_TestRun tr
    JOIN dbo.Dim_TestCase tc ON tc.TestCaseKey = tr.TestCaseKey
    JOIN dbo.Dim_Build b     ON b.BuildKey     = tr.BuildKey
    JOIN dbo.Dim_Date rd     ON rd.DateKey     = tr.RunDateKey
    JOIN dbo.Dim_Date bd     ON bd.DateKey     = b.BuildDateKey
    WHERE rd.[Date] < bd.[Date] AND b.BuildNumber <= @AsOfBuild

    UNION ALL
    -- D5
    SELECT 'RUN_ON_WITHDRAWN_REQ', 'TestRun', r.RequirementID,
           CONCAT('Test effort recorded against ', r.RequirementID, ', withdrawn from the baseline.')
    FROM dbo.Fact_TestRun tr
    JOIN dbo.Dim_TestCase tc  ON tc.TestCaseKey  = tr.TestCaseKey
    JOIN dbo.Dim_Requirement r ON r.RequirementKey = tc.RequirementKey
    JOIN dbo.Dim_Build b      ON b.BuildKey      = tr.BuildKey
    WHERE r.ReqStatus = 'Withdrawn' AND b.BuildNumber <= @AsOfBuild

    UNION ALL
    SELECT 'REQ_WITHOUT_TESTCASE', 'Requirement', r.RequirementID,
           CONCAT(r.Priority, ' ', r.ReqType, ' requirement with no test case defined.')
    FROM dbo.Dim_Requirement r
    WHERE r.ReqStatus = 'Baselined'
      AND NOT EXISTS (SELECT 1 FROM dbo.Dim_TestCase tc WHERE tc.RequirementKey = r.RequirementKey)

    UNION ALL
    SELECT 'TESTCASE_NEVER_RUN', 'TestCase', tc.TestCaseID,
           CONCAT('Test case defined for ', r.RequirementID, ' but never executed.')
    FROM dbo.Dim_TestCase tc
    JOIN dbo.Dim_Requirement r ON r.RequirementKey = tc.RequirementKey
    WHERE r.ReqStatus = 'Baselined'
      AND NOT EXISTS (SELECT 1 FROM dbo.Fact_TestRun tr
                      JOIN dbo.Dim_Build b ON b.BuildKey = tr.BuildKey
                      WHERE tr.TestCaseKey = tc.TestCaseKey AND b.BuildNumber <= @AsOfBuild)

    UNION ALL
    -- D3
    SELECT 'WORKITEM_CLOSED_BEFORE_OPEN', 'WorkItem', w.WorkItemID,
           CONCAT('Closed ', CONVERT(CHAR(10), cd.[Date], 120),
                  ', opened ', CONVERT(CHAR(10), od.[Date], 120), '.')
    FROM dbo.Fact_WorkItem w
    JOIN dbo.Dim_Date od ON od.DateKey = w.OpenedDateKey
    JOIN dbo.Dim_Date cd ON cd.DateKey = w.ClosedDateKey
    WHERE cd.[Date] < od.[Date]

    UNION ALL
    -- D4
    SELECT 'RAID_DUE_BEFORE_RAISED', 'RAID', x.RAIDID,
           CONCAT('Due ', CONVERT(CHAR(10), dd.[Date], 120),
                  ', raised ', CONVERT(CHAR(10), rd.[Date], 120), '.')
    FROM dbo.Fact_RAID x
    JOIN dbo.Dim_Date rd ON rd.DateKey = x.RaisedDateKey
    JOIN dbo.Dim_Date dd ON dd.DateKey = x.DueDateKey
    WHERE dd.[Date] < rd.[Date]

    UNION ALL
    -- EntityRef must be the TestCaseID, not the TestRunKey. The gate in
    -- usp_RunDataQualityChecks attributes 'TestRun' anomalies by matching
    -- EntityRef against Dim_TestCase.TestCaseID, so a bare integer key could
    -- never match and this check -- declared CountsTowardExposure = 1 -- was
    -- structurally unable to move the gate. Harmless today at 0 orphans, which
    -- is exactly why it survived: the catalog says the check counts, and it
    -- could not.
    SELECT 'ORPHAN_DIMENSION_KEY', 'TestRun',
           -- The TestCaseID where one exists, so the gate can attribute the
           -- anomaly to a requirement; the raw key where the orphan IS the
           -- missing test case, because then there is genuinely nothing to
           -- attribute it to and saying so is better than inventing a match.
           ISNULL(tc.TestCaseID, CONCAT('TestRunKey ', tr.TestRunKey)),
           'Fact row points at a dimension member that does not exist.'
    FROM dbo.Fact_TestRun tr
    LEFT JOIN dbo.Dim_TestCase tc ON tc.TestCaseKey = tr.TestCaseKey
    WHERE NOT EXISTS (SELECT 1 FROM dbo.Dim_TestCase t WHERE t.TestCaseKey = tr.TestCaseKey)
       OR NOT EXISTS (SELECT 1 FROM dbo.Dim_Build b    WHERE b.BuildKey    = tr.BuildKey)
       OR NOT EXISTS (SELECT 1 FROM dbo.Dim_Person p   WHERE p.PersonKey   = tr.RunByKey)

    UNION ALL
    SELECT 'SELF_VERIFIED_SAFETY', 'Requirement', v.RequirementID,
           CONCAT(v.ReqType, ' requirement on ', v.SubsystemCode,
                  ' whose passing evidence came from its own owner; policy requires an independent tester.')
    FROM dbo.fn_RequirementVerification(@AsOfBuild) v
    WHERE v.RequiresIndependentTester = 1
      AND v.PassingButSelfVerified > 0
      AND v.MeetsPolicy = 0

    UNION ALL
    SELECT 'UNDER_LEVELLED_VERIFICATION', 'Requirement', v.RequirementID,
           CONCAT(v.ReqType, ' requirement requiring ', v.MinTestLevel,
                  ' evidence, passing only at a lower test level.')
    FROM dbo.fn_RequirementVerification(@AsOfBuild) v
    WHERE v.PassingButUnderLevelled > 0 AND v.MeetsPolicy = 0

    UNION ALL
    SELECT 'BLOCKED_AT_RELEASE_CANDIDATE', 'TestCase', tc.TestCaseID,
           CONCAT('Most recent run of ', tc.TestCaseID, ' was Blocked, not Pass or Fail.')
    FROM dbo.fn_LatestRunPerCase(@AsOfBuild) lr
    JOIN dbo.Dim_TestCase tc   ON tc.TestCaseKey   = lr.TestCaseKey
    JOIN dbo.Dim_Requirement r ON r.RequirementKey = tc.RequirementKey
    WHERE lr.Result = 'Blocked' AND r.ReqStatus = 'Baselined'
);
GO

CREATE VIEW dbo.vw_DQ_Anomalies AS
SELECT a.*, c.EntityType AS CatalogEntityType, c.Severity, c.ImpactClass, c.CountsTowardExposure, c.WhatItMeans
FROM dbo.fn_DQ_Anomalies((SELECT TOP 1 BuildNumber FROM dbo.Dim_Build WHERE IsReleaseCandidate = 1 ORDER BY BuildNumber DESC)) a
JOIN dbo.vw_DQ_CheckCatalog c ON c.AnomalyType = a.AnomalyType;
GO

/*
--------------------------------------------------------------------------------
vw_DQ_Summary -- one row per check, INCLUDING the ones that found nothing.

Checks with zero findings are the point of the table. A summary that lists only
what was found cannot distinguish "we looked and it was clean" from "we never
looked", and those are very different statements to put in front of a review
board.
--------------------------------------------------------------------------------
*/
CREATE VIEW dbo.vw_DQ_Summary AS
WITH Pop AS (
    SELECT TestRuns     = (SELECT COUNT(*) FROM dbo.Fact_TestRun),
           TestCases    = (SELECT COUNT(*) FROM dbo.Dim_TestCase),
           Requirements = (SELECT COUNT(*) FROM dbo.Dim_Requirement WHERE ReqStatus = 'Baselined'),
           WorkItems    = (SELECT COUNT(*) FROM dbo.Fact_WorkItem),
           RAIDItems    = (SELECT COUNT(*) FROM dbo.Fact_RAID)
)
SELECT
    c.AnomalyType, c.EntityType, c.Severity, c.ImpactClass, c.CountsTowardExposure,
    Anomalies = ISNULL(f.n, 0),
    PopulationScanned = CASE c.EntityType
        WHEN 'TestRun'     THEN p.TestRuns
        WHEN 'TestCase'    THEN p.TestCases
        WHEN 'Requirement' THEN p.Requirements
        WHEN 'WorkItem'    THEN p.WorkItems
        ELSE p.RAIDItems END,
    AnomalyRatePct = CAST(100.0 * ISNULL(f.n, 0) / NULLIF(CASE c.EntityType
        WHEN 'TestRun'     THEN p.TestRuns
        WHEN 'TestCase'    THEN p.TestCases
        WHEN 'Requirement' THEN p.Requirements
        WHEN 'WorkItem'    THEN p.WorkItems
        ELSE p.RAIDItems END, 0) AS DECIMAL(9,3)),
    c.WhatItMeans
FROM dbo.vw_DQ_CheckCatalog c
CROSS JOIN Pop p
LEFT JOIN (SELECT AnomalyType, n = COUNT(*) FROM dbo.vw_DQ_Anomalies GROUP BY AnomalyType) f
       ON f.AnomalyType = c.AnomalyType;
GO

/*
--------------------------------------------------------------------------------
usp_RunDataQualityChecks -- the report and the gate.
--------------------------------------------------------------------------------
*/
CREATE PROCEDURE dbo.usp_RunDataQualityChecks
    @MaxAffectedMustShipPct DECIMAL(6,3) = 2.000
AS
BEGIN
    SET NOCOUNT ON;

    PRINT '=== Talon Robotics: Programme Data Quality Report (at the release candidate) ===';

    SELECT AnomalyType, Severity, ImpactClass, Anomalies, PopulationScanned,
           AnomalyRatePct, CountsTowardExposure
    FROM dbo.vw_DQ_Summary
    ORDER BY CASE Severity WHEN 'High' THEN 1 WHEN 'Medium' THEN 2 ELSE 3 END,
             Anomalies DESC, AnomalyType;

    DECLARE @defects INT = (SELECT SUM(Anomalies) FROM dbo.vw_DQ_Summary);

    -- The gate: MustShip requirements whose readiness answer rests on data a
    -- check has flagged. A count of rows would be easier to compute and would
    -- not be an answer to anything a programme board asks.
    DECLARE @mustShip INT = (SELECT COUNT(*) FROM dbo.vw_RequirementVerification WHERE Priority = 'MustShip');
    DECLARE @affected INT = (
        SELECT COUNT(DISTINCT v.RequirementKey)
        FROM dbo.vw_RequirementVerification v
        WHERE v.Priority = 'MustShip'
          AND EXISTS (
              SELECT 1
              FROM dbo.vw_DQ_Anomalies a
              WHERE a.CountsTowardExposure = 1
                AND (   (a.EntityType = 'Requirement' AND a.EntityRef = v.RequirementID)
                     OR (a.EntityType = 'TestRun'     AND a.EntityRef IN
                            (SELECT tc.TestCaseID FROM dbo.Dim_TestCase tc
                             WHERE tc.RequirementKey = v.RequirementKey)))));

    DECLARE @pct DECIMAL(9,3) = CAST(100.0 * @affected / NULLIF(@mustShip, 0) AS DECIMAL(9,3));

    PRINT '';
    PRINT CONCAT('Defects found: ', @defects,
                 ' | MustShip requirements whose readiness rests on flagged data: ',
                 @affected, ' of ', @mustShip, ' (', @pct, '%)');

    IF @pct > @MaxAffectedMustShipPct
    BEGIN
        PRINT CONCAT('QA GATE: FAIL -- ', @pct, '% of MustShip readiness rests on flagged data, above the ',
                     @MaxAffectedMustShipPct, '% tolerance. See dbo.vw_DQ_Anomalies.');
        -- A gate that only PRINTs is not a gate: it returns exit code 0, so no
        -- build step can act on it and the word FAIL is doing work the code is
        -- not. Every project in this portfolio shipped that defect at least
        -- once before it was caught by actually checking an exit code.
        DECLARE @msg VARCHAR(400) = CONCAT('QA gate failed: ', @pct,
            '% of MustShip readiness rests on flagged data, above the ',
            @MaxAffectedMustShipPct, '% tolerance. See dbo.vw_DQ_Anomalies.');
        ;THROW 52030, @msg, 1;
    END
    ELSE
        PRINT 'QA GATE: PASS -- readiness for MustShip requirements does not rest on flagged data beyond tolerance.';
END;
GO

-- =============================================================================
-- Creation gate
-- =============================================================================
DECLARE @missing VARCHAR(400) = '';
IF OBJECT_ID('dbo.fn_DuplicateTestRuns','IF')      IS NULL SET @missing += 'fn_DuplicateTestRuns ';
IF OBJECT_ID('dbo.vw_DQ_CheckCatalog','V')         IS NULL SET @missing += 'vw_DQ_CheckCatalog ';
IF OBJECT_ID('dbo.fn_DQ_Anomalies','IF')           IS NULL SET @missing += 'fn_DQ_Anomalies ';
IF OBJECT_ID('dbo.vw_DQ_Anomalies','V')            IS NULL SET @missing += 'vw_DQ_Anomalies ';
IF OBJECT_ID('dbo.vw_DQ_Summary','V')              IS NULL SET @missing += 'vw_DQ_Summary ';
IF OBJECT_ID('dbo.usp_RunDataQualityChecks','P')   IS NULL SET @missing += 'usp_RunDataQualityChecks ';
IF @missing <> '' THROW 52031, 'FAILED to create data-quality objects -- scroll up for the compile error.', 1;

-- Every catalogued check must appear in the summary, including the clean ones.
IF (SELECT COUNT(*) FROM dbo.vw_DQ_Summary) <> (SELECT COUNT(*) FROM dbo.vw_DQ_CheckCatalog)
    THROW 52032, 'vw_DQ_Summary does not report every catalogued check.', 1;

PRINT 'Data-quality objects created and verified: 11 behavioural checks behind a gate that raises.';
GO
