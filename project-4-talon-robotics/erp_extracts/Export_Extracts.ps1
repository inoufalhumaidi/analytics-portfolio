<#
================================================================================
Project 4 -- Talon Robotics: Payload Deployment Delivery
Script:  Export_Extracts.ps1
Purpose: Write the CSV extracts that stand in for a nightly export from the
         requirements tool, the test manager and the RAID register.

WHY CSV AND NOT A LIVE CONNECTION
    The Excel workbook is a deliverable in its own right, and a reviewer should
    be able to open it and see working formulas without installing SQL Server,
    restoring a database or holding credentials. These extracts are what the
    workbook reads through Power Query. Power BI connects to SQL Server
    directly, because a semantic model belongs next to the warehouse.

WHY POWERSHELL AND NOT bcp OR sqlcmd -s","
    Requirement titles and recommended-action text contain commas and
    apostrophes. sqlcmd's separator option does not quote fields, so those rows
    split into extra columns and every downstream count is quietly wrong --
    quietly, because the file still parses. Export-Csv quotes correctly.

DATA DISCLOSURE: Talon Robotics is fictional; all data is synthetic. No
confidential data and no production system is involved.
================================================================================
#>

param(
    [string]$ServerInstance = 'localhost\TEW_SQLEXPRESS',
    [string]$Database       = 'TalonDelivery',
    [string]$OutputPath     = $PSScriptRoot,
    # DECLARED, PRINTED IN THE BANNER, AND USED BY NOTHING.
    #
    # Every extract below reads a view that is pinned to a literal date inside
    # SQL -- vw_RAIDExposure is fn_RAIDExposure('2026-09-30'), vw_ReadinessKPI
    # is fn_ReadinessKPI(RC, '2026-09-30') -- so passing -AsOfDate 2025-06-30
    # printed a banner claiming a June-2025 extract and wrote byte-identical
    # September-2026 files, which then fed the Excel control and every figure
    # downstream.
    #
    # This is the third instance in this project of the class it documents
    # itself: a point-in-time parameter that is accepted and ignored. It is
    # kept here ONLY as the banner label, and the banner now says so, rather
    # than being removed and losing the record of the trap.
    [string]$AsOfDate       = '2026-09-30'
)

$ErrorActionPreference = 'Stop'

function Invoke-Extract {
    param([string]$Sql, [string]$FileName, [string]$Description)

    $conn = New-Object System.Data.SqlClient.SqlConnection
    $conn.ConnectionString = "Server=$ServerInstance;Database=$Database;Integrated Security=True;Encrypt=False;Connect Timeout=30"
    $conn.Open()
    try {
        $cmd = $conn.CreateCommand()
        $cmd.CommandText = $Sql
        $cmd.CommandTimeout = 600
        $adapter = New-Object System.Data.SqlClient.SqlDataAdapter $cmd
        $table = New-Object System.Data.DataTable
        [void]$adapter.Fill($table)
    }
    finally { $conn.Close() }

    $target = Join-Path $OutputPath $FileName
    $table | Export-Csv -Path $target -NoTypeInformation -Encoding UTF8
    "{0,-32} {1,6} rows  {2}" -f $FileName, $table.Rows.Count, $Description | Write-Host
    return $table.Rows.Count
}

Write-Host ""
Write-Host "Talon Robotics -- programme extract build (views are pinned to $AsOfDate in SQL; this switch is a label, not a filter)" -ForegroundColor Cyan
Write-Host ("-" * 88)
$total = 0

# The verification model itself -- one row per baselined requirement. This is
# the extract everything else is a cut of, and the one a reviewer should read
# first.
$total += Invoke-Extract -FileName 'requirement_verification.csv' -Description 'dbo.vw_RequirementVerification' -Sql @"
SELECT RequirementID, Title, ReqType, Priority, SubsystemCode, SubsystemName, Criticality,
       MinTestLevel, CAST(RequiresIndependentTester AS INT) AS RequiresIndependentTester,
       TotalCases, CasesWithRuns, PassingCases, BlockedCases,
       LastPassBuild, LastGoodBuild,
       PassingButUnderLevelled, PassingButSelfVerified,
       CAST(HasEvidence AS INT) AS HasEvidence,
       CAST(MeetsPolicy AS INT) AS MeetsPolicy,
       CAST(StaleByCode AS INT) AS StaleByCode,
       CAST(StaleByRequirement AS INT) AS StaleByRequirement,
       CAST(IsCurrent AS INT) AS IsCurrent,
       VerificationState, InterveningBuilds, AsOfBuild
FROM dbo.vw_RequirementVerification
ORDER BY RequirementID;
"@

$total += Invoke-Extract -FileName 'verification_queue.csv' -Description 'dbo.vw_VerificationQueue' -Sql @"
SELECT PriorityRank, RequirementID, Title, ReqType, Priority,
       SubsystemCode, SubsystemName, Criticality,
       OwnerID, OwnerName, OwnerTeam,
       MinTestLevel, LastGoodBuild, InterveningBuilds,
       Cases, RigHours, TestMinutes, CumulativeRigHours,
       PriorityScore, ActionCode, CAST(IsThisWeek AS INT) AS IsThisWeek, RecommendedAction
FROM dbo.vw_VerificationQueue
ORDER BY PriorityRank;
"@

$total += Invoke-Extract -FileName 'subsystem_readiness.csv' -Description 'dbo.vw_SubsystemReadiness' -Sql @"
SELECT SubsystemCode, SubsystemName, Criticality,
       CAST(RequiresHILRig AS INT) AS RequiresHILRig,
       Requirements, MustShip, CurrentAll, CurrentMustShip,
       NoEvidence, Insufficient, StaleByCode, StaleByReq,
       ReadinessPct, BuildsChanged, LinesChanged, LastChangeBuild, RigHoursOutstanding
FROM dbo.vw_SubsystemReadiness
ORDER BY ReadinessPct;
"@

$total += Invoke-Extract -FileName 'readiness_kpi.csv' -Description 'dbo.vw_ReadinessKPI (one row)' -Sql @"
SELECT AsOfBuild, CONVERT(CHAR(10), AsOfDate, 23) AS AsOfDate,
       Requirements, MustShipCount,
       WorkItemCompletionPct, VerificationCoveragePct, ShipReadinessPct,
       StaleVerificationPct, PolicyCompliancePct,
       NoEvidenceCount, InsufficientCount, StaleCount, CurrentCount,
       BlockedTestPct, CriticalRAIDOpen, OverdueRAIDPct, TotalRAIDExposure,
       RigHoursOutstanding
FROM dbo.vw_ReadinessKPI;
"@

# The trend is what turns one number into a trajectory. A programme board can
# tell a recovery from a slide; a single figure cannot.
$total += Invoke-Extract -FileName 'readiness_trend.csv' -Description 'usp_ReadinessTrend (build 40 onward)' -Sql @"
DECLARE @t TABLE (BuildNumber INT, BuildID VARCHAR(12), BuildDate DATE,
                  WorkItemCompletionPct DECIMAL(9,2), VerificationCoveragePct DECIMAL(9,2),
                  ShipReadinessPct DECIMAL(9,2), StaleVerificationPct DECIMAL(9,2),
                  ReadinessGapVsReported DECIMAL(9,2));
INSERT INTO @t EXEC dbo.usp_ReadinessTrend @FromBuild = 40, @StepSize = 4;
SELECT * FROM @t ORDER BY BuildNumber;
"@

$total += Invoke-Extract -FileName 'raid_register.csv' -Description 'dbo.vw_RAIDExposure' -Sql @"
SELECT RAIDID, RAIDType, Title, SubsystemCode, SubsystemName, Criticality,
       OwnerID, OwnerName, OwnerTeam,
       CONVERT(CHAR(10), RaisedDate, 23) AS RaisedDate,
       CONVERT(CHAR(10), DueDate, 23)    AS DueDate,
       CONVERT(CHAR(10), ClosedDate, 23) AS ClosedDate,
       Probability, Impact, ExposureScore, ExposureBand,
       RAIDStatusAsOf AS RAIDStatus, CAST(IsOpen AS INT) AS IsOpen, CAST(IsOverdue AS INT) AS IsOverdue,
       DaysOverdue, AgeDays, MitigationNote
FROM dbo.vw_RAIDExposure
ORDER BY ExposureScore DESC, RAIDID;
"@

$total += Invoke-Extract -FileName 'data_quality_summary.csv' -Description 'dbo.vw_DQ_Summary' -Sql @"
SELECT AnomalyType, EntityType, Severity, ImpactClass,
       CAST(CountsTowardExposure AS INT) AS CountsTowardExposure,
       Anomalies, PopulationScanned, AnomalyRatePct, WhatItMeans
FROM dbo.vw_DQ_Summary
ORDER BY CASE Severity WHEN 'High' THEN 1 WHEN 'Medium' THEN 2 ELSE 3 END, Anomalies DESC;
"@

$total += Invoke-Extract -FileName 'build_churn.csv' -Description 'build-by-subsystem change history' -Sql @"
SELECT b.BuildID, b.BuildNumber, CONVERT(CHAR(10), d.[Date], 23) AS BuildDate,
       CAST(b.IsReleaseCandidate AS INT) AS IsReleaseCandidate,
       s.SubsystemCode, s.SubsystemName, c.LinesChanged
FROM dbo.Fact_BuildSubsystemChange c
JOIN dbo.Dim_Build b     ON b.BuildKey = c.BuildKey
JOIN dbo.Dim_Date d      ON d.DateKey  = b.BuildDateKey
JOIN dbo.Dim_Subsystem s ON s.SubsystemKey = c.SubsystemKey
ORDER BY b.BuildNumber, s.SubsystemCode;
"@

# Work items at row grain. Exported so the workbook can compute the 93.86%
# ITSELF rather than reading it from readiness_kpi.csv -- the whole argument of
# this project is a comparison between two numbers, and a comparison in which
# one side is copied from the other tool proves nothing about either.
$total += Invoke-Extract -FileName 'work_items.csv' -Description 'dbo.Fact_WorkItem at row grain' -Sql @"
SELECT w.WorkItemID, r.RequirementID, r.Priority, s.SubsystemCode,
       p.PersonID AS AssignedToID, p.PersonName AS AssignedToName,
       CONVERT(CHAR(10), od.[Date], 23) AS OpenedDate,
       CONVERT(CHAR(10), cd.[Date], 23) AS ClosedDate,
       w.EstimateHours, w.ActualHours, w.WorkItemStatus
FROM dbo.Fact_WorkItem w
JOIN dbo.Dim_Requirement r ON r.RequirementKey = w.RequirementKey
JOIN dbo.Dim_Subsystem s   ON s.SubsystemKey   = r.SubsystemKey
JOIN dbo.Dim_Person p      ON p.PersonKey      = w.AssignedToKey
JOIN dbo.Dim_Date od       ON od.DateKey       = w.OpenedDateKey
LEFT JOIN dbo.Dim_Date cd  ON cd.DateKey       = w.ClosedDateKey
ORDER BY w.WorkItemID;
"@

# The three reference tables. They are the contestable inputs, and shipping
# them alongside the results is what lets a reader disagree with the analysis
# on its own terms rather than dismissing it.
$total += Invoke-Extract -FileName 'verification_policy.csv' -Description 'dbo.Ref_VerificationPolicy' -Sql @"
SELECT ReqType, MinTestLevel, MinTestLevelRank,
       CAST(RequiresIndependentTester AS INT) AS RequiresIndependentTester, PolicyNote
FROM dbo.Ref_VerificationPolicy ORDER BY MinTestLevelRank DESC, ReqType;
"@

$total += Invoke-Extract -FileName 'raid_matrix.csv' -Description 'dbo.Ref_RAIDMatrix' -Sql @"
SELECT Probability, Impact, ExposureScore, ExposureBand FROM dbo.Ref_RAIDMatrix
ORDER BY Probability, Impact;
"@

$total += Invoke-Extract -FileName 'readiness_targets.csv' -Description 'dbo.Ref_ReadinessTargets' -Sql @"
SELECT MetricName, TargetValue, WarningValue, Direction, Unit, [Description]
FROM dbo.Ref_ReadinessTargets ORDER BY MetricName;
"@

Write-Host ("-" * 88)
Write-Host ("{0,-32} {1,6} rows written in total" -f 'TOTAL', $total) -ForegroundColor Green
Write-Host ""
