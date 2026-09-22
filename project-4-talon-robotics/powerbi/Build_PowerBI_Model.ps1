<#
=============================================================================
Project 4 -- Talon Robotics: Payload Deployment Delivery
Script:  Build_PowerBI_Model.ps1
Purpose: Script the Power BI semantic model -- tables, relationships and every
         DAX measure -- into a running Power BI Desktop session via its local
         Analysis Services engine, the mechanism Tabular Editor and DAX Studio
         use.

WHY SCRIPT THE MODEL AT ALL
    Measures clicked into a UI cannot be reviewed, diffed or rebuilt. Scripted,
    they live in source control beside the SQL that feeds them, and a reviewer
    can check that the readiness definition in DAX is the same definition the
    SQL uses. That matters more here than usual: this project's entire claim is
    that two plausible numbers differ, so both had better be computed the same
    way wherever they appear.

    Report PAGES still have to be built by hand -- there is no supported
    scripting path for visuals.

HOW TO RUN
    1. Open Power BI Desktop on a blank report and leave it open.
    2. Run this script.
    3. The FIRST table refresh raises a credential prompt inside Desktop.
       Watch the Desktop window and click through it (Windows / Integrated
       authentication, any privacy level). The script appears to hang until you
       do. This happens once per Desktop session, and again after the tables
       are dropped and recreated.
    4. Save the file as powerbi/Talon_Readiness_Model.pbix.

The Analysis Services client DLLs download from nuget.org on first run into
.amo_libs/ next to this script; .gitignore excludes them.

DATA DISCLOSURE: Talon Robotics is fictional; all data is synthetic. No
confidential data and no production system is involved.
=============================================================================
#>

$ErrorActionPreference = "Stop"

$SqlServerInst = "localhost\TEW_SQLEXPRESS"
$SqlDb         = "TalonDelivery"
# There is deliberately no as-of parameter here. The reporting anchor is read
# from vw_ReadinessKPI into Ref_Reporting, so the model and the SQL scorecard
# take it from one source. A second copy in this script would be a second place
# for it to be wrong.

# -----------------------------------------------------------------------------
# 0. Analysis Services client libraries
# -----------------------------------------------------------------------------
$libDir = Join-Path $PSScriptRoot ".amo_libs"
$dllDir = Join-Path $libDir "lib\net45"
if (-not (Test-Path (Join-Path $dllDir "Microsoft.AnalysisServices.Tabular.dll"))) {
    Write-Host "Downloading Microsoft.AnalysisServices.retail.amd64 from nuget.org..." -ForegroundColor Yellow
    New-Item -ItemType Directory -Force -Path $libDir | Out-Null
    $zipPath = Join-Path $libDir "amo.nupkg.zip"
    Invoke-WebRequest -Uri "https://www.nuget.org/api/v2/package/Microsoft.AnalysisServices.retail.amd64/19.84.1" -OutFile $zipPath
    Expand-Archive -Path $zipPath -DestinationPath $libDir -Force
    Remove-Item $zipPath
}
Add-Type -Path (Join-Path $dllDir "Microsoft.AnalysisServices.Core.dll")
Add-Type -Path (Join-Path $dllDir "Microsoft.AnalysisServices.dll")
Add-Type -Path (Join-Path $dllDir "Microsoft.AnalysisServices.Tabular.dll")

# -----------------------------------------------------------------------------
# 1. Find the running Desktop instance's engine port
# -----------------------------------------------------------------------------
$msmdsrv = Get-Process -Name "msmdsrv" -ErrorAction SilentlyContinue
if (-not $msmdsrv) { throw "No Power BI Desktop session found. Open Power BI Desktop on a blank report first." }
$conn = Get-NetTCPConnection -OwningProcess $msmdsrv.Id -State Listen -ErrorAction SilentlyContinue | Select-Object -First 1
if (-not $conn) { throw "Could not find the local Analysis Services engine's listening port." }
$port = $conn.LocalPort
Write-Host "Found Power BI Desktop's local engine on port $port" -ForegroundColor Cyan

$server = New-Object Microsoft.AnalysisServices.Tabular.Server
$server.Connect("localhost:$port")
$db = $server.Databases[0]
$model = $db.Model
Write-Host "Connected to workspace database $($db.Name)"
Write-Host ("-" * 78)

# -----------------------------------------------------------------------------
# 2. Helpers
# -----------------------------------------------------------------------------
function New-Col($name, $type, [switch]$hidden) {
    $c = New-Object Microsoft.AnalysisServices.Tabular.DataColumn
    $c.Name = $name; $c.SourceColumn = $name
    $c.DataType = [Microsoft.AnalysisServices.Tabular.DataType]::$type
    if ($hidden) { $c.IsHidden = $true }
    return $c
}

function New-MTable($name, $query) {
    if ($model.Tables.ContainsName($name)) { $model.Tables.Remove($name) | Out-Null; $model.SaveChanges() | Out-Null }
    $t = New-Object Microsoft.AnalysisServices.Tabular.Table
    $t.Name = $name
    $mExpr = "let`n    Source = Sql.Database(""$SqlServerInst"", ""$SqlDb"", [Query=""$($query -replace '"','""')""])`nin`n    Source"
    $p = New-Object Microsoft.AnalysisServices.Tabular.Partition
    $p.Name = "$name-Partition"
    $mps = New-Object Microsoft.AnalysisServices.Tabular.MPartitionSource
    $mps.Expression = $mExpr
    $p.Source = $mps
    $t.Partitions.Add($p) | Out-Null
    return $t
}

function New-Measure($tbl, $name, $expr, $fmt, $folder) {
    if ($tbl.Measures.ContainsName($name)) { $tbl.Measures.Remove($name) | Out-Null }
    $m = New-Object Microsoft.AnalysisServices.Tabular.Measure
    $m.Name = $name; $m.Expression = $expr
    if ($fmt)    { $m.FormatString = $fmt }
    if ($folder) { $m.DisplayFolder = $folder }
    $tbl.Measures.Add($m) | Out-Null
}

function New-Rel($fromTable, $fromCol, $toTable, $toCol) {
    $r = New-Object Microsoft.AnalysisServices.Tabular.SingleColumnRelationship
    $r.FromColumn = $model.Tables[$fromTable].Columns[$fromCol]   # many side
    $r.ToColumn   = $model.Tables[$toTable].Columns[$toCol]       # one side
    $r.FromCardinality = [Microsoft.AnalysisServices.Tabular.RelationshipEndCardinality]::Many
    $r.ToCardinality   = [Microsoft.AnalysisServices.Tabular.RelationshipEndCardinality]::One
    $model.Relationships.Add($r) | Out-Null
}

# -----------------------------------------------------------------------------
# 3. Tables
#
# Loaded from the SQL VIEWS, not the raw facts. The verification-currency model
# is defined once in SQL and asserted by 23 acceptance tests; re-deriving
# staleness in DAX would create a second definition free to drift from the
# first, and this project's whole argument depends on one definition of "ready".
# -----------------------------------------------------------------------------
Write-Host "Creating tables..."

# Drop every relationship BEFORE touching the tables. Removing a table does NOT
# cascade to relationships that reference it -- TOM leaves them dangling and
# then refuses the save with "Relationship 'Relationship' points to deleted
# table", which only ever appears on a re-run. RelationshipCollection exposes
# Remove(object), not RemoveAt(index), and the collection must be snapshotted
# with @() before removing from it.
foreach ($rel in @($model.Relationships)) { $model.Relationships.Remove($rel) | Out-Null }
$model.SaveChanges() | Out-Null

$tDate = New-MTable "Dim_Date" @"
SELECT DateKey, [Date], [Year], [Quarter], [Month], MonthName, YearMonth, YearQuarter,
       MonthEndDate, IsMonthEnd, IsWeekend, IsWorkday
FROM dbo.Dim_Date WHERE [Date] BETWEEN '2025-01-01' AND '2027-06-30'
"@
@("DateKey|Int64","Date|DateTime","Year|Int64","Quarter|Int64","Month|Int64","MonthName|String",
  "YearMonth|String","YearQuarter|String","MonthEndDate|DateTime","IsMonthEnd|Boolean",
  "IsWeekend|Boolean","IsWorkday|Boolean") |
  ForEach-Object { $p = $_ -split '\|'; $tDate.Columns.Add((New-Col $p[0] $p[1])) | Out-Null }
$model.Tables.Add($tDate) | Out-Null

$tSub = New-MTable "Dim_Subsystem" @"
SELECT SubsystemCode, SubsystemName, Criticality,
       CAST(RequiresHILRig AS BIT) AS RequiresHILRig, LeadEngineerID
FROM dbo.Dim_Subsystem
"@
@("SubsystemCode|String","SubsystemName|String","Criticality|String",
  "RequiresHILRig|Boolean","LeadEngineerID|String") |
  ForEach-Object { $p = $_ -split '\|'; $tSub.Columns.Add((New-Col $p[0] $p[1])) | Out-Null }
$model.Tables.Add($tSub) | Out-Null

$tPerson = New-MTable "Dim_Person" @"
SELECT PersonID, PersonName, [Role], Team, WeeklyCapacityHours FROM dbo.Dim_Person
"@
@("PersonID|String","PersonName|String","Role|String","Team|String","WeeklyCapacityHours|Double") |
  ForEach-Object { $p = $_ -split '\|'; $tPerson.Columns.Add((New-Col $p[0] $p[1])) | Out-Null }
$model.Tables.Add($tPerson) | Out-Null

$tBuild = New-MTable "Dim_Build" @"
SELECT b.BuildID, b.BuildNumber, d.[Date] AS BuildDate, b.BuildLabel,
       CAST(b.IsReleaseCandidate AS BIT) AS IsReleaseCandidate
FROM dbo.Dim_Build b JOIN dbo.Dim_Date d ON d.DateKey = b.BuildDateKey
"@
@("BuildID|String","BuildNumber|Int64","BuildDate|DateTime","BuildLabel|String","IsReleaseCandidate|Boolean") |
  ForEach-Object { $p = $_ -split '\|'; $tBuild.Columns.Add((New-Col $p[0] $p[1])) | Out-Null }
$model.Tables.Add($tBuild) | Out-Null

$tReq = New-MTable "Requirement" @"
SELECT RequirementID, Title, ReqType, Priority, SubsystemCode, SubsystemName, Criticality,
       MinTestLevel, CAST(RequiresIndependentTester AS BIT) AS RequiresIndependentTester,
       TotalCases, CasesWithRuns, PassingCases, BlockedCases,
       LastPassBuild, LastGoodBuild, PassingButUnderLevelled, PassingButSelfVerified,
       CAST(HasEvidence AS BIT) AS HasEvidence,
       CAST(MeetsPolicy AS BIT) AS MeetsPolicy,
       CAST(StaleByCode AS BIT) AS StaleByCode,
       CAST(StaleByRequirement AS BIT) AS StaleByRequirement,
       CAST(IsCurrent AS BIT) AS IsCurrent,
       VerificationState, InterveningBuilds
FROM dbo.vw_RequirementVerification
"@
@("RequirementID|String","Title|String","ReqType|String","Priority|String","SubsystemCode|String",
  "SubsystemName|String","Criticality|String","MinTestLevel|String","RequiresIndependentTester|Boolean",
  "TotalCases|Int64","CasesWithRuns|Int64","PassingCases|Int64","BlockedCases|Int64",
  "LastPassBuild|Int64","LastGoodBuild|Int64","PassingButUnderLevelled|Int64","PassingButSelfVerified|Int64",
  "HasEvidence|Boolean","MeetsPolicy|Boolean","StaleByCode|Boolean","StaleByRequirement|Boolean",
  "IsCurrent|Boolean","VerificationState|String","InterveningBuilds|Int64") |
  ForEach-Object { $p = $_ -split '\|'; $tReq.Columns.Add((New-Col $p[0] $p[1])) | Out-Null }
$model.Tables.Add($tReq) | Out-Null

$tQueue = New-MTable "Queue" @"
SELECT PriorityRank, RequirementID, SubsystemCode, ReqType, Priority, Criticality,
       OwnerID, OwnerName, OwnerTeam, MinTestLevel, LastGoodBuild, InterveningBuilds,
       Cases, RigHours, TestMinutes, CumulativeRigHours, PriorityScore, ActionCode,
       CAST(IsThisWeek AS BIT) AS IsThisWeek, RecommendedAction
FROM dbo.vw_VerificationQueue
"@
@("PriorityRank|Int64","RequirementID|String","SubsystemCode|String","ReqType|String","Priority|String",
  "Criticality|String","OwnerID|String","OwnerName|String","OwnerTeam|String","MinTestLevel|String",
  "LastGoodBuild|Int64","InterveningBuilds|Int64","Cases|Int64","RigHours|Double","TestMinutes|Int64",
  "CumulativeRigHours|Double","PriorityScore|Double","ActionCode|String","IsThisWeek|Boolean",
  "RecommendedAction|String") |
  ForEach-Object { $p = $_ -split '\|'; $tQueue.Columns.Add((New-Col $p[0] $p[1])) | Out-Null }
$model.Tables.Add($tQueue) | Out-Null

$tRAID = New-MTable "RAID" @"
SELECT RAIDID, RAIDType, Title, SubsystemCode, OwnerID, OwnerName, OwnerTeam,
       RaisedDate, DueDate, ClosedDate, Probability, Impact, ExposureScore, ExposureBand,
       RAIDStatus, CAST(IsOpen AS BIT) AS IsOpen, CAST(IsOverdue AS BIT) AS IsOverdue,
       DaysOverdue, AgeDays, MitigationNote
FROM dbo.vw_RAIDExposure
"@
@("RAIDID|String","RAIDType|String","Title|String","SubsystemCode|String","OwnerID|String",
  "OwnerName|String","OwnerTeam|String","RaisedDate|DateTime","DueDate|DateTime","ClosedDate|DateTime",
  "Probability|Int64","Impact|Int64","ExposureScore|Int64","ExposureBand|String","RAIDStatus|String",
  "IsOpen|Boolean","IsOverdue|Boolean","DaysOverdue|Int64","AgeDays|Int64","MitigationNote|String") |
  ForEach-Object { $p = $_ -split '\|'; $tRAID.Columns.Add((New-Col $p[0] $p[1])) | Out-Null }
$model.Tables.Add($tRAID) | Out-Null

$tWork = New-MTable "WorkItem" @"
SELECT w.WorkItemID, r.RequirementID, s.SubsystemCode, p.PersonID AS AssignedToID,
       od.[Date] AS OpenedDate, cd.[Date] AS ClosedDate,
       w.EstimateHours, w.ActualHours, w.WorkItemStatus
FROM dbo.Fact_WorkItem w
JOIN dbo.Dim_Requirement r ON r.RequirementKey = w.RequirementKey
JOIN dbo.Dim_Subsystem s   ON s.SubsystemKey   = r.SubsystemKey
JOIN dbo.Dim_Person p      ON p.PersonKey      = w.AssignedToKey
JOIN dbo.Dim_Date od       ON od.DateKey       = w.OpenedDateKey
LEFT JOIN dbo.Dim_Date cd  ON cd.DateKey       = w.ClosedDateKey
"@
@("WorkItemID|String","RequirementID|String","SubsystemCode|String","AssignedToID|String",
  "OpenedDate|DateTime","ClosedDate|DateTime","EstimateHours|Double","ActualHours|Double",
  "WorkItemStatus|String") |
  ForEach-Object { $p = $_ -split '\|'; $tWork.Columns.Add((New-Col $p[0] $p[1])) | Out-Null }
$model.Tables.Add($tWork) | Out-Null

$tChurn = New-MTable "BuildChurn" @"
SELECT b.BuildNumber, s.SubsystemCode, c.LinesChanged
FROM dbo.Fact_BuildSubsystemChange c
JOIN dbo.Dim_Build b     ON b.BuildKey = c.BuildKey
JOIN dbo.Dim_Subsystem s ON s.SubsystemKey = c.SubsystemKey
"@
@("BuildNumber|Int64","SubsystemCode|String","LinesChanged|Int64") |
  ForEach-Object { $p = $_ -split '\|'; $tChurn.Columns.Add((New-Col $p[0] $p[1])) | Out-Null }
$model.Tables.Add($tChurn) | Out-Null

$tTargets = New-MTable "Ref_ReadinessTargets" @"
SELECT MetricName, TargetValue, WarningValue, Direction, Unit, [Description] FROM dbo.Ref_ReadinessTargets
"@
@("MetricName|String","TargetValue|Double","WarningValue|Double","Direction|String","Unit|String","Description|String") |
  ForEach-Object { $p = $_ -split '\|'; $tTargets.Columns.Add((New-Col $p[0] $p[1])) | Out-Null }
$model.Tables.Add($tTargets) | Out-Null

$tPolicy = New-MTable "Ref_VerificationPolicy" @"
SELECT ReqType, MinTestLevel, MinTestLevelRank,
       CAST(RequiresIndependentTester AS BIT) AS RequiresIndependentTester, PolicyNote
FROM dbo.Ref_VerificationPolicy
"@
@("ReqType|String","MinTestLevel|String","MinTestLevelRank|Int64",
  "RequiresIndependentTester|Boolean","PolicyNote|String") |
  ForEach-Object { $p = $_ -split '\|'; $tPolicy.Columns.Add((New-Col $p[0] $p[1])) | Out-Null }
$model.Tables.Add($tPolicy) | Out-Null

# The reporting anchor. One row, read straight out of the KPI view, so the
# model and the SQL scorecard take their as-of from the same source and cannot
# drift. Anchoring on MAX(BuildNumber) instead would look identical today and
# would silently reassess against an experimental build the moment one appeared
# after the release candidate.
$tRep = New-MTable "Ref_Reporting" @"
SELECT AsOfBuild, AsOfDate FROM dbo.vw_ReadinessKPI
"@
@("AsOfBuild|Int64","AsOfDate|DateTime") |
  ForEach-Object { $p = $_ -split '\|'; $tRep.Columns.Add((New-Col $p[0] $p[1])) | Out-Null }
$model.Tables.Add($tRep) | Out-Null

$model.SaveChanges() | Out-Null
# Counted, not asserted. This read "11 tables created" while the script added
# twelve -- Ref_Reporting arrived last and the literal above it never moved. A
# hardcoded count of a thing the code builds is a claim that rots silently,
# and the measure count forty lines below already does this correctly.
Write-Host "  $($model.Tables.Count) tables created"

# -----------------------------------------------------------------------------
# 4. Refresh
# -----------------------------------------------------------------------------
Write-Host ""
Write-Host "Refreshing tables. If a credential prompt appears in Power BI Desktop," -ForegroundColor Yellow
Write-Host "click through it (Windows / Integrated auth, any privacy level)." -ForegroundColor Yellow
$allTables = @("Dim_Date","Dim_Subsystem","Dim_Person","Dim_Build","Requirement","Queue",
               "RAID","WorkItem","BuildChurn","Ref_ReadinessTargets","Ref_VerificationPolicy","Ref_Reporting")
foreach ($tn in $allTables) {
    $model.Tables[$tn].RequestRefresh([Microsoft.AnalysisServices.Tabular.RefreshType]::Full)
}
$model.SaveChanges() | Out-Null
# Row counts are deliberately NOT printed. A Desktop refresh does not populate
# EstimatedRowCount, and anything structural would report success whether or
# not a single row loaded. Validate_PowerBI_Model.ps1 queries the model with
# DAX and reconciles against SQL, which is the only check worth trusting.
Write-Host "  Refresh submitted -- verify it with Validate_PowerBI_Model.ps1"

# -----------------------------------------------------------------------------
# 5. Relationships
# -----------------------------------------------------------------------------
Write-Host ""
Write-Host "Creating relationships..."
New-Rel "Requirement" "SubsystemCode" "Dim_Subsystem" "SubsystemCode"
New-Rel "Requirement" "ReqType"       "Ref_VerificationPolicy" "ReqType"
New-Rel "Queue"       "RequirementID" "Requirement"   "RequirementID"
New-Rel "WorkItem"    "RequirementID" "Requirement"   "RequirementID"
New-Rel "RAID"        "SubsystemCode" "Dim_Subsystem" "SubsystemCode"
New-Rel "RAID"        "OwnerID"       "Dim_Person"    "PersonID"
New-Rel "BuildChurn"  "BuildNumber"   "Dim_Build"     "BuildNumber"
New-Rel "BuildChurn"  "SubsystemCode" "Dim_Subsystem" "SubsystemCode"
New-Rel "Dim_Build"   "BuildDate"     "Dim_Date"      "Date"
$model.SaveChanges() | Out-Null
Write-Host "  9 relationships created"

$tDate2 = $model.Tables["Dim_Date"]
$tDate2.DataCategory = "Time"
$tDate2.Columns["Date"].IsKey = $true
$model.SaveChanges() | Out-Null

$h = New-Object Microsoft.AnalysisServices.Tabular.Hierarchy
$h.Name = "Calendar"
foreach ($lv in @(@("Year","Year"), @("Quarter","YearQuarter"), @("Month","YearMonth"), @("Date","Date"))) {
    $lvl = New-Object Microsoft.AnalysisServices.Tabular.Level
    $lvl.Name = $lv[0]; $lvl.Column = $tDate2.Columns[$lv[1]]; $lvl.Ordinal = $h.Levels.Count
    $h.Levels.Add($lvl) | Out-Null
}
if ($tDate2.Hierarchies.ContainsName("Calendar")) { $tDate2.Hierarchies.Remove("Calendar") | Out-Null }
$tDate2.Hierarchies.Add($h) | Out-Null
$model.SaveChanges() | Out-Null
Write-Host "  Dim_Date marked as a date table; Calendar hierarchy added"

# -----------------------------------------------------------------------------
# 6. Measures
# -----------------------------------------------------------------------------
Write-Host ""
Write-Host "Creating measures..."
$R = $model.Tables["Requirement"]
$Q = $model.Tables["Queue"]
$D = $model.Tables["RAID"]
$W = $model.Tables["WorkItem"]
$C = $model.Tables["BuildChurn"]
$P = $model.Tables["Ref_Reporting"]

# --- Reporting anchor -----------------------------------------------------
New-Measure $P "As Of Build" "MAX ( Ref_Reporting[AsOfBuild] )" '#,0' "00 Reporting"
New-Measure $P "As Of Date"  "MAX ( Ref_Reporting[AsOfDate] )"  'yyyy-mm-dd' "00 Reporting"

# --- Population -----------------------------------------------------------
New-Measure $R "Requirements"  "COUNTROWS ( Requirement )" '#,0' "01 Population"
New-Measure $R "Must Ship"     "CALCULATE ( COUNTROWS ( Requirement ), Requirement[Priority] = ""MustShip"" )" '#,0' "01 Population"
New-Measure $R "Should Ship"   "CALCULATE ( COUNTROWS ( Requirement ), Requirement[Priority] = ""ShouldShip"" )" '#,0' "01 Population"

# --- The contrast ---------------------------------------------------------
# The two headline numbers, defined next to each other on purpose. Anyone
# editing one should have to look at the other.
New-Measure $W "Work Items"           "COUNTROWS ( WorkItem )" '#,0' "02 The contrast"
New-Measure $W "Work Items Closed"    "CALCULATE ( COUNTROWS ( WorkItem ), WorkItem[WorkItemStatus] = ""Closed"" )" '#,0' "02 The contrast"
New-Measure $W "Work Item Completion %" "DIVIDE ( [Work Items Closed], [Work Items] ) * 100" '0.00' "02 The contrast"

# Readiness is over MUST-SHIP only. A readiness percentage that mixes
# must-ship and nice-to-have is arithmetic with no decision attached: it cannot
# tell you whether to ship.
New-Measure $R "Current Must Ship" @"
CALCULATE ( COUNTROWS ( Requirement ),
    Requirement[Priority] = "MustShip", Requirement[IsCurrent] = TRUE () )
"@ '#,0' "02 The contrast"
New-Measure $R "Ship Readiness %" "DIVIDE ( [Current Must Ship], [Must Ship] ) * 100" '0.00' "02 The contrast"
New-Measure $R "Readiness Gap vs Reported" "[Work Item Completion %] - [Ship Readiness %]" '0.00' "02 The contrast"

# --- Verification states --------------------------------------------------
New-Measure $R "Has Evidence"        "CALCULATE ( COUNTROWS ( Requirement ), Requirement[HasEvidence] = TRUE () )" '#,0' "03 Verification"
New-Measure $R "Verification Coverage %" "DIVIDE ( [Has Evidence], [Requirements] ) * 100" '0.00' "03 Verification"
New-Measure $R "Meets Policy"        "CALCULATE ( COUNTROWS ( Requirement ), Requirement[MeetsPolicy] = TRUE () )" '#,0' "03 Verification"
New-Measure $R "Policy Compliance %" "DIVIDE ( [Meets Policy], [Has Evidence] ) * 100" '0.00' "03 Verification"
New-Measure $R "Current"             "CALCULATE ( COUNTROWS ( Requirement ), Requirement[IsCurrent] = TRUE () )" '#,0' "03 Verification"

# Denominator is requirements that HAVE policy-compliant evidence, not all
# requirements. Mixing "never verified properly" into a staleness rate makes
# staleness look like a documentation problem instead of an evidence one.
New-Measure $R "Stale" @"
COUNTROWS (
    FILTER ( Requirement,
        Requirement[MeetsPolicy] = TRUE ()
        && ( Requirement[StaleByCode] = TRUE () || Requirement[StaleByRequirement] = TRUE () ) ) )
"@ '#,0' "03 Verification"
New-Measure $R "Stale Verification %" "DIVIDE ( [Stale], [Meets Policy] ) * 100" '0.00' "03 Verification"

New-Measure $R "No Evidence"     "CALCULATE ( COUNTROWS ( Requirement ), Requirement[HasEvidence] = FALSE () )" '#,0' "03 Verification"
New-Measure $R "Insufficient"    "CALCULATE ( COUNTROWS ( Requirement ), Requirement[HasEvidence] = TRUE (), Requirement[MeetsPolicy] = FALSE () )" '#,0' "03 Verification"

# The two stale buckets are made MUTUALLY EXCLUSIVE, and in the same order the
# SQL view uses: a requirement that is stale on both counts is reported as
# "requirement changed", because that is the one that needs a human to reread
# the requirement before anyone re-runs anything. Ordering them the other way
# in DAX than in SQL would make the two tools report different splits of the
# same total -- and the total would still agree, so nothing would look wrong.
New-Measure $R "Stale by Requirement" @"
CALCULATE ( COUNTROWS ( Requirement ),
    Requirement[MeetsPolicy] = TRUE (), Requirement[StaleByRequirement] = TRUE () )
"@ '#,0' "03 Verification"
New-Measure $R "Stale by Code" @"
CALCULATE ( COUNTROWS ( Requirement ),
    Requirement[MeetsPolicy] = TRUE (),
    Requirement[StaleByRequirement] = FALSE (), Requirement[StaleByCode] = TRUE () )
"@ '#,0' "03 Verification"

New-Measure $R "Self Verified"   "CALCULATE ( COUNTROWS ( Requirement ), Requirement[PassingButSelfVerified] > 0, Requirement[MeetsPolicy] = FALSE () )" '#,0' "03 Verification"
New-Measure $R "Under Levelled"  "CALCULATE ( COUNTROWS ( Requirement ), Requirement[PassingButUnderLevelled] > 0, Requirement[MeetsPolicy] = FALSE () )" '#,0' "03 Verification"

# --- Queue and schedule ---------------------------------------------------
New-Measure $Q "Outstanding"          "COUNTROWS ( Queue )" '#,0' "04 Queue"
New-Measure $Q "Rig Hours Outstanding" "SUM ( Queue[RigHours] )" '#,0.0' "04 Queue"
# The only figure here a programme board can act on without a further study.
New-Measure $Q "Rig Weeks Outstanding" "DIVIDE ( [Rig Hours Outstanding], 180 )" '0.0' "04 Queue"
New-Measure $Q "Schedulable This Week" "CALCULATE ( COUNTROWS ( Queue ), Queue[IsThisWeek] = TRUE () )" '#,0' "04 Queue"
New-Measure $Q "Week Coverage %"       "DIVIDE ( [Schedulable This Week], [Outstanding] ) * 100" '0.0' "04 Queue"
New-Measure $Q "Avg Intervening Builds" "AVERAGE ( Queue[InterveningBuilds] )" '0.0' "04 Queue"

# --- Churn ----------------------------------------------------------------
New-Measure $C "Builds Changed"  "DISTINCTCOUNT ( BuildChurn[BuildNumber] )" '#,0' "05 Churn"
New-Measure $C "Lines Changed"   "SUM ( BuildChurn[LinesChanged] )" '#,0' "05 Churn"
# Churn per readiness point. A subsystem that changed forty times and is a
# quarter ready is not behind on testing -- it is still being designed, which
# is a different conversation with a different owner.
New-Measure $C "Churn per Readiness Point" "DIVIDE ( [Builds Changed], [Ship Readiness %] )" '0.00' "05 Churn"

# --- RAID -----------------------------------------------------------------
New-Measure $D "RAID Items"      "COUNTROWS ( RAID )" '#,0' "06 RAID"
New-Measure $D "RAID Open"       "CALCULATE ( COUNTROWS ( RAID ), RAID[IsOpen] = TRUE () )" '#,0' "06 RAID"
New-Measure $D "RAID Overdue"    "CALCULATE ( COUNTROWS ( RAID ), RAID[IsOverdue] = TRUE () )" '#,0' "06 RAID"
New-Measure $D "Overdue RAID %"  "DIVIDE ( [RAID Overdue], [RAID Open] ) * 100" '0.00' "06 RAID"
New-Measure $D "Critical RAID Open" "CALCULATE ( COUNTROWS ( RAID ), RAID[ExposureBand] = ""Critical"", RAID[IsOpen] = TRUE () )" '#,0' "06 RAID"
New-Measure $D "Open RAID Exposure" "CALCULATE ( SUM ( RAID[ExposureScore] ), RAID[IsOpen] = TRUE () )" '#,0' "06 RAID"

# --- RAG colours ----------------------------------------------------------
# A colour measure returns a HEX STRING, not a number. Power BI's "format by
# field value" expects a colour; returning 1/2/3 and hoping the rule interprets
# it is the commonest reason a conditional-format rule silently does nothing.
#
# Direction is read from the target table rather than hardcoded, so adding a
# LowerBetter metric cannot silently invert its status.
$ragMetrics = @(
    @("Ship Readiness",       "ShipReadinessPct",        "[Ship Readiness %]",        $R),
    @("Verification Coverage","VerificationCoveragePct", "[Verification Coverage %]", $R),
    @("Stale Verification",   "StaleVerificationPct",    "[Stale Verification %]",    $R),
    @("Policy Compliance",    "PolicyCompliancePct",     "[Policy Compliance %]",     $R),
    @("Work Item Completion", "WorkItemCompletionPct",   "[Work Item Completion %]",  $W),
    @("Critical RAID",        "CriticalRAIDOpen",        "[Critical RAID Open]",      $D),
    @("Overdue RAID",         "OverdueRAIDPct",          "[Overdue RAID %]",          $D)
)
foreach ($m in $ragMetrics) {
    $label = $m[0]; $metric = $m[1]; $val = $m[2]; $tbl = $m[3]
    New-Measure $tbl "$label Colour" @"
VAR V = $val
VAR T = LOOKUPVALUE ( Ref_ReadinessTargets[TargetValue],  Ref_ReadinessTargets[MetricName], "$metric" )
VAR W = LOOKUPVALUE ( Ref_ReadinessTargets[WarningValue], Ref_ReadinessTargets[MetricName], "$metric" )
VAR Dir = LOOKUPVALUE ( Ref_ReadinessTargets[Direction],  Ref_ReadinessTargets[MetricName], "$metric" )
RETURN
    IF ( Dir = "HigherBetter",
         SWITCH ( TRUE (), V >= T, "#C6EFCE", V >= W, "#FFEB9C", "#FFC7CE" ),
         SWITCH ( TRUE (), V <= T, "#C6EFCE", V <= W, "#FFEB9C", "#FFC7CE" ) )
"@ $null "07 Status"
    New-Measure $tbl "$label Status" @"
VAR V = $val
VAR T = LOOKUPVALUE ( Ref_ReadinessTargets[TargetValue],  Ref_ReadinessTargets[MetricName], "$metric" )
VAR W = LOOKUPVALUE ( Ref_ReadinessTargets[WarningValue], Ref_ReadinessTargets[MetricName], "$metric" )
VAR Dir = LOOKUPVALUE ( Ref_ReadinessTargets[Direction],  Ref_ReadinessTargets[MetricName], "$metric" )
RETURN
    IF ( Dir = "HigherBetter",
         SWITCH ( TRUE (), V >= T, "Green", V >= W, "Amber", "Red" ),
         SWITCH ( TRUE (), V <= T, "Green", V <= W, "Amber", "Red" ) )
"@ $null "07 Status"
}

$model.SaveChanges() | Out-Null
$measureCount = ($model.Tables | ForEach-Object { $_.Measures.Count } | Measure-Object -Sum).Sum
# Counted, never asserted -- a hardcoded folder count is a number a reader
# takes on trust because it sits beside one that is computed.
$folderCount = ($model.Tables | ForEach-Object { $_.Measures } |
                Where-Object { $_.DisplayFolder } |
                Select-Object -ExpandProperty DisplayFolder -Unique).Count
Write-Host "  $measureCount measures created across $folderCount display folders"

# -----------------------------------------------------------------------------
# 7. Hide keys a reader might drag onto a visual and sum
# -----------------------------------------------------------------------------
foreach ($pair in @(@("Requirement","LastPassBuild"), @("Queue","PriorityScore"),
                    @("Queue","CumulativeRigHours"), @("Dim_Date","DateKey"))) {
    $model.Tables[$pair[0]].Columns[$pair[1]].IsHidden = $true
}
$model.SaveChanges() | Out-Null

Write-Host ("-" * 78)
Write-Host "Model built. Save the file in Power BI Desktop as powerbi/Talon_Readiness_Model.pbix," -ForegroundColor Green
Write-Host "then build the report pages using POWER_BI_BUILD_GUIDE.md." -ForegroundColor Green
$server.Disconnect()
