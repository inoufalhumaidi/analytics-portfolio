# =============================================================================
# Project 1 -- Ridgeline Field Services: Operational Efficiency
# Build_PowerBI_Model.ps1
#
# Scripts the Power BI data model (tables, relationships, DAX measures)
# directly into an OPEN Power BI Desktop session via its local Analysis
# Services engine -- the same mechanism Tabular Editor / DAX Studio use.
#
# WHAT THIS DOES: builds the star schema (7 tables imported from SQL Server
# via Power Query/M) + 6 relationships (5 active, 1 deliberately inactive) +
# 31 DAX measures + a Year > MonthName date hierarchy (see DAX_Measures.md)
# entirely by script.
#
# WHAT THIS DOES NOT DO: build report pages/visuals. Power BI's report
# layout has no supported scripting API -- follow POWER_BI_BUILD_GUIDE.md
# Section 5 to add visuals on top of the model this script builds.
#
# PREREQUISITES
#   1. Power BI Desktop open with a new BLANK report (File > New, don't add
#      any data source yet -- this script adds all of them).
#   2. sql/01_create_schema.sql through sql/02_generate_synthetic_data.sql
#      already run against your SQL Server instance.
#   3. First run of each new SQL Server connection: Power BI Desktop will
#      show a credentials/privacy-level dialog the first time it processes
#      each table. Watch the Desktop window and click through it (choose
#      Windows/Integrated auth, any privacy level) -- the script will
#      appear to hang until you do. This only happens once per session.
#
# This script auto-downloads the Microsoft.AnalysisServices.retail.amd64
# NuGet package (official Microsoft package, ~6MB) on first run if the
# client DLLs aren't already present in .amo_libs/ next to this script --
# it isn't vendored into the repo to keep it lean.
# =============================================================================

$ErrorActionPreference = "Stop"

$SqlServerInst = "localhost\TEW_SQLEXPRESS"
$SqlDb = "RidgelineFieldOps"

# -----------------------------------------------------------------------------
# 0. Ensure the Analysis Services client DLLs are available
# -----------------------------------------------------------------------------
$libDir = Join-Path $PSScriptRoot ".amo_libs"
$dllDir = Join-Path $libDir "lib\net45"
if (-not (Test-Path (Join-Path $dllDir "Microsoft.AnalysisServices.Tabular.dll"))) {
    Write-Host "Analysis Services client libraries not found -- downloading Microsoft.AnalysisServices.retail.amd64 from nuget.org..."
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
# 1. Find the running Power BI Desktop instance's local AS engine port
# -----------------------------------------------------------------------------
$msmdsrv = Get-Process -Name "msmdsrv" -ErrorAction SilentlyContinue
if (-not $msmdsrv) { throw "No Power BI Desktop session found. Open Power BI Desktop with a blank report first." }
$conn = Get-NetTCPConnection -OwningProcess $msmdsrv.Id -State Listen -ErrorAction SilentlyContinue | Select-Object -First 1
if (-not $conn) { throw "Could not find the local AS engine's listening port." }
$port = $conn.LocalPort
Write-Host "Found Power BI Desktop's local engine on port $port"

$server = New-Object Microsoft.AnalysisServices.Tabular.Server
$server.Connect("localhost:$port")
$db = $server.Databases[0]
$model = $db.Model
Write-Host "Connected to workspace DB $($db.Name)"

# -----------------------------------------------------------------------------
# 2. Helpers
# -----------------------------------------------------------------------------
function New-Col($name, $type, [switch]$hidden) {
    $c = New-Object Microsoft.AnalysisServices.Tabular.DataColumn
    $c.Name = $name
    $c.SourceColumn = $name
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

function New-Measure($fact, $name, $expr, $fmt, $folder) {
    if ($fact.Measures.ContainsName($name)) { $fact.Measures.Remove($name) | Out-Null }
    $m = New-Object Microsoft.AnalysisServices.Tabular.Measure
    $m.Name = $name
    $m.Expression = $expr
    if ($fmt) { $m.FormatString = $fmt }
    if ($folder) { $m.DisplayFolder = $folder }
    $fact.Measures.Add($m) | Out-Null
}

# -----------------------------------------------------------------------------
# 3. Tables
# -----------------------------------------------------------------------------
$t = New-MTable "Dim_Date" "SELECT DateKey, [Date], [Year], [Quarter], [Month], MonthName, [Week], DayOfWeek, DayName, CAST(IsWeekend AS INT) AS IsWeekend FROM dbo.Dim_Date"
$t.Columns.Add((New-Col "DateKey" "Int64" -hidden)) | Out-Null
$t.Columns.Add((New-Col "Date" "DateTime")) | Out-Null
$t.Columns.Add((New-Col "Year" "Int64")) | Out-Null
$t.Columns.Add((New-Col "Quarter" "Int64")) | Out-Null
$t.Columns.Add((New-Col "Month" "Int64")) | Out-Null
$t.Columns.Add((New-Col "MonthName" "String")) | Out-Null
$t.Columns.Add((New-Col "Week" "Int64")) | Out-Null
$t.Columns.Add((New-Col "DayOfWeek" "Int64")) | Out-Null
$t.Columns.Add((New-Col "DayName" "String")) | Out-Null
$t.Columns.Add((New-Col "IsWeekend" "Int64")) | Out-Null
$model.Tables.Add($t) | Out-Null
Write-Host "Added Dim_Date"

$t = New-MTable "Dim_Region" "SELECT RegionKey, RegionCode, RegionName, TechnicianTarget FROM dbo.Dim_Region"
$t.Columns.Add((New-Col "RegionKey" "Int64" -hidden)) | Out-Null
$t.Columns.Add((New-Col "RegionCode" "String")) | Out-Null
$t.Columns.Add((New-Col "RegionName" "String")) | Out-Null
$t.Columns.Add((New-Col "TechnicianTarget" "Int64")) | Out-Null
$model.Tables.Add($t) | Out-Null
Write-Host "Added Dim_Region"

$t = New-MTable "Dim_Technician" "SELECT TechnicianKey, TechnicianID, TechnicianName, RegionKey, SkillLevel, HireDate, EmploymentStatus, DailyCapacityHours FROM dbo.Dim_Technician"
$t.Columns.Add((New-Col "TechnicianKey" "Int64" -hidden)) | Out-Null
$t.Columns.Add((New-Col "TechnicianID" "String")) | Out-Null
$t.Columns.Add((New-Col "TechnicianName" "String")) | Out-Null
$t.Columns.Add((New-Col "RegionKey" "Int64" -hidden)) | Out-Null
$t.Columns.Add((New-Col "SkillLevel" "String")) | Out-Null
$t.Columns.Add((New-Col "HireDate" "DateTime")) | Out-Null
$t.Columns.Add((New-Col "EmploymentStatus" "String")) | Out-Null
$t.Columns.Add((New-Col "DailyCapacityHours" "Double")) | Out-Null
$model.Tables.Add($t) | Out-Null
Write-Host "Added Dim_Technician"

$t = New-MTable "Dim_Client" "SELECT ClientKey, ClientID, ClientName, ClientType, RegionKey, ContractTier FROM dbo.Dim_Client"
$t.Columns.Add((New-Col "ClientKey" "Int64" -hidden)) | Out-Null
$t.Columns.Add((New-Col "ClientID" "String")) | Out-Null
$t.Columns.Add((New-Col "ClientName" "String")) | Out-Null
$t.Columns.Add((New-Col "ClientType" "String")) | Out-Null
$t.Columns.Add((New-Col "RegionKey" "Int64" -hidden)) | Out-Null
$t.Columns.Add((New-Col "ContractTier" "String")) | Out-Null
$model.Tables.Add($t) | Out-Null
Write-Host "Added Dim_Client"

$t = New-MTable "Dim_ServiceType" "SELECT ServiceTypeKey, ServiceTypeID, ServiceCategory, ServiceName, StandardDurationMin, SLAHours, BaseRevenue, BaseCost FROM dbo.Dim_ServiceType"
$t.Columns.Add((New-Col "ServiceTypeKey" "Int64" -hidden)) | Out-Null
$t.Columns.Add((New-Col "ServiceTypeID" "String")) | Out-Null
$t.Columns.Add((New-Col "ServiceCategory" "String")) | Out-Null
$t.Columns.Add((New-Col "ServiceName" "String")) | Out-Null
$t.Columns.Add((New-Col "StandardDurationMin" "Int64")) | Out-Null
$t.Columns.Add((New-Col "SLAHours" "Double")) | Out-Null
$t.Columns.Add((New-Col "BaseRevenue" "Double")) | Out-Null
$t.Columns.Add((New-Col "BaseCost" "Double")) | Out-Null
$model.Tables.Add($t) | Out-Null
Write-Host "Added Dim_ServiceType"

$t = New-MTable "Ref_KPITargets" "SELECT MetricName, TargetValue, WarningValue, Direction, Description FROM dbo.Ref_KPITargets"
$t.Columns.Add((New-Col "MetricName" "String")) | Out-Null
$t.Columns.Add((New-Col "TargetValue" "Double")) | Out-Null
$t.Columns.Add((New-Col "WarningValue" "Double")) | Out-Null
$t.Columns.Add((New-Col "Direction" "String")) | Out-Null
$t.Columns.Add((New-Col "Description" "String")) | Out-Null
$model.Tables.Add($t) | Out-Null
Write-Host "Added Ref_KPITargets"

$factQuery = "SELECT JobKey, JobID, DateKey, TechnicianKey, ClientKey, ServiceTypeKey, RegionKey, ScheduledStart, ActualStart, ActualEnd, TravelTimeMin, JobDurationMin, CAST(FirstTimeFix AS INT) AS FirstTimeFix, CAST(CallbackFlag AS INT) AS CallbackFlag, CAST(SLAMet AS INT) AS SLAMet, JobCost, JobRevenue, JobStatus FROM dbo.Fact_ServiceJobs"
$t = New-MTable "Fact_ServiceJobs" $factQuery
$t.Columns.Add((New-Col "JobKey" "Int64" -hidden)) | Out-Null
$t.Columns.Add((New-Col "JobID" "String")) | Out-Null
$t.Columns.Add((New-Col "DateKey" "Int64" -hidden)) | Out-Null
$t.Columns.Add((New-Col "TechnicianKey" "Int64" -hidden)) | Out-Null
$t.Columns.Add((New-Col "ClientKey" "Int64" -hidden)) | Out-Null
$t.Columns.Add((New-Col "ServiceTypeKey" "Int64" -hidden)) | Out-Null
$t.Columns.Add((New-Col "RegionKey" "Int64" -hidden)) | Out-Null
$t.Columns.Add((New-Col "ScheduledStart" "DateTime")) | Out-Null
$t.Columns.Add((New-Col "ActualStart" "DateTime")) | Out-Null
$t.Columns.Add((New-Col "ActualEnd" "DateTime")) | Out-Null
$t.Columns.Add((New-Col "TravelTimeMin" "Int64")) | Out-Null
$t.Columns.Add((New-Col "JobDurationMin" "Int64")) | Out-Null
$t.Columns.Add((New-Col "FirstTimeFix" "Int64")) | Out-Null
$t.Columns.Add((New-Col "CallbackFlag" "Int64")) | Out-Null
$t.Columns.Add((New-Col "SLAMet" "Int64")) | Out-Null
$t.Columns.Add((New-Col "JobCost" "Double")) | Out-Null
$t.Columns.Add((New-Col "JobRevenue" "Double")) | Out-Null
$t.Columns.Add((New-Col "JobStatus" "String")) | Out-Null
$model.Tables.Add($t) | Out-Null
Write-Host "Added Fact_ServiceJobs"

Write-Host "Saving table structures..."
$model.SaveChanges() | Out-Null

Write-Host "Processing all tables (Power BI Desktop may prompt for credentials on first run -- watch the window)..."
foreach ($tn in @("Dim_Date","Dim_Region","Dim_Technician","Dim_Client","Dim_ServiceType","Ref_KPITargets","Fact_ServiceJobs")) {
    $model.Tables[$tn].RequestRefresh([Microsoft.AnalysisServices.Tabular.RefreshType]::Full)
}
$model.SaveChanges() | Out-Null
Write-Host "All tables processed."

# -----------------------------------------------------------------------------
# 4. Relationships
# -----------------------------------------------------------------------------
function New-Rel($fromTable, $fromCol, $toTable, $toCol, [bool]$active = $true) {
    $r = New-Object Microsoft.AnalysisServices.Tabular.SingleColumnRelationship
    $r.FromColumn = $model.Tables[$fromTable].Columns[$fromCol]
    $r.ToColumn = $model.Tables[$toTable].Columns[$toCol]
    $r.IsActive = $active
    $model.Relationships.Add($r) | Out-Null
}
New-Rel "Fact_ServiceJobs" "DateKey" "Dim_Date" "DateKey"
New-Rel "Fact_ServiceJobs" "TechnicianKey" "Dim_Technician" "TechnicianKey"
New-Rel "Fact_ServiceJobs" "ClientKey" "Dim_Client" "ClientKey"
New-Rel "Fact_ServiceJobs" "ServiceTypeKey" "Dim_ServiceType" "ServiceTypeKey"

# Region is deliberately wired as a SNOWFLAKE: Dim_Region -> Dim_Technician -> Fact,
# NOT the direct Fact -> Dim_Region path. Capacity/utilization measures iterate
# Dim_Technician, and a filter on Dim_Region only reaches the fact table -- it does
# NOT flow dimension -> fact -> dimension. With the direct path active, slicing a
# visual by Dim_Region[RegionName] narrowed [Worked Minutes] but left
# [Available Minutes] at the full-company total, silently understating regional
# Utilization %. Routing region through the technician fixes it with no ambiguity.
# Safe because Fact_ServiceJobs.RegionKey always equals the technician's RegionKey
# (the generator assigns jobs from the technician's own region) -- verified with:
#   SELECT COUNT(*) FROM Fact_ServiceJobs f JOIN Dim_Technician t
#     ON t.TechnicianKey=f.TechnicianKey WHERE f.RegionKey <> t.RegionKey;  -- returns 0
New-Rel "Dim_Technician" "RegionKey" "Dim_Region" "RegionKey" $true
New-Rel "Fact_ServiceJobs" "RegionKey" "Dim_Region" "RegionKey" $false
$model.SaveChanges() | Out-Null
Write-Host "Relationships created (region via snowflake through Dim_Technician)."

# -----------------------------------------------------------------------------
# 4b. Date table presentation: month-name sort order + Year/Month hierarchy.
#     Keep all hierarchy levels inside ONE table -- a slicer hierarchy mixing
#     tables (e.g. Dim_Region + Dim_Date) fails with
#     DataViewMappingError_ConditionNotHierarchicallyRelated.
#     Do NOT give Dim_Date[Date] a sort-by column; sorting it by Month
#     interleaves both years on any date axis.
# -----------------------------------------------------------------------------
$dimDate = $model.Tables["Dim_Date"]
$dimDate.Columns["MonthName"].SortByColumn = $dimDate.Columns["Month"]

$hierName = "Year/Month Hierarchy"
if ($dimDate.Hierarchies.ContainsName($hierName)) { $dimDate.Hierarchies.Remove($hierName) | Out-Null }
$h = New-Object Microsoft.AnalysisServices.Tabular.Hierarchy
$h.Name = $hierName
$ordinal = 0
foreach ($colName in @("Year", "MonthName")) {   # level names are what report visuals bind to -- keep "MonthName"
    $lvl = New-Object Microsoft.AnalysisServices.Tabular.Level
    $lvl.Name = $colName
    $lvl.Ordinal = $ordinal
    $lvl.Column = $dimDate.Columns[$colName]
    $h.Levels.Add($lvl) | Out-Null
    $ordinal++
}
$dimDate.Hierarchies.Add($h) | Out-Null
$model.SaveChanges() | Out-Null
Write-Host "Dim_Date: MonthName sorted by Month; '$hierName' (Year > MonthName) created."

# -----------------------------------------------------------------------------
# 5. DAX measures (all added to Fact_ServiceJobs, organized by DisplayFolder --
#    see DAX_Measures.md for the annotated version of each formula)
# -----------------------------------------------------------------------------
$fact = $model.Tables["Fact_ServiceJobs"]
$pct = "0.0%"; $money = "$#,##0"

New-Measure $fact "Total Jobs" "COUNTROWS ( Fact_ServiceJobs )" "#,##0" "Volume"
New-Measure $fact "Completed Jobs" 'CALCULATE ( [Total Jobs], Fact_ServiceJobs[JobStatus] = "Completed" )' "#,##0" "Volume"
New-Measure $fact "Cancelled Jobs" 'CALCULATE ( [Total Jobs], Fact_ServiceJobs[JobStatus] = "Cancelled" )' "#,##0" "Volume"
New-Measure $fact "Completion Rate %" "DIVIDE ( [Completed Jobs], [Total Jobs] )" $pct "Volume"

New-Measure $fact "First-Time-Fix %" 'CALCULATE ( AVERAGE ( Fact_ServiceJobs[FirstTimeFix] ), Fact_ServiceJobs[JobStatus] = "Completed" )' $pct "Quality"
New-Measure $fact "Callback %" 'CALCULATE ( AVERAGE ( Fact_ServiceJobs[CallbackFlag] ), Fact_ServiceJobs[JobStatus] = "Completed" )' $pct "Quality"
New-Measure $fact "SLA Compliance %" 'CALCULATE ( AVERAGE ( Fact_ServiceJobs[SLAMet] ), Fact_ServiceJobs[JobStatus] = "Completed" )' $pct "Quality"

# NOTE: RELATED() cannot be used here. CROSSJOIN(VALUES(...), VALUES(...)) builds a
# virtual two-column table that has no relationship to Dim_Technician, so
# RELATED(Dim_Technician[EmploymentStatus]) fails with "either doesn't exist or
# doesn't have a relationship to any table available in the current context".
# Iterating the dimension tables directly keeps every column in row context.
$availMinExpr = @'
VAR WeekdayDates =
    FILTER ( Dim_Date, Dim_Date[IsWeekend] = 0 )
RETURN
    SUMX (
        FILTER ( Dim_Technician, Dim_Technician[EmploymentStatus] = "Active" ),
        VAR HireD = Dim_Technician[HireDate]
        VAR CapacityMinutes = Dim_Technician[DailyCapacityHours] * 60
        RETURN
            COUNTROWS ( FILTER ( WeekdayDates, Dim_Date[Date] >= HireD ) ) * CapacityMinutes
    )
'@
New-Measure $fact "Available Minutes" $availMinExpr "#,##0" "Utilization"
# THE WEEKDAY FILTER BELOW IS THE WHOLE OF THE b1447a6 FIX, AND IT WAS MISSING
# HERE FOR THE ENTIRE LIFE OF THIS MODEL.
#
# [Available Minutes] above filters Dim_Date[IsWeekend] = 0, because capacity is
# modelled on scheduled weekdays. [Worked Minutes] filtered on JobStatus alone,
# so [Utilization %] divided EVERY completed job -- weekend call-outs included --
# by weekday-only capacity. That is precisely the defect
# docs/data_validation_report.md section 5b describes as corrected.
#
# It was corrected in SQL, in Excel, in the JSON exports, in the dashboard and
# in the prose. Nothing under powerbi/ was touched, so the shipped report
# evaluated to 70.455% portfolio utilization and 50.491% for South Metro while
# every other artefact in the project said 65.891% and 47.243%. All five regions
# differed, and two of them sat above a 75% target they do not meet.
#
# The lesson is the one the portfolio keeps relearning: a fix applied to "the
# code" is not applied to the artefacts that re-implement the same definition.
# Excel was checked because a validator reads it back; the .pbix was not,
# because verifying it needs Power BI Desktop running and nobody re-ran it.
New-Measure $fact "Worked Minutes" 'CALCULATE ( SUM ( Fact_ServiceJobs[JobDurationMin] ) + SUM ( Fact_ServiceJobs[TravelTimeMin] ), Fact_ServiceJobs[JobStatus] = "Completed", Dim_Date[IsWeekend] = 0 )' "#,##0" "Utilization"
New-Measure $fact "Utilization %" "DIVIDE ( [Worked Minutes], [Available Minutes] )" $pct "Utilization"

# Weekend work does not vanish -- it is overtime, and it is reported rather than
# folded into utilization. vw_TechnicianUtilization makes the same split
# (WorkedMinutes weekday, OvertimeMinutes weekend, TotalWorkedMinutes their
# sum), and the model carried no equivalent at all until now, so the figure the
# fix surfaced had nowhere to be shown.
New-Measure $fact "Overtime Minutes" 'CALCULATE ( SUM ( Fact_ServiceJobs[JobDurationMin] ) + SUM ( Fact_ServiceJobs[TravelTimeMin] ), Fact_ServiceJobs[JobStatus] = "Completed", Dim_Date[IsWeekend] = 1 )' "#,##0" "Utilization"
New-Measure $fact "Total Worked Minutes" "[Worked Minutes] + [Overtime Minutes]" "#,##0" "Utilization"
New-Measure $fact "Overtime % of Capacity" "DIVIDE ( [Overtime Minutes], [Available Minutes] )" $pct "Utilization"

New-Measure $fact "Total Revenue" 'CALCULATE ( SUM ( Fact_ServiceJobs[JobRevenue] ), Fact_ServiceJobs[JobStatus] = "Completed" )' $money "Financial"
New-Measure $fact "Total Cost" 'CALCULATE ( SUM ( Fact_ServiceJobs[JobCost] ), Fact_ServiceJobs[JobStatus] <> "Rescheduled" )' $money "Financial"
New-Measure $fact "Gross Margin" "[Total Revenue] - [Total Cost]" $money "Financial"
New-Measure $fact "Revenue per Completed Job" "DIVIDE ( [Total Revenue], [Completed Jobs] )" $money "Financial"

New-Measure $fact "Util Warning Threshold" 'CALCULATE ( VALUES ( Ref_KPITargets[WarningValue] ), Ref_KPITargets[MetricName] = "UtilizationPct" )' "0.0" "Targets"
New-Measure $fact "FTF Warning Threshold" 'CALCULATE ( VALUES ( Ref_KPITargets[WarningValue] ), Ref_KPITargets[MetricName] = "FirstTimeFixPct" )' "0.0" "Targets"
New-Measure $fact "SLA Warning Threshold" 'CALCULATE ( VALUES ( Ref_KPITargets[WarningValue] ), Ref_KPITargets[MetricName] = "SLACompliancePct" )' "0.0" "Targets"
New-Measure $fact "Callback Warning Threshold" 'CALCULATE ( VALUES ( Ref_KPITargets[WarningValue] ), Ref_KPITargets[MetricName] = "CallbackPct" )' "0.0" "Targets"
New-Measure $fact "Util Ceiling Threshold" 'CALCULATE ( VALUES ( Ref_KPITargets[WarningValue] ), Ref_KPITargets[MetricName] = "UtilizationCeilingPct" )' "0.0" "Targets"

# NOTE on the blank guards below: a technician with no capacity in the current
# context (Inactive, or a month before their HireDate) yields BLANK utilization,
# and in DAX BLANK compares as 0 -- so every "below target" test fired and those
# technicians scored a false Risk Score of 3, polluting the Priority Queue page.
# Guard on capacity, and only flag a quality metric that actually has a value.
$utilStatusExpr = @'
VAR Avail = [Available Minutes]
VAR U = [Utilization %] * 100
RETURN
    IF (
        ISBLANK ( Avail ) || Avail = 0,
        BLANK (),
        SWITCH (
            TRUE(),
            U < [Util Warning Threshold], "Underutilized",
            U > [Util Ceiling Threshold], "Over Capacity",
            "On Target"
        )
    )
'@
New-Measure $fact "Utilization Status" $utilStatusExpr $null "Status"

$riskExpr = @'
VAR Avail   = [Available Minutes]
VAR UtilPct = [Utilization %] * 100
VAR FTFPct  = [First-Time-Fix %] * 100
VAR SLAPct  = [SLA Compliance %] * 100
VAR CBPct   = [Callback %] * 100
VAR Score =
    IF ( UtilPct < [Util Warning Threshold], 1, 0 )
        + IF ( UtilPct > [Util Ceiling Threshold], 1, 0 )
        + IF ( NOT ISBLANK ( FTFPct ) && FTFPct < [FTF Warning Threshold], 1, 0 )
        + IF ( NOT ISBLANK ( SLAPct ) && SLAPct < [SLA Warning Threshold], 1, 0 )
        + IF ( NOT ISBLANK ( CBPct ) && CBPct > [Callback Warning Threshold], 1, 0 )
RETURN
    IF ( ISBLANK ( Avail ) || Avail = 0, BLANK (), Score )
'@
New-Measure $fact "Risk Score" $riskExpr "0" "Status"

# Text status labels for FTF / SLA / Callback (companions to Utilization Status)
New-Measure $fact "FTF Status" @'
VAR V = [First-Time-Fix %] * 100
RETURN IF ( ISBLANK ( V ), BLANK (), IF ( V < [FTF Warning Threshold], "Below Threshold", "Within Threshold" ) )
'@ $null "Status"
New-Measure $fact "SLA Status" @'
VAR V = [SLA Compliance %] * 100
RETURN IF ( ISBLANK ( V ), BLANK (), IF ( V < [SLA Warning Threshold], "Below Threshold", "Within Threshold" ) )
'@ $null "Status"
New-Measure $fact "Callback Status" @'
VAR V = [Callback %] * 100
RETURN IF ( ISBLANK ( V ), BLANK (), IF ( V > [Callback Warning Threshold], "Above Threshold", "Within Threshold" ) )
'@ $null "Status"

# Colour measures for conditional formatting ("Format style: Field value").
# Status labels can't drive formatting -- Power BI needs a colour string. These
# return a hex colour only on a breach and BLANK otherwise, so only exceptions
# are highlighted. Each rule is the exact condition that adds a point to
# [Risk Score], so a highlighted cell always explains part of the score.
New-Measure $fact "Utilization Color" @'
VAR Avail = [Available Minutes]
VAR U = [Utilization %] * 100
RETURN
    IF (
        ISBLANK ( Avail ) || Avail = 0,
        BLANK (),
        SWITCH (
            TRUE (),
            U < [Util Warning Threshold], "#D03B3B",
            U > [Util Ceiling Threshold], "#EC835A",
            BLANK ()
        )
    )
'@ $null "Formatting"
New-Measure $fact "FTF Color" @'
VAR V = [First-Time-Fix %] * 100
RETURN IF ( NOT ISBLANK ( V ) && V < [FTF Warning Threshold], "#D03B3B" )
'@ $null "Formatting"
New-Measure $fact "SLA Color" @'
VAR V = [SLA Compliance %] * 100
RETURN IF ( NOT ISBLANK ( V ) && V < [SLA Warning Threshold], "#D03B3B" )
'@ $null "Formatting"
New-Measure $fact "Callback Color" @'
VAR V = [Callback %] * 100
RETURN IF ( NOT ISBLANK ( V ) && V > [Callback Warning Threshold], "#D03B3B" )
'@ $null "Formatting"

Write-Host "Saving measures..."
$model.SaveChanges() | Out-Null

Write-Host ""
Write-Host "=== BUILD COMPLETE ==="
Write-Host "7 tables, 6 relationships (5 active), 31 measures. In Power BI Desktop: expand"
Write-Host "Fact_ServiceJobs in the Fields pane to find measures grouped into"
Write-Host "Volume/Quality/Utilization/Financial/Targets/Status/Formatting folders."
Write-Host "Now: File > Save As to save the .pbix, then follow"
Write-Host "POWER_BI_BUILD_GUIDE.md Section 5 to build the report pages."

$server.Disconnect()
