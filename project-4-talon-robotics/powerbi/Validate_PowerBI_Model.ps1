<#
=============================================================================
Project 4 -- Talon Robotics: Payload Deployment Delivery
Script:  Validate_PowerBI_Model.ps1
Purpose: Query the live Power BI semantic model with DAX and reconcile every
         headline figure against the SQL that is supposed to produce it.

WHY A SEPARATE SCRIPT
    Build_PowerBI_Model.ps1 reports complete success against a model that
    loaded zero rows. Creating tables, relationships and measures never touches
    the data, and a Desktop refresh does not even populate EstimatedRowCount.
    The only check worth trusting is the one the consuming application makes:
    ask the model, in DAX, for the numbers a report page would show.

    It proves Power BI and SQL AGREE. It does not prove either is right --
    Project 1 shipped an Excel workbook and a SQL view that agreed exactly and
    were both wrong, because both came from the same flawed written definition.
    Agreement is necessary, not sufficient.

    Two checks below go further than the totals on purpose. The verification
    states must partition the population, and the two stale buckets must be
    mutually exclusive: a total can agree while the split beneath it is wrong,
    and the split is what the report pages actually show.

DATA DISCLOSURE: Talon Robotics is fictional; all data is synthetic. No
confidential data and no production system is involved.
=============================================================================
#>

$ErrorActionPreference = "Stop"

$SqlServerInst = "localhost\TEW_SQLEXPRESS"
$SqlDb         = "TalonDelivery"

# --- ADOMD client (ships with Power BI Desktop / SSMS) -----------------------
$adomd = @(
    "C:\Program Files\Microsoft.NET\ADOMD.NET\160\Microsoft.AnalysisServices.AdomdClient.dll",
    "C:\Program Files\Microsoft.NET\ADOMD.NET\150\Microsoft.AnalysisServices.AdomdClient.dll"
) | Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $adomd) {
    $adomd = Get-ChildItem -Path "C:\Windows\Microsoft.NET\assembly\GAC_MSIL" `
             -Filter "Microsoft.AnalysisServices.AdomdClient.dll" -Recurse -ErrorAction SilentlyContinue |
             Select-Object -First 1 -ExpandProperty FullName
}
if (-not $adomd) { throw "ADOMD.NET client not found. It installs with Power BI Desktop or SSMS." }
Add-Type -Path $adomd

$msmdsrv = Get-Process -Name "msmdsrv" -ErrorAction SilentlyContinue
if (-not $msmdsrv) { throw "No Power BI Desktop session found. Open the model in Desktop first." }
$conn = Get-NetTCPConnection -OwningProcess $msmdsrv.Id -State Listen -ErrorAction SilentlyContinue | Select-Object -First 1
if (-not $conn) { throw "Could not find the local Analysis Services engine's listening port." }
$port = $conn.LocalPort

# AdomdConnection.Databases is empty until a catalog is chosen, so it cannot be
# used to discover the database name. AMO can, and its DLLs are already present
# from the build script.
$dllDir = Join-Path $PSScriptRoot ".amo_libs\lib\net45"
if (-not (Test-Path (Join-Path $dllDir "Microsoft.AnalysisServices.Tabular.dll"))) {
    throw "AMO libraries not found. Run Build_PowerBI_Model.ps1 first -- it downloads them."
}
Add-Type -Path (Join-Path $dllDir "Microsoft.AnalysisServices.Core.dll")
Add-Type -Path (Join-Path $dllDir "Microsoft.AnalysisServices.dll")
Add-Type -Path (Join-Path $dllDir "Microsoft.AnalysisServices.Tabular.dll")
$amo = New-Object Microsoft.AnalysisServices.Tabular.Server
$amo.Connect("localhost:$port")
if ($amo.Databases.Count -eq 0) { throw "The Desktop session has no database loaded." }
$dbName = $amo.Databases[0].Name
$amo.Disconnect()

$cnDax = New-Object Microsoft.AnalysisServices.AdomdClient.AdomdConnection("Data Source=localhost:$port;Initial Catalog=$dbName")
$cnDax.Open()
Write-Host "Connected to the Power BI model on port $port (database $dbName)" -ForegroundColor Cyan
Write-Host ("-" * 94)

function Get-Dax([string]$expr) {
    $cmd = $cnDax.CreateCommand()
    $cmd.CommandText = "EVALUATE ROW(""v"", $expr)"
    $rdr = $cmd.ExecuteReader()
    try { if ($rdr.Read()) { $v = $rdr[0]; if ($v -is [System.DBNull]) { return $null }; return [double]$v } }
    finally { $rdr.Close() }
    return $null
}

function Get-Sql([string]$query) {
    $cn = New-Object System.Data.SqlClient.SqlConnection("Server=$SqlServerInst;Database=$SqlDb;Integrated Security=True;")
    $cn.Open()
    try {
        $cmd = $cn.CreateCommand(); $cmd.CommandText = $query; $cmd.CommandTimeout = 300
        $v = $cmd.ExecuteScalar()
        if ($null -eq $v -or $v -is [System.DBNull]) { return $null }
        return [double]$v
    } finally { $cn.Close() }
}

# Each check: label, the DAX the model answers with, the SQL it must match, and
# the decimals SQL publishes that figure to. Comparing to more precision than
# SQL stores would fail on SQL's own rounding, not on a defect.
$checks = @(
    @{ N="Rows -- Requirement";        D="COUNTROWS ( Requirement )";       S="SELECT COUNT(*) FROM dbo.vw_RequirementVerification"; P=0 },
    @{ N="Rows -- Queue";              D="COUNTROWS ( Queue )";             S="SELECT COUNT(*) FROM dbo.vw_VerificationQueue";       P=0 },
    @{ N="Rows -- RAID";               D="COUNTROWS ( RAID )";              S="SELECT COUNT(*) FROM dbo.vw_RAIDExposure";            P=0 },
    @{ N="Rows -- WorkItem";           D="COUNTROWS ( WorkItem )";          S="SELECT COUNT(*) FROM dbo.Fact_WorkItem";              P=0 },
    @{ N="Rows -- Dim_Subsystem";      D="COUNTROWS ( Dim_Subsystem )";     S="SELECT COUNT(*) FROM dbo.Dim_Subsystem";              P=0 },
    @{ N="Rows -- Dim_Build";          D="COUNTROWS ( Dim_Build )";         S="SELECT COUNT(*) FROM dbo.Dim_Build";                  P=0 },
    @{ N="Rows -- BuildChurn";         D="COUNTROWS ( BuildChurn )";        S="SELECT COUNT(*) FROM dbo.Fact_BuildSubsystemChange";  P=0 },

    @{ N="Must-ship requirements";     D="[Must Ship]";                     S="SELECT MustShipCount FROM dbo.vw_ReadinessKPI";       P=0 },
    @{ N="Ship readiness %";           D="[Ship Readiness %]";              S="SELECT ShipReadinessPct FROM dbo.vw_ReadinessKPI";    P=2 },
    @{ N="Verification coverage %";    D="[Verification Coverage %]";       S="SELECT VerificationCoveragePct FROM dbo.vw_ReadinessKPI"; P=2 },
    @{ N="Stale verification %";       D="[Stale Verification %]";          S="SELECT StaleVerificationPct FROM dbo.vw_ReadinessKPI"; P=2 },
    @{ N="Policy compliance %";        D="[Policy Compliance %]";           S="SELECT PolicyCompliancePct FROM dbo.vw_ReadinessKPI"; P=2 },
    @{ N="Work-item completion %";     D="[Work Item Completion %]";        S="SELECT WorkItemCompletionPct FROM dbo.vw_ReadinessKPI"; P=2 },

    @{ N="No evidence (count)";        D="[No Evidence]";                   S="SELECT NoEvidenceCount FROM dbo.vw_ReadinessKPI";     P=0 },
    @{ N="Insufficient (count)";       D="[Insufficient]";                  S="SELECT InsufficientCount FROM dbo.vw_ReadinessKPI";   P=0 },
    @{ N="Stale (count)";              D="[Stale]";                         S="SELECT StaleCount FROM dbo.vw_ReadinessKPI";          P=0 },
    @{ N="Current (count)";            D="[Current]";                       S="SELECT CurrentCount FROM dbo.vw_ReadinessKPI";        P=0 },

    @{ N="Queue outstanding";          D="[Outstanding]";                   S="SELECT COUNT(*) FROM dbo.vw_VerificationQueue";       P=0 },
    @{ N="Rig hours outstanding";      D="[Rig Hours Outstanding]";         S="SELECT RigHoursOutstanding FROM dbo.vw_ReadinessKPI"; P=1 },
    @{ N="Schedulable this week";      D="[Schedulable This Week]";         S="SELECT SUM(CAST(IsThisWeek AS INT)) FROM dbo.vw_VerificationQueue"; P=0 },

    @{ N="RAID open";                  D="[RAID Open]";                     S="SELECT COUNT(*) FROM dbo.vw_RAIDExposure WHERE IsOpen=1"; P=0 },
    @{ N="Critical RAID open";         D="[Critical RAID Open]";            S="SELECT CriticalRAIDOpen FROM dbo.vw_ReadinessKPI";    P=0 },
    @{ N="Overdue RAID %";             D="[Overdue RAID %]";                S="SELECT OverdueRAIDPct FROM dbo.vw_ReadinessKPI";      P=2 },
    @{ N="Open RAID exposure";         D="[Open RAID Exposure]";            S="SELECT TotalRAIDExposure FROM dbo.vw_ReadinessKPI";   P=0 },

    @{ N="Reporting anchor -- build";  D="[As Of Build]";                   S="SELECT AsOfBuild FROM dbo.vw_ReadinessKPI";           P=0 }
)

$pass = 0; $fail = 0; $failures = @()
foreach ($c in $checks) {
    $dax = $null; $sql = $null; $err = $null
    try { $dax = Get-Dax $c.D } catch { $err = "DAX: $($_.Exception.Message)" }
    if (-not $err) { try { $sql = Get-Sql $c.S } catch { $err = "SQL: $($_.Exception.Message)" } }
    if ($err) {
        $fail++; $failures += "$($c.N) -- $err"
        Write-Host ("  {0,-30} {1}" -f $c.N, "ERROR: $err") -ForegroundColor Red
        continue
    }
    $tol = [math]::Pow(10, -[int]$c.P) / 2
    $d = [math]::Round($dax, [int]$c.P); $s = [math]::Round($sql, [int]$c.P)
    $ok = [math]::Abs($d - $s) -le $tol
    $fmt = if ($c.P -eq 0) { "N0" } else { "N$($c.P)" }
    if ($ok) {
        $pass++
        Write-Host ("  {0,-30} {1,16}  ==  {2,-16} MATCH" -f $c.N, $d.ToString($fmt), $s.ToString($fmt)) -ForegroundColor Green
    } else {
        $fail++; $failures += ("{0}: Power BI {1} vs SQL {2}" -f $c.N, $d.ToString($fmt), $s.ToString($fmt))
        Write-Host ("  {0,-30} {1,16}  <>  {2,-16} MISMATCH" -f $c.N, $d.ToString($fmt), $s.ToString($fmt)) -ForegroundColor Red
    }
}

# --- structural checks the totals cannot reach ------------------------------
# A total can agree while the split beneath it is wrong, and the split is what
# the report pages show.
$partition = Get-Dax "[No Evidence] + [Insufficient] + [Stale] + [Current]"
$reqs      = Get-Dax "[Requirements]"
if ($null -ne $partition -and [math]::Abs($partition - $reqs) -le 0.5) {
    $pass++
    Write-Host ("  {0,-30} {1,16}  ==  {2,-16} MATCH" -f "States partition the population", $partition, $reqs) -ForegroundColor Green
} else {
    $fail++; $failures += "Verification states do not partition the population: $partition vs $reqs"
    Write-Host ("  {0,-30} {1,16}  <>  {2,-16} MISMATCH" -f "States partition the population", $partition, $reqs) -ForegroundColor Red
}

$buckets = Get-Dax "[Stale by Code] + [Stale by Requirement]"
$stale   = Get-Dax "[Stale]"
if ($null -ne $buckets -and [math]::Abs($buckets - $stale) -le 0.5) {
    $pass++
    Write-Host ("  {0,-30} {1,16}  ==  {2,-16} MATCH" -f "Stale buckets are exclusive", $buckets, $stale) -ForegroundColor Green
} else {
    $fail++; $failures += "Stale buckets do not sum to Stale: $buckets vs $stale (they overlap or leave a gap)"
    Write-Host ("  {0,-30} {1,16}  <>  {2,-16} MISMATCH" -f "Stale buckets are exclusive", $buckets, $stale) -ForegroundColor Red
}

# Readiness per subsystem, every one -- a headline can agree while every cut
# beneath it is wrong, and the subsystem cut is what decides where people go.
$subs = @()
$cn = New-Object System.Data.SqlClient.SqlConnection("Server=$SqlServerInst;Database=$SqlDb;Integrated Security=True;")
$cn.Open()
try {
    $cmd = $cn.CreateCommand()
    $cmd.CommandText = "SELECT SubsystemCode, ReadinessPct FROM dbo.vw_SubsystemReadiness ORDER BY SubsystemCode;"
    $rdr = $cmd.ExecuteReader()
    while ($rdr.Read()) { $subs += @{ Code = [string]$rdr[0]; Pct = [double]$rdr[1] } }
    $rdr.Close()
} finally { $cn.Close() }

$subFail = 0; $subWorst = 0.0
foreach ($sv in $subs) {
    $d = Get-Dax "CALCULATE ( [Ship Readiness %], Dim_Subsystem[SubsystemCode] = ""$($sv.Code)"" )"
    if ($null -eq $d) { $subFail++; continue }
    $delta = [math]::Abs([math]::Round($d,2) - [math]::Round($sv.Pct,2))
    if ($delta -gt $subWorst) { $subWorst = $delta }
    if ($delta -gt 0.01) { $subFail++; $failures += ("Readiness {0}: Power BI {1:N2} vs SQL {2:N2}" -f $sv.Code, $d, $sv.Pct) }
}
if ($subFail -eq 0) {
    $pass++
    Write-Host ("  {0,-30} {1,16}  ==  {2,-16} MATCH" -f "Readiness -- all $($subs.Count) subsystems", "max delta", $subWorst.ToString("N4")) -ForegroundColor Green
} else {
    $fail++
    Write-Host ("  {0,-30} {1,16}  <>  {2,-16} MISMATCH" -f "Readiness -- all $($subs.Count) subsystems", "$subFail wrong", "") -ForegroundColor Red
}

$cnDax.Close()
Write-Host ("-" * 94)
if ($fail -eq 0) {
    Write-Host "VALIDATION PASS: $pass of $pass checks reconcile between Power BI and SQL." -ForegroundColor Green
} else {
    Write-Host "VALIDATION FAIL: $pass passed, $fail failed." -ForegroundColor Red
    $failures | ForEach-Object { Write-Host "    - $_" -ForegroundColor Red }
    exit 1
}
