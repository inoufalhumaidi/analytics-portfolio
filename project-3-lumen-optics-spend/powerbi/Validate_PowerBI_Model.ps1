<#
=============================================================================
Project 3 -- Lumen Optics Manufacturing: Photonics Spend Scorecard
Script:  Validate_PowerBI_Model.ps1
Purpose: Query the live Power BI semantic model with DAX and reconcile every
         headline figure against the SQL that is supposed to produce it.

WHY A SEPARATE SCRIPT
    Build_PowerBI_Model.ps1 can report complete success against a model that
    loaded zero rows. It creates tables, sets relationships and adds measures;
    none of that touches the data. A Desktop refresh does not even populate
    EstimatedRowCount, so printing row counts there prints blanks.

    The only check worth trusting is the one the consuming application would
    make: ask the model, in DAX, for the numbers a report page would show, and
    compare them character by character against SQL. Two independent
    implementations of the same definition, reconciled -- the same standard the
    Excel workbook is held to.

    Note what this can and cannot prove. It proves Power BI and SQL agree. It
    does not prove either is right: Project 1 in this portfolio had an Excel
    workbook and a SQL view that agreed exactly with each other and were both
    wrong, because both were derived from the same flawed written definition.
    Agreement is necessary, not sufficient.

USAGE
    Open the model in Power BI Desktop (or leave the Build script's session
    open), then:  powershell powerbi/Validate_PowerBI_Model.ps1

DATA DISCLOSURE: Lumen Optics Manufacturing is fictional; all data is
synthetic. No confidential data and no production system is involved.
=============================================================================
#>

$ErrorActionPreference = "Stop"

$SqlServerInst = "localhost\TEW_SQLEXPRESS"
$SqlDb         = "LumenSpend"
$AsOf          = "2025-12-31"

# --- ADOMD client (ships with Power BI Desktop / SSMS) -----------------------
$adomdCandidates = @(
    "C:\Program Files\Microsoft.NET\ADOMD.NET\160\Microsoft.AnalysisServices.AdomdClient.dll",
    "C:\Program Files\Microsoft.NET\ADOMD.NET\150\Microsoft.AnalysisServices.AdomdClient.dll",
    "C:\Program Files\Microsoft.NET\ADOMD.NET\140\Microsoft.AnalysisServices.AdomdClient.dll"
)
$adomd = $adomdCandidates | Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $adomd) {
    $adomd = Get-ChildItem -Path "C:\Windows\Microsoft.NET\assembly\GAC_MSIL" -Filter "Microsoft.AnalysisServices.AdomdClient.dll" -Recurse -ErrorAction SilentlyContinue |
             Select-Object -First 1 -ExpandProperty FullName
}
if (-not $adomd) { throw "ADOMD.NET client not found. It installs with Power BI Desktop or SSMS." }
Add-Type -Path $adomd

# --- find the running Desktop engine ----------------------------------------
$msmdsrv = Get-Process -Name "msmdsrv" -ErrorAction SilentlyContinue
if (-not $msmdsrv) { throw "No Power BI Desktop session found. Open the model in Desktop first." }
$conn = Get-NetTCPConnection -OwningProcess $msmdsrv.Id -State Listen -ErrorAction SilentlyContinue | Select-Object -First 1
if (-not $conn) { throw "Could not find the local Analysis Services engine's listening port." }
$port = $conn.LocalPort

# Power BI Desktop names its workspace database with a GUID, and ADOMD needs it
# as the catalog -- unlike AMO, AdomdConnection.Databases is not populated
# before a catalog is chosen, so it cannot be used to discover the name.
# AMO can, and its DLLs are already present from the build script.
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
Write-Host ("-" * 92)

function Get-Dax([string]$expr) {
    $cmd = $cnDax.CreateCommand()
    $cmd.CommandText = "EVALUATE ROW(""v"", $expr)"
    $rdr = $cmd.ExecuteReader()
    try   { if ($rdr.Read()) { $v = $rdr[0]; if ($v -is [System.DBNull]) { return $null }; return [double]$v } }
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
# the number of decimals SQL publishes that figure to. Comparing to more
# decimals than SQL stores would fail on SQL's own rounding, not on a defect.
$checks = @(
    @{ N="Row count -- POLine";            D="COUNTROWS ( POLine )";                 S="SELECT COUNT(*) FROM dbo.fn_POLineCost('$AsOf')";                                  P=0 },
    @{ N="Row count -- PriceErosion";      D="COUNTROWS ( PriceErosion )";           S="SELECT COUNT(*) FROM dbo.vw_PriceErosion";                                         P=0 },
    @{ N="Row count -- RenegotiationQueue";D="COUNTROWS ( RenegotiationQueue )";     S="SELECT COUNT(*) FROM dbo.vw_RenegotiationQueue";                                    P=0 },
    @{ N="Row count -- Dim_Vendor";        D="COUNTROWS ( Dim_Vendor )";             S="SELECT COUNT(*) FROM dbo.Dim_Vendor";                                              P=0 },
    @{ N="Row count -- Dim_Part";          D="COUNTROWS ( Dim_Part )";               S="SELECT COUNT(*) FROM dbo.Dim_Part";                                                P=0 },
    @{ N="Row count -- Dim_Buyer";         D="COUNTROWS ( Dim_Buyer )";              S="SELECT COUNT(*) FROM dbo.Dim_Buyer";                                               P=0 },
    @{ N="Row count -- Dim_Date";          D="COUNTROWS ( Dim_Date )";               S="SELECT COUNT(*) FROM dbo.Dim_Date WHERE [Date] BETWEEN '2023-01-01' AND '2026-06-30'"; P=0 },

    @{ N="Extended price -- all history";  D="[Extended Price]";                     S="SELECT SUM(ExtendedPrice) FROM dbo.fn_POLineCost('$AsOf')";                        P=2 },
    @{ N="Landed cost -- all history";     D="[Landed Cost]";                        S="SELECT SUM(LandedCost) FROM dbo.fn_POLineCost('$AsOf')";                           P=2 },
    @{ N="Qty accepted -- all history";    D="[Qty Accepted]";                       S="SELECT SUM(ISNULL(QtyAccepted,0)) FROM dbo.fn_POLineCost('$AsOf')";                P=0 },
    @{ N="Rejected value -- all history";  D="[Rejected Value]";                     S="SELECT SUM(RejectedValue) FROM dbo.fn_POLineCost('$AsOf')";                        P=2 },

    @{ N="TTM spend";                      D="[Spend (TTM)]";                        S="SELECT TotalSpend  FROM dbo.fn_SpendKPI('$AsOf')";                                 P=2 },
    @{ N="TTM landed cost";                D="[Landed Cost (TTM)]";                  S="SELECT LandedSpend FROM dbo.fn_SpendKPI('$AsOf')";                                 P=2 },
    @{ N="TTM PO lines";                   D="[PO Lines (TTM)]";                     S="SELECT Lines       FROM dbo.fn_SpendKPI('$AsOf')";                                 P=0 },
    @{ N="TTM maverick spend %";           D="[Maverick Spend % (TTM)]";             S="SELECT MaverickSpendPct     FROM dbo.fn_SpendKPI('$AsOf')";                        P=2 },
    @{ N="TTM purchase price variance %";  D="[PPV % (TTM)]";                        S="SELECT PPVPct               FROM dbo.fn_SpendKPI('$AsOf')";                        P=3 },
    @{ N="TTM acceptance rate %";          D="[Acceptance Rate % (TTM)]";            S="SELECT AcceptanceRatePct    FROM dbo.fn_SpendKPI('$AsOf')";                        P=2 },
    @{ N="TTM on-time delivery %";         D="[On Time Delivery % (TTM)]";           S="SELECT OnTimeDeliveryPct    FROM dbo.fn_SpendKPI('$AsOf')";                        P=2 },
    @{ N="TTM expedite spend %";           D="[Expedite Spend % (TTM)]";             S="SELECT ExpediteSpendPct     FROM dbo.fn_SpendKPI('$AsOf')";                        P=2 },
    @{ N="TTM single-source spend %";      D="[Single Source Spend % (TTM)]";        S="SELECT SingleSourceSpendPct FROM dbo.fn_SpendKPI('$AsOf')";                        P=2 },
    @{ N="TTM top-5 vendor share %";       D="[Top 5 Vendor Share % (TTM)]";         S="SELECT Top5VendorSharePct   FROM dbo.fn_SpendKPI('$AsOf')";                        P=2 },

    @{ N="Erosion capture % (spend-wtd)";  D="[Erosion Capture %]";                  S="SELECT ErosionCapturePct  FROM dbo.fn_SpendKPI('$AsOf')";                          P=2 },
    @{ N="Annual erosion opportunity";     D="[Annual Opportunity]";                 S="SELECT ErosionOpportunity FROM dbo.fn_SpendKPI('$AsOf')";                          P=2 },
    @{ N="Pairs assessed";                 D="[Pairs Assessed]";                     S="SELECT COUNT(*) FROM dbo.vw_PriceErosion";                                         P=0 },
    @{ N="Pairs capturing under 20%";      D="[Pairs Capturing Under 20%]";          S="SELECT COUNT(*) FROM dbo.vw_PriceErosion WHERE ErosionCapturePct < 20";            P=0 },

    @{ N="Queue -- total opportunity";     D="[Queue Opportunity]";                  S="SELECT SUM(AnnualOpportunity) FROM dbo.vw_RenegotiationQueue";                     P=2 },
    @{ N="Queue -- this quarter pairs";    D="[This Quarter Pairs]";                 S="SELECT SUM(CAST(IsThisQuarter AS INT)) FROM dbo.vw_RenegotiationQueue";            P=0 },
    @{ N="Queue -- this quarter value";    D="[This Quarter Opportunity]";           S="SELECT SUM(CASE WHEN IsThisQuarter=1 THEN AnnualOpportunity ELSE 0 END) FROM dbo.vw_RenegotiationQueue"; P=2 },

    @{ N="Reporting anchor -- TTM start";  D="VALUE ( FORMAT ( [TTM Start],  ""yyyyMMdd"" ) )"; S="SELECT CONVERT(INT, FORMAT(DATEADD(DAY,1,DATEADD(MONTH,-12,CAST('$AsOf' AS DATE))),'yyyyMMdd'))"; P=0 },
    @{ N="Reporting anchor -- as of";      D="VALUE ( FORMAT ( [As Of Date], ""yyyyMMdd"" ) )"; S="SELECT CONVERT(INT, FORMAT(CAST('$AsOf' AS DATE),'yyyyMMdd'))";          P=0 }
)

$pass = 0; $fail = 0
$failures = @()
foreach ($c in $checks) {
    $dax = $null; $sql = $null; $err = $null
    try { $dax = Get-Dax $c.D } catch { $err = "DAX: $($_.Exception.Message)" }
    if (-not $err) { try { $sql = Get-Sql $c.S } catch { $err = "SQL: $($_.Exception.Message)" } }

    if ($err) {
        $fail++; $failures += "$($c.N) -- $err"
        Write-Host ("  {0,-34} {1}" -f $c.N, "ERROR: $err") -ForegroundColor Red
        continue
    }

    $tol = [math]::Pow(10, -[int]$c.P) / 2
    $d = [math]::Round($dax, [int]$c.P)
    $s = [math]::Round($sql, [int]$c.P)
    $ok = [math]::Abs($d - $s) -le $tol

    $fmt = if ($c.P -eq 0) { "N0" } else { "N$($c.P)" }
    if ($ok) {
        $pass++
        Write-Host ("  {0,-34} {1,18}  ==  {2,-18} MATCH" -f $c.N, $d.ToString($fmt), $s.ToString($fmt)) -ForegroundColor Green
    } else {
        $fail++
        $failures += ("{0}: Power BI {1} vs SQL {2}" -f $c.N, $d.ToString($fmt), $s.ToString($fmt))
        Write-Host ("  {0,-34} {1,18}  <>  {2,-18} MISMATCH" -f $c.N, $d.ToString($fmt), $s.ToString($fmt)) -ForegroundColor Red
    }
}

# --- a like-for-like check at a grain the totals cannot reach ----------------
# Totals can agree while every per-vendor figure is wrong, so reconcile the cost
# index for EVERY vendor, not a sample. Two separate defects in this one measure
# were each correct for some vendors and wrong for the rest:
#   * the SQL version summed pair-level landed cost once per PO line, a fan-out
#     wrong for all 36 vendors by up to 3.07 points, which still left every
#     value above 100 and entirely plausible;
#   * the first DAX version returned exactly 100.00 for a filtered vendor,
#     because the inner MINX intersected with the outer vendor filter and so
#     compared each vendor against itself.
# A single-vendor spot check passes on the second of those roughly one time in
# thirty-six -- whenever the vendor sampled happens to be the cheapest.
$vendors = @()
$cn = New-Object System.Data.SqlClient.SqlConnection("Server=$SqlServerInst;Database=$SqlDb;Integrated Security=True;")
$cn.Open()
try {
    $cmd = $cn.CreateCommand()
    $cmd.CommandText = "SELECT v.VendorID, s.CostIndexVsBest FROM dbo.fn_VendorScorecard('$AsOf') s JOIN dbo.Dim_Vendor v ON v.VendorKey = s.VendorKey ORDER BY s.CostIndexVsBest DESC;"
    $rdr = $cmd.ExecuteReader()
    while ($rdr.Read()) { $vendors += @{ Id = [string]$rdr[0]; Idx = [double]$rdr[1] } }
    $rdr.Close()
} finally { $cn.Close() }

$idxPass = 0; $idxFail = 0; $worstDelta = 0.0
foreach ($v in $vendors) {
    # No duplicate filter applied here on purpose: the measure applies it
    # itself, and applying it twice would hide a measure that had stopped.
    $daxIdx = Get-Dax "CALCULATE ( [Cost Index vs Best], Dim_Vendor[VendorID] = ""$($v.Id)"" )"
    if ($null -eq $daxIdx) { $idxFail++; $failures += "Cost index $($v.Id): Power BI returned blank"; continue }
    $d = [math]::Round($daxIdx, 2); $s = [math]::Round($v.Idx, 2)
    $delta = [math]::Abs($d - $s)
    if ($delta -gt $worstDelta) { $worstDelta = $delta }
    if ($delta -le 0.01) { $idxPass++ }
    else { $idxFail++; $failures += ("Cost index {0}: Power BI {1} vs SQL {2}" -f $v.Id, $d, $s) }
}
if ($idxFail -eq 0) {
    $pass++
    Write-Host ("  {0,-34} {1,18}  ==  {2,-18} MATCH" -f "Cost index -- all $($vendors.Count) vendors", "max delta", $worstDelta.ToString("N4")) -ForegroundColor Green
} else {
    $fail++
    Write-Host ("  {0,-34} {1,18}  <>  {2,-18} MISMATCH" -f "Cost index -- all $($vendors.Count) vendors", "$idxFail wrong", "$idxPass right") -ForegroundColor Red
}

$cnDax.Close()
Write-Host ("-" * 92)
if ($fail -eq 0) {
    Write-Host "VALIDATION PASS: $pass of $pass checks reconcile between Power BI and SQL." -ForegroundColor Green
} else {
    Write-Host "VALIDATION FAIL: $pass passed, $fail failed." -ForegroundColor Red
    $failures | ForEach-Object { Write-Host "    - $_" -ForegroundColor Red }
    exit 1
}
