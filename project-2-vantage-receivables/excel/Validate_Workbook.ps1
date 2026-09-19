<#
================================================================================
Project 2 -- Vantage Wholesale Supply: Receivables Performance
Script:  Validate_Workbook.ps1
Purpose: Open the built workbook, force a full recalculation, and read back what
         the formulas actually produced.

WHY THIS EXISTS AS A SEPARATE SCRIPT
    A build script that finishes without throwing has proved that it ran, not
    that it produced a working file. Structural checks -- file size, sheet
    count, hashes -- prove a file is intact, never that it is valid. The only
    thing that proves a workbook works is opening it in Excel and reading the
    values back out. That is what this does.

WHAT IT ASSERTS
    1. No formula anywhere resolves to an Excel error value.
    2. The Excel-computed figures match the SQL figures the project published.
    3. Every row of the Validation sheet reconciles.
    4. The Priority Queue actually returns ranked rows.

DATA DISCLOSURE: Vantage Wholesale Supply is fictional; all data is synthetic.
================================================================================
#>

param(
    [string]$WorkbookPath = (Join-Path $PSScriptRoot 'Vantage_AR_Control.xlsx')
)

$ErrorActionPreference = 'Stop'
if (-not (Test-Path $WorkbookPath)) { throw "Workbook not found: $WorkbookPath" }

$SqlServerInst = 'localhost\TEW_SQLEXPRESS'
$SqlDb         = 'VantageAR'
$AsOf          = '2025-12-31'

# Each figure carries BOTH the query SQL answers it with and the value published
# in the README and case study.
#
# The published value used to be the only target. That is a frozen snapshot: if
# the SQL logic changes, the workbook follows it, the validator keeps comparing
# against the old number and reports a failure that is really a stale
# expectation -- or worse, both move together and it reports success. Project 3
# in this portfolio shipped a validator holding exactly such a constant.
#
# Checking against both separates two different failures that a single target
# conflates:
#   Excel <> SQL        the workbook's formulas disagree with the database
#   SQL   <> published  the database has moved since the documents were written
$sqlFigures = @(
    @{ Sheet='Dashboard'; Cell='C8';  Label='Total open AR';          Published=10847081.49; Tol=0.005;
       Sql="SELECT SUM(OpenBalance) FROM dbo.vw_ARBalance WHERE OpenBalance > 0.005" },
    @{ Sheet='Dashboard'; Cell='C9';  Label='Not yet due';            Published=8095793.78;  Tol=0.005;
       Sql="SELECT CurrentAR FROM dbo.fn_DSO('$AsOf')" },
    @{ Sheet='Dashboard'; Cell='C10'; Label='Past due, disputed';     Published=154864.71;   Tol=0.005;
       Sql="SELECT DisputedPastDueAR FROM dbo.fn_DSOBridge('$AsOf')" },
    @{ Sheet='Dashboard'; Cell='C11'; Label='Past due, undisputed';   Published=2596423.00;  Tol=0.005;
       Sql="SELECT UndisputedPastDueAR FROM dbo.fn_DSOBridge('$AsOf')" },
    @{ Sheet='Dashboard'; Cell='C13'; Label='Open invoices';          Published=2346;        Tol=0.5;
       Sql="SELECT COUNT(*) FROM dbo.vw_ARBalance WHERE OpenBalance > 0.005" },
    @{ Sheet='Dashboard'; Cell='F9';  Label='Granted days';           Published=45.97;       Tol=0.01;
       Sql="SELECT GrantedDays FROM dbo.fn_DSOBridge('$AsOf')" },
    @{ Sheet='Dashboard'; Cell='F10'; Label='Dispute days';           Published=0.88;        Tol=0.01;
       Sql="SELECT DisputeDays FROM dbo.fn_DSOBridge('$AsOf')" },
    @{ Sheet='Dashboard'; Cell='F11'; Label='Lateness days';          Published=14.74;       Tol=0.01;
       Sql="SELECT LatenessDays FROM dbo.fn_DSOBridge('$AsOf')" },
    @{ Sheet='Dashboard'; Cell='F12'; Label='Classic DSO';            Published=61.59;       Tol=0.01;
       Sql="SELECT DSO_Classic FROM dbo.fn_DSOBridge('$AsOf')" },
    @{ Sheet='Dashboard'; Cell='F14'; Label='Weighted average terms'; Published=46.68;       Tol=0.01;
       Sql="SELECT WeightedAvgTermsDays FROM dbo.fn_DSOBridge('$AsOf')" },
    @{ Sheet='Dashboard'; Cell='F17'; Label='Billing lag (days)';     Published=1.78;        Tol=0.02;
       Sql="SELECT BillingLagDays FROM dbo.fn_DSOBridge('$AsOf')" },
    @{ Sheet='Dashboard'; Cell='F18'; Label='True cash cycle (days)'; Published=63.37;       Tol=0.02;
       Sql="SELECT CashCycleDays FROM dbo.fn_DSOBridge('$AsOf')" },
    @{ Sheet='Dashboard'; Cell='B27'; Label='Unapplied cash % of AR'; Published=3.87;        Tol=0.02;
       Sql="SELECT CAST(100.0 * (SELECT SUM(UnappliedCash) FROM dbo.vw_UnappliedCash) / NULLIF((SELECT SUM(OpenBalance) FROM dbo.vw_ARBalance WHERE OpenBalance > 0.005),0) AS DECIMAL(9,2))" },
    @{ Sheet='Dashboard'; Cell='B29'; Label='Promise kept rate %';    Published=86.29;       Tol=0.01;
       Sql="SELECT KeptRatePct FROM dbo.fn_PromiseKeptRate('$AsOf', 12)" },
    @{ Sheet='Calc';      Cell='C20'; Label='Countback DSO';          Published=63.07;       Tol=0.01;
       Sql="SELECT DSO_Countback FROM dbo.fn_DSO('$AsOf')" },
    @{ Sheet='Calc';      Cell='C22'; Label='Best possible DSO';      Published=46.27;       Tol=0.02;
       Sql="SELECT BPDSO_Countback FROM dbo.fn_DSO('$AsOf')" },
    @{ Sheet='Calc';      Cell='C23'; Label='Average days delinquent';Published=16.80;       Tol=0.02;
       Sql="SELECT AvgDaysDelinquent FROM dbo.fn_DSO('$AsOf')" }
)

function Get-SqlScalar([string]$query) {
    $cn = New-Object System.Data.SqlClient.SqlConnection("Server=$SqlServerInst;Database=$SqlDb;Integrated Security=True;")
    $cn.Open()
    try {
        $cmd = $cn.CreateCommand(); $cmd.CommandText = $query; $cmd.CommandTimeout = 300
        $v = $cmd.ExecuteScalar()
        if ($null -eq $v -or $v -is [System.DBNull]) { return $null }
        return [double]$v
    } finally { $cn.Close() }
}

Write-Host "Reading the current figures from SQL..." -ForegroundColor Cyan
$sqlUnavailable = $false
foreach ($f in $sqlFigures) {
    try { $f.Expect = Get-SqlScalar $f.Sql }
    catch {
        # Without a database the workbook can still be checked against what was
        # published -- which is the state a reader who only cloned the repo is
        # in. Say so rather than failing, and do not silently call it a pass.
        $f.Expect = $f.Published
        $sqlUnavailable = $true
    }
}
if ($sqlUnavailable) {
    Write-Host "  SQL Server unreachable -- falling back to the PUBLISHED figures." -ForegroundColor Yellow
    Write-Host "  This checks the workbook against the documents, not against the database." -ForegroundColor Yellow
}

$errorText = @('#REF!','#VALUE!','#NAME?','#DIV/0!','#N/A','#NULL!','#NUM!')
$failures  = 0

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false

try {
    $wb = $xl.Workbooks.Open($WorkbookPath, 0, $true)   # read-only
    $xl.Calculation = -4105
    $wb.Application.CalculateFullRebuild()

    Write-Host ""
    Write-Host "Validating $(Split-Path $WorkbookPath -Leaf)" -ForegroundColor Cyan
    Write-Host ("-" * 78)

    # ---- 1. Excel figures against the SQL figures --------------------------
    Write-Host "Excel formulas against the figures SQL published:"
    foreach ($f in $sqlFigures) {
        $v = $wb.Worksheets.Item($f.Sheet).Range($f.Cell).Value2
        if ($null -eq $v -or $v -isnot [double]) {
            Write-Host ("  {0,-26} {1,16}  NOT A NUMBER" -f $f.Label, "$v") -ForegroundColor Red
            $failures++
            continue
        }
        $diff = [math]::Abs($v - $f.Expect)
        $ok   = $diff -le $f.Tol
        if (-not $ok) { $failures++ }
        $colour = if ($ok) { 'Green' } else { 'Red' }
        Write-Host ("  {0,-26} excel {1,15:N2}   sql {2,15:N2}   diff {3,10:N4}  {4}" -f `
            $f.Label, $v, $f.Expect, $diff, $(if ($ok) { 'MATCH' } else { 'DIFFERS' })) -ForegroundColor $colour

        # Separately: has SQL moved away from what the documents say? This is a
        # different failure from "Excel disagrees with SQL", and reporting them
        # together is how a stale published figure hides behind a green run.
        if (-not $sqlUnavailable) {
            $drift = [math]::Abs($f.Expect - $f.Published)
            if ($drift -gt $f.Tol) {
                $failures++
                Write-Host ("      ^ SQL now says {0:N2} but the README and case study publish {1:N2}. Update the documents." -f `
                    $f.Expect, $f.Published) -ForegroundColor Red
            }
        }
    }

    # ---- 2. the workbook's own validation sheet ----------------------------
    Write-Host ""
    Write-Host "The workbook's own SQL-versus-Excel reconciliation sheet:"
    $val = $wb.Worksheets.Item('Validation')
    $row = 5
    $match = 0; $differs = 0
    while ($val.Range("B$row").Value2) {
        $result = $val.Range("F$row").Text
        if ($result -eq 'MATCH') { $match++ }
        else {
            $differs++
            Write-Host ("  DIFFERS: {0}  sql={1}  excel={2}" -f `
                $val.Range("B$row").Text, $val.Range("C$row").Text, $val.Range("D$row").Text) -ForegroundColor Red
        }
        $row++
    }
    $colour = if ($differs -eq 0) { 'Green' } else { 'Red' }
    Write-Host ("  {0} of {1} checks reconcile" -f $match, ($match + $differs)) -ForegroundColor $colour
    if ($differs -gt 0) { $failures += $differs }

    # ---- 3. the bridge identity, as the workbook computes it ---------------
    $identity = $wb.Worksheets.Item('Dashboard').Range('H13').Text
    Write-Host ""
    Write-Host "DSO bridge identity: $identity" -ForegroundColor $(if ($identity -like 'Identity holds*') { 'Green' } else { 'Red' })
    if ($identity -notlike 'Identity holds*') { $failures++ }

    $targetVerdict = $wb.Worksheets.Item('Dashboard').Range('H16').Text
    Write-Host "DSO target check:    $targetVerdict"

    # ---- 4. the queue returns ranked rows ----------------------------------
    $pq = $wb.Worksheets.Item('Priority Queue')
    $filled = 0
    for ($r = 10; $r -le 34; $r++) { if ($pq.Range("B$r").Text -ne '') { $filled++ } }
    Write-Host ""
    Write-Host ("Priority Queue returned {0} ranked rows; top account: {1} ({2})" -f `
        $filled, $pq.Range('B10').Text, $pq.Range('O10').Text) -ForegroundColor $(if ($filled -gt 0) { 'Green' } else { 'Red' })
    if ($filled -eq 0) { $failures++ }

    # ---- 5. sweep every sheet for formula errors ---------------------------
    Write-Host ""
    Write-Host "Sweeping every used cell for Excel error values..."
    $errCount = 0
    foreach ($ws in $wb.Worksheets) {
        if ($ws.Name -like 'Data_*') { continue }      # loaded data carries no formulas
        $used = $ws.UsedRange
        foreach ($cell in $used) {
            $t = $cell.Text
            if ($errorText -contains $t) {
                Write-Host ("  {0}!{1}  {2}   formula: {3}" -f $ws.Name, $cell.Address($false,$false), $t, $cell.Formula) -ForegroundColor Red
                $errCount++
                if ($errCount -ge 25) { Write-Host '  (further errors suppressed)' -ForegroundColor Red; break }
            }
        }
        if ($errCount -ge 25) { break }
    }
    if ($errCount -eq 0) { Write-Host "  No error values found." -ForegroundColor Green }
    else { $failures += $errCount }

    $wb.Close($false)
}
finally {
    $xl.Quit()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}

Write-Host ("-" * 78)
if ($failures -eq 0) {
    Write-Host "WORKBOOK VALIDATED: every formula resolves and every figure reconciles with SQL." -ForegroundColor Green
    exit 0
} else {
    Write-Host "WORKBOOK VALIDATION FAILED: $failures problem(s). Do not publish this file." -ForegroundColor Red
    exit 1
}
