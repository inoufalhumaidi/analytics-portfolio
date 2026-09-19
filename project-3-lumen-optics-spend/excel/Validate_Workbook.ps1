<#
================================================================================
Project 3 -- Lumen Optics Manufacturing: Photonics Spend Scorecard
Script:  Validate_Workbook.ps1
Purpose: Open the built workbook, force a full recalculation, and read back what
         the formulas actually produced.

WHY THIS IS A SEPARATE SCRIPT
    A build script that finishes without throwing has proved that it ran, not
    that it produced a working file. Structural checks -- file size, sheet
    count, hashes -- prove a file is intact, never that it is valid. The only
    thing that proves a workbook works is opening it in Excel and reading the
    values back out.

DATA DISCLOSURE: Lumen Optics Manufacturing is fictional; all data is synthetic.
================================================================================
#>

param([string]$WorkbookPath = (Join-Path $PSScriptRoot 'Lumen_Spend_Scorecard.xlsx'))

$ErrorActionPreference = 'Stop'
if (-not (Test-Path $WorkbookPath)) { throw "Workbook not found: $WorkbookPath" }

$SqlServerInst = 'localhost\TEW_SQLEXPRESS'
$SqlDb         = 'LumenSpend'
$AsOf          = '2025-12-31'

# Each figure carries BOTH the query SQL answers it with and the value published
# in the README and case study.
#
# The published value used to be the only target, and this project is the reason
# that is not good enough: a rounding fix changed the erosion figures and this
# file kept asserting the old constant, so it reported a failure that was really
# a stale expectation. The opposite case is worse -- if the workbook and the
# constant drift together, it reports success.
#
# Checking against both separates two different failures:
#   Excel <> SQL        the workbook's formulas disagree with the database
#   SQL   <> published  the database has moved since the documents were written
$sqlFigures = @(
    @{ Sheet='Dashboard'; Cell='C8';  Label='Extended price (12m)';    Published=48586386.86; Tol=0.05;
       Sql="SELECT TotalSpend FROM dbo.fn_SpendKPI('$AsOf')" },
    @{ Sheet='Dashboard'; Cell='C9';  Label='Landed cost (12m)';       Published=49474785.92; Tol=0.05;
       Sql="SELECT LandedSpend FROM dbo.fn_SpendKPI('$AsOf')" },
    @{ Sheet='Dashboard'; Cell='G9';  Label='Erosion capture %';       Published=29.50;       Tol=0.02;
       Sql="SELECT ErosionCapturePct FROM dbo.fn_SpendKPI('$AsOf')" },
    @{ Sheet='Dashboard'; Cell='G10'; Label='Annual opportunity';      Published=4047608.72;  Tol=0.05;
       Sql="SELECT ErosionOpportunity FROM dbo.fn_SpendKPI('$AsOf')" },
    @{ Sheet='Dashboard'; Cell='B20'; Label='Erosion capture (score)'; Published=29.50;       Tol=0.02;
       Sql="SELECT ErosionCapturePct FROM dbo.fn_SpendKPI('$AsOf')" },
    @{ Sheet='Dashboard'; Cell='B21'; Label='Maverick spend %';        Published=27.09;       Tol=0.02;
       Sql="SELECT MaverickSpendPct FROM dbo.fn_SpendKPI('$AsOf')" },
    @{ Sheet='Dashboard'; Cell='B22'; Label='Price variance %';        Published=-0.19;       Tol=0.02;
       Sql="SELECT PPVPct FROM dbo.fn_SpendKPI('$AsOf')" },
    @{ Sheet='Dashboard'; Cell='B23'; Label='Acceptance rate %';       Published=99.11;       Tol=0.02;
       Sql="SELECT AcceptanceRatePct FROM dbo.fn_SpendKPI('$AsOf')" },
    @{ Sheet='Dashboard'; Cell='B24'; Label='On-time delivery %';      Published=80.62;       Tol=0.02;
       Sql="SELECT OnTimeDeliveryPct FROM dbo.fn_SpendKPI('$AsOf')" },
    @{ Sheet='Dashboard'; Cell='B25'; Label='Expedite spend %';        Published=0.23;        Tol=0.02;
       Sql="SELECT ExpediteSpendPct FROM dbo.fn_SpendKPI('$AsOf')" },
    @{ Sheet='Dashboard'; Cell='B26'; Label='Single-source spend %';   Published=27.17;       Tol=0.02;
       Sql="SELECT SingleSourceSpendPct FROM dbo.fn_SpendKPI('$AsOf')" },
    @{ Sheet='Dashboard'; Cell='B27'; Label='Top-5 vendor share %';    Published=27.20;       Tol=0.05;
       Sql="SELECT Top5VendorSharePct FROM dbo.fn_SpendKPI('$AsOf')" }
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
        # published -- the state a reader who only cloned the repo is in. Say
        # so rather than silently calling it a pass.
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
$xl.Visible = $false; $xl.DisplayAlerts = $false

try {
    $wb = $xl.Workbooks.Open($WorkbookPath, 0, $true)
    $xl.Calculation = -4105
    $wb.Application.CalculateFullRebuild()

    Write-Host ""
    Write-Host "Validating $(Split-Path $WorkbookPath -Leaf)" -ForegroundColor Cyan
    Write-Host ("-" * 78)
    Write-Host "Excel formulas against the figures SQL published:"
    foreach ($f in $sqlFigures) {
        $v = $wb.Worksheets.Item($f.Sheet).Range($f.Cell).Value2
        if ($null -eq $v -or $v -isnot [double]) {
            Write-Host ("  {0,-26} {1,16}  NOT A NUMBER" -f $f.Label, "$v") -ForegroundColor Red
            $failures++; continue
        }
        $diff = [math]::Abs($v - $f.Expect)
        $ok = $diff -le $f.Tol
        if (-not $ok) { $failures++ }
        Write-Host ("  {0,-26} excel {1,15:N2}   sql {2,15:N2}   diff {3,9:N4}  {4}" -f `
            $f.Label, $v, $f.Expect, $diff, $(if ($ok) { 'MATCH' } else { 'DIFFERS' })) `
            -ForegroundColor $(if ($ok) { 'Green' } else { 'Red' })

        # Separately: has SQL moved away from what the documents say? A
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

    Write-Host ""
    Write-Host "The workbook's own SQL-versus-Excel reconciliation sheet:"
    $val = $wb.Worksheets.Item('Validation')
    $row = 5; $match = 0; $differs = 0
    while ($val.Range("B$row").Value2) {
        if ($val.Range("F$row").Text -eq 'MATCH') { $match++ }
        else {
            $differs++
            Write-Host ("  DIFFERS: {0}  sql={1}  excel={2}" -f `
                $val.Range("B$row").Text, $val.Range("C$row").Text, $val.Range("D$row").Text) -ForegroundColor Red
        }
        $row++
    }
    Write-Host ("  {0} of {1} checks reconcile" -f $match, ($match + $differs)) `
        -ForegroundColor $(if ($differs -eq 0) { 'Green' } else { 'Red' })
    if ($differs -gt 0) { $failures += $differs }

    $qs = $wb.Worksheets.Item('Renegotiation Queue')
    $filled = 0
    for ($r = 10; $r -le 34; $r++) { if ($qs.Range("B$r").Text -ne '') { $filled++ } }
    Write-Host ""
    Write-Host ("Renegotiation Queue returned {0} ranked rows; top pair: {1} / {2} ({3})" -f `
        $filled, $qs.Range('B10').Text, $qs.Range('D10').Text, $qs.Range('N10').Text) `
        -ForegroundColor $(if ($filled -gt 0) { 'Green' } else { 'Red' })
    if ($filled -eq 0) { $failures++ }

    $vs = $wb.Worksheets.Item('Vendor Scorecard')
    Write-Host ("Vendor Scorecard top row: {0} -- cost index {1}, acceptance {2}%" -f `
        $vs.Range('A5').Text, $vs.Range('J5').Text, $vs.Range('G5').Text)

    Write-Host ""
    Write-Host "Sweeping every used cell for Excel error values..."
    $errCount = 0
    foreach ($ws in $wb.Worksheets) {
        if ($ws.Name -like 'Data_*') { continue }
        foreach ($cell in $ws.UsedRange) {
            if ($errorText -contains $cell.Text) {
                Write-Host ("  {0}!{1}  {2}   formula: {3}" -f `
                    $ws.Name, $cell.Address($false,$false), $cell.Text, $cell.Formula) -ForegroundColor Red
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
