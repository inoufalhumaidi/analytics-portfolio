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

# The figures SQL published, independently of anything Excel does.
$sqlFigures = @(
    @{ Sheet='Dashboard'; Cell='C8';  Label='Extended price (12m)';    Expect=48586386.86; Tol=0.05 },
    @{ Sheet='Dashboard'; Cell='C9';  Label='Landed cost (12m)';       Expect=49474785.92; Tol=0.05 },
    @{ Sheet='Dashboard'; Cell='G9';  Label='Erosion capture %';       Expect=29.50;       Tol=0.02 },
    @{ Sheet='Dashboard'; Cell='G10'; Label='Annual opportunity';      Expect=4047608.72;  Tol=0.05 },
    @{ Sheet='Dashboard'; Cell='B20'; Label='Erosion capture (score)'; Expect=29.50;       Tol=0.02 },
    @{ Sheet='Dashboard'; Cell='B21'; Label='Maverick spend %';        Expect=27.09;       Tol=0.02 },
    @{ Sheet='Dashboard'; Cell='B22'; Label='Price variance %';        Expect=-0.19;       Tol=0.02 },
    @{ Sheet='Dashboard'; Cell='B23'; Label='Acceptance rate %';       Expect=99.11;       Tol=0.02 },
    @{ Sheet='Dashboard'; Cell='B24'; Label='On-time delivery %';      Expect=80.62;       Tol=0.02 },
    @{ Sheet='Dashboard'; Cell='B25'; Label='Expedite spend %';        Expect=0.23;        Tol=0.02 },
    @{ Sheet='Dashboard'; Cell='B26'; Label='Single-source spend %';   Expect=27.17;       Tol=0.02 },
    @{ Sheet='Dashboard'; Cell='B27'; Label='Top-5 vendor share %';    Expect=27.20;       Tol=0.05 }
)

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
