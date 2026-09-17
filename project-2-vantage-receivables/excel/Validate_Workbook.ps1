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

# The figures SQL published, independently of anything Excel does.
$sqlFigures = @(
    @{ Sheet='Dashboard'; Cell='C8';  Label='Total open AR';          Expect=10847081.49; Tol=0.005 },
    @{ Sheet='Dashboard'; Cell='C9';  Label='Not yet due';            Expect=8095793.78;  Tol=0.005 },
    @{ Sheet='Dashboard'; Cell='C10'; Label='Past due, disputed';     Expect=154864.71;   Tol=0.005 },
    @{ Sheet='Dashboard'; Cell='C11'; Label='Past due, undisputed';   Expect=2596423.00;  Tol=0.005 },
    @{ Sheet='Dashboard'; Cell='C13'; Label='Open invoices';          Expect=2346;        Tol=0.5   },
    @{ Sheet='Dashboard'; Cell='F9';  Label='Granted days';           Expect=45.97;       Tol=0.01  },
    @{ Sheet='Dashboard'; Cell='F10'; Label='Dispute days';           Expect=0.88;        Tol=0.01  },
    @{ Sheet='Dashboard'; Cell='F11'; Label='Lateness days';          Expect=14.74;       Tol=0.01  },
    @{ Sheet='Dashboard'; Cell='F12'; Label='Classic DSO';            Expect=61.59;       Tol=0.01  },
    @{ Sheet='Dashboard'; Cell='F14'; Label='Weighted average terms'; Expect=46.68;       Tol=0.01  },
    @{ Sheet='Dashboard'; Cell='F17'; Label='Billing lag (days)';     Expect=1.78;        Tol=0.02  },
    @{ Sheet='Dashboard'; Cell='F18'; Label='True cash cycle (days)'; Expect=63.37;       Tol=0.02  },
    @{ Sheet='Dashboard'; Cell='B27'; Label='Unapplied cash % of AR'; Expect=3.87;        Tol=0.02  },
    @{ Sheet='Dashboard'; Cell='B29'; Label='Promise kept rate %';    Expect=86.29;       Tol=0.01  },
    @{ Sheet='Calc';      Cell='C20'; Label='Countback DSO';          Expect=63.07;       Tol=0.01  },
    @{ Sheet='Calc';      Cell='C22'; Label='Best possible DSO';      Expect=46.27;       Tol=0.02  },
    @{ Sheet='Calc';      Cell='C23'; Label='Average days delinquent';Expect=16.80;       Tol=0.02  }
)

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
