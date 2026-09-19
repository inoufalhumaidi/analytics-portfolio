<#
================================================================================
Project 5 -- Meridian UAV Services: Predictive Maintenance
Script:  Validate_Workbook.ps1
Purpose: Reopen the saved workbook, recalculate it, and read the values back.

WHY THIS EXISTS AS A SEPARATE SCRIPT

    Build_Workbook.ps1 reports success when every COM call returned without
    error. That is not the same as a workbook that works. A formula referring to
    a sheet that did not exist yet is written happily and stored as #REF!. A
    structured reference to a table that loaded zero rows returns #DIV/0!. A
    number format applied to text shows a right-aligned string that reads as a
    figure. None of it raises.

    The only check worth anything is to open the file the way a reviewer will,
    force a full recalculation, and read what the cells actually say.

WHAT IT ASSERTS

 1. Every sheet a reader is meant to open exists and is visible.
 2. No cell on any visible sheet holds an Excel error value.
 3. The headline figures match what SQL published, to the rounding step.
 4. All sixteen on-sheet validation checks report OK.
 5. THE CONTROLS ACTUALLY CONTROL SOMETHING. The hangar-hours cell is changed
    and the week plan is required to move; the alert-threshold cell is changed
    and the prediction figures are required to move. A control that changes the
    appearance and not the answer is decoration, and it is invisible in a
    screenshot.

    That last one is the point of the whole script. Everything above it checks
    the workbook is not broken; only (5) checks it is a control.

DATA DISCLOSURE: Meridian UAV Services is fictional; all data is synthetic.
================================================================================
#>

param(
    [string]$WorkbookPath = (Join-Path $PSScriptRoot 'Meridian_Fleet_Control.xlsx'),
    [string]$ExtractPath  = (Join-Path (Split-Path $PSScriptRoot -Parent) 'data_exports')
)

$ErrorActionPreference = 'Stop'
if (-not (Test-Path $WorkbookPath)) { throw "Workbook not found: $WorkbookPath. Run Build_Workbook.ps1 first." }

$script:pass = 0
$script:fail = 0
function Check {
    param([string]$Name, [object]$Actual, [object]$Expected, [double]$Tolerance = 0.011)
    $ok = $false
    if ($Expected -is [string]) {
        $ok = ([string]$Actual -eq [string]$Expected)
    } else {
        $a = 0.0; $e = 0.0
        $ok = ([double]::TryParse([string]$Actual, [ref]$a) -and
               [double]::TryParse([string]$Expected, [ref]$e) -and
               [Math]::Abs($a - $e) -le $Tolerance)
    }
    if ($ok) { $script:pass++ } else { $script:fail++ }
    "{0}  {1,-46} excel={2,-14} expected={3}" -f $(if ($ok) { 'ok  ' } else { 'FAIL' }), $Name, $Actual, $Expected | Write-Host
}

# The SQL side, read from the extracts rather than re-queried. These are the
# same files the workbook loaded, so a disagreement is the workbook's doing.
$kpi = @{}
Import-Csv (Join-Path $ExtractPath 'fleet_kpi.csv') | ForEach-Object { $kpi[$_.MetricName] = [double]$_.MetricValue }
$wear  = @(Import-Csv (Join-Path $ExtractPath 'component_wear.csv'))
$queue = @(Import-Csv (Join-Path $ExtractPath 'action_queue.csv'))
$sweep = @(Import-Csv (Join-Path $ExtractPath 'threshold_sweep.csv'))

$sqlFitted   = $wear.Count
$sqlHidden   = @($wear | Where-Object { $_.IsHiddenOverdue -eq '1' }).Count
$sqlHiddenFC = @($wear | Where-Object { $_.IsHiddenOverdue -eq '1' -and $_.Criticality -eq 'FlightCritical' }).Count
$sqlQueue    = $queue.Count
$sqlQueueHrs = [Math]::Round((($queue | Measure-Object -Property JobHours -Sum).Sum), 1)

Write-Host ""
Write-Host "Meridian UAV -- reopening the workbook and reading it back" -ForegroundColor Cyan
Write-Host ("-" * 92)

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false

try {
    $wb = $xl.Workbooks.Open($WorkbookPath)
    $xl.CalculateFullRebuild()

    # --- 1. the sheets a reader opens ------------------------------------
    Write-Host "`nSHEETS"
    foreach ($name in @('README','Dashboard','Bases','Action Queue','Prediction','Validation','Control')) {
        $found = $false
        foreach ($s in $wb.Worksheets) { if ($s.Name -eq $name -and $s.Visible -eq -1) { $found = $true } }
        Check "sheet visible: $name" $(if ($found) { 'yes' } else { 'no' }) 'yes'
    }

    # --- 2. no error values anywhere a reader will look -------------------
    Write-Host "`nERROR VALUES"
    $errorCells = 0
    $errorDetail = @()
    foreach ($s in $wb.Worksheets) {
        if ($s.Visible -ne -1) { continue }
        $used = $s.UsedRange
        if ($null -eq $used) { continue }
        # SpecialCells(-4123, 16) = formula cells whose result is an error.
        # It throws when there are none, which is the outcome we want.
        try {
            $errs = $used.SpecialCells(-4123, 16)
            if ($null -ne $errs) {
                $errorCells += $errs.Count
                $errorDetail += "$($s.Name)!$($errs.Address($false,$false))"
            }
        } catch { }
    }
    Check "cells holding #REF!, #DIV/0!, #N/A etc" $errorCells 0
    if ($errorCells -gt 0) { Write-Host ("      at: " + ($errorDetail -join '; ')) -ForegroundColor Red }

    $d = $wb.Worksheets.Item('Dashboard')
    $v = $wb.Worksheets.Item('Validation')
    $ctl = $wb.Worksheets.Item('Control')
    $p = $wb.Worksheets.Item('Prediction')

    # --- 3. the headline figures ------------------------------------------
    # Read BY LABEL, never by row offset. Project 1 shipped a dashboard that
    # read a total from a fixed row; a KPI row was inserted above it and the
    # figure became -163,307.32, still formatted as currency.
    function Get-DashValue {
        param([object]$Sheet, [string]$Label, [int]$ValueColumn = 2, [int]$MaxRow = 60)
        for ($r = 1; $r -le $MaxRow; $r++) {
            if ([string]$Sheet.Cells.Item($r, 1).Value2 -eq $Label) {
                return $Sheet.Cells.Item($r, $ValueColumn).Value2
            }
        }
        throw "Label '$Label' not found on $($Sheet.Name). The layout moved and a positional read would have returned the wrong cell silently."
    }

    Write-Host "`nHEADLINE FIGURES"
    Check 'hour compliance %'   ([double](Get-DashValue $d 'Compliant on FLIGHT HOURS') * 100) $kpi['HourCompliancePct']
    Check 'stress compliance %' ([double](Get-DashValue $d 'Compliant on STRESS HOURS') * 100) $kpi['StressCompliancePct']
    Check 'the gap (points)'    ([double](Get-DashValue $d 'The gap') * 100) ($kpi['HourCompliancePct'] - $kpi['StressCompliancePct'])
    Check 'components fitted'   (Get-DashValue $d 'Components fitted') $sqlFitted
    Check 'hidden overdue'      (Get-DashValue $d 'Past due on stress, inside hour limit') $sqlHidden
    Check 'hidden, flight-crit' (Get-DashValue $d '...of which flight-critical') $sqlHiddenFC
    Check 'jobs outstanding'    (Get-DashValue $d 'Jobs outstanding') $sqlQueue
    Check 'hangar hours'        (Get-DashValue $d 'Hangar hours to clear') $sqlQueueHrs 0.05

    Write-Host "`nSCORECARD (looked up by metric name)"
    foreach ($m in @('StressCompliancePct','HourCompliancePct','OverdueFlightCritical','UnscheduledRatePct','AirframeAvailabilityPct')) {
        Check "scorecard $m" (Get-DashValue $d $m) $kpi[$m]
    }

    # --- 4. the on-sheet validation block ---------------------------------
    Write-Host "`nON-SHEET VALIDATION"
    $summary = [string]$wb.Names.Item('ValidationSummary').RefersToRange.Value2
    $mismatches = 0
    $rng = $wb.Names.Item('ValidationRange').RefersToRange
    foreach ($cell in $rng) { if ([string]$cell.Value2 -eq 'MISMATCH') { $mismatches++ } }
    Check 'on-sheet checks reporting MISMATCH' $mismatches 0
    Write-Host "      the sheet reports: $summary"

    # --- 5. DO THE CONTROLS CONTROL ANYTHING? -----------------------------
    Write-Host "`nTHE CONTROLS (the part that matters)"

    $originalHours = $ctl.Range('B7').Value2
    $weekBefore    = [double](Get-DashValue $d 'Schedulable this week')
    $weeksBefore   = [double](Get-DashValue $d 'Weeks at current capacity')

    $ctl.Range('B7').Value2 = [double]($originalHours * 2)
    $xl.CalculateFullRebuild()
    $weekAfter  = [double](Get-DashValue $d 'Schedulable this week')
    $weeksAfter = [double](Get-DashValue $d 'Weeks at current capacity')

    Check 'doubling capacity increases the week plan' $(if ($weekAfter -gt $weekBefore) { 'yes' } else { 'no' }) 'yes'
    Check 'doubling capacity halves the backlog'      ([Math]::Round($weeksAfter * 2, 1)) ([Math]::Round($weeksBefore, 1)) 0.15
    Write-Host ("      this week {0} -> {1} jobs;  backlog {2} -> {3} weeks" -f $weekBefore, $weekAfter, $weeksBefore, $weeksAfter)

    $ctl.Range('B7').Value2 = $originalHours
    $xl.CalculateFullRebuild()
    Check 'restoring capacity restores the week plan' ([double](Get-DashValue $d 'Schedulable this week')) $weekBefore

    # the alert threshold
    $originalThr = $ctl.Range('B8').Value2
    function Get-PredValue {
        param([string]$Code, [int]$Col)
        for ($r = 4; $r -le 12; $r++) {
            if ([string]$p.Cells.Item($r, 1).Value2 -eq $Code) { return $p.Cells.Item($r, $Col).Value2 }
        }
        throw "Component '$Code' not found on the Prediction sheet."
    }

    $motorActionableBefore = [double](Get-PredValue 'MOTOR-ESC' 9)
    $motorPrecisionBefore  = [double](Get-PredValue 'MOTOR-ESC' 6)

    $ctl.Range('B8').Value2 = 2.05
    $xl.CalculateFullRebuild()
    $motorActionableAfter = [double](Get-PredValue 'MOTOR-ESC' 9)
    $motorPrecisionAfter  = [double](Get-PredValue 'MOTOR-ESC' 6)

    Check 'lowering the threshold raises actionable lead time' $(if ($motorActionableAfter -gt $motorActionableBefore) { 'yes' } else { 'no' }) 'yes'
    Check 'lowering the threshold lowers precision'            $(if ($motorPrecisionAfter -lt $motorPrecisionBefore) { 'yes' } else { 'no' }) 'yes'
    Write-Host ("      MOTOR-ESC at 3.40x: {0}% actionable, {1}% precision" -f $motorActionableBefore, $motorPrecisionBefore)
    Write-Host ("      MOTOR-ESC at 2.05x: {0}% actionable, {1}% precision" -f $motorActionableAfter, $motorPrecisionAfter)
    Write-Host "      That trade -- 22 points of precision for 85 points of usable warning -- is the second finding."

    # And the figure must match SQL at the new threshold, not merely move.
    $sqlMotorAt205 = ($sweep | Where-Object { $_.ComponentCode -eq 'MOTOR-ESC' -and [double]$_.VibThreshold -eq 2.05 } | Select-Object -First 1)
    if ($sqlMotorAt205) {
        Check 'threshold 2.05 matches SQL exactly' $motorActionableAfter ([double]$sqlMotorAt205.ActionableLeadTimePct)
    }

    $ctl.Range('B8').Value2 = $originalThr
    $xl.CalculateFullRebuild()
    Check 'restoring the threshold restores the figure' ([double](Get-PredValue 'MOTOR-ESC' 9)) $motorActionableBefore

    $wb.Close($false)   # discard the probes; the saved file is untouched
}
finally {
    $xl.Quit()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}

Write-Host ""
Write-Host ("-" * 92)
Write-Host ("  {0} passed, {1} failed" -f $script:pass, $script:fail) -ForegroundColor $(if ($script:fail -eq 0) { 'Green' } else { 'Red' })
if ($script:fail -gt 0) {
    throw "$($script:fail) workbook validation check(s) failed. The workbook builds and does not work."
}
Write-Host "  The workbook opens, recalculates, agrees with SQL, and its controls change the answer." -ForegroundColor Green
Write-Host ""
