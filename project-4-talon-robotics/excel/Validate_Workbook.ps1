<#
================================================================================
Project 4 -- Talon Robotics: Payload Deployment Delivery
Script:  Validate_Workbook.ps1
Purpose: Reopen the built workbook, recalculate it, and read the values back.

WHY THIS EXISTS SEPARATELY FROM THE BUILD
    Build_Workbook.ps1 can report complete success while producing a file full
    of #REF!. Writing a formula never evaluates it, so a build script is
    reporting that it wrote text, not that the text means anything. The only
    check worth trusting is the one the consuming application makes: open the
    file, force a full recalculation, and read what Excel actually computed.

    That is not hypothetical. Project 2 in this portfolio shipped a build that
    reported success against a workbook of #REF!, and Project 1 shipped one
    reporting Total Cost as a negative number, with no error value anywhere and
    every conditional format green.

WHAT IT COMPARES AGAINST
    Each figure is read from SQL at run time as well as from the workbook, and
    separately checked against the value published in the README and case
    study. Those are two different failures:

        Excel <> SQL        the workbook's formulas disagree with the database
        SQL   <> published  the database has moved since the documents were written

    A single frozen target conflates them, and a frozen target is what went
    stale in Project 3.

DATA DISCLOSURE: Talon Robotics is fictional; all data is synthetic. No
confidential data and no production system is involved.
================================================================================
#>

param(
    [string]$WorkbookPath  = (Join-Path $PSScriptRoot 'Talon_Readiness_Control.xlsx'),
    [string]$ServerInstance = 'localhost\TEW_SQLEXPRESS',
    [string]$Database       = 'TalonDelivery'
)

$ErrorActionPreference = 'Stop'
if (-not (Test-Path $WorkbookPath)) { throw "Workbook not found: $WorkbookPath" }

function Get-SqlScalar([string]$query) {
    $cn = New-Object System.Data.SqlClient.SqlConnection("Server=$ServerInstance;Database=$Database;Integrated Security=True;")
    $cn.Open()
    try {
        $cmd = $cn.CreateCommand(); $cmd.CommandText = $query; $cmd.CommandTimeout = 300
        $v = $cmd.ExecuteScalar()
        if ($null -eq $v -or $v -is [System.DBNull]) { return $null }
        return [double]$v
    } finally { $cn.Close() }
}

# Dashboard cell, label, the SQL that answers it, the published figure, tolerance.
$checks = @(
    @{ Sheet='Dashboard'; Cell='B6';  Label='Work-item completion %'; Published=93.86; Tol=0.02;
       Sql="SELECT WorkItemCompletionPct FROM dbo.vw_ReadinessKPI" },
    @{ Sheet='Dashboard'; Cell='B7';  Label='Verification coverage %'; Published=96.53; Tol=0.02;
       Sql="SELECT VerificationCoveragePct FROM dbo.vw_ReadinessKPI" },
    @{ Sheet='Dashboard'; Cell='B8';  Label='Ship readiness %';       Published=51.96; Tol=0.02;
       Sql="SELECT ShipReadinessPct FROM dbo.vw_ReadinessKPI" },
    @{ Sheet='Dashboard'; Cell='B13'; Label='Ship readiness (card)';  Published=51.96; Tol=0.02;
       Sql="SELECT ShipReadinessPct FROM dbo.vw_ReadinessKPI" },
    @{ Sheet='Dashboard'; Cell='B14'; Label='Verification coverage';  Published=96.53; Tol=0.02;
       Sql="SELECT VerificationCoveragePct FROM dbo.vw_ReadinessKPI" },
    @{ Sheet='Dashboard'; Cell='B15'; Label='Stale verification %';   Published=24.48; Tol=0.02;
       Sql="SELECT StaleVerificationPct FROM dbo.vw_ReadinessKPI" },
    @{ Sheet='Dashboard'; Cell='B16'; Label='Policy compliance %';    Published=77.45; Tol=0.02;
       Sql="SELECT PolicyCompliancePct FROM dbo.vw_ReadinessKPI" },
    @{ Sheet='Dashboard'; Cell='B17'; Label='Work-item completion';   Published=93.86; Tol=0.02;
       Sql="SELECT WorkItemCompletionPct FROM dbo.vw_ReadinessKPI" },
    @{ Sheet='Dashboard'; Cell='B18'; Label='Critical RAID open';     Published=9;     Tol=0.5;
       Sql="SELECT CriticalRAIDOpen FROM dbo.vw_ReadinessKPI" },
    @{ Sheet='Dashboard'; Cell='B19'; Label='Overdue RAID %';         Published=0;     Tol=0.05;
       Sql="SELECT OverdueRAIDPct FROM dbo.vw_ReadinessKPI" }
)

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
$failures = 0
$errorText = @('#REF!','#VALUE!','#NAME?','#DIV/0!','#N/A','#NULL!','#NUM!')

try {
    Write-Host ""
    Write-Host "Validating $(Split-Path $WorkbookPath -Leaf)" -ForegroundColor Cyan
    Write-Host ("-" * 92)

    Write-Host "Reading the current figures from SQL..." -ForegroundColor Cyan
    $sqlUnavailable = $false
    foreach ($c in $checks) {
        try { $c.Expect = Get-SqlScalar $c.Sql }
        catch { $c.Expect = $c.Published; $sqlUnavailable = $true }
    }
    if ($sqlUnavailable) {
        Write-Host "  SQL Server unreachable -- falling back to the PUBLISHED figures." -ForegroundColor Yellow
        Write-Host "  This checks the workbook against the documents, not against the database." -ForegroundColor Yellow
    }

    $wb = $xl.Workbooks.Open($WorkbookPath, 0, $true)   # read-only
    $xl.Calculation = -4105
    $wb.Application.CalculateFullRebuild()

    Write-Host ""
    Write-Host "Excel formulas against the figures SQL produced independently:"
    foreach ($c in $checks) {
        $v = $wb.Worksheets.Item($c.Sheet).Range($c.Cell).Value2
        if ($null -eq $v -or $v -isnot [double]) {
            Write-Host ("  {0,-26} {1,16}  NOT A NUMBER" -f $c.Label, "$v") -ForegroundColor Red
            $failures++; continue
        }
        $diff = [math]::Abs($v - $c.Expect)
        $ok = $diff -le $c.Tol
        if (-not $ok) { $failures++ }
        Write-Host ("  {0,-26} excel {1,12:N2}   sql {2,12:N2}   diff {3,8:N4}  {4}" -f `
            $c.Label, $v, $c.Expect, $diff, $(if ($ok) { 'MATCH' } else { 'DIFFERS' })) `
            -ForegroundColor $(if ($ok) { 'Green' } else { 'Red' })

        # Separately: has SQL moved away from what the documents say?
        if (-not $sqlUnavailable -and $c.Published -gt 0) {
            $drift = [math]::Abs($c.Expect - $c.Published)
            if ($drift -gt $c.Tol) {
                $failures++
                Write-Host ("      ^ SQL now says {0:N2} but the README and case study publish {1:N2}. Update the documents." -f `
                    $c.Expect, $c.Published) -ForegroundColor Red
            }
        }
    }

    # ---- the workbook's own reconciliation sheet --------------------------
    Write-Host ""
    Write-Host "The workbook's own Excel-versus-SQL reconciliation sheet:"
    $val = $wb.Worksheets.Item('Validation')
    $row = 5; $match = 0; $differs = 0
    while ($val.Range("B$row").Value2) {
        if ($val.Range("F$row").Text -eq 'MATCH') { $match++ }
        else {
            $differs++
            Write-Host ("  DIFFERS: {0}  excel={1}  sql={2}" -f `
                $val.Range("B$row").Text, $val.Range("C$row").Text, $val.Range("D$row").Text) -ForegroundColor Red
        }
        $row++
    }
    if ($differs -gt 0) { $failures += $differs }
    Write-Host ("  {0} of {1} checks reconcile" -f $match, ($match + $differs)) `
        -ForegroundColor $(if ($differs -eq 0) { 'Green' } else { 'Red' })

    # ---- the queue actually bounds itself ---------------------------------
    Write-Host ""
    # Read by NAME, not by cell. The schedule block sits below the subsystem
    # table, so its row moves whenever a subsystem is added, and a fixed cell
    # reference would quietly read a neighbouring figure and report it as fine.
    $qOutstanding = $wb.Names.Item('QueueOutstanding').RefersToRange.Value2
    $qThisWeek    = $wb.Names.Item('SchedulableThisWeek').RefersToRange.Value2
    $qRigWeeks    = $wb.Names.Item('RigWeeksOutstanding').RefersToRange.Value2
    Write-Host ("Queue: {0} outstanding, {1} schedulable this week, {2:N1} rig-weeks to clear the backlog" -f `
        $qOutstanding, $qThisWeek, $qRigWeeks)
    $sqlOutstanding = Get-SqlScalar "SELECT COUNT(*) FROM dbo.vw_VerificationQueue"
    if ([math]::Abs($qOutstanding - $sqlOutstanding) -gt 0.5) {
        $failures++
        Write-Host ("  Queue size disagrees: excel {0} vs sql {1}" -f $qOutstanding, $sqlOutstanding) -ForegroundColor Red
    }

    # ---- sweep every used cell for error values ---------------------------
    Write-Host ""
    Write-Host "Sweeping every used cell for Excel error values..."
    $errCount = 0
    foreach ($ws in $wb.Worksheets) {
        $used = $ws.UsedRange
        if ($null -eq $used) { continue }
        foreach ($t in $errorText) {
            $found = $used.Find($t, [System.Reflection.Missing]::Value, -4163, 2)
            if ($null -ne $found) {
                Write-Host ("  {0}!{1} = {2}" -f $ws.Name, $found.Address(0,0), $t) -ForegroundColor Red
                $errCount++
            }
        }
    }
    if ($errCount -eq 0) { Write-Host "  No error values found." -ForegroundColor Green }
    else { $failures += $errCount }

    Write-Host ("-" * 92)
    if ($failures -eq 0) {
        Write-Host "WORKBOOK VALIDATED: every formula resolves and every figure reconciles with SQL." -ForegroundColor Green
    } else {
        Write-Host "WORKBOOK VALIDATION FAILED: $failures problem(s) above." -ForegroundColor Red
    }
}
finally {
    if ($wb) { $wb.Close($false) }
    $xl.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
if ($failures -gt 0) { exit 1 }
Write-Host ""
