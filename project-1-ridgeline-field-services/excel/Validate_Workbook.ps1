$ErrorActionPreference = "Stop"
$Path = Join-Path $PSScriptRoot "Ridgeline_KPI_Dashboard.xlsx"

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
# Read-only: opening read-write lets Excel rewrite the file's timestamp and
# calculation chain just by looking at it, so a validation run that changed
# nothing still shows the artefact as modified. Projects 2-4 already do this.
$wb = $excel.Workbooks.Open($Path, 0, $true)   # 3rd arg = ReadOnly
$excel.CalculateFullRebuild()
$excel.CalculateUntilAsyncQueriesDone()

function Test-SheetForErrors($sheetName, $usedRangeOnly=$true) {
    $ws = $wb.Worksheets.Item($sheetName)
    $used = $ws.UsedRange
    $vals = $used.Value2
    $errCount = 0
    $errSamples = @()
    if ($vals -is [System.Array]) {
        $rows = $vals.GetLength(0); $cols = $vals.GetLength(1)
        for ($r=1; $r -le $rows; $r++) {
            for ($c=1; $c -le $cols; $c++) {
                $v = $vals[$r,$c]
                if ($v -is [string] -and $v.StartsWith("#")) {
                    $errCount++
                    if ($errSamples.Count -lt 5) { $errSamples += "R$($r)C$($c)=$v" }
                }
            }
        }
    }
    Write-Host "[$sheetName] error cells: $errCount $(if($errSamples.Count){'-- ' + ($errSamples -join ', ')})"
}

Write-Host "=== Scanning all sheets for formula errors ==="
foreach ($name in @("KPI_Dashboard","Priority_Action_Queue","Data_Quality","Traceability_UAT")) {
    Test-SheetForErrors $name
}

Write-Host ""
Write-Host "=== Data_Quality sheet: Excel vs SQL cross-check ==="
$wsDQ = $wb.Worksheets.Item("Data_Quality")
for ($r=5; $r -le 7; $r++) {
    $type = $wsDQ.Cells.Item($r,1).Value2
    $excelCount = $wsDQ.Cells.Item($r,3).Value2
    $sqlCount = $wsDQ.Cells.Item($r,5).Value2
    $result = $wsDQ.Cells.Item($r,7).Value2
    Write-Host "$type : Excel=$excelCount SQL=$sqlCount -> $result"
}

Write-Host ""
Write-Host "=== KPI_Dashboard: set filters to South Metro / 2025 / 12, read values ==="
$wsDash = $wb.Worksheets.Item("KPI_Dashboard")
$wsDash.Range("B4").Value2 = "South Metro"
$wsDash.Range("B5").Value2 = 2025
$wsDash.Range("B6").Value2 = 12
$excel.CalculateFullRebuild()
for ($r=9; $r -le 17; $r++) {
    $name = $wsDash.Cells.Item($r,1).Value2
    $val = $wsDash.Cells.Item($r,2).Value2
    Write-Host "$name = $val"
}

Write-Host ""
Write-Host "=== Priority_Action_Queue: top 5 ranked rows ==="
$wsQ = $wb.Worksheets.Item("Priority_Action_Queue")
for ($r=6; $r -le 10; $r++) {
    $name = $wsQ.Cells.Item($r,15).Value2
    $region = $wsQ.Cells.Item($r,17).Value2
    $risk = $wsQ.Cells.Item($r,22).Value2
    $action = $wsQ.Cells.Item($r,23).Value2
    Write-Host "Rank $($r-5): $name / $region / Risk=$risk / $action"
}

$wb.Close($false)
$excel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
Write-Host ""
Write-Host "Validation pass complete."
