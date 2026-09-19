<#
=============================================================================
Project 1 -- Ridgeline Field Services
Script:  Export_Dashboard_Data.ps1
Purpose: Regenerate every JSON in data_exports/ from the live database, and
         re-embed them into dashboard/index.html.

WHY THIS EXISTS
    It did not, and that was the defect. The seven JSON files and the DATA
    block inside dashboard/index.html were produced by hand, so there was no
    way to refresh them when the SQL beneath them changed -- and when the
    utilization view was corrected, every figure on the dashboard silently
    became the old, wrong one. Projects 2 and 3 both ship an equivalent
    exporter; this brings Project 1 to the same standard.

    The dashboard is deliberately a standalone file with its data inlined, so
    it opens from disk with no server and no CORS problem. That is a good
    property to keep, but it means the data has to be injected rather than
    fetched -- which is exactly the step that has to be scripted, or it will
    not be repeated.

USAGE
    powershell data_exports/Export_Dashboard_Data.ps1

DATA DISCLOSURE: Ridgeline Field Services is fictional; all data is synthetic.
No confidential data and no production system is involved.
=============================================================================
#>

$ErrorActionPreference = "Stop"

$SqlServerInst = "localhost\TEW_SQLEXPRESS"
$SqlDb         = "RidgelineFieldOps"
$outDir        = $PSScriptRoot
$dashboard     = Join-Path (Split-Path $PSScriptRoot -Parent) "dashboard\index.html"

function Invoke-Rows($query) {
    $cn = New-Object System.Data.SqlClient.SqlConnection("Server=$SqlServerInst;Database=$SqlDb;Integrated Security=True;")
    $cn.Open()
    try {
        $cmd = $cn.CreateCommand()
        $cmd.CommandText = $query
        $cmd.CommandTimeout = 300
        $da = New-Object System.Data.SqlClient.SqlDataAdapter($cmd)
        $dt = New-Object System.Data.DataTable
        $da.Fill($dt) | Out-Null
        $rows = @()
        foreach ($r in $dt.Rows) {
            $o = [ordered]@{}
            foreach ($c in $dt.Columns) {
                $v = $r[$c.ColumnName]
                if ($v -is [System.DBNull]) { $o[$c.ColumnName] = $null }
                elseif ($v -is [decimal])   { $o[$c.ColumnName] = [double]$v }
                else                        { $o[$c.ColumnName] = $v }
            }
            $rows += [pscustomobject]$o
        }
        return ,$rows
    } finally { $cn.Close() }
}

# ConvertTo-Json in PS 5.1 emits a bare object, not a one-element array, when
# handed a single row -- and the dashboard's JS indexes every one of these as
# an array. latestMonth is exactly one row, so this is not hypothetical.
function ConvertTo-JsonArray($rows) {
    if ($null -eq $rows -or @($rows).Count -eq 0) { return "[]" }
    $json = @($rows) | ConvertTo-Json -Compress -Depth 4
    if (-not $json.StartsWith("[")) { $json = "[$json]" }
    return $json
}

# The reporting month is read from the data, not hardcoded: the generator's
# window can move, and a hardcoded '2025-12' would quietly export an empty
# priority queue rather than failing.
$latest = Invoke-Rows @"
SELECT TOP 1 [Year], [Month] FROM dbo.vw_RegionKPIMonthly ORDER BY [Year] DESC, [Month] DESC;
"@
$ly = $latest[0].Year; $lm = $latest[0].Month
Write-Host "Reporting month: $ly-$('{0:00}' -f $lm)" -ForegroundColor Cyan

$sets = [ordered]@{
    regionMonthly = @"
SELECT RegionName, [Year], [Month], TotalJobs, CompletedJobs,
       FirstTimeFixPct, SLACompliancePct, CallbackPct, AvgUtilizationPct,
       TotalRevenue, TotalCost
FROM dbo.vw_RegionKPIMonthly
ORDER BY RegionName, [Year], [Month];
"@
    monthlyCategory = @"
SELECT d.[Year], d.[Month], s.ServiceCategory, COUNT(*) AS Jobs
FROM dbo.Fact_ServiceJobs f
JOIN dbo.Dim_Date d        ON d.DateKey = f.DateKey
JOIN dbo.Dim_ServiceType s ON s.ServiceTypeKey = f.ServiceTypeKey
GROUP BY d.[Year], d.[Month], s.ServiceCategory
ORDER BY d.[Year], d.[Month], s.ServiceCategory;
"@
    slaByRegion = @"
SELECT r.RegionName,
       OverallSLAPct = CAST(AVG(CASE WHEN f.JobStatus='Completed' THEN CAST(f.SLAMet AS FLOAT) END)*100 AS DECIMAL(6,1)),
       TightSLAPct   = CAST(AVG(CASE WHEN f.JobStatus='Completed' AND s.SLAHours=1.0 THEN CAST(f.SLAMet AS FLOAT) END)*100 AS DECIMAL(6,1))
FROM dbo.Fact_ServiceJobs f
JOIN dbo.Dim_Region r      ON r.RegionKey = f.RegionKey
JOIN dbo.Dim_ServiceType s ON s.ServiceTypeKey = f.ServiceTypeKey
GROUP BY r.RegionName
ORDER BY OverallSLAPct;
"@
    # Jobs is every dispatched job; CompletedJobs is the denominator the two
    # percentages beside it are actually computed over. Publishing only the
    # first, under a "Completed jobs" heading, is the error this column pair
    # exists to make impossible.
    skillQuality = @"
SELECT t.SkillLevel,
       FTFPct      = CAST(AVG(CASE WHEN f.JobStatus='Completed' THEN CAST(f.FirstTimeFix AS FLOAT) END)*100 AS DECIMAL(6,1)),
       CallbackPct = CAST(AVG(CASE WHEN f.JobStatus='Completed' THEN CAST(f.CallbackFlag AS FLOAT) END)*100 AS DECIMAL(6,1)),
       Jobs          = COUNT(*),
       CompletedJobs = SUM(CASE WHEN f.JobStatus='Completed' THEN 1 ELSE 0 END)
FROM dbo.Fact_ServiceJobs f
JOIN dbo.Dim_Technician t ON t.TechnicianKey = f.TechnicianKey
GROUP BY t.SkillLevel
ORDER BY t.SkillLevel;
"@
    dqSummary = @"
SELECT AnomalyType, AnomalyCount, TotalFactRows, AnomalyRatePct
FROM dbo.vw_DQ_Summary ORDER BY AnomalyType;
"@
    priorityQueue = @"
SELECT TechnicianID, TechnicianName, RegionName, SkillLevel,
       UtilizationPct, OvertimePctOfCapacity,
       FirstTimeFixPct, SLACompliancePct, CallbackPct,
       RiskScore, RecommendedAction
FROM dbo.vw_PriorityActionQueue
WHERE [Year] = $ly AND [Month] = $lm
ORDER BY RiskScore DESC, UtilizationPct ASC;
"@
    latestMonth = @"
SELECT TotalJobs     = SUM(TotalJobs),
       CompletedJobs = SUM(CompletedJobs),
       FTFPct        = CAST(SUM(FirstTimeFixPct*CompletedJobs)/NULLIF(SUM(CompletedJobs),0) AS DECIMAL(6,1)),
       SLAPct        = CAST(SUM(SLACompliancePct*CompletedJobs)/NULLIF(SUM(CompletedJobs),0) AS DECIMAL(6,1)),
       AvgUtilPct    = (SELECT CAST(SUM(WorkedMinutes)*100.0/NULLIF(SUM(AvailableMinutes),0) AS DECIMAL(6,1))
                        FROM dbo.vw_TechnicianUtilization WHERE [Year]=$ly AND [Month]=$lm),
       OvertimePct   = (SELECT CAST(SUM(OvertimeMinutes)*100.0/NULLIF(SUM(AvailableMinutes),0) AS DECIMAL(6,1))
                        FROM dbo.vw_TechnicianUtilization WHERE [Year]=$ly AND [Month]=$lm),
       Revenue       = SUM(TotalRevenue),
       Cost          = SUM(TotalCost)
FROM dbo.vw_RegionKPIMonthly WHERE [Year] = $ly AND [Month] = $lm;
"@
}

# The file names on disk are snake_case while the JS keys are camelCase.
$fileNames = @{
    regionMonthly   = "region_monthly.json"
    monthlyCategory = "monthly_category.json"
    slaByRegion     = "sla_by_region.json"
    skillQuality    = "skill_quality.json"
    dqSummary       = "dq_summary.json"
    priorityQueue   = "priority_queue.json"
    latestMonth     = "latest_month_summary.json"
}

$payloads = [ordered]@{}
foreach ($k in $sets.Keys) {
    $rows = Invoke-Rows $sets[$k]
    $json = ConvertTo-JsonArray $rows
    $payloads[$k] = $json
    $path = Join-Path $outDir $fileNames[$k]
    [System.IO.File]::WriteAllText($path, $json, (New-Object System.Text.UTF8Encoding($false)))
    Write-Host ("  {0,-16} {1,5} rows -> {2}" -f $k, @($rows).Count, $fileNames[$k])
}

# --- re-embed into the standalone dashboard --------------------------------
# Rewrite only the lines between 'const DATA = {' and its closing '};'. Editing
# by line offset would break the moment the file above it changed by one line.
$lines = [System.IO.File]::ReadAllLines($dashboard)
$start = -1; $end = -1
for ($i = 0; $i -lt $lines.Count; $i++) {
    if ($start -lt 0 -and $lines[$i].Trim() -eq "const DATA = {") { $start = $i; continue }
    if ($start -ge 0 -and $lines[$i].Trim() -eq "};") { $end = $i; break }
}
if ($start -lt 0 -or $end -lt 0) { throw "Could not find the 'const DATA = { ... };' block in $dashboard" }

$block = @("const DATA = {")
foreach ($k in $payloads.Keys) { $block += "  $k`: $($payloads[$k])," }
$block += "};"

$new = @()
$new += $lines[0..($start-1)]
$new += $block
if ($end + 1 -le $lines.Count - 1) { $new += $lines[($end+1)..($lines.Count-1)] }
[System.IO.File]::WriteAllLines($dashboard, $new, (New-Object System.Text.UTF8Encoding($false)))

Write-Host ""
Write-Host "Dashboard data block rewritten ($($block.Count - 2) datasets)." -ForegroundColor Green
Write-Host "Open dashboard/index.html to confirm it renders." -ForegroundColor Green
