# =============================================================================
# Project 1 -- Ridgeline Field Services: Operational Efficiency
# Build_Workbook.ps1
#
# Builds Ridgeline_KPI_Dashboard.xlsx directly from the live SQL Server
# database (RidgelineFieldOps) via ADO.NET + Excel COM automation. Documented
# here for transparency/reproducibility -- re-running this script regenerates
# the workbook from whatever is currently in the database.
#
# Sheets produced:
#   ReadMe                  - business question, data disclosure, sheet map
#   Raw_Data                - job-level fact extract as an Excel Table
#   Targets                 - editable KPI thresholds (drives every formula)
#   KPI_Dashboard            - formula-driven operational control (SUMIFS/AVERAGEIFS)
#   Priority_Action_Queue   - ranked, formula-driven technician worklist
#   Data_Quality            - formula-driven anomaly detection
#   Traceability_UAT        - requirements traceability + UAT summary
# =============================================================================

$ErrorActionPreference = "Stop"

$ConnStr = "Server=localhost\TEW_SQLEXPRESS;Database=RidgelineFieldOps;Integrated Security=True;TrustServerCertificate=True;"
$OutPath = Join-Path $PSScriptRoot "Ridgeline_KPI_Dashboard.xlsx"

# -----------------------------------------------------------------------------
# 1. Pull data via ADO.NET
# -----------------------------------------------------------------------------
Add-Type -AssemblyName System.Data

function Get-DataTable($query) {
    $conn = New-Object System.Data.SqlClient.SqlConnection($ConnStr)
    $cmd = New-Object System.Data.SqlClient.SqlCommand($query, $conn)
    $cmd.CommandTimeout = 120
    $da = New-Object System.Data.SqlClient.SqlDataAdapter($cmd)
    $dt = New-Object System.Data.DataTable
    $da.Fill($dt) | Out-Null
    $conn.Close()
    Write-Output -NoEnumerate $dt   # prevent PowerShell from unrolling the DataTable's rows into the pipeline
}

Write-Host "Querying Fact_ServiceJobs extract..."
$jobsQuery = @"
SELECT
    f.JobID,
    d.[Date]                AS JobDate,
    d.[Year],
    d.[Month],
    d.MonthName,
    r.RegionName,
    t.TechnicianID,
    t.TechnicianName,
    t.SkillLevel,
    c.ClientType,
    c.ContractTier,
    st.ServiceCategory,
    st.ServiceName,
    st.SLAHours,
    f.ScheduledStart,
    f.ActualStart,
    f.ActualEnd,
    f.TravelTimeMin,
    f.JobDurationMin,
    CAST(f.FirstTimeFix AS INT) AS FirstTimeFix,
    CAST(f.CallbackFlag AS INT) AS CallbackFlag,
    CAST(f.SLAMet AS INT) AS SLAMet,
    f.JobCost,
    f.JobRevenue,
    f.JobStatus
FROM dbo.Fact_ServiceJobs f
JOIN dbo.Dim_Date d ON d.DateKey = f.DateKey
JOIN dbo.Dim_Region r ON r.RegionKey = f.RegionKey
JOIN dbo.Dim_Technician t ON t.TechnicianKey = f.TechnicianKey
JOIN dbo.Dim_Client c ON c.ClientKey = f.ClientKey
JOIN dbo.Dim_ServiceType st ON st.ServiceTypeKey = f.ServiceTypeKey
ORDER BY d.[Date], t.TechnicianID;
"@
$jobsTable = Get-DataTable $jobsQuery
Write-Host ("Rows fetched: {0}" -f $jobsTable.Rows.Count)

Write-Host "Querying Dim_Technician (all, incl. inactive) for the Priority Action Queue helper table..."
$techQuery = @"
SELECT t.TechnicianID, t.TechnicianName, r.RegionName, t.SkillLevel, t.EmploymentStatus, t.DailyCapacityHours, t.HireDate
FROM dbo.Dim_Technician t JOIN dbo.Dim_Region r ON r.RegionKey = t.RegionKey
WHERE t.EmploymentStatus = 'Active'
ORDER BY r.RegionName, t.TechnicianID;
"@
$techTable = Get-DataTable $techQuery

Write-Host "Querying Ref_KPITargets..."
$targetsTable = Get-DataTable "SELECT MetricName, TargetValue, WarningValue, Direction, Description FROM dbo.Ref_KPITargets ORDER BY MetricName;"

Write-Host "Querying DQ summary for cross-check values..."
$dqTable = Get-DataTable "SELECT AnomalyType, AnomalyCount, TotalFactRows, AnomalyRatePct FROM dbo.vw_DQ_Summary ORDER BY AnomalyType;"

# -----------------------------------------------------------------------------
# 2. Open Excel
# -----------------------------------------------------------------------------
Write-Host "Launching Excel..."
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$wb = $excel.Workbooks.Add()

# Remove default extra sheets down to 1, we'll add the rest by name
while ($wb.Worksheets.Count -gt 1) { $wb.Worksheets.Item($wb.Worksheets.Count).Delete() }
$wsReadMe = $wb.Worksheets.Item(1)
$wsReadMe.Name = "ReadMe"

function New-Sheet($name) {
    $s = $wb.Worksheets.Add([System.Reflection.Missing]::Value, $wb.Worksheets.Item($wb.Worksheets.Count))
    $s.Name = $name
    return $s
}

$wsRaw = New-Sheet "Raw_Data"
$wsTargets = New-Sheet "Targets"
$wsDash = New-Sheet "KPI_Dashboard"
$wsQueue = New-Sheet "Priority_Action_Queue"
$wsDQ = New-Sheet "Data_Quality"
$wsRTM = New-Sheet "Traceability_UAT"

# -----------------------------------------------------------------------------
# 3. ReadMe sheet
# -----------------------------------------------------------------------------
$wsReadMe.Cells.Item(1,1) = "Ridgeline Field Services -- Operational Efficiency KPI Dashboard"
$wsReadMe.Cells.Item(1,1).Font.Size = 16
$wsReadMe.Cells.Item(1,1).Font.Bold = $true

$readmeLines = @(
    "",
    "BUSINESS QUESTION",
    "Where is technician capacity being wasted, and which dispatch decisions should change this week?",
    "",
    "DATA DISCLOSURE",
    "Ridgeline Field Services is a fictional company. All data in this workbook is 100% synthetically",
    "generated by sql/02_generate_synthetic_data.sql against a local SQL Server instance. No confidential,",
    "proprietary, or production data of any real organization is used or represented.",
    "",
    "HOW THIS WORKBOOK WORKS",
    "This is an operational control, not a static report. Change any cell on the Targets sheet and every",
    "formula on KPI_Dashboard and Priority_Action_Queue recalculates against the new threshold. Change the",
    "Region / Year / Month filter cells on KPI_Dashboard to re-slice the whole dashboard.",
    "",
    "SHEET MAP",
    "  Raw_Data                Job-level fact extract (Excel Table tblJobs) -- the single source all formulas reference",
    "  Targets                 Editable KPI thresholds -- mirrors sql/01_create_schema.sql: dbo.Ref_KPITargets",
    "  KPI_Dashboard            Formula-driven KPI summary with conditional-formatting thresholds",
    "  Priority_Action_Queue   Ranked worklist of technicians needing dispatch review, generated by formula",
    "  Data_Quality            Formula-driven anomaly detection (mirrors sql/03_data_quality_checks.sql)",
    "  Traceability_UAT        Requirements traceability matrix + UAT test summary",
    "",
    "REUSABLE SQL",
    "Every number here can be reproduced independently from the SQL views/procs in /sql -- see",
    "dbo.vw_RegionKPIMonthly, dbo.vw_TechnicianKPIMonthly, dbo.vw_PriorityActionQueue, dbo.vw_DQ_Summary.",
    "",
    "Generated: $(Get-Date -Format 'yyyy-MM-dd')"
)
for ($i=0; $i -lt $readmeLines.Count; $i++) {
    $wsReadMe.Cells.Item($i+3, 1) = $readmeLines[$i]
}
$wsReadMe.Columns.Item(1).ColumnWidth = 110
$wsReadMe.Cells.Item(4,1).Font.Bold = $true
$wsReadMe.Cells.Item(7,1).Font.Bold = $true
$wsReadMe.Cells.Item(12,1).Font.Bold = $true
$wsReadMe.Cells.Item(17,1).Font.Bold = $true
$wsReadMe.Cells.Item(25,1).Font.Bold = $true

Write-Host "ReadMe sheet written."

# -----------------------------------------------------------------------------
# 4. Raw_Data sheet -- bulk array write, then convert to an Excel Table
# -----------------------------------------------------------------------------
Write-Host "Building Raw_Data array ($($jobsTable.Rows.Count) rows x $($jobsTable.Columns.Count) cols)..."
$nRows = $jobsTable.Rows.Count
$nCols = $jobsTable.Columns.Count

$headerArr = New-Object 'object[,]' 1,$nCols
for ($c=0; $c -lt $nCols; $c++) { $headerArr[0,$c] = $jobsTable.Columns[$c].ColumnName }

$dataArr = New-Object 'object[,]' $nRows,$nCols
for ($r=0; $r -lt $nRows; $r++) {
    $row = $jobsTable.Rows[$r]
    for ($c=0; $c -lt $nCols; $c++) {
        $v = $row[$c]
        if ($v -is [DBNull]) { $dataArr[$r,$c] = $null } else { $dataArr[$r,$c] = $v }
    }
}
Write-Host "Array built. Writing to worksheet..."

$wsRaw.Range($wsRaw.Cells.Item(1,1), $wsRaw.Cells.Item(1,$nCols)).Value2 = $headerArr
$wsRaw.Range($wsRaw.Cells.Item(2,1), $wsRaw.Cells.Item($nRows+1,$nCols)).Value2 = $dataArr

# Date/datetime number formats (cols: 2=JobDate, 15=ScheduledStart, 16=ActualStart, 17=ActualEnd)
$wsRaw.Range($wsRaw.Cells.Item(2,2),  $wsRaw.Cells.Item($nRows+1,2)).NumberFormat  = "yyyy-mm-dd"
$wsRaw.Range($wsRaw.Cells.Item(2,15), $wsRaw.Cells.Item($nRows+1,17)).NumberFormat = "yyyy-mm-dd hh:mm"
$wsRaw.Range($wsRaw.Cells.Item(2,23), $wsRaw.Cells.Item($nRows+1,24)).NumberFormat = "$#,##0.00"

$fullRange = $wsRaw.Range($wsRaw.Cells.Item(1,1), $wsRaw.Cells.Item($nRows+1,$nCols))
$tblJobs = $wsRaw.ListObjects.Add(1, $fullRange, $null, 1)   # xlSrcRange=1, xlYes=1
$tblJobs.Name = "tblJobs"
$wsRaw.Columns.AutoFit() | Out-Null
Write-Host "Raw_Data table 'tblJobs' created."

# -----------------------------------------------------------------------------
# 5. Targets sheet -- editable thresholds + named ranges
# -----------------------------------------------------------------------------
$order = @('UtilizationPct','UtilizationCeilingPct','FirstTimeFixPct','SLACompliancePct','CallbackPct')
$targetsByName = @{}
foreach ($row in $targetsTable.Rows) { $targetsByName[[string]$row.MetricName] = $row }

$wsTargets.Cells.Item(1,1) = "MetricName"
$wsTargets.Cells.Item(1,2) = "TargetValue"
$wsTargets.Cells.Item(1,3) = "WarningValue"
$wsTargets.Cells.Item(1,4) = "Direction"
$wsTargets.Cells.Item(1,5) = "Description"
$wsTargets.Range("A1:E1").Font.Bold = $true

$r = 2
foreach ($name in $order) {
    $row = $targetsByName[$name]
    $wsTargets.Cells.Item($r,1) = $name
    $wsTargets.Cells.Item($r,2) = [double]$row.TargetValue
    $wsTargets.Cells.Item($r,3) = [double]$row.WarningValue
    $wsTargets.Cells.Item($r,4) = [string]$row.Direction
    $wsTargets.Cells.Item($r,5) = [string]$row.Description
    $wb.Names.Add("Target_$name", "=Targets!`$B`$$r") | Out-Null
    $wb.Names.Add("Warn_$name",   "=Targets!`$C`$$r") | Out-Null
    $r++
}
$wsTargets.Columns.Item(5).ColumnWidth = 90
$wsTargets.Columns.AutoFit() | Out-Null
Write-Host "Targets sheet written with named ranges (Target_*, Warn_*)."

# -----------------------------------------------------------------------------
# 6. Priority_Action_Queue sheet
#    Helper table (one row per active technician) with formula-driven KPIs,
#    then a ranked Top-15 display built with LARGE + MATCH + INDEX (works in
#    every Excel version, no dynamic-array functions required).
# -----------------------------------------------------------------------------
$wsQueue.Cells.Item(1,1) = "Priority Action Queue -- Technicians Needing Dispatch Review"
$wsQueue.Cells.Item(1,1).Font.Size = 14
$wsQueue.Cells.Item(1,1).Font.Bold = $true
$wsQueue.Cells.Item(2,1) = "Uses the Year/Month filter set on KPI_Dashboard (cells B4:B5)."

$headerRow = 4
$firstDataRow = 5
$nTech = $techTable.Rows.Count
$lastDataRow = $firstDataRow + $nTech - 1

$queueHeaders = @("TechnicianID","TechnicianName","RegionName","SkillLevel","HireDate","UtilizationPct","FirstTimeFixPct","SLACompliancePct","CallbackPct","RiskScore","RecommendedAction","SortKey")
for ($c=0; $c -lt $queueHeaders.Count; $c++) { $wsQueue.Cells.Item($headerRow, $c+1) = $queueHeaders[$c] }
$wsQueue.Range($wsQueue.Cells.Item($headerRow,1), $wsQueue.Cells.Item($headerRow,$queueHeaders.Count)).Font.Bold = $true

# Bulk-write technician static columns A-E
$techArr = New-Object 'object[,]' $nTech,5
for ($i=0; $i -lt $nTech; $i++) {
    $tr = $techTable.Rows[$i]
    $techArr[$i,0] = [string]$tr.TechnicianID
    $techArr[$i,1] = [string]$tr.TechnicianName
    $techArr[$i,2] = [string]$tr.RegionName
    $techArr[$i,3] = [string]$tr.SkillLevel
    $techArr[$i,4] = [datetime]$tr.HireDate
}
$wsQueue.Range($wsQueue.Cells.Item($firstDataRow,1), $wsQueue.Cells.Item($lastDataRow,5)).Value2 = $techArr
$wsQueue.Range($wsQueue.Cells.Item($firstDataRow,5), $wsQueue.Cells.Item($lastDataRow,5)).NumberFormat = "yyyy-mm-dd"

# Formula columns F-L, one string per row (built in PowerShell so ranges are already resolved)
$formulaArr = New-Object 'object[,]' $nTech,7
for ($i=0; $i -lt $nTech; $i++) {
    $rr = $firstDataRow + $i
    $formulaArr[$i,0] = "=IFERROR((SUMIFS(tblJobs[JobDurationMin],tblJobs[TechnicianID],`$A$rr,tblJobs[Year],KPI_Dashboard!`$B`$5,tblJobs[Month],KPI_Dashboard!`$B`$6,tblJobs[JobStatus],`"Completed`")+SUMIFS(tblJobs[TravelTimeMin],tblJobs[TechnicianID],`$A$rr,tblJobs[Year],KPI_Dashboard!`$B`$5,tblJobs[Month],KPI_Dashboard!`$B`$6,tblJobs[JobStatus],`"Completed`"))/(MAX(1,NETWORKDAYS(MAX(`$E$rr,DATE(KPI_Dashboard!`$B`$5,KPI_Dashboard!`$B`$6,1)),EOMONTH(DATE(KPI_Dashboard!`$B`$5,KPI_Dashboard!`$B`$6,1),0)))*480),0)"
    $formulaArr[$i,1] = "=IFERROR(AVERAGEIFS(tblJobs[FirstTimeFix],tblJobs[TechnicianID],`$A$rr,tblJobs[Year],KPI_Dashboard!`$B`$5,tblJobs[Month],KPI_Dashboard!`$B`$6,tblJobs[JobStatus],`"Completed`"),`"`")"
    $formulaArr[$i,2] = "=IFERROR(AVERAGEIFS(tblJobs[SLAMet],tblJobs[TechnicianID],`$A$rr,tblJobs[Year],KPI_Dashboard!`$B`$5,tblJobs[Month],KPI_Dashboard!`$B`$6,tblJobs[JobStatus],`"Completed`"),`"`")"
    $formulaArr[$i,3] = "=IFERROR(AVERAGEIFS(tblJobs[CallbackFlag],tblJobs[TechnicianID],`$A$rr,tblJobs[Year],KPI_Dashboard!`$B`$5,tblJobs[Month],KPI_Dashboard!`$B`$6,tblJobs[JobStatus],`"Completed`"),`"`")"
    $formulaArr[$i,4] = "=IF(F$rr<Warn_UtilizationPct/100,1,0)+IF(F$rr>Warn_UtilizationCeilingPct/100,1,0)+IF(AND(G$rr<>`"`",G$rr<Warn_FirstTimeFixPct/100),1,0)+IF(AND(H$rr<>`"`",H$rr<Warn_SLACompliancePct/100),1,0)+IF(AND(I$rr<>`"`",I$rr>Warn_CallbackPct/100),1,0)"
    $formulaArr[$i,5] = "=TRIM(IF(F$rr<Warn_UtilizationPct/100,`"Underutilized -- reassign jobs from over-capacity peers. `",`"`")&IF(F$rr>Warn_UtilizationCeilingPct/100,`"Over capacity -- redistribute jobs to underutilized techs. `",`"`")&IF(AND(G$rr<>`"`",G$rr<Warn_FirstTimeFixPct/100),`"Low first-time-fix -- pair with a Master technician. `",`"`")&IF(AND(H$rr<>`"`",H$rr<Warn_SLACompliancePct/100),`"SLA below target -- review dispatch routing. `",`"`")&IF(AND(I$rr<>`"`",I$rr>Warn_CallbackPct/100),`"High callback -- audit rework causes. `",`"`"))"
    $formulaArr[$i,6] = "=J$rr+ROW()/100000"
}
$wsQueue.Range($wsQueue.Cells.Item($firstDataRow,6), $wsQueue.Cells.Item($lastDataRow,12)).Formula = $formulaArr
$wsQueue.Range($wsQueue.Cells.Item($firstDataRow,6), $wsQueue.Cells.Item($lastDataRow,9)).NumberFormat = "0.0%"

# Ranked Top-15 display block (columns N-V), built with LARGE/MATCH/INDEX
$topN = 15
$rankHeaders = @("Rank","TechnicianID","TechnicianName","RegionName","UtilizationPct","FirstTimeFixPct","SLACompliancePct","CallbackPct","RiskScore","RecommendedAction")
$wsQueue.Cells.Item($headerRow,14) = "TOP $topN PRIORITY QUEUE (highest RiskScore first)"
$wsQueue.Cells.Item($headerRow,14).Font.Bold = $true
for ($c=0; $c -lt $rankHeaders.Count; $c++) { $wsQueue.Cells.Item($headerRow+1, 14+$c) = $rankHeaders[$c] }
$wsQueue.Range($wsQueue.Cells.Item($headerRow+1,14), $wsQueue.Cells.Item($headerRow+1,14+$rankHeaders.Count-1)).Font.Bold = $true

$rankArr = New-Object 'object[,]' $topN,10
for ($k=1; $k -le $topN; $k++) {
    $ri = $k - 1
    $rankArr[$ri,0] = $k
    $rankArr[$ri,1] = "=IFERROR(INDEX(`$A`$${firstDataRow}:`$A`$${lastDataRow},MATCH(LARGE(`$L`$${firstDataRow}:`$L`$${lastDataRow},$k),`$L`$${firstDataRow}:`$L`$${lastDataRow},0)),`"`")"
    $rankArr[$ri,2] = "=IFERROR(INDEX(`$B`$${firstDataRow}:`$B`$${lastDataRow},MATCH(LARGE(`$L`$${firstDataRow}:`$L`$${lastDataRow},$k),`$L`$${firstDataRow}:`$L`$${lastDataRow},0)),`"`")"
    $rankArr[$ri,3] = "=IFERROR(INDEX(`$C`$${firstDataRow}:`$C`$${lastDataRow},MATCH(LARGE(`$L`$${firstDataRow}:`$L`$${lastDataRow},$k),`$L`$${firstDataRow}:`$L`$${lastDataRow},0)),`"`")"
    $rankArr[$ri,4] = "=IFERROR(INDEX(`$F`$${firstDataRow}:`$F`$${lastDataRow},MATCH(LARGE(`$L`$${firstDataRow}:`$L`$${lastDataRow},$k),`$L`$${firstDataRow}:`$L`$${lastDataRow},0)),`"`")"
    $rankArr[$ri,5] = "=IFERROR(INDEX(`$G`$${firstDataRow}:`$G`$${lastDataRow},MATCH(LARGE(`$L`$${firstDataRow}:`$L`$${lastDataRow},$k),`$L`$${firstDataRow}:`$L`$${lastDataRow},0)),`"`")"
    $rankArr[$ri,6] = "=IFERROR(INDEX(`$H`$${firstDataRow}:`$H`$${lastDataRow},MATCH(LARGE(`$L`$${firstDataRow}:`$L`$${lastDataRow},$k),`$L`$${firstDataRow}:`$L`$${lastDataRow},0)),`"`")"
    $rankArr[$ri,7] = "=IFERROR(INDEX(`$I`$${firstDataRow}:`$I`$${lastDataRow},MATCH(LARGE(`$L`$${firstDataRow}:`$L`$${lastDataRow},$k),`$L`$${firstDataRow}:`$L`$${lastDataRow},0)),`"`")"
    $rankArr[$ri,8] = "=IFERROR(INDEX(`$J`$${firstDataRow}:`$J`$${lastDataRow},MATCH(LARGE(`$L`$${firstDataRow}:`$L`$${lastDataRow},$k),`$L`$${firstDataRow}:`$L`$${lastDataRow},0)),`"`")"
    $rankArr[$ri,9] = "=IFERROR(INDEX(`$K`$${firstDataRow}:`$K`$${lastDataRow},MATCH(LARGE(`$L`$${firstDataRow}:`$L`$${lastDataRow},$k),`$L`$${firstDataRow}:`$L`$${lastDataRow},0)),`"`")"
}
$wsQueue.Range($wsQueue.Cells.Item($headerRow+2,14), $wsQueue.Cells.Item($headerRow+1+$topN,23)).Formula = $rankArr
$wsQueue.Range($wsQueue.Cells.Item($headerRow+2,18), $wsQueue.Cells.Item($headerRow+1+$topN,21)).NumberFormat = "0.0%"

# Conditional formatting on RiskScore column of the ranked block (>=2 red, =1 amber)
$riskRange = $wsQueue.Range($wsQueue.Cells.Item($headerRow+2,22), $wsQueue.Cells.Item($headerRow+1+$topN,22))
$fc1 = $riskRange.FormatConditions.Add(1, 3, "=2")   # xlCellValue=1, xlGreaterEqual=3
$fc1.Interior.Color = 0x9999FF   # light red (BGR)
$wsQueue.Columns.Item(23).ColumnWidth = 70
$wsQueue.Columns.AutoFit() | Out-Null
Write-Host "Priority_Action_Queue sheet written ($nTech technicians, top $topN ranked)."

# -----------------------------------------------------------------------------
# 7. KPI_Dashboard sheet
# -----------------------------------------------------------------------------
$wsDash.Cells.Item(1,1) = "Ridgeline Field Services -- Regional Dispatch KPI Dashboard"
$wsDash.Cells.Item(1,1).Font.Size = 14
$wsDash.Cells.Item(1,1).Font.Bold = $true
$wsDash.Cells.Item(2,1) = "Business question: Where is technician capacity being wasted, and which dispatch decisions should change this week?"

$wsDash.Cells.Item(4,1) = "Region:"
$wsDash.Cells.Item(4,2) = "All"
$wsDash.Cells.Item(5,1) = "Year:"
$wsDash.Cells.Item(5,2) = 2025
$wsDash.Cells.Item(6,1) = "Month:"
$wsDash.Cells.Item(6,2) = 12
$wsDash.Range("A4:A6").Font.Bold = $true
$wsDash.Range("B4:B6").Interior.Color = 0xCCFFFF  # highlight the input cells (yellow-ish BGR)

$wb.Names.Add("FilterRegion", "=KPI_Dashboard!`$B`$4") | Out-Null
$wb.Names.Add("FilterYear",   "=KPI_Dashboard!`$B`$5") | Out-Null
$wb.Names.Add("FilterMonth",  "=KPI_Dashboard!`$B`$6") | Out-Null

# Data validation dropdowns
$wsDash.Range("B4").Validation.Delete()
$wsDash.Range("B4").Validation.Add(3, 1, 1, "All,North Metro,South Metro,East Valley,West Hills,Central District") | Out-Null
$wsDash.Range("B5").Validation.Delete()
$wsDash.Range("B5").Validation.Add(3, 1, 1, "2024,2025") | Out-Null
$wsDash.Range("B6").Validation.Delete()
$wsDash.Range("B6").Validation.Add(3, 1, 1, "1,2,3,4,5,6,7,8,9,10,11,12") | Out-Null

$dashHeaderRow = 8
$dashHeaders = @("Metric","Value","Target","Warning","Status")
for ($c=0; $c -lt $dashHeaders.Count; $c++) { $wsDash.Cells.Item($dashHeaderRow,$c+1) = $dashHeaders[$c] }
$wsDash.Range($wsDash.Cells.Item($dashHeaderRow,1), $wsDash.Cells.Item($dashHeaderRow,5)).Font.Bold = $true

$regionCrit = "IF(`$B`$4=`"All`",`"*`",`$B`$4)"
$dashRows = @(
    @{ Name="Total Jobs"; Formula="=COUNTIFS(tblJobs[RegionName],$regionCrit,tblJobs[Year],`$B`$5,tblJobs[Month],`$B`$6)"; Target=""; Warn=""; Fmt="#,##0" },
    @{ Name="Completed Jobs"; Formula="=COUNTIFS(tblJobs[RegionName],$regionCrit,tblJobs[Year],`$B`$5,tblJobs[Month],`$B`$6,tblJobs[JobStatus],`"Completed`")"; Target=""; Warn=""; Fmt="#,##0" },
    @{ Name="Utilization %"; Formula="=IFERROR((SUMIFS(tblJobs[JobDurationMin],tblJobs[RegionName],$regionCrit,tblJobs[Year],`$B`$5,tblJobs[Month],`$B`$6,tblJobs[JobStatus],`"Completed`")+SUMIFS(tblJobs[TravelTimeMin],tblJobs[RegionName],$regionCrit,tblJobs[Year],`$B`$5,tblJobs[Month],`$B`$6,tblJobs[JobStatus],`"Completed`"))/(NETWORKDAYS(DATE(`$B`$5,`$B`$6,1),EOMONTH(DATE(`$B`$5,`$B`$6,1),0))*480*COUNTIFS(Priority_Action_Queue!`$C`$${firstDataRow}:`$C`$${lastDataRow},$regionCrit)),0)"; Target="=Target_UtilizationPct/100"; Warn="=Warn_UtilizationPct/100"; Fmt="0.0%" },
    @{ Name="First-Time-Fix %"; Formula="=IFERROR(AVERAGEIFS(tblJobs[FirstTimeFix],tblJobs[RegionName],$regionCrit,tblJobs[Year],`$B`$5,tblJobs[Month],`$B`$6,tblJobs[JobStatus],`"Completed`"),0)"; Target="=Target_FirstTimeFixPct/100"; Warn="=Warn_FirstTimeFixPct/100"; Fmt="0.0%" },
    @{ Name="SLA Compliance %"; Formula="=IFERROR(AVERAGEIFS(tblJobs[SLAMet],tblJobs[RegionName],$regionCrit,tblJobs[Year],`$B`$5,tblJobs[Month],`$B`$6,tblJobs[JobStatus],`"Completed`"),0)"; Target="=Target_SLACompliancePct/100"; Warn="=Warn_SLACompliancePct/100"; Fmt="0.0%" },
    @{ Name="Callback %"; Formula="=IFERROR(AVERAGEIFS(tblJobs[CallbackFlag],tblJobs[RegionName],$regionCrit,tblJobs[Year],`$B`$5,tblJobs[Month],`$B`$6,tblJobs[JobStatus],`"Completed`"),0)"; Target="=Target_CallbackPct/100"; Warn="=Warn_CallbackPct/100"; Fmt="0.0%" },
    @{ Name="Total Revenue"; Formula="=SUMIFS(tblJobs[JobRevenue],tblJobs[RegionName],$regionCrit,tblJobs[Year],`$B`$5,tblJobs[Month],`$B`$6,tblJobs[JobStatus],`"Completed`")"; Target=""; Warn=""; Fmt="$#,##0" },
    @{ Name="Total Cost"; Formula="=SUMIFS(tblJobs[JobCost],tblJobs[RegionName],$regionCrit,tblJobs[Year],`$B`$5,tblJobs[Month],`$B`$6)"; Target=""; Warn=""; Fmt="$#,##0" },
    @{ Name="Gross Margin"; Formula="=C9-C10"; Target=""; Warn=""; Fmt="$#,##0" }
)

$dr = $dashHeaderRow + 1
foreach ($row in $dashRows) {
    $wsDash.Cells.Item($dr,1) = $row.Name
    $wsDash.Cells.Item($dr,2).Formula = $row.Formula
    $wsDash.Cells.Item($dr,2).NumberFormat = $row.Fmt
    if ($row.Target -ne "") { $wsDash.Cells.Item($dr,3).Formula = $row.Target; $wsDash.Cells.Item($dr,3).NumberFormat = $row.Fmt }
    if ($row.Warn -ne "")   { $wsDash.Cells.Item($dr,4).Formula = $row.Warn;   $wsDash.Cells.Item($dr,4).NumberFormat = $row.Fmt }
    $dr++
}
# Fix Gross Margin formula to reference actual rows (Total Revenue / Total Cost rows)
$revRow = $dashHeaderRow + 1 + 6   # 7th data row (index 6) = Total Revenue
$costRow = $dashHeaderRow + 1 + 7  # 8th data row = Total Cost
$marginRow = $dashHeaderRow + 1 + 8
$wsDash.Cells.Item($marginRow,2).Formula = "=B$revRow-B$costRow"
$wsDash.Cells.Item($marginRow,2).NumberFormat = "$#,##0"

# Conditional formatting: Value vs Warning for the 4 rate KPIs (Utilization, FTF, SLA, Callback)
$utilRow = $dashHeaderRow + 3
$ftfRow = $dashHeaderRow + 4
$slaRow = $dashHeaderRow + 5
$cbRow = $dashHeaderRow + 6
foreach ($rowNum in @($utilRow,$ftfRow,$slaRow)) {
    $cell = $wsDash.Range($wsDash.Cells.Item($rowNum,2), $wsDash.Cells.Item($rowNum,2))
    $fcGood = $cell.FormatConditions.Add(1, 3, "=`$D`$$rowNum")  # >= warning
    $fcGood.Interior.Color = 0xCCFFCC
    $fcBad = $cell.FormatConditions.Add(1, 6, "=`$D`$$rowNum")  # < warning (xlLess=6)
    $fcBad.Interior.Color = 0x9999FF
}
# Callback is LowerBetter: bad when ABOVE warning
$cbCell = $wsDash.Range($wsDash.Cells.Item($cbRow,2), $wsDash.Cells.Item($cbRow,2))
$fcGoodCB = $cbCell.FormatConditions.Add(1, 6, "=`$D`$$cbRow")   # < warning = good
$fcGoodCB.Interior.Color = 0xCCFFCC
$fcBadCB = $cbCell.FormatConditions.Add(1, 3, "=`$D`$$cbRow")    # >= warning = bad
$fcBadCB.Interior.Color = 0x9999FF

$wsDash.Columns.AutoFit() | Out-Null
Write-Host "KPI_Dashboard sheet written."

# -----------------------------------------------------------------------------
# 8. Data_Quality sheet -- independent formula re-implementation of the SQL
#    checks in sql/03_data_quality_checks.sql, plus a side-by-side SQL
#    cross-check so the two independent implementations can be compared.
# -----------------------------------------------------------------------------
$wsDQ.Cells.Item(1,1) = "Data Quality Controls"
$wsDQ.Cells.Item(1,1).Font.Size = 14
$wsDQ.Cells.Item(1,1).Font.Bold = $true
$wsDQ.Cells.Item(2,1) = "Formula-driven re-implementation of sql/03_data_quality_checks.sql, cross-checked against the SQL-computed values."

$dqHeaders = @("AnomalyType","Description","Excel Count","Excel Rate %","SQL Count (cross-check)","SQL Rate %","QA Result")
for ($c=0; $c -lt $dqHeaders.Count; $c++) { $wsDQ.Cells.Item(4,$c+1) = $dqHeaders[$c] }
$wsDQ.Range($wsDQ.Cells.Item(4,1), $wsDQ.Cells.Item(4,7)).Font.Bold = $true

$dqByType = @{}
foreach ($row in $dqTable.Rows) { $dqByType[[string]$row.AnomalyType] = $row }

$totalRowsFormula = "COUNTA(tblJobs[JobID])"
$dqRows = @(
    @{ Type="TIMESTAMP_INVERSION"; Desc="ActualEnd earlier than ActualStart (both populated)";
       Formula="=SUMPRODUCT((tblJobs[ActualStart]<>`"`")*(tblJobs[ActualEnd]<>`"`")*(tblJobs[ActualEnd]<tblJobs[ActualStart]))" },
    @{ Type="MISSING_ACTUALS"; Desc="Completed job with no ActualStart/ActualEnd";
       Formula="=SUMPRODUCT((tblJobs[JobStatus]=`"Completed`")*(((tblJobs[ActualStart]=`"`")+(tblJobs[ActualEnd]=`"`"))>0))" },
    @{ Type="EXTREME_DURATION"; Desc="JobDurationMin > 600 minutes";
       Formula="=COUNTIFS(tblJobs[JobDurationMin],`">600`")" }
)

$dqr = 5
foreach ($row in $dqRows) {
    $wsDQ.Cells.Item($dqr,1) = $row.Type
    $wsDQ.Cells.Item($dqr,2) = $row.Desc
    $wsDQ.Cells.Item($dqr,3).Formula = $row.Formula
    $wsDQ.Cells.Item($dqr,4).Formula = "=C$dqr/$totalRowsFormula*100"
    $wsDQ.Cells.Item($dqr,4).NumberFormat = "0.000"
    $sqlRow = $dqByType[$row.Type]
    $wsDQ.Cells.Item($dqr,5) = [int]$sqlRow.AnomalyCount
    $wsDQ.Cells.Item($dqr,6) = [double]$sqlRow.AnomalyRatePct
    $wsDQ.Cells.Item($dqr,6).NumberFormat = "0.000"
    $wsDQ.Cells.Item($dqr,7).Formula = "=IF(C$dqr=E$dqr,`"MATCH`",`"MISMATCH -- investigate`")"
    $dqr++
}
$wsDQ.Columns.Item(2).ColumnWidth = 55
$wsDQ.Columns.AutoFit() | Out-Null
Write-Host "Data_Quality sheet written."

# -----------------------------------------------------------------------------
# 9. Traceability_UAT sheet
# -----------------------------------------------------------------------------
$wsRTM.Cells.Item(1,1) = "Requirements Traceability Matrix + UAT Summary"
$wsRTM.Cells.Item(1,1).Font.Size = 14
$wsRTM.Cells.Item(1,1).Font.Bold = $true
$wsRTM.Cells.Item(2,1) = "Full detail: /docs/requirements_traceability_matrix.md. Executable proof: sql/06_uat_test_cases.sql."

$rtmHeaders = @("Req ID","Business Requirement","KPI / Field","Implementation","Validation / Test","UAT Result")
for ($c=0; $c -lt $rtmHeaders.Count; $c++) { $wsRTM.Cells.Item(4,$c+1) = $rtmHeaders[$c] }
$wsRTM.Range($wsRTM.Cells.Item(4,1), $wsRTM.Cells.Item(4,6)).Font.Bold = $true

$rtmData = @(
    @("BR-01","Identify regions/technicians with wasted (underutilized) capacity","UtilizationPct","vw_TechnicianUtilization; KPI_Dashboard","UAT-01: 255/11040 min = 2.31% matches exactly","PASS"),
    @("BR-02","Flag technicians who are over capacity (burnout/SLA risk)","FlagOverCapacity","vw_PriorityActionQueue; Priority_Action_Queue sheet","Manually verified against technicians exceeding 95% util","PASS"),
    @("BR-03","Quantify rework/quality risk","FirstTimeFixPct, CallbackPct","vw_TechnicianKPIMonthly","UAT-02a/02b: 1 of 2 jobs each -> 50.00% exactly","PASS"),
    @("BR-04","Quantify SLA compliance and its driver (travel/dispatch delay)","SLACompliancePct","vw_RegionKPIMonthly; drill by SLAHours","East Valley 93.1% vs Central 95.6% on 1-hr SLA segment reconciles to raw data","PASS"),
    @("BR-05","Produce a ranked, actionable worklist","RiskScore, RecommendedAction","vw_PriorityActionQueue; ranked Top-15 block","UAT-04: engineered low-utilization tech correctly flagged","PASS"),
    @("BR-06","No confidential/production data used","N/A (governance)","sql/02_generate_synthetic_data.sql header disclosure","Manual review: no real identifiers anywhere","PASS"),
    @("BR-07","Detect data quality problems before they reach the dashboard","vw_DQ_JobAnomalies","sql/03_data_quality_checks.sql; Data_Quality sheet","UAT-03: deliberately broken row caught by TIMESTAMP_INVERSION","PASS"),
    @("BR-08","Every KPI reproducible outside the BI tool","N/A (architecture)","usp_GetRegionKPISummary, usp_GetPriorityActionQueue, usp_GetTechnicianScorecard","Manually executed each proc, output matches views","PASS"),
    @("BR-09","Region rollups reconcile to job-level detail","TotalJobs","vw_RegionKPIMonthly","UAT-05: view count = raw COUNT(*) for the same scope (exact match)","PASS"),
    @("BR-10","A technician's monthly KPIs must not leak data from other months","FirstTimeFixPct/SLACompliancePct/CallbackPct (technician grain)","vw_TechnicianKPIMonthly","UAT-06a/06b: caught by Excel/SQL cross-check disagreeing -- see data_validation_report.md Sec.5","PASS")
)
$rtmArr = New-Object 'object[,]' $rtmData.Count,6
for ($i=0; $i -lt $rtmData.Count; $i++) { for ($c=0; $c -lt 6; $c++) { $rtmArr[$i,$c] = $rtmData[$i][$c] } }
$wsRTM.Range($wsRTM.Cells.Item(5,1), $wsRTM.Cells.Item(4+$rtmData.Count,6)).Value2 = $rtmArr

$uatStartRow = 5 + $rtmData.Count + 2
$wsRTM.Cells.Item($uatStartRow,1) = "UAT Test Case Log (from sql/06_uat_test_cases.sql -- transactional, rolled back after each run)"
$wsRTM.Cells.Item($uatStartRow,1).Font.Bold = $true
$uatHeaders = @("Test ID","Description","Expected","Actual","Result")
for ($c=0; $c -lt $uatHeaders.Count; $c++) { $wsRTM.Cells.Item($uatStartRow+1,$c+1) = $uatHeaders[$c] }
$wsRTM.Range($wsRTM.Cells.Item($uatStartRow+1,1), $wsRTM.Cells.Item($uatStartRow+1,5)).Font.Bold = $true
$uatData = @(
    @("UAT-01","Utilization % matches hand calculation (255/11040 min)","2.31","2.31","PASS"),
    @("UAT-02a","First-Time-Fix % = 1 of 2 completed jobs","50.00","50.00","PASS"),
    @("UAT-02b","Callback % = 1 of 2 completed jobs","50.00","50.00","PASS"),
    @("UAT-03","DQ view flags a job with ActualEnd < ActualStart","1","1","PASS"),
    @("UAT-04","Priority queue flags a technician far below utilization target","1","1","PASS"),
    @("UAT-05","Region monthly job count reconciles to underlying fact rows","matches","matches","PASS"),
    @("UAT-06a","Jan-2025 FTF% = 2 of its own 3 jobs, not pooled with Feb (regression test)","66.67","66.67","PASS"),
    @("UAT-06b","Feb-2025 FTF% = 1 of its own 1 job, not pooled with Jan (regression test)","100.00","100.00","PASS")
)
$uatArr = New-Object 'object[,]' $uatData.Count,5
for ($i=0; $i -lt $uatData.Count; $i++) { for ($c=0; $c -lt 5; $c++) { $uatArr[$i,$c] = $uatData[$i][$c] } }
$wsRTM.Range($wsRTM.Cells.Item($uatStartRow+2,1), $wsRTM.Cells.Item($uatStartRow+1+$uatData.Count,5)).Value2 = $uatArr

$wsRTM.Columns.Item(2).ColumnWidth = 55
$wsRTM.Columns.Item(4).ColumnWidth = 45
$wsRTM.Columns.Item(5).ColumnWidth = 45
$wsRTM.Columns.AutoFit() | Out-Null
Write-Host "Traceability_UAT sheet written."

# -----------------------------------------------------------------------------
# 10. Finalize: recalculate, order sheets, save, quit
# -----------------------------------------------------------------------------
$wsReadMe.Activate()
$wb.Worksheets.Item("ReadMe").Move($wb.Worksheets.Item(1))
$excel.CalculateFullRebuild()
if (Test-Path $OutPath) { Remove-Item $OutPath -Force }
# Strip author / last-modified-by from the saved file (Excel otherwise stamps the local Office user's name)
$wb.RemovePersonalInformation = $true
$wb.SaveAs($OutPath, 51)
if (-not (Test-Path $OutPath)) { throw "SaveAs did not produce a file at $OutPath -- aborting." }
$savedAt = (Get-Item $OutPath).LastWriteTime
if ((Get-Date) - $savedAt -gt [TimeSpan]::FromSeconds(30)) { throw "File at $OutPath was not just written (LastWriteTime=$savedAt) -- SaveAs likely failed silently." }
Write-Host "Final save complete: $OutPath (verified fresh write at $savedAt)"

$excel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()
Write-Host "Excel closed cleanly."
