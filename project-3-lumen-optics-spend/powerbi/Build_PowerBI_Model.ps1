<#
=============================================================================
Project 3 -- Lumen Optics Manufacturing: Photonics Spend Scorecard
Script:  Build_PowerBI_Model.ps1
Purpose: Script the Power BI semantic model -- tables, relationships,
         hierarchies and every DAX measure -- into a running Power BI Desktop
         session via its local Analysis Services engine, the same mechanism
         Tabular Editor and DAX Studio use.

WHY SCRIPT THE MODEL AT ALL
    Sixty-nine measures clicked into a UI cannot be reviewed, diffed or
    rebuilt. Scripted, they live in source control beside the SQL that feeds
    them, and a reviewer can see that the leverage weights in DAX are the same
    weights the SQL uses -- which is the only way the Power BI report and the
    Excel workbook can be trusted to rank the same negotiations identically.

    Report PAGES still have to be built by hand: there is no supported
    scripting path for visuals, and the one time this project edited a .pbix
    report layer directly it corrupted the file.

HOW TO RUN
    1. Open Power BI Desktop on a blank report and leave it open.
    2. Run this script.
    3. The FIRST table refresh raises a credential prompt inside Desktop.
       Watch the Desktop window and click through it -- choose Windows /
       Integrated authentication and any privacy level. The script will appear
       to hang until you do. This happens once per Desktop session.
    4. Save the file as powerbi/Lumen_Spend_Scorecard.pbix.

The Analysis Services client DLLs are downloaded from nuget.org on first run
(official Microsoft package, ~6MB) into .amo_libs/ next to this script. They
are not vendored into the repo to keep it lean, and .gitignore excludes them.

DATA DISCLOSURE: Lumen Optics Manufacturing is fictional; all data is
synthetic. No confidential data and no production system is involved.
=============================================================================
#>

$ErrorActionPreference = "Stop"

$SqlServerInst = "localhost\TEW_SQLEXPRESS"
$SqlDb         = "LumenSpend"
$AsOf          = "2025-12-31"

# -----------------------------------------------------------------------------
# 0. Analysis Services client libraries
# -----------------------------------------------------------------------------
$libDir = Join-Path $PSScriptRoot ".amo_libs"
$dllDir = Join-Path $libDir "lib\net45"
if (-not (Test-Path (Join-Path $dllDir "Microsoft.AnalysisServices.Tabular.dll"))) {
    Write-Host "Downloading Microsoft.AnalysisServices.retail.amd64 from nuget.org..." -ForegroundColor Yellow
    New-Item -ItemType Directory -Force -Path $libDir | Out-Null
    $zipPath = Join-Path $libDir "amo.nupkg.zip"
    Invoke-WebRequest -Uri "https://www.nuget.org/api/v2/package/Microsoft.AnalysisServices.retail.amd64/19.84.1" -OutFile $zipPath
    Expand-Archive -Path $zipPath -DestinationPath $libDir -Force
    Remove-Item $zipPath
}
Add-Type -Path (Join-Path $dllDir "Microsoft.AnalysisServices.Core.dll")
Add-Type -Path (Join-Path $dllDir "Microsoft.AnalysisServices.dll")
Add-Type -Path (Join-Path $dllDir "Microsoft.AnalysisServices.Tabular.dll")

# -----------------------------------------------------------------------------
# 1. Find the running Desktop instance's engine port
# -----------------------------------------------------------------------------
$msmdsrv = Get-Process -Name "msmdsrv" -ErrorAction SilentlyContinue
if (-not $msmdsrv) { throw "No Power BI Desktop session found. Open Power BI Desktop on a blank report first." }
$conn = Get-NetTCPConnection -OwningProcess $msmdsrv.Id -State Listen -ErrorAction SilentlyContinue | Select-Object -First 1
if (-not $conn) { throw "Could not find the local Analysis Services engine's listening port." }
$port = $conn.LocalPort
Write-Host "Found Power BI Desktop's local engine on port $port" -ForegroundColor Cyan

$server = New-Object Microsoft.AnalysisServices.Tabular.Server
$server.Connect("localhost:$port")
$db = $server.Databases[0]
$model = $db.Model
Write-Host "Connected to workspace database $($db.Name)"
Write-Host ("-" * 78)

# -----------------------------------------------------------------------------
# 2. Helpers
# -----------------------------------------------------------------------------
function New-Col($name, $type, [switch]$hidden) {
    $c = New-Object Microsoft.AnalysisServices.Tabular.DataColumn
    $c.Name = $name
    $c.SourceColumn = $name
    $c.DataType = [Microsoft.AnalysisServices.Tabular.DataType]::$type
    if ($hidden) { $c.IsHidden = $true }
    return $c
}

function New-MTable($name, $query) {
    if ($model.Tables.ContainsName($name)) { $model.Tables.Remove($name) | Out-Null; $model.SaveChanges() | Out-Null }
    $t = New-Object Microsoft.AnalysisServices.Tabular.Table
    $t.Name = $name
    $mExpr = "let`n    Source = Sql.Database(""$SqlServerInst"", ""$SqlDb"", [Query=""$($query -replace '"','""')""])`nin`n    Source"
    $p = New-Object Microsoft.AnalysisServices.Tabular.Partition
    $p.Name = "$name-Partition"
    $mps = New-Object Microsoft.AnalysisServices.Tabular.MPartitionSource
    $mps.Expression = $mExpr
    $p.Source = $mps
    $t.Partitions.Add($p) | Out-Null
    return $t
}

function New-Measure($tbl, $name, $expr, $fmt, $folder) {
    if ($tbl.Measures.ContainsName($name)) { $tbl.Measures.Remove($name) | Out-Null }
    $m = New-Object Microsoft.AnalysisServices.Tabular.Measure
    $m.Name = $name
    $m.Expression = $expr
    if ($fmt)    { $m.FormatString = $fmt }
    if ($folder) { $m.DisplayFolder = $folder }
    $tbl.Measures.Add($m) | Out-Null
}

function New-Rel($fromTable, $fromCol, $toTable, $toCol, [switch]$inactive) {
    $r = New-Object Microsoft.AnalysisServices.Tabular.SingleColumnRelationship
    $r.FromColumn = $model.Tables[$fromTable].Columns[$fromCol]   # many side
    $r.ToColumn   = $model.Tables[$toTable].Columns[$toCol]       # one side
    $r.FromCardinality = [Microsoft.AnalysisServices.Tabular.RelationshipEndCardinality]::Many
    $r.ToCardinality   = [Microsoft.AnalysisServices.Tabular.RelationshipEndCardinality]::One
    if ($inactive) { $r.IsActive = $false }
    $model.Relationships.Add($r) | Out-Null
}

# -----------------------------------------------------------------------------
# 3. Tables
#
# The facts are loaded from the SQL VIEWS, not the raw tables. The landed cost
# identity, the dated contract resolution and the erosion model are already
# defined once in SQL and asserted by 21 acceptance tests; re-deriving them in
# DAX would create a second definition that can drift from the first. DAX does
# the aggregation, the ratios, the time intelligence and the conditional
# formatting -- the things a semantic model is actually for.
# -----------------------------------------------------------------------------
Write-Host "Creating tables..."

# Drop every relationship BEFORE touching the tables. Removing a table does NOT
# cascade to the relationships that reference it -- TOM leaves them dangling and
# then refuses the save with "Relationship 'Relationship' points to deleted
# table 'Dim_Date'", which only ever appears on a re-run, never on a first run.
# Clearing here also makes relationship creation idempotent, so a re-run cannot
# stack a second copy of every filter path.
# NB: RelationshipCollection exposes Remove(object), not RemoveAt(index) -- and
# PowerShell's fallback for a missing method is to enumerate the collection and
# try the method on each MEMBER, so RemoveAt fails with a message naming
# SingleColumnRelationship rather than the collection. Snapshot with @() first:
# removing from a collection while enumerating it is its own bug.
foreach ($rel in @($model.Relationships)) { $model.Relationships.Remove($rel) | Out-Null }
$model.SaveChanges() | Out-Null

$tDate = New-MTable "Dim_Date" @"
SELECT DateKey, [Date], [Year], [Quarter], [Month], MonthName, YearMonth, YearQuarter,
       MonthEndDate, IsMonthEnd, IsWeekend
FROM dbo.Dim_Date WHERE [Date] BETWEEN '2023-01-01' AND '2026-06-30'
"@
@("DateKey|Int64","Date|DateTime","Year|Int64","Quarter|Int64","Month|Int64","MonthName|String",
  "YearMonth|String","YearQuarter|String","MonthEndDate|DateTime","IsMonthEnd|Boolean","IsWeekend|Boolean") |
  ForEach-Object { $p = $_ -split '\|'; $tDate.Columns.Add((New-Col $p[0] $p[1])) | Out-Null }
$model.Tables.Add($tDate) | Out-Null

$tVendor = New-MTable "Dim_Vendor" @"
SELECT VendorKey, VendorID, VendorName, Country, VendorTier, RelationshipStartDate,
       LumenRevenueSharePct
FROM dbo.Dim_Vendor
"@
@("VendorKey|Int64","VendorID|String","VendorName|String","Country|String","VendorTier|String",
  "RelationshipStartDate|DateTime","LumenRevenueSharePct|Double") |
  ForEach-Object { $p = $_ -split '\|'; $tVendor.Columns.Add((New-Col $p[0] $p[1])) | Out-Null }
$model.Tables.Add($tVendor) | Out-Null

$tPart = New-MTable "Dim_Part" @"
SELECT PartKey, PartNumber, PartName, Category, UnitOfMeasure, Criticality,
       QualifiedSupplierCount, QualificationMonths,
       CAST(CASE WHEN QualifiedSupplierCount = 1 THEN 1 ELSE 0 END AS BIT) AS IsSingleSource
FROM dbo.Dim_Part
"@
@("PartKey|Int64","PartNumber|String","PartName|String","Category|String","UnitOfMeasure|String",
  "Criticality|String","QualifiedSupplierCount|Int64","QualificationMonths|Int64","IsSingleSource|Boolean") |
  ForEach-Object { $p = $_ -split '\|'; $tPart.Columns.Add((New-Col $p[0] $p[1])) | Out-Null }
$model.Tables.Add($tPart) | Out-Null

$tBuyer = New-MTable "Dim_Buyer" @"
SELECT BuyerKey, BuyerID, BuyerName, Team FROM dbo.Dim_Buyer
"@
@("BuyerKey|Int64","BuyerID|String","BuyerName|String","Team|String") |
  ForEach-Object { $p = $_ -split '\|'; $tBuyer.Columns.Add((New-Col $p[0] $p[1])) | Out-Null }
$model.Tables.Add($tBuyer) | Out-Null

$tPO = New-MTable "POLine" @"
SELECT POLineKey, PONumber, PartKey, VendorKey, BuyerKey, OrderDate, PromisedDate,
       OrderQty, UnitPrice, FreightAmount, ExpediteFee, IsExpedited,
       ExtendedPrice, LandedCost, ContractedUnitPrice, AgreementsInForce, IsOnContract,
       PPVAmount, ContractedExtended,
       ISNULL(QtyReceived,0) AS QtyReceived, ISNULL(QtyAccepted,0) AS QtyAccepted,
       ISNULL(QtyRejected,0) AS QtyRejected, RejectedValue,
       DaysLate, CAST(ISNULL(IsOnTime,0) AS BIT) AS IsOnTime,
       CAST(IsReceived AS BIT) AS IsReceived,
       -- Carried so DAX can exclude duplicated requisitions exactly where SQL
       -- does. fn_SpendKPI counts them (committed spend is what was committed,
       -- duplicate or not); fn_VendorScorecard excludes them (a vendor should
       -- not be ranked on a requisition raised twice). Without this column the
       -- two populations cannot both be expressed in the model, and the Power
       -- BI cost index would disagree with SQL for a reason nobody could see.
       CAST(CASE WHEN EXISTS (SELECT 1 FROM dbo.fn_DuplicatePOLines('$AsOf') d
                              WHERE d.POLineKey = c.POLineKey)
                 THEN 1 ELSE 0 END AS BIT) AS IsDuplicatePO
FROM dbo.fn_POLineCost('$AsOf') c
"@
@("POLineKey|Int64","PONumber|String","PartKey|Int64","VendorKey|Int64","BuyerKey|Int64",
  "OrderDate|DateTime","PromisedDate|DateTime","OrderQty|Int64","UnitPrice|Double",
  "FreightAmount|Double","ExpediteFee|Double","IsExpedited|Boolean","ExtendedPrice|Double",
  "LandedCost|Double","ContractedUnitPrice|Double","AgreementsInForce|Int64","IsOnContract|Boolean",
  "PPVAmount|Double","ContractedExtended|Double","QtyReceived|Int64","QtyAccepted|Int64",
  "QtyRejected|Int64","RejectedValue|Double","DaysLate|Int64","IsOnTime|Boolean","IsReceived|Boolean",
  "IsDuplicatePO|Boolean") |
  ForEach-Object { $p = $_ -split '\|'; $tPO.Columns.Add((New-Col $p[0] $p[1])) | Out-Null }
$model.Tables.Add($tPO) | Out-Null

$tEro = New-MTable "PriceErosion" @"
SELECT e.PartKey, e.VendorKey, e.FirstDate, e.LastDate, e.YearsElapsed, e.Lines,
       e.TotalQty, e.TotalSpend, e.TTMQty, e.TTMSpend,
       e.FirstPrice, e.LastPrice, e.ExpectedPrice,
       e.BenchmarkErosionPct, e.ActualErosionPct, e.ErosionCapturePct,
       e.ErosionGapPerUnit, e.AnnualOpportunity
FROM dbo.fn_PriceErosion('$AsOf', 10, 500) e
"@
@("PartKey|Int64","VendorKey|Int64","FirstDate|DateTime","LastDate|DateTime","YearsElapsed|Double",
  "Lines|Int64","TotalQty|Int64","TotalSpend|Double","TTMQty|Int64","TTMSpend|Double",
  "FirstPrice|Double","LastPrice|Double","ExpectedPrice|Double","BenchmarkErosionPct|Double",
  "ActualErosionPct|Double","ErosionCapturePct|Double","ErosionGapPerUnit|Double","AnnualOpportunity|Double") |
  ForEach-Object { $p = $_ -split '\|'; $tEro.Columns.Add((New-Col $p[0] $p[1])) | Out-Null }
$model.Tables.Add($tEro) | Out-Null

$tQueue = New-MTable "RenegotiationQueue" @"
SELECT q.PriorityRank, q.BuyerRank, p.PartKey, v.VendorKey, b.BuyerKey,
       q.TTMSpend, q.AnnualOpportunity, q.RejectedValue, q.PPV, q.OffContractSpend,
       q.AcceptanceRatePct, q.LeverageScore, q.ActionCode,
       CAST(q.IsThisQuarter AS BIT) AS IsThisQuarter, q.RecommendedAction,
       q.ErosionCapturePct, q.ErosionGapPerUnit
FROM dbo.fn_RenegotiationQueue('$AsOf', 8) q
JOIN dbo.Dim_Part   p ON p.PartNumber = q.PartNumber
JOIN dbo.Dim_Vendor v ON v.VendorID   = q.VendorID
JOIN dbo.Dim_Buyer  b ON b.BuyerID    = q.BuyerID
"@
@("PriorityRank|Int64","BuyerRank|Int64","PartKey|Int64","VendorKey|Int64","BuyerKey|Int64",
  "TTMSpend|Double","AnnualOpportunity|Double","RejectedValue|Double","PPV|Double",
  "OffContractSpend|Double","AcceptanceRatePct|Double","LeverageScore|Double","ActionCode|String",
  "IsThisQuarter|Boolean","RecommendedAction|String","ErosionCapturePct|Double","ErosionGapPerUnit|Double") |
  ForEach-Object { $p = $_ -split '\|'; $tQueue.Columns.Add((New-Col $p[0] $p[1])) | Out-Null }
$model.Tables.Add($tQueue) | Out-Null

$tTargets = New-MTable "Ref_SpendTargets" @"
SELECT MetricName, TargetValue, WarningValue, Direction, Unit, Description FROM dbo.Ref_SpendTargets
"@
@("MetricName|String","TargetValue|Double","WarningValue|Double","Direction|String","Unit|String","Description|String") |
  ForEach-Object { $p = $_ -split '\|'; $tTargets.Columns.Add((New-Col $p[0] $p[1])) | Out-Null }
$model.Tables.Add($tTargets) | Out-Null

$tBench = New-MTable "Ref_PriceErosionBenchmark" @"
SELECT Category, AnnualErosionPct, ToleranceBandPct, MaturityProfile, SourceNote
FROM dbo.Ref_PriceErosionBenchmark
"@
@("Category|String","AnnualErosionPct|Double","ToleranceBandPct|Double","MaturityProfile|String","SourceNote|String") |
  ForEach-Object { $p = $_ -split '\|'; $tBench.Columns.Add((New-Col $p[0] $p[1])) | Out-Null }
$model.Tables.Add($tBench) | Out-Null

# The reporting anchor -- one row, read straight out of fn_SpendKPI. Both the
# TTM measures below and the SQL scorecard take their window from this, so the
# two cannot drift. Anchoring the window on MAX(OrderDate) instead would look
# identical today (the last order happens to fall on the as-of date) and would
# silently widen the window the moment it did not.
$tRep = New-MTable "Ref_Reporting" @"
SELECT AsOfDate, TTMStart = DATEADD(DAY, 1, DATEADD(MONTH, -12, AsOfDate))
FROM dbo.fn_SpendKPI('$AsOf')
"@
@("AsOfDate|DateTime","TTMStart|DateTime") |
  ForEach-Object { $p = $_ -split '\|'; $tRep.Columns.Add((New-Col $p[0] $p[1])) | Out-Null }
$model.Tables.Add($tRep) | Out-Null

$model.SaveChanges() | Out-Null
Write-Host "  9 tables created"

# -----------------------------------------------------------------------------
# 4. Refresh  (the first one raises a credential prompt inside Desktop)
# -----------------------------------------------------------------------------
Write-Host ""
Write-Host "Refreshing tables. If a credential prompt appears in Power BI Desktop," -ForegroundColor Yellow
Write-Host "click through it (Windows / Integrated auth, any privacy level)." -ForegroundColor Yellow
foreach ($tn in @("Dim_Date","Dim_Vendor","Dim_Part","Dim_Buyer","POLine","PriceErosion",
                  "RenegotiationQueue","Ref_SpendTargets","Ref_PriceErosionBenchmark","Ref_Reporting")) {
    $model.Tables[$tn].RequestRefresh([Microsoft.AnalysisServices.Tabular.RefreshType]::Full)
}
$model.SaveChanges() | Out-Null
# Row counts are deliberately NOT printed here. A Desktop refresh does not
# populate EstimatedRowCount, so printing it reports blanks -- and printing
# anything structural would report success whether or not a single row loaded.
# Validate_PowerBI_Model.ps1 queries the model with DAX and reconciles the
# answers against SQL, which is the only check worth trusting.
Write-Host "  Refresh submitted -- verify it with Validate_PowerBI_Model.ps1"

# -----------------------------------------------------------------------------
# 5. Relationships
#
# One direction of filter flow throughout. Dim_Part and Dim_Vendor each feed
# three facts, which is an ordinary star -- the ambiguity trap would be a
# relationship BETWEEN the facts, and there is none.
# -----------------------------------------------------------------------------
Write-Host ""
Write-Host "Creating relationships..."   # already cleared in section 3, above
New-Rel "POLine" "OrderDate"  "Dim_Date"   "Date"
New-Rel "POLine" "PartKey"    "Dim_Part"   "PartKey"
New-Rel "POLine" "VendorKey"  "Dim_Vendor" "VendorKey"
New-Rel "POLine" "BuyerKey"   "Dim_Buyer"  "BuyerKey"
New-Rel "PriceErosion" "PartKey"   "Dim_Part"   "PartKey"
New-Rel "PriceErosion" "VendorKey" "Dim_Vendor" "VendorKey"
New-Rel "RenegotiationQueue" "PartKey"   "Dim_Part"   "PartKey"
New-Rel "RenegotiationQueue" "VendorKey" "Dim_Vendor" "VendorKey"
New-Rel "RenegotiationQueue" "BuyerKey"  "Dim_Buyer"  "BuyerKey"
New-Rel "Dim_Part" "Category" "Ref_PriceErosionBenchmark" "Category"
$model.SaveChanges() | Out-Null
Write-Host "  10 relationships created"

# mark the date table so time intelligence works
$tDate2 = $model.Tables["Dim_Date"]
$tDate2.DataCategory = "Time"
$tDate2.Columns["Date"].IsKey = $true
$model.SaveChanges() | Out-Null

$h = New-Object Microsoft.AnalysisServices.Tabular.Hierarchy
$h.Name = "Calendar"
foreach ($lv in @(@("Year","Year"), @("Quarter","YearQuarter"), @("Month","YearMonth"), @("Date","Date"))) {
    $lvl = New-Object Microsoft.AnalysisServices.Tabular.Level
    $lvl.Name = $lv[0]; $lvl.Column = $tDate2.Columns[$lv[1]]; $lvl.Ordinal = $h.Levels.Count
    $h.Levels.Add($lvl) | Out-Null
}
if ($tDate2.Hierarchies.ContainsName("Calendar")) { $tDate2.Hierarchies.Remove("Calendar") | Out-Null }
$tDate2.Hierarchies.Add($h) | Out-Null
$model.SaveChanges() | Out-Null
Write-Host "  Dim_Date marked as a date table; Calendar hierarchy added"

# -----------------------------------------------------------------------------
# 6. Measures
# -----------------------------------------------------------------------------
Write-Host ""
Write-Host "Creating measures..."
$F = $model.Tables["POLine"]
$E = $model.Tables["PriceErosion"]
$Q = $model.Tables["RenegotiationQueue"]
$R = $model.Tables["Ref_Reporting"]

# --- Spend ---------------------------------------------------------------
New-Measure $F "Extended Price"  "SUM ( POLine[ExtendedPrice] )"  '\$#,0;(\$#,0)' "01 Spend"
New-Measure $F "Landed Cost"     "SUM ( POLine[LandedCost] )"     '\$#,0;(\$#,0)' "01 Spend"
New-Measure $F "Freight"         "SUM ( POLine[FreightAmount] )"  '\$#,0'          "01 Spend"
New-Measure $F "Expedite Fees"   "SUM ( POLine[ExpediteFee] )"    '\$#,0'          "01 Spend"
New-Measure $F "PO Lines"        "COUNTROWS ( POLine )"           '#,0'            "01 Spend"
New-Measure $F "Qty Ordered"     "SUM ( POLine[OrderQty] )"       '#,0'            "01 Spend"

New-Measure $R "As Of Date"  "MAX ( Ref_Reporting[AsOfDate] )" 'yyyy-mm-dd' "00 Reporting"
New-Measure $R "TTM Start"   "MAX ( Ref_Reporting[TTMStart] )"  'yyyy-mm-dd' "00 Reporting"

# --- Contract compliance --------------------------------------------------
New-Measure $F "On Contract Spend" "CALCULATE ( [Extended Price], POLine[IsOnContract] = TRUE () )" '\$#,0' "02 Contract"
New-Measure $F "Maverick Spend"    "CALCULATE ( [Extended Price], POLine[IsOnContract] = FALSE () )" '\$#,0' "02 Contract"
New-Measure $F "Maverick Spend %"  "DIVIDE ( [Maverick Spend], [Extended Price] ) * 100" '0.00' "02 Contract"

# PPV is restricted to contracted lines on BOTH sides. Including off-contract
# lines would treat a missing contracted price as zero and score "paid full
# price against a contract of nothing" as a variance.
New-Measure $F "Purchase Price Variance" "SUM ( POLine[PPVAmount] )" '\$#,0;(\$#,0)' "02 Contract"
New-Measure $F "PPV %" @"
DIVIDE (
    [Purchase Price Variance],
    CALCULATE ( SUM ( POLine[ContractedExtended] ), NOT ISBLANK ( POLine[ContractedUnitPrice] ) )
) * 100
"@ '0.00' "02 Contract"
New-Measure $F "Ambiguous Contract Lines" "CALCULATE ( COUNTROWS ( POLine ), POLine[AgreementsInForce] > 1 )" '#,0' "02 Contract"

# --- Price erosion --------------------------------------------------------
# Spend-weighted, never a simple average: a pair carrying $4m and one carrying
# $40k must not count equally in a company-level figure.
New-Measure $E "Erosion Capture %" @"
DIVIDE (
    SUMX ( PriceErosion, PriceErosion[ErosionCapturePct] * PriceErosion[TTMSpend] ),
    SUM ( PriceErosion[TTMSpend] )
)
"@ '0.00' "03 Erosion"

# Clamped at zero, not netted. A supplier already beating the curve is not a
# savings source and cannot fund another supplier's shortfall.
New-Measure $E "Annual Opportunity" `
    "SUMX ( PriceErosion, IF ( PriceErosion[AnnualOpportunity] > 0, PriceErosion[AnnualOpportunity], 0 ) )" `
    '\$#,0' "03 Erosion"
New-Measure $E "Value Already Ahead of Curve" `
    "SUMX ( PriceErosion, IF ( PriceErosion[AnnualOpportunity] < 0, PriceErosion[AnnualOpportunity], 0 ) )" `
    '\$#,0;(\$#,0)' "03 Erosion"
New-Measure $E "Pairs Assessed" "COUNTROWS ( PriceErosion )" '#,0' "03 Erosion"
New-Measure $E "Pairs Capturing Under 20%" `
    "CALCULATE ( COUNTROWS ( PriceErosion ), PriceErosion[ErosionCapturePct] < 20 )" '#,0' "03 Erosion"
New-Measure $E "Spend on Flat Pairs" `
    "CALCULATE ( SUM ( PriceErosion[TTMSpend] ), PriceErosion[ErosionCapturePct] < 20 )" '\$#,0' "03 Erosion"
New-Measure $E "Benchmark Erosion %" @"
DIVIDE (
    SUMX ( PriceErosion, PriceErosion[BenchmarkErosionPct] * PriceErosion[TTMSpend] ),
    SUM ( PriceErosion[TTMSpend] )
)
"@ '0.00' "03 Erosion"

# --- Quality --------------------------------------------------------------
New-Measure $F "Qty Received" "SUM ( POLine[QtyReceived] )" '#,0' "04 Quality"
New-Measure $F "Qty Accepted" "SUM ( POLine[QtyAccepted] )" '#,0' "04 Quality"
New-Measure $F "Qty Rejected" "SUM ( POLine[QtyRejected] )" '#,0' "04 Quality"
New-Measure $F "Acceptance Rate %" "DIVIDE ( [Qty Accepted], [Qty Received] ) * 100" '0.00' "04 Quality"
New-Measure $F "Rejected Value" "SUM ( POLine[RejectedValue] )" '\$#,0' "04 Quality"

# Unit price answers what we agreed to pay; this answers what each USABLE part
# cost. In optics the two diverge enough to reverse a supplier ranking.
New-Measure $F "Cost per Accepted Unit" "DIVIDE ( [Landed Cost], [Qty Accepted] )" '\$#,0.00' "04 Quality"

# Like-for-like: each vendor against the cheapest available ON THE SAME PARTS.
# An average across a mixed basket of parts compares nothing.
#
# Two things this has to get right, both of which the SQL version got wrong at
# some point. It aggregates per PART (SUMX over parts), never per PO line -- the
# SQL version summed a pair-level landed cost once per line, a fan-out that kept
# the index above 100 and moved 33 of 36 vendors' rank positions. And it drops
# duplicated requisitions, matching fn_VendorScorecard: a vendor should not be
# ranked on a requisition somebody raised twice.
# Two traps here, both of which returned a plausible number:
#
# 1. MINX over CALCULATETABLE(VALUES(Dim_Vendor[VendorKey]), ALL(Dim_Vendor))
#    builds the right LIST, but the context transition inside adds each row's
#    key to a filter context that still holds the outer vendor selection.
#    Filters on one table intersect, so every other vendor resolves to an empty
#    set and returns blank -- MINX then sees only the selected vendor's own
#    cost, and a vendor compared with itself is exactly 100.00. The inner
#    CALCULATE must therefore clear Dim_Vendor and re-apply just the iterated
#    key.
# 2. SQL restricts to part-vendor pairs with accepted quantity > 0 (the HAVING
#    in PartVendor). Without the same restriction, a part this vendor supplied
#    but had nothing accepted on adds landed cost to the numerator and nothing
#    to the denominator, inflating the index.
New-Measure $F "Cost Index vs Best" @"
VAR Clean = FILTER ( ALL ( POLine[IsDuplicatePO] ), POLine[IsDuplicatePO] = FALSE () )
RETURN
CALCULATE (
    VAR PartsSupplied = FILTER ( VALUES ( Dim_Part[PartKey] ), [Qty Accepted] > 0 )
    VAR MyCost = SUMX ( PartsSupplied, CALCULATE ( [Landed Cost] ) )
    VAR BestCost =
        SUMX (
            PartsSupplied,
            VAR Acc = CALCULATE ( [Qty Accepted] )
            VAR BestOnPart =
                MINX (
                    CALCULATETABLE ( VALUES ( Dim_Vendor[VendorKey] ), ALL ( Dim_Vendor ) ),
                    VAR vk = Dim_Vendor[VendorKey]
                    RETURN CALCULATE ( [Cost per Accepted Unit], ALL ( Dim_Vendor ), Dim_Vendor[VendorKey] = vk )
                )
            RETURN BestOnPart * Acc
        )
    RETURN DIVIDE ( MyCost, BestCost ) * 100,
    Clean
)
"@ '0.00' "04 Quality"

New-Measure $F "Duplicate PO Spend" `
    "CALCULATE ( [Extended Price], POLine[IsDuplicatePO] = TRUE () )" '\$#,0' "04 Quality"
New-Measure $F "Duplicate PO Lines" `
    "CALCULATE ( COUNTROWS ( POLine ), POLine[IsDuplicatePO] = TRUE () )" '#,0' "04 Quality"

# --- Delivery -------------------------------------------------------------
New-Measure $F "Receipts Booked" "CALCULATE ( COUNTROWS ( POLine ), POLine[IsReceived] = TRUE () )" '#,0' "05 Delivery"
New-Measure $F "On Time Receipts" `
    "CALCULATE ( COUNTROWS ( POLine ), POLine[IsOnTime] = TRUE (), POLine[IsReceived] = TRUE () )" '#,0' "05 Delivery"
New-Measure $F "On Time Delivery %" "DIVIDE ( [On Time Receipts], [Receipts Booked] ) * 100" '0.00' "05 Delivery"
New-Measure $F "Avg Days Late" `
    "CALCULATE ( AVERAGE ( POLine[DaysLate] ), POLine[IsReceived] = TRUE () )" '0.0' "05 Delivery"
New-Measure $F "Expedite Spend %" "DIVIDE ( [Expedite Fees], [Extended Price] ) * 100" '0.00' "05 Delivery"

# --- Leverage and concentration -------------------------------------------
New-Measure $F "Single Source Spend %" `
    "DIVIDE ( CALCULATE ( [Extended Price], Dim_Part[IsSingleSource] = TRUE () ), [Extended Price] ) * 100" `
    '0.00' "06 Leverage"
New-Measure $F "Top 5 Vendor Share %" @"
VAR VendorSpend = ADDCOLUMNS ( VALUES ( Dim_Vendor[VendorKey] ), "@Spend", [Extended Price] )
VAR Top5 = TOPN ( 5, VendorSpend, [@Spend], DESC )
RETURN DIVIDE ( SUMX ( Top5, [@Spend] ), SUMX ( VendorSpend, [@Spend] ) ) * 100
"@ '0.00' "06 Leverage"

# Identical weights to the SQL implementation, deliberately. If the two ever
# diverge, the Power BI report and the Excel workbook rank different
# negotiations and neither can be trusted.
New-Measure $Q "Leverage Score" @"
VAR Sources    = SELECTEDVALUE ( Dim_Part[QualifiedSupplierCount] )
VAR QualMonths = SELECTEDVALUE ( Dim_Part[QualificationMonths] )
VAR RevShare   = SELECTEDVALUE ( Dim_Vendor[LumenRevenueSharePct] )
VAR OffContract = SUM ( RenegotiationQueue[OffContractSpend] )
RETURN
      SWITCH ( TRUE (), Sources >= 3, 40, Sources = 2, 26, 4 )
    + SWITCH ( TRUE (), QualMonths <= 2, 20, QualMonths <= 5, 13, QualMonths <= 8, 6, 0 )
    + SWITCH ( TRUE (), RevShare >= 15, 30, RevShare >= 8, 20, RevShare >= 4, 11, 3 )
    + IF ( OffContract > 0, 10, 6 )
"@ '0.0' "06 Leverage"

# --- Queue ----------------------------------------------------------------
New-Measure $Q "Queue Opportunity" "SUM ( RenegotiationQueue[AnnualOpportunity] )" '\$#,0' "07 Queue"
New-Measure $Q "Queue Pairs" "COUNTROWS ( RenegotiationQueue )" '#,0' "07 Queue"
New-Measure $Q "This Quarter Pairs" `
    "CALCULATE ( COUNTROWS ( RenegotiationQueue ), RenegotiationQueue[IsThisQuarter] = TRUE () )" '#,0' "07 Queue"
New-Measure $Q "This Quarter Opportunity" `
    "CALCULATE ( SUM ( RenegotiationQueue[AnnualOpportunity] ), RenegotiationQueue[IsThisQuarter] = TRUE () )" `
    '\$#,0' "07 Queue"
New-Measure $Q "Quarter Coverage %" `
    "DIVIDE ( [This Quarter Opportunity], [Queue Opportunity] ) * 100" '0.0' "07 Queue"

# --- Trailing twelve months -----------------------------------------------
#
# WHY THESE EXIST AT ALL. The base ratios above deliberately respond to
# whatever the user has filtered: that is what makes a trend line or a
# per-vendor breakdown mean anything. But the published scorecard is a TTM
# figure, and a card with no slicer on it would show an all-history number
# that looks equally plausible and is not the number in the case study.
# So the headline KPIs get an explicit TTM pair, anchored on Ref_Reporting --
# the same window fn_SpendKPI uses -- and those are what the scorecard cards
# bind to. Nobody then has to remember to set a slicer for the page to be right.
$ttm = @(
    @($F, "Spend (TTM)",                  "[Extended Price]",         '\$#,0'),
    @($F, "Landed Cost (TTM)",            "[Landed Cost]",            '\$#,0'),
    @($F, "PO Lines (TTM)",               "[PO Lines]",               '#,0'),
    @($F, "Maverick Spend % (TTM)",       "[Maverick Spend %]",       '0.00'),
    @($F, "PPV % (TTM)",                  "[PPV %]",                  '0.00'),
    @($F, "Acceptance Rate % (TTM)",      "[Acceptance Rate %]",      '0.00'),
    @($F, "On Time Delivery % (TTM)",     "[On Time Delivery %]",     '0.00'),
    @($F, "Expedite Spend % (TTM)",       "[Expedite Spend %]",       '0.00'),
    @($F, "Single Source Spend % (TTM)",  "[Single Source Spend %]",  '0.00'),
    @($F, "Top 5 Vendor Share % (TTM)",   "[Top 5 Vendor Share %]",   '0.00')
)
foreach ($t in $ttm) {
    New-Measure $t[0] $t[1] @"
CALCULATE (
    $($t[2]),
    DATESBETWEEN ( Dim_Date[Date], MAX ( Ref_Reporting[TTMStart] ), MAX ( Ref_Reporting[AsOfDate] ) )
)
"@ $t[3] "09 TTM Scorecard"
}
# Erosion Capture % and Annual Opportunity get no TTM twin on purpose: the
# erosion model already weights by TTM spend internally and spans each pair's
# whole price history by design, so windowing it again would double-apply the
# window and shorten the very history the trend is measured over.

# --- Targets and RAG colours ----------------------------------------------
# A colour measure returns a HEX STRING, not a number. Power BI's conditional
# formatting by field value expects a colour; returning 1/2/3 and hoping the
# rule interprets it is the most common reason a format rule silently does
# nothing -- this project hit exactly that in Project 1.
# These bind to the TTM variants, not the base ratios. DATESBETWEEN replaces
# only the DATE filter, so a status measure still answers "how is THIS vendor
# doing" when a vendor is selected -- it just refuses to answer it over a
# window other than the published one. The alternative, colouring an
# all-history figure against a TTM target, is a red light with no meaning.
$ragMetrics = @(
    @("Erosion Capture",   "ErosionCapturePct",    "[Erosion Capture %]",           "H", $E),
    @("Maverick Spend",    "MaverickSpendPct",     "[Maverick Spend % (TTM)]",      "L", $F),
    @("PPV",               "PPVPct",               "[PPV % (TTM)]",                 "L", $F),
    @("Acceptance",        "AcceptanceRatePct",    "[Acceptance Rate % (TTM)]",     "H", $F),
    @("On Time Delivery",  "OnTimeDeliveryPct",    "[On Time Delivery % (TTM)]",    "H", $F),
    @("Expedite Spend",    "ExpediteSpendPct",     "[Expedite Spend % (TTM)]",      "L", $F),
    @("Single Source",     "SingleSourceSpendPct", "[Single Source Spend % (TTM)]", "L", $F),
    @("Top 5 Share",       "Top5VendorSharePct",   "[Top 5 Vendor Share % (TTM)]",  "L", $F)
)
foreach ($r in $ragMetrics) {
    $label = $r[0]; $metric = $r[1]; $val = $r[2]; $dir = $r[3]; $tbl = $r[4]
    $cmp = if ($dir -eq "H") { "V >= T, ""#C6EFCE"", V >= W, ""#FFEB9C""" } else { "V <= T, ""#C6EFCE"", V <= W, ""#FFEB9C""" }
    New-Measure $tbl "$label Colour" @"
VAR V = $val
VAR T = LOOKUPVALUE ( Ref_SpendTargets[TargetValue],  Ref_SpendTargets[MetricName], "$metric" )
VAR W = LOOKUPVALUE ( Ref_SpendTargets[WarningValue], Ref_SpendTargets[MetricName], "$metric" )
RETURN SWITCH ( TRUE (), $cmp, "#FFC7CE" )
"@ $null "08 Status"
    New-Measure $tbl "$label Status" @"
VAR V = $val
VAR T = LOOKUPVALUE ( Ref_SpendTargets[TargetValue],  Ref_SpendTargets[MetricName], "$metric" )
VAR W = LOOKUPVALUE ( Ref_SpendTargets[WarningValue], Ref_SpendTargets[MetricName], "$metric" )
RETURN SWITCH ( TRUE (), $(if ($dir -eq "H") { "V >= T, ""Green"", V >= W, ""Amber""" } else { "V <= T, ""Green"", V <= W, ""Amber""" }), "Red" )
"@ $null "08 Status"
}

$model.SaveChanges() | Out-Null
$measureCount = ($model.Tables | ForEach-Object { $_.Measures.Count } | Measure-Object -Sum).Sum
# Counted, not asserted. The folder count was hardcoded at 8 and stayed there
# while two more folders were added -- the kind of number a reader takes on
# trust because it sits beside one that is computed.
$folderCount = ($model.Tables | ForEach-Object { $_.Measures } |
                Where-Object { $_.DisplayFolder } |
                Select-Object -ExpandProperty DisplayFolder -Unique).Count
Write-Host "  $measureCount measures created across $folderCount display folders"

# -----------------------------------------------------------------------------
# 7. Hide keys from the report view -- a field list full of surrogate keys
#    invites someone to drag one onto a visual and sum it.
# -----------------------------------------------------------------------------
foreach ($pair in @(@("POLine","POLineKey"), @("POLine","PartKey"), @("POLine","VendorKey"),
                    @("POLine","BuyerKey"), @("PriceErosion","PartKey"), @("PriceErosion","VendorKey"),
                    @("RenegotiationQueue","PartKey"), @("RenegotiationQueue","VendorKey"),
                    @("RenegotiationQueue","BuyerKey"), @("Dim_Date","DateKey"))) {
    $model.Tables[$pair[0]].Columns[$pair[1]].IsHidden = $true
}
$model.SaveChanges() | Out-Null

Write-Host ("-" * 78)
Write-Host "Model built. Save the file in Power BI Desktop as powerbi/Lumen_Spend_Scorecard.pbix," -ForegroundColor Green
Write-Host "then build the report pages using POWER_BI_BUILD_GUIDE.md." -ForegroundColor Green
$server.Disconnect()
