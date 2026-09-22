<#
================================================================================
Project 3 -- Lumen Optics Manufacturing: Photonics Spend Scorecard
Script:  Build_Workbook.ps1
Purpose: Build Lumen_Spend_Scorecard.xlsx -- the control a category manager
         opens to decide which negotiation to run next.

WHAT MAKES THIS A CONTROL AND NOT A REPORT
    Nothing on the Dashboard is a pasted number. Every KPI is recomputed by
    formula from the purchase-order extract, the RAG status comes from
    INDEX/MATCH against the target table, and the Validation sheet then holds
    each Excel figure against the one SQL produced independently. If the two
    implementations ever disagree, the workbook says so on its own face.

WHY NO DYNAMIC ARRAYS
    FILTER, SORT and UNIQUE would shorten several formulas and break the file
    on Excel 2019 and LibreOffice. The queue ranking uses SUMPRODUCT and
    INDEX/MATCH, which work everywhere.

BUILD ORDER MATTERS
    Every worksheet is created before a single formula is written. Excel does
    not defer a reference to a sheet that does not exist yet: it accepts the
    formula, rewrites the reference to #REF!, and the build still reports
    success.

DATA DISCLOSURE: Lumen Optics Manufacturing is fictional; all data is
synthetic. No confidential data and no production system is involved.
================================================================================
#>

param(
    [string]$ProjectRoot = (Split-Path $PSScriptRoot -Parent),
    [string]$OutputFile  = (Join-Path $PSScriptRoot 'Lumen_Spend_Scorecard.xlsx')
)

$ErrorActionPreference = 'Stop'
$ExtractPath = Join-Path $ProjectRoot 'erp_extracts'
if (-not (Test-Path (Join-Path $ExtractPath 'po_line_cost.csv'))) {
    throw "Extracts not found in $ExtractPath. Run erp_extracts\Export_Extracts.ps1 first."
}

function RGBv([int]$r, [int]$g, [int]$b) { return $r + ($g * 256) + ($b * 65536) }
$Navy    = RGBv 26 54 93
$NavyLt  = RGBv 214 226 240
$Grey    = RGBv 242 242 242
$GreenBg = RGBv 198 239 206 ; $GreenFg = RGBv 0 97 0
$AmberBg = RGBv 255 235 156 ; $AmberFg = RGBv 156 101 0
$RedBg   = RGBv 255 199 206 ; $RedFg   = RGBv 156 0 6
$White   = RGBv 255 255 255

Write-Host ""
Write-Host "Building Lumen_Spend_Scorecard.xlsx" -ForegroundColor Cyan
Write-Host ("-" * 78)

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.ScreenUpdating = $false

try {
    $wb = $xl.Workbooks.Add()
    while ($wb.Worksheets.Count -gt 1) { $wb.Worksheets.Item($wb.Worksheets.Count).Delete() }

    # ---------------------------------------------------------------- Control
    $ctl = $wb.Worksheets.Item(1); $ctl.Name = 'Control'
    $ctl.Range('A1').Value2 = 'LUMEN OPTICS MANUFACTURING -- SPEND SCORECARD'
    $ctl.Range('A1').Font.Size = 16; $ctl.Range('A1').Font.Bold = $true; $ctl.Range('A1').Font.Color = $Navy
    $ctl.Range('A2').Value2 = 'Control panel. Every sheet reads its parameters from the cells below.'
    $ctl.Range('A2').Font.Italic = $true
    $ctl.Range('A4').Value2 = 'PARAMETER'; $ctl.Range('B4').Value2 = 'VALUE'; $ctl.Range('C4').Value2 = 'WHAT IT CONTROLS'
    $ctl.Range('A4:C4').Font.Bold = $true
    $ctl.Range('A4:C4').Interior.Color = $Navy
    $ctl.Range('A4:C4').Font.Color = $White

    $params = @(
        @('Reporting date',        'AsOfDate',        'DATE', '2025-12-31', 'The as-of date for every figure. Spend metrics use the twelve months ending here.'),
        @('Extract folder',        'SourceFolder',    'TEXT', $ExtractPath, 'Where Power Query looks for the CSV extracts. Edit if you move the repo, then Data > Refresh All.'),
        @('Negotiations / buyer',  'NegotiationsPerBuyer','NUM', 8,         'Capacity assumption behind the quarter''s worklist. Raising it lengthens every category manager''s list.'),
        @('Queue rows shown',      'QueueRowsShown',  'NUM',  25,           'How many ranked pairs the Renegotiation Queue sheet displays.'),
        @('Baseline lag (days)',   'MaterialThreshold','NUM', 100000,       'Opportunity above which a negotiation is worth opening regardless of leverage.')
    )
    $r = 5
    foreach ($p in $params) {
        $ctl.Range("A$r").Value2 = [string]$p[0]
        if ($p[2] -eq 'DATE') {
            $d = [datetime]::Parse([string]$p[3])
            $ctl.Range("B$r").NumberFormat = 'yyyy-mm-dd'
            $ctl.Range("B$r").Formula = "=DATE($($d.Year),$($d.Month),$($d.Day))"
        } else {
            $ctl.Range("B$r").Formula = [string]$p[3]
        }
        $ctl.Range("C$r").Value2 = [string]$p[4]
        $ctl.Range("B$r").Interior.Color = $NavyLt
        $ctl.Range("B$r").Font.Bold = $true
        $wb.Names.Add($p[1], $ctl.Range("B$r")) | Out-Null
        $r++
    }
    $ctl.Range('A12').Value2 = 'DATA DISCLOSURE'
    $ctl.Range('A12').Font.Bold = $true; $ctl.Range('A12').Font.Color = $Navy
    $ctl.Range('A13').Value2 = 'Lumen Optics Manufacturing is a fictional company. All data in this workbook is synthetic and generated by the scripts in the sql\ folder. No confidential data is used and no claim is made about any production system.'
    $ctl.Range('A13:C15').Merge(); $ctl.Range('A13').WrapText = $true; $ctl.Range('A13').VerticalAlignment = -4160
    $ctl.Columns.Item(1).ColumnWidth = 24
    $ctl.Columns.Item(2).ColumnWidth = 46
    $ctl.Columns.Item(3).ColumnWidth = 84
    Write-Host "  Control sheet            built"

    # ----------------------------------------------------------- Power Query
    function New-ExtractQuery {
        param([string]$QueryName, [string]$FileName, [string]$TypeList)
        $m = @"
let
    Folder   = Excel.CurrentWorkbook(){[Name="SourceFolder"]}[Content]{0}[Column1],
    Source   = Csv.Document(File.Contents(Folder & "\$FileName"), [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.Csv]),
    Promoted = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    Typed    = Table.TransformColumnTypes(Promoted, {$TypeList})
in
    Typed
"@
        $wb.Queries.Add($QueryName, $m) | Out-Null
    }

    # Types are declared explicitly. Letting Power Query guess means a part
    # number that happens to look numeric arrives as a number in one refresh and
    # text in the next, and every lookup against it breaks.
    $queries = @(
      @{ N='tPO'; F='po_line_cost.csv'; T='{"PONumber",type text},{"PartNumber",type text},{"PartName",type text},{"Category",type text},{"Criticality",type text},{"QualifiedSupplierCount",Int64.Type},{"QualificationMonths",Int64.Type},{"IsSingleSource",Int64.Type},{"VendorID",type text},{"VendorName",type text},{"Country",type text},{"VendorTier",type text},{"LumenRevenueSharePct",type number},{"BuyerID",type text},{"BuyerName",type text},{"Team",type text},{"OrderDate",type date},{"PromisedDate",type date},{"OrderYear",Int64.Type},{"OrderYearMonth",type text},{"OrderQty",Int64.Type},{"UnitPrice",type number},{"FreightAmount",type number},{"ExpediteFee",type number},{"IsExpedited",Int64.Type},{"ExtendedPrice",type number},{"LandedCost",type number},{"AgreementNo",type text},{"ContractedUnitPrice",type number},{"AgreementsInForce",Int64.Type},{"IsOnContract",Int64.Type},{"PPVAmount",type number},{"ContractedExtended",type number},{"QtyReceived",Int64.Type},{"QtyAccepted",Int64.Type},{"QtyRejected",Int64.Type},{"Receipts",Int64.Type},{"IsReceived",Int64.Type},{"AcceptanceRatePct",type number},{"CostPerAcceptedUnit",type number},{"RejectedValue",type number},{"LastReceiptDate",type date},{"DaysLate",Int64.Type},{"IsOnTime",Int64.Type}' },
      @{ N='tErosion'; F='price_erosion.csv'; T='{"PartNumber",type text},{"PartName",type text},{"Category",type text},{"Criticality",type text},{"QualifiedSupplierCount",Int64.Type},{"QualificationMonths",Int64.Type},{"VendorID",type text},{"VendorName",type text},{"VendorTier",type text},{"LumenRevenueSharePct",type number},{"FirstDate",type date},{"LastDate",type date},{"YearsElapsed",type number},{"Lines",Int64.Type},{"TotalQty",Int64.Type},{"TotalSpend",type number},{"TTMQty",Int64.Type},{"TTMSpend",type number},{"FirstPrice",type number},{"LastPrice",type number},{"ExpectedPrice",type number},{"BenchmarkErosionPct",type number},{"ToleranceBandPct",type number},{"ActualErosionPct",type number},{"ErosionCapturePct",type number},{"ErosionGapPerUnit",type number},{"AnnualOpportunity",type number}' },
      @{ N='tQueue'; F='renegotiation_queue.csv'; T='{"PriorityRank",Int64.Type},{"BuyerRank",Int64.Type},{"PartNumber",type text},{"PartName",type text},{"Category",type text},{"Criticality",type text},{"QualifiedSupplierCount",Int64.Type},{"QualificationMonths",Int64.Type},{"IsSingleSource",Int64.Type},{"VendorID",type text},{"VendorName",type text},{"VendorTier",type text},{"LumenRevenueSharePct",type number},{"BuyerID",type text},{"BuyerName",type text},{"Team",type text},{"FirstPrice",type number},{"LastPrice",type number},{"ExpectedPrice",type number},{"BenchmarkErosionPct",type number},{"ActualErosionPct",type number},{"ErosionCapturePct",type number},{"ErosionGapPerUnit",type number},{"TTMQty",Int64.Type},{"TTMSpend",type number},{"AnnualOpportunity",type number},{"RejectedValue",type number},{"PPV",type number},{"OffContractSpend",type number},{"AcceptanceRatePct",type number},{"LeverageScore",type number},{"ActionCode",type text},{"IsThisQuarter",Int64.Type},{"RecommendedAction",type text}' },
      @{ N='tVendor'; F='vendor_scorecard.csv'; T='{"VendorID",type text},{"VendorName",type text},{"Country",type text},{"VendorTier",type text},{"LumenRevenueSharePct",type number},{"Lines",Int64.Type},{"TotalSpend",type number},{"LandedSpend",type number},{"FreightSpend",type number},{"ExpediteSpend",type number},{"ExpeditePctOfSpend",type number},{"OnContractPct",type number},{"PPVAmount",type number},{"PPVPct",type number},{"QtyReceived",Int64.Type},{"QtyAccepted",Int64.Type},{"AcceptanceRatePct",type number},{"RejectedValue",type number},{"OnTimePct",type number},{"AvgDaysLate",type number},{"CostIndexVsBest",type number},{"SharedParts",Int64.Type}' },
      @{ N='tKPI'; F='spend_kpi_monthly.csv'; T='{"AsOfDate",type date},{"Lines",Int64.Type},{"TotalSpend",type number},{"LandedSpend",type number},{"PPVAmount",type number},{"PPVPct",type number},{"MaverickSpendPct",type number},{"AcceptanceRatePct",type number},{"RejectedValue",type number},{"OnTimeDeliveryPct",type number},{"ExpediteSpendPct",type number},{"ErosionCapturePct",type number},{"ErosionOpportunity",type number},{"SingleSourceSpendPct",type number},{"Top5VendorSharePct",type number}' },
      @{ N='tCategory'; F='category_summary.csv'; T='{"AsOfDate",type date},{"GroupedBy",type text},{"GroupValue",type text},{"Lines",Int64.Type},{"TotalSpend",type number},{"LandedSpend",type number},{"SpendSharePct",type number},{"MaverickPct",type number},{"AcceptancePct",type number},{"OnTimePct",type number},{"ExpeditePct",type number},{"SingleSourcePct",type number}' },
      @{ N='tBenchmark'; F='erosion_benchmark.csv'; T='{"Category",type text},{"AnnualErosionPct",type number},{"ToleranceBandPct",type number},{"MaturityProfile",type text},{"SourceNote",type text}' },
      @{ N='tTargets'; F='spend_targets.csv'; T='{"MetricName",type text},{"TargetValue",type number},{"WarningValue",type number},{"Direction",type text},{"Unit",type text},{"Description",type text}' },
      @{ N='tReject'; F='reject_reasons.csv'; T='{"RejectReasonCode",type text},{"Description",type text},{"Attribution",type text}' },
      @{ N='tDQ'; F='data_quality_summary.csv'; T='{"AnomalyType",type text},{"EntityType",type text},{"Severity",type text},{"ImpactClass",type text},{"Anomalies",Int64.Type},{"AmountAtRisk",type number},{"PopulationScanned",Int64.Type},{"AnomalyRatePct",type number},{"CountsTowardExposure",Int64.Type}' }
    )
    foreach ($q in $queries) { New-ExtractQuery -QueryName $q.N -FileName $q.F -TypeList $q.T }
    Write-Host "  Power Query queries      $($queries.Count) created"

    $loaded = @{}
    foreach ($q in $queries) {
        $ws = $wb.Worksheets.Add([System.Reflection.Missing]::Value, $wb.Worksheets.Item($wb.Worksheets.Count))
        $ws.Name = 'Data_' + $q.N.Substring(1)
        $connStr = "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=`$Workbook`$;Location=$($q.N);Extended Properties="""""
        $lo = $ws.ListObjects.Add(0, $connStr, $false, 1, $ws.Range('A1'))
        $lo.QueryTable.CommandType = 2
        $lo.QueryTable.CommandText = "SELECT * FROM [$($q.N)]"
        $lo.QueryTable.BackgroundQuery = $false
        $lo.QueryTable.AdjustColumnWidth = $true
        $lo.QueryTable.Refresh($false) | Out-Null
        $lo.Name = $q.N
        $loaded[$q.N] = $lo.ListRows.Count
        $ws.Rows.Item(1).Font.Bold = $true
        Write-Host ("    {0,-11} {1,6} rows loaded" -f $q.N, $loaded[$q.N])
    }
    if ($loaded['tPO'] -lt 100) { throw "tPO loaded only $($loaded['tPO']) rows -- the Power Query load did not work." }
    $nQueue = $loaded['tQueue']

    # Create every remaining sheet BEFORE writing a single formula.
    $sheetOrder = @('README','Dashboard','Renegotiation Queue','Vendor Scorecard','Validation','Calc')
    $made = @{}
    foreach ($name in $sheetOrder) {
        $ws = $wb.Worksheets.Add([System.Reflection.Missing]::Value, $wb.Worksheets.Item($wb.Worksheets.Count))
        $ws.Name = $name; $made[$name] = $ws
    }
    $rd = $made['README']; $dash = $made['Dashboard']; $q = $made['Renegotiation Queue']
    $vs = $made['Vendor Scorecard']; $val = $made['Validation']; $calc = $made['Calc']
    Write-Host "  Worksheets               $($sheetOrder.Count) created up front"

    # ------------------------------------------------------------------- Calc
    $calc.Range('A1').Value2 = 'WORKING -- the twelve-month window, the top-five concentration and the queue ranking'
    $calc.Range('A1').Font.Bold = $true; $calc.Range('A1').Font.Size = 13; $calc.Range('A1').Font.Color = $Navy
    $calc.Range('A2').Value2 = 'Nothing here is decorative. The Dashboard and the Queue read these cells.'
    $calc.Range('A2').Font.Italic = $true

    $calc.Range('A4').Value2 = 'Window opens after'
    $calc.Range('C4').Formula = '=EOMONTH(AsOfDate,-12)'
    $calc.Range('C4').NumberFormat = 'yyyy-mm-dd'
    $calc.Range('A5').Value2 = 'Window closes'
    $calc.Range('C5').Formula = '=AsOfDate'
    $calc.Range('C5').NumberFormat = 'yyyy-mm-dd'
    $calc.Range('A4:A5').Font.Bold = $true
    $wb.Names.Add('WinFrom', $calc.Range('C4')) | Out-Null
    $wb.Names.Add('WinTo',   $calc.Range('C5')) | Out-Null

    # Vendor spend in the SAME twelve-month window the SQL KPI uses.
    #
    # The vendor scorecard deliberately spans all history -- it is a supplier
    # profile, not a period measure -- so reading the concentration from it
    # answers a different question under the same label. Reading it that way
    # gave 26.80% against SQL's 27.20%: both correct, neither comparable.
    $calc.Range('A7').Value2 = 'VENDOR SPEND IN THE TWELVE-MONTH WINDOW (the top-five concentration, computed not read)'
    $calc.Range('A7').Font.Bold = $true
    $calc.Cells.Item(8,1).Value2 = 'Vendor'
    $calc.Cells.Item(8,3).Value2 = 'Spend in window'
    $calc.Range('A8:C8').Font.Bold = $true
    $calc.Range('A8:C8').Interior.Color = $Grey
    $nVend = $loaded['tVendor']
    $vFirst = 9
    $vLastRow = 8 + $nVend
    for ($i = 0; $i -lt $nVend; $i++) {
        $rw = $vFirst + $i
        $n  = $i + 1
        $calc.Cells.Item($rw,1).Formula = "=IFERROR(INDEX(tVendor[VendorID],$n),"""")"
        $calc.Cells.Item($rw,3).Formula =
            "=IF(`$A$rw="""","""",SUMIFS(tPO[ExtendedPrice],tPO[VendorID],`$A$rw,tPO[OrderDate],"">""&WinFrom,tPO[OrderDate],""<=""&WinTo))"
        $calc.Cells.Item($rw,3).NumberFormat = '#,##0.00'
    }
    $t5Row = $vLastRow + 2
    $allRow = $vLastRow + 3
    $rng = "`$C`$$vFirst`:`$C`$$vLastRow"
    $calc.Cells.Item($t5Row,1).Value2 = 'Top five total'
    $calc.Cells.Item($t5Row,3).Formula =
        "=LARGE($rng,1)+LARGE($rng,2)+LARGE($rng,3)+LARGE($rng,4)+LARGE($rng,5)"
    $calc.Cells.Item($allRow,1).Value2 = 'All vendors in window'
    $calc.Cells.Item($allRow,3).Formula = "=SUM($rng)"
    $calc.Range($calc.Cells.Item($t5Row,1), $calc.Cells.Item($allRow,3)).Font.Bold = $true
    $calc.Range($calc.Cells.Item($t5Row,3), $calc.Cells.Item($allRow,3)).NumberFormat = '#,##0.00'
    $wb.Names.Add('xlTop5',      $calc.Cells.Item($t5Row,3))  | Out-Null
    $wb.Names.Add('xlWindowAll', $calc.Cells.Item($allRow,3)) | Out-Null

    # queue ranking under whatever filters the Queue sheet has set
    # Placed BELOW the vendor block, and derived from where that block actually
    # ended rather than at a fixed row. A hard-coded offset here is how the
    # Project 2 workbook wrote its ageing lists over its own countback flags --
    # the symptom was not an error but a plausible wrong number.
    $qTitle = $allRow + 3
    $qHead  = $qTitle + 1
    $qFirst = $qHead + 1
    $calc.Cells.Item($qTitle,1).Value2 = 'QUEUE RANKING -- re-ranks the SQL queue under the filters set on the Renegotiation Queue sheet'
    $calc.Cells.Item($qTitle,1).Font.Bold = $true
    $calc.Cells.Item($qHead,1).Value2 = 'Row'
    $calc.Cells.Item($qHead,2).Value2 = 'Matches filters'
    $calc.Cells.Item($qHead,3).Value2 = 'Rank within filter'
    $calc.Range($calc.Cells.Item($qHead,1), $calc.Cells.Item($qHead,3)).Font.Bold = $true
    $calc.Range($calc.Cells.Item($qHead,1), $calc.Cells.Item($qHead,3)).Interior.Color = $Grey
    $qLast = $qFirst + $nQueue - 1
    for ($i = 0; $i -lt $nQueue; $i++) {
        $rw = $qFirst + $i; $n = $i + 1
        $calc.Cells.Item($rw,1).Formula = [string]$n
        $calc.Cells.Item($rw,2).Formula = "=IF(AND(" +
            "OR('Renegotiation Queue'!`$C`$4=""(All)"",INDEX(tQueue[BuyerName],$n)='Renegotiation Queue'!`$C`$4)," +
            "OR('Renegotiation Queue'!`$C`$5=""(All)"",INDEX(tQueue[ActionCode],$n)='Renegotiation Queue'!`$C`$5)," +
            "OR('Renegotiation Queue'!`$C`$6=""(All)"",INDEX(tQueue[Category],$n)='Renegotiation Queue'!`$C`$6)," +
            "OR('Renegotiation Queue'!`$C`$7=""No"",INDEX(tQueue[IsThisQuarter],$n)=1)),1,0)"
        $calc.Cells.Item($rw,3).Formula = "=IF(`$B$rw=0,"""",SUMPRODUCT((`$B`$$qFirst`:`$B`$$qLast=1)*(tQueue[PriorityRank]<INDEX(tQueue[PriorityRank],$n)))+1)"
    }
    $calc.Columns.Item(1).ColumnWidth = 34
    Write-Host "  Calc sheet               built ($nQueue queue rows)"

    # -------------------------------------------------------------- Dashboard
    $dash.Range('A1').Value2 = 'LUMEN OPTICS MANUFACTURING -- PHOTONICS SPEND SCORECARD'
    $dash.Range('A1').Font.Size = 18; $dash.Range('A1').Font.Bold = $true; $dash.Range('A1').Font.Color = $Navy
    $dash.Range('A2').Formula = '="Twelve months to "&TEXT(AsOfDate,"d mmmm yyyy")&"   |   Fictional company, synthetic data, no production system."'
    $dash.Range('A2').Font.Italic = $true

    $dash.Range('A4').Value2 = 'THE QUESTION'
    $dash.Range('A4').Font.Bold = $true; $dash.Range('A4').Font.Color = $White
    $dash.Range('A4:H4').Interior.Color = $Navy
    $dash.Range('A5').Value2 = 'Our purchase prices are flat or falling and every variance report is green -- so why is material cost per ACCEPTED unit rising, and which supplier-part combinations should sourcing renegotiate first?'
    $dash.Range('A5:H5').Merge(); $dash.Range('A5').WrapText = $true; $dash.Range('A5').RowHeight = 30

    # base figures, recomputed from the extract
    $dash.Range('A7').Value2 = 'SPEND -- recomputed in Excel from the purchase-order extract'
    $dash.Range('A7').Font.Bold = $true
    $ttm = 'tPO[OrderDate],">"&WinFrom,tPO[OrderDate],"<="&WinTo'
    $base = @(
        @(8,  'Extended price (12 months)', "=SUMIFS(tPO[ExtendedPrice],$ttm)"),
        @(9,  'Landed cost (12 months)',    "=SUMIFS(tPO[LandedCost],$ttm)"),
        @(10, 'Freight',                    "=SUMIFS(tPO[FreightAmount],$ttm)"),
        @(11, 'Expedite fees',              "=SUMIFS(tPO[ExpediteFee],$ttm)"),
        @(12, 'Off-contract spend',         "=SUMIFS(tPO[ExtendedPrice],$ttm,tPO[IsOnContract],0)"),
        @(13, 'Single-source spend',        "=SUMIFS(tPO[ExtendedPrice],$ttm,tPO[IsSingleSource],1)"),
        @(14, 'Units received',             "=SUMIFS(tPO[QtyReceived],$ttm)"),
        @(15, 'Units accepted',             "=SUMIFS(tPO[QtyAccepted],$ttm)"),
        @(16, 'Value of rejected material', "=SUMIFS(tPO[RejectedValue],$ttm)")
    )
    foreach ($b in $base) {
        $dash.Cells.Item($b[0],1).Value2 = $b[1]
        $dash.Cells.Item($b[0],3).Formula = $b[2]
        $dash.Cells.Item($b[0],3).NumberFormat = '#,##0.00'
    }
    $dash.Range('C14:C15').NumberFormat = '#,##0'

    # the erosion block -- the centrepiece
    $dash.Range('E7').Value2 = 'PRICE EROSION -- what the learning curve should have handed back'
    $dash.Range('E7').Font.Bold = $true
    $ero = @(
        @(8,  'Vendor-part pairs assessed',       '=COUNTA(tErosion[PartNumber])'),
        @(9,  'Spend-weighted erosion capture %', '=SUMPRODUCT(tErosion[ErosionCapturePct],tErosion[TTMSpend])/SUM(tErosion[TTMSpend])'),
        @(10, 'Annual opportunity (gross)',       '=SUMIF(tErosion[AnnualOpportunity],">0")'),
        @(11, 'Pairs already ahead of the curve', '=COUNTIF(tErosion[AnnualOpportunity],"<0")'),
        @(12, 'Value where already ahead',        '=SUMIF(tErosion[AnnualOpportunity],"<0")'),
        @(13, 'Pairs capturing under 20%',        '=COUNTIFS(tErosion[ErosionCapturePct],"<20")'),
        @(14, 'Spend on those pairs',             '=SUMIFS(tErosion[TTMSpend],tErosion[ErosionCapturePct],"<20")')
    )
    foreach ($e in $ero) {
        $dash.Cells.Item($e[0],5).Value2 = $e[1]
        $dash.Cells.Item($e[0],7).Formula = $e[2]
        $dash.Cells.Item($e[0],7).NumberFormat = '#,##0.00'
    }
    $dash.Range('G8').NumberFormat = '#,##0'
    $dash.Range('G11').NumberFormat = '#,##0'
    $dash.Range('G13').NumberFormat = '#,##0'
    $dash.Range('H10').Value2 = 'Gross, not netted: a supplier already beating the curve is not a savings source and cannot fund another supplier''s shortfall.'
    $dash.Range('H10').Font.Italic = $true

    # scorecard
    $dash.Range('A18').Value2 = 'SCORECARD -- status computed against the target table, not typed in'
    $dash.Range('A18').Font.Bold = $true
    $kh = @('Metric','Value','Target','Warning','Status','Variance','What it means')
    for ($i = 0; $i -lt $kh.Count; $i++) { $dash.Cells.Item(19, $i + 1).Value2 = $kh[$i] }
    $dash.Range('A19:G19').Font.Bold = $true
    $dash.Range('A19:G19').Interior.Color = $Navy
    $dash.Range('A19:G19').Font.Color = $White

    # metric key, label, Excel formula, direction, commentary
    $kpis = @(
      @('ErosionCapturePct','Erosion capture %','=$G$9','H','Share of the expected category erosion actually realised. No price-variance report can see this gap.'),
      @('MaverickSpendPct','Maverick spend %','=100*$C$12/$C$8','L','Spend with no agreement in force. There is nothing to negotiate against until it is on contract.'),
      @('PPVPct','Purchase price variance %',"=100*SUMIFS(tPO[PPVAmount],$ttm)/SUMIFS(tPO[ContractedExtended],$ttm)",'L','A supplier who renews flat scores zero here forever. Read it beside erosion capture, never alone.'),
      @('AcceptanceRatePct','Acceptance rate %','=100*$C$15/$C$14','H','Portfolio figure. Check the vendor-category cut before concluding anything.'),
      @('OnTimeDeliveryPct','On-time delivery %',"=100*COUNTIFS(tPO[IsOnTime],1,$ttm)/COUNTIFS(tPO[IsOnTime],"">=0"",$ttm)",'H','Receipts on or before the promised date.'),
      @('ExpediteSpendPct','Expedite spend %','=100*$C$11/$C$8','L','Caused by planning, not by the supplier. Negotiable only by fixing the plan.'),
      @('SingleSourceSpendPct','Single-source spend %','=100*$C$13/$C$8','L','Where the largest erosion gaps sit and the least leverage exists.'),
      @('Top5VendorSharePct','Top-5 vendor share %','=100*xlTop5/xlWindowAll','L','Concentration of spend in the five largest suppliers.')
    )
    $kr = 20
    foreach ($k in $kpis) {
        $m = $k[0]
        $dash.Cells.Item($kr,1).Value2  = $k[1]
        $dash.Cells.Item($kr,2).Formula = $k[2]
        $dash.Cells.Item($kr,3).Formula = "=INDEX(tTargets[TargetValue],MATCH(""$m"",tTargets[MetricName],0))"
        $dash.Cells.Item($kr,4).Formula = "=INDEX(tTargets[WarningValue],MATCH(""$m"",tTargets[MetricName],0))"
        if ($k[3] -eq 'L') {
            $dash.Cells.Item($kr,5).Formula = "=IF(B$kr<=C$kr,""Green"",IF(B$kr<=D$kr,""Amber"",""Red""))"
            $dash.Cells.Item($kr,6).Formula = "=B$kr-C$kr"
        } else {
            $dash.Cells.Item($kr,5).Formula = "=IF(B$kr>=C$kr,""Green"",IF(B$kr>=D$kr,""Amber"",""Red""))"
            $dash.Cells.Item($kr,6).Formula = "=C$kr-B$kr"
        }
        $dash.Cells.Item($kr,7).Value2 = $k[4]
        $kr++
    }
    $kpiLast = $kr - 1
    $dash.Range("B20:D$kpiLast").NumberFormat = '0.00'
    $dash.Range("F20:F$kpiLast").NumberFormat = '+0.00;-0.00'
    $dash.Range("E20:E$kpiLast").HorizontalAlignment = -4108
    foreach ($rule in @(@('Green',$GreenBg,$GreenFg), @('Amber',$AmberBg,$AmberFg), @('Red',$RedBg,$RedFg))) {
        $fc = $dash.Range("E20:E$kpiLast").FormatConditions.Add(1, 3, """$($rule[0])""")
        $fc.Interior.Color = $rule[1]; $fc.Font.Color = $rule[2]; $fc.Font.Bold = $true
    }

    # category cut
    $catTop = $kpiLast + 2
    $dash.Cells.Item($catTop,1).Value2 = 'BY CATEGORY -- spend, benchmark and what was actually captured'
    $dash.Cells.Item($catTop,1).Font.Bold = $true
    $ch = @('Category','Spend (12m)','Share %','Benchmark erosion %','Captured %','Opportunity','Maverick %')
    for ($i = 0; $i -lt $ch.Count; $i++) { $dash.Cells.Item($catTop+1, $i+1).Value2 = $ch[$i] }
    $dash.Range($dash.Cells.Item($catTop+1,1), $dash.Cells.Item($catTop+1,7)).Font.Bold = $true
    $dash.Range($dash.Cells.Item($catTop+1,1), $dash.Cells.Item($catTop+1,7)).Interior.Color = $Grey
    for ($i = 0; $i -lt 8; $i++) {
        $rw = $catTop + 2 + $i; $n = $i + 1
        $dash.Cells.Item($rw,1).Formula = "=IFERROR(INDEX(tBenchmark[Category],$n),"""")"
        $dash.Cells.Item($rw,2).Formula = "=IF(`$A$rw="""","""",SUMIFS(tPO[ExtendedPrice],$ttm,tPO[Category],`$A$rw))"
        $dash.Cells.Item($rw,3).Formula = "=IF(`$A$rw="""","""",100*B$rw/`$C`$8)"
        $dash.Cells.Item($rw,4).Formula = "=IFERROR(INDEX(tBenchmark[AnnualErosionPct],$n),"""")"
        $dash.Cells.Item($rw,5).Formula = "=IF(`$A$rw="""","""",IFERROR(SUMPRODUCT((tErosion[Category]=`$A$rw)*tErosion[ErosionCapturePct]*tErosion[TTMSpend])/SUMPRODUCT((tErosion[Category]=`$A$rw)*tErosion[TTMSpend]),""""))"
        $dash.Cells.Item($rw,6).Formula = "=IF(`$A$rw="""","""",SUMIFS(tErosion[AnnualOpportunity],tErosion[Category],`$A$rw,tErosion[AnnualOpportunity],"">0""))"
        $dash.Cells.Item($rw,7).Formula = "=IF(`$A$rw="""","""",IFERROR(100*SUMIFS(tPO[ExtendedPrice],$ttm,tPO[Category],`$A$rw,tPO[IsOnContract],0)/SUMIFS(tPO[ExtendedPrice],$ttm,tPO[Category],`$A$rw),""""))"
    }
    $catTot = $catTop + 10
    $dash.Cells.Item($catTot,1).Value2 = 'TOTAL'
    $dash.Cells.Item($catTot,2).Formula = "=SUM(B$($catTop+2):B$($catTop+9))"
    $dash.Cells.Item($catTot,3).Formula = "=SUM(C$($catTop+2):C$($catTop+9))"
    $dash.Cells.Item($catTot,6).Formula = "=SUM(F$($catTop+2):F$($catTop+9))"
    $dash.Range($dash.Cells.Item($catTot,1), $dash.Cells.Item($catTot,7)).Font.Bold = $true
    $dash.Range($dash.Cells.Item($catTop+2,2), $dash.Cells.Item($catTot,2)).NumberFormat = '#,##0'
    $dash.Range($dash.Cells.Item($catTop+2,3), $dash.Cells.Item($catTot,3)).NumberFormat = '0.0'
    $dash.Range($dash.Cells.Item($catTop+2,4), $dash.Cells.Item($catTot,5)).NumberFormat = '0.00'
    $dash.Range($dash.Cells.Item($catTop+2,6), $dash.Cells.Item($catTot,6)).NumberFormat = '#,##0'
    $dash.Range($dash.Cells.Item($catTop+2,7), $dash.Cells.Item($catTot,7)).NumberFormat = '0.0'

    $widths = @(32,15,14,13,30,13,58,60)
    for ($i = 0; $i -lt $widths.Count; $i++) { $dash.Columns.Item($i+1).ColumnWidth = $widths[$i] }
    Write-Host "  Dashboard                built"

    # ----------------------------------------------- Renegotiation Queue
    $q.Range('A1').Value2 = 'RENEGOTIATION QUEUE'
    $q.Range('A1').Font.Size = 16; $q.Range('A1').Font.Bold = $true; $q.Range('A1').Font.Color = $Navy
    $q.Range('A2').Value2 = 'Which supplier-part combination to renegotiate first. Ranked by annual opportunity; the ACTION is decided by leverage, because the biggest gaps sit where there is least of it.'
    $q.Range('A2:P2').Merge(); $q.Range('A2').WrapText = $true; $q.Range('A2').RowHeight = 28

    $q.Range('B4').Value2 = 'Category manager'; $q.Range('C4').Value2 = '(All)'
    $q.Range('B5').Value2 = 'Action';           $q.Range('C5').Value2 = '(All)'
    $q.Range('B6').Value2 = 'Category';         $q.Range('C6').Value2 = '(All)'
    $q.Range('B7').Value2 = "This quarter only"; $q.Range('C7').Value2 = 'Yes'
    $q.Range('B4:B7').Font.Bold = $true
    $q.Range('C4:C7').Interior.Color = $NavyLt

    $qcsv = Import-Csv (Join-Path $ExtractPath 'renegotiation_queue.csv')
    $buyers  = @('(All)') + ($qcsv | Select-Object -ExpandProperty BuyerName -Unique | Sort-Object)
    $actions = @('(All)') + ($qcsv | Select-Object -ExpandProperty ActionCode -Unique | Sort-Object)
    $cats    = @('(All)') + ($qcsv | Select-Object -ExpandProperty Category -Unique | Sort-Object)
    $q.Range('C4').Validation.Add(3,1,1,($buyers -join ',')) | Out-Null
    $q.Range('C5').Validation.Add(3,1,1,($actions -join ',')) | Out-Null
    $q.Range('C6').Validation.Add(3,1,1,($cats -join ',')) | Out-Null
    $q.Range('C7').Validation.Add(3,1,1,'Yes,No') | Out-Null

    $q.Range('F4').Value2  = 'Pairs matching'
    $q.Range('H4').Formula = "=SUM(Calc!`$B`$$qFirst`:`$B`$$qLast)"
    $q.Range('F5').Value2  = 'Annual opportunity in scope'
    $q.Range('H5').Formula = "=SUMPRODUCT(Calc!`$B`$$qFirst`:`$B`$$qLast,tQueue[AnnualOpportunity])"
    $q.Range('F6').Value2  = 'Spend in scope (12m)'
    $q.Range('H6').Formula = "=SUMPRODUCT(Calc!`$B`$$qFirst`:`$B`$$qLast,tQueue[TTMSpend])"
    $q.Range('F7').Formula = '="Showing top "&MIN(QueueRowsShown,H4)&" of "&H4'
    $q.Range('F4:F6').Font.Bold = $true; $q.Range('F7').Font.Bold = $true
    $q.Range('H5:H6').NumberFormat = '#,##0.00'

    $qh = @('#','Part','Category','Vendor','Manager','Sources','Bench %','Actual %','Capture %',
            'Gap / unit','TTM spend','Opportunity','Leverage','Action','What to do')
    for ($i = 0; $i -lt $qh.Count; $i++) { $q.Cells.Item(9, $i+1).Value2 = $qh[$i] }
    $q.Range('A9:O9').Font.Bold = $true
    $q.Range('A9:O9').Interior.Color = $Navy
    $q.Range('A9:O9').Font.Color = $White
    $q.Range('A9:O9').WrapText = $true

    $qCols = @('PartNumber','Category','VendorID','BuyerName','QualifiedSupplierCount',
               'BenchmarkErosionPct','ActualErosionPct','ErosionCapturePct','ErosionGapPerUnit',
               'TTMSpend','AnnualOpportunity','LeverageScore','ActionCode','RecommendedAction')
    $qRows = 25
    for ($i = 0; $i -lt $qRows; $i++) {
        $rw = 10 + $i; $n = $i + 1
        $q.Cells.Item($rw,1).Formula = "=IF($n>MIN(QueueRowsShown,`$H`$4),"""",$n)"
        for ($c = 0; $c -lt $qCols.Count; $c++) {
            $q.Cells.Item($rw, $c+2).Formula =
                "=IFERROR(INDEX(tQueue[$($qCols[$c])],MATCH($n,Calc!`$C`$$qFirst`:`$C`$$qLast,0)),"""")"
        }
    }
    $q.Range("G10:J$(9+$qRows)").NumberFormat = '0.00'
    $q.Range("K10:L$(9+$qRows)").NumberFormat = '#,##0'
    $q.Range("M10:M$(9+$qRows)").NumberFormat = '0.0'
    # colour by action, computed rather than applied by hand
    $fcSole = $q.Range("A10:O$(9+$qRows)").FormatConditions.Add(2, 0, '=OR($N10="DUAL_SOURCE",$N10="QUALIFY_ALTERNATE")')
    $fcSole.Interior.Color = $AmberBg; $fcSole.Font.Color = $AmberFg
    $fcCon = $q.Range("A10:O$(9+$qRows)").FormatConditions.Add(2, 0, '=$N10="PUT_ON_CONTRACT"')
    $fcCon.Interior.Color = $RedBg; $fcCon.Font.Color = $RedFg
    $qw = @(5,14,17,10,17,9,10,10,11,12,14,14,11,19,72)
    for ($i = 0; $i -lt $qw.Count; $i++) { $q.Columns.Item($i+1).ColumnWidth = $qw[$i] }
    Write-Host "  Renegotiation Queue      built ($qRows rows shown)"

    # ------------------------------------------------------ Vendor Scorecard
    $vs.Range('A1').Value2 = 'VENDOR SCORECARD'
    $vs.Range('A1').Font.Size = 16; $vs.Range('A1').Font.Bold = $true; $vs.Range('A1').Font.Color = $Navy
    $vs.Range('A2').Value2 = 'Cost index compares each vendor with the cheapest cost per ACCEPTED unit on the same parts, so 100 is the best available and 108 is eight per cent dearer per usable part. Unit price alone cannot show this.'
    $vs.Range('A2:L2').Merge(); $vs.Range('A2').WrapText = $true; $vs.Range('A2').RowHeight = 28

    $vh = @('Vendor','Name','Tier','Spend (12m)','Our share of their revenue %','On contract %',
            'Acceptance %','Rejected value','On-time %','Cost index vs best','Shared parts','Read')
    for ($i = 0; $i -lt $vh.Count; $i++) { $vs.Cells.Item(4, $i+1).Value2 = $vh[$i] }
    $vs.Range('A4:L4').Font.Bold = $true
    $vs.Range('A4:L4').Interior.Color = $Navy
    $vs.Range('A4:L4').Font.Color = $White
    $vs.Range('A4:L4').WrapText = $true
    $nVendor = $loaded['tVendor']
    $vCols = @('VendorID','VendorName','VendorTier','TotalSpend','LumenRevenueSharePct',
               'OnContractPct','AcceptanceRatePct','RejectedValue','OnTimePct','CostIndexVsBest','SharedParts')
    for ($i = 0; $i -lt $nVendor; $i++) {
        $rw = 5 + $i; $n = $i + 1
        for ($c = 0; $c -lt $vCols.Count; $c++) {
            $vs.Cells.Item($rw, $c+1).Formula = "=IFERROR(INDEX(tVendor[$($vCols[$c])],$n),"""")"
        }
        $vs.Cells.Item($rw,12).Formula =
            "=IF(J$rw="""","""",IF(J$rw>105,""Dearer per usable part than the best alternative"",IF(G$rw<97,""Quality is the problem, not price"",""Competitive"")))"
    }
    $vs.Range("D5:D$(4+$nVendor)").NumberFormat = '#,##0'
    $vs.Range("E5:J$(4+$nVendor)").NumberFormat = '0.00'
    $vs.Range("H5:H$(4+$nVendor)").NumberFormat = '#,##0'
    $vw = @(10,30,14,15,16,13,13,15,11,14,13,48)
    for ($i = 0; $i -lt $vw.Count; $i++) { $vs.Columns.Item($i+1).ColumnWidth = $vw[$i] }
    Write-Host "  Vendor Scorecard         built ($nVendor vendors)"

    # -------------------------------------------------------------- Validation
    $val.Range('A1').Value2 = 'CROSS-IMPLEMENTATION VALIDATION'
    $val.Range('A1').Font.Size = 16; $val.Range('A1').Font.Bold = $true; $val.Range('A1').Font.Color = $Navy
    $val.Range('A2').Value2 = 'Each figure below was computed twice: once in T-SQL against the database, once in Excel formulas against the extract. A dashboard that only displays what the database told it cannot detect that the database was wrong.'
    $val.Range('A2:F2').Merge(); $val.Range('A2').WrapText = $true; $val.Range('A2').RowHeight = 30

    $vhh = @('#','Figure','SQL','Excel','Difference','Result')
    for ($i = 0; $i -lt $vhh.Count; $i++) { $val.Cells.Item(4, $i+1).Value2 = $vhh[$i] }
    $val.Range('A4:F4').Font.Bold = $true
    $val.Range('A4:F4').Interior.Color = $Navy
    $val.Range('A4:F4').Font.Color = $White

    $lastKPI = 'MATCH(MAX(tKPI[AsOfDate]),tKPI[AsOfDate],0)'
    $checks = @(
      @('Extended price (12 months)',  "=INDEX(tKPI[TotalSpend],$lastKPI)",        '=Dashboard!$C$8',  0.05),
      @('Landed cost (12 months)',     "=INDEX(tKPI[LandedSpend],$lastKPI)",       '=Dashboard!$C$9',  0.05),
      @('Erosion capture %',           "=INDEX(tKPI[ErosionCapturePct],$lastKPI)", '=Dashboard!$B$20', 0.01),
      @('Annual erosion opportunity',  "=INDEX(tKPI[ErosionOpportunity],$lastKPI)",'=Dashboard!$G$10', 0.05),
      @('Maverick spend %',            "=INDEX(tKPI[MaverickSpendPct],$lastKPI)",  '=Dashboard!$B$21', 0.01),
      @('Purchase price variance %',   "=INDEX(tKPI[PPVPct],$lastKPI)",            '=Dashboard!$B$22', 0.01),
      @('Acceptance rate %',           "=INDEX(tKPI[AcceptanceRatePct],$lastKPI)", '=Dashboard!$B$23', 0.01),
      @('On-time delivery %',          "=INDEX(tKPI[OnTimeDeliveryPct],$lastKPI)", '=Dashboard!$B$24', 0.01),
      @('Expedite spend %',            "=INDEX(tKPI[ExpediteSpendPct],$lastKPI)",  '=Dashboard!$B$25', 0.01),
      @('Single-source spend %',       "=INDEX(tKPI[SingleSourceSpendPct],$lastKPI)",'=Dashboard!$B$26', 0.01),
      @('Top-5 vendor share %',        "=INDEX(tKPI[Top5VendorSharePct],$lastKPI)",'=Dashboard!$B$27', 0.01),
      # This check compared SUM(tPO[RejectedValue]) against
      # SUMIFS(tPO[RejectedValue], tPO[OrderDate], ">"&MIN(tPO[OrderDate])-1).
      # That criterion is true for EVERY row -- a date is always greater than
      # the minimum date minus one day -- so the SUMIFS collapsed to the SUM
      # beside it. Both sides were the same arithmetic over the same column of
      # the same extract, so the difference was exactly 0.0000 under any data
      # and any defect. It was the only one of the fourteen that was not a
      # cross-implementation check, and it could not fail.
      #
      # It now compares SQL's own TTM figure from the KPI extract against the
      # Dashboard cell Excel computes from the purchase-order extract, which is
      # the pattern every other check here follows.
      @('Value of rejected material',  "=INDEX(tKPI[RejectedValue],$lastKPI)",     '=Dashboard!$C$16', 0.05),
      @('Queue: pairs assessed',       '=COUNTA(tQueue[PartNumber])',              '=COUNTA(tErosion[PartNumber])', 0.5),
      @('Queue: total opportunity',    '=SUM(tQueue[AnnualOpportunity])',          '=SUMIF(tErosion[AnnualOpportunity],">0")', 0.05)
    )
    $vr = 5
    foreach ($c in $checks) {
        $val.Cells.Item($vr,1).Formula = [string]($vr - 4)
        $val.Cells.Item($vr,2).Value2  = $c[0]
        $val.Cells.Item($vr,3).Formula = $c[1]
        $val.Cells.Item($vr,4).Formula = $c[2]
        $val.Cells.Item($vr,5).Formula = "=C$vr-D$vr"
        $val.Cells.Item($vr,6).Formula = "=IF(ABS(E$vr)<=$($c[3]),""MATCH"",""DIFFERS"")"
        $vr++
    }
    $vLast = $vr - 1
    $val.Range("C5:E$vLast").NumberFormat = '#,##0.0000'
    $val.Range("F5:F$vLast").HorizontalAlignment = -4108
    $fcM = $val.Range("F5:F$vLast").FormatConditions.Add(1,3,'"MATCH"')
    $fcM.Interior.Color = $GreenBg; $fcM.Font.Color = $GreenFg; $fcM.Font.Bold = $true
    $fcD = $val.Range("F5:F$vLast").FormatConditions.Add(1,3,'"DIFFERS"')
    $fcD.Interior.Color = $RedBg; $fcD.Font.Color = $RedFg; $fcD.Font.Bold = $true
    $val.Cells.Item($vLast+2,2).Value2  = 'Checks reconciling'
    $val.Cells.Item($vLast+2,3).Formula = "=COUNTIF(`$F`$5:`$F`$$vLast,""MATCH"")&"" of ""&COUNTA(`$F`$5:`$F`$$vLast)"
    $val.Cells.Item($vLast+2,2).Font.Bold = $true; $val.Cells.Item($vLast+2,3).Font.Bold = $true
    $val.Cells.Item($vLast+3,2).Formula = "=IF(COUNTIF(`$F`$5:`$F`$$vLast,""DIFFERS"")=0,""All figures reconcile between the two independent implementations."",""RECONCILIATION FAILED -- do not publish these figures."")"
    $val.Cells.Item($vLast+3,2).Font.Bold = $true
    $val.Columns.Item(1).ColumnWidth = 5
    $val.Columns.Item(2).ColumnWidth = 36
    $val.Columns.Item(3).ColumnWidth = 18
    $val.Columns.Item(4).ColumnWidth = 18
    $val.Columns.Item(5).ColumnWidth = 15
    $val.Columns.Item(6).ColumnWidth = 12
    Write-Host "  Validation               built ($($vLast - 4) checks)"

    # ----------------------------------------------------------------- README
    $lines = @(
      @('LUMEN OPTICS MANUFACTURING -- PHOTONICS SPEND SCORECARD','title'),
      @('',''),
      @('Lumen Optics Manufacturing is a fictional photonics manufacturer. Every figure in this workbook is synthetic, generated by the SQL scripts in the sql\ folder of this project. No confidential data is used and no claim is made about any production system.','body'),
      @('',''),
      @('THE QUESTION THIS ANSWERS','head'),
      @('Our purchase prices are flat or falling and every variance report is green -- so why is material cost per ACCEPTED unit rising, and which supplier-part combinations should sourcing renegotiate first?','body'),
      @('',''),
      @('WHY PRICE VARIANCE CANNOT ANSWER IT','head'),
      @('Photonic components follow learning curves: unit prices decline as volumes rise and processes mature. Purchase price variance compares what was paid against what was AGREED, so a supplier who renews flat every year scores zero variance forever while quietly keeping everything the curve was supposed to hand back. Erosion capture measures that gap; price variance is structurally blind to it.','body'),
      @('',''),
      @('SHEETS','head'),
      @('Dashboard            The scorecard, the erosion position and the category cut. Every number is a formula over the extract.','body'),
      @('Renegotiation Queue  Which supplier-part pair to open first, with the action and the annual value. Filters re-rank the list.','body'),
      @('Vendor Scorecard     Cost index against the best available on the same parts, plus quality, delivery and compliance.','body'),
      @('Validation           Each headline figure computed twice -- once in SQL, once in Excel -- and reconciled.','body'),
      @('Control              Every parameter the workbook obeys.','body'),
      @('Calc                 The working behind the twelve-month window, the concentration and the queue ranking. Visible on purpose.','body'),
      @('Data_*               The loaded extracts. Do not edit; they are replaced on refresh.','body'),
      @('',''),
      @('HOW TO REFRESH','head'),
      @('1. Rebuild the extracts:  powershell erp_extracts\Export_Extracts.ps1','body'),
      @('2. If the repo has moved, update the Extract folder cell on the Control sheet.','body'),
      @('3. Data > Refresh All.','body'),
      @('4. Check the Validation sheet still reads "All figures reconcile".','body'),
      @('',''),
      @('WHAT THE MODEL DOES NOT DO','head'),
      @('Erosion is measured from the FIRST OBSERVED PRICE, so it detects a failure to erode, not whether the opening price was fair. A part bought badly in 2023 and eroded perfectly since will look healthy here. Answering that needs a should-cost build-up from materials and process, which is a different exercise with different inputs.','body')
    )
    $rr = 1
    foreach ($l in $lines) {
        $rd.Cells.Item($rr,1).Value2 = $l[0]
        switch ($l[1]) {
            'title' { $rd.Cells.Item($rr,1).Font.Size = 16; $rd.Cells.Item($rr,1).Font.Bold = $true; $rd.Cells.Item($rr,1).Font.Color = $Navy }
            'head'  { $rd.Cells.Item($rr,1).Font.Bold = $true; $rd.Cells.Item($rr,1).Font.Color = $Navy }
        }
        $rr++
    }
    $rd.Columns.Item(1).ColumnWidth = 140

    # ------------------------------------------------------------------ finish
    # Self-locating folder: shipping an absolute path would bake one machine's
    # directory -- and its username -- into a public file.
    $ctl.Range('B6').Formula =
        '=IF(OR(CELL("filename")="",LEFT(CELL("filename"),4)="http"),' +
        '"Set this cell to the full path of the erp_extracts folder, then Data > Refresh All",' +
        'LEFT(CELL("filename"),FIND("[",CELL("filename"))-1)&"..\erp_extracts")'

    $finalOrder = @('README','Dashboard','Renegotiation Queue','Vendor Scorecard','Validation','Calc','Control')
    for ($i = $finalOrder.Count - 1; $i -ge 0; $i--) {
        $wb.Worksheets.Item($finalOrder[$i]).Move($wb.Worksheets.Item(1))
    }
    $wb.Worksheets.Item('README').Activate()
    $xl.Calculation = -4105
    $wb.Application.CalculateFullRebuild()

    foreach ($prop in @(@('Author','Lumen Sourcing Analytics'),
                        @('Company','Lumen Optics Manufacturing (fictional)'),
                        @('Title','Lumen Optics -- Photonics Spend Scorecard'),
                        @('Comments','Synthetic demonstration data. No confidential data and no production system.'))) {
        try { $wb.BuiltinDocumentProperties($prop[0]).Value = $prop[1] } catch { }
    }
    try { $wb.RemovePersonalInformation = $true } catch { }

    if (Test-Path $OutputFile) { Remove-Item $OutputFile -Force }
    $wb.SaveAs($OutputFile, 51)
    Write-Host ("-" * 78)
    Write-Host "  Saved: $OutputFile" -ForegroundColor Green
    $wb.Close($false)
}
catch {
    Write-Host ("FAILED at line {0}: {1}" -f $_.InvocationInfo.ScriptLineNumber, $_.Exception.Message) -ForegroundColor Red
    Write-Host ("  >> " + $_.InvocationInfo.Line.Trim()) -ForegroundColor DarkYellow
    throw
}
finally {
    $xl.Quit()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
Write-Host ""
