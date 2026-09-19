<#
================================================================================
Project 2 -- Vantage Wholesale Supply: Receivables Performance
Script:  Build_Workbook.ps1
Purpose: Build Vantage_AR_Control.xlsx -- the operational control the collections
         team actually works from.

WHAT MAKES THIS A CONTROL AND NOT A REPORT
    Nothing on the Dashboard is a pasted number. Every KPI is recomputed by
    formula from the open-item extract: SUMIFS for the balances, a real
    countback loop for DSO, INDEX/MATCH against the target table for the RAG
    status. The Validation sheet then compares each Excel-computed figure with
    the figure SQL produced independently, and shows the difference. If the two
    implementations ever disagree, the workbook says so on its own face.

    That is the whole point. A dashboard that only displays what a database
    told it cannot detect that the database was wrong. Two independent
    implementations that reconcile to the cent can.

WHY NO DYNAMIC ARRAYS
    FILTER, SORT and UNIQUE would make several of these formulas shorter, and
    would break the workbook for anyone on Excel 2019 or LibreOffice. The queue
    ranking therefore uses SUMPRODUCT and INDEX/MATCH, which work everywhere
    and are the technique an analyst actually needs when a file has to open on
    someone else's desk.

WHY POWER QUERY READS A FOLDER FROM A CELL
    An absolute path baked into M breaks the moment the repo is cloned
    somewhere else. Every query here resolves its folder from the named cell
    SourceFolder on the Control sheet, so relocating the workbook is a
    one-cell edit rather than eight query edits.

DATA DISCLOSURE: Vantage Wholesale Supply is fictional; all data is synthetic.
No confidential data and no production system is involved.
================================================================================
#>

param(
    [string]$ProjectRoot = (Split-Path $PSScriptRoot -Parent),
    [string]$OutputFile  = (Join-Path $PSScriptRoot 'Vantage_AR_Control.xlsx')
)

$ErrorActionPreference = 'Stop'
$ExtractPath = Join-Path $ProjectRoot 'erp_extracts'

if (-not (Test-Path (Join-Path $ExtractPath 'ar_open_items.csv'))) {
    throw "Extracts not found in $ExtractPath. Run erp_extracts\Export_Extracts.ps1 first."
}

# ---- palette (Excel wants BGR packed into an integer) -----------------------
function RGBv([int]$r, [int]$g, [int]$b) { return $r + ($g * 256) + ($b * 65536) }
$Navy      = RGBv 31 56 100
$NavyLight = RGBv 217 225 242
$Grey      = RGBv 242 242 242
$Border    = RGBv 191 191 191
$GreenBg   = RGBv 198 239 206 ; $GreenFg = RGBv 0 97 0
$AmberBg   = RGBv 255 235 156 ; $AmberFg = RGBv 156 101 0
$RedBg     = RGBv 255 199 206 ; $RedFg   = RGBv 156 0 6
$White     = RGBv 255 255 255

Write-Host ""
Write-Host "Building Vantage_AR_Control.xlsx" -ForegroundColor Cyan
Write-Host ("-" * 78)

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.ScreenUpdating = $false

try {
    $wb = $xl.Workbooks.Add()
    # start from a single sheet; everything else is added deliberately
    while ($wb.Worksheets.Count -gt 1) { $wb.Worksheets.Item($wb.Worksheets.Count).Delete() }

    # =========================================================================
    # 1. CONTROL SHEET -- every parameter the workbook obeys, in one place
    # =========================================================================
    $ctl = $wb.Worksheets.Item(1)
    $ctl.Name = 'Control'

    $ctl.Range('A1').Value2 = 'VANTAGE WHOLESALE SUPPLY -- RECEIVABLES CONTROL'
    $ctl.Range('A1').Font.Size = 16
    $ctl.Range('A1').Font.Bold = $true
    $ctl.Range('A1').Font.Color = $Navy
    $ctl.Range('A2').Value2 = 'Control panel. Every sheet in this workbook reads its parameters from the cells below.'
    $ctl.Range('A2').Font.Italic = $true

    $ctl.Range('A4').Value2 = 'PARAMETER'
    $ctl.Range('B4').Value2 = 'VALUE'
    $ctl.Range('C4').Value2 = 'WHAT IT CONTROLS'
    $ctl.Range('A4:C4').Font.Bold = $true
    $ctl.Range('A4:C4').Interior.Color = $Navy
    $ctl.Range('A4:C4').Font.Color = $White

    $params = @(
        @('Reporting date',        'AsOfDate',                'DATE',   '2025-12-31', 'The as-of date for every balance, ageing bucket and KPI.'),
        @('Extract folder',        'SourceFolder',            'TEXT',   $ExtractPath, 'Where Power Query looks for the CSV extracts. Edit this cell if you move the repo, then Data > Refresh All.'),
        @('Calls per collector',   'CallsPerCollectorPerDay', 'NUM',    10,           'Capacity assumption behind the daily worklist. Raising it lengthens every collector''s list.'),
        @('Dispute discount',      'DisputeDiscount',         'PCT',    0.75,         'How much of a disputed balance is treated as un-chaseable by a phone call.'),
        @('90+ threshold (days)',  'Bucket90Threshold',       'NUM',    90,           'Days past due at which a balance counts as seriously delinquent.'),
        @('Queue rows shown',      'QueueRowsShown',          'NUM',    25,           'How many ranked accounts the Priority Queue sheet displays.')
    )
    $r = 5
    foreach ($p in $params) {
        $ctl.Range("A$r").Value2 = [string]$p[0]
        if ($p[2] -eq 'DATE') {
            # written as a formula rather than a marshalled DateTime: it survives
            # the COM boundary intact and is legible in the cell itself
            $d = [datetime]::Parse([string]$p[3])
            $ctl.Range("B$r").NumberFormat = 'yyyy-mm-dd'
            $ctl.Range("B$r").Formula = "=DATE($($d.Year),$($d.Month),$($d.Day))"
        } else {
            $ctl.Range("B$r").Formula = [string]$p[3]
            if ($p[2] -eq 'PCT') { $ctl.Range("B$r").NumberFormat = '0%' }
        }
        $ctl.Range("C$r").Value2 = [string]$p[4]
        $ctl.Range("B$r").Interior.Color = $NavyLight
        $ctl.Range("B$r").Font.Bold = $true
        $wb.Names.Add($p[1], $ctl.Range("B$r")) | Out-Null
        $r++
    }

    $ctl.Range('A13').Value2 = 'DATA DISCLOSURE'
    $ctl.Range('A13').Font.Bold = $true
    $ctl.Range('A13').Font.Color = $Navy
    $ctl.Range('A14').Value2 = 'Vantage Wholesale Supply is a fictional company. All data in this workbook is synthetic and generated by the scripts in the sql\ folder. No confidential data is used and no claim is made about any production system.'
    $ctl.Range('A14:C16').Merge()
    $ctl.Range('A14').WrapText = $true
    $ctl.Range('A14').VerticalAlignment = -4160

    $ctl.Columns.Item(1).ColumnWidth = 24
    $ctl.Columns.Item(2).ColumnWidth = 46
    $ctl.Columns.Item(3).ColumnWidth = 78
    Write-Host "  Control sheet            built"

    # =========================================================================
    # 2. POWER QUERY -- one query per extract, folder resolved from a cell
    # =========================================================================
    function New-Extract-Query {
        param([string]$QueryName, [string]$FileName, [string]$TypeList)
        $m = @"
let
    Folder    = Excel.CurrentWorkbook(){[Name="SourceFolder"]}[Content]{0}[Column1],
    Source    = Csv.Document(File.Contents(Folder & "\$FileName"), [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.Csv]),
    Promoted  = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    Typed     = Table.TransformColumnTypes(Promoted, {$TypeList})
in
    Typed
"@
        $wb.Queries.Add($QueryName, $m) | Out-Null
    }

    # Types are declared explicitly. Letting Power Query guess means a column of
    # invoice numbers that happen to look numeric arrives as a number in one
    # refresh and text in the next, and every lookup against it breaks.
    $queries = @(
        @{ Name='tOpenItems'; File='ar_open_items.csv'; Types='{"InvoiceNo",type text},{"CustomerID",type text},{"CustomerName",type text},{"Segment",type text},{"Region",type text},{"SalesRep",type text},{"RiskTier",type text},{"CreditLimit",type number},{"CollectorID",type text},{"CollectorName",type text},{"Team",type text},{"TermsCode",type text},{"NetDays",Int64.Type},{"DiscountPct",type number},{"InvoiceDate",type date},{"DueDate",type date},{"InvoiceAmount",type number},{"PaidToDate",type number},{"DiscountToDate",type number},{"CreditMemoToDate",type number},{"WriteOffToDate",type number},{"OpenBalance",type number},{"OverAppliedAmount",type number},{"IsOpen",Int64.Type},{"IsPastDue",Int64.Type},{"DaysPastDue",Int64.Type},{"BucketKey",Int64.Type},{"BucketName",type text},{"RiskWeight",type number},{"IsDisputedOpen",Int64.Type},{"DisputeReason",type text}' },
        @{ Name='tQueue'; File='priority_action_queue.csv'; Types='{"PriorityRank",Int64.Type},{"CollectorRank",Int64.Type},{"CustomerID",type text},{"CustomerName",type text},{"Segment",type text},{"Region",type text},{"RiskTier",type text},{"CollectorID",type text},{"CollectorName",type text},{"Team",type text},{"TermsCode",type text},{"OpenBalance",type number},{"PastDueBalance",type number},{"Balance90Plus",type number},{"DisputedBalance",type number},{"WeightedExposure",type number},{"DisputedWeightedExposure",type number},{"CollectableExposure",type number},{"NetCollectablePosition",type number},{"IsFullyOffsetByCash",Int64.Type},{"OldestDaysPastDue",Int64.Type},{"CreditUtilizationPct",type number},{"AvgDaysLateHistoric",type number},{"UnappliedCash",type number},{"NetExposure",type number},{"BrokenPromises90d",Int64.Type},{"ActionCode",type text},{"CreditHoldFlag",Int64.Type},{"IsTodaysWorklist",Int64.Type},{"RecommendedAction",type text}' },
        @{ Name='tKPI'; File='ar_kpi_monthly.csv'; Types='{"YearMonth",type text},{"AsOfDate",type date},{"BeginAR",type number},{"MonthSales",type number},{"EndAR",type number},{"EndCurrentAR",type number},{"PastDueAR",type number},{"AR90Plus",type number},{"DisputedAR",type number},{"OpenInvoices",Int64.Type},{"PastDueInvoices",Int64.Type},{"Invoices90Plus",Int64.Type},{"CashApplied",type number},{"CashReceived",type number},{"CashAppliedLessReceived",type number},{"UnappliedCash",type number},{"UnappliedCashPct",type number},{"WriteOffs",type number},{"CreditMemos",type number},{"DSO_Countback",type number},{"DSO_Simple",type number},{"BPDSO_Countback",type number},{"AvgDaysDelinquent",type number},{"DSO_MethodGap",type number},{"CollectableBase",type number},{"CEI_Book",type number},{"CEI_Cash",type number},{"PaperCollectionsGap",type number},{"PctPastDue",type number},{"Pct90Plus",type number},{"PctDisputed",type number},{"PctPastDueByCount",type number},{"Pct90PlusByCount",type number},{"IsComparablePeriod",Int64.Type}' },
        @{ Name='tBridge'; File='dso_bridge_monthly.csv'; Types='{"AsOfDate",type date},{"TotalAR",type number},{"CurrentAR",type number},{"DisputedPastDueAR",type number},{"UndisputedPastDueAR",type number},{"SalesPerDay",type number},{"DSO_Classic",type number},{"GrantedDays",type number},{"DisputeDays",type number},{"LatenessDays",type number},{"WeightedAvgTermsDays",type number},{"TermsPositionDays",type number},{"RecoverableDays",type number},{"RecoverableCash",type number},{"BillingLagDays",type number},{"CashCycleDays",type number},{"CashPerLagDay",type number},{"IsComparablePeriod",Int64.Type}' },
        @{ Name='tSales'; File='monthly_credit_sales.csv'; Types='{"YearMonth",type text},{"MonthEndDate",type date},{"DaysInMonth",Int64.Type},{"CreditSales",type number},{"InvoiceCount",Int64.Type},{"WeightedAvgTermsDays",type number},{"PctDollarsNet30",type number},{"PctDollarsNet45",type number},{"PctDollarsNet60",type number},{"PctDollarsExtended",type number}' },
        @{ Name='tTargets'; File='ar_targets.csv'; Types='{"MetricName",type text},{"TargetValue",type number},{"WarningValue",type number},{"Direction",type text},{"Unit",type text},{"Description",type text}' },
        @{ Name='tBuckets'; File='aging_buckets.csv'; Types='{"BucketKey",Int64.Type},{"BucketName",type text},{"MinDaysPastDue",Int64.Type},{"MaxDaysPastDue",Int64.Type},{"RiskWeight",type number}' },
        @{ Name='tCashApp'; File='cash_application_worklist.csv'; Types='{"CustomerKey",Int64.Type},{"CustomerID",type text},{"CustomerName",type text},{"Segment",type text},{"Region",type text},{"CollectorName",type text},{"UnappliedCash",type number},{"UnappliedReceipts",Int64.Type},{"FullyUnappliedReceipts",Int64.Type},{"NoRemittanceAdvice",Int64.Type},{"OldestUnappliedDays",Int64.Type},{"OpenBalance",type number},{"PastDueBalance",type number},{"NetExposure",type number},{"WronglyChaseable",type number}' },
        @{ Name='tPromise'; File='promise_status.csv'; Types='{"PromiseNo",type text},{"CustomerID",type text},{"CustomerName",type text},{"RiskTier",type text},{"Segment",type text},{"CollectorName",type text},{"PromiseMadeDate",type date},{"PromisedPayDate",type date},{"PromisedAmount",type number},{"CashInWindow",type number},{"PctOfPromise",type number},{"DaysSincePromised",Int64.Type},{"PromiseStatus",type text}' },
        @{ Name='tLag'; File='billing_lag_monthly.csv'; Types='{"YearMonth",type text},{"MonthEndDate",type date},{"Region",type text},{"InvoiceCount",Int64.Type},{"CreditSales",type number},{"BillingLagDays",type number}' },
        @{ Name='tDQ'; File='data_quality_summary.csv'; Types='{"AnomalyType",type text},{"Severity",type text},{"ImpactClass",type text},{"Anomalies",Int64.Type},{"AmountAtRisk",type number},{"PopulationScanned",Int64.Type},{"AnomalyRatePct",type number},{"CountsTowardExposure",Int64.Type}' }
    )

    foreach ($q in $queries) { New-Extract-Query -QueryName $q.Name -FileName $q.File -TypeList $q.Types }
    Write-Host "  Power Query queries      $($queries.Count) created"

    # ---- load each query onto its own sheet as a table ----------------------
    $loaded = @{}
    foreach ($q in $queries) {
        $ws = $wb.Worksheets.Add([System.Reflection.Missing]::Value, $wb.Worksheets.Item($wb.Worksheets.Count))
        $ws.Name = 'Data_' + $q.Name.Substring(1)
        $connStr = "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=`$Workbook`$;Location=$($q.Name);Extended Properties="""""
        $lo = $ws.ListObjects.Add(0, $connStr, $false, 1, $ws.Range('A1'))
        $lo.QueryTable.CommandType = 2                # xlCmdSql
        $lo.QueryTable.CommandText = "SELECT * FROM [$($q.Name)]"
        $lo.QueryTable.BackgroundQuery = $false
        $lo.QueryTable.AdjustColumnWidth = $true
        $lo.QueryTable.Refresh($false) | Out-Null
        $lo.Name = $q.Name
        $rows = $lo.ListRows.Count
        $loaded[$q.Name] = $rows
        $ws.Rows.Item(1).Font.Bold = $true
        Write-Host ("    {0,-12} {1,6} rows loaded" -f $q.Name, $rows)
    }

    if ($loaded['tOpenItems'] -lt 100) { throw "tOpenItems loaded only $($loaded['tOpenItems']) rows -- the Power Query load did not work." }

    # -------------------------------------------------------------------------
    # Create every remaining sheet BEFORE writing a single formula.
    #
    # Excel does not defer a reference to a sheet that does not exist yet. Write
    # ='Priority Queue'!$C$4 while that sheet is still to come and Excel accepts
    # the formula, silently rewrites the reference to #REF!, and every formula
    # downstream inherits the error. The build still reports success.
    #
    # The sheets are created here in their final left-to-right order, then
    # populated below.
    # -------------------------------------------------------------------------
    $sheetOrder = @('README','Dashboard','Priority Queue','Cash Application','Ageing Matrix','Validation','Calc')
    $made = @{}
    foreach ($name in $sheetOrder) {
        $ws = $wb.Worksheets.Add([System.Reflection.Missing]::Value, $wb.Worksheets.Item($wb.Worksheets.Count))
        $ws.Name = $name
        $made[$name] = $ws
    }
    $rd   = $made['README']
    $dash = $made['Dashboard']
    $pq   = $made['Priority Queue']
    $am   = $made['Ageing Matrix']
    $val  = $made['Validation']
    $ca   = $made['Cash Application']
    $calc = $made['Calc']
    Write-Host "  Worksheets               $($sheetOrder.Count) created up front"

    $nOpen  = $loaded['tOpenItems']
    $nQueue = $loaded['tQueue']
    $nKPI   = $loaded['tKPI']

    # =========================================================================
    # 3. CALC SHEET -- the working. Visible on purpose: a reviewer should be
    #    able to see how the countback loop and the queue ranking are done.
    # =========================================================================
    $calc.Range('A1').Value2 = 'WORKING -- countback DSO and queue ranking'
    $calc.Range('A1').Font.Bold = $true
    $calc.Range('A1').Font.Size = 13
    $calc.Range('A1').Font.Color = $Navy
    $calc.Range('A2').Value2 = 'Nothing here is decorative. The Dashboard reads these cells.'
    $calc.Range('A2').Font.Italic = $true

    # ---- 3a. countback DSO -------------------------------------------------
    $calc.Range('A4').Value2 = 'COUNTBACK DSO -- consume the open balance against the most recent months of sales'
    $calc.Range('A4').Font.Bold = $true
    $hdr = @('Month end','Credit sales','Days in month','Cumulative sales','Cumulative days','Sales before this month','Days before this month','Covers total AR?','Covers current AR?')
    for ($i = 0; $i -lt $hdr.Count; $i++) { $calc.Cells.Item(5, $i + 1).Value2 = $hdr[$i] }
    $calc.Range('A5:I5').Font.Bold = $true
    $calc.Range('A5:I5').Interior.Color = $Grey
    $calc.Range('A5:I5').WrapText = $true

    $cbFirst = 6; $cbLast = 17          # 12 months
    for ($i = 0; $i -lt 12; $i++) {
        $rw = $cbFirst + $i
        $n  = $i + 1
        $calc.Cells.Item($rw,1).Formula = "=AGGREGATE(14,6,tSales[MonthEndDate]/(tSales[MonthEndDate]<=AsOfDate),$n)"
        $calc.Cells.Item($rw,1).NumberFormat = 'yyyy-mm-dd'
        $calc.Cells.Item($rw,2).Formula = "=SUMIFS(tSales[CreditSales],tSales[MonthEndDate],`$A$rw)"
        $calc.Cells.Item($rw,3).Formula = "=SUMIFS(tSales[DaysInMonth],tSales[MonthEndDate],`$A$rw)"
        $calc.Cells.Item($rw,4).Formula = "=SUM(`$B`$$cbFirst`:`$B$rw)"
        $calc.Cells.Item($rw,5).Formula = "=SUM(`$C`$$cbFirst`:`$C$rw)"
        $calc.Cells.Item($rw,6).Formula = "=`$D$rw-`$B$rw"
        $calc.Cells.Item($rw,7).Formula = "=`$E$rw-`$C$rw"
        $calc.Cells.Item($rw,8).Formula = "=IF(`$D$rw>=Dashboard!`$C`$8,1,0)"
        $calc.Cells.Item($rw,9).Formula = "=IF(`$D$rw>=Dashboard!`$C`$9,1,0)"
    }
    $calc.Range("B$cbFirst`:G$cbLast").NumberFormat = '#,##0.00'

    $calc.Range('A19').Value2 = 'Month that exhausts total AR'
    $calc.Range('C19').Formula = "=MATCH(1,`$H`$$cbFirst`:`$H`$$cbLast,0)"
    $calc.Range('A20').Value2 = 'Countback DSO (days)'
    $calc.Range('C20').Formula = "=INDEX(`$G`$$cbFirst`:`$G`$$cbLast,`$C`$19)+(Dashboard!`$C`$8-INDEX(`$F`$$cbFirst`:`$F`$$cbLast,`$C`$19))/INDEX(`$B`$$cbFirst`:`$B`$$cbLast,`$C`$19)*INDEX(`$C`$$cbFirst`:`$C`$$cbLast,`$C`$19)"
    $calc.Range('A21').Value2 = 'Month that exhausts current AR'
    $calc.Range('C21').Formula = "=MATCH(1,`$I`$$cbFirst`:`$I`$$cbLast,0)"
    $calc.Range('A22').Value2 = 'Best Possible DSO (days)'
    $calc.Range('C22').Formula = "=INDEX(`$G`$$cbFirst`:`$G`$$cbLast,`$C`$21)+(Dashboard!`$C`$9-INDEX(`$F`$$cbFirst`:`$F`$$cbLast,`$C`$21))/INDEX(`$B`$$cbFirst`:`$B`$$cbLast,`$C`$21)*INDEX(`$C`$$cbFirst`:`$C`$$cbLast,`$C`$21)"
    $calc.Range('A23').Value2 = 'Average Days Delinquent (days)'
    # Derived from the two figures as PUBLISHED, matching the SQL rule: a reader
    # who subtracts the two rows above must get this number exactly.
    $calc.Range('C23').Formula = '=ROUND(C20,2)-ROUND(C22,2)'
    $calc.Range('A19:A23').Font.Bold = $true
    $calc.Range('C19:C23').NumberFormat = '0.00'
    $calc.Range('C19').NumberFormat = '0'
    $calc.Range('C21').NumberFormat = '0'
    $wb.Names.Add('xlDSOCountback', $calc.Range('C20')) | Out-Null
    $wb.Names.Add('xlBPDSO',        $calc.Range('C22')) | Out-Null
    $wb.Names.Add('xlADD',          $calc.Range('C23')) | Out-Null

    # ---- 3b. queue ranking under the sheet filters -------------------------
    $qFirst = 28
    $calc.Range("A26").Value2 = 'QUEUE RANKING -- re-ranks the SQL queue under whatever filters the Priority Queue sheet has set'
    $calc.Range("A26").Font.Bold = $true
    $calc.Cells.Item(27,1).Value2 = 'Row'
    $calc.Cells.Item(27,2).Value2 = 'Matches filters'
    $calc.Cells.Item(27,3).Value2 = 'Rank within filter'
    $calc.Range("A27:C27").Font.Bold = $true
    $calc.Range("A27:C27").Interior.Color = $Grey
    $qLast = $qFirst + $nQueue - 1
    for ($i = 0; $i -lt $nQueue; $i++) {
        $rw = $qFirst + $i
        $n  = $i + 1
        $calc.Cells.Item($rw,1).Formula = [string]$n
        $calc.Cells.Item($rw,2).Formula = "=IF(AND(" +
            "OR('Priority Queue'!`$C`$4=""(All)"",INDEX(tQueue[CollectorName],$n)='Priority Queue'!`$C`$4)," +
            "OR('Priority Queue'!`$C`$5=""(All)"",INDEX(tQueue[ActionCode],$n)='Priority Queue'!`$C`$5)," +
            "OR('Priority Queue'!`$C`$6=""No"",INDEX(tQueue[IsTodaysWorklist],$n)=1)," +
            "INDEX(tQueue[PastDueBalance],$n)>='Priority Queue'!`$C`$7),1,0)"
        $calc.Cells.Item($rw,3).Formula = "=IF(`$B$rw=0,"""",SUMPRODUCT((`$B`$$qFirst`:`$B`$$qLast=1)*(tQueue[PriorityRank]<INDEX(tQueue[PriorityRank],$n)))+1)"
    }
    $calc.Columns.Item(1).ColumnWidth = 30
    Write-Host "  Calc sheet               built ($nQueue queue rows, 12 countback months)"

    # =========================================================================
    # 4. DASHBOARD
    # =========================================================================

    $dash.Range('A1').Value2 = 'VANTAGE WHOLESALE SUPPLY -- RECEIVABLES PERFORMANCE'
    $dash.Range('A1').Font.Size = 18
    $dash.Range('A1').Font.Bold = $true
    $dash.Range('A1').Font.Color = $Navy
    $dash.Range('A2').Formula = '="Reporting date: "&TEXT(AsOfDate,"d mmmm yyyy")&"   |   Fictional company, synthetic data, no production system."'
    $dash.Range('A2').Font.Italic = $true

    $dash.Range('A4').Value2 = 'THE QUESTION'
    $dash.Range('A4').Font.Bold = $true
    $dash.Range('A4').Font.Color = $White
    $dash.Range('A4:H4').Interior.Color = $Navy
    $dash.Range('A5').Value2 = 'DSO has risen. How many of those days did Vantage GRANT itself through the terms it sold, how many are customers TAKING beyond those terms -- and who should collections call first?'
    $dash.Range('A5:H5').Merge()
    $dash.Range('A5').WrapText = $true
    $dash.Range('A5').RowHeight = 30

    # ---- balances recomputed from the open-item extract --------------------
    $dash.Range('A7').Value2 = 'BALANCES -- recomputed in Excel from the open-item extract'
    $dash.Range('A7').Font.Bold = $true
    $bal = @(
        @(8,  'Total open AR',              '=SUM(tOpenItems[OpenBalance])'),
        @(9,  'Not yet due (granted)',      '=SUMIFS(tOpenItems[OpenBalance],tOpenItems[DaysPastDue],"<=0")'),
        @(10, 'Past due, disputed',         '=SUMIFS(tOpenItems[OpenBalance],tOpenItems[DaysPastDue],">0",tOpenItems[IsDisputedOpen],1)'),
        @(11, 'Past due, undisputed',       '=SUMIFS(tOpenItems[OpenBalance],tOpenItems[DaysPastDue],">0",tOpenItems[IsDisputedOpen],0)'),
        @(12, 'Of which 90+ days',          '=SUMIFS(tOpenItems[OpenBalance],tOpenItems[DaysPastDue],">"&Bucket90Threshold)'),
        @(13, 'Open invoices',              '=COUNTA(tOpenItems[InvoiceNo])'),
        @(14, 'Credit sales, last 12 mths', '=SUMIFS(tSales[CreditSales],tSales[MonthEndDate],">"&EOMONTH(AsOfDate,-12),tSales[MonthEndDate],"<="&AsOfDate)'),
        @(15, 'Days in that window',        '=SUMIFS(tSales[DaysInMonth],tSales[MonthEndDate],">"&EOMONTH(AsOfDate,-12),tSales[MonthEndDate],"<="&AsOfDate)'),
        @(16, 'Credit sales per day',       '=C14/C15')
    )
    foreach ($b in $bal) {
        $dash.Cells.Item($b[0],1).Value2 = $b[1]
        $dash.Cells.Item($b[0],3).Formula = $b[2]
        $dash.Cells.Item($b[0],3).NumberFormat = '#,##0.00'
    }
    $dash.Range('C13').NumberFormat = '#,##0'
    $dash.Range('C15').NumberFormat = '#,##0'

    # ---- the DSO bridge ----------------------------------------------------
    $dash.Range('E7').Value2 = 'DSO BRIDGE -- granted versus taken (these three add to classic DSO exactly)'
    $dash.Range('E7').Font.Bold = $true
    $dash.Range('E8').Value2 = 'Component'      ; $dash.Range('F8').Value2 = 'Days'
    $dash.Range('G8').Value2 = 'Cash'           ; $dash.Range('H8').Value2 = 'Who owns it'
    $dash.Range('E8:H8').Font.Bold = $true
    $dash.Range('E8:H8').Interior.Color = $Grey

    $dash.Range('E9').Value2  = 'Granted -- not yet due'
    $dash.Range('F9').Formula = '=C9*$C$15/$C$14'
    $dash.Range('G9').Formula = '=C9'
    $dash.Range('H9').Value2  = 'Sales / terms policy'
    $dash.Range('E10').Value2  = 'Taken -- disputed'
    $dash.Range('F10').Formula = '=C10*$C$15/$C$14'
    $dash.Range('G10').Formula = '=C10'
    $dash.Range('H10').Value2  = 'Billing / sales'
    $dash.Range('E11').Value2  = 'Taken -- undisputed'
    $dash.Range('F11').Formula = '=C11*$C$15/$C$14'
    $dash.Range('G11').Formula = '=C11'
    $dash.Range('H11').Value2  = 'Collections'
    $dash.Range('E12').Value2  = 'TOTAL -- classic DSO'
    $dash.Range('F12').Formula = '=C8*$C$15/$C$14'
    $dash.Range('G12').Formula = '=C8'
    $dash.Range('H12').Value2  = ''
    $dash.Range('E12:H12').Font.Bold = $true
    $dash.Range('E13').Value2  = 'Check: parts less total'
    $dash.Range('F13').Formula = '=ROUND(F9+F10+F11-F12,10)'
    $dash.Range('H13').Formula = '=IF(ABS(F13)<0.0000001,"Identity holds","BRIDGE DOES NOT CLOSE")'
    $dash.Range('E14').Value2  = 'Weighted average terms sold'
    $dash.Range('F14').Formula = '=SUMPRODUCT((tSales[MonthEndDate]>EOMONTH(AsOfDate,-12))*(tSales[MonthEndDate]<=AsOfDate)*tSales[CreditSales]*tSales[WeightedAvgTermsDays])/$C$14'
    $dash.Range('H14').Value2  = 'The floor: DSO if every customer paid exactly on the due date'
    $dash.Range('E15').Value2  = 'DSO target'
    $dash.Range('F15').Formula = '=INDEX(tTargets[TargetValue],MATCH("DSO",tTargets[MetricName],0))'
    $dash.Range('E16').Value2  = 'Allowance the target leaves over terms'
    $dash.Range('F16').Formula = '=F15-F14'
    $dash.Range('H16').Formula = '=IF(F16<0,"TARGET IS BELOW THE TERMS FLOOR -- unachievable","Achievable: "&TEXT(F16,"0.0")&" days of collection allowance")'
    $dash.Range('H16').Font.Bold = $true
    # Billing lag sits before the invoice exists, so it EXTENDS the measure
    # rather than decomposing it. Adding it into the bridge above would break
    # the identity on row 13.
    $dash.Range('E17').Value2  = 'Billing lag (despatch to invoice)'
    $dash.Range('F17').Formula = '=SUMPRODUCT((tLag[MonthEndDate]>EOMONTH(AsOfDate,-12))*(tLag[MonthEndDate]<=AsOfDate)*tLag[CreditSales]*tLag[BillingLagDays])/SUMPRODUCT((tLag[MonthEndDate]>EOMONTH(AsOfDate,-12))*(tLag[MonthEndDate]<=AsOfDate)*tLag[CreditSales])'
    $dash.Range('H17').Value2  = 'Before the DSO clock starts, so additive -- and the only days fixable without a customer conversation'
    $dash.Range('E18').Value2  = 'TRUE CASH CYCLE (despatch to cash)'
    $dash.Range('F18').Formula = '=F17+F12'
    $dash.Range('E18:H18').Font.Bold = $true
    $dash.Range('H18').Formula = '="Larger than the DSO anyone is reporting, by "&TEXT(F17,"0.00")&" days"'
    $dash.Range('F9:F18').NumberFormat = '0.00'
    $dash.Range('G9:G12').NumberFormat = '#,##0.00'
    $dash.Range('F13').NumberFormat = '0.000000000'

    # ---- KPI scorecard with formula-driven RAG -----------------------------
    $dash.Range('A20').Value2 = 'SCORECARD -- status is computed against the target table, not typed in'
    $dash.Range('A20').Font.Bold = $true
    $kh = @('Metric','Value','Target','Warning','Status','Variance to target','What it means')
    for ($i = 0; $i -lt $kh.Count; $i++) { $dash.Cells.Item(21, $i + 1).Value2 = $kh[$i] }
    $dash.Range('A21:G21').Font.Bold = $true
    $dash.Range('A21:G21').Interior.Color = $Navy
    $dash.Range('A21:G21').Font.Color = $White

    # metric key, label, Excel-computed value formula, commentary
    $kpis = @(
        @('DSO',               'DSO (countback)',                '=xlDSOCountback', 'Days of recent sales sitting in receivables.'),
        @('AvgDaysDelinquent', 'Average days delinquent',        '=xlADD',          'Days lost purely to late payment, terms stripped out.'),
        @('PctPastDue',        'Past due share of AR',           '=100*($C$10+$C$11)/$C$8', 'By value. Compare with the count version before concluding anything.'),
        @('Pct90Plus',         '90+ share of AR',                '=100*$C$12/$C$8', 'Seriously delinquent balance as a share of the book.')
    )
    $kr = 22
    foreach ($k in $kpis) {
        $metric = $k[0]
        $dash.Cells.Item($kr,1).Value2  = $k[1]
        $dash.Cells.Item($kr,2).Formula = $k[2]
        $dash.Cells.Item($kr,3).Formula = "=INDEX(tTargets[TargetValue],MATCH(""$metric"",tTargets[MetricName],0))"
        $dash.Cells.Item($kr,4).Formula = "=INDEX(tTargets[WarningValue],MATCH(""$metric"",tTargets[MetricName],0))"
        $dash.Cells.Item($kr,5).Formula = "=IF(INDEX(tTargets[Direction],MATCH(""$metric"",tTargets[MetricName],0))=""LowerBetter""," +
                                          "IF(B$kr<=C$kr,""Green"",IF(B$kr<=D$kr,""Amber"",""Red""))," +
                                          "IF(B$kr>=C$kr,""Green"",IF(B$kr>=D$kr,""Amber"",""Red"")))"
        $dash.Cells.Item($kr,6).Formula = "=IF(INDEX(tTargets[Direction],MATCH(""$metric"",tTargets[MetricName],0))=""LowerBetter"",B$kr-C$kr,C$kr-B$kr)"
        $dash.Cells.Item($kr,7).Value2   = $k[3]
        $kr++
    }
    # CEI comes straight from the monthly extract: it needs a prior-month balance
    # that the open-item file does not carry, so Excel cannot recompute it
    # independently and says so rather than pretending.
    $dash.Cells.Item($kr,1).Value2  = 'Collection effectiveness (book)'
    $dash.Cells.Item($kr,2).Formula = "=INDEX(tKPI[CEI_Book],MATCH(MAX(tKPI[AsOfDate]),tKPI[AsOfDate],0))"
    $dash.Cells.Item($kr,3).Formula = '=INDEX(tTargets[TargetValue],MATCH("CEI",tTargets[MetricName],0))'
    $dash.Cells.Item($kr,4).Formula = '=INDEX(tTargets[WarningValue],MATCH("CEI",tTargets[MetricName],0))'
    $dash.Cells.Item($kr,5).Formula = "=IF(B$kr>=C$kr,""Green"",IF(B$kr>=D$kr,""Amber"",""Red""))"
    $dash.Cells.Item($kr,6).Formula = "=C$kr-B$kr"
    $dash.Cells.Item($kr,7).Formula = "=""Cash variant reads ""&TEXT(INDEX(tKPI[CEI_Cash],MATCH(MAX(tKPI[AsOfDate]),tKPI[AsOfDate],0)),""0.00"")&""%. The gap is write-offs counted as collections.""&"" (from SQL: needs a prior-month balance the open-item file does not carry.)"""
    $kr++

    # ---- the three metrics the schema change made possible -----------------
    # All three are recomputed here from their own extracts rather than read as
    # headline figures, so the Validation sheet can hold them against SQL.
    $dash.Cells.Item($kr,1).Value2  = 'Unapplied cash % of AR'
    $dash.Cells.Item($kr,2).Formula = '=100*SUM(tCashApp[UnappliedCash])/$C$8'
    $dash.Cells.Item($kr,3).Formula = '=INDEX(tTargets[TargetValue],MATCH("UnappliedCashPct",tTargets[MetricName],0))'
    $dash.Cells.Item($kr,4).Formula = '=INDEX(tTargets[WarningValue],MATCH("UnappliedCashPct",tTargets[MetricName],0))'
    $dash.Cells.Item($kr,5).Formula = "=IF(B$kr<=C$kr,""Green"",IF(B$kr<=D$kr,""Amber"",""Red""))"
    $dash.Cells.Item($kr,6).Formula = "=B$kr-C$kr"
    $dash.Cells.Item($kr,7).Formula = "=""Cash banked against ""&COUNTA(tCashApp[CustomerID])&"" accounts that nobody has matched. Every dollar of it makes an invoice look unpaid when it is not."""
    $kr++

    $dash.Cells.Item($kr,1).Value2  = 'Billing lag (days)'
    $dash.Cells.Item($kr,2).Formula = '=SUMPRODUCT((tLag[MonthEndDate]>EOMONTH(AsOfDate,-12))*(tLag[MonthEndDate]<=AsOfDate)*tLag[CreditSales]*tLag[BillingLagDays])/SUMPRODUCT((tLag[MonthEndDate]>EOMONTH(AsOfDate,-12))*(tLag[MonthEndDate]<=AsOfDate)*tLag[CreditSales])'
    $dash.Cells.Item($kr,3).Formula = '=INDEX(tTargets[TargetValue],MATCH("BillingLagDays",tTargets[MetricName],0))'
    $dash.Cells.Item($kr,4).Formula = '=INDEX(tTargets[WarningValue],MATCH("BillingLagDays",tTargets[MetricName],0))'
    $dash.Cells.Item($kr,5).Formula = "=IF(B$kr<=C$kr,""Green"",IF(B$kr<=D$kr,""Amber"",""Red""))"
    $dash.Cells.Item($kr,6).Formula = "=B$kr-C$kr"
    $dash.Cells.Item($kr,7).Value2  = 'Despatch to invoice, dollar-weighted. Sits BEFORE the DSO clock starts, so it adds to the cycle rather than being part of DSO.'
    $kr++

    $dash.Cells.Item($kr,1).Value2  = 'Promise kept rate %'
    $dash.Cells.Item($kr,2).Formula = '=100*SUMIFS(tPromise[PromisedAmount],tPromise[PromiseStatus],"Kept",tPromise[PromisedPayDate],">"&EOMONTH(AsOfDate,-12))/SUMIFS(tPromise[PromisedAmount],tPromise[PromiseStatus],"<>Outstanding",tPromise[PromisedPayDate],">"&EOMONTH(AsOfDate,-12))'
    $dash.Cells.Item($kr,3).Formula = '=INDEX(tTargets[TargetValue],MATCH("PromiseKeptRate",tTargets[MetricName],0))'
    $dash.Cells.Item($kr,4).Formula = '=INDEX(tTargets[WarningValue],MATCH("PromiseKeptRate",tTargets[MetricName],0))'
    $dash.Cells.Item($kr,5).Formula = "=IF(B$kr>=C$kr,""Green"",IF(B$kr>=D$kr,""Amber"",""Red""))"
    $dash.Cells.Item($kr,6).Formula = "=C$kr-B$kr"
    $dash.Cells.Item($kr,7).Formula = "=""Portfolio figure. The High-risk tier reads ""&TEXT(100*SUMIFS(tPromise[PromisedAmount],tPromise[PromiseStatus],""Kept"",tPromise[RiskTier],""High"")/SUMIFS(tPromise[PromisedAmount],tPromise[PromiseStatus],""<>Outstanding"",tPromise[RiskTier],""High""),""0.0"")&""% -- read the cohort, never the headline."""

    $kpiLast = $kr
    $dash.Range("B22:D$kpiLast").NumberFormat = '0.00'
    $dash.Range("F22:F$kpiLast").NumberFormat = '+0.00;-0.00'
    $dash.Range("E22:E$kpiLast").HorizontalAlignment = -4108

    # RAG conditional formatting, driven by the computed text
    foreach ($rule in @(@('Green',$GreenBg,$GreenFg), @('Amber',$AmberBg,$AmberFg), @('Red',$RedBg,$RedFg))) {
        $fc = $dash.Range("E22:E$kpiLast").FormatConditions.Add(1, 3, """$($rule[0])""")   # xlCellValue, xlEqual
        $fc.Interior.Color = $rule[1]
        $fc.Font.Color = $rule[2]
        $fc.Font.Bold = $true
    }

    # ---- ageing profile, bucketed by formula from days past due ------------
    $ageTop = $kpiLast + 2
    $dash.Cells.Item($ageTop,1).Value2 = 'AGEING -- buckets re-derived in Excel from days past due, not taken from the extract'
    $dash.Cells.Item($ageTop,1).Font.Bold = $true
    $ah = @('Bucket','Invoices','Balance','Share of AR','Risk-weighted exposure')
    for ($i = 0; $i -lt $ah.Count; $i++) { $dash.Cells.Item($ageTop + 1, $i + 1).Value2 = $ah[$i] }
    $dash.Range($dash.Cells.Item($ageTop+1,1), $dash.Cells.Item($ageTop+1,5)).Font.Bold = $true
    $dash.Range($dash.Cells.Item($ageTop+1,1), $dash.Cells.Item($ageTop+1,5)).Interior.Color = $Grey
    for ($i = 0; $i -lt 5; $i++) {
        $rw = $ageTop + 2 + $i
        $bi = $i + 1
        $dash.Cells.Item($rw,1).Formula = "=INDEX(tBuckets[BucketName],$bi)"
        $dash.Cells.Item($rw,2).Formula = "=COUNTIFS(tOpenItems[DaysPastDue],"">=""&INDEX(tBuckets[MinDaysPastDue],$bi),tOpenItems[DaysPastDue],""<=""&INDEX(tBuckets[MaxDaysPastDue],$bi))"
        $dash.Cells.Item($rw,3).Formula = "=SUMIFS(tOpenItems[OpenBalance],tOpenItems[DaysPastDue],"">=""&INDEX(tBuckets[MinDaysPastDue],$bi),tOpenItems[DaysPastDue],""<=""&INDEX(tBuckets[MaxDaysPastDue],$bi))"
        $dash.Cells.Item($rw,4).Formula = "=C$rw/`$C`$8"
        $dash.Cells.Item($rw,5).Formula = "=C$rw*INDEX(tBuckets[RiskWeight],$bi)"
    }
    $ageTotal = $ageTop + 7
    $dash.Cells.Item($ageTotal,1).Value2  = 'TOTAL'
    $dash.Cells.Item($ageTotal,2).Formula = "=SUM(B$($ageTop+2):B$($ageTop+6))"
    $dash.Cells.Item($ageTotal,3).Formula = "=SUM(C$($ageTop+2):C$($ageTop+6))"
    $dash.Cells.Item($ageTotal,4).Formula = "=SUM(D$($ageTop+2):D$($ageTop+6))"
    $dash.Cells.Item($ageTotal,5).Formula = "=SUM(E$($ageTop+2):E$($ageTop+6))"
    $dash.Range($dash.Cells.Item($ageTotal,1), $dash.Cells.Item($ageTotal,5)).Font.Bold = $true
    $dash.Range($dash.Cells.Item($ageTop+2,2), $dash.Cells.Item($ageTotal,2)).NumberFormat = '#,##0'
    $dash.Range($dash.Cells.Item($ageTop+2,3), $dash.Cells.Item($ageTotal,3)).NumberFormat = '#,##0.00'
    $dash.Range($dash.Cells.Item($ageTop+2,4), $dash.Cells.Item($ageTotal,4)).NumberFormat = '0.0%'
    $dash.Range($dash.Cells.Item($ageTop+2,5), $dash.Cells.Item($ageTotal,5)).NumberFormat = '#,##0.00'

    $dash.Columns.Item(1).ColumnWidth = 32
    $dash.Columns.Item(2).ColumnWidth = 13
    $dash.Columns.Item(3).ColumnWidth = 15
    $dash.Columns.Item(4).ColumnWidth = 13
    $dash.Columns.Item(5).ColumnWidth = 30
    $dash.Columns.Item(6).ColumnWidth = 12
    $dash.Columns.Item(7).ColumnWidth = 16
    $dash.Columns.Item(8).ColumnWidth = 44
    Write-Host "  Dashboard                built"

    # =========================================================================
    # 5. PRIORITY QUEUE -- the operational control
    # =========================================================================

    $pq.Range('A1').Value2 = 'PRIORITY ACTION QUEUE'
    $pq.Range('A1').Font.Size = 16
    $pq.Range('A1').Font.Bold = $true
    $pq.Range('A1').Font.Color = $Navy
    $pq.Range('A2').Value2 = 'Who collections should call first, ranked by risk-weighted money a call can actually move. Change the filters; the list re-ranks itself.'
    $pq.Range('A2').Font.Italic = $true

    $pq.Range('B4').Value2 = 'Collector'                 ; $pq.Range('C4').Value2 = '(All)'
    $pq.Range('B5').Value2 = 'Action'                    ; $pq.Range('C5').Value2 = '(All)'
    $pq.Range('B6').Value2 = "Today's worklist only"     ; $pq.Range('C6').Value2 = 'Yes'
    $pq.Range('B7').Value2 = 'Minimum past due ($)'      ; $pq.Range('C7').Formula = '0'
    $pq.Range('B4:B7').Font.Bold = $true
    $pq.Range('C4:C7').Interior.Color = $NavyLight
    $pq.Range('C7').NumberFormat = '#,##0'

    # dropdowns sourced from the data, so they cannot drift out of date
    $collectors = @('(All)') + (Import-Csv (Join-Path $ExtractPath 'priority_action_queue.csv') | Select-Object -ExpandProperty CollectorName -Unique | Sort-Object)
    $actions    = @('(All)') + (Import-Csv (Join-Path $ExtractPath 'priority_action_queue.csv') | Select-Object -ExpandProperty ActionCode -Unique | Sort-Object)
    $pq.Range('C4').Validation.Add(3, 1, 1, ($collectors -join ',')) | Out-Null
    $pq.Range('C5').Validation.Add(3, 1, 1, ($actions -join ',')) | Out-Null
    $pq.Range('C6').Validation.Add(3, 1, 1, 'Yes,No') | Out-Null

    $pq.Range('E4').Value2  = 'Accounts matching'
    $pq.Range('F4').Formula = "=SUM(Calc!`$B`$$qFirst`:`$B`$$qLast)"
    $pq.Range('E5').Value2  = 'Past due in scope'
    $pq.Range('F5').Formula = "=SUMPRODUCT(Calc!`$B`$$qFirst`:`$B`$$qLast,tQueue[PastDueBalance])"
    $pq.Range('E6').Value2  = 'Collectable exposure in scope'
    $pq.Range('F6').Formula = "=SUMPRODUCT(Calc!`$B`$$qFirst`:`$B`$$qLast,tQueue[CollectableExposure])"
    $pq.Range('E7').Formula = '="Showing top "&MIN(QueueRowsShown,F4)&" of "&F4'
    $pq.Range('E4:E6').Font.Bold = $true
    $pq.Range('F5:F6').NumberFormat = '#,##0.00'
    $pq.Range('E7').Font.Bold = $true

    $qh = @('#','Customer','Segment','Risk','Collector','Terms','Open balance','Past due','90+','Disputed',
            'Collectable exposure','Oldest days','Credit use %','Hold','Action','What to do')
    for ($i = 0; $i -lt $qh.Count; $i++) { $pq.Cells.Item(9, $i + 1).Value2 = $qh[$i] }
    $pq.Range('A9:P9').Font.Bold = $true
    $pq.Range('A9:P9').Interior.Color = $Navy
    $pq.Range('A9:P9').Font.Color = $White
    $pq.Range('A9:P9').WrapText = $true

    $qCols = @('CustomerName','Segment','RiskTier','CollectorName','TermsCode','OpenBalance','PastDueBalance',
               'Balance90Plus','DisputedBalance','CollectableExposure','OldestDaysPastDue','CreditUtilizationPct',
               'CreditHoldFlag','ActionCode','RecommendedAction')
    $qRows = 25
    for ($i = 0; $i -lt $qRows; $i++) {
        $rw = 10 + $i
        $n  = $i + 1
        $pq.Cells.Item($rw,1).Formula = "=IF($n>MIN(QueueRowsShown,`$F`$4),"""",$n)"
        for ($c = 0; $c -lt $qCols.Count; $c++) {
            $col = $qCols[$c]
            $pq.Cells.Item($rw, $c + 2).Formula =
                "=IFERROR(INDEX(tQueue[$col],MATCH($n,Calc!`$C`$$qFirst`:`$C`$$qLast,0)),"""")"
        }
    }
    $pq.Range("G10:K$(9+$qRows)").NumberFormat = '#,##0.00'
    $pq.Range("M10:M$(9+$qRows)").NumberFormat = '0.0'
    $pq.Range("L10:L$(9+$qRows)").NumberFormat = '#,##0'

    # highlight the rows that need escalating, computed not coloured by hand
    $fcEsc = $pq.Range("A10:P$(9+$qRows)").FormatConditions.Add(2, 0, '=$O10="ESCALATE"')
    $fcEsc.Interior.Color = $RedBg
    $fcEsc.Font.Color = $RedFg
    $fcHold = $pq.Range("A10:P$(9+$qRows)").FormatConditions.Add(2, 0, '=AND($N10=1,$O10<>"ESCALATE")')
    $fcHold.Interior.Color = $AmberBg
    $fcHold.Font.Color = $AmberFg

    $widths = @(5,30,16,9,18,11,14,13,12,12,16,12,12,7,18,66)
    for ($i = 0; $i -lt $widths.Count; $i++) { $pq.Columns.Item($i + 1).ColumnWidth = $widths[$i] }
    Write-Host "  Priority Queue           built ($qRows rows shown)"

    # =========================================================================
    # 5b. CASH APPLICATION -- the other worklist
    #
    # Collections gets a call list; this is what cash application gets. Both are
    # ranked by money, but this one is ranked by money Vantage ALREADY HAS and
    # cannot see, which is cheaper to recover than anything a phone call can
    # achieve: no negotiation, no concession, no relationship cost.
    #
    # Ordered oldest-first rather than largest-first, because a receipt that has
    # sat unmatched longest is the one most likely to have already caused a
    # wrong call.
    # =========================================================================
    $ca.Range('A1').Value2 = 'CASH APPLICATION WORKLIST'
    $ca.Range('A1').Font.Size = 16
    $ca.Range('A1').Font.Bold = $true
    $ca.Range('A1').Font.Color = $Navy
    $ca.Range('A2').Value2 = 'Money already in the bank that nobody has matched to an invoice. Until it is applied, these invoices read as unpaid and the call list will chase customers who have already paid.'
    $ca.Range('A2:I2').Merge()
    $ca.Range('A2').WrapText = $true
    $ca.Range('A2').RowHeight = 30

    $ca.Range('B4').Value2  = 'Accounts holding unapplied cash'
    $ca.Range('F4').Formula = '=COUNTA(tCashApp[CustomerID])'
    $ca.Range('B5').Value2  = 'Cash banked and unmatched'
    $ca.Range('F5').Formula = '=SUM(tCashApp[UnappliedCash])'
    $ca.Range('B6').Value2  = 'Of which fully unapplied receipts'
    $ca.Range('F6').Formula = '=SUM(tCashApp[FullyUnappliedReceipts])'
    $ca.Range('B7').Value2  = 'Past due that would be WRONGLY chased'
    $ca.Range('F7').Formula = '=SUM(tCashApp[WronglyChaseable])'
    $ca.Range('B8').Value2  = 'Oldest unmatched receipt (days)'
    $ca.Range('F8').Formula = '=MAX(tCashApp[OldestUnappliedDays])'
    $ca.Range('B4:B8').Font.Bold = $true
    $ca.Range('F5:F5').NumberFormat = '#,##0.00'
    $ca.Range('F7:F7').NumberFormat = '#,##0.00'
    $ca.Range('F4').NumberFormat = '#,##0'
    $ca.Range('F6').NumberFormat = '#,##0'
    $ca.Range('F8').NumberFormat = '#,##0'
    $ca.Range('H7').Formula = '="This is cash Vantage already holds -- recovering it needs no customer conversation at all."'
    $ca.Range('H7').Font.Bold = $true

    $cah = @('#','Customer','Segment','Region','Collector','Unapplied cash','Receipts',
             'Fully unapplied','No remittance','Oldest days','Open balance','Past due',
             'Net exposure','Wrongly chaseable','Root cause')
    for ($i = 0; $i -lt $cah.Count; $i++) { $ca.Cells.Item(10, $i + 1).Value2 = $cah[$i] }
    $ca.Range('A10:O10').Font.Bold = $true
    $ca.Range('A10:O10').Interior.Color = $Navy
    $ca.Range('A10:O10').Font.Color = $White
    $ca.Range('A10:O10').WrapText = $true

    # The extract is already ordered oldest-first, so the sheet indexes straight
    # into it rather than re-ranking.
    $caCols = @('CustomerName','Segment','Region','CollectorName','UnappliedCash','UnappliedReceipts',
                'FullyUnappliedReceipts','NoRemittanceAdvice','OldestUnappliedDays','OpenBalance',
                'PastDueBalance','NetExposure','WronglyChaseable')
    $caRows = 30
    for ($i = 0; $i -lt $caRows; $i++) {
        $rw = 11 + $i
        $n  = $i + 1
        $ca.Cells.Item($rw,1).Formula = "=IF($n>COUNTA(tCashApp[CustomerID]),"""",$n)"
        for ($c = 0; $c -lt $caCols.Count; $c++) {
            $col = $caCols[$c]
            $ca.Cells.Item($rw, $c + 2).Formula = "=IFERROR(INDEX(tCashApp[$col],$n),"""")"
        }
        $ca.Cells.Item($rw, 15).Formula =
            "=IFERROR(IF(INDEX(tCashApp[NoRemittanceAdvice],$n)>0," +
            """No remittance advice -- request it, then apply""," +
            """Advice received; the cash simply has not been matched""),"""")"
    }
    $ca.Range("F11:F$(10+$caRows)").NumberFormat = '#,##0.00'
    $ca.Range("L11:O$(10+$caRows)").NumberFormat = '#,##0.00'

    # anything unmatched for more than 90 days is a defect, not a backlog
    $fcOld = $ca.Range("A11:O$(10+$caRows)").FormatConditions.Add(2, 0, '=$J11>90')
    $fcOld.Interior.Color = $RedBg
    $fcOld.Font.Color = $RedFg

    $caw = @(5,30,16,12,18,15,10,13,13,12,14,14,14,16,46)
    for ($i = 0; $i -lt $caw.Count; $i++) { $ca.Columns.Item($i + 1).ColumnWidth = $caw[$i] }
    Write-Host "  Cash Application         built ($caRows rows shown)"

    # =========================================================================
    # 6. AGEING MATRIX -- one cut, any dimension, driven by INDEX/MATCH
    # =========================================================================
    $am.Range('A1').Value2 = 'AGEING MATRIX'
    $am.Range('A1').Font.Size = 16
    $am.Range('A1').Font.Bold = $true
    $am.Range('A1').Font.Color = $Navy
    $am.Range('A2').Value2 = 'Pick a dimension. SUMIFS resolves the criteria column by name, so one set of formulas serves every cut.'
    $am.Range('A2').Font.Italic = $true

    $dims = @('Segment','Region','RiskTier','TermsCode','CollectorName','Team')
    $am.Range('B4').Value2 = 'Cut by'
    $am.Range('B4').Font.Bold = $true
    $am.Range('C4').Value2 = 'Segment'
    $am.Range('C4').Interior.Color = $NavyLight
    $am.Range('C4').Validation.Add(3, 1, 1, ($dims -join ',')) | Out-Null

    # distinct values per dimension, resolved at build time into a lookup block
    $openItems = Import-Csv (Join-Path $ExtractPath 'ar_open_items.csv')
    $distinct = @{}
    $maxVals = 0
    foreach ($d in $dims) {
        $vals = $openItems | Select-Object -ExpandProperty $d -Unique | Sort-Object
        $distinct[$d] = $vals
        if ($vals.Count -gt $maxVals) { $maxVals = $vals.Count }
    }
    # Park the lists on Calc, well clear of the other blocks. Column N onwards:
    # the countback block owns A-I rows 6-17 and the queue ranking owns A-C from
    # row 28, and an overlap here is not obvious when it happens -- it overwrote
    # the countback's "covers the balance?" flags with segment names, and MATCH
    # then skipped the text rows and returned a plausible wrong answer instead
    # of an error.
    $listCol = 14
    $listTop = 4
    $calc.Cells.Item(3, $listCol).Value2 = 'DIMENSION VALUE LISTS (used by the Ageing Matrix)'
    $calc.Cells.Item(3, $listCol).Font.Bold = $true
    for ($d = 0; $d -lt $dims.Count; $d++) {
        $calc.Cells.Item($listTop, $listCol + $d).Value2 = $dims[$d]
        $calc.Cells.Item($listTop, $listCol + $d).Font.Bold = $true
        $vals = $distinct[$dims[$d]]
        for ($v = 0; $v -lt $vals.Count; $v++) {
            $calc.Cells.Item($listTop + 1 + $v, $listCol + $d).Value2 = $vals[$v]
        }
    }
    $listFirstCol = $listCol
    $listLastCol  = $listCol + $dims.Count - 1
    $listFirstRow = $listTop + 1
    $listLastRow  = $listTop + $maxVals

    $am.Range('A6').Value2 = 'Value'
    for ($i = 0; $i -lt 5; $i++) { $am.Cells.Item(6, $i + 2).Formula = "=INDEX(tBuckets[BucketName],$($i+1))" }
    $am.Range('G6').Value2 = 'Total'
    $am.Range('H6').Value2 = 'Past due'
    $am.Range('I6').Value2 = 'Share past due'
    $am.Range('A6:I6').Font.Bold = $true
    $am.Range('A6:I6').Interior.Color = $Navy
    $am.Range('A6:I6').Font.Color = $White

    $critCol = "INDEX(tOpenItems,0,MATCH(`$C`$4,tOpenItems[#Headers],0))"
    for ($i = 0; $i -lt $maxVals; $i++) {
        $rw = 7 + $i
        $n  = $i + 1
        $am.Cells.Item($rw,1).Formula =
            "=IFERROR(INDEX(Calc!`$$([char](64+$listFirstCol))`$$listFirstRow`:`$$([char](64+$listLastCol))`$$listLastRow,$n,MATCH(`$C`$4,Calc!`$$([char](64+$listFirstCol))`$$listTop`:`$$([char](64+$listLastCol))`$$listTop,0)),"""")"
        for ($b = 0; $b -lt 5; $b++) {
            $bi = $b + 1
            $am.Cells.Item($rw, $b + 2).Formula =
                "=IF(`$A$rw="""","""",SUMIFS(tOpenItems[OpenBalance],$critCol,`$A$rw," +
                "tOpenItems[DaysPastDue],"">=""&INDEX(tBuckets[MinDaysPastDue],$bi)," +
                "tOpenItems[DaysPastDue],""<=""&INDEX(tBuckets[MaxDaysPastDue],$bi)))"
        }
        $am.Cells.Item($rw,7).Formula = "=IF(`$A$rw="""","""",SUM(B$rw`:F$rw))"
        $am.Cells.Item($rw,8).Formula = "=IF(`$A$rw="""","""",SUM(C$rw`:F$rw))"
        $am.Cells.Item($rw,9).Formula = "=IF(OR(`$A$rw="""",G$rw=0),"""",H$rw/G$rw)"
    }
    $amTotal = 7 + $maxVals
    $am.Cells.Item($amTotal,1).Value2 = 'TOTAL'
    for ($c = 2; $c -le 8; $c++) {
        $L = [char](64 + $c)
        $am.Cells.Item($amTotal,$c).Formula = "=SUM($L`7:$L$($amTotal-1))"
    }
    $am.Cells.Item($amTotal,9).Formula = "=IF(G$amTotal=0,"""",H$amTotal/G$amTotal)"
    $am.Range($am.Cells.Item($amTotal,1), $am.Cells.Item($amTotal,9)).Font.Bold = $true
    $am.Range("B7:H$amTotal").NumberFormat = '#,##0.00'
    $am.Range("I7:I$amTotal").NumberFormat = '0.0%'
    $am.Columns.Item(1).ColumnWidth = 30
    for ($c = 2; $c -le 9; $c++) { $am.Columns.Item($c).ColumnWidth = 15 }
    Write-Host "  Ageing Matrix            built ($maxVals rows)"

    # =========================================================================
    # 7. VALIDATION -- Excel against SQL, to the cent
    # =========================================================================
    $val.Range('A1').Value2 = 'CROSS-IMPLEMENTATION VALIDATION'
    $val.Range('A1').Font.Size = 16
    $val.Range('A1').Font.Bold = $true
    $val.Range('A1').Font.Color = $Navy
    $val.Range('A2').Value2 = 'Each figure below was computed twice: once in T-SQL against the database, once in Excel formulas against the extract. A dashboard that only displays what the database told it cannot detect that the database was wrong.'
    $val.Range('A2:F2').Merge()
    $val.Range('A2').WrapText = $true
    $val.Range('A2').RowHeight = 30

    $vh = @('#','Figure','SQL','Excel','Difference','Result')
    for ($i = 0; $i -lt $vh.Count; $i++) { $val.Cells.Item(4, $i + 1).Value2 = $vh[$i] }
    $val.Range('A4:F4').Font.Bold = $true
    $val.Range('A4:F4').Interior.Color = $Navy
    $val.Range('A4:F4').Font.Color = $White

    $lastKPI = "MATCH(MAX(tKPI[AsOfDate]),tKPI[AsOfDate],0)"
    $lastBr  = "MATCH(MAX(tBridge[AsOfDate]),tBridge[AsOfDate],0)"
    $checks = @(
        @('Total open AR',              "=INDEX(tKPI[EndAR],$lastKPI)",              '=Dashboard!$C$8',  0.005),
        @('Not yet due',                "=INDEX(tKPI[EndCurrentAR],$lastKPI)",       '=Dashboard!$C$9',  0.005),
        @('Past due AR',                "=INDEX(tKPI[PastDueAR],$lastKPI)",          '=Dashboard!$C$10+Dashboard!$C$11', 0.005),
        @('90+ AR',                     "=INDEX(tKPI[AR90Plus],$lastKPI)",           '=Dashboard!$C$12', 0.005),
        @('Open invoice count',         "=INDEX(tKPI[OpenInvoices],$lastKPI)",       '=Dashboard!$C$13', 0.5),
        # Compared per day, not per year. SQL stores SalesPerDay as DECIMAL(14,2),
        # so multiplying it back up by 365 magnifies that half-cent rounding into
        # a dollar and the check fails on its own arithmetic rather than on a
        # real disagreement.
        @('Credit sales per day',       "=INDEX(tBridge[SalesPerDay],$lastBr)",      '=Dashboard!$C$16', 0.005),
        @('Classic DSO',                "=INDEX(tBridge[DSO_Classic],$lastBr)",      '=Dashboard!$F$12', 0.005),
        @('Countback DSO',              "=INDEX(tKPI[DSO_Countback],$lastKPI)",      '=xlDSOCountback',  0.005),
        @('Best possible DSO',          "=INDEX(tKPI[BPDSO_Countback],$lastKPI)",    '=xlBPDSO',         0.005),
        @('Average days delinquent',    "=INDEX(tKPI[AvgDaysDelinquent],$lastKPI)",  '=xlADD',           0.005),
        @('Granted days',               "=INDEX(tBridge[GrantedDays],$lastBr)",      '=Dashboard!$F$9',  0.005),
        @('Dispute days',               "=INDEX(tBridge[DisputeDays],$lastBr)",      '=Dashboard!$F$10', 0.005),
        @('Lateness days',              "=INDEX(tBridge[LatenessDays],$lastBr)",     '=Dashboard!$F$11', 0.005),
        @('Weighted average terms',     "=INDEX(tBridge[WeightedAvgTermsDays],$lastBr)", '=Dashboard!$F$14', 0.005),
        @('Past due % of AR',           "=INDEX(tKPI[PctPastDue],$lastKPI)",         '=Dashboard!$B$24', 0.005),
        @('90+ % of AR',                "=INDEX(tKPI[Pct90Plus],$lastKPI)",          '=Dashboard!$B$25', 0.005),
        # the three metrics the schema change made possible, each recomputed in
        # Excel from its own extract and held against the SQL figure
        @('Unapplied cash',             "=INDEX(tKPI[UnappliedCash],$lastKPI)",      '=SUM(tCashApp[UnappliedCash])', 0.005),
        @('Unapplied cash % of AR',     "=INDEX(tKPI[UnappliedCashPct],$lastKPI)",   '=Dashboard!$B$27', 0.005),
        @('Billing lag (days)',         "=INDEX(tBridge[BillingLagDays],$lastBr)",   '=Dashboard!$B$28', 0.01),
        @('True cash cycle (days)',     "=INDEX(tBridge[CashCycleDays],$lastBr)",    '=Dashboard!$F$18', 0.01),
        @('Net exposure (AR less cash held)', "=INDEX(tKPI[EndAR],$lastKPI)-INDEX(tKPI[UnappliedCash],$lastKPI)",
                                                                                     '=Dashboard!$C$8-SUM(tCashApp[UnappliedCash])', 0.005)
    )
    $vr = 5
    foreach ($c in $checks) {
        $val.Cells.Item($vr,1).Formula  = [string]($vr - 4)
        $val.Cells.Item($vr,2).Value2  = $c[0]
        $val.Cells.Item($vr,3).Formula = $c[1]
        $val.Cells.Item($vr,4).Formula = $c[2]
        $val.Cells.Item($vr,5).Formula = "=C$vr-D$vr"
        $val.Cells.Item($vr,6).Formula = "=IF(ABS(E$vr)<=$($c[3]),""MATCH"",""DIFFERS"")"
        $vr++
    }
    # re-derived ageing buckets, SQL's bucket label against Excel's own arithmetic
    for ($i = 0; $i -lt 5; $i++) {
        $bi = $i + 1
        $val.Cells.Item($vr,1).Formula  = [string]($vr - 4)
        $val.Cells.Item($vr,2).Formula = "=""Bucket: ""&INDEX(tBuckets[BucketName],$bi)"
        $val.Cells.Item($vr,3).Formula = "=SUMIFS(tOpenItems[OpenBalance],tOpenItems[BucketKey],$bi)"
        $val.Cells.Item($vr,4).Formula = "=SUMIFS(tOpenItems[OpenBalance],tOpenItems[DaysPastDue],"">=""&INDEX(tBuckets[MinDaysPastDue],$bi),tOpenItems[DaysPastDue],""<=""&INDEX(tBuckets[MaxDaysPastDue],$bi))"
        $val.Cells.Item($vr,5).Formula = "=C$vr-D$vr"
        $val.Cells.Item($vr,6).Formula = "=IF(ABS(E$vr)<=0.005,""MATCH"",""DIFFERS"")"
        $vr++
    }
    # days past due re-derived from the reporting date
    $val.Cells.Item($vr,1).Formula  = [string]($vr - 4)
    $val.Cells.Item($vr,2).Value2  = 'Days past due (sum, re-derived from dates)'
    $val.Cells.Item($vr,3).Formula = '=SUM(tOpenItems[DaysPastDue])'
    $val.Cells.Item($vr,4).Formula = '=SUMPRODUCT(AsOfDate-tOpenItems[DueDate])'
    $val.Cells.Item($vr,5).Formula = "=C$vr-D$vr"
    $val.Cells.Item($vr,6).Formula = "=IF(ABS(E$vr)<=0.5,""MATCH"",""DIFFERS"")"
    $vLast = $vr

    $val.Range("C5:E$vLast").NumberFormat = '#,##0.0000'
    $val.Range("F5:F$vLast").HorizontalAlignment = -4108
    $fcM = $val.Range("F5:F$vLast").FormatConditions.Add(1, 3, '"MATCH"')
    $fcM.Interior.Color = $GreenBg ; $fcM.Font.Color = $GreenFg ; $fcM.Font.Bold = $true
    $fcD = $val.Range("F5:F$vLast").FormatConditions.Add(1, 3, '"DIFFERS"')
    $fcD.Interior.Color = $RedBg ; $fcD.Font.Color = $RedFg ; $fcD.Font.Bold = $true

    $val.Cells.Item($vLast + 2, 2).Value2  = 'Checks reconciling'
    $val.Cells.Item($vLast + 2, 3).Formula = "=COUNTIF(`$F`$5:`$F`$$vLast,""MATCH"")&"" of ""&COUNTA(`$F`$5:`$F`$$vLast)"
    $val.Cells.Item($vLast + 2, 2).Font.Bold = $true
    $val.Cells.Item($vLast + 2, 3).Font.Bold = $true
    $val.Cells.Item($vLast + 3, 2).Formula = "=IF(COUNTIF(`$F`$5:`$F`$$vLast,""DIFFERS"")=0,""All figures reconcile between the two independent implementations."",""RECONCILIATION FAILED -- do not publish these figures."")"
    $val.Cells.Item($vLast + 3, 2).Font.Bold = $true

    $val.Columns.Item(1).ColumnWidth = 5
    $val.Columns.Item(2).ColumnWidth = 40
    $val.Columns.Item(3).ColumnWidth = 18
    $val.Columns.Item(4).ColumnWidth = 18
    $val.Columns.Item(5).ColumnWidth = 15
    $val.Columns.Item(6).ColumnWidth = 12
    Write-Host "  Validation               built ($($vLast - 4) checks)"

    # =========================================================================
    # 8. README
    # =========================================================================
    $lines = @(
        @('VANTAGE WHOLESALE SUPPLY -- RECEIVABLES CONTROL', 'title'),
        @('', ''),
        @('Vantage Wholesale Supply is a fictional industrial parts distributor. Every figure in this workbook is synthetic, generated by the SQL scripts in the sql\ folder of this project. No confidential data is used and no claim is made about any production system.', 'body'),
        @('', ''),
        @('THE QUESTION THIS ANSWERS', 'head'),
        @('DSO has risen. How many of those days did Vantage grant itself through the terms it sold, how many are customers taking beyond those terms, and who should collections call first?', 'body'),
        @('', ''),
        @('SHEETS', 'head'),
        @('Dashboard         The scorecard, the DSO bridge and the ageing profile. Every number is a formula over the extract.', 'body'),
        @('Priority Queue    The operational control: who to call, in what order, with the instruction and the dollar ask.', 'body'),
        @('Cash Application   The OTHER worklist: money already banked that nobody matched. These accounts must not be called.', 'body'),
        @('Ageing Matrix     Open AR by bucket, cut by any dimension you pick.', 'body'),
        @('Validation        Each headline figure computed twice -- once in SQL, once in Excel -- and reconciled.', 'body'),
        @('Control           Every parameter the workbook obeys.', 'body'),
        @('Calc              The working behind countback DSO and the queue ranking. Visible on purpose.', 'body'),
        @('Data_*            The loaded extracts. Do not edit; they are replaced on refresh.', 'body'),
        @('', ''),
        @('HOW TO REFRESH', 'head'),
        @('1. Rebuild the extracts:  powershell erp_extracts\Export_Extracts.ps1', 'body'),
        @('2. If the repo has moved, update the Extract folder cell on the Control sheet.', 'body'),
        @('3. Data > Refresh All.', 'body'),
        @('4. Check the Validation sheet still reads "All figures reconcile".', 'body'),
        @('', ''),
        @('TO READ FROM SQL SERVER INSTEAD OF THE CSVs', 'head'),
        @('Replace the Source and Promoted steps of any query with a direct connection. For example, for tOpenItems:', 'body'),
        @('let', 'code'),
        @('    Source = Sql.Database("localhost\SQLEXPRESS", "VantageAR", [Query="SELECT * FROM dbo.vw_ARBalanceDetail WHERE IsOpen = 1"])', 'code'),
        @('in', 'code'),
        @('    Source', 'code'),
        @('', ''),
        @('WHY THERE ARE NO DYNAMIC-ARRAY FORMULAS', 'head'),
        @('FILTER, SORT and UNIQUE would shorten several formulas here and would break the workbook on Excel 2019 and LibreOffice. The queue ranking uses SUMPRODUCT and INDEX/MATCH instead, which work everywhere.', 'body'),
        @('', ''),
        @('WHAT MAKES THIS A CONTROL RATHER THAN A REPORT', 'head'),
        @('Nothing on the Dashboard is a pasted value. The balances come from SUMIFS over the open-item extract, the countback DSO from a real month-by-month loop on the Calc sheet, and the red/amber/green status from INDEX/MATCH against the target table -- so changing a target changes the status, and changing the data changes the numbers. The Validation sheet then checks the Excel implementation against the SQL one and says on its own face if they ever disagree.', 'body')
    )
    $rr = 1
    foreach ($l in $lines) {
        $rd.Cells.Item($rr,1).Value2 = $l[0]
        switch ($l[1]) {
            'title' { $rd.Cells.Item($rr,1).Font.Size = 16; $rd.Cells.Item($rr,1).Font.Bold = $true; $rd.Cells.Item($rr,1).Font.Color = $Navy }
            'head'  { $rd.Cells.Item($rr,1).Font.Bold = $true; $rd.Cells.Item($rr,1).Font.Color = $Navy }
            'code'  { $rd.Cells.Item($rr,1).Font.Name = 'Consolas'; $rd.Cells.Item($rr,1).Interior.Color = $Grey }
        }
        $rr++
    }
    $rd.Columns.Item(1).ColumnWidth = 130
    $rd.Range('A1:A60').WrapText = $false

    # =========================================================================
    # 9. finish
    # =========================================================================
    # -------------------------------------------------------------------------
    # Make the extract folder self-locating before saving.
    #
    # The queries needed a real path to load against during the build, but
    # shipping that path would bake one machine's directory -- and its username
    # -- into a public file, and would break the moment the repo is cloned
    # anywhere else. CELL("filename") returns the workbook's own location, so
    # the folder resolves relative to wherever the file actually sits.
    # -------------------------------------------------------------------------
    # CELL("filename") returns a https://d.docs.live.net/... URL when the file
    # lives in a synced OneDrive folder, and an empty string before the workbook
    # has ever been saved. Neither is a path Power Query can open, so the formula
    # detects both and asks for the folder rather than resolving to something
    # broken and blaming the query.
    $ctl.Range('B6').Formula =
        '=IF(OR(CELL("filename")="",LEFT(CELL("filename"),4)="http"),' +
        '"Set this cell to the full path of the erp_extracts folder, then Data > Refresh All",' +
        'LEFT(CELL("filename"),FIND("[",CELL("filename"))-1)&"..\erp_extracts")'
    $ctl.Range('C6').Value2  = 'Resolves from this workbook''s own location so the repo can be cloned anywhere. On OneDrive, Excel reports a URL rather than a path, so set the folder here by hand. The workbook ships with its data already loaded, so this only matters when you refresh.'

    # Working sheets first: the tab a reader lands on should be one that answers
    # the question, not the raw extract. Moving each to position 1 in reverse
    # order leaves them in the intended order with the Data_ sheets behind them.
    $finalOrder = @('README','Dashboard','Priority Queue','Cash Application','Ageing Matrix','Validation','Calc','Control')
    for ($i = $finalOrder.Count - 1; $i -ge 0; $i--) {
        $wb.Worksheets.Item($finalOrder[$i]).Move($wb.Worksheets.Item(1))
    }
    $wb.Worksheets.Item('README').Activate()
    $xl.Calculation = -4105          # xlCalculationAutomatic
    $wb.Application.CalculateFullRebuild()

    # A portfolio file should not carry a personal name. Set neutral properties
    # first, then ask Excel to strip personal information on save -- the reverse
    # order writes a name and immediately discards it. These are cosmetic, so a
    # failure here must not cost us the workbook.
    foreach ($prop in @(
        @('Author',  'Vantage AR Analytics'),
        @('Company', 'Vantage Wholesale Supply (fictional)'),
        @('Title',   'Vantage Wholesale Supply -- Receivables Control'),
        @('Comments','Synthetic demonstration data. No confidential data and no production system.')
    )) {
        try { $wb.BuiltinDocumentProperties($prop[0]).Value = $prop[1] }
        catch { Write-Host "    (could not set document property '$($prop[0])' -- continuing)" -ForegroundColor DarkGray }
    }
    try { $wb.RemovePersonalInformation = $true } catch { }

    if (Test-Path $OutputFile) { Remove-Item $OutputFile -Force }
    $wb.SaveAs($OutputFile, 51)      # xlOpenXMLWorkbook
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
