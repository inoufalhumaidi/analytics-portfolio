<#
================================================================================
Project 5 -- Meridian UAV Services: Predictive Maintenance
Script:  Build_Workbook.ps1
Purpose: Build Meridian_Fleet_Control.xlsx from the CSV extracts.

A CONTROL, NOT A REPORT

    Nothing on the Dashboard is a pasted number. Both sides of the headline
    comparison are computed IN EXCEL, with COUNTIF over the component extract --
    hour compliance and stress compliance alike. A comparison where one side is
    copied from the tool that produced the other proves nothing about either.

    Two cells change the answer rather than the appearance:
      * Hangar hours per week re-cuts the queue and re-states the backlog;
      * Alert threshold re-states the prediction metrics from the sweep.

    Change either and the workbook answers a different question, with no
    rebuild and no SQL.

FOUR THINGS THIS SCRIPT DOES IN A PARTICULAR ORDER, AND WHY

 1. EVERY SHEET IS CREATED BEFORE ANY FORMULA IS WRITTEN.
    Excel rewrites a reference to a sheet that does not exist yet as #REF!, and
    does NOT repair it when the sheet later appears. The workbook then opens
    looking finished and reads #REF! in the cells nobody scrolled to.

 2. NO DYNAMIC ARRAYS.
    FILTER, SORT, UNIQUE and XLOOKUP do not exist in Excel 2019 or LibreOffice.
    A portfolio workbook that only opens in a current Microsoft 365 build is a
    workbook most reviewers cannot open. SUMPRODUCT, INDEX/MATCH and COUNTIFS
    throughout.

 3. THE DASHBOARD LOOKS ROWS UP BY NAME.
    Project 1 shipped a dashboard that read a total from a fixed row offset.
    Someone inserted a KPI row above it and the cost went to -163,307.32 --
    still formatted as currency, still plausible. Lookups here are by label and
    the helper raises if the label is not found.

 4. PERSONAL METADATA IS STRIPPED BEFORE SAVING.
    RemovePersonalInformation. A workbook published from a personal machine
    otherwise carries the author's name and path in its properties.

DATA DISCLOSURE: Meridian UAV Services is fictional; all data is synthetic. No
confidential data is used and no claim is made about any production system.
================================================================================
#>

param(
    [string]$ExtractPath = (Join-Path (Split-Path $PSScriptRoot -Parent) 'data_exports'),
    [string]$OutputPath  = (Join-Path $PSScriptRoot 'Meridian_Fleet_Control.xlsx')
)

$ErrorActionPreference = 'Stop'

<#
EVERY Value2 ASSIGNMENT IS EXPLICITLY CAST, AND IT HAS TO BE.

Windows PowerShell 5.1 binds Range.Value2 through COM late binding, and passing
an unpinned variable -- even a plain string from a foreach -- can fail with
"Specified cast is not valid". It is not deterministic across assignment sites:
this script wrote strings to Value2 in eight places and failed at one of them,
the scorecard loop, because that was the only site handing over a bare loop
variable rather than an indexed element of a literal array.

The failure is a hard stop with no cell reference, thirty lines into a build
that has already reported success for everything before it. Casting at every
site costs nothing and removes the class.
#>
function RGBv([int]$r, [int]$g, [int]$b) { return $r + ($g * 256) + ($b * 65536) }
$NAVY  = RGBv 31 78 121
$RED   = RGBv 192 0 0
$AMBER = RGBv 191 143 0
$GREEN = RGBv 84 130 53
$GREY  = RGBv 89 89 89
$LIGHT = RGBv 242 242 242

if (-not (Test-Path $ExtractPath)) { throw "Extract folder not found: $ExtractPath. Run data_exports/Export_Extracts.ps1 first." }

Write-Host ""
Write-Host "Meridian UAV -- building the fleet control workbook" -ForegroundColor Cyan
Write-Host ("-" * 78)

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.ScreenUpdating = $false

try {
    $wb  = $xl.Workbooks.Add()
    while ($wb.Worksheets.Count -gt 1) { $wb.Worksheets.Item($wb.Worksheets.Count).Delete() }

    # ------------------------------------------------------------- Control
    $ctl = $wb.Worksheets.Item(1)
    $ctl.Name = 'Control'
    $ctl.Range('A1').Value2 = 'MERIDIAN UAV SERVICES -- FLEET MAINTENANCE CONTROL'
    $ctl.Range('A1').Font.Size = 15; $ctl.Range('A1').Font.Bold = $true
    $ctl.Range('A1').Font.Color = $NAVY
    $ctl.Range('A2').Value2 = 'Control panel. Every sheet reads its parameters from the cells below.'
    $ctl.Range('A2').Font.Italic = $true; $ctl.Range('A2').Font.Color = $GREY

    $ctl.Range('A4').Value2 = 'PARAMETER'
    $ctl.Range('B4').Value2 = 'VALUE'
    $ctl.Range('C4').Value2 = 'WHAT IT CONTROLS'
    $ctl.Range('A4:C4').Font.Bold = $true
    $ctl.Range('A4:C4').Interior.Color = $NAVY
    $ctl.Range('A4:C4').Font.Color = RGBv 255 255 255

    $params = @(
        @('Source folder',      'SourceFolder',       $ExtractPath, 'text',
          'Where the CSV extracts live. Power Query reads this cell, so the workbook moves with its data.'),
        @('Reporting date',     'AsOfDate',           '2026-09-30', 'text',
          'Pinned, not derived from the data. A figure anchored on "the latest row" stops matching the report that quoted it the moment anything is loaded.'),
        @('Hangar hours / week','HangarHoursPerWeek', 120,          'number',
          'Maintenance capacity for the whole fleet. Change it and the week plan re-cuts itself -- this is how "what does another bay buy us" gets answered.'),
        @('Alert threshold',    'AlertThreshold',     3.40,         'number',
          'Vibration multiple of a component''s own baseline at which the alarm fires. Change it and the prediction figures re-state from the sweep.')
    )
    $r = 5
    foreach ($p in $params) {
        $ctl.Range("A$r").Value2 = [string]$p[0]
        if ($p[3] -eq 'number') { $ctl.Range("B$r").Value2 = [double]$p[2] }
        else                    { $ctl.Range("B$r").Value2 = [string]$p[2] }
        $ctl.Range("B$r").Interior.Color = RGBv 255 242 204
        $ctl.Range("B$r").Font.Bold = $true
        $ctl.Range("C$r").Value2 = [string]$p[4]
        $wb.Names.Add($p[1], $ctl.Range("B$r")) | Out-Null
        $r++
    }

    $ctl.Range('A11').Value2 = 'DATA DISCLOSURE'
    $ctl.Range('A11').Font.Bold = $true
    $ctl.Range('A12').Value2 = 'Meridian UAV Services is a fictional UAV operator. All data in this workbook is synthetic and generated by the scripts in the sql\ folder. Degradation behaviour is modelled on publicly described characteristics of rotating machinery. No confidential data is used and no claim is made about any real aircraft, operator or manufacturer.'
    $ctl.Range('A12:C14').Merge(); $ctl.Range('A12').WrapText = $true
    $ctl.Range('A12').VerticalAlignment = -4160
    $ctl.Columns.Item(1).ColumnWidth = 24
    $ctl.Columns.Item(2).ColumnWidth = 44
    $ctl.Columns.Item(3).ColumnWidth = 96
    Write-Host "  Control sheet            built"

    # --------------------------------------------------------- Power Query
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

    # Types are declared explicitly. Letting Power Query guess means a serial
    # that happens to look numeric arrives as a number in one refresh and text
    # in the next, and every lookup against it breaks on the refresh nobody was
    # watching.
    $queries = @(
      @{ N='tWear'; F='component_wear.csv'; T='{"InstallKey",Int64.Type},{"ComponentSerial",type text},{"ComponentCode",type text},{"ComponentName",type text},{"Criticality",type text},{"LeadTimeFlightHours",type number},{"TailNumber",type text},{"ModelCode",type text},{"BaseCode",type text},{"BaseName",type text},{"PositionNo",Int64.Type},{"InstalledDate",type date},{"Sorties",Int64.Type},{"Landings",Int64.Type},{"FlightHours",type number},{"StressHours",type number},{"IntervalFlightHours",type number},{"IntervalStressHours",type number},{"PctOfHourInterval",type number},{"PctOfStressInterval",type number},{"IsOverdueByHours",Int64.Type},{"IsOverdueByStress",Int64.Type},{"IsHiddenOverdue",Int64.Type},{"StressPerFlightHour",type number},{"RemainingStressHours",type number},{"RemainingFlightHours",type number},{"ProjectedFlightHoursToStressLimit",type number},{"LastSortieDate",type date}' },
      @{ N='tQueue'; F='action_queue.csv'; T='{"PriorityRank",Int64.Type},{"InstallKey",Int64.Type},{"ComponentSerial",type text},{"ComponentCode",type text},{"ComponentName",type text},{"Criticality",type text},{"TailNumber",type text},{"BaseCode",type text},{"BaseName",type text},{"PositionNo",Int64.Type},{"InstalledDate",type date},{"FlightHours",type number},{"StressHours",type number},{"IntervalFlightHours",type number},{"IntervalStressHours",type number},{"PctOfHourInterval",type number},{"PctOfStressInterval",type number},{"IsOverdueByHours",Int64.Type},{"IsOverdueByStress",Int64.Type},{"IsHiddenOverdue",Int64.Type},{"StressPerFlightHour",type number},{"RemainingStressHours",type number},{"ProjectedFlightHoursToStressLimit",type number},{"LeadTimeFlightHours",type number},{"ActionCode",type text},{"JobHours",type number},{"CumulativeJobHours",type number},{"IsThisWeek",Int64.Type},{"RecommendedAction",type text}' },
      @{ N='tBase'; F='base_scorecard.csv'; T='{"BaseCode",type text},{"BaseName",type text},{"Region",type text},{"HangarBays",Int64.Type},{"Airframes",Int64.Type},{"FlightHours",type number},{"StressHours",type number},{"StressRatio",type number},{"FittedComponents",Int64.Type},{"HourCompliancePct",type number},{"StressCompliancePct",type number},{"ComplianceGapPoints",type number},{"HiddenOverdue",Int64.Type},{"HiddenOverdueFlightCritical",Int64.Type},{"UnscheduledRatePct",type number},{"MaintenanceEvents",Int64.Type},{"DowntimeHours",type number}' },
      @{ N='tKPI'; F='fleet_kpi.csv'; T='{"MetricName",type text},{"MetricValue",type number},{"Numerator",type number},{"Denominator",type number},{"TargetValue",type number},{"WarningValue",type number},{"Direction",type text},{"Unit",type text},{"RAGStatus",type text},{"Description",type text}' },
      @{ N='tSweep'; F='threshold_sweep.csv'; T='{"VibThreshold",type number},{"ComponentCode",type text},{"Population",Int64.Type},{"Failures",Int64.Type},{"AlertsRaised",Int64.Type},{"TP",Int64.Type},{"FP",Int64.Type},{"FN",Int64.Type},{"TN",Int64.Type},{"ActionableTP",Int64.Type},{"LateTP",Int64.Type},{"PrecisionPct",type number},{"RecallPct",type number},{"ActionableLeadTimePct",type number},{"PreventableFailurePct",type number},{"MedianLeadTimeFlightHours",type number},{"LeadTimeFlightHours",type number}' },
      @{ N='tTargets'; F='ref_fleet_targets.csv'; T='{"MetricName",type text},{"TargetValue",type number},{"WarningValue",type number},{"Direction",type text},{"Unit",type text},{"Description",type text}' },
      @{ N='tStress'; F='ref_stress_model.csv'; T='{"ProfileCode",type text},{"StressMultiplier",type number},{"CyclePenaltyHours",type number},{"PayloadPenaltyPerKg",type number},{"SourceNote",type text}' },
      @{ N='tInterval'; F='ref_service_interval.csv'; T='{"ComponentCode",type text},{"IntervalFlightHours",type number},{"IntervalStressHours",type number},{"WeibullShape",type number},{"IntervalNote",type text}' },
      @{ N='tTrend'; F='compliance_trend.csv'; T='{"MonthEndDate",type date},{"MetricName",type text},{"MetricValue",type number}' },
      @{ N='tDQ'; F='dq_findings.csv'; T='{"CheckCode",type text},{"CheckName",type text},{"Severity",type text},{"EntityType",type text},{"EntityRef",type text},{"Detail",type text},{"AffectsCurrentWear",Int64.Type}' }
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
        Write-Host ("    {0,-10} {1,6} rows loaded" -f $q.N, $loaded[$q.N])
    }

    # A load that returns nothing still produces a workbook full of zeros, and
    # zero overdue components reads as a healthy fleet.
    if ($loaded['tWear']  -lt 100) { throw "tWear loaded only $($loaded['tWear']) rows -- the Power Query load did not work." }
    if ($loaded['tQueue'] -lt 50)  { throw "tQueue loaded only $($loaded['tQueue']) rows." }
    if ($loaded['tSweep'] -lt 10)  { throw "tSweep loaded only $($loaded['tSweep']) rows." }
    $nQueue = $loaded['tQueue']

    # Create every remaining sheet BEFORE writing a single formula. See note 1.
    $sheetOrder = @('README','Dashboard','Action Queue','Bases','Prediction','Validation','Calc')
    $made = @{}
    foreach ($name in $sheetOrder) {
        $s = $wb.Worksheets.Add([System.Reflection.Missing]::Value, $wb.Worksheets.Item($wb.Worksheets.Count))
        $s.Name = $name
        $made[$name] = $s
    }
    Write-Host "  Worksheets               $($sheetOrder.Count) created up front"

    # ------------------------------------------------------------------ Calc
    # The queue cut, computed in Excel from the queue's own job hours. The
    # RANKING comes from SQL; only the cumulative total and the week boundary
    # are recomputed here, so the workbook cannot hold a second opinion about
    # what is urgent.
    $calc = $made['Calc']
    $calc.Range('A1').Value2 = 'CALC -- the week cut, recomputed in Excel. Not a reporting sheet.'
    $calc.Range('A1').Font.Bold = $true; $calc.Range('A1').Font.Color = $GREY
    $calc.Range('A2').Value2 = 'SQL ranks the queue. Excel re-adds the hours and finds the line, so the Hangar hours cell on Control actually does something.'
    $calc.Range('A2').Font.Italic = $true; $calc.Range('A2').Font.Color = $GREY
    $calc.Range('A4').Value2 = 'Rank'
    $calc.Range('B4').Value2 = 'JobHours'
    $calc.Range('C4').Value2 = 'Cumulative (Excel)'
    $calc.Range('D4').Value2 = 'Within the week'
    $calc.Range('A4:D4').Font.Bold = $true
    for ($i = 1; $i -le $nQueue; $i++) {
        $row = 4 + $i
        $calc.Range("A$row").Value2 = [double]$i
        $calc.Range("B$row").Formula = "=SUMIFS(tQueue[JobHours],tQueue[PriorityRank],A$row)"
        $calc.Range("C$row").Formula = "=SUMIFS(tQueue[JobHours],tQueue[PriorityRank],`"<=`"&A$row)"
        $calc.Range("D$row").Formula = "=IF(C$row<=HangarHoursPerWeek,1,0)"
    }
    $calcLast = 4 + $nQueue
    $wb.Names.Add('CalcInWeek', $calc.Range("D5:D$calcLast")) | Out-Null
    Write-Host "  Calc sheet               $nQueue queue rows"

    # ------------------------------------------------------------- Dashboard
    $d = $made['Dashboard']
    $d.Range('A1').Value2 = 'IS THIS FLEET INSIDE ITS MAINTENANCE INTERVALS?'
    $d.Range('A1').Font.Size = 16; $d.Range('A1').Font.Bold = $true; $d.Range('A1').Font.Color = $NAVY
    $d.Range('A2').Formula = '="Reporting date " & AsOfDate & "   |   both figures below are computed in this workbook from the component extract, not copied from SQL."'
    $d.Range('A2').Font.Italic = $true; $d.Range('A2').Font.Color = $GREY

    $d.Range('A4').Value2 = 'THE SAME COMPONENTS, THE SAME INTERVALS, THE SAME DAY'
    $d.Range('A4').Font.Bold = $true; $d.Range('A4').Font.Color = $NAVY

    $d.Range('A5').Value2 = 'MEASURE'; $d.Range('B5').Value2 = 'VALUE'; $d.Range('C5').Value2 = 'WHAT IT ACTUALLY MEANS'
    $d.Range('A5:C5').Font.Bold = $true; $d.Range('A5:C5').Interior.Color = $LIGHT

    $headline = @(
        @('Compliant on FLIGHT HOURS',
          '=COUNTIFS(tWear[IsOverdueByHours],0)/COUNTA(tWear[InstallKey])',
          '0.00%',
          'What the maintenance system reports. The hour meter is what it enforces against.'),
        @('Compliant on STRESS HOURS',
          '=COUNTIFS(tWear[IsOverdueByStress],0)/COUNTA(tWear[InstallKey])',
          '0.00%',
          'The same interval measured in the unit the wear actually accrues in.'),
        @('The gap',
          '=B6-B7',
          '0.00%',
          'Components past due that no current report shows.')
    )
    $row = 6
    foreach ($h in $headline) {
        $d.Range("A$row").Value2 = [string]$h[0]
        $d.Range("B$row").Formula = [string]$h[1]
        $d.Range("B$row").NumberFormat = [string]$h[2]
        $d.Range("B$row").Font.Bold = $true
        $d.Range("B$row").Font.Size = 13
        $d.Range("C$row").Value2 = [string]$h[3]
        $row++
    }
    $d.Range('B6').Font.Color = $AMBER
    $d.Range('B7').Font.Color = $RED
    $d.Range('B8').Font.Color = $RED

    $d.Range('A10').Value2 = 'An hour is not a unit of wear. An hour of mountain survey at 22kg and an hour of coastal survey at 12kg damage a rotor bearing by very different amounts, and an hour of circuit training produces ten landings where cargo relay produces one. The interval is enforced in hours because that is what the hour meter produces.'
    $d.Range('A10:C11').Merge(); $d.Range('A10').WrapText = $true
    $d.Range('A10').Font.Italic = $true; $d.Range('A10').Font.Color = $GREY
    $d.Range('A10').VerticalAlignment = -4160

    $d.Range('A13').Value2 = 'WHAT THAT MEANS IN COMPONENTS'
    $d.Range('A13').Font.Bold = $true; $d.Range('A13').Font.Color = $NAVY
    $counts = @(
        @('Components fitted',                     '=COUNTA(tWear[InstallKey])', '#,##0', 'The population every percentage above is over.'),
        @('Past due on stress, inside hour limit', '=COUNTIFS(tWear[IsHiddenOverdue],1)', '#,##0', 'Invisible to every report the operator currently runs.'),
        @('...of which flight-critical',           '=COUNTIFS(tWear[IsHiddenOverdue],1,tWear[Criticality],"FlightCritical")', '#,##0', 'Failure grounds the aircraft.'),
        @('Past due on stress (all)',              '=COUNTIFS(tWear[IsOverdueByStress],1)', '#,##0', 'Including those the hour meter can already see.'),
        @('Past due on flight hours',              '=COUNTIFS(tWear[IsOverdueByHours],1)', '#,##0', 'What the maintenance system is chasing today.')
    )
    $d.Range('A14').Value2 = 'MEASURE'; $d.Range('B14').Value2 = 'COUNT'; $d.Range('C14').Value2 = 'NOTE'
    $d.Range('A14:C14').Font.Bold = $true; $d.Range('A14:C14').Interior.Color = $LIGHT
    $row = 15
    foreach ($c in $counts) {
        $d.Range("A$row").Value2 = [string]$c[0]
        $d.Range("B$row").Formula = [string]$c[1]
        $d.Range("B$row").NumberFormat = [string]$c[2]
        $d.Range("B$row").Font.Bold = $true
        $d.Range("C$row").Value2 = [string]$c[3]
        $row++
    }

    # --- the scorecard, bound to Ref_FleetTargets by NAME --------------------
    $d.Range('A21').Value2 = 'FLEET SCORECARD -- against the targets in Ref_FleetTargets'
    $d.Range('A21').Font.Bold = $true; $d.Range('A21').Font.Color = $NAVY
    $d.Range('A22').Value2 = 'METRIC'; $d.Range('B22').Value2 = 'VALUE'
    $d.Range('C22').Value2 = 'TARGET'; $d.Range('D22').Value2 = 'WARN'; $d.Range('E22').Value2 = 'STATUS'
    $d.Range('A22:E22').Font.Bold = $true; $d.Range('A22:E22').Interior.Color = $LIGHT

    $metrics = @('StressCompliancePct','HourCompliancePct','OverdueFlightCritical','UnscheduledRatePct','AirframeAvailabilityPct')
    $row = 23
    foreach ($m in $metrics) {
        $d.Range("A$row").Value2 = [string]$m
        # The value comes from SQL; the TARGET and DIRECTION are looked up by
        # metric NAME, never by row position -- see note 3.
        $d.Range("B$row").Formula = "=INDEX(tKPI[MetricValue],MATCH(A$row,tKPI[MetricName],0))"
        $d.Range("C$row").Formula = "=INDEX(tTargets[TargetValue],MATCH(A$row,tTargets[MetricName],0))"
        $d.Range("D$row").Formula = "=INDEX(tTargets[WarningValue],MATCH(A$row,tTargets[MetricName],0))"
        $d.Range("E$row").Formula = "=IF(INDEX(tTargets[Direction],MATCH(A$row,tTargets[MetricName],0))=""HigherBetter"",IF(B$row>=C$row,""Green"",IF(B$row>=D$row,""Amber"",""Red"")),IF(B$row<=C$row,""Green"",IF(B$row<=D$row,""Amber"",""Red"")))"
        $d.Range("B$row:D$row").NumberFormat = '#,##0.00'
        $row++
    }
    $lastKpi = $row - 1
    foreach ($spec in @(@('Red', $RED), @('Amber', $AMBER), @('Green', $GREEN))) {
        $fc = $d.Range("E23:E$lastKpi").FormatConditions.Add(1, 3, "=""$($spec[0])""")
        $fc.Font.Color = $spec[1]; $fc.Font.Bold = $true
    }

    # --- the week's plan -----------------------------------------------------
    $d.Range('A30').Value2 = 'THE WEEK -- recomputed here from the hangar hours on Control'
    $d.Range('A30').Font.Bold = $true; $d.Range('A30').Font.Color = $NAVY
    $plan = @(
        @('Jobs outstanding',        '=COUNTA(tQueue[PriorityRank])', '#,##0',   'Across the whole fleet.'),
        @('Hangar hours to clear',   '=SUM(tQueue[JobHours])',        '#,##0.0', 'A procurement action costs no bay time and carries zero hours.'),
        @('Schedulable this week',   '=SUM(CalcInWeek)',              '#,##0',   'Where the cumulative hours cross the capacity on Control.'),
        @('Weeks at current capacity','=SUM(tQueue[JobHours])/HangarHoursPerWeek', '#,##0.0', 'Change the capacity cell and this moves.'),
        @('Ground and replace now',  '=COUNTIFS(tQueue[ActionCode],"GROUND_AND_REPLACE")', '#,##0', 'Flight-critical and already past its stress interval.'),
        @('Order a part now',        '=COUNTIFS(tQueue[ActionCode],"ORDER_PART_NOW")', '#,##0', 'A purchase order, not a bay. Different person, different day.')
    )
    $d.Range('A31').Value2 = 'MEASURE'; $d.Range('B31').Value2 = 'VALUE'; $d.Range('C31').Value2 = 'NOTE'
    $d.Range('A31:C31').Font.Bold = $true; $d.Range('A31:C31').Interior.Color = $LIGHT
    $row = 32
    foreach ($p in $plan) {
        $d.Range("A$row").Value2 = [string]$p[0]
        $d.Range("B$row").Formula = [string]$p[1]
        $d.Range("B$row").NumberFormat = [string]$p[2]
        $d.Range("B$row").Font.Bold = $true
        $d.Range("C$row").Value2 = [string]$p[3]
        $row++
    }

    $d.Columns.Item(1).ColumnWidth = 40
    $d.Columns.Item(2).ColumnWidth = 14
    $d.Columns.Item(3).ColumnWidth = 78
    $d.Columns.Item(4).ColumnWidth = 10
    $d.Columns.Item(5).ColumnWidth = 10
    Write-Host "  Dashboard                built"

    # ------------------------------------------------------------- Bases
    $b = $made['Bases']
    $b.Range('A1').Value2 = 'THE FLEET AVERAGE DESCRIBES NONE OF THE FOUR BASES'
    $b.Range('A1').Font.Size = 14; $b.Range('A1').Font.Bold = $true; $b.Range('A1').Font.Color = $NAVY
    $b.Range('A2').Value2 = 'StressRatio is stress hours accrued per logged flight hour. Read across: the hour meter cannot tell these bases apart, and the work they fly is not comparable.'
    $b.Range('A2').Font.Italic = $true; $b.Range('A2').Font.Color = $GREY

    $bh = @('Base','Airframes','Stress / flight hr','Hour compliance %','Stress compliance %','Gap (pts)','Hidden overdue','Unscheduled %')
    for ($i = 0; $i -lt $bh.Count; $i++) { $b.Cells.Item(4, $i + 1).Value2 = [string]$bh[$i] }
    $b.Range('A4:H4').Font.Bold = $true; $b.Range('A4:H4').Interior.Color = $NAVY
    $b.Range('A4:H4').Font.Color = RGBv 255 255 255

    for ($i = 1; $i -le $loaded['tBase']; $i++) {
        $row = 4 + $i
        # Ordered by stress compliance, worst first, using SMALL/MATCH rather
        # than SORT -- which does not exist outside current Microsoft 365.
        $b.Range("A$row").Formula = "=INDEX(tBase[BaseCode],MATCH(SMALL(tBase[StressCompliancePct],$i),tBase[StressCompliancePct],0))"
        $b.Range("B$row").Formula = "=INDEX(tBase[Airframes],MATCH(`$A$row,tBase[BaseCode],0))"
        $b.Range("C$row").Formula = "=INDEX(tBase[StressRatio],MATCH(`$A$row,tBase[BaseCode],0))"
        $b.Range("D$row").Formula = "=INDEX(tBase[HourCompliancePct],MATCH(`$A$row,tBase[BaseCode],0))"
        $b.Range("E$row").Formula = "=INDEX(tBase[StressCompliancePct],MATCH(`$A$row,tBase[BaseCode],0))"
        $b.Range("F$row").Formula = "=D$row-E$row"
        $b.Range("G$row").Formula = "=INDEX(tBase[HiddenOverdue],MATCH(`$A$row,tBase[BaseCode],0))"
        $b.Range("H$row").Formula = "=INDEX(tBase[UnscheduledRatePct],MATCH(`$A$row,tBase[BaseCode],0))"
        $b.Range("C$row").NumberFormat = '0.0000'
        $b.Range("D$row:F$row").NumberFormat = '0.00'
        $b.Range("H$row").NumberFormat = '0.00'
    }
    $lastBase = 4 + $loaded['tBase']
    $fcB = $b.Range("E5:E$lastBase").FormatConditions.Add(1, 6, '=70')
    $fcB.Font.Color = $RED; $fcB.Font.Bold = $true

    $b.Range("A$($lastBase + 2)").Value2 = 'Highland and Port are not worse-run bases; they fly mountain and cargo work, which is what the multiplier measures. The Academy is the one worth pausing on: it flies the lightest payloads and the fewest hours in the fleet, and circuit training produces so many landings that it wears its aircraft harder than the cargo hub.'
    $b.Range("A$($lastBase + 2):H$($lastBase + 4)").Merge()
    $b.Range("A$($lastBase + 2)").WrapText = $true
    $b.Range("A$($lastBase + 2)").Font.Italic = $true
    $b.Range("A$($lastBase + 2)").VerticalAlignment = -4160
    $b.Columns.Item(1).ColumnWidth = 14
    for ($i = 2; $i -le 8; $i++) { $b.Columns.Item($i).ColumnWidth = 20 }
    Write-Host "  Bases sheet              built"

    # --------------------------------------------------------- Prediction
    $p = $made['Prediction']
    $p.Range('A1').Value2 = 'THE ALARM IS ACCURATE. THAT IS NOT THE SAME AS USEFUL.'
    $p.Range('A1').Font.Size = 14; $p.Range('A1').Font.Bold = $true; $p.Range('A1').Font.Color = $NAVY
    $p.Range('A2').Value2 = 'Change the Alert threshold on Control and every figure below re-states. Precision rises with it; lead time falls with it. There is no setting where both clear their target for both components.'
    $p.Range('A2').Font.Italic = $true; $p.Range('A2').Font.Color = $GREY

    $ph = @('Component','Part lead time (h)','Threshold','Alerts','Caught','Precision %','Recall %','Median warning (h)','Actionable %','Warning covers supply?')
    for ($i = 0; $i -lt $ph.Count; $i++) { $p.Cells.Item(4, $i + 1).Value2 = [string]$ph[$i] }
    $p.Range('A4:J4').Font.Bold = $true; $p.Range('A4:J4').Interior.Color = $NAVY
    $p.Range('A4:J4').Font.Color = RGBv 255 255 255

    $codes = @('ROTOR-ASSY','MOTOR-ESC')
    $row = 5
    foreach ($code in $codes) {
        $p.Range("A$row").Value2 = [string]$code
        # Two-key lookup without dynamic arrays: SUMIFS over the sweep, matching
        # both the component and the threshold from Control.
        $p.Range("B$row").Formula = "=SUMIFS(tSweep[LeadTimeFlightHours],tSweep[ComponentCode],`$A$row,tSweep[VibThreshold],AlertThreshold)"
        $p.Range("C$row").Formula = "=AlertThreshold"
        $p.Range("D$row").Formula = "=SUMIFS(tSweep[AlertsRaised],tSweep[ComponentCode],`$A$row,tSweep[VibThreshold],AlertThreshold)"
        $p.Range("E$row").Formula = "=SUMIFS(tSweep[TP],tSweep[ComponentCode],`$A$row,tSweep[VibThreshold],AlertThreshold)"
        $p.Range("F$row").Formula = "=SUMIFS(tSweep[PrecisionPct],tSweep[ComponentCode],`$A$row,tSweep[VibThreshold],AlertThreshold)"
        $p.Range("G$row").Formula = "=SUMIFS(tSweep[RecallPct],tSweep[ComponentCode],`$A$row,tSweep[VibThreshold],AlertThreshold)"
        $p.Range("H$row").Formula = "=SUMIFS(tSweep[MedianLeadTimeFlightHours],tSweep[ComponentCode],`$A$row,tSweep[VibThreshold],AlertThreshold)"
        $p.Range("I$row").Formula = "=SUMIFS(tSweep[ActionableLeadTimePct],tSweep[ComponentCode],`$A$row,tSweep[VibThreshold],AlertThreshold)"
        $p.Range("J$row").Formula = "=IF(H$row>=B$row,""yes"",""NO -- the median warning is shorter than the time to get the part"")"
        $p.Range("B$row:C$row").NumberFormat = '0.00'
        $p.Range("F$row:I$row").NumberFormat = '0.00'
        $row++
    }
    $lastPred = $row - 1

    # The fleet figure SUMS THE COUNTS. It does not average the two
    # percentages: an average weights a population of 130 the same as one of
    # 430 and produces a third number describing neither.
    $p.Range("A$($lastPred + 2)").Value2 = 'FLEET (counts summed, not percentages averaged)'
    $p.Range("A$($lastPred + 2)").Font.Bold = $true
    $p.Range("D$($lastPred + 2)").Formula = "=SUM(D5:D$lastPred)"
    $p.Range("E$($lastPred + 2)").Formula = "=SUM(E5:E$lastPred)"
    $p.Range("F$($lastPred + 2)").Formula = "=E$($lastPred + 2)/D$($lastPred + 2)*100"
    $p.Range("I$($lastPred + 2)").Formula = "=SUMPRODUCT(SUMIFS(tSweep[ActionableTP],tSweep[ComponentCode],A5:A$lastPred,tSweep[VibThreshold],AlertThreshold))/E$($lastPred + 2)*100"
    $p.Range("F$($lastPred + 2)").NumberFormat = '0.00'
    $p.Range("I$($lastPred + 2)").NumberFormat = '0.00'
    $p.Range("D$($lastPred + 2):I$($lastPred + 2)").Font.Bold = $true
    $wb.Names.Add('FleetActionablePct', $p.Range("I$($lastPred + 2)")) | Out-Null
    $wb.Names.Add('FleetPrecisionPct',  $p.Range("F$($lastPred + 2)")) | Out-Null

    $p.Range("A$($lastPred + 5)").Value2 = 'A prediction is worth its LEAD TIME, not its accuracy. A rotor assembly takes 18 flight hours to obtain and fit, so an alert with 6 hours of warning is a correct prediction and a grounded aircraft. Precision and recall cannot tell those two alerts apart -- both are true positives. Set the threshold to 2.05 and watch the motor controller: precision falls, and the share of its failures caught in time to act rises from single figures to over 90%.'
    $p.Range("A$($lastPred + 5):J$($lastPred + 8)").Merge()
    $p.Range("A$($lastPred + 5)").WrapText = $true
    $p.Range("A$($lastPred + 5)").Font.Italic = $true
    $p.Range("A$($lastPred + 5)").VerticalAlignment = -4160

    $p.Columns.Item(1).ColumnWidth = 16
    for ($i = 2; $i -le 9; $i++) { $p.Columns.Item($i).ColumnWidth = 18 }
    $p.Columns.Item(10).ColumnWidth = 58
    Write-Host "  Prediction sheet         built"

    # --------------------------------------------------------- Action Queue
    $q = $made['Action Queue']
    $q.Range('A1').Value2 = 'WHAT TO DO THIS WEEK'
    $q.Range('A1').Font.Size = 14; $q.Range('A1').Font.Bold = $true; $q.Range('A1').Font.Color = $NAVY
    $q.Range('A2').Formula = '="Capacity " & TEXT(HangarHoursPerWeek,"#,##0") & " hangar hours. SQL ranks the queue; the Within-week column is recomputed in this workbook."'
    $q.Range('A2').Font.Italic = $true; $q.Range('A2').Font.Color = $GREY

    $qh = @('Rank','Tail','Base','Component','Pos','Criticality','% hour interval','% stress interval','Projected hrs to limit','Part lead (h)','Action','Job hrs','Cumulative','Within week')
    for ($i = 0; $i -lt $qh.Count; $i++) { $q.Cells.Item(4, $i + 1).Value2 = [string]$qh[$i] }
    $q.Range('A4:N4').Font.Bold = $true; $q.Range('A4:N4').Interior.Color = $NAVY
    $q.Range('A4:N4').Font.Color = RGBv 255 255 255

    $showRows = [Math]::Min($nQueue, 120)
    for ($i = 1; $i -le $showRows; $i++) {
        $row = 4 + $i
        $q.Range("A$row").Value2 = [double]$i
        foreach ($map in @(@('B','TailNumber'), @('C','BaseCode'), @('D','ComponentCode'),
                           @('E','PositionNo'), @('F','Criticality'),
                           @('G','PctOfHourInterval'), @('H','PctOfStressInterval'),
                           @('I','ProjectedFlightHoursToStressLimit'), @('J','LeadTimeFlightHours'),
                           @('K','ActionCode'), @('L','JobHours'))) {
            $q.Range("$($map[0])$row").Formula =
                "=INDEX(tQueue[$($map[1])],MATCH(`$A$row,tQueue[PriorityRank],0))"
        }
        $q.Range("M$row").Formula = "=INDEX(Calc!`$C`$5:`$C`$$calcLast,MATCH(`$A$row,Calc!`$A`$5:`$A`$$calcLast,0))"
        $q.Range("N$row").Formula = "=IF(M$row<=HangarHoursPerWeek,""yes"",""-"")"
        $q.Range("G$row:J$row").NumberFormat = '0.00'
        $q.Range("L$row:M$row").NumberFormat = '0.0'
    }
    $lastQ = 4 + $showRows
    $fcQ = $q.Range("N5:N$lastQ").FormatConditions.Add(1, 3, '="yes"')
    $fcQ.Font.Color = $GREEN; $fcQ.Font.Bold = $true
    $q.Range("A$($lastQ + 2)").Formula = "=""Showing the top $showRows of "" & COUNTA(tQueue[PriorityRank]) & "" outstanding jobs. The full queue is on Data_Queue -- it is kept complete for audit, and cut here only for reading."""
    $q.Range("A$($lastQ + 2)").Font.Italic = $true; $q.Range("A$($lastQ + 2)").Font.Color = $GREY
    $q.Columns.Item(1).ColumnWidth = 7
    for ($i = 2; $i -le 14; $i++) { $q.Columns.Item($i).ColumnWidth = 17 }
    $q.Columns.Item(11).ColumnWidth = 24
    Write-Host "  Action Queue sheet       top $showRows of $nQueue"

    # ---------------------------------------------------------- Validation
    $v = $made['Validation']
    $v.Range('A1').Value2 = 'VALIDATION -- Excel against SQL'
    $v.Range('A1').Font.Size = 14; $v.Range('A1').Font.Bold = $true; $v.Range('A1').Font.Color = $NAVY
    $v.Range('A2').Value2 = 'Each row recomputes a figure in Excel from the raw extract and compares it to what SQL published. A workbook that agrees with itself proves nothing.'
    $v.Range('A2').Font.Italic = $true; $v.Range('A2').Font.Color = $GREY
    $v.Range('A3').Value2 = 'This proves the two AGREE. It does not prove either is right -- Project 1 shipped a workbook and a view that agreed exactly and were both wrong, because both came from the same flawed written definition.'
    $v.Range('A3').Font.Italic = $true; $v.Range('A3').Font.Color = $RED

    $v.Range('A5').Value2 = '#'; $v.Range('B5').Value2 = 'CHECK'
    $v.Range('C5').Value2 = 'EXCEL'; $v.Range('D5').Value2 = 'SQL'; $v.Range('E5').Value2 = 'RESULT'
    $v.Range('A5:E5').Font.Bold = $true; $v.Range('A5:E5').Interior.Color = $NAVY
    $v.Range('A5:E5').Font.Color = RGBv 255 255 255

    $checks = @(
        @('Stress compliance %',        '=COUNTIFS(tWear[IsOverdueByStress],0)/COUNTA(tWear[InstallKey])*100', '=INDEX(tKPI[MetricValue],MATCH("StressCompliancePct",tKPI[MetricName],0))'),
        @('Hour compliance %',          '=COUNTIFS(tWear[IsOverdueByHours],0)/COUNTA(tWear[InstallKey])*100',  '=INDEX(tKPI[MetricValue],MATCH("HourCompliancePct",tKPI[MetricName],0))'),
        @('Overdue flight-critical',    '=COUNTIFS(tWear[IsOverdueByStress],1,tWear[Criticality],"FlightCritical")', '=INDEX(tKPI[MetricValue],MATCH("OverdueFlightCritical",tKPI[MetricName],0))'),
        @('Components fitted',          '=COUNTA(tWear[InstallKey])',                     '=INDEX(tKPI[Denominator],MATCH("StressCompliancePct",tKPI[MetricName],0))'),
        @('Hidden overdue',             '=COUNTIFS(tWear[IsHiddenOverdue],1)',            '=SUM(tBase[HiddenOverdue])'),
        @('Hidden overdue, flt-crit',   '=COUNTIFS(tWear[IsHiddenOverdue],1,tWear[Criticality],"FlightCritical")', '=SUM(tBase[HiddenOverdueFlightCritical])'),
        @('Base fitted components sum', '=SUM(tBase[FittedComponents])',                  '=COUNTA(tWear[InstallKey])'),
        @('Queue rows',                 '=COUNTA(tQueue[PriorityRank])',                  '=MAX(tQueue[PriorityRank])'),
        @('Queue hangar hours',         '=SUM(tQueue[JobHours])',                         '=MAX(tQueue[CumulativeJobHours])'),
        @('Ground-and-replace jobs',    '=COUNTIFS(tQueue[ActionCode],"GROUND_AND_REPLACE")', '=COUNTIFS(tWear[IsOverdueByStress],1,tWear[Criticality],"FlightCritical")'),
        @('Rotor precision at threshold','=SUMIFS(tSweep[PrecisionPct],tSweep[ComponentCode],"ROTOR-ASSY",tSweep[VibThreshold],AlertThreshold)', '=SUMIFS(tSweep[TP],tSweep[ComponentCode],"ROTOR-ASSY",tSweep[VibThreshold],AlertThreshold)/SUMIFS(tSweep[AlertsRaised],tSweep[ComponentCode],"ROTOR-ASSY",tSweep[VibThreshold],AlertThreshold)*100'),
        @('Motor precision at threshold','=SUMIFS(tSweep[PrecisionPct],tSweep[ComponentCode],"MOTOR-ESC",tSweep[VibThreshold],AlertThreshold)', '=SUMIFS(tSweep[TP],tSweep[ComponentCode],"MOTOR-ESC",tSweep[VibThreshold],AlertThreshold)/SUMIFS(tSweep[AlertsRaised],tSweep[ComponentCode],"MOTOR-ESC",tSweep[VibThreshold],AlertThreshold)*100'),
        @('Actionable + late = caught', '=SUMIFS(tSweep[ActionableTP],tSweep[VibThreshold],AlertThreshold)+SUMIFS(tSweep[LateTP],tSweep[VibThreshold],AlertThreshold)', '=SUMIFS(tSweep[TP],tSweep[VibThreshold],AlertThreshold)'),
        @('Confusion matrix partitions','=SUMIFS(tSweep[TP],tSweep[VibThreshold],AlertThreshold)+SUMIFS(tSweep[FP],tSweep[VibThreshold],AlertThreshold)+SUMIFS(tSweep[FN],tSweep[VibThreshold],AlertThreshold)+SUMIFS(tSweep[TN],tSweep[VibThreshold],AlertThreshold)', '=SUMIFS(tSweep[Population],tSweep[VibThreshold],AlertThreshold)'),
        @('Week cut within capacity',   '=SUMPRODUCT(CalcInWeek)',                        '=COUNTIFS(tQueue[IsThisWeek],1)'),
        @('DQ findings total',          '=COUNTA(tDQ[CheckCode])',                        '=COUNTA(tDQ[Detail])')
    )
    $row = 6
    $n = 1
    foreach ($c in $checks) {
        $v.Range("A$row").Value2 = [double]$n
        $v.Range("B$row").Value2 = [string]$c[0]
        $v.Range("C$row").Formula = [string]$c[1]
        $v.Range("D$row").Formula = [string]$c[2]
        # 0.01 absolute, because the SQL side is published rounded to two
        # decimals. Anything bigger than a rounding step is a disagreement.
        $v.Range("E$row").Formula = "=IF(ABS(C$row-D$row)<0.011,""OK"",""MISMATCH"")"
        $v.Range("C$row:D$row").NumberFormat = '#,##0.00'
        $row++; $n++
    }
    $lastV = $row - 1
    $fcOK = $v.Range("E6:E$lastV").FormatConditions.Add(1, 3, '="OK"')
    $fcOK.Font.Color = $GREEN; $fcOK.Font.Bold = $true
    $fcNo = $v.Range("E6:E$lastV").FormatConditions.Add(1, 3, '="MISMATCH"')
    $fcNo.Font.Color = $RED; $fcNo.Font.Bold = $true; $fcNo.Interior.Color = RGBv 255 230 230

    $v.Range("B$($lastV + 2)").Value2 = 'CHECKS PASSING'
    $v.Range("B$($lastV + 2)").Font.Bold = $true
    $v.Range("C$($lastV + 2)").Formula = "=COUNTIF(E6:E$lastV,""OK"") & "" of "" & COUNTA(E6:E$lastV)"
    $v.Range("C$($lastV + 2)").Font.Bold = $true; $v.Range("C$($lastV + 2)").Font.Size = 12
    $wb.Names.Add('ValidationSummary', $v.Range("C$($lastV + 2)")) | Out-Null
    $wb.Names.Add('ValidationRange',   $v.Range("E6:E$lastV")) | Out-Null

    $v.Columns.Item(1).ColumnWidth = 5
    $v.Columns.Item(2).ColumnWidth = 34
    $v.Columns.Item(3).ColumnWidth = 16
    $v.Columns.Item(4).ColumnWidth = 16
    $v.Columns.Item(5).ColumnWidth = 14
    Write-Host "  Validation sheet         $($checks.Count) checks"

    # ------------------------------------------------------------- README
    $rd = $made['README']
    $rd.Range('A1').Value2 = 'MERIDIAN UAV SERVICES -- FLEET MAINTENANCE CONTROL'
    $rd.Range('A1').Font.Size = 16; $rd.Range('A1').Font.Bold = $true; $rd.Range('A1').Font.Color = $NAVY
    $lines = @(
        '',
        'THE QUESTION',
        'Is this fleet inside its maintenance intervals? Both answers on the Dashboard are correct. They are the same components, the same intervals and the same day, and they differ only in the unit the interval is measured in.',
        '',
        'HOW TO READ IT',
        '  Control        Four cells. Two of them change the answer: hangar hours per week, and the alert threshold.',
        '  Dashboard      The headline comparison, computed here in Excel from the component extract. Nothing on it is pasted.',
        '  Bases          The cut the fleet average hides. StressRatio explains the rest of the row.',
        '  Action Queue   What to do this week, cut at the capacity on Control.',
        '  Prediction     What the alarm is worth at the threshold on Control -- accuracy and usefulness side by side.',
        '  Validation     Sixteen figures recomputed in Excel and reconciled against SQL.',
        '  Data_*         The extracts, loaded through Power Query and kept complete.',
        '',
        'TRY THIS',
        '  1. Set Hangar hours / week to 240 on Control. The week plan roughly doubles and the backlog halves.',
        '  2. Set Alert threshold to 2.05. Motor-controller precision falls and its actionable share rises from single figures to over 90%. That trade is the second finding.',
        '',
        'WHAT IS DELIBERATELY NOT HERE',
        '  No dynamic-array formulas. FILTER, SORT, UNIQUE and XLOOKUP would break this file on Excel 2019 and LibreOffice.',
        '  No macros. A portfolio workbook that asks a reviewer to enable content is a workbook they will not open.',
        '  No pasted headline figures. Both sides of the comparison are computed here, because a comparison where one side is copied from the other proves nothing.',
        '',
        'DATA DISCLOSURE',
        'Meridian UAV Services is a fictional UAV operator. All data is synthetic and generated by the scripts in the sql\ folder. Degradation behaviour is modelled on publicly described characteristics of rotating machinery. No confidential data is used and no claim is made about any real aircraft, operator or manufacturer.'
    )
    $row = 2
    foreach ($line in $lines) {
        $rd.Range("A$row").Value2 = [string]$line
        if ($line -cmatch '^[A-Z][A-Z ,\-]+$') {
            $rd.Range("A$row").Font.Bold = $true
            $rd.Range("A$row").Font.Color = $NAVY
        }
        $row++
    }
    $rd.Columns.Item(1).ColumnWidth = 130
    $rd.Range("A2:A$row").WrapText = $false
    Write-Host "  README sheet             built"

    # ------------------------------------------------------------ finalise
    foreach ($name in @('Data_Wear','Data_Queue','Data_Base','Data_KPI','Data_Sweep',
                        'Data_Targets','Data_Stress','Data_Interval','Data_Trend','Data_DQ','Calc')) {
        try { $wb.Worksheets.Item($name).Visible = 0 } catch { }   # hidden, not deleted
    }
    $order = @('README','Dashboard','Bases','Action Queue','Prediction','Validation','Control')
    for ($i = 0; $i -lt $order.Count; $i++) {
        $wb.Worksheets.Item($order[$i]).Move($wb.Worksheets.Item($i + 1))
    }
    $wb.Worksheets.Item('Dashboard').Activate()

    $xl.CalculateFullRebuild()
    $wb.RemovePersonalInformation = $true

    if (Test-Path $OutputPath) { Remove-Item $OutputPath -Force }
    $wb.SaveAs($OutputPath, 51)   # 51 = xlOpenXMLWorkbook
    Write-Host ("-" * 78)
    Write-Host "  Saved to $OutputPath" -ForegroundColor Green
    $wb.Close($false)
}
finally {
    $xl.Quit()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}

Write-Host ""
Write-Host "Built. Now run Validate_Workbook.ps1 -- building a file is not the same as the file working." -ForegroundColor Yellow
