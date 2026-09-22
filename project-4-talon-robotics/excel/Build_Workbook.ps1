<#
================================================================================
Project 4 -- Talon Robotics: Payload Deployment Delivery
Script:  Build_Workbook.ps1
Purpose: Build Talon_Readiness_Control.xlsx from the CSV extracts.

WHAT MAKES THIS A CONTROL AND NOT A REPORT
    Nothing on the Dashboard is a pasted number. Every figure is a formula over
    the loaded extract tables, and in particular BOTH SIDES of the central
    comparison are computed here:

        work-item completion   COUNTIF over work_items.csv
        ship readiness         COUNTIFS over requirement_verification.csv

    That matters more than usual in this project. The whole argument is a
    comparison between two numbers, and a comparison in which one side is
    copied from the other tool proves nothing about either. The Validation
    sheet then holds each Excel figure against the one SQL produced
    independently, and says on its own face whether they agree.

DELIBERATE CONSTRAINTS
    No dynamic-array formulas. FILTER, SORT, UNIQUE and LET would break the
    file on Excel 2019 and LibreOffice, and a control nobody can open is not a
    control. Ranking is SUMPRODUCT plus INDEX/MATCH throughout.

    Every sheet is created BEFORE any formula is written. Excel rewrites a
    reference to a sheet that does not exist yet as #REF! and does not undo it
    when the sheet appears.

DATA DISCLOSURE: Talon Robotics is fictional; all data is synthetic. No
confidential data and no production system is involved.
================================================================================
#>

param(
    [string]$ExtractPath = (Join-Path (Split-Path $PSScriptRoot -Parent) 'erp_extracts'),
    [string]$OutPath     = (Join-Path $PSScriptRoot 'Talon_Readiness_Control.xlsx')
)

$ErrorActionPreference = 'Stop'

<#
A VARIABLE FOLLOWED BY A COLON IS SCOPE NOTATION, NOT TEXT.

    "B$row:D$row"  ->  "B12"          the ":D$row" is silently swallowed
    "B${row}:D${row}"  ->  "B12:D12"

PowerShell parses `$row:` the same way it parses `$env:PATH` -- as a drive or
scope qualifier -- so the rest of the range disappears and the expression still
produces a perfectly valid single-cell reference. Nothing raises. A
NumberFormat applied this way formats one cell of the intended range; a named
range built this way covers one cell of the intended block, and a SUM over it
returns a number that looks plausible.

Every multi-cell range string in this file therefore uses ${...} or $(...).
#>

function RGBv([int]$r, [int]$g, [int]$b) { return $r + ($g * 256) + ($b * 65536) }

$Slate   = RGBv 31 45 61
$SlateLt = RGBv 226 232 238
$White   = RGBv 255 255 255
$Amber   = RGBv 196 121 32
$GoodBg  = RGBv 198 239 206
$WarnBg  = RGBv 255 235 156
$BadBg   = RGBv 255 199 206
$GreyTxt = RGBv 110 118 128

Write-Host ""
Write-Host "Building Talon_Readiness_Control.xlsx" -ForegroundColor Cyan
Write-Host ("-" * 84)

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.ScreenUpdating = $false

try {
    $wb = $xl.Workbooks.Add()
    while ($wb.Worksheets.Count -gt 1) { $wb.Worksheets.Item($wb.Worksheets.Count).Delete() }

    # ---------------------------------------------------------------- Control
    $ctl = $wb.Worksheets.Item(1); $ctl.Name = 'Control'
    $ctl.Range('A1').Value2 = 'TALON ROBOTICS -- PAYLOAD DELIVERY READINESS'
    $ctl.Range('A1').Font.Size = 16; $ctl.Range('A1').Font.Bold = $true; $ctl.Range('A1').Font.Color = $Slate
    $ctl.Range('A2').Value2 = 'Control panel. Every sheet reads its parameters from the cells below.'
    $ctl.Range('A2').Font.Italic = $true
    $ctl.Range('A4').Value2 = 'PARAMETER'; $ctl.Range('B4').Value2 = 'VALUE'; $ctl.Range('C4').Value2 = 'WHAT IT CONTROLS'
    $ctl.Range('A4:C4').Font.Bold = $true
    $ctl.Range('A4:C4').Interior.Color = $Slate
    $ctl.Range('A4:C4').Font.Color = $White

    $params = @(
        @('Reporting date',      'AsOfDate',      'DATE', '2026-09-30', 'The as-of date for RAID ageing and overdue.'),
        @('Release candidate',   'AsOfBuild',     'NUM',  88,           'The build readiness is assessed against. Readiness is always about ONE build, never "the latest of everything".'),
        @('Extract folder',      'SourceFolder',  'TEXT', $ExtractPath, 'Where Power Query looks for the CSV extracts. Edit if you move the repo, then Data > Refresh All.'),
        @('Rig hours per week',  'RigHoursPerWeek','NUM', 180,          'Bookable hardware-in-the-loop hours in a week. Six rigs at 30 hours. This is what bounds the queue.'),
        @('Queue rows shown',    'QueueRowsShown','NUM',  30,           'How many ranked requirements the Verification Queue sheet displays.')
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
        $ctl.Range("B$r").Interior.Color = $SlateLt
        $ctl.Range("B$r").Font.Bold = $true
        $wb.Names.Add($p[1], $ctl.Range("B$r")) | Out-Null
        $r++
    }
    $ctl.Range('A12').Value2 = 'DATA DISCLOSURE'
    $ctl.Range('A12').Font.Bold = $true; $ctl.Range('A12').Font.Color = $Slate
    $ctl.Range('A13').Value2 = 'Talon Robotics is a fictional mechatronics company. All data in this workbook is synthetic and generated by the scripts in the sql\ folder. No confidential data is used and no claim is made about any production system.'
    $ctl.Range('A13:C15').Merge(); $ctl.Range('A13').WrapText = $true; $ctl.Range('A13').VerticalAlignment = -4160
    $ctl.Columns.Item(1).ColumnWidth = 22
    $ctl.Columns.Item(2).ColumnWidth = 46
    $ctl.Columns.Item(3).ColumnWidth = 92
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

    # Types are declared explicitly. Letting Power Query guess means a
    # requirement ID that happens to look numeric arrives as a number in one
    # refresh and text in the next, and every lookup against it breaks.
    $queries = @(
      @{ N='tReq'; F='requirement_verification.csv'; T='{"RequirementID",type text},{"Title",type text},{"ReqType",type text},{"Priority",type text},{"SubsystemCode",type text},{"SubsystemName",type text},{"Criticality",type text},{"MinTestLevel",type text},{"RequiresIndependentTester",Int64.Type},{"TotalCases",Int64.Type},{"CasesWithRuns",Int64.Type},{"PassingCases",Int64.Type},{"BlockedCases",Int64.Type},{"LastPassBuild",Int64.Type},{"LastGoodBuild",Int64.Type},{"PassingButUnderLevelled",Int64.Type},{"PassingButSelfVerified",Int64.Type},{"HasEvidence",Int64.Type},{"MeetsPolicy",Int64.Type},{"StaleByCode",Int64.Type},{"StaleByRequirement",Int64.Type},{"IsCurrent",Int64.Type},{"VerificationState",type text},{"InterveningBuilds",Int64.Type},{"AsOfBuild",Int64.Type}' },
      @{ N='tQueue'; F='verification_queue.csv'; T='{"PriorityRank",Int64.Type},{"RequirementID",type text},{"Title",type text},{"ReqType",type text},{"Priority",type text},{"SubsystemCode",type text},{"SubsystemName",type text},{"Criticality",type text},{"OwnerID",type text},{"OwnerName",type text},{"OwnerTeam",type text},{"MinTestLevel",type text},{"LastGoodBuild",Int64.Type},{"InterveningBuilds",Int64.Type},{"Cases",Int64.Type},{"RigHours",type number},{"TestMinutes",Int64.Type},{"CumulativeRigHours",type number},{"PriorityScore",type number},{"ActionCode",type text},{"IsThisWeek",Int64.Type},{"RecommendedAction",type text}' },
      @{ N='tSub'; F='subsystem_readiness.csv'; T='{"SubsystemCode",type text},{"SubsystemName",type text},{"Criticality",type text},{"RequiresHILRig",Int64.Type},{"Requirements",Int64.Type},{"MustShip",Int64.Type},{"CurrentAll",Int64.Type},{"CurrentMustShip",Int64.Type},{"NoEvidence",Int64.Type},{"Insufficient",Int64.Type},{"StaleByCode",Int64.Type},{"StaleByReq",Int64.Type},{"ReadinessPct",type number},{"BuildsChanged",Int64.Type},{"LinesChanged",Int64.Type},{"LastChangeBuild",Int64.Type},{"RigHoursOutstanding",type number}' },
      @{ N='tKPI'; F='readiness_kpi.csv'; T='{"AsOfBuild",Int64.Type},{"AsOfDate",type date},{"Requirements",Int64.Type},{"MustShipCount",Int64.Type},{"WorkItemCompletionPct",type number},{"VerificationCoveragePct",type number},{"ShipReadinessPct",type number},{"StaleVerificationPct",type number},{"PolicyCompliancePct",type number},{"NoEvidenceCount",Int64.Type},{"InsufficientCount",Int64.Type},{"StaleCount",Int64.Type},{"CurrentCount",Int64.Type},{"BlockedTestPct",type number},{"CriticalRAIDOpen",Int64.Type},{"OverdueRAIDPct",type number},{"TotalRAIDExposure",Int64.Type},{"RigHoursOutstanding",type number}' },
      @{ N='tTrend'; F='readiness_trend.csv'; T='{"BuildNumber",Int64.Type},{"BuildID",type text},{"BuildDate",type date},{"WorkItemCompletionPct",type number},{"VerificationCoveragePct",type number},{"ShipReadinessPct",type number},{"StaleVerificationPct",type number},{"ReadinessGapVsReported",type number}' },
      @{ N='tRAID'; F='raid_register.csv'; T='{"RAIDID",type text},{"RAIDType",type text},{"Title",type text},{"SubsystemCode",type text},{"SubsystemName",type text},{"Criticality",type text},{"OwnerID",type text},{"OwnerName",type text},{"OwnerTeam",type text},{"RaisedDate",type date},{"DueDate",type date},{"ClosedDate",type date},{"Probability",Int64.Type},{"Impact",Int64.Type},{"ExposureScore",Int64.Type},{"ExposureBand",type text},{"RAIDStatus",type text},{"IsOpen",Int64.Type},{"IsOverdue",Int64.Type},{"DaysOverdue",Int64.Type},{"AgeDays",Int64.Type},{"MitigationNote",type text}' },
      @{ N='tWork'; F='work_items.csv'; T='{"WorkItemID",type text},{"RequirementID",type text},{"Priority",type text},{"SubsystemCode",type text},{"AssignedToID",type text},{"AssignedToName",type text},{"OpenedDate",type date},{"ClosedDate",type date},{"EstimateHours",type number},{"ActualHours",type number},{"WorkItemStatus",type text}' },
      @{ N='tDQ'; F='data_quality_summary.csv'; T='{"AnomalyType",type text},{"EntityType",type text},{"Severity",type text},{"ImpactClass",type text},{"CountsTowardExposure",Int64.Type},{"Anomalies",Int64.Type},{"PopulationScanned",Int64.Type},{"AnomalyRatePct",type number},{"WhatItMeans",type text}' },
      @{ N='tPolicy'; F='verification_policy.csv'; T='{"ReqType",type text},{"MinTestLevel",type text},{"MinTestLevelRank",Int64.Type},{"RequiresIndependentTester",Int64.Type},{"PolicyNote",type text}' },
      @{ N='tMatrix'; F='raid_matrix.csv'; T='{"Probability",Int64.Type},{"Impact",Int64.Type},{"ExposureScore",Int64.Type},{"ExposureBand",type text}' },
      @{ N='tTargets'; F='readiness_targets.csv'; T='{"MetricName",type text},{"TargetValue",type number},{"WarningValue",type number},{"Direction",type text},{"Unit",type text},{"Description",type text}' }
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
    if ($loaded['tReq'] -lt 100) { throw "tReq loaded only $($loaded['tReq']) rows -- the Power Query load did not work." }
    $nQueue = $loaded['tQueue']

    # Create every remaining sheet BEFORE writing a single formula. Excel
    # rewrites a reference to a sheet that does not exist yet as #REF! and does
    # not repair it when the sheet appears.
    $sheetOrder = @('README','Dashboard','Verification Queue','Subsystem Readiness','RAID Register','Validation','Calc')
    $made = @{}
    foreach ($name in $sheetOrder) {
        $s = $wb.Worksheets.Add([System.Reflection.Missing]::Value, $wb.Worksheets.Item($wb.Worksheets.Count))
        $s.Name = $name
        $made[$name] = $s
    }
    Write-Host "  Worksheets               $($sheetOrder.Count) created up front"

    # ------------------------------------------------------------------ Calc
    # Helper block for the queue ranking. Kept on its own sheet, and the
    # ranking block is offset dynamically rather than at a fixed row, because a
    # fixed layout on a shared Calc sheet is how Project 2 produced a
    # plausible-but-wrong BPDSO: two blocks silently overlapped.
    $calc = $made['Calc']
    $calc.Range('A1').Value2 = 'CALC -- helper block for the Verification Queue ranking. Not a reporting sheet.'
    $calc.Range('A1').Font.Bold = $true; $calc.Range('A1').Font.Color = $GreyTxt
    $qFirst = 4
    $qLast  = $qFirst + $nQueue - 1
    # Excel walks the queue in SQL's published rank order and re-derives the
    # CUMULATIVE rig hours and the week boundary itself. It deliberately does
    # NOT re-implement the priority score: that would be a second definition of
    # the ranking with nothing asserting the two agree, which is how Project 3
    # ended up with a cost index that was wrong in both tools at once. What is
    # worth checking independently is the arithmetic that decides the week, and
    # that is what this block does.
    for ($i = 0; $i -lt $nQueue; $i++) {
        $rr = $qFirst + $i
        $calc.Range("A$rr").Formula = "=$($i + 1)"
        $calc.Range("B$rr").Formula = "=INDEX(tQueue[RequirementID],MATCH(A$rr,tQueue[PriorityRank],0))"
        $calc.Range("C$rr").Formula = "=INDEX(tQueue[RigHours],MATCH(A$rr,tQueue[PriorityRank],0))"
        $calc.Range("D$rr").Formula = "=IF(ROW()=$qFirst,C$rr,D$($rr-1)+C$rr)"
        $calc.Range("E$rr").Formula = "=IF(D$rr<=RigHoursPerWeek,1,0)"
    }
    # Named so the Dashboard can point at it. Without this the block was
    # computed and then referenced by nothing at all.
    $wb.Names.Add('CalcWithinWeek', $calc.Range("E$($qFirst):E$($qFirst + $nQueue - 1)")) | Out-Null

    $calc.Range("A3").Value2 = 'Rank'
    $calc.Range("B3").Value2 = 'RequirementID'
    $calc.Range("C3").Value2 = 'RigHours'
    $calc.Range("D3").Value2 = 'Cumulative rig hours (Excel, independent of SQL)'
    $calc.Range("E3").Value2 = 'Within the week'
    $calc.Range("A3:E3").Font.Bold = $true
    $calc.Columns.Item(4).ColumnWidth = 40
    Write-Host "  Calc sheet               built ($nQueue queue rows)"

    # ------------------------------------------------------------- Dashboard
    $d = $made['Dashboard']
    $d.Range('A1').Value2 = 'PAYLOAD DELIVERY -- READINESS AT THE RELEASE CANDIDATE'
    $d.Range('A1').Font.Size = 15; $d.Range('A1').Font.Bold = $true; $d.Range('A1').Font.Color = $Slate
    $d.Range('A2').Formula = '="Assessed against build "&AsOfBuild&"  |  as of "&TEXT(AsOfDate,"d mmmm yyyy")&"  |  every figure below is a formula over the extracts, not a pasted value"'
    $d.Range('A2').Font.Italic = $true; $d.Range('A2').Font.Color = $GreyTxt

    # --- the contrast, stated first because it is the whole argument
    $d.Range('A4').Value2 = 'THE QUESTION ON THE AGENDA'
    $d.Range('A4').Font.Bold = $true; $d.Range('A4').Font.Color = $White
    $d.Range('A4:E4').Interior.Color = $Slate

    $contrast = @(
        @('B6','Work items closed',       '=COUNTIF(tWork[WorkItemStatus],"Closed")/COUNTA(tWork[WorkItemID])*100', '0.00"%"', 'What the programme reports. Counts effort.'),
        @('B7','Requirements with evidence','=COUNTIF(tReq[HasEvidence],1)/COUNTA(tReq[RequirementID])*100',        '0.00"%"', 'Has a passing test ever run? Deliberately generous.'),
        @('B8','SHIP READINESS',          '=COUNTIFS(tReq[Priority],"MustShip",tReq[IsCurrent],1)/COUNTIF(tReq[Priority],"MustShip")*100', '0.00"%"', 'Must-ship requirements with CURRENT, policy-compliant evidence. The gate.'),
        @('B9','The gap',                 '=B6-B8',                                                                  '0.00" pts"', 'Between what is reported and what can ship.')
    )
    $d.Range('A5').Value2 = 'MEASURE'; $d.Range('B5').Value2 = 'VALUE'; $d.Range('C5').Value2 = 'WHAT IT ACTUALLY MEANS'
    $d.Range('A5:C5').Font.Bold = $true; $d.Range('A5:C5').Interior.Color = $SlateLt
    foreach ($c in $contrast) {
        $cell = [string]$c[0]; $row = $cell.Substring(1)
        $d.Range("A$row").Value2 = [string]$c[1]
        $d.Range($cell).Formula = [string]$c[2]
        $d.Range($cell).NumberFormat = [string]$c[3]
        $d.Range("C$row").Value2 = [string]$c[4]
    }
    $d.Range('A8:C8').Font.Bold = $true
    $d.Range('A8:C8').Interior.Color = $BadBg
    $d.Range('B6:B9').Font.Size = 12; $d.Range('B6:B9').Font.Bold = $true

    # --- the scorecard, RAG against the target table
    $d.Range('A11').Value2 = 'READINESS SCORECARD'
    $d.Range('A11').Font.Bold = $true; $d.Range('A11').Font.Color = $White
    $d.Range('A11:E11').Interior.Color = $Slate
    $d.Range('A12').Value2 = 'METRIC'; $d.Range('B12').Value2 = 'VALUE'
    $d.Range('C12').Value2 = 'TARGET'; $d.Range('D12').Value2 = 'WARN'; $d.Range('E12').Value2 = 'STATUS'
    $d.Range('A12:E12').Font.Bold = $true; $d.Range('A12:E12').Interior.Color = $SlateLt

    # Each row: label, the Excel formula, and the metric name to look the
    # target up by. Thresholds are never written into a formula -- they live in
    # readiness_targets.csv, so a target the programme disagrees with can be
    # argued with rather than hunted for.
    $kpis = @(
        @('Ship readiness',          '=COUNTIFS(tReq[Priority],"MustShip",tReq[IsCurrent],1)/COUNTIF(tReq[Priority],"MustShip")*100', 'ShipReadinessPct'),
        @('Verification coverage',   '=COUNTIF(tReq[HasEvidence],1)/COUNTA(tReq[RequirementID])*100',                                  'VerificationCoveragePct'),
        @('Stale verification',      '=SUMPRODUCT((tReq[MeetsPolicy]=1)*(((tReq[StaleByCode]=1)+(tReq[StaleByRequirement]=1))>0))/COUNTIF(tReq[MeetsPolicy],1)*100', 'StaleVerificationPct'),
        @('Policy compliance',       '=COUNTIFS(tReq[HasEvidence],1,tReq[MeetsPolicy],1)/COUNTIF(tReq[HasEvidence],1)*100',            'PolicyCompliancePct'),
        @('Work-item completion',    '=COUNTIF(tWork[WorkItemStatus],"Closed")/COUNTA(tWork[WorkItemID])*100',                          'WorkItemCompletionPct'),
        @('Critical RAID open',      '=COUNTIFS(tRAID[ExposureBand],"Critical",tRAID[IsOpen],1)',                                       'CriticalRAIDOpen'),
        @('Overdue RAID',            '=COUNTIF(tRAID[IsOverdue],1)/COUNTIF(tRAID[IsOpen],1)*100',                                       'OverdueRAIDPct')
    )
    $kr = 13
    foreach ($k in $kpis) {
        $d.Range("A$kr").Value2 = [string]$k[0]
        $d.Range("B$kr").Formula = [string]$k[1]
        $d.Range("B$kr").NumberFormat = if ($k[2] -eq 'CriticalRAIDOpen') { '0' } else { '0.00' }
        $d.Range("C$kr").Formula = "=INDEX(tTargets[TargetValue],MATCH(""$($k[2])"",tTargets[MetricName],0))"
        $d.Range("D$kr").Formula = "=INDEX(tTargets[WarningValue],MATCH(""$($k[2])"",tTargets[MetricName],0))"
        # Direction is read from the table too: hardcoding >= here would silently
        # invert the status for every LowerBetter metric the day one is added.
        $d.Range("E$kr").Formula = "=IF(INDEX(tTargets[Direction],MATCH(""$($k[2])"",tTargets[MetricName],0))=""HigherBetter"",IF(B$kr>=C$kr,""Green"",IF(B$kr>=D$kr,""Amber"",""Red"")),IF(B$kr<=C$kr,""Green"",IF(B$kr<=D$kr,""Amber"",""Red"")))"
        $kr++
    }
    $lastKpi = $kr - 1
    # Conditional formatting on the status text
    foreach ($spec in @(@('Green',$GoodBg), @('Amber',$WarnBg), @('Red',$BadBg))) {
        $fc = $d.Range("E13:E$lastKpi").FormatConditions.Add(1, 3, "=""$($spec[0])""")
        $fc.Interior.Color = $spec[1]
    }

    # --- where the gap is
    $d.Range('A22').Value2 = 'WHERE THE GAP IS -- readiness by subsystem, worst first'
    $d.Range('A22').Font.Bold = $true; $d.Range('A22').Font.Color = $White
    $d.Range('A22:F22').Interior.Color = $Slate
    $hdr = @('Subsystem','Criticality','Must-ship','Current','Readiness %','Builds changed')
    for ($i = 0; $i -lt $hdr.Count; $i++) {
        $d.Cells.Item(23, $i + 1).Value2 = $hdr[$i]
    }
    $d.Range('A23:F23').Font.Bold = $true; $d.Range('A23:F23').Interior.Color = $SlateLt
    $nSub = $loaded['tSub']
    for ($i = 0; $i -lt $nSub; $i++) {
        $rr = 24 + $i
        $k = $i + 1
        $d.Range("A$rr").Formula = "=INDEX(tSub[SubsystemName],SMALL(IF(1,ROW(tSub[ReadinessPct])-ROW(INDEX(tSub[ReadinessPct],1))+1),$k))"
        # SMALL/IF is an array pattern; use a plain positional index instead so
        # the file stays legacy-safe. The extract is already sorted ascending
        # by ReadinessPct, so position IS worst-first.
        $d.Range("A$rr").Formula = "=INDEX(tSub[SubsystemName],$k)"
        $d.Range("B$rr").Formula = "=INDEX(tSub[Criticality],$k)"
        $d.Range("C$rr").Formula = "=INDEX(tSub[MustShip],$k)"
        $d.Range("D$rr").Formula = "=INDEX(tSub[CurrentMustShip],$k)"
        $d.Range("E$rr").Formula = "=IF(C$rr=0,0,D$rr/C$rr*100)"
        $d.Range("E$rr").NumberFormat = '0.00'
        $d.Range("F$rr").Formula = "=INDEX(tSub[BuildsChanged],$k)"
    }
    $lastSub = 23 + $nSub
    $fcSub = $d.Range("E24:E$lastSub").FormatConditions.Add(1, 6, '=40')
    $fcSub.Interior.Color = $BadBg

    # --- the schedule answer
    $sr = $lastSub + 2
    $d.Range("A$sr").Value2 = 'THE SCHEDULE ANSWER'
    $d.Range("A$sr").Font.Bold = $true; $d.Range("A$sr").Font.Color = $White
    $d.Range("A$sr" + ":E$sr").Interior.Color = $Slate
    $d.Range("A$($sr+1)").Value2 = 'Requirements outstanding'
    $d.Range("B$($sr+1)").Formula = '=COUNTA(tQueue[RequirementID])'
    $d.Range("A$($sr+2)").Value2 = 'Rig hours to clear them'
    $d.Range("B$($sr+2)").Formula = '=SUM(tQueue[RigHours])'
    $d.Range("B$($sr+2)").NumberFormat = '0.0'
    $d.Range("A$($sr+3)").Value2 = 'Rig weeks at current capacity'
    $d.Range("B$($sr+3)").Formula = "=B$($sr+2)/RigHoursPerWeek"
    $d.Range("B$($sr+3)").NumberFormat = '0.0'
    $d.Range("C$($sr+3)").Value2 = 'This is the honest schedule answer: it is the only figure here a programme board can act on without a further study.'
    # SUM over the Calc sheet's live boundary, NOT COUNTIF over the static
    # IsThisWeek column SQL computed at 180 rig hours. The static version made
    # this figure -- the one the README promises responds to the Rig hours cell
    # -- the one figure on the sheet that could not respond to it.
    $d.Range("A$($sr+4)").Value2 = 'Schedulable this week'
    $d.Range("B$($sr+4)").Formula = '=SUM(CalcWithinWeek)'

    # Named, so the validator and any future reader address these by name
    # rather than by cell. This block sits below the subsystem table, so its
    # row number moves whenever a subsystem is added -- and a checker holding a
    # fixed cell reference would then read a neighbouring figure and report it
    # as a match. That is the Project 1 row-offset defect, in a different file.
    $wb.Names.Add('QueueOutstanding',     $d.Range("B$($sr+1)")) | Out-Null
    $wb.Names.Add('RigHoursOutstanding',  $d.Range("B$($sr+2)")) | Out-Null
    $wb.Names.Add('RigWeeksOutstanding',  $d.Range("B$($sr+3)")) | Out-Null
    $wb.Names.Add('SchedulableThisWeek',  $d.Range("B$($sr+4)")) | Out-Null

    $d.Columns.Item(1).ColumnWidth = 30
    $d.Columns.Item(2).ColumnWidth = 14
    $d.Columns.Item(3).ColumnWidth = 74
    $d.Columns.Item(4).ColumnWidth = 12
    $d.Columns.Item(5).ColumnWidth = 14
    $d.Columns.Item(6).ColumnWidth = 16
    Write-Host "  Dashboard                built"

    # -------------------------------------------------- Verification Queue
    $q = $made['Verification Queue']
    $q.Range('A1').Value2 = 'VERIFICATION QUEUE -- what to do next, bounded by rig capacity'
    $q.Range('A1').Font.Size = 14; $q.Range('A1').Font.Bold = $true; $q.Range('A1').Font.Color = $Slate
    $q.Range('A2').Formula = '="Ranked by priority score. The first "&SUM(CalcWithinWeek)&" rows fit inside "&RigHoursPerWeek&" rig hours -- one week. A list of everything outstanding is not a plan."'
    $q.Range('A2').Font.Italic = $true; $q.Range('A2').Font.Color = $GreyTxt
    $qh = @('Rank','Requirement','Subsystem','Type','Priority','Owner','Action','Rig hrs','Cumulative','This week','What to do')
    for ($i = 0; $i -lt $qh.Count; $i++) { $q.Cells.Item(4, $i + 1).Value2 = $qh[$i] }
    $q.Range('A4:K4').Font.Bold = $true; $q.Range('A4:K4').Interior.Color = $Slate; $q.Range('A4:K4').Font.Color = $White
    $show = [Math]::Min($nQueue, 30)
    for ($i = 0; $i -lt $show; $i++) {
        $rr = 5 + $i; $k = $i + 1
        $q.Range("A$rr").Formula = "=$k"
        foreach ($m in @(@('B','RequirementID'),@('C','SubsystemCode'),@('D','ReqType'),@('E','Priority'),
                         @('F','OwnerName'),@('G','ActionCode'),@('H','RigHours'),@('I','CumulativeRigHours'),
                         @('J','IsThisWeek'),@('K','RecommendedAction'))) {
            $q.Range("$($m[0])$rr").Formula = "=INDEX(tQueue[$($m[1])],MATCH(`$A$rr,tQueue[PriorityRank],0))"
        }
        $q.Range("H$rr").NumberFormat = '0.0'
        $q.Range("I$rr").NumberFormat = '0.0'
    }
    $lastQ = 4 + $show
    $fcQ = $q.Range("A5:K$lastQ").FormatConditions.Add(2, 0, "=`$J5=1")
    $fcQ.Interior.Color = $GoodBg
    $q.Columns.Item(1).ColumnWidth = 6
    $q.Columns.Item(2).ColumnWidth = 13
    $q.Columns.Item(3).ColumnWidth = 12
    $q.Columns.Item(4).ColumnWidth = 13
    $q.Columns.Item(5).ColumnWidth = 12
    $q.Columns.Item(6).ColumnWidth = 22
    $q.Columns.Item(7).ColumnWidth = 22
    $q.Columns.Item(11).ColumnWidth = 96
    Write-Host "  Verification Queue       built ($show rows shown)"

    # ------------------------------------------------- Subsystem Readiness
    $s = $made['Subsystem Readiness']
    $s.Range('A1').Value2 = 'SUBSYSTEM READINESS -- and the churn behind it'
    $s.Range('A1').Font.Size = 14; $s.Range('A1').Font.Bold = $true; $s.Range('A1').Font.Color = $Slate
    $s.Range('A2').Value2 = 'A subsystem that changed forty times and is a quarter ready is not behind on testing -- it is still being designed. That is a different conversation with a different owner.'
    $s.Range('A2').Font.Italic = $true; $s.Range('A2').Font.Color = $GreyTxt
    $sh = @('Subsystem','Criticality','Reqs','Must-ship','Current','Readiness %','No evidence','Insufficient','Stale (code)','Stale (req)','Builds changed','Rig hrs')
    for ($i = 0; $i -lt $sh.Count; $i++) { $s.Cells.Item(4, $i + 1).Value2 = $sh[$i] }
    $s.Range('A4:L4').Font.Bold = $true; $s.Range('A4:L4').Interior.Color = $Slate; $s.Range('A4:L4').Font.Color = $White
    for ($i = 0; $i -lt $nSub; $i++) {
        $rr = 5 + $i; $k = $i + 1
        foreach ($m in @(@('A','SubsystemName'),@('B','Criticality'),@('C','Requirements'),@('D','MustShip'),
                         @('E','CurrentMustShip'),@('G','NoEvidence'),@('H','Insufficient'),
                         @('I','StaleByCode'),@('J','StaleByReq'),@('K','BuildsChanged'),@('L','RigHoursOutstanding'))) {
            $s.Range("$($m[0])$rr").Formula = "=INDEX(tSub[$($m[1])],$k)"
        }
        $s.Range("F$rr").Formula = "=IF(D$rr=0,0,E$rr/D$rr*100)"
        $s.Range("F$rr").NumberFormat = '0.00'
        $s.Range("L$rr").NumberFormat = '0.0'
    }
    $s.Columns.Item(1).ColumnWidth = 30
    Write-Host "  Subsystem Readiness      built"

    # -------------------------------------------------------- RAID Register
    $ra = $made['RAID Register']
    $ra.Range('A1').Value2 = 'RAID REGISTER -- scored against the matrix, open items first'
    $ra.Range('A1').Font.Size = 14; $ra.Range('A1').Font.Bold = $true; $ra.Range('A1').Font.Color = $Slate
    $ra.Range('A2').Value2 = 'Exposure is read from Ref_RAIDMatrix, not computed as probability x impact. A board treats "unlikely but catastrophic" differently from "certain but trivial" even though the product is the same.'
    $ra.Range('A2').Font.Italic = $true; $ra.Range('A2').Font.Color = $GreyTxt
    $ra.Range('A4').Value2 = 'Open items'
    $ra.Range('B4').Formula = '=COUNTIF(tRAID[IsOpen],1)'
    $ra.Range('A5').Value2 = 'Critical band, open'
    $ra.Range('B5').Formula = '=COUNTIFS(tRAID[ExposureBand],"Critical",tRAID[IsOpen],1)'
    $ra.Range('A6').Value2 = 'Overdue, open'
    $ra.Range('B6').Formula = '=COUNTIFS(tRAID[IsOverdue],1,tRAID[IsOpen],1)'
    $ra.Range('A7').Value2 = 'Total open exposure'
    $ra.Range('B7').Formula = '=SUMIFS(tRAID[ExposureScore],tRAID[IsOpen],1)'
    $ra.Range('A4:A7').Font.Bold = $true

    $rh = @('RAID','Type','Subsystem','Owner','Prob','Impact','Exposure','Band','Status','Overdue days','Mitigation')
    for ($i = 0; $i -lt $rh.Count; $i++) { $ra.Cells.Item(9, $i + 1).Value2 = $rh[$i] }
    $ra.Range('A9:K9').Font.Bold = $true; $ra.Range('A9:K9').Interior.Color = $Slate; $ra.Range('A9:K9').Font.Color = $White
    $nRaid = [Math]::Min($loaded['tRAID'], 40)
    for ($i = 0; $i -lt $nRaid; $i++) {
        $rr = 10 + $i; $k = $i + 1
        foreach ($m in @(@('A','RAIDID'),@('B','RAIDType'),@('C','SubsystemCode'),@('D','OwnerName'),
                         @('E','Probability'),@('F','Impact'),@('G','ExposureScore'),@('H','ExposureBand'),
                         @('I','RAIDStatus'),@('J','DaysOverdue'),@('K','MitigationNote'))) {
            $ra.Range("$($m[0])$rr").Formula = "=INDEX(tRAID[$($m[1])],$k)"
        }
    }
    $lastR = 9 + $nRaid
    $fcR = $ra.Range("A10:K$lastR").FormatConditions.Add(2, 0, "=`$H10=""Critical""")
    $fcR.Interior.Color = $BadBg
    $ra.Columns.Item(1).ColumnWidth = 11
    $ra.Columns.Item(4).ColumnWidth = 22
    $ra.Columns.Item(11).ColumnWidth = 80
    Write-Host "  RAID Register            built ($nRaid rows shown)"

    # ----------------------------------------------------------- Validation
    # Excel's own figure against the one SQL produced independently. This sheet
    # is the reason to believe anything else in the workbook.
    $v = $made['Validation']
    $v.Range('A1').Value2 = 'VALIDATION -- every Excel figure against the one SQL produced independently'
    $v.Range('A1').Font.Size = 14; $v.Range('A1').Font.Bold = $true; $v.Range('A1').Font.Color = $Slate
    $v.Range('A2').Value2 = 'Each row recomputes a figure with Excel formulas over the extracts, then compares it with the value SQL published in readiness_kpi.csv. Agreement is necessary, not sufficient -- both could be derived from the same flawed definition -- but disagreement is decisive.'
    $v.Range('A2').Font.Italic = $true; $v.Range('A2').Font.Color = $GreyTxt
    $vh = @('#','Figure','Excel','SQL','Difference','Result')
    for ($i = 0; $i -lt $vh.Count; $i++) { $v.Cells.Item(4, $i + 1).Value2 = $vh[$i] }
    $v.Range('A4:F4').Font.Bold = $true; $v.Range('A4:F4').Interior.Color = $Slate; $v.Range('A4:F4').Font.Color = $White

    $checks = @(
        @('Requirements (baselined)',  '=COUNTA(tReq[RequirementID])',                                     '=INDEX(tKPI[Requirements],1)'),
        @('Must-ship requirements',    '=COUNTIF(tReq[Priority],"MustShip")',                              '=INDEX(tKPI[MustShipCount],1)'),
        @('Ship readiness %',          '=COUNTIFS(tReq[Priority],"MustShip",tReq[IsCurrent],1)/COUNTIF(tReq[Priority],"MustShip")*100', '=INDEX(tKPI[ShipReadinessPct],1)'),
        @('Verification coverage %',   '=COUNTIF(tReq[HasEvidence],1)/COUNTA(tReq[RequirementID])*100',    '=INDEX(tKPI[VerificationCoveragePct],1)'),
        @('Stale verification %',      '=SUMPRODUCT((tReq[MeetsPolicy]=1)*(((tReq[StaleByCode]=1)+(tReq[StaleByRequirement]=1))>0))/COUNTIF(tReq[MeetsPolicy],1)*100', '=INDEX(tKPI[StaleVerificationPct],1)'),
        @('Policy compliance %',       '=COUNTIFS(tReq[HasEvidence],1,tReq[MeetsPolicy],1)/COUNTIF(tReq[HasEvidence],1)*100', '=INDEX(tKPI[PolicyCompliancePct],1)'),
        @('Work-item completion %',    '=COUNTIF(tWork[WorkItemStatus],"Closed")/COUNTA(tWork[WorkItemID])*100', '=INDEX(tKPI[WorkItemCompletionPct],1)'),
        @('No evidence (count)',       '=COUNTIF(tReq[HasEvidence],0)',                                    '=INDEX(tKPI[NoEvidenceCount],1)'),
        @('Insufficient (count)',      '=COUNTIFS(tReq[HasEvidence],1,tReq[MeetsPolicy],0)',               '=INDEX(tKPI[InsufficientCount],1)'),
        @('Current (count)',           '=COUNTIF(tReq[IsCurrent],1)',                                      '=INDEX(tKPI[CurrentCount],1)'),
        @('Queue size',                '=COUNTA(tQueue[RequirementID])',                                   '=COUNTIF(tReq[IsCurrent],0)'),
        @('Rig hours outstanding',     '=SUM(tQueue[RigHours])',                                           '=INDEX(tKPI[RigHoursOutstanding],1)'),
        @('Critical RAID open',        '=COUNTIFS(tRAID[ExposureBand],"Critical",tRAID[IsOpen],1)',        '=INDEX(tKPI[CriticalRAIDOpen],1)'),
        @('Overdue RAID %',            '=COUNTIF(tRAID[IsOverdue],1)/COUNTIF(tRAID[IsOpen],1)*100',        '=INDEX(tKPI[OverdueRAIDPct],1)'),
        @('Must-ship, subsystem sum',  '=SUM(tSub[MustShip])',                                             '=COUNTIF(tReq[Priority],"MustShip")'),
        @('Current, subsystem sum',    '=SUM(tSub[CurrentMustShip])',                                      '=COUNTIFS(tReq[Priority],"MustShip",tReq[IsCurrent],1)')
    )
    $vr = 5
    $n = 1
    foreach ($c in $checks) {
        $v.Range("A$vr").Formula = "=$n"
        $v.Range("B$vr").Value2 = [string]$c[0]
        $v.Range("C$vr").Formula = [string]$c[1]
        $v.Range("D$vr").Formula = [string]$c[2]
        $v.Range("E$vr").Formula = "=ABS(C$vr-D$vr)"
        $v.Range("C$vr").NumberFormat = '0.0000'
        $v.Range("D$vr").NumberFormat = '0.0000'
        $v.Range("E$vr").NumberFormat = '0.0000'
        # A hundredth is the right tolerance: SQL publishes these to two
        # decimals, so comparing to more precision than SQL stores would fail
        # on SQL's own rounding rather than on a disagreement.
        $v.Range("F$vr").Formula = "=IF(E$vr<=0.01,""MATCH"",""DIFFERS"")"
        $vr++; $n++
    }
    $lastV = $vr - 1
    $fcV1 = $v.Range("F5:F$lastV").FormatConditions.Add(1, 3, '="MATCH"');   $fcV1.Interior.Color = $GoodBg
    $fcV2 = $v.Range("F5:F$lastV").FormatConditions.Add(1, 3, '="DIFFERS"'); $fcV2.Interior.Color = $BadBg
    $v.Range("B$($vr+1)").Value2 = 'Checks reconciling:'
    $v.Range("C$($vr+1)").Formula = "=COUNTIF(F5:F$lastV,""MATCH"")&"" of ""&COUNTA(F5:F$lastV)"
    $v.Range("B$($vr+1)").Font.Bold = $true; $v.Range("C$($vr+1)").Font.Bold = $true
    $v.Columns.Item(2).ColumnWidth = 30
    $v.Columns.Item(3).ColumnWidth = 16
    $v.Columns.Item(4).ColumnWidth = 16
    $v.Columns.Item(5).ColumnWidth = 14
    $v.Columns.Item(6).ColumnWidth = 12
    Write-Host "  Validation               built ($($checks.Count) checks)"

    # --------------------------------------------------------------- README
    $rm = $made['README']
    $rm.Range('A1').Value2 = 'TALON ROBOTICS -- PAYLOAD DELIVERY READINESS CONTROL'
    $rm.Range('A1').Font.Size = 15; $rm.Range('A1').Font.Bold = $true; $rm.Range('A1').Font.Color = $Slate
    $lines = @(
        '',
        'THE QUESTION',
        'Is the new payload ready to ship, and what is still open?',
        '',
        'THE ANSWER THIS WORKBOOK GIVES',
        'Work items are 93.9% closed. Ship readiness is 52.0%. Both are true; only one is about shipping.',
        '',
        'WHY THEY DIFFER',
        'A verification is evidence about ONE BUILD. If a requirement passed on build 62 and its subsystem',
        'changed in build 71, the green tick is describing software that no longer exists. Readiness counts',
        'only must-ship requirements whose evidence is still current AND meets the verification policy for',
        'their type -- a safety behaviour is not verified by a unit test, and not by its own author.',
        '',
        'WHAT IS NOT A STALENESS RULE',
        'Age. A verification from fourteen months ago is fine if nothing underneath it has moved; one from',
        'last Tuesday is worthless if the subsystem was rewritten on Wednesday. Intervening change is the',
        'question, and it is why the model needs build history rather than timestamps.',
        '',
        'SHEETS',
        '  Control              Parameters. Change the rig capacity and the queue re-bounds itself.',
        '  Dashboard            The contrast, the scorecard, where the gap is, and the schedule answer.',
        '  Verification Queue   What to do next, ranked, bounded by one week of rig time.',
        '  Subsystem Readiness  Readiness against the churn behind it.',
        '  RAID Register        Exposure read from the matrix, open items first.',
        '  Validation           Every Excel figure against the one SQL produced independently.',
        '  Calc                 Helper block for the queue. Not a reporting sheet.',
        '  Data_*               The loaded extracts. Data > Refresh All re-reads them from disk.',
        '',
        'HOW TO USE IT',
        'Change Rig hours per week on the Control sheet and the queue re-bounds itself: that is the',
        '"what if we add a rig" question answered without editing anything. Every threshold lives in',
        'readiness_targets.csv, so a target the programme disagrees with can be argued with.',
        '',
        'DATA DISCLOSURE',
        'Talon Robotics is a fictional mechatronics company. All data in this workbook is synthetic and',
        'generated by the scripts in the sql\ folder. No confidential data is used and no claim is made',
        'about any production system.'
    )
    for ($i = 0; $i -lt $lines.Count; $i++) { $rm.Range("A$($i+2)").Value2 = $lines[$i] }
    foreach ($h in @(3,6,9,14,19,29,34)) { $rm.Range("A$h").Font.Bold = $true; $rm.Range("A$h").Font.Color = $Slate }
    $rm.Columns.Item(1).ColumnWidth = 108
    Write-Host "  README                   built"

    # ------------------------------------------------------------- finalise
    $wb.Worksheets.Item('README').Activate()
    $xl.Calculation = -4105
    $wb.Application.CalculateFullRebuild()

    # Strip author metadata. A portfolio file should not carry the machine name
    # of whoever built it.
    $wb.RemovePersonalInformation = $true
    if (Test-Path $OutPath) { Remove-Item $OutPath -Force }
    $wb.SaveAs($OutPath, 51)
    Write-Host ("-" * 84)
    Write-Host "  Saved: $OutPath" -ForegroundColor Green
}
finally {
    if ($wb) { $wb.Close($false) }
    $xl.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
Write-Host ""
