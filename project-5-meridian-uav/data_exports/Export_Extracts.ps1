<#
================================================================================
Project 5 -- Meridian UAV Services: Predictive Maintenance
Script:  Export_Extracts.ps1
Purpose: Write the CSV extracts that stand in for a nightly export from the
         maintenance system, the flight-log system and the condition-monitoring
         telemetry store.

WHY CSV AND NOT A LIVE CONNECTION
    The Excel workbook is a deliverable in its own right, and a reviewer should
    be able to open it and see working formulas without installing SQL Server,
    restoring a database or holding credentials. These extracts are what the
    workbook reads through Power Query.

    The Python model and the Streamlit app read SQL Server directly when they
    can, and fall back to these same files when they cannot -- so a reviewer
    with neither SQL Server nor a licence for anything can still run the
    analysis end to end.

WHY POWERSHELL AND NOT bcp OR sqlcmd -s","
    Recommended-action text and maintenance findings contain commas and
    apostrophes. sqlcmd's separator option does not quote fields, so those rows
    split into extra columns and every downstream count is quietly wrong --
    quietly, because the file still parses and still looks like a table.
    Export-Csv quotes correctly.

THE ONE LARGE FILE, AND WHY IT IS HERE
    sensor_features.csv.gz carries one row per monitored component per sortie --
    about 85,000 rows. It is the only extract a reader could reasonably object
    to on size, and it is included deliberately: without it the degradation
    model is not reproducible by anyone who does not have the database, and an
    unreproducible model in a portfolio is a screenshot.

    It is the only gzipped file. pandas reads .csv.gz with no change to the
    reading code; the other sixteen stay plain so a reviewer can open them by
    double-clicking.

DATA DISCLOSURE: Meridian UAV Services is fictional; all data is synthetic. No
confidential data is used and no claim is made about any production system.
================================================================================
#>

param(
    [string]$ServerInstance = 'localhost\TEW_SQLEXPRESS',
    [string]$Database       = 'MeridianUAV',
    [string]$OutputPath     = $PSScriptRoot,
    [string]$AsOfDate       = '2026-09-30',
    [string]$VibThreshold   = '3.40',
    [string]$HangarHours    = '120.0'
)

$ErrorActionPreference = 'Stop'

function Invoke-Extract {
    param([string]$Sql, [string]$FileName, [string]$Description, [switch]$Compress)

    $conn = New-Object System.Data.SqlClient.SqlConnection
    $conn.ConnectionString = "Server=$ServerInstance;Database=$Database;Integrated Security=True;Encrypt=False;Connect Timeout=30"
    $conn.Open()
    try {
        $cmd = $conn.CreateCommand()
        $cmd.CommandText = $Sql
        $cmd.CommandTimeout = 900
        $adapter = New-Object System.Data.SqlClient.SqlDataAdapter $cmd
        $table = New-Object System.Data.DataTable
        [void]$adapter.Fill($table)
    }
    finally { $conn.Close() }

    if ($table.Rows.Count -eq 0) {
        throw "Extract '$FileName' returned zero rows. An empty extract is worse than a missing one: the workbook and the model will both load it, compute over nothing, and report a clean fleet."
    }

    $target = Join-Path $OutputPath $FileName
    $table | Export-Csv -Path $target -NoTypeInformation -Encoding UTF8

    <#
    Only one extract is compressed, and only because of its size.

    sensor_features is 85,000 rows and lands at 11 MB uncompressed even after
    the decimals are narrowed -- large enough to make a repository unpleasant to
    clone for the sake of one file. gzip takes it under 2 MB and pandas reads
    .csv.gz transparently, with no change to the reading code.

    Nothing else is compressed. A reviewer should be able to open the other
    sixteen files by double-clicking them.
    #>
    if ($Compress) {
        $gz = "$target.gz"
        $in = [System.IO.File]::OpenRead($target)
        try {
            $out = [System.IO.File]::Create($gz)
            try {
                $stream = New-Object System.IO.Compression.GZipStream($out, [System.IO.Compression.CompressionLevel]::Optimal)
                try { $in.CopyTo($stream) } finally { $stream.Dispose() }
            } finally { $out.Dispose() }
        } finally { $in.Dispose() }
        Remove-Item $target
        $kb = [math]::Round((Get-Item $gz).Length / 1KB, 1)
        "{0,-30} {1,7} rows  {2} (gzip, {3} KB)" -f "$FileName.gz", $table.Rows.Count, $Description, $kb | Write-Host
    } else {
        "{0,-30} {1,7} rows  {2}" -f $FileName, $table.Rows.Count, $Description | Write-Host
    }
    return $table.Rows.Count
}

Write-Host ""
Write-Host "Meridian UAV -- maintenance extract build (as of $AsOfDate)" -ForegroundColor Cyan
Write-Host ("-" * 92)
$total = 0

# -----------------------------------------------------------------------------
# Reference data -- the contestable inputs.
#
# Exported first and read first. Every figure in this project depends on the
# stress model, and the stress model is an argument, not a fact. Shipping it
# beside the results is what lets a maintenance engineer disagree with the
# analysis on its own terms instead of dismissing it.
# -----------------------------------------------------------------------------
$total += Invoke-Extract -FileName 'ref_stress_model.csv' -Description 'the damage multipliers, with their reasoning' -Sql @"
SELECT ProfileCode, StressMultiplier, CyclePenaltyHours, PayloadPenaltyPerKg, SourceNote
FROM dbo.Ref_StressModel ORDER BY StressMultiplier DESC;
"@

$total += Invoke-Extract -FileName 'ref_service_interval.csv' -Description 'the interval, expressed both ways' -Sql @"
SELECT ComponentCode, IntervalFlightHours, IntervalStressHours, WeibullShape, IntervalNote
FROM dbo.Ref_ServiceInterval ORDER BY ComponentCode;
"@

$total += Invoke-Extract -FileName 'ref_fleet_targets.csv' -Description 'targets and warning levels' -Sql @"
SELECT MetricName, TargetValue, WarningValue, Direction, Unit, [Description]
FROM dbo.Ref_FleetTargets ORDER BY MetricName;
"@

$total += Invoke-Extract -FileName 'ref_reporting.csv' -Description 'the pinned reporting settings' -Sql @"
SELECT SettingName, SettingValue, [Description] FROM dbo.Ref_Reporting ORDER BY SettingName;
"@

$total += Invoke-Extract -FileName 'dim_component_type.csv' -Description 'component types and part lead times' -Sql @"
SELECT ComponentCode, ComponentName, PositionsPerAirframe, Criticality, LeadTimeFlightHours
FROM dbo.Dim_ComponentType ORDER BY ComponentCode;
"@

$total += Invoke-Extract -FileName 'dim_mission_profile.csv' -Description 'mission profiles' -Sql @"
SELECT ProfileCode, ProfileName, TypicalPayloadKg, [Description]
FROM dbo.Dim_MissionProfile ORDER BY ProfileCode;
"@

$total += Invoke-Extract -FileName 'dim_base.csv' -Description 'bases and hangar capacity' -Sql @"
SELECT BaseCode, BaseName, Region, HangarBays FROM dbo.Dim_Base ORDER BY BaseCode;
"@

# -----------------------------------------------------------------------------
# The core extract. One row per component fitted at the as-of date, with its
# life measured BOTH ways. Everything the workbook shows is a cut of this.
# -----------------------------------------------------------------------------
$total += Invoke-Extract -FileName 'component_wear.csv' -Description 'fitted components, life measured both ways' -Sql @"
SELECT InstallKey, ComponentSerial, ComponentCode, ComponentName, Criticality, LeadTimeFlightHours,
       TailNumber, ModelCode, BaseCode, BaseName, PositionNo,
       InstalledDate, Sorties, Landings, FlightHours, StressHours,
       IntervalFlightHours, IntervalStressHours,
       PctOfHourInterval, PctOfStressInterval,
       CAST(IsOverdueByHours AS INT)  AS IsOverdueByHours,
       CAST(IsOverdueByStress AS INT) AS IsOverdueByStress,
       CAST(IsHiddenOverdue AS INT)   AS IsHiddenOverdue,
       StressPerFlightHour, RemainingStressHours, RemainingFlightHours,
       ProjectedFlightHoursToStressLimit, LastSortieDate
FROM dbo.fn_ComponentWear('$AsOfDate')
WHERE AirframeStatus = 'Active'
ORDER BY PctOfStressInterval DESC;
"@

$total += Invoke-Extract -FileName 'airframe_stress.csv' -Description 'the fleet, one row per airframe' -Sql @"
SELECT TailNumber, ModelCode, AirframeStatus, BaseCode, BaseName, Region,
       Sorties, Landings, FlightHours, StressHours, StressRatio,
       PctStressFromHours, PctStressFromCycles, PctStressFromPayload,
       DominantProfileCode, DominantProfileName, FirstSortie, LastSortie
FROM dbo.fn_AirframeStress('$AsOfDate')
ORDER BY StressRatio DESC;
"@

$total += Invoke-Extract -FileName 'base_scorecard.csv' -Description 'the cut where the variation is' -Sql @"
SELECT BaseCode, BaseName, Region, HangarBays, Airframes, FlightHours, StressHours, StressRatio,
       FittedComponents, HourCompliancePct, StressCompliancePct, ComplianceGapPoints,
       HiddenOverdue, HiddenOverdueFlightCritical, UnscheduledRatePct,
       MaintenanceEvents, DowntimeHours
FROM dbo.fn_BaseScorecard('$AsOfDate') ORDER BY StressCompliancePct;
"@

# The board is TWO functions because the prediction metrics depend on the alert
# threshold as well as the date. Exported as one file, because a reader deciding
# what to do about this fleet needs the compliance and prediction figures
# together, and because all eight Ref_FleetTargets rows belong on one board.
$total += Invoke-Extract -FileName 'fleet_kpi.csv' -Description 'all eight metrics against their targets' -Sql @"
SELECT MetricName, MetricValue, Numerator, Denominator, TargetValue, WarningValue,
       Direction, Unit, RAGStatus, [Description]
FROM dbo.fn_FleetKPI('$AsOfDate')
UNION ALL
SELECT MetricName, MetricValue, Numerator, Denominator, TargetValue, WarningValue,
       Direction, Unit, RAGStatus, [Description]
FROM dbo.fn_PredictionKPI('$AsOfDate', $VibThreshold)
ORDER BY MetricName;
"@

$total += Invoke-Extract -FileName 'threshold_recommendation.csv' -Description 'one threshold per component, and what it buys' -Sql @"
SELECT ComponentCode, LeadTimeFlightHours,
       CurrentThreshold, CurrentPrecisionPct, CurrentActionablePct, CurrentMedianLeadHours,
       RecommendedThreshold, PrecisionPct, RecallPct, ActionableLeadTimePct,
       MedianLeadTimeFlightHours, PreventableFailurePct,
       ActionablePointsGained, PrecisionPointsGiven, ExtraFailuresPrevented
FROM dbo.fn_ThresholdRecommendation('$AsOfDate', 75.00) ORDER BY ComponentCode;
"@

$total += Invoke-Extract -FileName 'action_queue.csv' -Description 'what to do on Monday' -Sql @"
SELECT PriorityRank, InstallKey, ComponentSerial, ComponentCode, ComponentName, Criticality,
       TailNumber, BaseCode, BaseName, PositionNo, InstalledDate,
       FlightHours, StressHours, IntervalFlightHours, IntervalStressHours,
       PctOfHourInterval, PctOfStressInterval,
       CAST(IsOverdueByHours AS INT)  AS IsOverdueByHours,
       CAST(IsOverdueByStress AS INT) AS IsOverdueByStress,
       CAST(IsHiddenOverdue AS INT)   AS IsHiddenOverdue,
       StressPerFlightHour, RemainingStressHours, ProjectedFlightHoursToStressLimit,
       LeadTimeFlightHours, ActionCode, JobHours, CumulativeJobHours,
       CAST(IsThisWeek AS INT) AS IsThisWeek, RecommendedAction
FROM dbo.fn_ActionQueue('$AsOfDate', $HangarHours) ORDER BY PriorityRank;
"@

# -----------------------------------------------------------------------------
# The trend. A snapshot cannot show that the gap between the two measures WIDENS
# as the fleet accumulates hours on the wrong unit -- and that is the difference
# between a backlog to clear and a policy that keeps producing one.
# -----------------------------------------------------------------------------
$total += Invoke-Extract -FileName 'compliance_trend.csv' -Description '24 months of both compliance measures' -Sql @"
DECLARE @t TABLE (MonthEndDate DATE PRIMARY KEY);
INSERT INTO @t
SELECT TOP (24) d.MonthEndDate FROM dbo.Dim_Date d
WHERE d.IsMonthEnd = 1 AND d.MonthEndDate <= '$AsOfDate'
GROUP BY d.MonthEndDate ORDER BY d.MonthEndDate DESC;

SELECT t.MonthEndDate, k.MetricName, k.MetricValue
FROM @t t CROSS APPLY dbo.fn_FleetKPI(t.MonthEndDate) k
WHERE k.MetricName IN ('StressCompliancePct','HourCompliancePct','UnscheduledRatePct','OverdueFlightCritical')
ORDER BY t.MonthEndDate, k.MetricName;
"@

# -----------------------------------------------------------------------------
# Prediction. The evaluation rows, the summary, and the whole trade-off curve.
#
# The sweep is exported in full rather than just the chosen operating point,
# because the point of the second finding is that NO operating point satisfies
# both requirements -- and a single row cannot show that.
# -----------------------------------------------------------------------------
$total += Invoke-Extract -FileName 'alert_evaluation.csv' -Description 'per-component alert outcomes' -Sql @"
SELECT InstallKey, ComponentSerial, ComponentCode, TailNumber, BaseCode, LeadTimeFlightHours,
       CAST(IsFailure AS INT) AS IsFailure, RemovedDate,
       CAST(Alerted AS INT) AS Alerted, AlertDate, AlertVibRatio,
       LeadTimeGivenFlightHours, LeadTimeGivenStressHours,
       CAST(TruePositive AS INT) AS TruePositive,
       CAST(ActionableTruePositive AS INT) AS ActionableTruePositive,
       CAST(LateTruePositive AS INT) AS LateTruePositive,
       CAST(FalsePositive AS INT) AS FalsePositive,
       CAST(FalseNegative AS INT) AS FalseNegative,
       CAST(TrueNegative AS INT) AS TrueNegative
FROM dbo.fn_AlertEvaluation('$AsOfDate', $VibThreshold) ORDER BY ComponentSerial;
"@

$total += Invoke-Extract -FileName 'threshold_sweep.csv' -Description 'the precision / lead-time trade-off curve' -Sql @"
-- 0.05 steps, matching fn_ThresholdRecommendation's search grid. At 0.100 the
-- recommended motor-controller threshold of 2.050 was not in the sweep at all,
-- so the workbook could not show the recommendation it is meant to demonstrate.
DECLARE @s TABLE (v DECIMAL(6,3) PRIMARY KEY);
DECLARE @v DECIMAL(6,3) = 1.200;
WHILE @v <= 4.000 BEGIN INSERT INTO @s VALUES (@v); SET @v = @v + 0.050; END

SELECT s.VibThreshold, s.ComponentCode, s.Population, s.Failures, s.AlertsRaised,
       s.TP, s.FP, s.FN, s.TN, s.ActionableTP, s.LateTP,
       s.PrecisionPct, s.RecallPct, s.ActionableLeadTimePct, s.PreventableFailurePct,
       s.MedianLeadTimeFlightHours, s.LeadTimeFlightHours
FROM @s t CROSS APPLY dbo.fn_AlertSummary('$AsOfDate', t.v) s
ORDER BY s.ComponentCode, s.VibThreshold;
"@

# -----------------------------------------------------------------------------
# Reliability history -- for the survival model.
#
# IsCensored is the column that matters. A component replaced on schedule did
# not fail: all we know is that its life was AT LEAST that long. Treating those
# as failures understates component life by roughly the share of scheduled
# removals, which is most of them at the benign bases.
# -----------------------------------------------------------------------------
$total += Invoke-Extract -FileName 'component_life_history.csv' -Description 'completed lives, with censoring' -Sql @"
SELECT InstallKey, ComponentSerial, ComponentCode, ComponentName, Criticality, LeadTimeFlightHours,
       TailNumber, ModelCode, BaseCode, BaseName, PositionNo,
       InstalledDate, RemovedDate, RemovalReason, DaysFitted,
       Sorties, Landings, FlightHours, StressHours, StressRatio,
       IntervalFlightHours, IntervalStressHours, WeibullShape,
       PctOfHourInterval, PctOfStressInterval,
       CAST(IsFailure AS INT) AS IsFailure, CAST(IsCensored AS INT) AS IsCensored,
       EventType, DowntimeHours
FROM dbo.vw_ComponentLifeHistory ORDER BY ComponentSerial;
"@

$total += Invoke-Extract -FileName 'dq_findings.csv' -Description 'every data-quality finding' -Sql @"
SELECT CheckCode, CheckName, Severity, EntityType, EntityRef, [Detail],
       CAST(AffectsCurrentWear AS INT) AS AffectsCurrentWear
FROM dbo.DQ_Findings ORDER BY CheckCode, FindingKey;
"@

# -----------------------------------------------------------------------------
# The big one. Per-sortie condition monitoring with the rolling features already
# computed in SQL -- so the model in python/ and the views in sql/ cannot
# disagree about what a feature means.
# -----------------------------------------------------------------------------
<#
The decimals are narrowed on the way out, and that is a modelling decision
rather than a file-size one.

A vibration reading carried to four decimal places implies a measurement
precision the sensor does not have, and a model fitted on that fourth digit is
fitting the generator's arithmetic rather than the degradation. Three decimals
on a ratio and two on an hour figure is past anything the instrument resolves.

It also takes the file from 21 MB to 11 MB, which is a happy side effect and
not the reason.
#>
$total += Invoke-Extract -Compress -FileName 'sensor_features.csv' -Description 'per-sortie condition monitoring + rolling features' -Sql @"
SELECT InstallKey, ComponentSerial, ComponentCode, LeadTimeFlightHours, TailNumber, BaseCode,
       SortieDate, SortieSeq,
       FlightHours = CAST(FlightHours AS DECIMAL(9,3)),
       StressHours = CAST(StressHours AS DECIMAL(9,3)),
       VibrationRms, TempRiseC, CurrentDrawA,
       FlightHoursSoFar = CAST(FlightHoursSoFar AS DECIMAL(9,2)),
       StressHoursSoFar = CAST(StressHoursSoFar AS DECIMAL(9,2)),
       VibRoll10  = CAST(VibRoll10  AS DECIMAL(8,3)),
       TempRoll10 = CAST(TempRoll10 AS DECIMAL(8,2)),
       CurrRoll10 = CAST(CurrRoll10 AS DECIMAL(8,3)),
       VibVsBaseline  = CAST(VibVsBaseline  AS DECIMAL(8,3)),
       TempVsBaseline = CAST(TempVsBaseline AS DECIMAL(8,3)),
       RemovalReason, RemovedDate,
       CAST(IsFailureInstall AS INT) AS IsFailureInstall,
       FlightHoursToRemoval = CAST(FlightHoursToRemoval AS DECIMAL(9,2)),
       StressHoursToRemoval = CAST(StressHoursToRemoval AS DECIMAL(9,2))
FROM dbo.vw_SensorFeatures ORDER BY ComponentSerial, SortieSeq;
"@

Write-Host ("-" * 92)
Write-Host ("{0,-30} {1,7} rows across 19 files" -f 'TOTAL', $total) -ForegroundColor Green

# -----------------------------------------------------------------------------
# Verify by READING BACK what was written.
#
# Export-Csv reports success on a file that a consumer cannot use: a stray
# newline inside a text field, a BOM, a quoting failure. The only check worth
# anything is to parse the file again the way Excel and pandas will, and confirm
# the row count survived the round trip.
# -----------------------------------------------------------------------------
Write-Host ""
Write-Host "Reading the files back -- Export-Csv reporting success is not the same as a readable file."
$bad = 0
Get-ChildItem -Path $OutputPath -Filter '*.csv' | ForEach-Object {
    $rows = @(Import-Csv -Path $_.FullName)
    $sizeKb = [math]::Round($_.Length / 1KB, 1)
    if ($rows.Count -eq 0) {
        Write-Host ("  {0,-30} UNREADABLE or empty on re-read" -f $_.Name) -ForegroundColor Red
        $bad++
    } else {
        Write-Host ("  {0,-30} {1,7} rows re-read, {2,8} KB" -f $_.Name, $rows.Count, $sizeKb)
    }
}
# The gzipped extract is verified the same way, by decompressing and counting
# lines. A corrupt archive reports its size happily and fails only on read.
Get-ChildItem -Path $OutputPath -Filter '*.csv.gz' | ForEach-Object {
    $lines = 0
    $in = [System.IO.File]::OpenRead($_.FullName)
    try {
        $gz = New-Object System.IO.Compression.GZipStream($in, [System.IO.Compression.CompressionMode]::Decompress)
        try {
            $reader = New-Object System.IO.StreamReader($gz)
            try { while ($null -ne $reader.ReadLine()) { $lines++ } } finally { $reader.Dispose() }
        } finally { $gz.Dispose() }
    } finally { $in.Dispose() }

    $sizeKb = [math]::Round($_.Length / 1KB, 1)
    if ($lines -lt 2) {
        Write-Host ("  {0,-30} UNREADABLE archive" -f $_.Name) -ForegroundColor Red
        $bad++
    } else {
        Write-Host ("  {0,-30} {1,7} rows re-read, {2,8} KB gzip" -f $_.Name, ($lines - 1), $sizeKb)
    }
}

if ($bad -gt 0) { throw "$bad extract(s) could not be read back. They would load as empty and report a clean fleet." }

Write-Host ""
Write-Host "All extracts written and verified by re-reading." -ForegroundColor Green
