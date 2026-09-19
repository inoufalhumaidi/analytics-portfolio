<#
================================================================================
Project 3 -- Lumen Optics Manufacturing: Photonics Spend Scorecard
Script:  Export_Extracts.ps1
Purpose: Write the CSV extracts that stand in for a nightly ERP download.

WHY CSV AND NOT A LIVE CONNECTION
    The Excel workbook is a deliverable in its own right, and a reviewer should
    be able to open it and see working formulas without installing SQL Server,
    restoring a database or holding credentials. These extracts are what the
    workbook reads through Power Query. Power BI connects to SQL Server
    directly, because a semantic model belongs next to the warehouse.

WHY POWERSHELL AND NOT bcp OR sqlcmd -s","
    Vendor names contain commas and the recommended-action text contains both
    commas and apostrophes. sqlcmd's separator option does not quote fields, so
    those rows split into extra columns and every downstream count is quietly
    wrong. Export-Csv quotes correctly.

DATA DISCLOSURE: Lumen Optics Manufacturing is fictional; all data is synthetic.
No confidential data and no production system is involved.
================================================================================
#>

param(
    [string]$ServerInstance = 'localhost\TEW_SQLEXPRESS',
    [string]$Database       = 'LumenSpend',
    [string]$OutputPath     = $PSScriptRoot,
    [string]$AsOf           = '2025-12-31'
)

$ErrorActionPreference = 'Stop'

function Invoke-Extract {
    param([string]$Sql, [string]$FileName, [string]$Description)

    $conn = New-Object System.Data.SqlClient.SqlConnection
    $conn.ConnectionString = "Server=$ServerInstance;Database=$Database;Integrated Security=True;Encrypt=False;Connect Timeout=30"
    $conn.Open()
    try {
        $cmd = $conn.CreateCommand()
        $cmd.CommandText = $Sql
        $cmd.CommandTimeout = 600
        $adapter = New-Object System.Data.SqlClient.SqlDataAdapter $cmd
        $table = New-Object System.Data.DataTable
        [void]$adapter.Fill($table)
    }
    finally { $conn.Close() }

    $target = Join-Path $OutputPath $FileName
    $table | Export-Csv -Path $target -NoTypeInformation -Encoding UTF8
    "{0,-30} {1,7} rows  {2}" -f $FileName, $table.Rows.Count, $Description | Write-Host
    return $table.Rows.Count
}

Write-Host ""
Write-Host "Lumen Optics Manufacturing -- ERP extract build (as of $AsOf)" -ForegroundColor Cyan
Write-Host ("-" * 78)
$total = 0

$total += Invoke-Extract -FileName 'po_line_cost.csv' -Description 'dbo.vw_POLineCost' -Sql @"
SELECT PONumber, PartNumber, PartName, Category, Criticality,
       QualifiedSupplierCount, QualificationMonths, CAST(IsSingleSource AS INT) AS IsSingleSource,
       VendorID, VendorName, Country, VendorTier, LumenRevenueSharePct,
       BuyerID, BuyerName, Team,
       CONVERT(CHAR(10), OrderDate, 23) AS OrderDate,
       CONVERT(CHAR(10), PromisedDate, 23) AS PromisedDate,
       OrderYear, OrderYearMonth,
       OrderQty, UnitPrice, FreightAmount, ExpediteFee, CAST(IsExpedited AS INT) AS IsExpedited,
       ExtendedPrice, LandedCost,
       ISNULL(AgreementNo,'') AS AgreementNo, ContractedUnitPrice, AgreementsInForce,
       CAST(IsOnContract AS INT) AS IsOnContract, PPVAmount, ContractedExtended,
       QtyReceived, QtyAccepted, QtyRejected, Receipts, CAST(IsReceived AS INT) AS IsReceived,
       AcceptanceRatePct, CostPerAcceptedUnit, RejectedValue,
       CONVERT(CHAR(10), LastReceiptDate, 23) AS LastReceiptDate,
       DaysLate, CAST(IsOnTime AS INT) AS IsOnTime
FROM dbo.vw_POLineCost
ORDER BY OrderDate, PONumber;
"@

$total += Invoke-Extract -FileName 'price_erosion.csv' -Description 'dbo.vw_PriceErosion' -Sql @"
SELECT PartNumber, PartName, Category, Criticality, QualifiedSupplierCount, QualificationMonths,
       VendorID, VendorName, VendorTier, LumenRevenueSharePct,
       CONVERT(CHAR(10), FirstDate, 23) AS FirstDate,
       CONVERT(CHAR(10), LastDate, 23) AS LastDate,
       YearsElapsed, Lines, TotalQty, TotalSpend, TTMQty, TTMSpend,
       FirstPrice, LastPrice, ExpectedPrice,
       BenchmarkErosionPct, ToleranceBandPct, ActualErosionPct, ErosionCapturePct,
       ErosionGapPerUnit, AnnualOpportunity
FROM dbo.vw_PriceErosion ORDER BY AnnualOpportunity DESC;
"@

$total += Invoke-Extract -FileName 'renegotiation_queue.csv' -Description 'dbo.vw_RenegotiationQueue' -Sql @"
SELECT PriorityRank, BuyerRank, PartNumber, PartName, Category, Criticality,
       QualifiedSupplierCount, QualificationMonths, CAST(IsSingleSource AS INT) AS IsSingleSource,
       VendorID, VendorName, VendorTier, LumenRevenueSharePct,
       BuyerID, BuyerName, Team,
       FirstPrice, LastPrice, ExpectedPrice,
       BenchmarkErosionPct, ActualErosionPct, ErosionCapturePct, ErosionGapPerUnit,
       TTMQty, TTMSpend, AnnualOpportunity, RejectedValue, PPV, OffContractSpend,
       AcceptanceRatePct, LeverageScore, ActionCode,
       CAST(IsThisQuarter AS INT) AS IsThisQuarter, RecommendedAction
FROM dbo.vw_RenegotiationQueue ORDER BY PriorityRank;
"@

$total += Invoke-Extract -FileName 'vendor_scorecard.csv' -Description 'dbo.usp_VendorScorecard' -Sql @"
SELECT v.VendorID, v.VendorName, v.Country, v.VendorTier, v.LumenRevenueSharePct,
       s.Lines, s.TotalSpend, s.LandedSpend, s.FreightSpend, s.ExpediteSpend,
       s.ExpeditePctOfSpend, s.OnContractPct, s.PPVAmount, s.PPVPct,
       s.QtyReceived, s.QtyAccepted, s.AcceptanceRatePct, s.RejectedValue,
       s.OnTimePct, s.AvgDaysLate, s.CostIndexVsBest, s.SharedParts
FROM dbo.fn_VendorScorecard('$AsOf') s
JOIN dbo.Dim_Vendor v ON v.VendorKey = s.VendorKey
ORDER BY s.TotalSpend DESC;
"@

$total += Invoke-Extract -FileName 'spend_kpi_monthly.csv' -Description 'dbo.vw_SpendKPIMonthly' -Sql @"
SELECT CONVERT(CHAR(10), AsOfDate, 23) AS AsOfDate, Lines, TotalSpend, LandedSpend,
       PPVAmount, PPVPct, MaverickSpendPct, AcceptanceRatePct, RejectedValue,
       OnTimeDeliveryPct, ExpediteSpendPct, ErosionCapturePct, ErosionOpportunity,
       SingleSourceSpendPct, Top5VendorSharePct
FROM dbo.vw_SpendKPIMonthly ORDER BY AsOfDate;
"@

# A procedure fills a DataTable perfectly well, so these summaries come from the
# procedure itself rather than being re-implemented here. Re-implementing would
# create a second definition that can drift from the one the database serves.
$total += Invoke-Extract -FileName 'category_summary.csv' -Description 'dbo.usp_CategorySummary (by Category)' `
    -Sql "EXEC dbo.usp_CategorySummary @AsOf = '$AsOf', @GroupBy = 'Category';"

$total += Invoke-Extract -FileName 'team_summary.csv' -Description 'dbo.usp_CategorySummary (by Team)' `
    -Sql "EXEC dbo.usp_CategorySummary @AsOf = '$AsOf', @GroupBy = 'Team';"

$total += Invoke-Extract -FileName 'erosion_benchmark.csv' -Description 'dbo.Ref_PriceErosionBenchmark' -Sql @"
SELECT Category, AnnualErosionPct, ToleranceBandPct, MaturityProfile, SourceNote
FROM dbo.Ref_PriceErosionBenchmark ORDER BY Category;
"@

$total += Invoke-Extract -FileName 'spend_targets.csv' -Description 'dbo.Ref_SpendTargets' -Sql @"
SELECT MetricName, TargetValue, WarningValue, Direction, Unit, Description
FROM dbo.Ref_SpendTargets ORDER BY MetricName;
"@

$total += Invoke-Extract -FileName 'reject_reasons.csv' -Description 'dbo.Ref_RejectReason' -Sql @"
SELECT RejectReasonCode, Description, Attribution FROM dbo.Ref_RejectReason ORDER BY Attribution, RejectReasonCode;
"@

$total += Invoke-Extract -FileName 'data_quality_summary.csv' -Description 'dbo.vw_DQ_Summary' -Sql @"
SELECT AnomalyType, EntityType, Severity, ImpactClass, Anomalies, AmountAtRisk,
       PopulationScanned, AnomalyRatePct, CAST(CountsTowardExposure AS INT) AS CountsTowardExposure
FROM dbo.vw_DQ_Summary ORDER BY Severity, AnomalyType;
"@

Write-Host ("-" * 78)
Write-Host ("{0,-30} {1,7} rows written in total" -f 'TOTAL', $total) -ForegroundColor Green
Write-Host ""
