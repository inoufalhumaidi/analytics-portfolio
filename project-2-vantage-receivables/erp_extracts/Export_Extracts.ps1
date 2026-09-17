<#
================================================================================
Project 2 -- Vantage Wholesale Supply: Receivables Performance
Script:  Export_Extracts.ps1
Purpose: Write the CSV extracts that stand in for a nightly ERP download.

WHY CSV AND NOT A LIVE CONNECTION
    The Excel workbook is the deliverable, and a reviewer should be able to
    open it and see working formulas without installing SQL Server, restoring a
    database or holding credentials. These extracts are what the workbook reads
    through Power Query. The SQL Server queries that produce them are named on
    each file, and the workbook's README sheet carries the M code to repoint it
    at a live database instead.

WHY POWERSHELL AND NOT bcp OR sqlcmd -s","
    Customer names contain commas ("Westbrook Industries - West"). sqlcmd's
    separator option does not quote fields, so those rows split into extra
    columns and every downstream count is quietly wrong. Export-Csv quotes
    correctly.

DATA DISCLOSURE: Vantage Wholesale Supply is fictional; all data is synthetic.
No confidential data and no production system is involved.
================================================================================
#>

param(
    [string]$ServerInstance = 'localhost\TEW_SQLEXPRESS',
    [string]$Database       = 'VantageAR',
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
    # -NoEnumerate keeps PowerShell from unrolling the DataTable into DataRows,
    # which silently drops the column schema and produces a file of type names.
    $table | Export-Csv -Path $target -NoTypeInformation -Encoding UTF8
    $rows = $table.Rows.Count
    "{0,-34} {1,8} rows  {2}" -f $FileName, $rows, $Description | Write-Host
    return $rows
}

Write-Host ""
Write-Host "Vantage Wholesale Supply -- ERP extract build (as of $AsOf)" -ForegroundColor Cyan
Write-Host ("-" * 78)

$total = 0

$total += Invoke-Extract -FileName 'ar_open_items.csv' -Description 'dbo.vw_ARBalanceDetail (open items)' -Sql @"
SELECT InvoiceNo, CustomerID, CustomerName, Segment, Region, SalesRep, RiskTier, CreditLimit,
       CollectorID, CollectorName, Team, TermsCode, NetDays, DiscountPct,
       CONVERT(CHAR(10), InvoiceDate, 23) AS InvoiceDate,
       CONVERT(CHAR(10), DueDate, 23)     AS DueDate,
       InvoiceAmount, PaidToDate, DiscountToDate, CreditMemoToDate, WriteOffToDate,
       OpenBalance, OverAppliedAmount,
       CAST(IsOpen AS INT) AS IsOpen, CAST(IsPastDue AS INT) AS IsPastDue,
       DaysPastDue, BucketKey, BucketName, RiskWeight,
       CAST(IsDisputedOpen AS INT) AS IsDisputedOpen, ISNULL(DisputeReason,'') AS DisputeReason
FROM dbo.vw_ARBalanceDetail
WHERE IsOpen = 1          -- an ERP ageing download is open items, not all history
ORDER BY InvoiceNo;
"@

$total += Invoke-Extract -FileName 'priority_action_queue.csv' -Description 'dbo.vw_PriorityActionQueue' -Sql @"
SELECT PriorityRank, CollectorRank, CustomerID, CustomerName, Segment, Region, RiskTier,
       CollectorID, CollectorName, Team, TermsCode,
       OpenBalance, PastDueBalance, Balance90Plus, DisputedBalance,
       WeightedExposure, DisputedWeightedExposure, CollectableExposure,
       OldestDaysPastDue, CreditUtilizationPct, AvgDaysLateHistoric,
       UnappliedCash, NetExposure, BrokenPromises90d,
       ActionCode, CAST(CreditHoldFlag AS INT) AS CreditHoldFlag,
       CAST(IsTodaysWorklist AS INT) AS IsTodaysWorklist, RecommendedAction
FROM dbo.vw_PriorityActionQueue ORDER BY PriorityRank;
"@

$total += Invoke-Extract -FileName 'ar_kpi_monthly.csv' -Description 'dbo.vw_ARKPIMonthly' -Sql @"
SELECT YearMonth, CONVERT(CHAR(10), AsOfDate, 23) AS AsOfDate,
       BeginAR, MonthSales, EndAR, EndCurrentAR, PastDueAR, AR90Plus, DisputedAR,
       OpenInvoices, PastDueInvoices, Invoices90Plus,
       CashApplied, CashReceived, CashAppliedLessReceived, UnappliedCash, UnappliedCashPct,
       WriteOffs, CreditMemos,
       DSO_Countback, DSO_Simple, BPDSO_Countback, AvgDaysDelinquent, DSO_MethodGap,
       CollectableBase, CEI_Book, CEI_Cash, PaperCollectionsGap,
       PctPastDue, Pct90Plus, PctDisputed, PctPastDueByCount, Pct90PlusByCount,
       CAST(IsComparablePeriod AS INT) AS IsComparablePeriod
FROM dbo.vw_ARKPIMonthly ORDER BY AsOfDate;
"@

$total += Invoke-Extract -FileName 'dso_bridge_monthly.csv' -Description 'dbo.vw_DSOBridgeMonthly' -Sql @"
SELECT CONVERT(CHAR(10), AsOfDate, 23) AS AsOfDate,
       TotalAR, CurrentAR, DisputedPastDueAR, UndisputedPastDueAR, SalesPerDay,
       DSO_Classic, GrantedDays, DisputeDays, LatenessDays,
       WeightedAvgTermsDays, TermsPositionDays, RecoverableDays, RecoverableCash,
       BillingLagDays, CashCycleDays, CashPerLagDay,
       CAST(IsComparablePeriod AS INT) AS IsComparablePeriod
FROM dbo.vw_DSOBridgeMonthly ORDER BY AsOfDate;
"@

$total += Invoke-Extract -FileName 'monthly_credit_sales.csv' -Description 'dbo.vw_MonthlyCreditSales' -Sql @"
SELECT s.YearMonth, CONVERT(CHAR(10), s.MonthEndDate, 23) AS MonthEndDate,
       s.DaysInMonth, s.CreditSales, s.InvoiceCount,
       m.WeightedAvgTermsDays, m.PctDollarsNet30, m.PctDollarsNet45,
       m.PctDollarsNet60, m.PctDollarsExtended
FROM dbo.vw_MonthlyCreditSales s
JOIN dbo.vw_TermsMixMonthly m ON m.YearMonth = s.YearMonth
ORDER BY s.MonthEndDate;
"@

$total += Invoke-Extract -FileName 'cash_application_worklist.csv' -Description 'dbo.usp_CashApplicationWorklist' -Sql @"
SELECT u.CustomerKey, c.CustomerID, c.CustomerName, c.Segment, c.Region, col.CollectorName,
       u.UnappliedCash, u.UnappliedReceipts, u.FullyUnappliedReceipts,
       u.NoRemittanceAdvice, u.OldestUnappliedDays,
       ca.OpenBalance, ca.PastDueBalance, ca.NetExposure,
       WronglyChaseable = CAST(CASE WHEN ca.PastDueBalance < u.UnappliedCash
                                    THEN ca.PastDueBalance ELSE u.UnappliedCash END AS DECIMAL(14,2))
FROM dbo.fn_UnappliedCash('2025-12-31') u
JOIN dbo.Dim_Customer c    ON c.CustomerKey = u.CustomerKey
JOIN dbo.Dim_Collector col ON col.CollectorKey = c.CollectorKey
JOIN dbo.fn_CustomerAR('2025-12-31') ca ON ca.CustomerKey = u.CustomerKey
WHERE u.UnappliedCash > 0
ORDER BY u.OldestUnappliedDays DESC;
"@

$total += Invoke-Extract -FileName 'promise_status.csv' -Description 'dbo.vw_PromiseStatus' -Sql @"
SELECT PromiseNo, CustomerID, CustomerName, RiskTier, Segment, CollectorName,
       CONVERT(CHAR(10), PromiseMadeDate, 23) AS PromiseMadeDate,
       CONVERT(CHAR(10), PromisedPayDate, 23) AS PromisedPayDate,
       PromisedAmount, CashInWindow, PctOfPromise, DaysSincePromised, PromiseStatus
FROM dbo.vw_PromiseStatus ORDER BY PromisedPayDate;
"@

$total += Invoke-Extract -FileName 'billing_lag_monthly.csv' -Description 'dbo.vw_BillingLagMonthly' -Sql @"
SELECT YearMonth, CONVERT(CHAR(10), MonthEndDate, 23) AS MonthEndDate, Region,
       InvoiceCount, CreditSales, BillingLagDays
FROM dbo.vw_BillingLagMonthly ORDER BY MonthEndDate, Region;
"@

$total += Invoke-Extract -FileName 'ar_targets.csv' -Description 'dbo.Ref_ARTargets' -Sql @"
SELECT MetricName, TargetValue, WarningValue, Direction, Unit, Description FROM dbo.Ref_ARTargets ORDER BY MetricName;
"@

$total += Invoke-Extract -FileName 'aging_buckets.csv' -Description 'dbo.Ref_AgingBucket' -Sql @"
SELECT BucketKey, BucketName, MinDaysPastDue, MaxDaysPastDue, RiskWeight FROM dbo.Ref_AgingBucket ORDER BY BucketKey;
"@

$total += Invoke-Extract -FileName 'data_quality_summary.csv' -Description 'dbo.vw_DQ_Summary' -Sql @"
SELECT AnomalyType, Severity, ImpactClass, Anomalies, AmountAtRisk,
       PopulationScanned, AnomalyRatePct, CAST(CountsTowardExposure AS INT) AS CountsTowardExposure
FROM dbo.vw_DQ_Summary ORDER BY Severity, AnomalyType;
"@

Write-Host ("-" * 78)
Write-Host ("{0,-34} {1,8} rows written in total" -f 'TOTAL', $total) -ForegroundColor Green
Write-Host ""
