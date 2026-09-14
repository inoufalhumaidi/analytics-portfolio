# Requirements Traceability Matrix — Ridgeline Field Services (Project 1)

**Business question:** Where is technician capacity being wasted, and which dispatch decisions should change this week?

This matrix also exists as a formatted sheet (`Traceability_UAT`) inside `excel/Ridgeline_KPI_Dashboard.xlsx`. This file is the source of truth; the Excel sheet mirrors it.

| Req ID | Business Requirement | Functional Requirement | KPI / Data Field | Implementation | Validation / Test | UAT Result |
|---|---|---|---|---|---|---|
| BR-01 | Identify regions/technicians with wasted (underutilized) capacity | Compute technician utilization % (worked minutes ÷ available capacity minutes) per technician per month | `UtilizationPct` | `dbo.vw_TechnicianUtilization`, `dbo.vw_RegionKPIMonthly`; Excel `KPI_Dashboard!UtilizationPct` cell | UAT-01: hand-calculated 2-job fixture (255 worked min / 11,040 available min = 2.31%) matches view output exactly | PASS |
| BR-02 | Flag technicians who are over capacity (burnout/SLA risk) | Compare `UtilizationPct` against `Ref_KPITargets.UtilizationCeilingPct` warning value | `FlagOverCapacity` | `dbo.vw_PriorityActionQueue`; Excel `Priority_Action_Queue` sheet | UAT-04 variant (inverse case): manually verified against a technician exceeding 95% in the generated dataset (e.g., TECH-025, see case study) | PASS |
| BR-03 | Quantify rework/quality risk by first-time-fix and callback rate | `FirstTimeFixPct`, `CallbackPct` computed from `FirstTimeFix`/`CallbackFlag` bit columns, Completed jobs only | `FirstTimeFixPct`, `CallbackPct` | `dbo.vw_TechnicianKPIMonthly`, `dbo.vw_RegionKPIMonthly` | UAT-02a/02b: 2-job fixture with 1 FTF success / 1 failure → 50.00% each, matches exactly | PASS |
| BR-04 | Quantify SLA (arrival-window) compliance and identify the driver (travel/dispatch delay) | `SLAMet` = 1 when `DelayMin <= ServiceType.SLAHours * 60` | `SLACompliancePct` | `dbo.vw_RegionKPIMonthly`; drill-down by `Dim_ServiceType.SLAHours` | Manually verified East Valley's 1-hour-SLA segment (93.1%) vs Central District (95.6%) reconciles to `AVG(SLAMet)` on the filtered raw extract | PASS |
| BR-05 | Produce a ranked, actionable worklist (not just a report) | Composite `RiskScore` = count of breached targets; plain-language `RecommendedAction` per breach type | `RiskScore`, `RecommendedAction` | `dbo.vw_PriorityActionQueue`; Excel `Priority_Action_Queue` ranked list (LARGE + INDEX/MATCH) | UAT-04: a technician engineered to breach only the utilization target appears with `FlagLowUtilization = 1` | PASS |
| BR-06 | Guarantee no confidential/production data is used | All dimension and fact data synthetically generated with documented, seeded generation logic | N/A (governance requirement) | `sql/02_generate_synthetic_data.sql` header disclosure; `docs/data_validation_report.md` | Manual review: no real company/technician/client names or real identifiers used anywhere in the schema or data | PASS |
| BR-07 | Detect data quality problems before they reach the dashboard | Row-level anomaly detection view + rollup + QA-gate procedure | `vw_DQ_JobAnomalies`, `vw_DQ_Summary` | `sql/03_data_quality_checks.sql`; Excel `Data_Quality` sheet (independent COUNTIFS re-implementation) | UAT-03: a deliberately broken row (`ActualEnd < ActualStart`) is caught by `TIMESTAMP_INVERSION` | PASS |
| BR-08 | Every KPI must be reproducible outside the BI tool | KPI logic implemented as parameterized, reusable SQL views/procs, not embedded only in Power BI/Excel | `usp_GetRegionKPISummary`, `usp_GetPriorityActionQueue`, `usp_GetTechnicianScorecard` | `sql/05_stored_procedures.sql` | Manually executed each proc with test parameters and confirmed output matches the underlying views | PASS |
| BR-09 | Region-level rollups must reconcile to job-level detail | Aggregation integrity between `Fact_ServiceJobs` and `vw_RegionKPIMonthly` | `TotalJobs` | `dbo.vw_RegionKPIMonthly` | UAT-05: region job count from the view exactly equals a raw `COUNT(*)` on `Fact_ServiceJobs` for the same scope (563 = 563 on the current dataset; the count itself varies per generator run, the equality does not) | PASS |
| BR-10 | A technician's monthly KPIs must not leak data from other months | Month-scoped join between `Fact_ServiceJobs` and the target `Dim_Date` month, not a technician-wide join | `FirstTimeFixPct`, `SLACompliancePct`, `CallbackPct` at technician grain | `dbo.vw_TechnicianKPIMonthly` | UAT-06a/06b: a technician with jobs in two different months shows each month's isolated FTF% (66.67% Jan, 100.00% Feb), not a pooled figure (75.00%) | PASS |

### A bug this project's cross-checking actually caught

`vw_TechnicianKPIMonthly` originally joined `Fact_ServiceJobs` to a technician with no date filter (only `Dim_Date` was filtered to the target month, via a join condition that nulled non-matching dates without excluding the corresponding job rows). The practical effect: every technician's monthly First-Time-Fix/SLA/Callback percentages were silently computed across their **entire 2-year job history**, not the selected month. Region-level KPIs (`vw_RegionKPIMonthly`) were unaffected — that view joins `Dim_Date` directly on the fact table's own date.

This was not caught by the original UAT suite (UAT-02a/02b's test technician only ever had jobs in one month, so the bug had nothing to leak across). It was caught by BR-05's Excel `Priority_Action_Queue` sheet — an independent formula re-implementation of the same KPI — disagreeing with the SQL view for real technicians with multi-month histories. The fix and a new regression test (UAT-06a/06b) are both in `sql/04_kpi_views.sql` / `sql/06_uat_test_cases.sql`; the full story is in `docs/data_validation_report.md` §6.

## Requirements NOT in scope for Project 1 (documented, not silently dropped)

- Real-time/streaming dispatch data (this is a monthly-refresh analytical model, not a live dispatch system)
- Technician GPS/routing optimization (out of scope — SLA analysis flags *that* travel is a driver, not a routing engine)
- Multi-year forecasting (the 2-year window supports trend visibility, not a forecast model)

## UAT Execution

Run `sql/06_uat_test_cases.sql` against the database at any time — it inserts a fully isolated fixture inside a transaction, asserts the UAT cases referenced above, prints a PASS/FAIL report, and rolls back automatically. See the script header for details. Last executed run (2026-09-14): all 8 test cases (UAT-01, 02a, 02b, 03, 04, 05, 06a, 06b) returned PASS.
