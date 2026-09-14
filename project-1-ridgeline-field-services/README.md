# Project 1 — Ridgeline Field Services: Operational Efficiency

**Business question:** Where is technician capacity being wasted, and which dispatch decisions should change this week?

Ridgeline Field Services is a fictional HVAC/electrical field-service dispatch company. All data in this project is synthetically generated — no confidential, proprietary, or production data of any real organization is used or represented anywhere in this project.

## Stack

SQL Server (star schema, reusable views/procs) → Power BI (DAX) → Excel (formula-driven operational control) → live HTML dashboard.

## How to reproduce this project from scratch

1. **Database.** Run these against a SQL Server instance, in order:
   ```
   sqlcmd -S <your_instance> -E -C -i sql/01_create_schema.sql
   sqlcmd -S <your_instance> -E -C -i sql/02_generate_synthetic_data.sql
   sqlcmd -S <your_instance> -E -C -i sql/03_data_quality_checks.sql
   sqlcmd -S <your_instance> -E -C -i sql/04_kpi_views.sql
   sqlcmd -S <your_instance> -E -C -i sql/05_stored_procedures.sql
   ```
   Then validate the logic before trusting anything built on top of it:
   ```
   sqlcmd -S <your_instance> -E -C -i sql/06_uat_test_cases.sql
   ```
   This inserts an isolated test fixture inside a transaction, asserts 8 KPI test cases against hand-calculated expected values, prints a PASS/FAIL report, and rolls back automatically — the production dataset is never touched by running it.

2. **Excel workbook.** Edit the connection string at the top of `excel/Build_Workbook.ps1` if your instance name differs from `localhost\TEW_SQLEXPRESS`, then run it:
   ```
   powershell -File excel/Build_Workbook.ps1
   ```
   This produces `excel/Ridgeline_KPI_Dashboard.xlsx` directly from the live database via ADO.NET + Excel COM automation — every number in it is a live formula against the `tblJobs` table, not a paste of query results.

3. **Power BI.** `powerbi/Ridgeline_KPI_Dashboard.pbix` — its data model (7 tables, 6 relationships of which 5 active, 28 DAX measures, a Year › MonthName date hierarchy) was built by running `powerbi/Build_PowerBI_Model.ps1` against an open Power BI Desktop session (it scripts the local Analysis Services engine Desktop exposes while a file is open — the same mechanism Tabular Editor uses). Report pages and visuals were built by hand in Desktop following `powerbi/POWER_BI_BUILD_GUIDE.md` Section 5. Three report-layer fixes were then applied directly to the report definition inside the `.pbix` (Power BI's PBIR format, which follows Microsoft's published JSON schemas): navigation-button actions, slicer sync to pages 2 and 3, and the per-technician scatter. The edit was checked against those schemas, made from a backup, and accepted only after the edited file opened successfully in Power BI Desktop. The data model is byte-for-byte identical to before. The package's `SecurityBindings` part had to be dropped for the file to open, so Desktop may ask you to confirm the SQL Server credentials or privacy level on the next refresh. See `docs/data_validation_report.md` §8. Re-run the script against a blank report to reproduce the data model from scratch; see `powerbi/DAX_Measures.md` for the annotated measure definitions.

4. **Live dashboard.** `dashboard/index.html` is a self-contained interactive dashboard: download it and open it in any browser, with no server or install. Its data was exported from the SQL views into `data_exports/`, and the file follows light or dark mode from your system setting.

5. **Case study.** `case_study/Project1_Ridgeline_Case_Study.pdf` — the decision-focused narrative built from everything above.

## Repository map

```
sql/            Star schema, synthetic data generator, data quality checks, KPI views, stored procs, UAT tests
excel/          Build_Workbook.ps1 (source of truth) + Ridgeline_KPI_Dashboard.xlsx (generated artifact)
powerbi/        Build_PowerBI_Model.ps1 (scripts the data model) + Ridgeline_KPI_Dashboard.pbix + DAX measure library + build guide
dashboard/      Live HTML/JS dashboard
docs/           Requirements traceability matrix, data validation/profiling report
case_study/     Case study PDF + HTML source
```

## What makes this an operational control, not a report

- Every KPI threshold lives in one table (`dbo.Ref_KPITargets`), mirrored onto the Excel `Targets` sheet — change a number there and every dashboard, flag, and recommended action recalculates against it.
- The `Priority_Action_Queue` (both the SQL view and the Excel sheet) outputs a ranked, actionable worklist with a specific recommended action per flagged technician — not just a chart.
- Every SQL view has a reusable, parameterized stored procedure wrapper (`sql/05_stored_procedures.sql`) so the same logic is callable from a scheduled job, a BI tool, or an analyst's ad hoc query without re-deriving it.
- The data quality layer is implemented **twice, independently** — once in SQL (`sql/03_data_quality_checks.sql`) and once in Excel formulas (`Data_Quality` sheet) — and the two are cross-checked against each other rather than assumed to agree.
