# Analytics Portfolio

Five end-to-end analytics projects covering operational efficiency, financial operations, supply-chain spend, technical project management and predictive maintenance. Each is built around a different fictional company and a single business question, and each ships working artifacts rather than screenshots.

Every project follows the same standard:

- a clear **business question** and the decision it drives
- **validated synthetic data**, generated and profiled before any analysis
- **reusable SQL**: parameterized views and procedures, not one-off queries
- a **formula-driven Excel workbook** that works as an operational control, not a static report
- a **priority action queue**: a ranked worklist rather than just a chart
- **data quality controls**, **requirements traceability** and a **UAT suite**
- **management recommendations** in a written case study

> **Data disclosure:** every company in this repository is fictional and all data is synthetically generated. No confidential, proprietary or production data is used.

## Projects

| # | Project | Business question | Stack | Status |
|---|---|---|---|---|
| 1 | [Ridgeline Field Services: Operational Efficiency](project-1-ridgeline-field-services/) | Where is technician capacity being wasted, and which dispatch decisions should change this week? | SQL Server · Power BI · DAX · Excel | ✅ Complete |
| 2 | Vantage Wholesale Supply: Receivables Performance | Which accounts and terms are driving DSO up, and who should collections call first? | Excel · Power Query · SQL | Planned |
| 3 | Lumen Optics Manufacturing: Spend Scorecard | Which vendors and categories are eroding margin, and where should sourcing renegotiate first? | Power BI · Excel · SQL | Planned |
| 4 | Talon Robotics: Payload Deployment Delivery | Is the new payload ready to ship, and what is still open? | Requirements · Traceability · UAT · RAID | Planned |
| 5 | Meridian UAV Services: Predictive Maintenance | Which airframes need maintenance before their next mission? | Python · SQL · Streamlit · public + synthetic data | Planned |

## Project 1 at a glance

- **Star schema** in SQL Server with a synthetic-data generator (~37,000 dispatched jobs over two years), data-quality views, KPI views, stored procedures and an 8-case UAT suite that runs in a rolled-back transaction.
- **Excel operational control** built directly from the database by script, with its KPIs implemented independently of SQL and cross-checked against it.
- **Power BI report** whose data model is scripted (7 tables, 28 DAX measures) and reconciled figure-by-figure against SQL.
- **Case study** (PDF) with findings and management recommendations.

The validation report records the real bugs that cross-checking caught along the way, including a KPI view that silently pooled months and a regional utilization figure that would have been understated about fivefold. See [`docs/data_validation_report.md`](project-1-ridgeline-field-services/docs/data_validation_report.md).

Start with the project [README](project-1-ridgeline-field-services/README.md) for the reproduction steps.
