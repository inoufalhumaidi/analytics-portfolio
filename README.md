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
| 2 | [Vantage Wholesale Supply: Receivables Performance](project-2-vantage-receivables/) | Of the days DSO has risen, how many did we grant through longer terms, how many are customers taking, and how many are our own unapplied cash — and who should collections call first? | Excel · Power Query · SQL | ✅ Complete |
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

## Project 2 at a glance

- **Deterministic synthetic data** — 29,971 invoices, 27,601 cash receipts, 29,312 cash applications, 1,425 adjustments and 1,347 promises to pay, generated from hash-keyed pseudo-randomness rather than `RAND()`, so a rebuild reproduces the dataset byte for byte and every quoted figure stays checkable. A UAT case asserts all five fingerprints.
- **An exact DSO decomposition.** The rise in DSO splits into days Vantage *granted* through longer terms and days customers *took* beyond them, and the three components sum to classic DSO to the cent — no residual to plug. A shift-share decomposition then separates a deliberate terms change (+1.71 days, the same 400 customers re-papered) from ordinary sales mix.
- **A schema decision that changed a conclusion.** Modelling the cash receipt and its application as separate facts made *unapplied cash* expressible for the first time — and moved **2.38 days of DSO** out of "customers are paying late" and into "we have not applied our own $419,392". A shortcut that cannot represent a category does not leave a gap in the output; it attributes that category to whatever it can represent.
- **A target shown to be arithmetically unachievable.** The 45-day DSO target sits 1.68 days *below* the weighted average terms actually sold, so it would be missed even if every customer paid exactly on the due date.
- **Excel as an operational control**, not a report: countback DSO implemented as a real month-by-month exhaustion loop in worksheet formulas, a capacity-bounded call list of 65 accounts carrying 81.3% of collectable exposure, a second worklist routing 51 already-paid accounts *out* of collections, and a validation sheet that reconciles all 27 headline figures against SQL and says so on its own face.
- **A data-quality gate that fails on purpose** — 96 planted duplicate receipts mis-state AR by 3.421%, above a 1% tolerance, and the build does not suppress it.

Three defects in the validation report are worth reading: a generator bug that produced a clean, plausible and entirely false finding ("47% of discount-terms invoices are 90+ days past due") which no balance-level check could see; an Excel build that reported success while producing a file full of `#REF!`; and a UAT case whose premise a later feature invalidated, so the test failed while the data was right. See [`docs/data_validation_report.md`](project-2-vantage-receivables/docs/data_validation_report.md).

Start with the project [README](project-2-vantage-receivables/README.md) for the reproduction steps.
