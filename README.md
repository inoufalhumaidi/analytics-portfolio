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
| 3 | [Lumen Optics Manufacturing: Photonics Spend Scorecard](project-3-lumen-optics-spend/) | Our prices are flat and every variance report is green — so why is material cost per accepted unit rising, and where should sourcing renegotiate first? | SQL Server · Power BI · DAX · Excel | ✅ Complete |
| 4 | Talon Robotics: Payload Deployment Delivery | Is the new payload ready to ship, and what is still open? | Requirements · Traceability · UAT · RAID | Planned |
| 5 | Meridian UAV Services: Predictive Maintenance | Which airframes need maintenance before their next mission? | Python · SQL · Streamlit · public + synthetic data | Planned |

## Project 1 at a glance

- **Star schema** in SQL Server with a synthetic-data generator (~37,000 dispatched jobs over two years), data-quality views, KPI views, stored procedures and a 12-assertion UAT suite that runs in a rolled-back transaction and exits non-zero on failure.
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

## Project 3 at a glance

- **A metric that variance reporting is structurally blind to.** Optical components follow learning curves, so a supplier who renews a contract *flat* keeps everything the curve was supposed to hand back — and because the price never rose, purchase price variance records zero forever. Purchase price variance reads **−0.19% (green)** while **$4,047,609 a year**, 8.3% of direct-materials spend, is going uncollected.
- **The gap is worst exactly where leverage is weakest.** Parts with two or more qualified suppliers capture 38.26% of the expected erosion; sole-source parts capture **8.91%**. Suppliers hold price precisely where you cannot leave — obvious economically, and now quantified at $1.47M.
- **The cheapest supplier is the most expensive one.** On a shared laser diode part, the vendor with the lowest unit price ($184.14 against $197.43) is *dearer per accepted unit* once a 91.6% acceptance rate is counted. Across the book, **$1,571,168** of material was paid for and cannot be used.
- **A queue that decides the action by leverage, not by value.** 185 vendor-part pairs ranked, bounded to the **43** a six-person sourcing team can actually run in a quarter — covering 65.6% of the gap — with sole-source pairs routed to `DUAL_SOURCE` rather than into negotiations nobody can win, and 73 low-value pairs honestly parked as `ACCEPT`.
- **A scripted Power BI model, reconciled rather than eyeballed.** 9 tables, 10 relationships and
69 DAX measures are built by script into a live Desktop session, and a second script queries the
model *in DAX* and reconciles 31 figures against the SQL that produces them — including the
like-for-like cost index for all 36 vendors individually. That validator earned its place
immediately: the first version of the cost-index measure returned exactly **100.00** for every
filtered vendor, because the inner `MINX` intersected with the outer vendor filter and compared
each vendor against itself. On an index where 100 means "best available", that reads as a result.
- **Dated price agreements, deterministic under overlap.** A contracted price is true between two dates; open-ended agreements are still in force, and the 61 lines with two agreements at once are *reported as ambiguous* rather than silently resolved, because a variance computed off a coin-flip has no error bar.

Three defects in the validation report are worth reading: a planted pattern that silently never existed because an `UPDATE` matched zero rows; a contract-lapse rule that produced the opposite of a lapse, caught only because a figure refused to move when its driver doubled; and a window function inside a `CROSS APPLY` that saw a single row and reported "100% of vendor-part pairs holding price flat" directly beneath a table showing prices falling. See [`docs/data_validation_report.md`](project-3-lumen-optics-spend/docs/data_validation_report.md).

Start with the project [README](project-3-lumen-optics-spend/README.md) for the reproduction steps.
