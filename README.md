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
| 4 | [Talon Robotics: Payload Deployment Delivery](project-4-talon-robotics/) | Is the new payload ready to ship, and what is still open? | SQL Server · Power BI · DAX · Excel | ✅ Complete |
| 5 | [Meridian UAV Services: Predictive Maintenance](project-5-meridian-uav/) | If we are 99.73% compliant, why do we keep having failures? | Python · SQL Server · Streamlit · Excel | ✅ Complete |

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

## Project 4 at a glance

- **A metric that answers a different question from the one on the agenda.** The programme reports **93.86% complete**; it is **51.96% ready to ship**. Both are correct — work-item completion counts effort, and shipping depends on evidence. The 42-point gap sat invisible because only one of the two was ever reported.
- **A verification is evidence about ONE BUILD.** If a requirement passed on build 62 and its subsystem changed in build 71, the requirements tool still says Verified and is describing code that has since been replaced. Modelling a test run as a dated fact carrying a `BuildKey`, alongside a record of what each build touched, makes staleness *computable* rather than asserted — the same shape as Project 3's dated price agreements. Age is explicitly **not** the rule: intervening change is.
- **Rig contention is not a second problem, it is the cause.** A release-candidate regression re-runs everything cheap and automated, so what misses it is whatever needs the scarce resource — the hardware rig at ~74% and the airframe at ~54%, against ~90% for automated tests. Stale evidence therefore lands hardest on exactly the Safety and Regulatory requirements whose policy *demands* hardware evidence. That makes "verification is behind" and "we cannot book rig time" one item on the risk register rather than two.
- **113 requirements are verified by evidence that does not count** — 110 passing below the test level their type demands, 20 signed off by their own author against a policy requiring an independent witness. Both read as a pass on every dashboard the programme runs. Policy compliance: **77.45%**.
- **The gap sits where the design is still moving.** Flight Control (45 builds changed, **23.40%** ready) and Release Mechanism (42 builds, **32.26%**) carry 39% of all must-ship requirements. They are not behind on testing; re-running their suites buys evidence the next build invalidates.
- **A queue bounded by the constraining resource**: 226 outstanding, **610 rig hours — 3.4 rig-weeks** — with 40 schedulable in the first week, and the action chosen by state rather than rank, because a requirement whose *wording* changed needs a systems engineer before it needs a rig.

Three defects in the validation report are worth reading: a function that accepted an as-of parameter and ignored it for one metric, drawing a flat line across twenty months; a UAT harness that dropped its own results table on rollback and so reported `Invalid object name` instead of the failures, reachable only when a test failed; and a fixture that silently stopped testing anything after the generator was tuned. See [`docs/data_validation_report.md`](project-4-talon-robotics/docs/data_validation_report.md).

Start with the project [README](project-4-talon-robotics/README.md) for the reproduction steps.


## Project 5 at a glance

- **Two correct numbers in the wrong unit.** The fleet is **99.73% compliant** measured in flight hours and **68.41%** measured in stress-weighted duty cycles — the same components, the same intervals, the same day. **89 flight-critical components are past due and read as compliant on every report the operator runs.** MU-008 is at 93.16% of its flight-hour interval and **205.46%** of its stress interval.
- **An hour is not a unit of wear, and the data says so.** Highland accrues **2.15 stress hours per logged flight hour**; Coastal accrues **1.17**. The hour meter cannot tell them apart. The sting is the training base: lightest payloads, fewest hours, and the *worst* stress compliance in the fleet — circuit training produces landings rather than airborne time.
- **The finding is not planted.** Each component is given a Weibull life in stress hours and comes off by whichever clock finishes first: flight hours reaching the published interval, or stress reaching its drawn life. On a coastal airframe the clocks run together; on a mountain airframe the physics clock runs 2.4× faster. Nothing says "make Highland worse."
- **A failure alarm that passes both its targets and prevents almost nothing.** Precision **81.56%** (target 75), recall **98.26%** (target 80) — and **only 37.81% actionable lead time**. Of 283 correctly predicted failures, **176 arrived too late to order the part**. Precision and recall cannot distinguish a warning of 30 hours from one of 6.
- **The cause is one threshold across two different physics, and the fix is a number.** At the deployed setting the motor-controller alarm is **97.17% precise and 8.74% actionable**. Retuning that component alone to 2.05× gives **93.20% actionable at 75.18% precision — 87 more failures caught in time, for nothing.** Fleet-wide: 37.81% → **66.43%**.
- **The warning was tested for, not assumed away.** A model over the *same telemetry* reaches 91.11% actionable, so the information exists and the fixed threshold discards it. A model over *exposure only* never reaches the precision floor at any threshold — so the stress model is **not** a substitute for condition monitoring, and the two findings are about genuinely different things.
- **An optimisation with a test that makes it safe.** A 57-threshold sweep took over ten minutes recomputing window functions; `dbo.AlertFrontier` keeps only the points where a component's running-maximum vibration increases — 22,006 rows instead of 84,000 — and the sweep now takes **1.2 seconds**. UAT-35 recomputes the alert set the slow way and requires a match install-for-install, because a fast answer that disagrees with the slow one is worse than the slow one.

An adversarial multi-agent review of the SQL raised 51 findings — and **49 of its 55 verification agents died on a session quota**, so the run reported 50 as "refuted" when they had simply never been checked. *Unverified is not refuted.* Re-verified by hand, ten were real, including a population filtered on a today-attribute that froze an eighteen-month trend's denominator at 28, component windows that charged **1,144 sorties to two components at once**, and a feature view with no window predicate that fed **362 orphan readings** to the model. One review finding turned out to be my own misreading, and that is recorded too. See [`docs/data_validation_report.md`](project-5-meridian-uav/docs/data_validation_report.md).

Start with the project [README](project-5-meridian-uav/README.md) for the reproduction steps.
