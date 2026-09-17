# Requirements Traceability Matrix
## Vantage Wholesale Supply — Receivables Performance

**Reporting date:** 2025-12-31 · **Status:** all 24 requirements implemented, tested and accepted

> **Data disclosure.** Vantage Wholesale Supply is a fictional company. All data is synthetic.
> No confidential data is used and no claim is made about any production system.

---

## How to read this

Each row traces one business requirement from the question a stakeholder asked, through the object
that implements it, to the test that proves it works, to the figure it produced. A requirement
with no test is an intention; a test with no requirement is trivia. Both columns are filled for
every row, and the **Evidence** column carries the actual value at the reporting date so a
reviewer can check the claim rather than take it.

**Priority:** `M` must-have · `S` should-have · `C` could-have.

---

## 1. Business requirements

| ID | Requirement | Asked by | Pri | Implemented in | Verified by | Evidence at 2025-12-31 |
|---|---|---|:--:|---|---|---|
| **BR-01** | Report a single, defensible open balance per invoice that every other figure derives from | Controller | M | `fn_ARBalance(@AsOf)` — the settlement identity | UAT-01 | Open AR **$10,847,081.49** across 2,346 invoices |
| **BR-02** | Age the book as at any chosen date, not just today | Controller | M | `@AsOf` on every function; `vw_*` wrappers pin the reporting date | UAT-02 | Queue rebuilt at 2025-06-30, 09-30, 12-31 |
| **BR-03** | Report DSO on a basis that survives seasonal sales swings | CFO | M | `fn_DSO` — countback (exhaustion) method | UAT-05, UAT-11 | **63.07 days** (classic reads 61.59) |
| **BR-04** | Separate the DSO days Vantage granted from the days customers took | CFO / Collections | M | `fn_DSOBridge(@AsOf)` — exact three-way decomposition | UAT-04 | Granted **45.97** + Dispute **0.88** + Lateness **14.74** = **61.59** |
| **BR-05** | Say whether a terms policy changed, or whether long-terms customers simply bought more | Sales director | S | `fn_WATShiftShare(...)` — shift-share with visible interaction | UAT-06 | ΔWAT **+2.28 d** = mix −0.33, **rate +1.71**, interaction +0.90 |
| **BR-06** | State whether the DSO target is achievable given the terms actually sold | CFO | M | `fn_DSOBridge.WeightedAvgTermsDays`; Dashboard `H16` | Workbook validation | Target **45.00** vs WAT **46.68** → allowance **−1.68 days** |
| **BR-07** | Measure collections performance without rewarding write-offs | Collections | M | `fn_ARKPI` — `CEI_Book`, `CEI_Cash`, `PaperCollectionsGap` | UAT-13 | CEI book **66.91%**, cash **65.26%**, gap **1.65 pts** |
| **BR-08** | Produce a ranked worklist answering "who do we call first" | Collections | M | `fn_PriorityActionQueue(@AsOf)` | UAT-08, 09, 10 | **272** accounts ranked; **65** on today's worklist |
| **BR-09** | Bound the worklist to what six collectors can work in a day | Collections | M | `IsTodaysWorklist`, `CollectorRank`, `CallsPerCollectorPerDay` | UAT-10 | 10–12 per collector, **81.3%** of collectable exposure |
| **BR-10** | Tell each collector what to do, not just who to ring | Collections | M | `ActionCode` + `RecommendedAction` ladder | UAT-09 | 7 action types; every row carries an instruction |
| **BR-11** | Route disputed balances away from collections | Collections | S | Dispute discount in `CollectableExposure`; `RESOLVE_DISPUTE` | UAT-08 | **45** accounts / $195,694 routed to billing |
| **BR-12** | Flag accounts at or over credit limit without hijacking the collections action | Credit manager | S | `CreditHoldFlag` as a separate bit | UAT-09 | 3 accounts flagged; highest utilisation **106.10%** |
| **BR-13** | Detect receivables mis-statement and refuse to publish when material | Controller / Audit | M | `usp_RunDataQualityChecks` with `@MaxMisstatementPctOfAR` | UAT-17 | Gate **FAILS** at **3.421%** vs 1.000% tolerance |
| **BR-14** | Quantify mis-statement without double-counting cause and effect | Controller | M | `ImpactClass` + `CountsTowardExposure` in `vw_DQ_CheckCatalog` | §5 validation report | **$371,039.99** (not $954k as a naive sum gave) |
| **BR-15** | Give collections an operational control that works on any desk | Collections | M | `excel/Vantage_AR_Control.xlsx` — formula-driven, no dynamic arrays | `Validate_Workbook.ps1` | 27/27 reconcile; 0 formula errors |
| **BR-16** | Let a reviewer verify the numbers without a database | Reviewer | S | `erp_extracts/*.csv` + Power Query reading a folder from a cell | Workbook opens standalone | 4,279 rows across 11 extracts |
| **BR-17** | Make every published figure re-checkable | Reviewer / Audit | M | Hash-keyed generator; `CHECKSUM_AGG` fingerprints | UAT-16 | All 5 fingerprints identical across regeneration |
| **BR-18** | Report ageing by value **and** by count | Controller | S | `PctPastDue`/`…ByCount`, `Pct90Plus`/`…ByCount` | Workbook validation | 90+ **3.35%** by value, **5.33%** by count |
| **BR-19** | Provide a reusable query interface for any consumer | Data team | S | 8 stored procedures, `@AsOf`-parameterised, no dynamic SQL | UAT-14 | All 8 execute; invalid arguments raise |
| **BR-20** | Exclude the ledger ramp-up from trend conclusions | Controller | S | `IsComparablePeriod` on the monthly views | §8 validation report | 4 months excluded, **19** comparable |
| **BR-21** | Show days-beyond-terms without survivorship bias | Collections | C | `fn_DBT(@AsOf, @Months)` — settled and at-risk variants | UAT-25 | Settled **10.56 d** vs at-risk **12.41 d**; survivorship gap widened from **0.62 d** (2024) to **1.85 d** (2025) |
| **BR-22** | **Track unapplied cash sitting against paid accounts** | **Cash application** | **M** | **`Fact_CashReceipt` + `Fact_CashApplication`; `fn_UnappliedCash(@AsOf)`; `usp_CashApplicationWorklist`; `APPLY_CASH` routing** | **UAT-18, 19, 20, 24** | **$419,392.13 across 97 accounts; 57 receipts wholly unapplied; oldest 679 days; 51 accounts routed out of the call list** |
| **BR-23** | **Measure billing lag from despatch to invoice** | **Operations** | **S** | **`Fact_Invoice.ShipDateKey`; `fn_BillingLag`; `vw_BillingLagMonthly`; `usp_BillingLagReport`; `CashCycleDays`** | **UAT-21** | **1.78 days overall; Southeast 3.22 vs 1.49–1.51 elsewhere; true cash cycle 63.37 days** |
| **BR-24** | **Report promise-to-pay kept rate** | **Collections** | **S** | **`Fact_PromiseToPay`; `fn_PromiseStatus`; `fn_PromiseKeptRate`; `usp_PromiseReport`; broken-promise escalation in the queue** | **UAT-22, 23** | **86.29% by value overall; High-risk tier 53.87%; 244 broken promises worth $625,642** |

---

## 2. What the three newly-met requirements changed

They were originally deferred as schema limitations. Closing them did not simply add three
metrics — one of them changed a headline conclusion.

| ID | What it added | What it changed |
|---|---|---|
| **BR-22** | Unapplied cash as a first-class category | **2.38 days of DSO moved out of "customers are paying late" and into "we have not applied our own cash".** The granted-versus-taken framing the project was built to answer turned out to need a third term. |
| **BR-23** | Billing lag, and the true cash cycle | DSO was understating the order-to-cash cycle by 1.78 days. Lag extends DSO rather than decomposing it, so the bridge identity is preserved (UAT-21). |
| **BR-24** | Promise kept rate, derived from cash | Surfaced the cohort pattern: a portfolio figure of 86.29% sitting on top of a High-risk tier at 53.87%. Also became an escalation trigger in the queue. |

### Schema changes required

| Change | Reason |
|---|---|
| `Fact_Payment` split into `Fact_CashReceipt` + `Fact_CashApplication` | A receipt and its application are different events. Holding them in one row makes unapplied cash inexpressible. |
| `Fact_Invoice.ShipDateKey` added | Billing lag is `InvoiceDate − ShipDate`. No constraint forbids a negative lag: pre-billing is a revenue-recognition question, raised by `PRE_BILLING` rather than clipped to zero. |
| `Fact_PromiseToPay` added, with **no** status column | Status is derived from cash in `fn_PromiseStatus`. A stored outcome would be written by the collectors the metric measures. UAT-22 asserts the column does not exist. |
| `Dim_Date` extended back to 2023-11-01 | Ship dates precede invoice dates. A date dimension must cover every date any fact can reference, or the first billing lag inserted violates a foreign key. |
| `Ref_ARTargets` gained 3 rows | `UnappliedCashPct` (0.50 / 1.50), `BillingLagDays` (2.00 / 4.00), `PromiseKeptRate` (80.00 / 65.00) |

---

## 3. Data-quality requirements

| ID | Control | Implemented in | Entity | Severity | Result |
|---|---|---|---|---|---|
| **DQ-01** | No receipt banked twice for the same customer, date and amount | `DUPLICATE_RECEIPT` | Receipt | High | **96 found** (planted; retained to prove the detector fires) |
| **DQ-02** | No invoice may carry applied cash exceeding its value | `OVER_APPLIED_CASH` | Invoice | High | **101 found** — traced entirely to DQ-01 |
| **DQ-03** | No receipt may be dated before the invoice it settles | `RECEIPT_BEFORE_INVOICE` | Receipt | High | **30 found** ($100,496.95) |
| **DQ-04** | Every fact row must resolve to a dimension row | `ORPHAN_DIMENSION_KEY` | Invoice | High | 0 |
| **DQ-05** | Adjustments may not exceed what is owed | `ADJUSTMENT_EXCEEDS_INVOICE` | Invoice | High | 0 *(5 found and fixed during the build)* |
| **DQ-06** | Due date must equal invoice date plus the terms net days | `DUE_DATE_TERMS_MISMATCH` | Invoice | Medium | **52 found** ($12,727.40) |
| **DQ-07** | A settlement discount may only be taken where terms grant one | `DISCOUNT_WITHOUT_TERMS` | Application | Medium | 0 |
| **DQ-08** | Dispute flag, reason and dates must agree | `DISPUTE_FLAG_INCONSISTENT` | Invoice | Medium | 0 |
| **DQ-09** | No receipt may be dated after the reporting date | `RECEIPT_AFTER_ASOF` | Receipt | Medium | 0 |
| **DQ-10** | **Cash may not be applied before it reaches the bank** | `APPLICATION_BEFORE_RECEIPT` | Application | High | 0 |
| **DQ-11** | **No receipt may sit unmatched beyond 90 days** | `UNAPPLIED_CASH_AGED` | Receipt | High | **65 found** ($265,487.11) |
| **DQ-12** | **No invoice may be issued before the goods shipped** | `PRE_BILLING` | Invoice | Medium | **20 found** ($76,841.74) |

**Gate:** `usp_RunDataQualityChecks @MaxAcceptableRatePct = 1.000, @MaxMisstatementPctOfAR = 1.000`
→ **FAIL** at 3.421% of open AR.

---

## 4. Acceptance criteria and UAT coverage

All **25** cases pass. Definitions in `sql/08_uat_test_cases.sql`.

| Test | Requirement | Asserts | Result |
|---|---|---|---|
| UAT-01 | BR-01 | Settlement identity holds for all 29,971 invoices | PASS — 0 drift |
| UAT-02 | BR-02 | As-of balance counts only cash applied by then, to invoices raised by then | PASS — exact |
| UAT-03 | BR-01, 18 | Buckets exhaustive and mutually exclusive | PASS — 0 / 0 |
| UAT-04 | BR-04 | Bridge sums to classic DSO | PASS — residual 0.0000 |
| UAT-05 | BR-03 | Two objects publishing classic DSO agree | PASS — 0.0000 |
| UAT-06 | BR-05 | Mix + rate + interaction = ΔWAT | PASS — 0.0000 |
| UAT-07 | BR-01 | Customer rollup ties to invoice grain | PASS — exact |
| UAT-08 | BR-08, 11 | Exposure within [0, WeightedExposure] | PASS — 0 / 0 |
| UAT-09 | BR-10, 12 | One known action and an instruction per row | PASS — 0 bad |
| UAT-10 | BR-09 | Capacity respected except escalations | PASS — 0 over |
| UAT-11 | BR-03 | ADD uses countback both sides, and matches its components | PASS — 0.0000 |
| UAT-12 | BR-01, DQ-02 | Over-application surfaced, not netted | PASS — 0 / 0 |
| UAT-13 | BR-07 | CEI book ≥ CEI cash where write-offs exist | PASS — 0 months |
| UAT-14 | BR-19 | Procedures raise on invalid input | PASS / PASS |
| UAT-15 | §3.2 validation report | Instalment defect stays fixed | PASS — 1 / 0 |
| UAT-16 | BR-17 | All five fingerprints reproduce | PASS |
| UAT-17 | BR-13, DQ-01 | Planted defects still detected | PASS — 96 / 101 |
| **UAT-18** | **BR-22, DQ-11** | Unapplied cash non-negative, floored per receipt | PASS — 0 / $419,392.13 |
| **UAT-19** | **BR-22** | NetExposure = OpenBalance − UnappliedCash | PASS — 0 mismatches |
| **UAT-20** | **BR-22, BR-08** | No `APPLY_CASH` account on a worklist | PASS — 0 of 51 |
| **UAT-21** | **BR-23** | CashCycle = BillingLag + classic DSO | PASS — 0.0000 |
| **UAT-22** | **BR-24** | No stored outcome column; no Kept without cash | PASS — 0 / 0 |
| **UAT-23** | **BR-24** | Unripe promise is Outstanding, not Broken | PASS — 0 |
| **UAT-24** | **DQ-10** | No application precedes its receipt | PASS — 0 |
| **UAT-25** | **BR-21** | At-risk DBT is at least settled DBT while past-due exists | PASS — 12.41 ≥ 10.56 |

### Workbook acceptance

`excel/Validate_Workbook.ps1` opens the saved file, forces a full rebuild and reads values back.
Separate from the build on purpose: a build that finishes without throwing has proved that it
ran, not that it produced a working file.

| Criterion | Result |
|---|---|
| No formula resolves to an Excel error | PASS — 0 across all non-data sheets |
| Excel figures match SQL | PASS — 17/17 spot checks within tolerance |
| Workbook's own reconciliation sheet | PASS — 27/27 |
| DSO bridge identity holds in Excel | PASS |
| Priority Queue returns ranked rows under its filters | PASS — 25 rows |
| No personal information in the saved file | PASS — creator, last-modified-by and company empty |

Three defects were caught by this script that the build reported as successful: sheets
referencing each other before they existed (`#REF!` throughout), a layout collision that returned
a *plausible* wrong Best Possible DSO, and an `AvgDaysDelinquent` that disagreed with its own
published components.

---

## 5. Non-functional requirements

| ID | Requirement | How met | Evidence |
|---|---|---|---|
| **NFR-01** | Any consumer can query without knowing the internals | Eight `@AsOf`-parameterised procedures | `sql/07` |
| **NFR-02** | No SQL injection surface | `@GroupBy` resolved by `CASE`, never concatenated | `usp_AgingSummary` raises on `'DROP TABLE'` (UAT-14) |
| **NFR-03** | Thresholds defined once | `Ref_ARTargets` read by SQL and by Excel `INDEX/MATCH` | Changing a target changes the RAG status in both |
| **NFR-04** | Workbook opens on Excel 2019 and LibreOffice | No `FILTER`/`SORT`/`UNIQUE`; `SUMPRODUCT` + `INDEX`/`MATCH` | `excel/Build_Workbook.ps1` header |
| **NFR-05** | Workbook survives relocation | Power Query resolves its folder from a named cell | `SourceFolder` on Control |
| **NFR-06** | Scripts fail loudly | Creation gates `THROW`; UAT `THROW`s; validator exits non-zero | `sql/03`–`08` |
| **NFR-07** | No personal or confidential data | Fictional company, synthetic data, personal info stripped on save | Disclosure on every artefact |
| **NFR-08** | Derived figures agree with their published components | One rounding rule applied in SQL and Excel alike | UAT-05, UAT-11 |

---

## 6. Coverage summary

| Category | Total | Met | Coverage |
|---|---:|---:|---:|
| Business requirements | 24 | 24 | **100%** |
| Must-have (`M`) | 13 | 13 | **100%** |
| Data-quality controls | 12 | 12 | **100%** |
| Non-functional | 8 | 8 | **100%** |
| UAT cases | 25 | 25 pass | **100%** |
| Workbook acceptance | 6 | 6 pass | **100%** |

Every requirement is met and tested. The three that were previously deferred — unapplied cash,
billing lag and promise-to-pay — required the schema change recorded in §2, and closing them
changed a headline conclusion rather than merely adding metrics.
