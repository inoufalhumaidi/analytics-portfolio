# Requirements Traceability Matrix
## Lumen Optics Manufacturing — Photonics Spend Scorecard

**Reporting date:** 2025-12-31 · **Status:** all 20 requirements implemented, tested and accepted

> **Data disclosure.** Lumen Optics Manufacturing is a fictional company. All data is synthetic.
> No confidential data is used and no claim is made about any production system.

---

## How to read this

Each row traces one business requirement from the question a stakeholder asked, through the object
that implements it, to the test that proves it works, to the figure it produced. A requirement
with no test is an intention; a test with no requirement is trivia. Both columns are filled for
every row, and the **Evidence** column carries the actual value so a reviewer can check the claim
rather than take it.

**Priority:** `M` must-have · `S` should-have · `C` could-have.

---

## 1. Business requirements

| ID | Requirement | Asked by | Pri | Implemented in | Verified by | Evidence at 2025-12-31 |
|---|---|---|:--:|---|---|---|
| **BR-01** | Report one defensible landed cost per purchase line that every figure derives from | Controller | M | `fn_POLineCost(@AsOf)` — the landed cost identity | UAT-01 | Landed cost **$49,474,785.92** over 12 months; identity drift 0 |
| **BR-02** | Report spend as at any chosen date, not just today | Controller | M | `@AsOf` on every function; `vw_*` wrappers pin the reporting date | UAT-02 | Mid-year snapshot ties exactly to a direct count |
| **BR-03** | Measure price variance against the agreement actually in force at the order date | CPO | M | `fn_ContractedPrice`; dated `Fact_PriceAgreement` | UAT-03, UAT-04 | PPV **−0.19%** of contracted spend |
| **BR-04** | Show what the learning curve should have delivered and what was actually captured | CPO | M | `fn_PriceErosion(@AsOf,…)`; `Ref_PriceErosionBenchmark` | UAT-05, UAT-06, UAT-07 | Erosion capture **29.50%** against an 80% target |
| **BR-05** | Quantify the annual value of the erosion gap | CFO | M | `AnnualOpportunity` in `fn_PriceErosion` | UAT-21 | **$4,047,609** a year, gross |
| **BR-06** | Identify spend with no agreement in force | Sourcing ops | M | `IsOnContract` in `fn_POLineCost` | UAT-03 | Maverick spend **27.09%** against a 5% target |
| **BR-07** | Compare suppliers on cost per USABLE unit, not unit price | Category managers | M | `CostPerAcceptedUnit`; `CostIndexVsBest` in `fn_VendorScorecard` | UAT-12 | VEN-004 cheapest per unit on LOM-0072 and **dearest per accepted unit** |
| **BR-08** | Produce a ranked worklist of which supplier-part pair to renegotiate first | CPO | M | `fn_RenegotiationQueue(@AsOf, @PerBuyer)` | UAT-08, 09, 10 | **185** pairs ranked; **43** workable this quarter |
| **BR-09** | Decide the ACTION by leverage, not only by value | Category managers | M | `LeverageScore` + the `ActionCode` ladder | UAT-08 | 5 action types; sole-source routed to `DUAL_SOURCE` |
| **BR-10** | Bound the worklist to what a category manager can actually run | CPO | M | `IsThisQuarter`, `BuyerRank`, `@NegotiationsPerBuyer` | UAT-10 | 43 pairs covering **65.6%** of the gap |
| **BR-11** | Assign each negotiation to the person who actually buys the part | Sourcing ops | S | Spend-weighted `PrimaryBuyer` in the queue | UAT-11 | 0 mismatches across 185 pairs |
| **BR-12** | Say plainly where there is no leverage and no case for creating any | Category managers | S | `ACCEPT` action with a written reason | UAT-09 | 73 pairs / $276,461, largest single $30,142 |
| **BR-13** | Separate quality loss from price loss | Supplier quality | S | `FIX_QUALITY` action; `RejectedValue`; reject attribution | UAT-08 | **$1,571,168** of material paid for and unusable |
| **BR-14** | Detect procurement data defects and refuse to publish when material | Controller / Audit | M | `usp_RunDataQualityChecks` with `@MaxMisstatementPctOfSpend` | UAT-15, 17 | Gate **PASSES at 0.497%** against a 0.500% tolerance |
| **BR-15** | Quantify mis-statement without double-counting cause and effect | Controller | M | `ImpactClass` + `CountsTowardExposure` in `vw_DQ_CheckCatalog` | UAT-16 | $722,990 counted once; defect count 114 → 94 |
| **BR-16** | Give sourcing an operational control that works on any desk | CPO | M | `excel/Lumen_Spend_Scorecard.xlsx` — formula-driven, no dynamic arrays | `Validate_Workbook.ps1` | 14/14 reconcile; 0 formula errors |
| **BR-17** | Let a reviewer verify the numbers without a database | Reviewer | S | `erp_extracts/*.csv` + Power Query resolving its folder from a cell | Workbook opens standalone | 6,803 rows across 11 extracts |
| **BR-18** | Make every published figure re-checkable | Reviewer / Audit | M | Hash-keyed generator; `CHECKSUM_AGG` fingerprints | UAT-20 | All three fingerprints identical across regeneration |
| **BR-19** | Provide a reusable query interface for any consumer | Data team | S | 7 stored procedures, `@AsOf`-parameterised, no dynamic SQL | UAT-14 | All 7 execute; invalid arguments raise |
| **BR-20** | Publish the same logic in Power BI without re-deriving it | BI team | S | `powerbi/DAX_Measures.md` — thresholds read from `Ref_SpendTargets` | Weights match SQL by inspection | Leverage weights identical in both implementations |

---

## 2. Data-quality requirements

| ID | Control | Implemented in | Entity | Severity | Result |
|---|---|---|---|---|---|
| **DQ-01** | No requisition raised twice under two PO numbers | `DUPLICATE_PO_LINE` | POLine | High | **26 found** (planted; retained to prove the detector fires) |
| **DQ-02** | Only one price agreement may be in force per vendor-part per date | `AMBIGUOUS_CONTRACT` | POLine | High | **61 lines** from 6 overlapping agreements |
| **DQ-03** | A unit price may not be a multiple of the agreement in force | `PRICE_OUTLIER` | POLine | High | **3 found** ($705,310) |
| **DQ-04** | Goods may not be booked in before the order was raised | `RECEIPT_BEFORE_ORDER` | Receipt | High | **4 found** ($74,253) |
| **DQ-05** | Cumulative receipts may not exceed the quantity ordered | `RECEIPT_EXCEEDS_ORDER` | Receipt | High | 0 |
| **DQ-06** | Every fact row must resolve to a dimension row | `ORPHAN_DIMENSION_KEY` | POLine | High | 0 |
| **DQ-07** | An agreement may not end before it starts | `AGREEMENT_INVERTED` | Agreement | Medium | 0 |
| **DQ-08** | Expedite fee and expedite flag must agree | `EXPEDITE_FLAG_MISMATCH` | POLine | Medium | 0 |
| **DQ-09** | No order may be dated after the reporting date | `ORDER_AFTER_ASOF` | POLine | Medium | 0 |
| **DQ-10** | An order over 180 days old must have a receipt | `NEVER_RECEIVED` | POLine | Medium | 0 *(excludes duplicates — see §3.5 of the validation report)* |

**Gate:** `usp_RunDataQualityChecks @MaxAcceptableRatePct = 1.000, @MaxMisstatementPctOfSpend = 0.500`
→ **PASS at 0.497%**, by three thousandths of a percentage point. Reported as a margin, not a
green light.

---

## 3. Acceptance criteria and UAT coverage

All **21** cases pass. Definitions in `sql/07_uat_test_cases.sql`.

| Test | Requirement | Asserts | Result |
|---|---|---|---|
| UAT-01 | BR-01 | Landed cost identity on all 6,327 lines | PASS — 0 drift |
| UAT-02 | BR-02 | As-of snapshot ties to a direct count | PASS — exact |
| UAT-03 | BR-03, BR-06 | Open-ended agreements still in force | PASS — 1,333 lines |
| UAT-04 | BR-03 | Off-contract lines carry NULL variance | PASS — 0 |
| UAT-05 | BR-04 | Expected price matches the benchmark curve | PASS — < 0.01 |
| UAT-06 | BR-04 | Erosion windows cannot overlap | PASS — 0 |
| UAT-07 | BR-04 | Capture agrees with its own components | PASS — < 0.01 |
| UAT-08 | BR-08, BR-09, BR-13 | One known action and an instruction per pair | PASS — 0 bad |
| UAT-09 | BR-12 | Nothing ≥ $100k parked as `ACCEPT` | PASS — max $30,142 |
| UAT-10 | BR-10 | Capacity respected | PASS — 0 buyers over |
| UAT-11 | BR-11 | Buyer assignment follows the spend | PASS — 0 mismatches |
| UAT-12 | BR-07 | Cost index never below 100 | PASS — 0 |
| UAT-13 | BR-01 | Vendor rollup ties to line level | PASS — exact |
| UAT-14 | BR-19 | Procedures raise on invalid input | PASS / PASS |
| UAT-15 | BR-14, DQ-01 | Duplicate detector fires by behaviour | PASS — 26 / 26 |
| UAT-16 | BR-15 | No defect reported twice | PASS — 0 |
| UAT-17 | DQ-02 | Overlapping agreements detected | PASS — 61 lines |
| UAT-18 | BR-03 | Contract resolution is deterministic | PASS — 0 |
| UAT-19 | BR-07 | Accepted + rejected = received | PASS — 0 |
| UAT-20 | BR-18 | Fingerprints reproduce | PASS — all three |
| UAT-21 | BR-05 | Opportunity never negative; offset disclosed | PASS — 0 / 19 |

### Workbook acceptance

`excel/Validate_Workbook.ps1` opens the saved file, forces a full rebuild and reads values back.
Separate from the build on purpose: a build that finishes without throwing has proved it ran, not
that it produced a working file.

| Criterion | Result |
|---|---|
| No formula resolves to an Excel error | PASS — 0 across all non-data sheets |
| Excel figures match SQL | PASS — 12/12 spot checks |
| Workbook's own reconciliation sheet | PASS — 14/14 |
| Renegotiation Queue returns ranked rows under its filters | PASS — 25 rows |
| No personal information in the saved file | PASS |

Two defects were caught by this script that the build reported as successful: a top-five
concentration computed over the wrong time window, and a stale expectation left over from the
rounding fix.

---

## 4. Non-functional requirements

| ID | Requirement | How met | Evidence |
|---|---|---|---|
| **NFR-01** | Any consumer can query without knowing the internals | Seven `@AsOf`-parameterised procedures | `sql/06` |
| **NFR-02** | No SQL injection surface | `@GroupBy` resolved by `CASE`, never concatenated | `usp_CategorySummary` raises on `'DROP TABLE'` (UAT-14) |
| **NFR-03** | Thresholds defined once | `Ref_SpendTargets` read by SQL, DAX `LOOKUPVALUE` and Excel `INDEX/MATCH` | Changing a target changes the status in all three |
| **NFR-04** | The benchmark assumption is visible and challengeable | `Ref_PriceErosionBenchmark` with a `SourceNote` per row | 8 categories, each with provenance |
| **NFR-05** | Workbook opens on Excel 2019 and LibreOffice | No `FILTER`/`SORT`/`UNIQUE` | `excel/Build_Workbook.ps1` header |
| **NFR-06** | Workbook survives relocation | Power Query resolves its folder from a named cell | `SourceFolder` on Control |
| **NFR-07** | Scripts fail loudly | Creation gates `THROW`; UAT `THROW`s; validator exits non-zero | `sql/03`–`07` |
| **NFR-08** | Derived figures agree with their published components | One rounding rule applied in SQL, DAX and Excel | UAT-05, UAT-07 |
| **NFR-09** | No personal or confidential data | Fictional company, synthetic data, personal info stripped on save | Disclosure on every artefact |

---

## 5. Coverage summary

| Category | Total | Met | Coverage |
|---|---:|---:|---:|
| Business requirements | 20 | 20 | **100%** |
| Must-have (`M`) | 12 | 12 | **100%** |
| Data-quality controls | 10 | 10 | **100%** |
| Non-functional | 9 | 9 | **100%** |
| UAT cases | 21 | 21 pass | **100%** |
| Workbook acceptance | 5 | 5 pass | **100%** |

---

## 6. Deliberately out of scope

Recorded rather than quietly omitted, because a model that silently lacks a measure invites
someone to compute it badly in a visual.

| Not built | Why | What it would take |
|---|---|---|
| Should-cost variance | The erosion model measures failure to decline from the first observed price; it cannot say whether that price was fair. | A bottom-up build-up from materials, process time, yield and overhead per part. |
| Supplier risk score | Financial health, single-site exposure and geography are not in this model. A score built from delivery and quality alone would be named for something it does not measure. | Supplier financial data and site-level master data. |
| Savings realised | Requires tracking negotiated outcomes back into subsequent prices. | A negotiation register, and one quarter of the queue actually worked. It is the measure that would prove the exercise paid for itself. |
