# Requirements Traceability Matrix — Talon Robotics

## Payload Deployment Delivery

Every requirement traced to the code that implements it, the test that proves it, and the evidence
that test produced. Where a row's evidence is prose rather than a runnable test, it says so — that
is a gap, not a pass, and naming it is the point of the document.

> **Data disclosure.** Talon Robotics is fictional and all data is synthetic. No confidential data
> is used and no claim is made about any production system.

**Source of truth.** This file. The `Traceability` view of the Power BI report and the summary in
the case study both derive from it; if they disagree, this is right.

---

## Business requirements

| ID | Requirement | Stakeholder | Pri | Implemented by | Verified by | Evidence | Status |
|---|---|---|:--:|---|---|---|:--:|
| **BR-01** | Answer whether the payload can ship, not how much work is done | Programme board | M | `fn_ReadinessKPI.ShipReadinessPct` over must-ship only | UAT-15 | Recomputed at requirement grain: 51.96 = 51.96 | PASS |
| **BR-02** | Distinguish evidence that EXISTS from evidence that still APPLIES | Chief engineer | M | `fn_RequirementVerification` — `HasEvidence` vs `IsCurrent`; `Fact_TestRun.BuildKey` + `Fact_BuildSubsystemChange` | UAT-02 | `IsCurrent` recomputed from raw facts for all 519 requirements: 0 mismatches | PASS |
| **BR-03** | Verification must meet the evidence grade its requirement type demands | Quality | M | `Ref_VerificationPolicy.MinTestLevelRank`; `MeetsPolicy` | UAT-07 | 0 requirements current without a test case at or above the required level | PASS |
| **BR-04** | Safety and regulatory requirements need an independent tester | Quality / certification | M | `Ref_VerificationPolicy.RequiresIndependentTester`; `MeetsIndependence` | UAT-08 | Fixture: Safety requirement passed at the right level **by its own owner** is still not current | PASS |
| **BR-05** | A withdrawn requirement is out of scope, not unverified | Systems engineering | S | `WHERE ReqStatus = 'Baselined'` in `fn_RequirementVerification` | UAT-09 | 0 withdrawn requirements in the model; 0 in the queue | PASS |
| **BR-06** | Readiness must be answerable at any build, not only today | Programme board | M | `@AsOfBuild` on every function; `usp_ReadinessTrend` | UAT-04, UAT-05 | Build 40 leaks no later run (0/0); work-item completion differs early vs late (84.6 < 93.9) | PASS |
| **BR-07** | Produce a ranked worklist bounded by the constraining resource | Test manager | M | `fn_VerificationQueue`, `CumulativeRigHours`, `IsThisWeek` | UAT-10, UAT-13 | Queue holds exactly the 226 non-current requirements; nothing in the week exceeds 180 rig hours; doubling the rig widens the week 40 → 60 | PASS |
| **BR-08** | Name the action, not just the gap | Test manager | M | `ActionCode` + `RecommendedAction` (5 codes) | UAT-11, UAT-12 | Every item has a recognised code and an instruction; 0 action codes disagree with the requirement's state | PASS |
| **BR-09** | Quantify the remaining work in the constraining resource | Programme board | M | `RigHoursOutstanding`, `RigWeeksOutstanding` | Workbook check 12 | 610.0 rig hours = 3.4 rig-weeks; Excel and SQL agree | PASS |
| **BR-10** | Score RAID from an agreed matrix, not a multiplication | Programme board | S | `Ref_RAIDMatrix`; `fn_RAIDExposure` joins it | UAT-19 | 0 mismatches against the matrix, and 1×5 ≠ 5×1 confirms it is not a product | PASS |
| **BR-11** | An item due before it was raised is a data defect, not an overdue item | Programme board | S | `IsOverdue` excludes `DueDate < RaisedDate` | UAT-20 | 7 such items exist; 0 counted overdue | PASS |
| **BR-23** | The RAID register must answer for the date asked, not for today | Programme board | M | `fn_RAIDExposure` derives `IsOpen` from the dates; `RAIDStatusAsOf` reconstructs status | UAT-23 | At 2025-06-30: 18 reported open, 18 genuinely open; 0 overdue-but-closed | PASS |
| **BR-12** | Detect data-quality defects by behaviour, never by a marker | Data owner | M | `fn_DuplicateTestRuns` matches the business signature | UAT-16 | Planted duplicate found from its signature alone, with no prefix to read | PASS |
| **BR-13** | Separate policy breaches from data defects | Quality | M | `ImpactClass` + `CountsTowardExposure` declared in `vw_DQ_CheckCatalog` | UAT-18 | 0 Policy-class checks contribute to the gate | PASS |
| **BR-14** | The data-quality gate must be able to stop the build | Release engineering | M | `usp_RunDataQualityChecks` raises 52030 | Exit-code check | Passes at default (exit 0); raises and exits 1 at a 1.000% tolerance | PASS |
| **BR-15** | The dataset must rebuild byte for byte | Reviewer | M | `dbo.fn_Rand` — SHA2_256 of a stable key | UAT-22 | Fingerprints 8806 / 135 reproduce on a rebuild | PASS |
| **BR-16** | Every figure reproducible outside the BI tool | Analyst | M | 7 `@AsOf`-aware procedures in `sql/06_stored_procedures.sql` | UAT-21 | Procedures raise on an unknown build, a zero rig capacity and a reversed range | PASS |
| **BR-17** | An Excel control, not a static report | Test manager | M | `excel/Build_Workbook.ps1` — every dashboard figure a formula over the extracts | `Validate_Workbook.ps1` | 16 of 16 reconcile; no error values anywhere in the file | PASS |
| **BR-18** | Headline figures implemented twice and reconciled | Reviewer | M | T-SQL and Excel formulas, independently | Validation sheet | 16/16 agree to 0.01; both sides of the headline comparison computed in Excel, not copied | PASS |
| **BR-19** | Publish the reported metric ALONGSIDE the real one | Programme board | M | `WorkItemCompletionPct` and `ShipReadinessPct` in the same view and adjacent on the dashboard | Workbook checks 3, 7 | 93.86% and 51.96% shown together; gap 41.90 points | PASS |
| **BR-20** | Show where the gap is, not just that it exists | Chief engineer | M | `fn_SubsystemReadiness` | UAT-14 | Subsystem cut sums to the programme totals: 146/281 both ways | PASS |
| **BR-21** | The verification states must partition the population | Reviewer | S | `VerificationState` in `vw_RequirementVerification` | UAT-01 | 519 requirements, 519 classified, no leftovers | PASS |
| **BR-22** | Latest evidence must be deterministic when runs tie | Reviewer | S | `fn_LatestRunPerCase` orders by build, then date, then key | UAT-03 | Fixture plants two runs on the same case, build and date; exactly one is returned | PASS |

---

## Requirements explicitly NOT in scope

Recorded so that a reader can tell a deliberate boundary from an oversight.

| ID | Requirement | Why it is out of scope |
|---|---|---|
| NS-01 | Cost and effort forecasting | Rig hours bound the queue, but there is no budget, labour rate or critical path. "3.4 rig-weeks" is a capacity statement, not a schedule forecast. |
| NS-02 | Partial credit for staleness | A subsystem that changed once and one that changed eleven times both mark a verification stale. `InterveningBuilds` drives the ranking, but readiness stays binary — a partial-credit ship gate is not a ship decision. |
| NS-03 | Test-level substitution rules | `Unit < Integration < HIL < Field` assumes field evidence subsumes rig evidence. On a real programme some instrumented rig data cannot be replaced by a flight, and the policy table cannot express that. |
| NS-04 | Requirement dependency graph | Requirements are treated as independent. A real programme has parent/child decomposition where verifying a child contributes to a parent. |
| NS-05 | Multiple release candidates | One candidate, build 88. The model is parameterised by build, so "would build 74 have shipped better" is answerable — but it is not answered here. |
| NS-06 | Per-receipt reject attribution | Test runs record Pass/Fail/Blocked, not a failure taxonomy. Root-cause analysis of failures is a different project. |

---

## Test case coverage

| Test | Area | What it asserts | Can it fail? |
|---|---|---|---|
| UAT-01 | Verification model | States partition the population | Yes — a new state, or an unclassified requirement |
| UAT-02 | Verification model | `IsCurrent` recomputed from raw facts | Yes — any error in assembling the three flags |
| UAT-03 | Evidence selection | One latest run per case, deterministic tie-break | Yes — fixture plants a genuine tie |
| UAT-04 | As-of integrity | No later evidence leaks into an earlier assessment | Yes |
| UAT-05 | As-of integrity | Work-item completion respects its as-of date | Yes — **fails against the pre-fix code** |
| UAT-06 | Staleness rules | Clarified does not invalidate; Modified does | Yes — self-contained fixture build |
| UAT-07 | Verification policy | No requirement current below its required test level | Yes |
| UAT-08 | Verification policy | Independence enforced even at the right level | Yes — fixture is at the right level and fails only on the author |
| UAT-09 | Population | Withdrawn requirements excluded everywhere | Yes |
| UAT-10 | Priority queue | Queue holds exactly the non-current requirements | Yes |
| UAT-11 | Priority queue | Every item has a code and an instruction | Yes |
| UAT-12 | Priority queue | Action codes match the recomputed state | Yes |
| UAT-13 | Priority queue | Rig bound respected, and it responds to capacity | Yes — the second half is what makes it more than a row count |
| UAT-14 | Aggregation | Subsystem cut sums to the programme totals | Yes |
| UAT-15 | Scorecard | Ship readiness recomputed at requirement grain | Yes |
| UAT-16 | Data quality | Duplicates found by business signature | Yes |
| UAT-17 | Data quality | Every catalogued check reported, including clean ones | Yes |
| UAT-18 | Data quality | Policy breaches excluded from the gate | Yes |
| UAT-19 | RAID | Exposure read from the matrix, which is not a product | Yes |
| UAT-20 | RAID | Due-before-raised not counted as overdue | Yes |
| UAT-21 | Interface | Procedures raise on invalid arguments | Yes |
| UAT-22 | Reproducibility | Fingerprints reproduce | Yes |
| UAT-23 | RAID | Openness is as-of correct at a HISTORICAL date | Yes — **fails against the pre-fix code** (12 vs 18) |

**23 of 23 pass.** The suite asserts an expected case count before reporting and raises if the two
disagree — a run that dies partway through otherwise prints a pass count covering only the cases
that executed, which is indistinguishable from a clean run of a shorter suite.

**On the "can it fail?" column.** Three projects in this portfolio shipped defects past a green
suite, each because a test asserted a *property* of the answer rather than its *value* — "the index
is at least 100" is true under any positive weighting, including the join fan-out it was written to
catch. Every case above was written or rewritten against that standard, and three of them
(UAT-03, UAT-06, UAT-08) plant a fixture specifically designed to make the code fail if the rule
under test were removed.

---

## Traceability gaps

Stated rather than quietly omitted:

- **BR-09 and BR-17/18 are verified by the workbook validator, not by the SQL UAT suite.** That is
  appropriate — they are claims about Excel — but it means a reader running only `sql/07` has not
  checked them. The README's run sequence includes both for that reason.
- **BR-14 is verified by an exit-code check, not by a UAT case.** A test inside the suite cannot
  usefully assert that a *different* script returns a non-zero exit code; it is checked by running
  the gate at a tightened tolerance and reading `$?`. The command is in the validation report.
- **No requirement covers the Power BI layer's correctness.** `Validate_PowerBI_Model.ps1`
  reconciles 28 figures against SQL, but no BR row demands it, so it is currently an
  over-delivery rather than a traced requirement.
