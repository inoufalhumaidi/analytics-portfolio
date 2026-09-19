# Data Validation Report — Talon Robotics

## Payload Deployment Delivery

**Purpose:** record how the synthetic programme dataset was generated, what behaviour was
deliberately built into it, what was tested, what broke, and what was deliberately retained.

> **Data disclosure.** Talon Robotics is a fictional mechatronics company. Every row was generated
> by `sql/02_generate_synthetic_data.sql`. No confidential, proprietary or production data is used
> at any point, and no claim is made about any real programme.

---

## 1. What "validated" means here

Three things, in order of how much they are worth:

1. **The model answers the question it claims to answer.** Readiness is must-ship requirements
   whose evidence is current *and* meets the verification policy. Every acceptance test that can
   recompute a figure does so from the raw facts, at a different grain from the implementation.
2. **Two independent implementations agree.** Every headline figure is computed in T-SQL and again
   in Excel formulas, and reconciled on the workbook's own Validation sheet.
3. **The planted behaviour is present.** A generator that quietly failed to plant a pattern would
   produce a clean dataset and a finding about nothing — which happened in Project 3 of this
   portfolio, where an `UPDATE` matched zero rows and nobody noticed for a week.

Agreement between tools is necessary, not sufficient. Project 1 shipped an Excel workbook and a
SQL view that agreed exactly and were both wrong, because both were derived from the same flawed
written definition. That caveat is repeated in the workbook and in the Power BI validator, because
it is the one that is easiest to forget while looking at a column of green.

---

## 2. Reproducibility

Every random draw comes from `dbo.fn_Rand(<stable text key>)`, which hashes the key with SHA2_256
and scales it into `[0,1)`. It is a pure function of the key: dropping and regenerating the database
reproduces the dataset byte for byte, so every figure quoted in the case study stays checkable by a
reader who runs the scripts.

`UAT-22` asserts the fingerprints. A rebuild that produced different data would fail the suite
rather than silently publishing different numbers under the same headings.

| Fingerprint | Value |
|---|---:|
| Test runs | `8806` |
| RAID register | `135` |
| Work items | `7942` |

**Volumes:** 540 requirements (519 baselined, 21 withdrawn), 1,354 test cases, 3,011 test runs,
88 builds, 196 build-subsystem changes, 107 requirement changes, 1,091 work items, 96 RAID items.

---

## 3. What was deliberately built into the data

### 3a. The behaviour the analysis is meant to surface

| | Mechanism | Signal it creates |
|---|---|---|
| **Late churn** | Release Mechanism and Flight Control change in most builds after 61; everything else largely stabilises | Their verifications go stale in bulk — 23% and 32% ready against 59–83% elsewhere |
| **Effort completes, evidence does not** | Work items close steadily to ~94% | The reported figure and the ship gate diverge by 42 points |
| **Policy breaches** | Some Safety requirements covered only at Integration level; some verified by their own owner | 110 requirements "verified" by evidence that does not count |
| **Rig contention** | HIL and Field tests are far less likely to make the release-candidate regression | The *cause* of staleness, not a separate finding |
| **RAID concentration** | Items on the churning subsystems stay open far more often | Exposure and overdue both cluster where the change is |
| **Cohort effect** | Safety and Regulatory requirements demand hardware evidence | They are the worst-covered group, while the aggregate looks tolerable |

### 3b. The causal spine, stated explicitly

The generator does **not** plant staleness directly. It plants a release-candidate regression, and
makes participation in that regression depend on test level:

| Test level | In the RC regression | Why |
|---|---:|---|
| Unit | 90% | Automated, free, runs every time |
| Integration | 88% | Automated, cheap |
| HIL | 74% | Rig-limited |
| Field | 54% | Airframe-limited |

…multiplied by **0.42** for Release Mechanism and Flight Control, whose teams spend the endgame
changing code rather than testing it.

Staleness then *emerges*: anything re-run on the release candidate is current by definition, and
what does not get re-run is whatever needed the scarce resource. This matters for the analysis, not
just the generator — it means rig contention and stale evidence are one problem, and it is why the
recommended action for most of the queue is "book rig time", not "write more tests".

### 3c. Planted data defects, found by behaviour

Each is created by behaviour, never by a marker in an identifier, so the data-quality layer has to
find it the way it would have to in real data.

| Defect | Planted | Found | Detector |
|---|---:|---:|---|
| Duplicate test runs | 22 | 22 | Same case, build, date and result — the business signature |
| Test run dated before its build | 14 | 14 | Run date earlier than build date |
| Work item closed before opened | 9 | 9 | Close date earlier than open date |
| RAID due before raised | 7 | 7 | Due date earlier than raised date |
| Runs against withdrawn requirements | 16 | 16 | Requirement status, not an ID prefix |

---

## 4. Defects found during the build

### 4a. A function that took an as-of parameter and ignored it

**Symptom:** `usp_ReadinessTrend` drew a **flat line**. `WorkItemCompletionPct` read 93.86% at
every build from 46 to 88 — across twenty months of a programme that was demonstrably closing work
items throughout.

**Cause:** `fn_ReadinessKPI` accepts `@AsOfDate` and used it for every metric except work items,
which were counted with no date filter at all.

This is worse than a function that never offered the parameter. A caller passing an as-of date is
entitled to assume the whole answer respects it, and the one metric that ignored it was the metric
the entire comparison rests on.

**Fix:** an item not yet opened does not exist; an item closed after the as-of date was still open
at the as-of date. Completion now climbs 87.6% → 93.9% across the programme.

**Regression cover:** `UAT-05` asserts completion at an early date is strictly lower than at the
reporting date. Against the old code the two figures are equal and it fails.

### 4b. A UAT harness that hid its own failures

**Symptom:** a failing suite reported `Invalid object name '#UATResults'` instead of the failures.

**Cause:** `#UATResults` is created inside the test transaction, so `ROLLBACK` drops it — and the
failure-detail `SELECT` ran *after* the rollback.

The branch only executes when a case fails, which is exactly why a suite that had always passed
would never reveal it. The first real failure would have been masked by an unrelated error message
pointing at the harness rather than the defect.

**Fix:** the failure detail is selected before the rollback; only the `THROW` comes after. The
other three projects were checked — Project 1 reads its counts before its rollback, Projects 2 and
3 use no transaction, so none shared the bug.

### 4c. A fixture that stopped testing anything

**Symptom:** after the regression rates were tuned, `UAT-06` returned `NULL` rather than failing.

**Cause:** the fixture selected "a requirement that is Current with evidence predating the release
candidate". After tuning, **every** Current requirement drew its evidence from the RC regression at
build 88 — anything verified earlier has had its subsystem move since — so the selection matched
nothing and the assertion compared two nulls.

A fixture that depends on the shape of the generated data is one parameter change away from testing
nothing, and it fails silently rather than loudly.

**Fix:** the fixture now creates its own build. Build 89 changes no subsystem, so the only thing
that can invalidate a verification at build 89 is a requirement change — which is precisely the
rule under test, with every other cause held out by construction.

### 4d. A workbook that reported success while holding `#N/A`

**Symptom:** `Build_Workbook.ps1` completed cleanly; `Validate_Workbook.ps1` found `Calc!B4 = #N/A`.

**Cause:** while removing an unused code block from the build script, the line that populates the
Calc sheet's rank column went with it. `MATCH` then had nothing to look up.

Writing a formula never evaluates it, so a build script reports that it wrote text, not that the
text means anything. This is the third time in this portfolio that reopening a workbook caught
something the build could not.

**Also caught in the same pass:** the validator read the schedule figures from fixed cells
(`B33`/`B34`), but that block sits below the subsystem table and moves whenever a subsystem is
added. Those cells are now named ranges and read by name — the same fix as Project 1's Gross Margin
row-offset defect, in a different file.

### 4e. A second as-of leak, found by reviewing for the first one

**Symptom:** none at the reporting date. `fn_RAIDExposure` reported 45 open items and 45 were
genuinely open, so every published RAID figure was correct.

Run at **2025-06-30** it reported **12** open when **18** were open.

**Cause:** `IsOpen` was derived from the stored `RAIDStatus` — today's status — rather than from the
dates. An item closed in March 2026 shows as closed in a register run for June 2025, when it was
demonstrably open at the time. `IsOverdue` had the same fault, since it also tested `RAIDStatus`.

It is harmless at the reporting date only because every close in this dataset predates it, which is
exactly why it survived the first pass. The function still advertises `@AsOfDate` in its signature
and `usp_RAIDRegister(@AsOfDate)` still offers it to a caller, and both were answering about today.

**This is the same defect as §4a**, in a different function, found because §4a was written up and
the review then went looking for other functions that accept a point-in-time parameter and ignore
it. That is the argument for writing defects up rather than just fixing them: the class is more
useful than the instance.

**Fix:** `IsOpen` is now `ClosedDate IS NULL OR ClosedDate > @AsOfDate`, `IsOverdue` uses the same
test, `ClosedDate` is suppressed when the close has not happened yet, and a new `RAIDStatusAsOf`
column reconstructs the status at the as-of date. An item closed *after* the as-of date was open
then; the stored status cannot say whether it was `Open` or `Mitigating`, so it reports `Open` —
the conservative reading, and the one that does not invent a mitigation nobody had started.

**Regression cover:** `UAT-23` asserts the open count at **2025-06-30**, not at the reporting date.
Run at the reporting date the case passes against the broken code too, and would be a test that
cannot fail for the reason it was written.

**No published figure changed.** Critical RAID open 9, overdue 57.78%, exposure 466 — all identical
before and after.

---

---

## 5. Results at the release candidate (build 88)

### The comparison the project exists to make

| | |
|---|---:|
| Work items closed | **93.86%** |
| Requirements with evidence | 96.53% |
| **Ship readiness** | **51.96%** |
| The gap | **41.90 points** |

### Why the gap exists

| Verification state | Requirements | Of which must-ship |
|---|---:|---:|
| Current | 293 | 146 |
| Stale — subsystem changed | 88 | 54 |
| Stale — requirement changed | 7 | 3 |
| Insufficient evidence | 113 | 70 |
| No evidence | 18 | 8 |
| **Total baselined** | **519** | **281** |

### Where the gap is

| Subsystem | Criticality | Must-ship | Current | Readiness | Builds changed |
|---|---|---:|---:|---:|---:|
| Flight Control Interface | Safety | 47 | 11 | **23.40%** | 45 |
| Payload Release Mechanism | Safety | 62 | 20 | **32.26%** | 42 |
| Power Management | Mission | 37 | 22 | 59.46% | 15 |
| Command & Telemetry Link | Mission | 43 | 26 | 60.47% | 23 |
| Structural Interface | Safety | 19 | 13 | 68.42% | 15 |
| Ground Control Software | Support | 28 | 20 | 71.43% | 21 |
| Navigation & Sensor Fusion | Mission | 33 | 24 | 72.73% | 17 |
| Diagnostics & Logging | Support | 12 | 10 | 83.33% | 18 |

The two Safety subsystems carry **39% of all must-ship requirements** and the most change. The
correlation is the finding: they are not behind on testing, they are still being designed.

### The schedule answer

226 requirements outstanding, **610 rig hours**, **3.4 rig-weeks** at 180 bookable hours. 40 are
schedulable in the first week.

---

## 6. Data-quality results

The gate measures the share of **must-ship** requirements whose readiness answer rests on data a
check has flagged — not a count of rows. A count of defects says how untidy the data is; it does
not say whether to trust the answer.

**It passes at 1.779% against a 2.000% tolerance** — by a fifth of a percentage point. That is a
margin, not comfort, and this report says so rather than reporting a green light.

| Check | Found | Impact class | Counts toward the gate |
|---|---:|---|:--:|
| `UNDER_LEVELLED_VERIFICATION` | 110 | Policy | — |
| `SELF_VERIFIED_SAFETY` | 20 | Policy | — |
| `RUN_BEFORE_BUILD` | 14 | Timing | ✔ |
| `WORKITEM_CLOSED_BEFORE_OPEN` | 9 | Timing | — |
| `BLOCKED_AT_RELEASE_CANDIDATE` | 45 | Evidence | — |
| `DUPLICATE_TEST_RUN` | 22 | Evidence | — |
| `RUN_ON_WITHDRAWN_REQ` | 16 | Waste | — |
| `RAID_DUE_BEFORE_RAISED` | 7 | Timing | — |
| `ORPHAN_DIMENSION_KEY` | 0 | Integrity | ✔ |
| `REQ_WITHOUT_TESTCASE` | 0 | Coverage | ✔ |
| `TESTCASE_NEVER_RUN` | 0 | Coverage | — |

**Three checks found nothing, and they are listed anyway.** A summary that reports only what was
found cannot distinguish "we looked and it was clean" from "we never looked", and those are very
different things to put in front of a review board.

**Policy breaches are not data defects.** A Safety requirement verified by its own owner is
accurate data recording a practice that should not have happened. Its `ImpactClass` is `Policy` and
it is excluded from the gate; the KPI layer holds it to account instead. Conflating "the record is
wrong" with "the practice is wrong" makes both harder to fix, and lets the gate be dismissed as
noise.

---

## 7. Defects deliberately retained

**The 110 under-levelled verifications stay.** They are the second-largest finding in the
programme, and removing them would remove the evidence for it. They are reported, attributed and
excluded from readiness — which is what should happen to them.

**The seven RAID items due before they were raised stay.** They are overdue on the day they are
created, and they demonstrate why `IsOverdue` explicitly excludes them: counting them would inflate
the overdue rate with a typing error and blame the programme for it. `UAT-20` asserts both halves —
that they exist, and that they are not counted.

---

## 8. Known limitations

**One release candidate.** Readiness is assessed against build 88. The model is parameterised by
build and `usp_ReadinessTrend` walks it backwards, but there is no second candidate to compare
against, so "would build 74 have been a better ship point" is answerable and unanswered.

**Test levels are ordinal, not a lattice.** `Unit < Integration < HIL < Field` assumes a field test
subsumes a rig test. On a real programme some field evidence does not substitute for instrumented
rig data, and the policy table cannot currently express that.

**Staleness is binary.** A subsystem that changed once and one that changed eleven times both mark
a verification stale. `InterveningBuilds` is published and drives the queue ranking, but readiness
itself does not distinguish them. That is deliberate — a partial-credit readiness figure is not a
ship decision — but it is a simplification worth naming.

**No cost or effort model.** Rig hours bound the queue, but there is no budget, no labour cost and
no critical path. "Three point four rig-weeks" is a capacity statement, not a schedule forecast.

---

## 9. How to reproduce this report

```bash
sqlcmd -S <server> -E -d master        -i sql/01_create_schema.sql
sqlcmd -S <server> -E -d TalonDelivery -i sql/02_generate_synthetic_data.sql
sqlcmd -S <server> -E -d TalonDelivery -i sql/03_core_views.sql
sqlcmd -S <server> -E -d TalonDelivery -i sql/04_data_quality_checks.sql
sqlcmd -S <server> -E -d TalonDelivery -i sql/05_kpi_views.sql
sqlcmd -S <server> -E -d TalonDelivery -i sql/06_stored_procedures.sql
sqlcmd -S <server> -E -d TalonDelivery -i sql/07_uat_test_cases.sql     # expect 23/23

powershell erp_extracts/Export_Extracts.ps1
powershell excel/Build_Workbook.ps1
powershell excel/Validate_Workbook.ps1                                  # expect 16/16
```

Every figure in this report comes from those scripts. The generator is deterministic, so a rebuild
reproduces them exactly, and `UAT-22` asserts the fingerprints that guarantee it.
