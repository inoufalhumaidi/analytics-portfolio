# DAX Measures — Talon Robotics Payload Delivery Readiness

**Model:** `TalonDelivery` on SQL Server · **Assessed against:** build 88 (RC1), 2026-09-30

> **Data disclosure.** Talon Robotics is fictional and all data is synthetic. No confidential data
> is used and no claim is made about any production system.

`Build_PowerBI_Model.ps1` is the source of truth; this file explains it. If the two disagree, the
script is right and this document is stale — which is a defect, and one this portfolio has shipped
before.

**This page is checked, not trusted.** `Sync_Measures.ps1 -DocOnly` reads every `dax` block below
and compares it with the build script's definition, token by token: layout and `--` comments are
ignored, and every name, string, number and operator must match. It fails on a difference, on a
measure the script builds that no block defines, and on a block that defines anything the script
does not build. So each block holds exactly one measure, every measure appears exactly once — the
ones the sections below do not need to discuss are in the [Complete reference](#complete-reference)
at the end — and anything merely illustrative is written in prose, never in a `dax` block.

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File Sync_Measures.ps1 -DocOnly
```

---

## The rule that governs every threshold here

**Thresholds are never written into DAX.** Every target, warning level and direction lives in
`Ref_ReadinessTargets` and is read from that table when the measure runs. A threshold hardcoded
into a measure is one that will eventually disagree with the SQL layer, the Excel workbook and the
printed pack — and the disagreement will surface in a programme board rather than in a test.

The same applies to `Ref_VerificationPolicy`, which decides what counts as evidence. It is the most
contestable input in the analysis, so it lives where a systems engineer can find it, argue with it
and change it.

---

## One rule every filter below follows

**A filter inside `CALCULATE` replaces; `KEEPFILTERS` intersects.** Written plainly,
`CALCULATE ( COUNTROWS ( Requirement ), Requirement[Priority] = "MustShip" )` discards any filter
the report already has on `Priority` and puts its own in its place. That is how this model used to
be written, and with the Priority slicer on *ShouldShip* it set must-ship readiness (51.96%) beside
should-ship work-item completion (93.66%) as though the two described one population. Wrapped as
`KEEPFILTERS ( Requirement[Priority] = "MustShip" )`, the filter intersects with the slicer
instead: there are no must-ship requirements among the should-ship ones, so `[Must Ship]` and
`[Ship Readiness %]` are BLANK — which is the true answer.

Every `CALCULATE` filter in this model is wrapped that way, except the two kept plain on purpose
below. With nothing selected the two forms return the same number, so no headline figure moved; they part only when a report filter lands on the same column. `[Stale]` needs no wrapper: it
iterates `FILTER ( Requirement, … )`, which only ever sees the rows the report has left.

**Ratio numerators carry `+ 0`.** In `[Ship Readiness %]`, a selection that has must-ship
requirements but none of them current has a BLANK numerator. Without the `+ 0` that genuine 0%
would show as an empty cell, and the worst possible answer would look like no answer at all.
`DIVIDE` still returns BLANK when the denominator is BLANK, so BLANK keeps exactly one meaning:
nothing to measure here. Every `%` measure in the model is written this way.

**Every Colour and Status measure returns BLANK for a BLANK value.** DAX compares BLANK as 0.
Unguarded, a cell whose denominator a cross-filter had emptied still got a colour: BLANK is below the
ship-readiness warning level of 95, so it read Red, and below the stale-verification target of 15,
so that one read Green. Both were statements about requirements that were not there. The guard,
`IF ( ISBLANK ( V ) || ISBLANK ( T ) || ISBLANK ( W ), BLANK (), … )`, leaves the cell empty
instead — and because the ratios carry `+ 0`, an empty cell can only mean there was nothing to rate.

**`[Critical RAID Open] + 0` is what the Critical RAID pair rates.** A count is not a ratio. With no
critical item open, BLANK *is* the answer zero, and zero is the target, so the guard must not hide
it. The consequence is deliberate: that pair is never BLANK, and a selection with no critical item
open reads Green.

**Thresholds are read with `REMOVEFILTERS`, not `LOOKUPVALUE`.** `LOOKUPVALUE` searches the table
as the current filters leave it, so a filter on `Ref_ReadinessTargets` could blank a threshold — and,
with the guard above, the status along with it. Each threshold is read as
`CALCULATE ( VALUES ( … ), REMOVEFILTERS ( Ref_ReadinessTargets ), Ref_ReadinessTargets[MetricName] = "…" )`:
clear the table, then pick its own row.

**The deliberate exceptions.** Two filters stay plain, on purpose:

- the `MetricName` filter inside each threshold, which runs after `REMOVEFILTERS` has cleared the
  table. A threshold must not move with the report's selection, so there is nothing to intersect
  with.
- the numerator of `[Week Coverage %]`, which counts the rows that fit this week whatever a slicer
  on `Queue[IsThisWeek]` says, so that the measure stays a share of the *whole* queue; its
  denominator clears that one column for the same reason. It is explained with the measure in
  [Queue and schedule](#4-queue-and-schedule).

---

## 1. Model shape

Twelve tables, star-shaped, one direction of filter flow throughout.

| Table | Role | Source | Grain |
|---|---|---|---|
| `Dim_Date` | dimension, marked as date table | `Dim_Date` | one row per calendar date |
| `Dim_Subsystem` | dimension | `Dim_Subsystem` | one row per subsystem |
| `Dim_Person` | dimension | `Dim_Person` | one row per engineer |
| `Dim_Build` | dimension | `Dim_Build` + `Dim_Date` | one row per build |
| `Requirement` | fact | `vw_RequirementVerification` | one row per baselined requirement, verdict attached |
| `Queue` | fact | `vw_VerificationQueue` | one row per non-current requirement |
| `RAID` | fact | `vw_RAIDExposure` | one row per RAID item, scored |
| `WorkItem` | fact | `Fact_WorkItem` | one row per work item |
| `BuildChurn` | fact | `Fact_BuildSubsystemChange` | one row per build × subsystem changed |
| `Ref_ReadinessTargets`, `Ref_VerificationPolicy` | reference | tables | one row per metric / per requirement type |
| `Ref_Reporting` | reference | `vw_ReadinessKPI` | exactly one row: the reporting anchor |

### Relationships

```
Requirement[SubsystemCode]  * → 1  Dim_Subsystem[SubsystemCode]
Requirement[ReqType]        * → 1  Ref_VerificationPolicy[ReqType]
Queue[RequirementID]        * → 1  Requirement[RequirementID]
WorkItem[RequirementID]     * → 1  Requirement[RequirementID]
RAID[SubsystemCode]         * → 1  Dim_Subsystem[SubsystemCode]
RAID[OwnerID]               * → 1  Dim_Person[PersonID]
BuildChurn[BuildNumber]     * → 1  Dim_Build[BuildNumber]
BuildChurn[SubsystemCode]   * → 1  Dim_Subsystem[SubsystemCode]
Dim_Build[BuildDate]        * → 1  Dim_Date[Date]
```

All active, all single-direction. `Dim_Subsystem` feeds three facts, which is an ordinary star —
the ambiguity trap would be a relationship *between* two facts, and there is none.

**Why the facts load from views.** The verification-currency model is defined once in SQL and
asserted by 25 acceptance tests. Re-deriving staleness in DAX would create a second definition free
to drift from the first, and this project's entire argument depends on there being exactly one
definition of "ready". DAX does aggregation, ratios, time intelligence and conditional formatting —
what a semantic model is actually for.

**`Ref_Reporting`** holds one row read straight out of `vw_ReadinessKPI`, so the model and the SQL
scorecard take their as-of from the same source. Anchoring on `MAX(Dim_Build[BuildNumber])` instead
would look identical today and would silently reassess against an experimental build the moment one
appeared after the release candidate. Readiness is always about *the thing being shipped*, which is
a decision somebody made, not whatever is newest.

---

## 2. The contrast

The two headline measures are defined next to each other on purpose. Anyone editing one should have
to look at the other.

```dax
Work Item Completion % =
DIVIDE ( [Work Items Closed] + 0, [Work Items] ) * 100
```

```dax
Current Must Ship =
CALCULATE ( COUNTROWS ( Requirement ),
    KEEPFILTERS ( Requirement[Priority] = "MustShip" ), KEEPFILTERS ( Requirement[IsCurrent] = TRUE () ) )
```

```dax
Ship Readiness % =
DIVIDE ( [Current Must Ship] + 0, [Must Ship] ) * 100
```

```dax
Readiness Gap vs Reported =
VAR Reported = [Work Item Completion %]
VAR Ready = [Ship Readiness %]
RETURN IF ( NOT ISBLANK ( Reported ) && NOT ISBLANK ( Ready ), Reported - Ready )
```

**Readiness is over must-ship only.** A readiness percentage that mixes must-ship and
nice-to-have is arithmetic with no decision attached: it cannot tell you whether to ship. The
should-ship and nice populations are still measurable — `[Should Ship]` exists — but they do not
dilute the gate.

**The gap exists only where both sides do.** Under a Priority slicer that leaves no must-ship
requirement, `[Ship Readiness %]` is BLANK, and a bare subtraction would report the whole
completion figure — 93.66 under *ShouldShip* — as the gap. A gap between a number and nothing is not a
finding, so the measure returns BLANK there.

Neither number is wrong. They answer different questions, and only one of them is the question on
the agenda. Putting them adjacent, from the same model, is the entire argument.

---

## 3. Verification states

```dax
Has Evidence =
CALCULATE ( COUNTROWS ( Requirement ), KEEPFILTERS ( Requirement[HasEvidence] = TRUE () ) )
```

```dax
Meets Policy =
CALCULATE ( COUNTROWS ( Requirement ), KEEPFILTERS ( Requirement[MeetsPolicy] = TRUE () ) )
```

```dax
Current =
CALCULATE ( COUNTROWS ( Requirement ), KEEPFILTERS ( Requirement[IsCurrent] = TRUE () ) )
```

```dax
Verification Coverage % =
DIVIDE ( [Has Evidence] + 0, [Requirements] ) * 100
```

```dax
Policy Compliance % =
DIVIDE ( [Meets Policy] + 0, [Has Evidence] ) * 100
```

```dax
Stale =
COUNTROWS (
    FILTER ( Requirement,
        Requirement[MeetsPolicy] = TRUE ()
        && ( Requirement[StaleByCode] = TRUE () || Requirement[StaleByRequirement] = TRUE () ) ) )
```

```dax
Stale Verification % =
DIVIDE ( [Stale] + 0, [Meets Policy] ) * 100
```

**The denominator is requirements that HAVE policy-compliant evidence, not all requirements.**
Mixing "never verified properly" into a staleness rate makes staleness look like a documentation
problem rather than an evidence one, and the two need different work from different people.

The four counts `[No Evidence]`, `[Insufficient]`, `[Stale]` and `[Current]` partition the
population, and `Validate_PowerBI_Model.ps1` checks that they sum to `[Requirements]`. The first
two, and the two counts that say *why* evidence falls short of policy, are in the
[Complete reference](#03-verification).

### The two stale buckets are mutually exclusive, in the same order as SQL

```dax
Stale by Requirement =
CALCULATE ( COUNTROWS ( Requirement ),
    KEEPFILTERS ( Requirement[MeetsPolicy] = TRUE () ), KEEPFILTERS ( Requirement[StaleByRequirement] = TRUE () ) )
```

```dax
Stale by Code =
CALCULATE ( COUNTROWS ( Requirement ),
    KEEPFILTERS ( Requirement[MeetsPolicy] = TRUE () ),
    KEEPFILTERS ( Requirement[StaleByRequirement] = FALSE () ),   -- stale on both counts: counted by [Stale by Requirement]
    KEEPFILTERS ( Requirement[StaleByCode] = TRUE () ) )
```

A requirement stale on both counts is reported as **"requirement changed"**, because that is the
one needing a human to reread the requirement before anyone re-runs anything. Re-running first
risks certifying against the old wording.

The precedence deliberately matches `VerificationState` in `vw_RequirementVerification`, which
tests `StaleByRequirement` before `StaleByCode`. Ordering them differently in DAX would make the
two tools report different splits of the same total — and the *total* would still agree, so nothing
would look wrong.

`Validate_PowerBI_Model.ps1` asserts `[Stale by Code] + [Stale by Requirement] = [Stale]`, but that
proves only that the buckets are exclusive and leave no gap. It cannot see their order: 7 by
requirement + 88 by code and 0 + 95 both sum to 95. The risk is not hypothetical —
`vw_SubsystemReadiness` once resolved the buckets the other way round and put all 95 under
"subsystem changed" while this model said 7 and 88, and a sum that never looks at SQL could not have
noticed. What catches an ordering difference is the validator's stale-split check, which reconciles
each bucket with SQL on its own, subsystem by subsystem, against `StaleByReq` and `StaleByCode` in
`vw_SubsystemReadiness`.

---

## 4. Queue and schedule

```dax
Outstanding =
COUNTROWS ( Queue )
```

```dax
Rig Hours Outstanding =
SUM ( Queue[RigHours] )
```

```dax
Rig Weeks Outstanding =
DIVIDE ( [Rig Hours Outstanding], 180 )
```

```dax
Airframe Hours Outstanding =
SUM ( Queue[AirframeHours] )
```

```dax
Schedulable This Week =
CALCULATE ( COUNTROWS ( Queue ), KEEPFILTERS ( Queue[IsThisWeek] = TRUE () ) )
```

```dax
Week Coverage % =
DIVIDE (
    CALCULATE ( COUNTROWS ( Queue ), Queue[IsThisWeek] = TRUE () ) + 0,   -- plain on purpose: see below
    CALCULATE ( [Outstanding], REMOVEFILTERS ( Queue[IsThisWeek] ) )
) * 100
```

`Rig Weeks Outstanding` is the only figure in the model a programme board can act on without a
further study. "We are behind" is not a decision; "we are 2.4 rig-weeks behind and the gate is in
two" is.

**Each item is costed by what its action needs, in two resources that are never added.** SQL's
`fn_VerificationCost` charges the tests whose pass could make the requirement current: every test
at or above the policy level, one new test where none exists yet, nothing for a self-verified pass
that only needs a countersignature. Hardware-in-the-loop hours go to `Queue[RigHours]` (425.0) and
Field hours to `Queue[AirframeHours]` (272.0, thirty-four eight-hour flights). The old model summed
every existing test of an item into one figure, 610.0, and divided it by rig capacity: that charged
signatures as rig time, made items with no test at their level look free, and counted flights
against the rigs. `[Airframe Hours Outstanding]` is never divided by 180, and there is no
airframe-weeks measure, because the data holds no airframe capacity to divide by.

The 180 is the one hardcoded parameter in the DAX, and it is a deliberate exception: six rigs at 30
bookable hours, a fact about the building rather than a threshold anyone will argue about in a
review. It is a copy, not a link. It matches the `RigHoursPerWeek` parameter on the workbook's
Control sheet and the `180.00` that `vw_VerificationQueue` passes to `fn_VerificationQueue` to set
`Queue[IsThisWeek]`, but changing the workbook's parameter moves neither.

**`[Week Coverage %]` is the share of the whole queue that fits this week, whatever a slicer on
`IsThisWeek` says.** Written as `[Schedulable This Week]` over `[Outstanding]`, it compared this
week with itself under a slicer on *True* — the build guide's Page 3 default — and read 100.0. Its
denominator now clears that one column with `REMOVEFILTERS ( Queue[IsThisWeek] )`, and its
numerator keeps a plain filter, so that it too overrides the slicer rather than intersecting with
it. Of the model's filters on data columns, it is the only one left without `KEEPFILTERS`, and on
purpose. Every other filter the report applies, such as a subsystem, still reaches both
sides.

---

## 5. Churn

```dax
Builds Changed =
DISTINCTCOUNT ( BuildChurn[BuildNumber] )
```

```dax
Lines Changed =
SUM ( BuildChurn[LinesChanged] )
```

```dax
Churn per Readiness Point =
DIVIDE ( [Builds Changed], [Ship Readiness %] )
```

**`DISTINCTCOUNT`, not `COUNTROWS`.** `BuildChurn` is at build × subsystem grain, so counting rows
across several subsystems would count the same build repeatedly and report more builds than exist.

`Churn per Readiness Point` is the measure that changes the conversation. A subsystem that changed
forty-five times and is 23% ready is not behind on testing — it is still being designed, and that
is a different conversation with a different owner and a different remedy. Where a subsystem is 0%
ready the ratio is undefined, and `DIVIDE` returns BLANK there rather than an error.

---

## 6. RAID

```dax
RAID Open =
CALCULATE ( COUNTROWS ( RAID ), KEEPFILTERS ( RAID[IsOpen] = TRUE () ) )
```

```dax
RAID Overdue =
CALCULATE ( COUNTROWS ( RAID ), KEEPFILTERS ( RAID[IsOverdue] = TRUE () ) )
```

```dax
Overdue RAID % =
DIVIDE ( [RAID Overdue] + 0, [RAID Open] ) * 100
```

```dax
Critical RAID Open =
CALCULATE ( COUNTROWS ( RAID ),
    KEEPFILTERS ( RAID[ExposureBand] = "Critical" ), KEEPFILTERS ( RAID[IsOpen] = TRUE () ) )
```

```dax
Open RAID Exposure =
CALCULATE ( SUM ( RAID[ExposureScore] ), KEEPFILTERS ( RAID[IsOpen] = TRUE () ) )
```

`ExposureScore` and `ExposureBand` come from `Ref_RAIDMatrix` in SQL, **not** from probability ×
impact. The matrix deliberately distorts the corners — a 1×5 "unlikely but catastrophic" scores 10
while a 5×1 "certain but trivial" scores 5, though the product is identical — because that is how a
programme board actually treats them, and a formula cannot express it.

`IsOverdue` excludes items due before they were raised. Those are a data defect, reported by the
data-quality layer; counting them would inflate the overdue rate with a typing error and blame the
programme for it.

`[Critical RAID Open]` is BLANK when no critical item is open. The Critical RAID status pair adds
`+ 0` to it, so that zero reads Green rather than vanishing — see [Status colours](#7-status-colours).

---

## 7. Status colours

```dax
Ship Readiness Colour =
VAR V = [Ship Readiness %]
VAR T = CALCULATE ( VALUES ( Ref_ReadinessTargets[TargetValue] ),  REMOVEFILTERS ( Ref_ReadinessTargets ), Ref_ReadinessTargets[MetricName] = "ShipReadinessPct" )
VAR W = CALCULATE ( VALUES ( Ref_ReadinessTargets[WarningValue] ), REMOVEFILTERS ( Ref_ReadinessTargets ), Ref_ReadinessTargets[MetricName] = "ShipReadinessPct" )
VAR Dir = CALCULATE ( VALUES ( Ref_ReadinessTargets[Direction] ),  REMOVEFILTERS ( Ref_ReadinessTargets ), Ref_ReadinessTargets[MetricName] = "ShipReadinessPct" )
RETURN
    IF ( ISBLANK ( V ) || ISBLANK ( T ) || ISBLANK ( W ), BLANK (),   -- nothing to rate: no colour
         IF ( Dir = "HigherBetter",
              SWITCH ( TRUE (), V >= T, "#C6EFCE", V >= W, "#FFEB9C", "#FFC7CE" ),
              SWITCH ( TRUE (), V <= T, "#C6EFCE", V <= W, "#FFEB9C", "#FFC7CE" ) ) )
```

**A colour measure returns a hex string, not a number.** Power BI's *format by field value* expects
a colour; returning `1`/`2`/`3` and hoping the rule interprets it is the commonest reason a
conditional-format rule silently does nothing — this portfolio hit exactly that in Project 1.

**Direction is read from the table, not hardcoded.** Writing `>=` into the measure would silently
invert the status for every `LowerBetter` metric, and three of the seven here are: Stale
Verification, Critical RAID and Overdue RAID. Seven metrics get a `… Colour` and a `… Status`
pair, generated from one loop so they cannot drift apart; the Status measure is the same template
returning `"Green"`, `"Amber"` or `"Red"`.

**Every pair carries the same two guards** — BLANK in, BLANK out, and thresholds read with
`REMOVEFILTERS` — for the reasons in [One rule every filter below follows](#one-rule-every-filter-below-follows).
Every pair rates its measure as it stands except Critical RAID, which rates
`[Critical RAID Open] + 0`, because no critical item open is a genuine zero and should read Green:

```dax
Critical RAID Status =
VAR V = [Critical RAID Open] + 0
VAR T = CALCULATE ( VALUES ( Ref_ReadinessTargets[TargetValue] ),  REMOVEFILTERS ( Ref_ReadinessTargets ), Ref_ReadinessTargets[MetricName] = "CriticalRAIDOpen" )
VAR W = CALCULATE ( VALUES ( Ref_ReadinessTargets[WarningValue] ), REMOVEFILTERS ( Ref_ReadinessTargets ), Ref_ReadinessTargets[MetricName] = "CriticalRAIDOpen" )
VAR Dir = CALCULATE ( VALUES ( Ref_ReadinessTargets[Direction] ),  REMOVEFILTERS ( Ref_ReadinessTargets ), Ref_ReadinessTargets[MetricName] = "CriticalRAIDOpen" )
RETURN
    IF ( ISBLANK ( V ) || ISBLANK ( T ) || ISBLANK ( W ), BLANK (),
         IF ( Dir = "HigherBetter",
              SWITCH ( TRUE (), V >= T, "Green", V >= W, "Amber", "Red" ),
              SWITCH ( TRUE (), V <= T, "Green", V <= W, "Amber", "Red" ) ) )
```

**The ship-gate target is 100.** A must-ship requirement is by definition one you cannot ship
without, so anything below 100% is not green. That is not a stretch target, it is the definition.

The other twelve Colour and Status measures are in the [Complete reference](#07-status).

---

## 8. Measures deliberately not built

**No trend measures over `Dim_Date`.** Readiness is a function of *build*, not of date. Two builds
can share a date and a build date can be corrected; the sequence is what "superseded by" means.
`usp_ReadinessTrend` in SQL walks the build axis and the extract feeds the trend visual, rather than
a DAX time-intelligence measure that would quietly reintroduce dates as the ordering.

**No `USERELATIONSHIP` on receipt or run dates.** The verification model resolves as-of questions in
SQL, where they are tested. A DAX measure that forgot the cut-off would silently count evidence from
after the reporting date, and nothing in the report would look wrong.

**No partial-credit readiness.** A requirement is current or it is not. `InterveningBuilds` is
published and drives the queue ranking, so "how stale" is visible — but a readiness figure that
awarded fractions would not be a ship decision.

---

## 9. Verifying this document is true

Two scripts answer two different questions. Run both after any change here.

`Sync_Measures.ps1 -DocOnly` proves this document says what the build script builds: every `dax`
block token for token, and every measure exactly once. It needs no Power BI. Run without
`-DocOnly`, it also compares the `.pbix` open in Power BI Desktop with the script.

`Validate_PowerBI_Model.ps1` proves the live model's numbers agree with SQL. It queries the model in
DAX and reconciles the headline figures against the SQL that produces them, and readiness and the
stale split for every subsystem against `vw_SubsystemReadiness`. It also checks two structural
properties the totals cannot reach: the four verification states partition the population, and the
two stale buckets are exclusive.

A measure file that has drifted from the model is exactly as misleading as a model-shape section
describing a model that was never built — which is what Project 3 shipped, and had to have rewritten.

---

## Complete reference

Every measure the sections above do not show, one per block, grouped by the display folder it sits
in. With the blocks above, this is every measure the build script creates, each exactly once.

### 00 Reporting

The as-of anchor, read from `Ref_Reporting` so the model and the SQL scorecard agree on it.

```dax
As Of Build =
MAX ( Ref_Reporting[AsOfBuild] )
```

```dax
As Of Date =
MAX ( Ref_Reporting[AsOfDate] )
```

### 01 Population

```dax
Requirements =
COUNTROWS ( Requirement )
```

```dax
Must Ship =
CALCULATE ( COUNTROWS ( Requirement ), KEEPFILTERS ( Requirement[Priority] = "MustShip" ) )
```

```dax
Should Ship =
CALCULATE ( COUNTROWS ( Requirement ), KEEPFILTERS ( Requirement[Priority] = "ShouldShip" ) )
```

### 02 The contrast

Also in this folder, in [section 2](#2-the-contrast): `Work Item Completion %`,
`Current Must Ship`, `Ship Readiness %`, `Readiness Gap vs Reported`.

```dax
Work Items =
COUNTROWS ( WorkItem )
```

```dax
Work Items Closed =
CALCULATE ( COUNTROWS ( WorkItem ), KEEPFILTERS ( WorkItem[WorkItemStatus] = "Closed" ) )
```

### 03 Verification

Also in this folder, in [section 3](#3-verification-states): `Has Evidence`, `Meets Policy`,
`Current`, `Verification Coverage %`, `Policy Compliance %`, `Stale`, `Stale Verification %`,
`Stale by Requirement`, `Stale by Code`.

`No Evidence` and `Insufficient` complete the partition with `Stale` and `Current`.

```dax
No Evidence =
CALCULATE ( COUNTROWS ( Requirement ), KEEPFILTERS ( Requirement[HasEvidence] = FALSE () ) )
```

```dax
Insufficient =
CALCULATE ( COUNTROWS ( Requirement ),
    KEEPFILTERS ( Requirement[HasEvidence] = TRUE () ), KEEPFILTERS ( Requirement[MeetsPolicy] = FALSE () ) )
```

`Self Verified` and `Under Levelled` count two reasons evidence falls short of policy: a requirement
that fails policy although it has a passing run made by its own owner where its type requires an
independent tester, or a passing run below the test level its type requires. One requirement can be
counted in both.

```dax
Self Verified =
CALCULATE ( COUNTROWS ( Requirement ),
    KEEPFILTERS ( Requirement[PassingButSelfVerified] > 0 ), KEEPFILTERS ( Requirement[MeetsPolicy] = FALSE () ) )
```

```dax
Under Levelled =
CALCULATE ( COUNTROWS ( Requirement ),
    KEEPFILTERS ( Requirement[PassingButUnderLevelled] > 0 ), KEEPFILTERS ( Requirement[MeetsPolicy] = FALSE () ) )
```

### 04 Queue

Also in this folder, in [section 4](#4-queue-and-schedule): `Outstanding`,
`Rig Hours Outstanding`, `Rig Weeks Outstanding`, `Schedulable This Week`, `Week Coverage %`.

```dax
Avg Intervening Builds =
AVERAGE ( Queue[InterveningBuilds] )
```

### 05 Churn

All three measures in this folder are in [section 5](#5-churn).

### 06 RAID

Also in this folder, in [section 6](#6-raid): `RAID Open`, `RAID Overdue`, `Overdue RAID %`,
`Critical RAID Open`, `Open RAID Exposure`.

```dax
RAID Items =
COUNTROWS ( RAID )
```

### 07 Status

Also in this folder, in [section 7](#7-status-colours): `Ship Readiness Colour`,
`Critical RAID Status`. The rest follow the same template: the value, three thresholds read with
`REMOVEFILTERS`, BLANK in and BLANK out, and the direction taken from the table.

```dax
Ship Readiness Status =
VAR V = [Ship Readiness %]
VAR T = CALCULATE ( VALUES ( Ref_ReadinessTargets[TargetValue] ),  REMOVEFILTERS ( Ref_ReadinessTargets ), Ref_ReadinessTargets[MetricName] = "ShipReadinessPct" )
VAR W = CALCULATE ( VALUES ( Ref_ReadinessTargets[WarningValue] ), REMOVEFILTERS ( Ref_ReadinessTargets ), Ref_ReadinessTargets[MetricName] = "ShipReadinessPct" )
VAR Dir = CALCULATE ( VALUES ( Ref_ReadinessTargets[Direction] ),  REMOVEFILTERS ( Ref_ReadinessTargets ), Ref_ReadinessTargets[MetricName] = "ShipReadinessPct" )
RETURN
    IF ( ISBLANK ( V ) || ISBLANK ( T ) || ISBLANK ( W ), BLANK (),
         IF ( Dir = "HigherBetter",
              SWITCH ( TRUE (), V >= T, "Green", V >= W, "Amber", "Red" ),
              SWITCH ( TRUE (), V <= T, "Green", V <= W, "Amber", "Red" ) ) )
```

```dax
Verification Coverage Colour =
VAR V = [Verification Coverage %]
VAR T = CALCULATE ( VALUES ( Ref_ReadinessTargets[TargetValue] ),  REMOVEFILTERS ( Ref_ReadinessTargets ), Ref_ReadinessTargets[MetricName] = "VerificationCoveragePct" )
VAR W = CALCULATE ( VALUES ( Ref_ReadinessTargets[WarningValue] ), REMOVEFILTERS ( Ref_ReadinessTargets ), Ref_ReadinessTargets[MetricName] = "VerificationCoveragePct" )
VAR Dir = CALCULATE ( VALUES ( Ref_ReadinessTargets[Direction] ),  REMOVEFILTERS ( Ref_ReadinessTargets ), Ref_ReadinessTargets[MetricName] = "VerificationCoveragePct" )
RETURN
    IF ( ISBLANK ( V ) || ISBLANK ( T ) || ISBLANK ( W ), BLANK (),
         IF ( Dir = "HigherBetter",
              SWITCH ( TRUE (), V >= T, "#C6EFCE", V >= W, "#FFEB9C", "#FFC7CE" ),
              SWITCH ( TRUE (), V <= T, "#C6EFCE", V <= W, "#FFEB9C", "#FFC7CE" ) ) )
```

```dax
Verification Coverage Status =
VAR V = [Verification Coverage %]
VAR T = CALCULATE ( VALUES ( Ref_ReadinessTargets[TargetValue] ),  REMOVEFILTERS ( Ref_ReadinessTargets ), Ref_ReadinessTargets[MetricName] = "VerificationCoveragePct" )
VAR W = CALCULATE ( VALUES ( Ref_ReadinessTargets[WarningValue] ), REMOVEFILTERS ( Ref_ReadinessTargets ), Ref_ReadinessTargets[MetricName] = "VerificationCoveragePct" )
VAR Dir = CALCULATE ( VALUES ( Ref_ReadinessTargets[Direction] ),  REMOVEFILTERS ( Ref_ReadinessTargets ), Ref_ReadinessTargets[MetricName] = "VerificationCoveragePct" )
RETURN
    IF ( ISBLANK ( V ) || ISBLANK ( T ) || ISBLANK ( W ), BLANK (),
         IF ( Dir = "HigherBetter",
              SWITCH ( TRUE (), V >= T, "Green", V >= W, "Amber", "Red" ),
              SWITCH ( TRUE (), V <= T, "Green", V <= W, "Amber", "Red" ) ) )
```

```dax
Stale Verification Colour =
VAR V = [Stale Verification %]
VAR T = CALCULATE ( VALUES ( Ref_ReadinessTargets[TargetValue] ),  REMOVEFILTERS ( Ref_ReadinessTargets ), Ref_ReadinessTargets[MetricName] = "StaleVerificationPct" )
VAR W = CALCULATE ( VALUES ( Ref_ReadinessTargets[WarningValue] ), REMOVEFILTERS ( Ref_ReadinessTargets ), Ref_ReadinessTargets[MetricName] = "StaleVerificationPct" )
VAR Dir = CALCULATE ( VALUES ( Ref_ReadinessTargets[Direction] ),  REMOVEFILTERS ( Ref_ReadinessTargets ), Ref_ReadinessTargets[MetricName] = "StaleVerificationPct" )
RETURN
    IF ( ISBLANK ( V ) || ISBLANK ( T ) || ISBLANK ( W ), BLANK (),
         IF ( Dir = "HigherBetter",
              SWITCH ( TRUE (), V >= T, "#C6EFCE", V >= W, "#FFEB9C", "#FFC7CE" ),
              SWITCH ( TRUE (), V <= T, "#C6EFCE", V <= W, "#FFEB9C", "#FFC7CE" ) ) )
```

```dax
Stale Verification Status =
VAR V = [Stale Verification %]
VAR T = CALCULATE ( VALUES ( Ref_ReadinessTargets[TargetValue] ),  REMOVEFILTERS ( Ref_ReadinessTargets ), Ref_ReadinessTargets[MetricName] = "StaleVerificationPct" )
VAR W = CALCULATE ( VALUES ( Ref_ReadinessTargets[WarningValue] ), REMOVEFILTERS ( Ref_ReadinessTargets ), Ref_ReadinessTargets[MetricName] = "StaleVerificationPct" )
VAR Dir = CALCULATE ( VALUES ( Ref_ReadinessTargets[Direction] ),  REMOVEFILTERS ( Ref_ReadinessTargets ), Ref_ReadinessTargets[MetricName] = "StaleVerificationPct" )
RETURN
    IF ( ISBLANK ( V ) || ISBLANK ( T ) || ISBLANK ( W ), BLANK (),
         IF ( Dir = "HigherBetter",
              SWITCH ( TRUE (), V >= T, "Green", V >= W, "Amber", "Red" ),
              SWITCH ( TRUE (), V <= T, "Green", V <= W, "Amber", "Red" ) ) )
```

```dax
Policy Compliance Colour =
VAR V = [Policy Compliance %]
VAR T = CALCULATE ( VALUES ( Ref_ReadinessTargets[TargetValue] ),  REMOVEFILTERS ( Ref_ReadinessTargets ), Ref_ReadinessTargets[MetricName] = "PolicyCompliancePct" )
VAR W = CALCULATE ( VALUES ( Ref_ReadinessTargets[WarningValue] ), REMOVEFILTERS ( Ref_ReadinessTargets ), Ref_ReadinessTargets[MetricName] = "PolicyCompliancePct" )
VAR Dir = CALCULATE ( VALUES ( Ref_ReadinessTargets[Direction] ),  REMOVEFILTERS ( Ref_ReadinessTargets ), Ref_ReadinessTargets[MetricName] = "PolicyCompliancePct" )
RETURN
    IF ( ISBLANK ( V ) || ISBLANK ( T ) || ISBLANK ( W ), BLANK (),
         IF ( Dir = "HigherBetter",
              SWITCH ( TRUE (), V >= T, "#C6EFCE", V >= W, "#FFEB9C", "#FFC7CE" ),
              SWITCH ( TRUE (), V <= T, "#C6EFCE", V <= W, "#FFEB9C", "#FFC7CE" ) ) )
```

```dax
Policy Compliance Status =
VAR V = [Policy Compliance %]
VAR T = CALCULATE ( VALUES ( Ref_ReadinessTargets[TargetValue] ),  REMOVEFILTERS ( Ref_ReadinessTargets ), Ref_ReadinessTargets[MetricName] = "PolicyCompliancePct" )
VAR W = CALCULATE ( VALUES ( Ref_ReadinessTargets[WarningValue] ), REMOVEFILTERS ( Ref_ReadinessTargets ), Ref_ReadinessTargets[MetricName] = "PolicyCompliancePct" )
VAR Dir = CALCULATE ( VALUES ( Ref_ReadinessTargets[Direction] ),  REMOVEFILTERS ( Ref_ReadinessTargets ), Ref_ReadinessTargets[MetricName] = "PolicyCompliancePct" )
RETURN
    IF ( ISBLANK ( V ) || ISBLANK ( T ) || ISBLANK ( W ), BLANK (),
         IF ( Dir = "HigherBetter",
              SWITCH ( TRUE (), V >= T, "Green", V >= W, "Amber", "Red" ),
              SWITCH ( TRUE (), V <= T, "Green", V <= W, "Amber", "Red" ) ) )
```

```dax
Work Item Completion Colour =
VAR V = [Work Item Completion %]
VAR T = CALCULATE ( VALUES ( Ref_ReadinessTargets[TargetValue] ),  REMOVEFILTERS ( Ref_ReadinessTargets ), Ref_ReadinessTargets[MetricName] = "WorkItemCompletionPct" )
VAR W = CALCULATE ( VALUES ( Ref_ReadinessTargets[WarningValue] ), REMOVEFILTERS ( Ref_ReadinessTargets ), Ref_ReadinessTargets[MetricName] = "WorkItemCompletionPct" )
VAR Dir = CALCULATE ( VALUES ( Ref_ReadinessTargets[Direction] ),  REMOVEFILTERS ( Ref_ReadinessTargets ), Ref_ReadinessTargets[MetricName] = "WorkItemCompletionPct" )
RETURN
    IF ( ISBLANK ( V ) || ISBLANK ( T ) || ISBLANK ( W ), BLANK (),
         IF ( Dir = "HigherBetter",
              SWITCH ( TRUE (), V >= T, "#C6EFCE", V >= W, "#FFEB9C", "#FFC7CE" ),
              SWITCH ( TRUE (), V <= T, "#C6EFCE", V <= W, "#FFEB9C", "#FFC7CE" ) ) )
```

```dax
Work Item Completion Status =
VAR V = [Work Item Completion %]
VAR T = CALCULATE ( VALUES ( Ref_ReadinessTargets[TargetValue] ),  REMOVEFILTERS ( Ref_ReadinessTargets ), Ref_ReadinessTargets[MetricName] = "WorkItemCompletionPct" )
VAR W = CALCULATE ( VALUES ( Ref_ReadinessTargets[WarningValue] ), REMOVEFILTERS ( Ref_ReadinessTargets ), Ref_ReadinessTargets[MetricName] = "WorkItemCompletionPct" )
VAR Dir = CALCULATE ( VALUES ( Ref_ReadinessTargets[Direction] ),  REMOVEFILTERS ( Ref_ReadinessTargets ), Ref_ReadinessTargets[MetricName] = "WorkItemCompletionPct" )
RETURN
    IF ( ISBLANK ( V ) || ISBLANK ( T ) || ISBLANK ( W ), BLANK (),
         IF ( Dir = "HigherBetter",
              SWITCH ( TRUE (), V >= T, "Green", V >= W, "Amber", "Red" ),
              SWITCH ( TRUE (), V <= T, "Green", V <= W, "Amber", "Red" ) ) )
```

```dax
Critical RAID Colour =
VAR V = [Critical RAID Open] + 0
VAR T = CALCULATE ( VALUES ( Ref_ReadinessTargets[TargetValue] ),  REMOVEFILTERS ( Ref_ReadinessTargets ), Ref_ReadinessTargets[MetricName] = "CriticalRAIDOpen" )
VAR W = CALCULATE ( VALUES ( Ref_ReadinessTargets[WarningValue] ), REMOVEFILTERS ( Ref_ReadinessTargets ), Ref_ReadinessTargets[MetricName] = "CriticalRAIDOpen" )
VAR Dir = CALCULATE ( VALUES ( Ref_ReadinessTargets[Direction] ),  REMOVEFILTERS ( Ref_ReadinessTargets ), Ref_ReadinessTargets[MetricName] = "CriticalRAIDOpen" )
RETURN
    IF ( ISBLANK ( V ) || ISBLANK ( T ) || ISBLANK ( W ), BLANK (),
         IF ( Dir = "HigherBetter",
              SWITCH ( TRUE (), V >= T, "#C6EFCE", V >= W, "#FFEB9C", "#FFC7CE" ),
              SWITCH ( TRUE (), V <= T, "#C6EFCE", V <= W, "#FFEB9C", "#FFC7CE" ) ) )
```

```dax
Overdue RAID Colour =
VAR V = [Overdue RAID %]
VAR T = CALCULATE ( VALUES ( Ref_ReadinessTargets[TargetValue] ),  REMOVEFILTERS ( Ref_ReadinessTargets ), Ref_ReadinessTargets[MetricName] = "OverdueRAIDPct" )
VAR W = CALCULATE ( VALUES ( Ref_ReadinessTargets[WarningValue] ), REMOVEFILTERS ( Ref_ReadinessTargets ), Ref_ReadinessTargets[MetricName] = "OverdueRAIDPct" )
VAR Dir = CALCULATE ( VALUES ( Ref_ReadinessTargets[Direction] ),  REMOVEFILTERS ( Ref_ReadinessTargets ), Ref_ReadinessTargets[MetricName] = "OverdueRAIDPct" )
RETURN
    IF ( ISBLANK ( V ) || ISBLANK ( T ) || ISBLANK ( W ), BLANK (),
         IF ( Dir = "HigherBetter",
              SWITCH ( TRUE (), V >= T, "#C6EFCE", V >= W, "#FFEB9C", "#FFC7CE" ),
              SWITCH ( TRUE (), V <= T, "#C6EFCE", V <= W, "#FFEB9C", "#FFC7CE" ) ) )
```

```dax
Overdue RAID Status =
VAR V = [Overdue RAID %]
VAR T = CALCULATE ( VALUES ( Ref_ReadinessTargets[TargetValue] ),  REMOVEFILTERS ( Ref_ReadinessTargets ), Ref_ReadinessTargets[MetricName] = "OverdueRAIDPct" )
VAR W = CALCULATE ( VALUES ( Ref_ReadinessTargets[WarningValue] ), REMOVEFILTERS ( Ref_ReadinessTargets ), Ref_ReadinessTargets[MetricName] = "OverdueRAIDPct" )
VAR Dir = CALCULATE ( VALUES ( Ref_ReadinessTargets[Direction] ),  REMOVEFILTERS ( Ref_ReadinessTargets ), Ref_ReadinessTargets[MetricName] = "OverdueRAIDPct" )
RETURN
    IF ( ISBLANK ( V ) || ISBLANK ( T ) || ISBLANK ( W ), BLANK (),
         IF ( Dir = "HigherBetter",
              SWITCH ( TRUE (), V >= T, "Green", V >= W, "Amber", "Red" ),
              SWITCH ( TRUE (), V <= T, "Green", V <= W, "Amber", "Red" ) ) )
```
