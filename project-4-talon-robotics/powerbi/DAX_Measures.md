# DAX Measures — Talon Robotics Payload Delivery Readiness

**Model:** `TalonDelivery` on SQL Server · **Assessed against:** build 88 (RC1), 2026-09-30

> **Data disclosure.** Talon Robotics is fictional and all data is synthetic. No confidential data
> is used and no claim is made about any production system.

`Build_PowerBI_Model.ps1` is the source of truth; this file explains it. If the two disagree, the
script is right and this document is stale — which is a defect, and one this portfolio has shipped
before.

---

## The rule that governs every measure here

**Thresholds are never written into DAX.** Every target, warning level and direction lives in
`Ref_ReadinessTargets` and is read with `LOOKUPVALUE`. A threshold hardcoded into a measure is one
that will eventually disagree with the SQL layer, the Excel workbook and the printed pack — and the
disagreement will surface in a programme board rather than in a test.

The same applies to `Ref_VerificationPolicy`, which decides what counts as evidence. It is the most
contestable input in the analysis, so it lives where a systems engineer can find it, argue with it
and change it.

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
asserted by 23 acceptance tests. Re-deriving staleness in DAX would create a second definition free
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
DIVIDE ( [Work Items Closed], [Work Items] ) * 100

Current Must Ship =
CALCULATE ( COUNTROWS ( Requirement ),
    Requirement[Priority] = "MustShip", Requirement[IsCurrent] = TRUE () )

Ship Readiness % =
DIVIDE ( [Current Must Ship], [Must Ship] ) * 100

Readiness Gap vs Reported = [Work Item Completion %] - [Ship Readiness %]
```

**Readiness is over must-ship only.** A readiness percentage that mixes must-ship and
nice-to-have is arithmetic with no decision attached: it cannot tell you whether to ship. The
should-ship and nice populations are still measurable — `[Should Ship]` exists — but they do not
dilute the gate.

Neither number is wrong. They answer different questions, and only one of them is the question on
the agenda. Putting them adjacent, from the same model, is the entire argument.

---

## 3. Verification states

```dax
Has Evidence = CALCULATE ( COUNTROWS ( Requirement ), Requirement[HasEvidence] = TRUE () )
Meets Policy = CALCULATE ( COUNTROWS ( Requirement ), Requirement[MeetsPolicy] = TRUE () )
Current      = CALCULATE ( COUNTROWS ( Requirement ), Requirement[IsCurrent] = TRUE () )

Verification Coverage % = DIVIDE ( [Has Evidence], [Requirements] ) * 100
Policy Compliance %     = DIVIDE ( [Meets Policy], [Has Evidence] ) * 100
```

```dax
Stale =
COUNTROWS (
    FILTER ( Requirement,
        Requirement[MeetsPolicy] = TRUE ()
        && ( Requirement[StaleByCode] = TRUE () || Requirement[StaleByRequirement] = TRUE () ) ) )

Stale Verification % = DIVIDE ( [Stale], [Meets Policy] ) * 100
```

**The denominator is requirements that HAVE policy-compliant evidence, not all requirements.**
Mixing "never verified properly" into a staleness rate makes staleness look like a documentation
problem rather than an evidence one, and the two need different work from different people.

### The two stale buckets are mutually exclusive, in the same order as SQL

```dax
Stale by Requirement =
CALCULATE ( COUNTROWS ( Requirement ),
    Requirement[MeetsPolicy] = TRUE (), Requirement[StaleByRequirement] = TRUE () )

Stale by Code =
CALCULATE ( COUNTROWS ( Requirement ),
    Requirement[MeetsPolicy] = TRUE (),
    Requirement[StaleByRequirement] = FALSE (), Requirement[StaleByCode] = TRUE () )
```

A requirement stale on both counts is reported as **"requirement changed"**, because that is the
one needing a human to reread the requirement before anyone re-runs anything. Re-running first
risks certifying against the old wording.

The precedence deliberately matches `VerificationState` in the SQL view. Ordering them differently
in DAX would make the two tools report different splits of the same total — and the *total* would
still agree, so nothing would look wrong. `Validate_PowerBI_Model.ps1` asserts
`[Stale by Code] + [Stale by Requirement] = [Stale]` for exactly that reason.

---

## 4. Queue and schedule

```dax
Outstanding           = COUNTROWS ( Queue )
Rig Hours Outstanding = SUM ( Queue[RigHours] )
Rig Weeks Outstanding = DIVIDE ( [Rig Hours Outstanding], 180 )
Schedulable This Week = CALCULATE ( COUNTROWS ( Queue ), Queue[IsThisWeek] = TRUE () )
```

`Rig Weeks Outstanding` is the only figure in the model a programme board can act on without a
further study. "We are behind" is not a decision; "we are 3.4 rig-weeks behind and the gate is in
two" is.

The 180 is the one hardcoded number in the model, and it is a deliberate exception: it mirrors the
`RigHoursPerWeek` parameter on the workbook's Control sheet, and the rig count is a fact about the
building rather than a threshold anyone will argue about in a review.

---

## 5. Churn

```dax
Builds Changed = DISTINCTCOUNT ( BuildChurn[BuildNumber] )
Lines Changed  = SUM ( BuildChurn[LinesChanged] )

Churn per Readiness Point = DIVIDE ( [Builds Changed], [Ship Readiness %] )
```

**`DISTINCTCOUNT`, not `COUNTROWS`.** `BuildChurn` is at build × subsystem grain, so counting rows
across several subsystems would count the same build repeatedly and report more builds than exist.

`Churn per Readiness Point` is the measure that changes the conversation. A subsystem that changed
forty-five times and is 23% ready is not behind on testing — it is still being designed, and that
is a different conversation with a different owner and a different remedy.

---

## 6. RAID

```dax
RAID Open          = CALCULATE ( COUNTROWS ( RAID ), RAID[IsOpen] = TRUE () )
RAID Overdue       = CALCULATE ( COUNTROWS ( RAID ), RAID[IsOverdue] = TRUE () )
Overdue RAID %     = DIVIDE ( [RAID Overdue], [RAID Open] ) * 100
Critical RAID Open = CALCULATE ( COUNTROWS ( RAID ), RAID[ExposureBand] = "Critical", RAID[IsOpen] = TRUE () )
Open RAID Exposure = CALCULATE ( SUM ( RAID[ExposureScore] ), RAID[IsOpen] = TRUE () )
```

`ExposureScore` and `ExposureBand` come from `Ref_RAIDMatrix` in SQL, **not** from probability ×
impact. The matrix deliberately distorts the corners — a 1×5 "unlikely but catastrophic" scores 10
while a 5×1 "certain but trivial" scores 5, though the product is identical — because that is how a
programme board actually treats them, and a formula cannot express it.

`IsOverdue` excludes items due before they were raised. Those are a data defect, reported by the
data-quality layer; counting them would inflate the overdue rate with a typing error and blame the
programme for it.

---

## 7. Status colours

```dax
Ship Readiness Colour =
VAR V = [Ship Readiness %]
VAR T = LOOKUPVALUE ( Ref_ReadinessTargets[TargetValue],  Ref_ReadinessTargets[MetricName], "ShipReadinessPct" )
VAR W = LOOKUPVALUE ( Ref_ReadinessTargets[WarningValue], Ref_ReadinessTargets[MetricName], "ShipReadinessPct" )
VAR Dir = LOOKUPVALUE ( Ref_ReadinessTargets[Direction],  Ref_ReadinessTargets[MetricName], "ShipReadinessPct" )
RETURN
    IF ( Dir = "HigherBetter",
         SWITCH ( TRUE (), V >= T, "#C6EFCE", V >= W, "#FFEB9C", "#FFC7CE" ),
         SWITCH ( TRUE (), V <= T, "#C6EFCE", V <= W, "#FFEB9C", "#FFC7CE" ) )
```

**A colour measure returns a hex string, not a number.** Power BI's *format by field value* expects
a colour; returning `1`/`2`/`3` and hoping the rule interprets it is the commonest reason a
conditional-format rule silently does nothing — this portfolio hit exactly that in Project 1.

**Direction is read from the table, not hardcoded.** Writing `>=` into the measure would silently
invert the status for every `LowerBetter` metric, and two of the seven here are. Seven metrics get a
`… Colour` and a `… Status` pair, generated from one loop so they cannot drift apart.

**The ship-gate target is 100.** A must-ship requirement is by definition one you cannot ship
without, so anything below 100% is red. That is not a stretch target, it is the definition.

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

`Validate_PowerBI_Model.ps1` queries the live model in DAX and reconciles **28 figures** against the
SQL that produces them, including the state partition, the exclusivity of the two stale buckets, and
readiness for all eight subsystems individually. Run it after any change here.

A measure file that has drifted from the model is exactly as misleading as a model-shape section
describing a model that was never built — which is what Project 3 shipped, and had to have rewritten.
