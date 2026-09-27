# Ridgeline Field Services — DAX Measure Library

All measures below assume the star schema from `sql/01_create_schema.sql` has been imported into Power BI as-is (table names unchanged: `Fact_ServiceJobs`, `Dim_Date`, `Dim_Technician`, `Dim_Client`, `Dim_ServiceType`, `Dim_Region`, `Ref_KPITargets`), with relationships built exactly as described in `POWER_BI_BUILD_GUIDE.md`. Create a dedicated measures table called `_Measures` (Modeling → New Table → `_Measures = ROW("x", 0)`, then delete the `x` column) and paste every measure below into it — keeping measures out of the data tables is standard practice and keeps the field list clean. *(If you built the model with `Build_PowerBI_Model.ps1`, skip this: the script places all 31 measures on `Fact_ServiceJobs`, grouped into Volume / Quality / Utilization / Financial / Targets / Status / Formatting display folders.)*

**This page is checked, not trusted.** `Build_PowerBI_Model.ps1` is the source of truth. `Sync_Measures.ps1 -DocOnly` compares every code block below with it and fails on any difference, missing measure or extra one. That check exists because this page once drifted: it kept the defective `[Worked Minutes]` and was missing five measures while the script moved on.

Business question these measures answer: **Where is technician capacity being wasted, and which dispatch decisions should change this week?**

---

## One rule every filter below follows

> **A filter inside `CALCULATE` replaces; `KEEPFILTERS` intersects.** `CALCULATE ( [Total Jobs], Fact_ServiceJobs[JobStatus] = "Completed" )` discards any filter the report already has on `JobStatus` and puts its own in its place. With a slicer on *Cancelled*, `[Completed Jobs]` would ignore the slicer and return every completed job, and `[Completion Rate %]` would divide that by the cancelled count and go past 100%. Under a weekends-only filter, `[Worked Minutes]` would return the *weekday* total. Wrapping the filter in `KEEPFILTERS` makes it intersect with the report's selection instead: completed jobs *among the cancelled ones* is none, which is the true answer.
>
> Every measure that filters `JobStatus` or `IsWeekend` uses it. The **Target** measures deliberately do not: a threshold must stay fixed whatever the report has selected. No page of the shipped report filters `JobStatus` or `IsWeekend`, so adding `KEEPFILTERS` changed no figure it shows. That was checked value by value when the change was made.
>
> **One consequence needs a guard.** Utilization is defined on *completed* work. With the report filtered to cancelled jobs it is not 0%, it does not apply. `KEEPFILTERS` makes `[Utilization %]` BLANK there, and DAX compares BLANK as 0, so without a guard every technician would read "Underutilized", turn red and gain a `[Risk Score]` point. `Utilization Status`, `Utilization Color` and `Risk Score` therefore return BLANK when the `JobStatus` selection excludes *Completed*. The test is `ISFILTERED` plus `FILTERS`, which look only at a filter placed on `JobStatus` itself. A technician whose only jobs in a month were cancelled is not caught by it and still, correctly, reads "Underutilized".

---

## Core volume measures

```dax
Total Jobs =
COUNTROWS ( Fact_ServiceJobs )
```

```dax
Completed Jobs =
CALCULATE ( [Total Jobs], KEEPFILTERS ( Fact_ServiceJobs[JobStatus] = "Completed" ) )
```

```dax
Cancelled Jobs =
CALCULATE ( [Total Jobs], KEEPFILTERS ( Fact_ServiceJobs[JobStatus] = "Cancelled" ) )
```

```dax
Completion Rate % =
DIVIDE ( [Completed Jobs], [Total Jobs] )
```

---

## Quality / reliability measures

```dax
First-Time-Fix % =
CALCULATE (
    AVERAGE ( Fact_ServiceJobs[FirstTimeFix] ),
    KEEPFILTERS ( Fact_ServiceJobs[JobStatus] = "Completed" )
)
```

```dax
Callback % =
CALCULATE (
    AVERAGE ( Fact_ServiceJobs[CallbackFlag] ),
    KEEPFILTERS ( Fact_ServiceJobs[JobStatus] = "Completed" )
)
```

```dax
SLA Compliance % =
CALCULATE (
    AVERAGE ( Fact_ServiceJobs[SLAMet] ),
    KEEPFILTERS ( Fact_ServiceJobs[JobStatus] = "Completed" )
)
```

Note: `FirstTimeFix`, `CallbackFlag`, and `SLAMet` are stored as 0/1 (bit) — `AVERAGE` over 0/1 is the proportion, so format these measures as Percentage in the Power BI field pane.

---

## Utilization measures (mirrors `dbo.vw_TechnicianUtilization`)

```dax
Available Minutes =
VAR WeekdayDates =
    FILTER ( Dim_Date, Dim_Date[IsWeekend] = 0 )
RETURN
    SUMX (
        FILTER ( Dim_Technician, Dim_Technician[EmploymentStatus] = "Active" ),
        VAR HireD = Dim_Technician[HireDate]
        VAR CapacityMinutes = Dim_Technician[DailyCapacityHours] * 60
        RETURN
            COUNTROWS ( FILTER ( WeekdayDates, Dim_Date[Date] >= HireD ) ) * CapacityMinutes
    )
```

> **Do not reach for `RELATED()` here.** An earlier version of this measure iterated `CROSSJOIN ( VALUES ( Dim_Technician[TechnicianKey] ), VALUES ( Dim_Date[DateKey] ) )` and pulled the other columns with `RELATED()`. That fails — `CROSSJOIN` of two `VALUES()` produces a *virtual* two-column table which sits on no relationship, so `RELATED ( Dim_Technician[EmploymentStatus] )` raises **"either doesn't exist or doesn't have a relationship to any table available in the current context"**, and every measure downstream of it inherits the error. `RELATED()` only works when the row context is on the *many* side of a real relationship. Iterating `Dim_Technician` and `Dim_Date` directly keeps their columns in row context, so no `RELATED()` is needed at all.

> `FILTER` here iterates the date table *within* the report's current selection, so it already intersects with a date or weekday slicer. That is why this measure needs no `KEEPFILTERS`.

```dax
Worked Minutes =
CALCULATE (
    SUM ( Fact_ServiceJobs[JobDurationMin] ) + SUM ( Fact_ServiceJobs[TravelTimeMin] ),
    KEEPFILTERS ( Fact_ServiceJobs[JobStatus] = "Completed" ),
    KEEPFILTERS ( Dim_Date[IsWeekend] = 0 )
)
```

```dax
Utilization % =
DIVIDE ( [Worked Minutes], [Available Minutes] )
```

> **Worked Minutes counts weekdays only, because capacity does.** `[Available Minutes]` is built from scheduled weekdays, so the numerator must be weekday work too. Counting every completed job against weekday-only capacity put 2,194 weekend call-outs (343,381 minutes) into the numerator with no capacity behind them, and inflated utilization from 65.891% to 70.455% across the book. This is the defect `docs/data_validation_report.md` §5b corrects.

Weekend work is not discarded. It is reported as overtime, matching the split in `vw_TechnicianUtilization`:

```dax
Overtime Minutes =
CALCULATE (
    SUM ( Fact_ServiceJobs[JobDurationMin] ) + SUM ( Fact_ServiceJobs[TravelTimeMin] ),
    KEEPFILTERS ( Fact_ServiceJobs[JobStatus] = "Completed" ),
    KEEPFILTERS ( Dim_Date[IsWeekend] = 1 )
)
```

```dax
Total Worked Minutes =
[Worked Minutes] + [Overtime Minutes]
```

```dax
Overtime % of Capacity =
DIVIDE ( [Overtime Minutes], [Available Minutes] )
```

> `Available Minutes` iterates every active technician against every weekday since their hire date, because capacity must be counted even on days a technician did *no* job — the whole point of a utilization KPI. It is the most expensive measure in the model. If report performance ever matters more than exact parity with the SQL view, import `dbo.vw_TechnicianDailyCapacity` as its own table and replace this measure with a plain `SUM` over it. Both are defensible; the model uses the iterator because it needs no SQL objects beyond the star schema.

---

## Financial measures

```dax
Total Revenue =
CALCULATE ( SUM ( Fact_ServiceJobs[JobRevenue] ), KEEPFILTERS ( Fact_ServiceJobs[JobStatus] = "Completed" ) )
```

```dax
Total Cost =
CALCULATE ( SUM ( Fact_ServiceJobs[JobCost] ), KEEPFILTERS ( Fact_ServiceJobs[JobStatus] <> "Rescheduled" ) )
```

```dax
Gross Margin =
[Total Revenue] - [Total Cost]
```

```dax
Revenue per Completed Job =
DIVIDE ( [Total Revenue], [Completed Jobs] )
```

---

## Target / conditional-formatting measures (mirrors `dbo.Ref_KPITargets`)

These are the one place `KEEPFILTERS` would be wrong: a threshold must not move with the report's selection. Each clears every filter on `Ref_KPITargets` with `REMOVEFILTERS` and then picks its own row, so a slicer on, say, `Direction` cannot blank a threshold and set off every status test against it.

```dax
Util Warning Threshold =
CALCULATE ( VALUES ( Ref_KPITargets[WarningValue] ), REMOVEFILTERS ( Ref_KPITargets ), Ref_KPITargets[MetricName] = "UtilizationPct" )
```

```dax
FTF Warning Threshold =
CALCULATE ( VALUES ( Ref_KPITargets[WarningValue] ), REMOVEFILTERS ( Ref_KPITargets ), Ref_KPITargets[MetricName] = "FirstTimeFixPct" )
```

```dax
SLA Warning Threshold =
CALCULATE ( VALUES ( Ref_KPITargets[WarningValue] ), REMOVEFILTERS ( Ref_KPITargets ), Ref_KPITargets[MetricName] = "SLACompliancePct" )
```

```dax
Callback Warning Threshold =
CALCULATE ( VALUES ( Ref_KPITargets[WarningValue] ), REMOVEFILTERS ( Ref_KPITargets ), Ref_KPITargets[MetricName] = "CallbackPct" )
```

```dax
Util Ceiling Threshold =
CALCULATE ( VALUES ( Ref_KPITargets[WarningValue] ), REMOVEFILTERS ( Ref_KPITargets ), Ref_KPITargets[MetricName] = "UtilizationCeilingPct" )
```

```dax
Utilization Status =
VAR Avail = [Available Minutes]
VAR U = [Utilization %] * 100
VAR CompletedInScope =
    NOT ISFILTERED ( Fact_ServiceJobs[JobStatus] )
        || "Completed" IN FILTERS ( Fact_ServiceJobs[JobStatus] )
RETURN
    IF (
        ISBLANK ( Avail ) || Avail = 0 || NOT CompletedInScope,
        BLANK (),
        SWITCH (
            TRUE(),
            U < [Util Warning Threshold], "Underutilized",
            U > [Util Ceiling Threshold], "Over Capacity",
            "On Target"
        )
    )
```

Companion status labels (text — for tooltips or an extra table column):

```dax
FTF Status =
VAR V = [First-Time-Fix %] * 100
RETURN IF ( ISBLANK ( V ), BLANK (), IF ( V < [FTF Warning Threshold], "Below Threshold", "Within Threshold" ) )
```

```dax
SLA Status =
VAR V = [SLA Compliance %] * 100
RETURN IF ( ISBLANK ( V ), BLANK (), IF ( V < [SLA Warning Threshold], "Below Threshold", "Within Threshold" ) )
```

```dax
Callback Status =
VAR V = [Callback %] * 100
RETURN IF ( ISBLANK ( V ), BLANK (), IF ( V > [Callback Warning Threshold], "Above Threshold", "Within Threshold" ) )
```

### Colour measures (for conditional formatting)

> **A status label can't drive conditional formatting.** *Format style: Field value* only applies a colour if the field returns a colour string (a hex code like `#D03B3B` or a colour name). A measure returning "Underutilized" is silently ignored. These measures return a hex colour **only when a threshold is breached**, and BLANK otherwise, so only exceptions get highlighted. Each rule is the exact condition that adds a point to `[Risk Score]`, so a red cell always accounts for part of the score, and all thresholds still come from `Ref_KPITargets`.

```dax
Utilization Color =
VAR Avail = [Available Minutes]
VAR U = [Utilization %] * 100
VAR CompletedInScope =
    NOT ISFILTERED ( Fact_ServiceJobs[JobStatus] )
        || "Completed" IN FILTERS ( Fact_ServiceJobs[JobStatus] )
RETURN
    IF (
        ISBLANK ( Avail ) || Avail = 0 || NOT CompletedInScope,
        BLANK (),
        SWITCH (
            TRUE (),
            U < [Util Warning Threshold], "#D03B3B",   -- underutilized
            U > [Util Ceiling Threshold], "#EC835A",   -- over capacity
            BLANK ()
        )
    )
```

```dax
FTF Color =
VAR V = [First-Time-Fix %] * 100
RETURN IF ( NOT ISBLANK ( V ) && V < [FTF Warning Threshold], "#D03B3B" )
```

```dax
SLA Color =
VAR V = [SLA Compliance %] * 100
RETURN IF ( NOT ISBLANK ( V ) && V < [SLA Warning Threshold], "#D03B3B" )
```

```dax
Callback Color =
VAR V = [Callback %] * 100
RETURN IF ( NOT ISBLANK ( V ) && V > [Callback Warning Threshold], "#D03B3B" )
```

**Applying one:** select the table visual → *Format* → *Cell elements* → pick the series (e.g. `Utilization %`) → turn on **Background color** → *fx* → *Format style: Field value* → *What field should we base this on:* `Utilization Color`. Use **Background color only**, not Font color too — bound to the same colour measure, both would paint the text the same colour as its cell and the number would disappear.

---

## Priority Action Queue measure (feeds a table visual, sorted by this descending)

```dax
Risk Score =
VAR Avail   = [Available Minutes]
VAR UtilPct = [Utilization %] * 100
VAR FTFPct  = [First-Time-Fix %] * 100
VAR SLAPct  = [SLA Compliance %] * 100
VAR CBPct   = [Callback %] * 100
VAR CompletedInScope =
    NOT ISFILTERED ( Fact_ServiceJobs[JobStatus] )
        || "Completed" IN FILTERS ( Fact_ServiceJobs[JobStatus] )
VAR Score =
    IF ( UtilPct < [Util Warning Threshold], 1, 0 )
        + IF ( UtilPct > [Util Ceiling Threshold], 1, 0 )
        + IF ( NOT ISBLANK ( FTFPct ) && FTFPct < [FTF Warning Threshold], 1, 0 )
        + IF ( NOT ISBLANK ( SLAPct ) && SLAPct < [SLA Warning Threshold], 1, 0 )
        + IF ( NOT ISBLANK ( CBPct ) && CBPct > [Callback Warning Threshold], 1, 0 )
RETURN
    IF ( ISBLANK ( Avail ) || Avail = 0 || NOT CompletedInScope, BLANK (), Score )
```

> **Why the blank guards.** A technician with no capacity in the current filter context — Inactive, or a month earlier than their `HireDate` — returns BLANK utilization. DAX compares BLANK as 0, so without the guard every "below target" test fired and those technicians scored a false `Risk Score` of 3, putting five phantom names at the top of the Priority Queue page. Returning BLANK instead makes them drop out of visuals entirely, which matches `dbo.vw_PriorityActionQueue` (built only over active technicians). Verified: DAX and SQL both return **25** flagged technicians for December 2025, with identical per-technician values.

Build the Priority Action Queue page as a table visual on `Dim_Technician[TechnicianName]`, `Dim_Region[RegionName]`, `[Utilization %]`, `[First-Time-Fix %]`, `[SLA Compliance %]`, `[Callback %]`, `[Risk Score]` — sort by `[Risk Score]` descending, filter to `[Risk Score] > 0`.
