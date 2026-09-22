# DAX Measures — Lumen Optics Photonics Spend Scorecard

**Model:** `LumenSpend` on SQL Server · **Reporting date:** 2025-12-31

> **Data disclosure.** Lumen Optics Manufacturing is fictional and all data is synthetic. No
> confidential data is used and no claim is made about any production system.

---

## The one rule that governs every measure here

**Thresholds are never written into DAX.** Every target and warning level lives in
`Ref_SpendTargets` and is read with `LOOKUPVALUE`. A threshold hard-coded into a measure is a
threshold that will eventually disagree with the SQL layer, the Excel workbook and the printed
report — and the disagreement will surface in a meeting rather than in a test.

The same applies to the price-erosion benchmark, which lives in `Ref_PriceErosionBenchmark`
because it is the single most contestable input in the analysis.

---

## 1. Model shape

> **This section was rewritten after the model was actually built.** An earlier version of this
> document described a seven-table model loading the raw `Fact_PurchaseOrderLine` and
> `Fact_GoodsReceipt` tables, with an inactive `Dim_Date → Fact_GoodsReceipt[ReceiptDateKey]`
> relationship reached through `USERELATIONSHIP`. That design was never built, and documenting a
> model that does not exist is worse than documenting none: a reader checks the DAX against the
> wrong shape and concludes the code is wrong. `Build_PowerBI_Model.ps1` is the source of truth;
> this file explains it. §1a records what changed and why.

Ten tables, star-shaped, one direction of filter flow throughout.

| Table | Role | Source | Grain |
|---|---|---|---|
| `Dim_Date` | dimension, marked as date table | `Dim_Date` | one row per calendar date |
| `Dim_Vendor` | dimension | `Dim_Vendor` | one row per supplier |
| `Dim_Part` | dimension | `Dim_Part` | one row per part number |
| `Dim_Buyer` | dimension | `Dim_Buyer` | one row per category manager |
| `POLine` | fact | `fn_POLineCost(@AsOf)` | one row per PO line, receipts pre-aggregated on |
| `PriceErosion` | fact | `fn_PriceErosion(@AsOf,10,500)` | one row per vendor–part pair |
| `RenegotiationQueue` | fact | `fn_RenegotiationQueue(@AsOf,8)` | one row per queued pair |
| `Ref_SpendTargets`, `Ref_PriceErosionBenchmark` | reference | tables | one row per metric / per category |
| `Ref_Reporting` | reference | `fn_SpendKPI(@AsOf)` | exactly one row: the reporting anchor |

### Relationships

```
POLine[OrderDate]              * → 1  Dim_Date[Date]
POLine[PartKey]                * → 1  Dim_Part[PartKey]
POLine[VendorKey]              * → 1  Dim_Vendor[VendorKey]
POLine[BuyerKey]               * → 1  Dim_Buyer[BuyerKey]
PriceErosion[PartKey]          * → 1  Dim_Part[PartKey]
PriceErosion[VendorKey]        * → 1  Dim_Vendor[VendorKey]
RenegotiationQueue[PartKey]    * → 1  Dim_Part[PartKey]
RenegotiationQueue[VendorKey]  * → 1  Dim_Vendor[VendorKey]
RenegotiationQueue[BuyerKey]   * → 1  Dim_Buyer[BuyerKey]
Dim_Part[Category]             * → 1  Ref_PriceErosionBenchmark[Category]
```

All active, all single-direction. `Dim_Part` and `Dim_Vendor` each feed three facts, which is an
ordinary star — the ambiguity trap would be a relationship *between* two facts, and there is none.
`Ref_Reporting` is deliberately unrelated to everything: it is read with `MAX()`, which is why it
works from any filter context.

**Why the benchmark joins to `Dim_Part`, not to the fact.** The erosion rate is a property of the
category, and `Dim_Part` already carries the category. Joining it to the fact would duplicate the
rate on every line and make a simple `AVERAGE` over it silently volume-weighted.

---

## 1a. What changed from the original design, and why

**The facts load from views, not raw tables.** The landed-cost identity, the dated contract
resolution and the erosion model are each defined once in SQL and asserted by 21 acceptance tests.
Re-deriving them in DAX would create a second definition, free to drift from the first, with
nothing asserting the two agree. DAX does the aggregation, the ratios, the time intelligence and
the conditional formatting — which is what a semantic model is actually for.

**`Fact_GoodsReceipt` is not in the model, and neither is the inactive relationship.**
`fn_POLineCost` already aggregates receipts onto the PO line as `QtyReceived`, `QtyAccepted`,
`QtyRejected`, `DaysLate` and `IsOnTime`, at an as-of date. The receipt grain is therefore not
reachable from the model — and that is a genuine trade-off, stated rather than hidden:

- *Lost:* per-receipt analysis (reject reasons by receipt, partial-delivery patterns) and any
  measure that needs to slice by receipt date rather than order date. `RejectReasonCode` lives on
  `Fact_GoodsReceipt` and nothing wraps it, so that question has to be asked in SQL against the
  table directly — see §9, where the gap is stated rather than mitigated.
- *Gained:* one definition of "what arrived and what was usable", with the as-of cut-off applied
  in one place. The `USERELATIONSHIP` design put that cut-off in every measure that touched it,
  and a measure that forgot it would have silently counted receipts booked after the reporting
  date.

The original reasoning for the inactive relationship still stands on its own terms — order date and
receipt date answer different questions, and letting spend move when deliveries slip is wrong. It
was the right call for a model built on raw facts. This model does not need it, because the
distinction is resolved in SQL where it can be tested.

**`Ref_Reporting` was added.** One row, holding the as-of date and the start of the trailing twelve
months, read straight out of `fn_SpendKPI`. Every TTM measure takes its window from it, so the
model and the SQL scorecard cannot drift. The first draft of `Spend (TTM)` anchored on
`MAX(POLine[OrderDate])` instead; that is identical today only because the last order happens to
fall on the as-of date, and would have silently widened the window the moment it did not. See §10.

---

## 2. Base measures

```dax
Extended Price =
SUMX ( POLine, POLine[OrderQty] * POLine[UnitPrice] )
```
`SUMX` rather than a calculated column: the product is needed only in aggregate, and a stored
column would cost memory on every one of the rows without being useful on its own.

```dax
Freight  = SUM ( POLine[FreightAmount] )
Expedite = SUM ( POLine[ExpediteFee] )

Landed Cost = [Extended Price] + [Freight] + [Expedite]
```

**The landed cost identity, stated once.** Freight and expedite are deliberately not folded into
the unit price: a unit price is negotiated with the supplier, an expedite fee is caused by
planning, and combining them hides which of the two is costing the money.

```dax
Qty Ordered  = SUM ( POLine[OrderQty] )
Qty Received = SUM ( POLine[QtyReceived] )
Qty Accepted = SUM ( POLine[QtyAccepted] )
Qty Rejected = SUM ( POLine[QtyRejected] )
```

---

## 3. The measure that changes decisions

```dax
Cost per Accepted Unit =
DIVIDE ( [Landed Cost], [Qty Accepted] )
```

Unit price answers *what did we agree to pay*. This answers *what did each usable part cost*. In
optics the two diverge enough to reverse a supplier ranking, because a component that arrives
outside specification was still paid for, still freighted, and cannot be built into a module.

```dax
Cost Index vs Best =
VAR Clean = FILTER ( ALL ( POLine[IsDuplicatePO] ), POLine[IsDuplicatePO] = FALSE () )
VAR BestCost =
    CALCULATE (
        SUMX (
            VALUES ( Dim_Part[PartKey] ),
            VAR BestOnPart =
                MINX (
                    CALCULATETABLE ( VALUES ( Dim_Vendor[VendorKey] ), ALL ( Dim_Vendor ) ),
                    CALCULATE ( [Cost per Accepted Unit] )
                )
            RETURN BestOnPart * CALCULATE ( [Qty Accepted] )
        ),
        Clean
    )
RETURN DIVIDE ( CALCULATE ( [Landed Cost], Clean ), BestCost ) * 100
```

**It aggregates per PART, never per PO line.** The SQL implementation of this index originally
joined line-level rows to a pair-level aggregate and summed the pair's landed cost once per line —
a join fan-out. It inflated numerator and denominator by *per-pair line counts*, which are not a
common factor, so the result stayed above 100, stayed entirely plausible, and was quietly a
line-count-weighted average of the index rather than the index. It was wrong for all 36 vendors,
by up to 3.07 index points, and moved 33 of their rank positions. `UAT-12` originally asserted only
"index ≥ 100", which is true under *any* positive weighting — a test that could not fail for the
reason it was written. It now recomputes the index at pair grain and compares.

**It excludes duplicated requisitions**, matching `fn_VendorScorecard`.

**Why this is not a simple average across vendors.** Vendors supply different parts at different
base prices, so an average cost per accepted unit across a mixed basket compares nothing. This
compares each vendor with the cheapest available *on the same part*, which is the only
like-for-like question. 100 means best available; 108 means eight per cent dearer per usable part.

`ALLEXCEPT` keeps the part filter and clears the vendor filter, so `MINX` iterates every supplier
qualified on that part rather than every supplier in the company.

---

## 4. Contract compliance

```dax
On Contract Spend =
CALCULATE ( [Extended Price], POLine[IsOnContract] = TRUE () )

Maverick Spend =
CALCULATE ( [Extended Price], POLine[IsOnContract] = FALSE () )

Maverick Spend % =
DIVIDE ( [Maverick Spend], [Extended Price] ) * 100
```

```dax
Purchase Price Variance =
SUMX (
    FILTER ( POLine, NOT ISBLANK ( POLine[ContractedUnitPrice] ) ),
    ( POLine[UnitPrice] - POLine[ContractedUnitPrice] )
        * POLine[OrderQty]
)

PPV % =
DIVIDE (
    [Purchase Price Variance],
    CALCULATE ( [Extended Price], NOT ISBLANK ( POLine[ContractedUnitPrice] ) )
) * 100
```

**The `FILTER` is load-bearing.** Off-contract lines have no contracted price, and including them
would treat a blank as zero — scoring a variance of "paid the full price against a contract of
nothing". Both numerator and denominator are restricted to contracted lines, so PPV answers only
the question it can answer, and maverick spend answers the rest.

---

## 5. The price erosion model

This is the centrepiece and the thing purchase price variance cannot see.

> **Five of the blocks in this section are NOT measures in the built model.**
> `First Price`, `Last Price`, `Years Elapsed`, `Expected Price` and
> `Actual Erosion %` are computed in SQL by `fn_PriceErosion` and arrive in the
> model as **columns** on the `PriceErosion` table. `Erosion Capture %
> (Spend Weighted)` is likewise a SQL figure, surfaced through
> `[Erosion Capture %]`.
>
> The DAX is kept because it is the clearest statement of the arithmetic, and
> because anyone rebuilding this in a tool without a SQL layer will need it.
> But it is **illustrative**: search the `.pbix` for these names and you will
> not find them. Section 1 says the build script is the source of truth, and
> that applies to this section too — it previously did not say so, and a reader
> had no way to tell these eight blocks from the sixty-nine that are real.

```dax
Benchmark Erosion % =
AVERAGEX ( VALUES ( Dim_Part[Category] ),
    LOOKUPVALUE ( Ref_PriceErosionBenchmark[AnnualErosionPct],
                  Ref_PriceErosionBenchmark[Category], Dim_Part[Category] ) )
```

```dax
First Price =
VAR FirstOrder = CALCULATE ( MIN ( POLine[OrderDateKey] ), ALLSELECTED ( Dim_Date ) )
RETURN
    CALCULATE (
        DIVIDE ( [Extended Price], [Qty Ordered] ),
        POLine[OrderDateKey] <= FirstOrder + 90,
        ALLSELECTED ( Dim_Date )
    )

Last Price =
VAR LastOrder = CALCULATE ( MAX ( POLine[OrderDateKey] ), ALLSELECTED ( Dim_Date ) )
RETURN
    CALCULATE (
        DIVIDE ( [Extended Price], [Qty Ordered] ),
        POLine[OrderDateKey] >= LastOrder - 90,
        ALLSELECTED ( Dim_Date )
    )
```

**Volume-weighted over a 90-day window, not a single transaction.** One small rush order at a bad
price should not set the baseline the whole opportunity is measured from. The SQL layer enforces a
180-day minimum history so the two windows cannot overlap; the same caveat applies here and the
report filters to pairs with enough history.

```dax
Years Elapsed =
VAR FirstOrder = CALCULATE ( MIN ( POLine[OrderDateKey] ), ALLSELECTED ( Dim_Date ) )
VAR LastOrder  = CALCULATE ( MAX ( POLine[OrderDateKey] ), ALLSELECTED ( Dim_Date ) )
RETURN
    ROUND ( DIVIDE ( DATEDIFF ( FirstOrder, LastOrder, DAY ), 365.25 ), 3 )

Expected Price =
ROUND ( [First Price] * POWER ( 1 - DIVIDE ( [Benchmark Erosion %], 100 ), [Years Elapsed] ), 4 )

Actual Erosion % =
ROUND (
    ( 1 - POWER ( DIVIDE ( [Last Price], [First Price] ), DIVIDE ( 1, [Years Elapsed] ) ) ) * 100,
    3
)

Erosion Capture % =
DIVIDE ( [Actual Erosion %], [Benchmark Erosion %] ) * 100
```

**Every one of these rounds explicitly, and the ones below use the rounded values.** A derived
figure must agree with the components printed beside it. During the SQL build, `ExpectedPrice` was
computed from an unrounded exponent while `YearsElapsed` was published rounded, and a reader
redoing the arithmetic on the page got a different answer. UAT-05 and UAT-07 exist because of it.

```dax
Erosion Gap per Unit = [Last Price] - [Expected Price]

Annual Opportunity =
VAR Gap = [Erosion Gap per Unit]
VAR TTMQty =
    CALCULATE ( [Qty Ordered], DATESINPERIOD ( Dim_Date[Date], MAX ( Dim_Date[Date] ), -12, MONTH ) )
RETURN
    IF ( Gap > 0, Gap * TTMQty, 0 )
```

**Clamped at zero, not netted.** A supplier already beating the curve is not a savings source, and
letting its negative gap offset another supplier's shortfall would understate the actionable
target and imply an offset nobody can realise. The pairs that are ahead are reported separately:

```dax
Value Already Ahead of Curve =
VAR Gap = [Erosion Gap per Unit]
VAR TTMQty =
    CALCULATE ( [Qty Ordered], DATESINPERIOD ( Dim_Date[Date], MAX ( Dim_Date[Date] ), -12, MONTH ) )
RETURN
    IF ( Gap < 0, Gap * TTMQty, 0 )
```

```dax
Erosion Capture % (Spend Weighted) =
VAR Pairs =
    SUMMARIZE ( POLine, Dim_Part[PartKey], Dim_Vendor[VendorKey] )
VAR WithWeights =
    ADDCOLUMNS ( Pairs, "@Capture", [Erosion Capture %], "@Spend", [Extended Price] )
RETURN
    DIVIDE (
        SUMX ( WithWeights, [@Capture] * [@Spend] ),
        SUMX ( WithWeights, [@Spend] )
    )
```

**Spend-weighted, never a simple average.** A vendor-part pair carrying $4m of spend and one
carrying $40k should not count equally in a company-level figure, and a plain `AVERAGE` would say
they do.

---

## 6. Quality and delivery

```dax
Acceptance Rate % = DIVIDE ( [Qty Accepted], [Qty Received] ) * 100

Rejected Value = SUM ( POLine[RejectedValue] )
```

`RejectedValue` is computed in `fn_POLineCost` as rejected quantity × the line's unit price, so the
DAX is a plain `SUM`. Doing the multiplication here instead would need the receipt grain, and would
be a second definition of "what the unusable material cost" with nothing asserting it agrees with
the first.

**Reject attribution lives in SQL, not here.** Only supplier-caused rejects are negotiable — a
rejection against a drawing revision issued *after* the order is Lumen's problem, and charging it
to the supplier in a business review destroys credibility faster than any number wins. That split
needs `RejectReasonCode` at receipt grain, which this model does not carry (§1a).

**It is not answered anywhere yet, and that is a gap rather than a design choice.** The column
exists on `Fact_GoodsReceipt`; no view or procedure wraps it, so the question has to be asked with
SQL directly against that table. An earlier draft of this document named a `usp_RejectAnalysis` in
`sql/06_stored_procedures.sql` as the answer. There is no such procedure — the file creates six,
and none of them is it. Saying so plainly is the point: the alternative is a reader assuming the
Power BI rejected-value figure is already attributed, which is exactly the assumption this section
exists to prevent.

```dax
Receipts Booked  = CALCULATE ( COUNTROWS ( POLine ), POLine[IsReceived] = TRUE () )
On Time Receipts = CALCULATE ( COUNTROWS ( POLine ),
                               POLine[IsOnTime] = TRUE (), POLine[IsReceived] = TRUE () )

On Time Delivery % = DIVIDE ( [On Time Receipts], [Receipts Booked] ) * 100

Avg Days Late = CALCULATE ( AVERAGE ( POLine[DaysLate] ), POLine[IsReceived] = TRUE () )
```

**Both sides of the ratio are restricted to lines that have actually been received.** A line still
in transit is neither on time nor late yet; counting it in the denominator only would drag on-time
delivery down every time the order book grew, and counting it in neither is the only reading that
does not move for a reason unrelated to supplier performance.

`IsOnTime` is `NULL` in SQL for a line with no receipt — not `0`. That distinction is load-bearing:
`ISNULL(IsOnTime,0)` in the M query would silently reclassify every undelivered line as *late*,
which is why the model carries `IsReceived` alongside it and every delivery measure filters on it.

```dax
Cost per Accepted Unit = DIVIDE ( [Landed Cost], [Qty Accepted] )
Duplicate PO Lines     = CALCULATE ( COUNTROWS ( POLine ), POLine[IsDuplicatePO] = TRUE () )
Duplicate PO Spend     = CALCULATE ( [Extended Price], POLine[IsDuplicatePO] = TRUE () )
```

`IsDuplicatePO` comes from `fn_DuplicatePOLines`, which matches a repeated requisition on its
business signature — same part, vendor, date, quantity and price — and never on the generator's
`PO-D` prefix. The model needs the flag because the two populations genuinely differ: `fn_SpendKPI`
*counts* duplicates (committed spend is what was committed), while `fn_VendorScorecard` *excludes*
them (a vendor should not be ranked on a requisition somebody raised twice). Without the column,
only one of those two views could be expressed in DAX.

---

## 7. Leverage and concentration

```dax
Single Source Spend % =
DIVIDE (
    CALCULATE ( [Extended Price], Dim_Part[QualifiedSupplierCount] = 1 ),
    [Extended Price]
) * 100

Top 5 Vendor Share % =
VAR VendorSpend =
    ADDCOLUMNS ( VALUES ( Dim_Vendor[VendorKey] ), "@Spend", [Extended Price] )
VAR Top5 = TOPN ( 5, VendorSpend, [@Spend], DESC )
RETURN
    DIVIDE ( SUMX ( Top5, [@Spend] ), SUMX ( VendorSpend, [@Spend] ) ) * 100
```

```dax
Leverage Score =
VAR Sources = SELECTEDVALUE ( Dim_Part[QualifiedSupplierCount] )
VAR QualMonths = SELECTEDVALUE ( Dim_Part[QualificationMonths] )
VAR RevShare = SELECTEDVALUE ( Dim_Vendor[LumenRevenueSharePct] )
RETURN
      SWITCH ( TRUE (), Sources >= 3, 40, Sources = 2, 26, 4 )
    + SWITCH ( TRUE (), QualMonths <= 2, 20, QualMonths <= 5, 13, QualMonths <= 8, 6, 0 )
    + SWITCH ( TRUE (), RevShare >= 15, 30, RevShare >= 8, 20, RevShare >= 4, 11, 3 )
    + IF ( [Maverick Spend] > 0, 10, 6 )
```

Identical weights to the SQL implementation, deliberately. The two must agree or the Power BI
report and the Excel workbook will rank the same negotiations differently.

---

## 8. Targets and conditional formatting

> **`Target Value` and `Warning Value` are not measures in the built model.**
> They are **columns** on `Ref_SpendTargets`, read directly by the `LOOKUPVALUE`
> calls inside each `… Colour` and `… Status` measure — which *are* built, 16 of
> them, generated from `$ragMetrics`. The two blocks below show the lookup in
> isolation for readability.


```dax
Target Value =
LOOKUPVALUE ( Ref_SpendTargets[TargetValue],
              Ref_SpendTargets[MetricName], SELECTEDVALUE ( Ref_SpendTargets[MetricName] ) )

Warning Value =
LOOKUPVALUE ( Ref_SpendTargets[WarningValue],
              Ref_SpendTargets[MetricName], SELECTEDVALUE ( Ref_SpendTargets[MetricName] ) )
```

```dax
Erosion Capture Colour =
VAR V = [Erosion Capture % (Spend Weighted)]
VAR T = LOOKUPVALUE ( Ref_SpendTargets[TargetValue],  Ref_SpendTargets[MetricName], "ErosionCapturePct" )
VAR W = LOOKUPVALUE ( Ref_SpendTargets[WarningValue], Ref_SpendTargets[MetricName], "ErosionCapturePct" )
RETURN
    SWITCH ( TRUE (), V >= T, "#C6EFCE", V >= W, "#FFEB9C", "#FFC7CE" )
```

**A colour measure returns a hex string, not a number.** Power BI's conditional formatting by
field value expects a colour string; returning 1, 2 or 3 and hoping the rule interprets it is the
most common reason a formatting rule silently does nothing. One such measure exists per KPI tile,
each reading its own thresholds from the reference table.

---

## 9. Measures deliberately not built

| Not built | Why |
|---|---|
| Should-cost variance | Needs a bottom-up build-up from materials, process time and overhead. The erosion model measures failure to erode from the first observed price; it cannot say whether that first price was fair, and pretending otherwise would be the more dangerous error. |
| Supplier risk score | Financial health, geographic and single-site exposure are not in this model. A risk score computed from delivery and quality alone would be named for something it does not measure. |
| Savings realised | Requires tracking negotiated outcomes back to subsequent prices. It is the natural follow-on once the queue has been worked for a quarter, and it is the measure that would prove the whole exercise paid for itself. |

Naming these matters as much as building the others: a model that quietly lacks a measure invites
someone to compute it badly in a visual.

---

## 10. The trailing-twelve-month window, and why it is a table

Every headline KPI in this project is a trailing-twelve-month figure. `fn_SpendKPI` windows its PO
lines with `OrderDate > DATEADD(MONTH, -12, @AsOf)`, and the case study quotes the results.

The model expresses that window **twice over**, on purpose.

**Base ratios stay filter-responsive.** `[Maverick Spend %]` answers over whatever the user has
sliced. That is what makes a monthly trend line or a per-vendor breakdown mean anything, and it is
the correct default for a semantic model.

**The headline KPIs get an explicit TTM twin.**

```dax
Maverick Spend % (TTM) =
CALCULATE (
    [Maverick Spend %],
    DATESBETWEEN ( Dim_Date[Date], MAX ( Ref_Reporting[TTMStart] ), MAX ( Ref_Reporting[AsOfDate] ) )
)
```

A card bound to the base measure with no slicer on the page shows an **all-history** number. It
looks exactly as plausible as the right one, it is not the number in the case study, and nothing
about the report makes the difference visible. Pinning the window inside the measure removes the
requirement to remember a page filter in order for the page to be correct.

`DATESBETWEEN` replaces only the *date* filter, so `[Acceptance Rate % (TTM)]` still answers
"how is **this vendor** doing" when a vendor is selected — it just refuses to answer it over a
window other than the published one. That is why the RAG colour measures bind to the TTM variants:
colouring an all-history figure against a TTM target is a red light with no meaning.

**Why the window comes from a table rather than from the data.** The first draft was:

```dax
VAR EndDate = CALCULATE ( MAX ( POLine[OrderDate] ), ALL ( POLine ) )
RETURN CALCULATE ( [Extended Price], DATESINPERIOD ( Dim_Date[Date], EndDate, -12, MONTH ) )
```

which is correct today and correct by coincidence. The last order in this dataset happens to fall
on 2025-12-31, so "twelve months back from the newest row" and "twelve months back from the
reporting date" are the same window. Had the last order landed on 29 December, `DATESINPERIOD`
would have started on 30 December 2024 and pulled two extra days of the prior year into every
headline figure — silently, with no error and no visible symptom.

An anchor that means *wherever the data happens to end* is not the same as one that means *the
reporting date*. `Ref_Reporting` holds one row read straight out of `fn_SpendKPI`, so the model and
the SQL scorecard take their window from the same source and cannot drift apart.

**`Erosion Capture %` and `Annual Opportunity` have no TTM twin, deliberately.** The erosion model
already weights by TTM spend internally and spans each pair's whole price history by design.
Windowing it again would double-apply the window and shorten the very history the trend is
measured over.

---

## 11. Verifying this document is true

`Validate_PowerBI_Model.ps1` queries the live model in DAX and reconciles 31 figures against the
SQL that produces them. Run it after any change here. A measure file that has drifted from the
model is exactly as misleading as the model-shape section this document had to have rewritten.
