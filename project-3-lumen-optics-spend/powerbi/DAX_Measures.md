# DAX Measures — Lumen Optics Photonics Spend Scorecard

**Model:** `LumenSpend` on SQL Server · **Reporting date:** 2025-12-31

> **Data disclosure.** Lumen Optics Manufacturing is fictional and all data is synthetic. No
> confidential data is used and no claim is made about any production system.

**This page is checked, not trusted.** `Build_PowerBI_Model.ps1` is the source of truth, and
`Sync_Measures.ps1 -DocOnly` compares every `dax` block below with it, token for token, failing on
any difference, any missing measure and any extra one. All 69 of the model's measures appear
exactly once: the ones the text discusses in their sections, the rest in the
[complete reference](#complete-reference) at the end. Anything that is *not* a model measure (an
earlier draft, or arithmetic that lives in SQL) is shown as plain `text` and labelled illustrative.
The check exists because this page drifted once: it showed a `Cost Index vs Best` with both of the
bugs its own prose said were fixed, and most of the measures on it were either missing or not the
ones the script builds.

---

## Thresholds are never written into DAX

Every target and warning level lives in `Ref_SpendTargets` and is read inside the measure that
needs it. A threshold hard-coded into a measure is a threshold that will eventually disagree with
the SQL layer, the Excel workbook and the printed report — and the disagreement will surface in a
meeting rather than in a test.

The same applies to the price-erosion benchmark, which lives in `Ref_PriceErosionBenchmark`
because it is the single most contestable input in the analysis. It reaches the measures already
applied: `fn_PriceErosion` puts the category's rate on every vendor–part pair as
`BenchmarkErosionPct`.

Two kinds of literal do appear in the DAX, and they are definitions rather than targets. The 20 in
`[Pairs Capturing Under 20%]` and `[Spend on Flat Pairs]` is the line below which
`usp_PriceErosionDetail` stops calling a pair even a partial capture, and the validator checks the
count against SQL's own `< 20`. The `[Leverage Score]` weights must be the SQL weights, digit for
digit (§7).

---

## One rule every filter below follows

> **A filter inside `CALCULATE` replaces; `KEEPFILTERS` intersects.**
> `CALCULATE ( [Extended Price], POLine[IsOnContract] = FALSE () )` discards any filter the report
> already has on `IsOnContract` and puts its own in its place. Under a slicer set to on-contract
> lines, `[Maverick Spend]` ignored the slicer and the scorecard's `[Maverick Spend % (TTM)]` read
> 37.15%, where the true answer is that there is no maverick spend among on-contract lines. Wrapped in `KEEPFILTERS`, the
> filter intersects with the report's selection instead, and the answer is none.
>
> Every filter on a column the report can also filter is wrapped this way: `IsOnContract`,
> `ContractedUnitPrice`, `AgreementsInForce`, `IsDuplicatePO`, `IsReceived`, `IsOnTime`,
> `Dim_Part[IsSingleSource]`, `PriceErosion[ErosionCapturePct]` and
> `RenegotiationQueue[IsThisQuarter]`.
>
> **A ratio whose numerator can come back empty carries `+ 0`.** An empty subset is BLANK, and
> `DIVIDE ( BLANK, x )` is BLANK, so a vendor with no off-contract spend would show an empty cell
> where the true figure is 0.00%. `[Maverick Spend %]`, `[On Time Delivery %]` and
> `[Single Source Spend %]` therefore add `+ 0` to the numerator. BLANK then means *there is nothing
> to measure here* (the denominator is empty too), and never hides a genuine 0%. The other
> percentage ratios, except `[Quarter Coverage %]` (§7a), take numerator and denominator from the
> same rows, so they are already 0 rather than BLANK wherever there is something to measure.
>
> **Status measures return BLANK for a BLANK value, and read their thresholds with
> `REMOVEFILTERS`.** DAX compares BLANK as 0, so without a guard a vendor with no spend read Green
> on every lower-is-better metric and Red on the rest. Every `… Colour` and `… Status` measure now
> returns BLANK when the value or either threshold is BLANK. That guard is safe only because of the
> `+ 0` above; without it, a genuine 0% would lose its Green. The thresholds are read with
> `CALCULATE ( VALUES ( … ), REMOVEFILTERS ( Ref_SpendTargets ), Ref_SpendTargets[MetricName] = "…" )`.
> The `LOOKUPVALUE` used before respected filters on `Ref_SpendTargets`, so a filter on that table
> could blank a threshold.
>
> **The deliberate exceptions**, each marked where it occurs in the build script:
>
> - **The TTM window.** `DATESBETWEEN` in the ten `(TTM)` measures replaces the date filter on
>   purpose: it pins the published period, so a scorecard card shows the case-study figure whatever
>   date slicer is on the page (§10).
> - **`ALL ( Dim_Vendor )` and `REMOVEFILTERS ( Dim_Buyer )` in `[Cost Index vs Best]`.** Clearing
>   the vendor selection is the whole point: the cheapest cost on a part has to be found among every
>   vendor, not only the selected one. The buyer is cleared for the same reason: who placed an order
>   does not change the best price Lumen was offered (§3). The same measure's duplicate filter is
>   `KEEPFILTERS`.
> - **The numerator of `[Quarter Coverage %]`** stays plain, and its denominator clears
>   `IsThisQuarter`, so it reports this quarter's share of the whole queue whatever an
>   `IsThisQuarter` slicer says (§7a).
> - **The threshold reads** clear every filter on `Ref_SpendTargets`: a target has to stay fixed
>   whatever the report has selected.
>
> None of the figures `Validate_PowerBI_Model.ps1` reconciles depends on the difference. Its checks
> either run unfiltered or filter `Dim_Vendor[VendorID]` and `Dim_Buyer[BuyerID]`, and no
> `KEEPFILTERS` in the model is on either of those columns.

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
| `Dim_Part` | dimension | `Dim_Part`, plus `IsSingleSource` | one row per part number |
| `Dim_Buyer` | dimension | `Dim_Buyer` | one row per category manager |
| `POLine` | fact | `fn_POLineCost(@AsOf)`, plus `IsDuplicatePO` | one row per PO line, receipts pre-aggregated on |
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
`Ref_Reporting` and `Ref_SpendTargets` are deliberately unrelated to everything: `Ref_Reporting` is
read with `MAX()` and the targets with `REMOVEFILTERS`, which is why both work from any filter
context. `PriceErosion` and `RenegotiationQueue` have no date relationship, so no date slicer moves
the erosion or queue measures.

**Why the benchmark table joins to `Dim_Part`, not to a fact.** The erosion rate is a property of
the category, and `Dim_Part` already carries the category, so a page can show a category's rate
beside its parts without copying it onto every line. No measure reads the table:
`fn_PriceErosion` has already applied the benchmark to each pair, and `[Benchmark Erosion %]`
weights that pair-level rate by trailing-twelve-month spend (§5).

---

## 1a. What changed from the original design, and why

**The facts load from the SQL layer's functions, not raw tables.** The landed-cost identity, the
dated contract resolution and the erosion model are each defined once in SQL and asserted by 21
acceptance tests. Re-deriving them in DAX would create a second definition, free to drift from the
first, with nothing asserting the two agree. DAX does the aggregation, the ratios, the time
intelligence and the conditional formatting — which is what a semantic model is actually for.

**`Fact_GoodsReceipt` is not in the model, and neither is the inactive relationship.**
`fn_POLineCost` already aggregates receipts onto the PO line as `QtyReceived`, `QtyAccepted`,
`QtyRejected`, `DaysLate` and `IsOnTime`, at an as-of date. The receipt grain is therefore not
reachable from the model — and that is a genuine trade-off, stated rather than hidden:

- *Lost:* per-receipt analysis (reject reasons by receipt, partial-delivery patterns) and any
  measure that needs to slice by receipt date rather than order date. `RejectReasonCode` lives on
  `Fact_GoodsReceipt` and nothing wraps it, so that question has to be asked in SQL against the
  table directly — see §6, where the gap is stated rather than mitigated.
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
SUM ( POLine[ExtendedPrice] )
```

```dax
Freight =
SUM ( POLine[FreightAmount] )
```

```dax
Expedite Fees =
SUM ( POLine[ExpediteFee] )
```

```dax
Landed Cost =
SUM ( POLine[LandedCost] )
```

**The landed cost identity, stated once — in SQL.** `fn_POLineCost` computes `ExtendedPrice`
(quantity × unit price) and `LandedCost` (extended price + freight + expedite fee) per line, and
the model only sums them. An earlier version of this page built `Extended Price` with a `SUMX` over
quantity × price and `Landed Cost` by adding three measures together: a second definition of the
identity, free to drift from the first, which is exactly what §1a says the model avoids.

Freight and expedite are deliberately not folded into the unit price: a unit price is negotiated
with the supplier, an expedite fee is caused by planning, and combining them hides which of the two
is costing the money.

```dax
PO Lines =
COUNTROWS ( POLine )
```

```dax
Qty Ordered =
SUM ( POLine[OrderQty] )
```

```dax
Qty Received =
SUM ( POLine[QtyReceived] )
```

```dax
Qty Accepted =
SUM ( POLine[QtyAccepted] )
```

```dax
Qty Rejected =
SUM ( POLine[QtyRejected] )
```

The three receipt quantities sum what `fn_POLineCost` has already aggregated onto each line as of
the reporting date (§1a). The model's source query loads them with `ISNULL(…, 0)`, so a line with
no receipt yet contributes 0 rather than BLANK. `[PO Lines]` counts every line, duplicated
requisitions included; committed spend is what was committed (§6).

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
CALCULATE (
    -- only parts with accepted quantity, as SQL's HAVING SUM(QtyAccepted) > 0
    VAR PartsSupplied = FILTER ( VALUES ( Dim_Part[PartKey] ), [Qty Accepted] > 0 )
    VAR MyCost = SUMX ( PartsSupplied, CALCULATE ( [Landed Cost] ) )
    VAR BestCost =
        SUMX (
            PartsSupplied,
            VAR Acc = CALCULATE ( [Qty Accepted] )
            VAR BestOnPart =
                MINX (
                    CALCULATETABLE ( VALUES ( Dim_Vendor[VendorKey] ), ALL ( Dim_Vendor ) ),
                    VAR vk = Dim_Vendor[VendorKey]
                    -- clear the outer vendor selection and the buyer, then apply only the iterated key
                    RETURN CALCULATE ( [Cost per Accepted Unit], ALL ( Dim_Vendor ), Dim_Vendor[VendorKey] = vk, REMOVEFILTERS ( Dim_Buyer ) )
                )
            RETURN BestOnPart * Acc
        )
    RETURN DIVIDE ( MyCost, BestCost ) * 100,
    KEEPFILTERS ( POLine[IsDuplicatePO] = FALSE () )
)
```

**Why this is not a simple average across vendors.** Vendors supply different parts at different
base prices, so an average cost per accepted unit across a mixed basket compares nothing. This
compares each vendor with the cheapest available *on the same part*, which is the only
like-for-like question. For every part the vendor had usable stock accepted on, it prices that
accepted quantity at the best cost per accepted unit any vendor achieved on the part, and divides
what the vendor actually cost by that. 100 means best available on everything it supplies; 108
means eight per cent dearer per usable part.

Four things in it have to be right, and each of them, when wrong, still returned a plausible
number.

**1. It aggregates per PART, never per PO line.** The SQL implementation of this index originally
joined line-level rows to a pair-level aggregate and summed the pair's landed cost once per line —
a join fan-out. It inflated numerator and denominator by *per-pair line counts*, which are not a
common factor, so the result stayed above 100, stayed entirely plausible, and was quietly a
line-count-weighted average of the index rather than the index. It was wrong for all 36 vendors,
by up to 3.07 index points, and moved 33 of their rank positions. `UAT-12` originally asserted only
"index ≥ 100", which is true under *any* positive weighting — a test that could not fail for the
reason it was written. It now recomputes the index at pair grain and compares. Here, both `SUMX`
calls iterate parts.

**2. Every rival's cost is computed with the vendor filter cleared.** The list `MINX` iterates is
right, but the context transition inside it adds each vendor key to a filter context that still
holds the outer vendor selection. Filters on one table intersect, so every other vendor resolves
to nothing and returns BLANK; `MINX` then sees only the selected vendor's own cost, and a vendor
compared with itself scores exactly 100.00. The first DAX version did exactly that. The inner
`CALCULATE` therefore clears `Dim_Vendor` and applies only the iterated key, `vk`. This
`ALL ( Dim_Vendor )` is one of the deliberate exceptions to the `KEEPFILTERS` rule above: clearing
the vendor selection is the point.

**3. Only parts with accepted quantity count**, as in SQL's `PartVendor` step. Without the
restriction, a part the vendor supplied but had nothing accepted on adds landed cost to the
numerator and nothing to the denominator, inflating the index.

**4. The benchmark clears the buyer too.** `POLine` carries a `BuyerKey`, so under a buyer filter
every rival's cost was computed from that buyer's orders only, and "the best any vendor achieved"
quietly became "the best vendor this buyer happened to use". BUY-01, BUY-03 and BUY-05 each read
exactly **100.00** (best available on everything) against true figures of 101.90, 102.66 and
103.95. Who placed an order does not change what the market offered, so the inner `CALCULATE`
clears `Dim_Buyer`. Dates and `POLine` attributes still scope the benchmark, so a period is
compared with the same period.

**It excludes duplicated requisitions**, matching `fn_VendorScorecard`: a vendor should not be
ranked on a requisition somebody raised twice. The exclusion is a `KEEPFILTERS`, so it narrows a
report filter on `IsDuplicatePO` instead of replacing it. The earlier
`FILTER ( ALL ( POLine[IsDuplicatePO] ), … )` cleared any outer filter on that column first, so
under a duplicates-only slicer it still scored the clean lines; now it is BLANK there, because
there are no clean lines to score.

`Validate_PowerBI_Model.ps1` reconciles this index against `fn_VendorScorecard` for every vendor,
not a sample. A single-vendor spot check would pass the self-comparison bug whenever the vendor
sampled happened to be the cheapest. It also reconciles the index for each of the six buyers
against SQL that prices every part at the Lumen-wide best, which is the check that fails on a
buyer-scoped benchmark.

---

## 4. Contract compliance

```dax
On Contract Spend =
CALCULATE ( [Extended Price], KEEPFILTERS ( POLine[IsOnContract] = TRUE () ) )
```

```dax
Maverick Spend =
CALCULATE ( [Extended Price], KEEPFILTERS ( POLine[IsOnContract] = FALSE () ) )
```

```dax
Maverick Spend % =
DIVIDE ( [Maverick Spend] + 0, [Extended Price] ) * 100
```

`IsOnContract` is decided in SQL: `fn_POLineCost` resolves the agreement in force at the order date,
and a line with none is maverick. The `+ 0` makes a vendor with no maverick spend read 0.00%
instead of BLANK.

```dax
Purchase Price Variance =
SUM ( POLine[PPVAmount] )
```

```dax
PPV % =
DIVIDE (
    [Purchase Price Variance],
    CALCULATE ( SUM ( POLine[ContractedExtended] ), KEEPFILTERS ( NOT ISBLANK ( POLine[ContractedUnitPrice] ) ) )
) * 100
```

**PPV is restricted to contracted lines on both sides, and the restriction is made in SQL.**
`fn_POLineCost` computes `PPVAmount` as (paid − contracted) × quantity and `ContractedExtended` as
contracted price × quantity only where an agreement was in force, and leaves both NULL on an
off-contract line. Scoring such a line as zero variance would report "paid the full price against
a contract of nothing" as perfect compliance, for exactly the spend nobody negotiated; `UAT-04`
asserts it does not happen. PPV answers only the question it can answer, and maverick spend answers
the rest.

**The denominator is contracted spend, not spend.** PPV % is the variance as a share of what the
contracted prices would have cost, the same `PPVPct` that `fn_SpendKPI` publishes. The
`KEEPFILTERS ( NOT ISBLANK ( … ) )` on it restates the SQL restriction rather than creating it:
`ContractedExtended` is already NULL on every line without a contracted price, so the filter
changes no figure. An earlier version of this page summed a `FILTER` over `POLine` and divided by
`[Extended Price]`; the second of those is a different figure, not a different style.

```dax
Ambiguous Contract Lines =
CALCULATE ( COUNTROWS ( POLine ), KEEPFILTERS ( POLine[AgreementsInForce] > 1 ) )
```

**PPV is only as good as the agreement it is measured against.** When two agreements overlap on the
order date, `fn_ContractedPrice` resolves the tie deterministically (latest start wins, then
highest key) and reports how many agreements were in force. This counts the lines where that was
more than one, so a variance figure can be read beside how much of it rests on a tie-break.

---

## 5. The price erosion model

This is the centrepiece and the thing purchase price variance cannot see.

**The erosion arithmetic lives in SQL, one row per vendor–part pair.** `fn_PriceErosion` works out
each pair's first and last price, the years between them, the price the category benchmark says it
should have reached by now, and how much of that erosion was actually captured. The results arrive
in the model as **columns** on `PriceErosion`, and the measures in this section only weight and
total them.

An earlier version of this section presented that arithmetic as DAX measures: `First Price`,
`Last Price`, `Years Elapsed`, `Expected Price`, `Actual Erosion %`, `Erosion Gap per Unit` and
`Erosion Capture % (Spend Weighted)`. None of them exists in the model, and the DAX would not even
have run against it: it read a `POLine[OrderDateKey]` column the model does not load. Here is what
the SQL actually does, summarised:

```text
ILLUSTRATIVE ONLY. Not DAX, and not model measures: this is the per-pair arithmetic in
fn_PriceErosion (sql/05_kpi_views.sql). Each result is a COLUMN on the PriceErosion table.

FirstPrice         volume-weighted unit price over the pair's first 90 days
LastPrice          volume-weighted unit price over the pair's last 90 days
YearsElapsed       days from first to last order / 365.25, rounded to 3 dp
ExpectedPrice      FirstPrice x (1 - BenchmarkErosionPct / 100) ^ YearsElapsed
ActualErosionPct   100 x (1 - (LastPrice / FirstPrice) ^ (365.25 / days)), rounded to 3 dp
ErosionCapturePct  100 x ActualErosionPct / BenchmarkErosionPct
ErosionGapPerUnit  LastPrice - ExpectedPrice
AnnualOpportunity  ErosionGapPerUnit x TTMQty      (negative when ahead of the curve)
```

**Volume-weighted over a 90-day window, not a single transaction.** One small rush order at a bad
price should not set the baseline the whole opportunity is measured from. Below 180 days of history
the two windows would overlap, so `fn_PriceErosion` refuses any shorter pair inside the function
rather than trusting its caller's arguments (`UAT-06`). The model loads it with at least ten lines
and 500 days of history per pair: `fn_PriceErosion(@AsOf, 10, 500)`.

**Rounded where it is published, and derived from the rounded values.** A derived figure must agree
with the components printed beside it. During the SQL build, `ExpectedPrice` was computed from an
unrounded exponent while `YearsElapsed` was published rounded, and a reader redoing the arithmetic
on the page got a different answer. `UAT-05` and `UAT-07` exist because of it.

```dax
Erosion Capture % =
DIVIDE (
    SUMX ( PriceErosion, PriceErosion[ErosionCapturePct] * PriceErosion[TTMSpend] ),
    SUM ( PriceErosion[TTMSpend] )
)
```

**Spend-weighted, never a simple average.** A vendor–part pair carrying $4m of spend and one
carrying $40k should not count equally in a company-level figure, and a plain `AVERAGE` over the
column would say they do. This is the same weighting `fn_SpendKPI` uses for its
`ErosionCapturePct`, which the validator reconciles against.

```dax
Benchmark Erosion % =
DIVIDE (
    SUMX ( PriceErosion, PriceErosion[BenchmarkErosionPct] * PriceErosion[TTMSpend] ),
    SUM ( PriceErosion[TTMSpend] )
)
```

The benchmark the same spend was held to, weighted the same way, so the two can sit side by side
on a page. It is not an average of the category table: a category with little spend should not
move the figure as much as one with most of it.

```dax
Annual Opportunity =
SUMX ( PriceErosion, IF ( PriceErosion[AnnualOpportunity] > 0, PriceErosion[AnnualOpportunity], 0 ) )
```

**Clamped at zero, not netted.** A supplier already beating the curve is not a savings source, and
letting its negative gap offset another supplier's shortfall would understate the actionable
target and imply an offset nobody can realise. The pairs that are ahead are reported separately:

```dax
Value Already Ahead of Curve =
SUMX ( PriceErosion, IF ( PriceErosion[AnnualOpportunity] < 0, PriceErosion[AnnualOpportunity], 0 ) )
```

Both annualise against each pair's trailing-twelve-month quantity, which `fn_PriceErosion` has
already applied, so neither needs a date filter here (§10). The counts of pairs assessed and of
pairs capturing under 20% of their benchmark are in the [complete reference](#complete-reference).

---

## 6. Quality and delivery

```dax
Acceptance Rate % =
DIVIDE ( [Qty Accepted], [Qty Received] ) * 100
```

```dax
Rejected Value =
SUM ( POLine[RejectedValue] )
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
Receipts Booked =
CALCULATE ( COUNTROWS ( POLine ), KEEPFILTERS ( POLine[IsReceived] = TRUE () ) )
```

```dax
On Time Receipts =
CALCULATE (
    COUNTROWS ( POLine ),
    KEEPFILTERS ( POLine[IsOnTime] = TRUE () ),
    KEEPFILTERS ( POLine[IsReceived] = TRUE () )
)
```

```dax
On Time Delivery % =
DIVIDE ( [On Time Receipts] + 0, [Receipts Booked] ) * 100
```

```dax
Avg Days Late =
CALCULATE ( AVERAGE ( POLine[DaysLate] ), KEEPFILTERS ( POLine[IsReceived] = TRUE () ) )
```

**Both sides of the ratio are restricted to lines that have actually been received.** A line still
in transit is neither on time nor late yet; counting it in the denominator only would drag on-time
delivery down every time the order book grew, and counting it in neither is the only reading that
does not move for a reason unrelated to supplier performance. Receipts are pre-aggregated onto the
line (§1a), so `[Receipts Booked]` counts PO lines with a receipt booked by the reporting date, not
individual receipts. The `+ 0` lets a vendor that delivered nothing on time read 0.00%.

`IsOnTime` is `NULL` in SQL for a line with no receipt — not `0`. The model's source query turns it
into a Boolean with `ISNULL(IsOnTime,0)`, which on its own would reclassify every undelivered line
as *late*. That is why the model carries `IsReceived` alongside it and every delivery measure
filters on it.

```dax
Duplicate PO Lines =
CALCULATE ( COUNTROWS ( POLine ), KEEPFILTERS ( POLine[IsDuplicatePO] = TRUE () ) )
```

```dax
Duplicate PO Spend =
CALCULATE ( [Extended Price], KEEPFILTERS ( POLine[IsDuplicatePO] = TRUE () ) )
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
    CALCULATE ( [Extended Price], KEEPFILTERS ( Dim_Part[IsSingleSource] = TRUE () ) ) + 0,
    [Extended Price]
) * 100
```

`Dim_Part[IsSingleSource]` is computed in the model's source query as `QualifiedSupplierCount = 1`,
the same test `fn_SpendKPI` applies, so a single-source part is defined once and the flag can be
sliced like any other. The `+ 0` makes a vendor with no single-source spend read 0.00%.

```dax
Top 5 Vendor Share % =
VAR VendorSpend = ADDCOLUMNS ( VALUES ( Dim_Vendor[VendorKey] ), "@Spend", [Extended Price] )
VAR Active = FILTER ( VendorSpend, NOT ISBLANK ( [@Spend] ) )
VAR Top5 = TOPN ( 5, Active, [@Spend], DESC )
RETURN
    IF (
        COUNTROWS ( Active ) <= 5,
        BLANK (),
        DIVIDE ( SUMX ( Top5, [@Spend] ), SUMX ( Active, [@Spend] ) ) * 100
    )
```

**Concentration needs more than five vendors to mean anything.** With five or fewer in context the
share is 100% by construction, and its status read Red for a single selected vendor, so the measure
returns BLANK instead. Only vendors *with spend* count: `Dim_Vendor` is not filtered by `POLine`, so
`VALUES ( Dim_Vendor[VendorKey] )` on its own lists all 36 vendors under any category filter, and
the test would never fire.

Just above five vendors the share still has a floor: the top five of *N* vendors hold at least
500 / *N* per cent, so 83.3% of six and 71.4% of seven, both past the 70 warning line. Under a filter
that narrow the status is Red whatever the buying. The value is still true, since six vendors *are*
a concentrated base, but the target is set for the whole vendor base, so read the number rather
than the colour there.

```dax
Leverage Score =
VAR Sources    = SELECTEDVALUE ( Dim_Part[QualifiedSupplierCount] )
VAR QualMonths = SELECTEDVALUE ( Dim_Part[QualificationMonths] )
VAR RevShare   = SELECTEDVALUE ( Dim_Vendor[LumenRevenueSharePct] )
VAR OffContract = SUM ( RenegotiationQueue[OffContractSpend] )
RETURN
    -- A score belongs to ONE part-vendor pair on the queue. At a total, or for
    -- a pair not on it, the SELECTEDVALUEs are BLANK and BLANK <= 2 is TRUE, so
    -- this used to invent a score (33 or 37) for rows that have none.
    IF (
        ISEMPTY ( RenegotiationQueue ) || NOT HASONEVALUE ( Dim_Part[PartKey] ) || NOT HASONEVALUE ( Dim_Vendor[VendorKey] ),
        BLANK (),
          SWITCH ( TRUE (), Sources >= 3, 40, Sources = 2, 26, 4 )
        + SWITCH ( TRUE (), QualMonths <= 2, 20, QualMonths <= 5, 13, QualMonths <= 8, 6, 0 )
        + SWITCH ( TRUE (), RevShare >= 15, 30, RevShare >= 8, 20, RevShare >= 4, 11, 3 )
        + IF ( OffContract > 0, 10, 6 )
    )
```

Identical weights to the SQL implementation in `fn_RenegotiationQueue`, deliberately. The two must
agree or the Power BI report and the Excel workbook will rank the same negotiations differently.
They agree by inspection; the validator does not compare this measure with the queue's own
`LeverageScore` column.

**It is BLANK everywhere except on one part–vendor pair that is on the queue.** The three
`SELECTEDVALUE`s are BLANK at a total or with several parts or vendors in context, and DAX compares
BLANK as 0, so the score used to be invented for rows that have none: 4 + 20 + 3, plus 6 or 10.

**The contract term reads the queue, not `[Maverick Spend]`.** `RenegotiationQueue[OffContractSpend]`
is the pair's off-contract spend over the trailing twelve months with duplicated requisitions
excluded, which is the figure SQL scored. An earlier version of this page used `[Maverick Spend]`,
which answers over all history, duplicates included, and over whatever dates the page has selected.

---

## 7a. The renegotiation queue

`fn_RenegotiationQueue(@AsOf, 8)` ranks every assessed vendor–part pair with spend in the last twelve
months by annual opportunity. It flags `IsThisQuarter` on each buyer's eight highest-ranked pairs,
except the ones whose action is `ACCEPT`. An `ACCEPT` pair in a buyer's top eight still takes one of
the eight places, so that buyer gets fewer than eight flags (BUY-02 and BUY-04 have seven, BUY-06
five). Eight is the number of serious negotiations a category manager can realistically run in a
quarter. A list of 185 pairs is not a plan.

```dax
Queue Opportunity =
SUM ( RenegotiationQueue[AnnualOpportunity] )
```

```dax
This Quarter Opportunity =
CALCULATE (
    SUM ( RenegotiationQueue[AnnualOpportunity] ),
    KEEPFILTERS ( RenegotiationQueue[IsThisQuarter] = TRUE () )
)
```

`[This Quarter Opportunity]`, like `[This Quarter Pairs]`, intersects with an `IsThisQuarter`
slicer: with the slicer on False it is BLANK, because none of the pairs held back is scheduled for
this quarter.

```dax
Quarter Coverage % =
DIVIDE (
    -- plain on purpose: this quarter, whatever an IsThisQuarter slicer says
    CALCULATE ( [Queue Opportunity], RenegotiationQueue[IsThisQuarter] = TRUE () ) + 0,
    -- the whole queue: clear that one column, keep every other filter
    CALCULATE ( [Queue Opportunity], REMOVEFILTERS ( RenegotiationQueue[IsThisQuarter] ) )
) * 100
```

**A ratio that must ignore a slicer on its own table.** It answers *what share of the whole queue
does this quarter's work cover*, and `POWER_BI_BUILD_GUIDE.md` defaults the queue page's
`IsThisQuarter` slicer to True. With the slicer on True, the old ratio of `[This Quarter Opportunity]`
to `[Queue Opportunity]` compared this quarter with itself and read 100.0. So the numerator's
filter stays plain, replacing the slicer, and the denominator clears `IsThisQuarter` alone. A buyer,
category or vendor filter still applies to both sides, so the measure reads each buyer's own
coverage. The `+ 0` makes a slice with queued pairs but none this quarter (12 of the 36 vendors,
and every `ACCEPT` pair) read 0.0% rather than BLANK.

Clearing one column is exact only while `IsThisQuarter` is the sole slicer on a
`RenegotiationQueue` column. Power BI merges slicers on the same table into one combined filter
(auto-exist), so with a second slicer on, say, `ActionCode`, clearing `IsThisQuarter` would leave
`ActionCode` cut down to the actions that have pairs this quarter. The build guide keeps the page's
cards out of the slicer's reach, which avoids the question.

---

## 8. Targets and conditional formatting

Each of the eight scorecard KPIs has a `… Colour` measure (a hex string, for conditional
formatting) and a `… Status` measure (Green, Amber or Red, for text): 16 in all, generated in the
build script from one list, `$ragMetrics`. Three metrics are higher-is-better (erosion capture,
acceptance, on-time delivery) and five lower-is-better (maverick spend, PPV, expedite spend,
single-source spend, top-5 share). The pair for erosion capture is shown here; the other fourteen
are in the [complete reference](#complete-reference) and differ only in the value they read, the
metric name and the direction of the comparison.

> **`Target Value` and `Warning Value` are not measures.** They are the `TargetValue` and
> `WarningValue` **columns** of `Ref_SpendTargets`, read inside each Colour and Status measure. An
> earlier version of this section showed them as two stand-alone measures for readability; they
> were never in the model, and have been removed from this page.

```dax
Erosion Capture Colour =
VAR V = [Erosion Capture %]
VAR T = CALCULATE ( VALUES ( Ref_SpendTargets[TargetValue] ),  REMOVEFILTERS ( Ref_SpendTargets ), Ref_SpendTargets[MetricName] = "ErosionCapturePct" )
VAR W = CALCULATE ( VALUES ( Ref_SpendTargets[WarningValue] ), REMOVEFILTERS ( Ref_SpendTargets ), Ref_SpendTargets[MetricName] = "ErosionCapturePct" )
RETURN
    IF (
        ISBLANK ( V ) || ISBLANK ( T ) || ISBLANK ( W ),
        BLANK (),
        SWITCH ( TRUE (), V >= T, "#C6EFCE", V >= W, "#FFEB9C", "#FFC7CE" )
    )
```

```dax
Erosion Capture Status =
VAR V = [Erosion Capture %]
VAR T = CALCULATE ( VALUES ( Ref_SpendTargets[TargetValue] ),  REMOVEFILTERS ( Ref_SpendTargets ), Ref_SpendTargets[MetricName] = "ErosionCapturePct" )
VAR W = CALCULATE ( VALUES ( Ref_SpendTargets[WarningValue] ), REMOVEFILTERS ( Ref_SpendTargets ), Ref_SpendTargets[MetricName] = "ErosionCapturePct" )
RETURN
    IF (
        ISBLANK ( V ) || ISBLANK ( T ) || ISBLANK ( W ),
        BLANK (),
        SWITCH ( TRUE (), V >= T, "Green", V >= W, "Amber", "Red" )
    )
```

**A colour measure returns a hex string, not a number.** Power BI's conditional formatting by
field value expects a colour string; returning 1, 2 or 3 and hoping the rule interprets it is the
most common reason a formatting rule silently does nothing. Each measure reads its own thresholds
from the reference table.

**BLANK in, BLANK out.** A vendor with no erosion pairs has no capture figure; without the guard,
DAX would compare that BLANK as 0 and paint it Red. The thresholds are read with `REMOVEFILTERS`,
so no filter on `Ref_SpendTargets` can blank them. PPV's target is 0.00, and a threshold of zero is
not BLANK, so the guard leaves it alone.

**Every other pair binds to the TTM variant, not the base ratio.** Colouring an all-history figure
against a TTM target is a red light with no meaning (§10). Erosion capture is the exception because
it has no TTM twin: the erosion model already works on trailing-twelve-month spend.

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

```dax
As Of Date =
MAX ( Ref_Reporting[AsOfDate] )
```

```dax
TTM Start =
MAX ( Ref_Reporting[TTMStart] )
```

`TTMStart` is the day after the as-of date minus twelve months, so the inclusive window from
`[TTM Start]` to `[As Of Date]` is exactly SQL's `OrderDate > DATEADD(MONTH, -12, @AsOf)`. The
validator reconciles both anchors against SQL.

The model expresses that window **twice over**, on purpose.

**Base ratios stay filter-responsive.** `[Maverick Spend %]` answers over whatever the user has
sliced. That is what makes a monthly trend line or a per-vendor breakdown mean anything, and it is
the correct default for a semantic model.

**The headline KPIs get an explicit TTM twin.** There are ten, all of the same form; the other nine
are in the [complete reference](#complete-reference).

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

This filter deliberately replaces what the report has selected, and it replaces only the dates: `Dim_Date` is marked as a date table, so a filter on its `Date` column
clears its year, quarter and month columns too, and touches nothing else. `[Acceptance Rate % (TTM)]`
therefore still answers "how is **this vendor** doing" when a vendor is selected — it just refuses
to answer it over a window other than the published one. That is why the RAG colour measures bind
to the TTM variants: colouring an all-history figure against a TTM target is a red light with no
meaning.

**Why the window comes from a table rather than from the data.** The first draft, never built,
was:

```text
ILLUSTRATIVE ONLY: the first draft of Spend (TTM), shown for the bug it had. Not a model measure.

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
measured over. `PriceErosion` has no date relationship in any case (§1), so a date slicer leaves
these measures where they are.

---

## 11. Verifying this document is true

Two scripts check this page and the model, so neither has to be taken on trust.

- **`Sync_Measures.ps1 -DocOnly`** reads every measure out of `Build_PowerBI_Model.ps1` and fails
  unless each `dax` block on this page matches one of them token for token, every measure appears
  exactly once, and no block defines anything else. Layout and `--` comments are ignored; names,
  strings, operators and numbers are not. Without `-DocOnly` it also compares the `.pbix` open in
  Power BI Desktop, and `-Apply` brings that model into line with the script.
- **`Validate_PowerBI_Model.ps1`** queries the live model in DAX and reconciles it against the SQL
  that produces each figure: the row counts, the all-history totals, the TTM scorecard, the erosion
  figures, the queue, the reporting anchor, the cost index for every vendor and every buyer, and the
  queue for every buyer and row by row. Run it after any change here. It proves Power BI and SQL agree, not that either is right.

A measure file that has drifted from the model is exactly as misleading as the model-shape section
this document had to have rewritten.

---

## Complete reference

Every measure the text above does not show, grouped by the display folder it sits in. With the
blocks above, this covers all 69 measures in the model exactly once.

### 03 Erosion

```dax
Pairs Assessed =
COUNTROWS ( PriceErosion )
```

```dax
Pairs Capturing Under 20% =
CALCULATE ( COUNTROWS ( PriceErosion ), KEEPFILTERS ( PriceErosion[ErosionCapturePct] < 20 ) )
```

```dax
Spend on Flat Pairs =
CALCULATE ( SUM ( PriceErosion[TTMSpend] ), KEEPFILTERS ( PriceErosion[ErosionCapturePct] < 20 ) )
```

### 05 Delivery

```dax
Expedite Spend % =
DIVIDE ( [Expedite Fees], [Extended Price] ) * 100
```

### 07 Queue

```dax
Queue Pairs =
COUNTROWS ( RenegotiationQueue )
```

```dax
This Quarter Pairs =
CALCULATE ( COUNTROWS ( RenegotiationQueue ), KEEPFILTERS ( RenegotiationQueue[IsThisQuarter] = TRUE () ) )
```

### 08 Status

Lower-is-better metrics compare with `<=`, higher-is-better with `>=`. Each reads the TTM variant
of its KPI.

```dax
Maverick Spend Colour =
VAR V = [Maverick Spend % (TTM)]
VAR T = CALCULATE ( VALUES ( Ref_SpendTargets[TargetValue] ),  REMOVEFILTERS ( Ref_SpendTargets ), Ref_SpendTargets[MetricName] = "MaverickSpendPct" )
VAR W = CALCULATE ( VALUES ( Ref_SpendTargets[WarningValue] ), REMOVEFILTERS ( Ref_SpendTargets ), Ref_SpendTargets[MetricName] = "MaverickSpendPct" )
RETURN IF ( ISBLANK ( V ) || ISBLANK ( T ) || ISBLANK ( W ), BLANK (), SWITCH ( TRUE (), V <= T, "#C6EFCE", V <= W, "#FFEB9C", "#FFC7CE" ) )
```

```dax
Maverick Spend Status =
VAR V = [Maverick Spend % (TTM)]
VAR T = CALCULATE ( VALUES ( Ref_SpendTargets[TargetValue] ),  REMOVEFILTERS ( Ref_SpendTargets ), Ref_SpendTargets[MetricName] = "MaverickSpendPct" )
VAR W = CALCULATE ( VALUES ( Ref_SpendTargets[WarningValue] ), REMOVEFILTERS ( Ref_SpendTargets ), Ref_SpendTargets[MetricName] = "MaverickSpendPct" )
RETURN IF ( ISBLANK ( V ) || ISBLANK ( T ) || ISBLANK ( W ), BLANK (), SWITCH ( TRUE (), V <= T, "Green", V <= W, "Amber", "Red" ) )
```

```dax
PPV Colour =
VAR V = [PPV % (TTM)]
VAR T = CALCULATE ( VALUES ( Ref_SpendTargets[TargetValue] ),  REMOVEFILTERS ( Ref_SpendTargets ), Ref_SpendTargets[MetricName] = "PPVPct" )
VAR W = CALCULATE ( VALUES ( Ref_SpendTargets[WarningValue] ), REMOVEFILTERS ( Ref_SpendTargets ), Ref_SpendTargets[MetricName] = "PPVPct" )
RETURN IF ( ISBLANK ( V ) || ISBLANK ( T ) || ISBLANK ( W ), BLANK (), SWITCH ( TRUE (), V <= T, "#C6EFCE", V <= W, "#FFEB9C", "#FFC7CE" ) )
```

```dax
PPV Status =
VAR V = [PPV % (TTM)]
VAR T = CALCULATE ( VALUES ( Ref_SpendTargets[TargetValue] ),  REMOVEFILTERS ( Ref_SpendTargets ), Ref_SpendTargets[MetricName] = "PPVPct" )
VAR W = CALCULATE ( VALUES ( Ref_SpendTargets[WarningValue] ), REMOVEFILTERS ( Ref_SpendTargets ), Ref_SpendTargets[MetricName] = "PPVPct" )
RETURN IF ( ISBLANK ( V ) || ISBLANK ( T ) || ISBLANK ( W ), BLANK (), SWITCH ( TRUE (), V <= T, "Green", V <= W, "Amber", "Red" ) )
```

```dax
Acceptance Colour =
VAR V = [Acceptance Rate % (TTM)]
VAR T = CALCULATE ( VALUES ( Ref_SpendTargets[TargetValue] ),  REMOVEFILTERS ( Ref_SpendTargets ), Ref_SpendTargets[MetricName] = "AcceptanceRatePct" )
VAR W = CALCULATE ( VALUES ( Ref_SpendTargets[WarningValue] ), REMOVEFILTERS ( Ref_SpendTargets ), Ref_SpendTargets[MetricName] = "AcceptanceRatePct" )
RETURN IF ( ISBLANK ( V ) || ISBLANK ( T ) || ISBLANK ( W ), BLANK (), SWITCH ( TRUE (), V >= T, "#C6EFCE", V >= W, "#FFEB9C", "#FFC7CE" ) )
```

```dax
Acceptance Status =
VAR V = [Acceptance Rate % (TTM)]
VAR T = CALCULATE ( VALUES ( Ref_SpendTargets[TargetValue] ),  REMOVEFILTERS ( Ref_SpendTargets ), Ref_SpendTargets[MetricName] = "AcceptanceRatePct" )
VAR W = CALCULATE ( VALUES ( Ref_SpendTargets[WarningValue] ), REMOVEFILTERS ( Ref_SpendTargets ), Ref_SpendTargets[MetricName] = "AcceptanceRatePct" )
RETURN IF ( ISBLANK ( V ) || ISBLANK ( T ) || ISBLANK ( W ), BLANK (), SWITCH ( TRUE (), V >= T, "Green", V >= W, "Amber", "Red" ) )
```

```dax
On Time Delivery Colour =
VAR V = [On Time Delivery % (TTM)]
VAR T = CALCULATE ( VALUES ( Ref_SpendTargets[TargetValue] ),  REMOVEFILTERS ( Ref_SpendTargets ), Ref_SpendTargets[MetricName] = "OnTimeDeliveryPct" )
VAR W = CALCULATE ( VALUES ( Ref_SpendTargets[WarningValue] ), REMOVEFILTERS ( Ref_SpendTargets ), Ref_SpendTargets[MetricName] = "OnTimeDeliveryPct" )
RETURN IF ( ISBLANK ( V ) || ISBLANK ( T ) || ISBLANK ( W ), BLANK (), SWITCH ( TRUE (), V >= T, "#C6EFCE", V >= W, "#FFEB9C", "#FFC7CE" ) )
```

```dax
On Time Delivery Status =
VAR V = [On Time Delivery % (TTM)]
VAR T = CALCULATE ( VALUES ( Ref_SpendTargets[TargetValue] ),  REMOVEFILTERS ( Ref_SpendTargets ), Ref_SpendTargets[MetricName] = "OnTimeDeliveryPct" )
VAR W = CALCULATE ( VALUES ( Ref_SpendTargets[WarningValue] ), REMOVEFILTERS ( Ref_SpendTargets ), Ref_SpendTargets[MetricName] = "OnTimeDeliveryPct" )
RETURN IF ( ISBLANK ( V ) || ISBLANK ( T ) || ISBLANK ( W ), BLANK (), SWITCH ( TRUE (), V >= T, "Green", V >= W, "Amber", "Red" ) )
```

```dax
Expedite Spend Colour =
VAR V = [Expedite Spend % (TTM)]
VAR T = CALCULATE ( VALUES ( Ref_SpendTargets[TargetValue] ),  REMOVEFILTERS ( Ref_SpendTargets ), Ref_SpendTargets[MetricName] = "ExpediteSpendPct" )
VAR W = CALCULATE ( VALUES ( Ref_SpendTargets[WarningValue] ), REMOVEFILTERS ( Ref_SpendTargets ), Ref_SpendTargets[MetricName] = "ExpediteSpendPct" )
RETURN IF ( ISBLANK ( V ) || ISBLANK ( T ) || ISBLANK ( W ), BLANK (), SWITCH ( TRUE (), V <= T, "#C6EFCE", V <= W, "#FFEB9C", "#FFC7CE" ) )
```

```dax
Expedite Spend Status =
VAR V = [Expedite Spend % (TTM)]
VAR T = CALCULATE ( VALUES ( Ref_SpendTargets[TargetValue] ),  REMOVEFILTERS ( Ref_SpendTargets ), Ref_SpendTargets[MetricName] = "ExpediteSpendPct" )
VAR W = CALCULATE ( VALUES ( Ref_SpendTargets[WarningValue] ), REMOVEFILTERS ( Ref_SpendTargets ), Ref_SpendTargets[MetricName] = "ExpediteSpendPct" )
RETURN IF ( ISBLANK ( V ) || ISBLANK ( T ) || ISBLANK ( W ), BLANK (), SWITCH ( TRUE (), V <= T, "Green", V <= W, "Amber", "Red" ) )
```

```dax
Single Source Colour =
VAR V = [Single Source Spend % (TTM)]
VAR T = CALCULATE ( VALUES ( Ref_SpendTargets[TargetValue] ),  REMOVEFILTERS ( Ref_SpendTargets ), Ref_SpendTargets[MetricName] = "SingleSourceSpendPct" )
VAR W = CALCULATE ( VALUES ( Ref_SpendTargets[WarningValue] ), REMOVEFILTERS ( Ref_SpendTargets ), Ref_SpendTargets[MetricName] = "SingleSourceSpendPct" )
RETURN IF ( ISBLANK ( V ) || ISBLANK ( T ) || ISBLANK ( W ), BLANK (), SWITCH ( TRUE (), V <= T, "#C6EFCE", V <= W, "#FFEB9C", "#FFC7CE" ) )
```

```dax
Single Source Status =
VAR V = [Single Source Spend % (TTM)]
VAR T = CALCULATE ( VALUES ( Ref_SpendTargets[TargetValue] ),  REMOVEFILTERS ( Ref_SpendTargets ), Ref_SpendTargets[MetricName] = "SingleSourceSpendPct" )
VAR W = CALCULATE ( VALUES ( Ref_SpendTargets[WarningValue] ), REMOVEFILTERS ( Ref_SpendTargets ), Ref_SpendTargets[MetricName] = "SingleSourceSpendPct" )
RETURN IF ( ISBLANK ( V ) || ISBLANK ( T ) || ISBLANK ( W ), BLANK (), SWITCH ( TRUE (), V <= T, "Green", V <= W, "Amber", "Red" ) )
```

```dax
Top 5 Share Colour =
VAR V = [Top 5 Vendor Share % (TTM)]
VAR T = CALCULATE ( VALUES ( Ref_SpendTargets[TargetValue] ),  REMOVEFILTERS ( Ref_SpendTargets ), Ref_SpendTargets[MetricName] = "Top5VendorSharePct" )
VAR W = CALCULATE ( VALUES ( Ref_SpendTargets[WarningValue] ), REMOVEFILTERS ( Ref_SpendTargets ), Ref_SpendTargets[MetricName] = "Top5VendorSharePct" )
RETURN IF ( ISBLANK ( V ) || ISBLANK ( T ) || ISBLANK ( W ), BLANK (), SWITCH ( TRUE (), V <= T, "#C6EFCE", V <= W, "#FFEB9C", "#FFC7CE" ) )
```

```dax
Top 5 Share Status =
VAR V = [Top 5 Vendor Share % (TTM)]
VAR T = CALCULATE ( VALUES ( Ref_SpendTargets[TargetValue] ),  REMOVEFILTERS ( Ref_SpendTargets ), Ref_SpendTargets[MetricName] = "Top5VendorSharePct" )
VAR W = CALCULATE ( VALUES ( Ref_SpendTargets[WarningValue] ), REMOVEFILTERS ( Ref_SpendTargets ), Ref_SpendTargets[MetricName] = "Top5VendorSharePct" )
RETURN IF ( ISBLANK ( V ) || ISBLANK ( T ) || ISBLANK ( W ), BLANK (), SWITCH ( TRUE (), V <= T, "Green", V <= W, "Amber", "Red" ) )
```

### 09 TTM Scorecard

All nine share one form with `[Maverick Spend % (TTM)]` in §10: the base measure, inside the
published window.

```dax
Spend (TTM) =
CALCULATE (
    [Extended Price],
    DATESBETWEEN ( Dim_Date[Date], MAX ( Ref_Reporting[TTMStart] ), MAX ( Ref_Reporting[AsOfDate] ) )
)
```

```dax
Landed Cost (TTM) =
CALCULATE (
    [Landed Cost],
    DATESBETWEEN ( Dim_Date[Date], MAX ( Ref_Reporting[TTMStart] ), MAX ( Ref_Reporting[AsOfDate] ) )
)
```

```dax
PO Lines (TTM) =
CALCULATE (
    [PO Lines],
    DATESBETWEEN ( Dim_Date[Date], MAX ( Ref_Reporting[TTMStart] ), MAX ( Ref_Reporting[AsOfDate] ) )
)
```

```dax
PPV % (TTM) =
CALCULATE (
    [PPV %],
    DATESBETWEEN ( Dim_Date[Date], MAX ( Ref_Reporting[TTMStart] ), MAX ( Ref_Reporting[AsOfDate] ) )
)
```

```dax
Acceptance Rate % (TTM) =
CALCULATE (
    [Acceptance Rate %],
    DATESBETWEEN ( Dim_Date[Date], MAX ( Ref_Reporting[TTMStart] ), MAX ( Ref_Reporting[AsOfDate] ) )
)
```

```dax
On Time Delivery % (TTM) =
CALCULATE (
    [On Time Delivery %],
    DATESBETWEEN ( Dim_Date[Date], MAX ( Ref_Reporting[TTMStart] ), MAX ( Ref_Reporting[AsOfDate] ) )
)
```

```dax
Expedite Spend % (TTM) =
CALCULATE (
    [Expedite Spend %],
    DATESBETWEEN ( Dim_Date[Date], MAX ( Ref_Reporting[TTMStart] ), MAX ( Ref_Reporting[AsOfDate] ) )
)
```

```dax
Single Source Spend % (TTM) =
CALCULATE (
    [Single Source Spend %],
    DATESBETWEEN ( Dim_Date[Date], MAX ( Ref_Reporting[TTMStart] ), MAX ( Ref_Reporting[AsOfDate] ) )
)
```

```dax
Top 5 Vendor Share % (TTM) =
CALCULATE (
    [Top 5 Vendor Share %],
    DATESBETWEEN ( Dim_Date[Date], MAX ( Ref_Reporting[TTMStart] ), MAX ( Ref_Reporting[AsOfDate] ) )
)
```
