"""
================================================================================
Project 3 - Lumen Optics Manufacturing: Photonics Spend Scorecard
App:     streamlit/app.py
Purpose: The web demo. Six screens, each answering one question the case
         study answers, and each ending in the decision and who makes it.

WHAT THIS IS FOR

    A category manager opens this at the start of a quarter to answer: our
    purchase prices are flat and every variance report is green, so where is
    the money going -- and which supplier-part pairs do we renegotiate first?
    Screen 1 is the contrast the project rests on: purchase price variance
    reads green while erosion capture reads under a third.

NO ANALYTICAL LOGIC LIVES HERE, AND THAT IS DELIBERATE

    Every definition -- erosion capture, annual opportunity, leverage score,
    action code, the this-quarter flag, the vendor cost index -- is computed in
    SQL (sql/05_kpi_views.sql) and arrives in an extract column. This app
    filters, sorts, counts and sums rows SQL already flagged, and draws. Its
    few ratios are of two totals whose result the case study quotes.

    Where the case study quotes a figure at a grain no extract carries, the
    app QUOTES it, citing the section, rather than rebuild it (QUOTED below):
    cost per accepted unit on LOM-0072 by order year, VEN-009's on-time rate
    by order year, and erosion capture spend-weighted by sourcing. The first
    needs duplicate PO lines removed the way fn_DuplicatePOLines removes them,
    and po_line_cost.csv carries no duplicate flag; the others would be new
    aggregations of SQL's definitions. Rebuilt in pandas, each would be a
    second definition free to drift from the first.

    Card colours: a card for a scorecard metric is drawn in the status the
    case study publishes for it (PUBLISHED_STATUS -- quoted, never computed;
    dbo.usp_SpendScorecard computes it and no extract carries it). Every other
    card is grey, because nothing computes a status for it.

DATA DISCLOSURE
    Lumen Optics Manufacturing is fictional and all data is synthetic. No
    confidential data is used and no claim is made about any production system.
================================================================================
"""

from __future__ import annotations

from pathlib import Path

import pandas as pd
import streamlit as st

EXTRACTS = Path(__file__).resolve().parent.parent / "erp_extracts"

st.set_page_config(page_title="Lumen Optics - Photonics Spend Scorecard", layout="wide")

RED, AMBER, GREEN, GREY = "#c00000", "#bf8f00", "#548235", "#808080"
RAG = {"Red": RED, "Amber": AMBER, "Green": GREEN, "Grey": GREY}

# The case study's own exhibits (sections 4.2 and 4.4). These name WHAT to
# show, not a rule for choosing it -- finding the worst cohort is SQL's job.
EXHIBIT_PART = "LOM-0072"
COHORT_VENDOR, COHORT_CATEGORY = "VEN-004", "Laser Diode"
DELIVERY_VENDOR = "VEN-009"

# The data-quality gate's tolerance: the @MaxMisstatementPctOfSpend default of
# dbo.usp_RunDataQualityChecks (sql/04_data_quality_checks.sql). The procedure
# PRINTs its verdict rather than returning it, so no extract carries this number.
GATE_TOLERANCE_PCT = 0.500

# The status the case study's appendix (and section 4.4) publishes for each
# scorecard metric a card shows, at 2025-12-31. Quoted, not derived.
PUBLISHED_STATUS = {"ErosionCapturePct": "Red", "PPVPct": "Green", "AcceptanceRatePct": "Green",
                    "OnTimeDeliveryPct": "Red", "ExpediteSpendPct": "Green"}

# QUOTED, NOT COMPUTED.
# (1) Case study section 4.2: cost per accepted unit on LOM-0072 by order year.
#     SQL's figure is SUM(LandedCost) / SUM(QtyAccepted) per vendor, as in
#     fn_VendorScorecard's PartVendor CTE (sql/05), AFTER fn_DuplicatePOLines
#     (sql/04) removes duplicate PO lines. po_line_cost.csv has no column that
#     flags a duplicate -- only the generator's 'PO-D' prefix, which sql/04
#     refuses to rely on -- so the published figures are quoted, not rebuilt.
LOM_0072_QUOTED = (
    "- **VEN-004**: \\$227.76 in 2023 (the dearest), \\$192.58 in 2024 and \\$168.28 in 2025 "
    "(the cheapest both years); \\$204.71 over all history. Average unit price \\$184.14, the "
    "cheapest of the three; acceptance 91.61%.\n"
    "- **VEN-012**: \\$212.00 in 2023, \\$204.35 in 2024, \\$189.42 in 2025; \\$204.44 over all "
    "history. Average unit price \\$197.43; acceptance 98.99%.\n"
    "- **VEN-009**: \\$223.59 in 2023, no orders in 2024, \\$193.55 in 2025; \\$213.94 over all "
    "history. Average unit price \\$211.49; acceptance 99.09%."
)
# (2) Case study sections 4.4 and 6 (Recommendation 6): VEN-009 on time by order year.
VEN_009_BY_YEAR = "83% of its 2023 orders and 20.8% of 2025's"
# (3) Case study section 4.1: erosion capture, spend-weighted, by sourcing.
CAPTURE_SOLE, CAPTURE_COMPETITIVE = "8.91%", "38.26%"

SCREENS = ["Green, and the money is gone", "Where the gap sits", "Price is not cost",
           "Green lights, failing cohorts", "The queue", "Can the data carry it?"]


@st.cache_data(show_spinner="Loading the committed extracts...")
def load_all() -> dict[str, pd.DataFrame]:
    files = {
        "kpi": "spend_kpi_monthly.csv", "targets": "spend_targets.csv",
        "queue": "renegotiation_queue.csv", "erosion": "price_erosion.csv",
        "benchmark": "erosion_benchmark.csv", "vendors": "vendor_scorecard.csv",
        "po": "po_line_cost.csv", "category": "category_summary.csv",
        "dq": "data_quality_summary.csv",
    }
    return {key: pd.read_csv(EXTRACTS / name, encoding="utf-8-sig") for key, name in files.items()}


def metric_card(column, label: str, value: str, rag: str, sub: str = "") -> None:
    colour = RAG.get(rag, GREY)
    column.markdown(
        f"""<div style="border-left:6px solid {colour};padding:0.4rem 0 0.4rem 0.9rem;margin-bottom:0.6rem">
        <div style="font-size:0.78rem;color:#555;text-transform:uppercase;letter-spacing:0.04em">{label}</div>
        <div style="font-size:2.0rem;font-weight:700;color:{colour};line-height:1.1">{value}</div>
        <div style="font-size:0.78rem;color:#666">{sub}</div></div>""",
        unsafe_allow_html=True,
    )


def usd(value: float, decimals: int = 0) -> str:
    """Dollars for a metric card. The card is raw HTML, where a $ is literal."""
    return f"${value:,.{decimals}f}"


def usd_md(value: float, decimals: int = 0) -> str:
    """Dollars inside markdown text, where two bare $ signs can start LaTeX."""
    return "\\" + usd(value, decimals)


def decision(*recommendations: tuple[str, str, str]) -> None:
    st.subheader("The decision")
    for action, owner, why in recommendations:
        st.success(f"**{action}**  \nOwner: {owner}.  \n{why}")


data = load_all()
kpi = data["kpi"].sort_values("AsOfDate")
now = kpi.iloc[-1]
as_of = str(now["AsOfDate"])
targets = data["targets"].set_index("MetricName")
queue = data["queue"].sort_values("PriorityRank")
gap_total = float(queue["AnnualOpportunity"].sum())
vendors = data["vendors"]
vend = vendors.set_index("VendorID")

st.sidebar.title("Lumen Optics")
st.sidebar.caption("Direct-materials sourcing control")
st.sidebar.markdown(f"**Reporting date** {as_of}")
st.sidebar.caption("Source: committed CSV extracts (erp_extracts/), exported from SQL Server")
st.sidebar.markdown("---")
st.sidebar.caption("Lumen Optics Manufacturing is fictional. All data is synthetic and generated "
                   "by the scripts in `sql/`. No confidential data is used and no claim is made "
                   "about any production system.")
st.sidebar.caption("Card colours: a scorecard metric is drawn in the status the case study "
                   "publishes for it. Every other card is grey: nothing computes a status for it.")

page = st.sidebar.radio("Screen", SCREENS)


# =============================================================================
# 1. GREEN, AND THE MONEY IS GONE
# =============================================================================
if page == SCREENS[0]:
    st.title("Every variance report is green. So where is the money going?")
    st.markdown(
        "Purchase price variance compares what was paid with what was **agreed**. Erosion "
        "capture compares it with what the category's learning curve says the part **should "
        "cost by now**. Same purchase orders, same day - two different questions, and only "
        "one of them can see the loss."
    )

    spend, gap = float(now["TotalSpend"]), float(now["ErosionOpportunity"])
    c1, c2, c3 = st.columns(3)
    metric_card(c1, "Purchase price variance", f"{now['PPVPct']:.2f}%", PUBLISHED_STATUS["PPVPct"],
                "What procurement reports every month")
    metric_card(c2, "Erosion capture", f"{now['ErosionCapturePct']:.2f}%",
                PUBLISHED_STATUS["ErosionCapturePct"],
                f"Target {targets.at['ErosionCapturePct', 'TargetValue']:.2f}% - share of the "
                "expected price decline actually paid")
    metric_card(c3, "Annual value of the gap", usd(gap), "Grey",
                f"{100 * gap / spend:.1f}% of {usd(spend, 2)} direct-materials spend (12 months)")

    bench = data["benchmark"].set_index("Category")["AnnualErosionPct"]
    st.info(
        f"**Optical components follow learning curves.** The benchmark expects laser diodes to "
        f"fall about {bench['Laser Diode']:.0f}% a year, fibre connectors "
        f"{bench['Fibre Connector']:.0f}%, optical coatings {bench['Optical Coating']:.0f}% and "
        f"optical fibre {bench['Optical Fibre']:.0f}%. A supplier who renews a contract **flat** "
        "keeps everything the curve was supposed to hand back - and because the price never "
        "*rose*, purchase price variance records zero. Forever."
    )

    st.subheader(f"Both measures, {len(kpi)} month-ends")
    st.line_chart(kpi.set_index("AsOfDate")[["PPVPct", "ErosionCapturePct"]], height=280)
    st.caption("Capture is blank before June 2024: fn_PriceErosion judges a pair only once it has "
               "10 order lines spanning 500 days, and orders begin in January 2023.")

    st.subheader(f"The whole scorecard at {as_of}")
    board = data["targets"][["MetricName", "TargetValue", "WarningValue", "Direction",
                             "Description"]].copy()
    board.insert(1, "Value", board["MetricName"].map(now).astype(float).round(2))
    st.dataframe(board, width="stretch", hide_index=True)
    st.caption("dbo.fn_SpendKPI at the reporting date. No Status column, on purpose: "
               "red/amber/green is computed inside "
               "dbo.usp_SpendScorecard, which is not exported. Read Value against Target and "
               "Warning in the stated Direction.")

    decision((
        "Recommendation 3: add erosion capture to the monthly pack, beside purchase price variance",
        "the CPO",
        f"PPV reads {now['PPVPct']:.2f}% while {usd_md(gap)} a year of expected price decline is "
        "not being collected, and no report in the business shows it. It costs nothing but a "
        "decision.",
    ))


# =============================================================================
# 2. WHERE THE GAP SITS
# =============================================================================
elif page == SCREENS[1]:
    label = {0: "Parts with 2+ qualified suppliers", 1: "Parts with one qualified supplier"}
    groups = (queue.groupby("IsSingleSource")
                   .agg(Pairs=("PartNumber", "size"), Opportunity=("AnnualOpportunity", "sum"),
                        AvgLeverage=("LeverageScore", "mean"))
                   .reset_index())
    groups.insert(0, "Sourcing", groups["IsSingleSource"].map(label))
    # A ratio of the two columns beside it: case study section 5, "per pair, the
    # largest gaps sit on sole-source parts".
    groups["OpportunityPerPair"] = groups["Opportunity"] / groups["Pairs"]
    sole = groups[groups["IsSingleSource"] == 1].iloc[0]
    comp = groups[groups["IsSingleSource"] == 0].iloc[0]

    st.title("Per pair, the gap is largest where leverage is weakest")
    st.markdown(
        "Every row below is a vendor-part pair with enough history to judge. **LeverageScore** "
        "(0-100, computed in SQL) combines the four things that decide a photonics negotiation: "
        "how many qualified suppliers exist, how long requalification takes, how much of the "
        "supplier's revenue Lumen represents, and whether there is a contract to reopen."
    )

    c1, c2, c3, c4 = st.columns(4)
    metric_card(c1, "Vendor-part pairs assessed", f"{len(data['erosion'])}", "Grey",
                f"{usd(gap_total)} of annual opportunity between them")
    metric_card(c2, "Sole-source pairs", f"{sole['Pairs']}", "Grey",
                f"carry {usd(sole['Opportunity'])} - {100 * sole['Opportunity'] / gap_total:.0f}% "
                "of the gap")
    metric_card(c3, "Gap per sole-source pair", usd(sole["OpportunityPerPair"]), "Grey",
                f"against {usd(comp['OpportunityPerPair'])} where an alternative exists")
    metric_card(c4, "Average leverage score", f"{sole['AvgLeverage']:.1f}", "Grey",
                f"on sole-source parts, against {comp['AvgLeverage']:.1f} where an alternative exists")
    st.dataframe(groups[["Sourcing", "Pairs", "Opportunity", "OpportunityPerPair", "AvgLeverage"]]
                 .round({"Opportunity": 0, "OpportunityPerPair": 0, "AvgLeverage": 1}),
                 width="stretch", hide_index=True)
    st.caption(
        f"Suppliers hold price precisely where Lumen cannot leave: spend-weighted, sole-source "
        f"pairs capture {CAPTURE_SOLE} of the expected erosion against {CAPTURE_COMPETITIVE} "
        "where an alternative exists (case study, section 4.1). No extract carries that "
        "weighting, so the two figures are quoted, not recomputed."
    )

    st.subheader("Annual opportunity by category")
    by_cat = (queue.assign(Sourcing=queue["IsSingleSource"].map(label))
                   .pivot_table(index="Category", columns="Sourcing", values="AnnualOpportunity",
                                aggfunc="sum", fill_value=0))
    st.bar_chart(by_cat, height=300)

    st.subheader("The benchmark every pair is judged against")
    st.dataframe(data["benchmark"], width="stretch", hide_index=True)
    st.caption("The most contestable input, so it lives in Ref_PriceErosionBenchmark with a "
               "SourceNote on every row, where a category manager can challenge and change it.")

    st.subheader("The sole-source pairs, in SQL's priority order")
    sole_pairs = queue[queue["IsSingleSource"] == 1]
    st.dataframe(sole_pairs[
        ["PriorityRank", "PartNumber", "Category", "VendorID", "VendorName", "QualificationMonths",
         "LeverageScore", "BenchmarkErosionPct", "ActualErosionPct", "ErosionCapturePct",
         "TTMSpend", "AnnualOpportunity", "ActionCode"]], width="stretch", hide_index=True)

    dual = queue[queue["ActionCode"] == "DUAL_SOURCE"]
    dual_tq = int(dual["IsThisQuarter"].sum())
    top = sole_pairs.iloc[0]  # PriorityRank is by money, so the first row is the largest gap
    decision((
        f"Recommendation 5: qualify a second source on the {len(dual)} DUAL_SOURCE pairs",
        "category managers",
        f"Sole-source capture is {CAPTURE_SOLE} against {CAPTURE_COMPETITIVE} where alternatives "
        "exist. SQL routes a pair to DUAL_SOURCE when it is sole source, carries at least "
        "\\$40,000 a year of opportunity, can be requalified within 8 months, and has not "
        "already been routed to a quality fix, a contract or a renegotiation: "
        f"{len(dual)} pairs, {usd_md(dual['AnnualOpportunity'].sum())} a year, "
        f"{'all ' + str(dual_tq) if dual_tq == len(dual) else str(dual_tq)} flagged for this "
        f"quarter. The largest sole-source gap, {top['PartNumber']} at "
        f"{usd_md(top['AnnualOpportunity'])}, is not one of them: an earlier rule routes it to "
        f"{top['ActionCode']}.",
    ))


# =============================================================================
# 3. PRICE IS NOT COST
# =============================================================================
elif page == SCREENS[2]:
    st.title("The cheapest price was the dearest usable part - until the price fell further")
    st.markdown(
        "A component that arrives outside specification was still paid for and still freighted, "
        "and cannot be built into a module. Unit price cannot see any of that. **Cost per "
        "accepted unit** - landed cost divided by the quantity that passed incoming inspection - "
        "is the only comparison that can."
    )

    unusable = float(vendors["RejectedValue"].sum())  # SQL's per-vendor totals, summed
    own = float(vend.at[COHORT_VENDOR, "RejectedValue"])
    c1, c2 = st.columns(2)
    metric_card(c1, "Material paid for and unusable", usd(unusable), "Grey",
                f"Rejected at inspection, all {len(vendors)} vendors, all history")
    metric_card(c2, f"Of which {COHORT_VENDOR}", usd(own), "Grey",
                f"{100 * own / unusable:.0f}% of it, from one supplier")

    st.subheader(f"Part {EXHIBIT_PART}, three qualified suppliers: cost per accepted unit by order year")
    st.markdown(LOM_0072_QUOTED)
    st.caption("Quoted from the case study, section 4.2. SQL computes it as fn_VendorScorecard's "
               "PartVendor CTE does - landed cost over accepted quantity - after "
               "fn_DuplicatePOLines removes duplicate PO lines. po_line_cost.csv carries no "
               "duplicate flag, so this app could not exclude them as SQL does, and does not try.")
    st.markdown(
        f"In 2023 {COHORT_VENDOR} was the cheapest per unit on {EXHIBIT_PART} and the **dearest "
        "per usable part**: it rejected 9.2% of what it shipped that year, and on all three "
        "laser-diode parts it shares with another supplier it cost more per accepted unit than "
        "the best alternative. Over all history it rejects 8.39% on this part against 1.00% for "
        "the other two suppliers - 8.4 times the rate. Its prices "
        "then fell faster than its rejection rate moved, and in 2024 and 2025 it was the "
        "**cheapest per accepted unit on all three** - on this part in 2025, by \\$21 a unit. "
        "Over all history the two effects nearly cancel: \\$204.71 against VEN-012's \\$204.44 "
        "is a blend of one bad year and two good ones, not a finding."
    )
    st.info("**Compare suppliers on cost per accepted unit over a recent window** - never on unit "
            "price, which got 2023 wrong, and never on all history, which calls 2025 a 27-cent tie.")

    st.subheader("Every vendor, all history")
    st.dataframe(vendors.sort_values("CostIndexVsBest", ascending=False)[
        ["VendorID", "VendorName", "VendorTier", "Lines", "TotalSpend", "AcceptanceRatePct",
         "RejectedValue", "OnTimePct", "CostIndexVsBest", "SharedParts"]], width="stretch", hide_index=True)
    st.caption("fn_VendorScorecard: all history, duplicate PO lines excluded. CostIndexVsBest "
               "(100 = the best cost per accepted unit on the same parts) blends every year - the "
               "comparison the lesson above warns against - so read it as a position, not a "
               "verdict. VendorID sits beside every name because VEN-024 and VEN-036 are both "
               "called Aperture Technologies.")
    st.subheader("Rejected material, by vendor")
    st.bar_chart(vendors.set_index("VendorID")["RejectedValue"], height=260)

    decision((
        f"Recommendation 4: keep {COHORT_VENDOR} on laser diodes, and put its rejection rate "
        "under corrective action", "Optoelectronics, with supplier quality",
        "Cheapest per accepted unit on all three shared laser-diode parts in 2024 and 2025, but "
        f"rejecting 8.4 times as often as its competitors on {EXHIBIT_PART}. {usd_md(own)} of "
        f"the {usd_md(unusable)} of unusable material is {COHORT_VENDOR}'s; fixing it makes the "
        "best-value supplier on these parts cheaper still.",
    ))


# =============================================================================
# 4. GREEN LIGHTS, FAILING COHORTS
# =============================================================================
elif page == SCREENS[3]:
    st.title("Portfolio figures hide the cohorts that are failing")
    st.markdown(
        "A portfolio average cannot breach until the whole book does, and by then the cohort "
        "that caused it has been failing for a year. The top row is the scorecard, in the status "
        "the case study publishes; under each figure is the cut it hides."
    )

    # One vendor on one category: SUM(QtyAccepted) / SUM(QtyReceived) over the
    # rows SQL produced -- the aggregation dbo.usp_SpendScorecard runs over
    # fn_POLineCost to find its worst vendor-category cohort (sql/06,
    # @WorstAcceptPct). Duplicate PO lines have no receipt, so add to neither sum.
    po = data["po"]
    cohort_lines = po[(po["VendorID"] == COHORT_VENDOR) & (po["Category"] == COHORT_CATEGORY)]
    cohort = 100 * cohort_lines["QtyAccepted"].sum() / cohort_lines["QtyReceived"].sum()
    first = kpi.iloc[0]

    st.subheader(f"The portfolio, trailing 12 months to {as_of}")
    c1, c2, c3 = st.columns(3)
    for col, metric, name in [(c1, "AcceptanceRatePct", "Acceptance rate"),
                              (c2, "OnTimeDeliveryPct", "On-time delivery"),
                              (c3, "ExpediteSpendPct", "Expedite spend")]:
        metric_card(col, name, f"{now[metric]:.2f}%", PUBLISHED_STATUS[metric],
                    f"Target {targets.at[metric, 'TargetValue']:.2f}% "
                    f"({targets.at[metric, 'Direction']})")
    st.subheader("The cohort underneath")
    c1, c2, c3 = st.columns(3)
    metric_card(c1, f"{COHORT_VENDOR} on {COHORT_CATEGORY}", f"{cohort:.1f}%", "Grey",
                f"All history. As a vendor it reads "
                f"{vend.at[COHORT_VENDOR, 'AcceptanceRatePct']:.1f}% - unremarkable")
    metric_card(c2, f"{DELIVERY_VENDOR} on time", f"{vend.at[DELIVERY_VENDOR, 'OnTimePct']:.2f}%",
                "Grey", f"All history. By order year: {VEN_009_BY_YEAR}")
    metric_card(c3, f"Expedite spend at {first['AsOfDate']}", f"{first['ExpediteSpendPct']:.2f}%",
                "Grey", "Trailing 12 months. It has doubled since, across the whole book")
    st.caption(f"{DELIVERY_VENDOR}'s rate by order year is quoted from the case study, section 4.4: "
               "no extract carries on-time delivery by vendor and year.")

    st.subheader("Expedite spend, trailing 12 months, at each month-end")
    st.line_chart(kpi.set_index("AsOfDate")["ExpediteSpendPct"], height=240)
    st.caption("A rise across the book, not in one vendor or one set of parts. It stays green "
               "because the target is 1% of spend.")

    st.subheader("The lowest on-time rates, all history")
    st.dataframe(vendors.sort_values("OnTimePct")[
        ["VendorID", "VendorName", "VendorTier", "Lines", "OnTimePct", "AvgDaysLate"]].head(5),
        width="stretch", hide_index=True)

    decision(
        ("Recommendation 7: report every aggregate with its cohort cut", "Analytics",
         f"Acceptance is {now['AcceptanceRatePct']:.2f}% company-wide and {cohort:.1f}% for "
         f"{COHORT_VENDOR} on {COHORT_CATEGORY}. A green portfolio figure is not evidence that "
         "no cohort inside it has failed."),
        (f"Recommendation 6: put {DELIVERY_VENDOR} on a delivery corrective action",
         "Supplier quality",
         f"{DELIVERY_VENDOR} ({vend.at[DELIVERY_VENDOR, 'VendorName']}) was on time for "
         f"{VEN_009_BY_YEAR} ({vend.at[DELIVERY_VENDOR, 'OnTimePct']:.2f}% over all history): "
         f"the worst delivery record of the {len(vendors)} vendors."),
    )


# =============================================================================
# 5. THE QUEUE
# =============================================================================
elif page == SCREENS[4]:
    st.title("What sourcing works on this quarter")
    st.markdown(
        f"{len(queue)} vendor-part pairs is not a plan. A category manager can run perhaps eight "
        "serious negotiations a quarter, so the queue stays complete for audit and SQL flags "
        "what is workable (**IsThisQuarter**). **Ranking is by money; the action is by "
        "leverage** - both arrive from SQL. This page filters and totals; it re-ranks nothing."
    )

    tq = queue[queue["IsThisQuarter"] == 1]
    tq_value = float(tq["AnnualOpportunity"].sum())
    share = 100 * tq_value / gap_total
    c1, c2, c3, c4 = st.columns(4)
    metric_card(c1, "Pairs in the queue", f"{len(queue)}", "Grey", "Complete, for audit")
    metric_card(c2, "Workable this quarter", f"{len(tq)}", "Grey", "Flagged in SQL")
    metric_card(c3, "Annual value of this quarter's pairs", usd(tq_value), "Grey",
                f"of {usd(gap_total)} a year across all {len(queue)} pairs")
    metric_card(c4, "Share of the gap", f"{share:.1f}%", "Grey",
                f"carried by {len(tq)} of {len(queue)} pairs")

    st.subheader("The mix of work, by action")
    mix = (queue.groupby("ActionCode")
                .agg(Pairs=("PartNumber", "size"), Opportunity=("AnnualOpportunity", "sum"),
                     Spend12m=("TTMSpend", "sum"), ThisQuarter=("IsThisQuarter", "sum"))
                .reset_index().sort_values("Opportunity", ascending=False))
    st.dataframe(mix.round({"Opportunity": 0, "Spend12m": 0}), width="stretch", hide_index=True)
    st.caption("Opportunity is annual. These need different people. PUT_ON_CONTRACT needs a "
               "contract before a price; FIX_QUALITY needs a corrective action, because the loss "
               "is rejected material; ACCEPT is parked with a reason rather than listed forever.")

    st.subheader("Who works what")
    buyers = (queue.groupby(["BuyerID", "BuyerName", "Team"])
                   .agg(QueuePairs=("PartNumber", "size"), ThisQuarter=("IsThisQuarter", "sum"))
                   .reset_index())
    buyers["AnnualOppOnQuarterPairs"] = (
        buyers["BuyerID"].map(tq.groupby("BuyerID")["AnnualOpportunity"].sum()).fillna(0).round(2))
    st.dataframe(buyers, width="stretch", hide_index=True)
    st.caption("AnnualOppOnQuarterPairs is the annual opportunity on the pairs flagged this "
               "quarter - not an amount recovered this quarter. Capacity is eight pairs per buyer "
               "(vw_RenegotiationQueue passes 8). A buyer with fewer than eight flagged has ACCEPT "
               "pairs in their top eight by value: they take a place and are not flagged, because "
               "accepting is the decision for them.")

    idle = vendors[~vendors["VendorID"].isin(tq["VendorID"])]
    st.markdown(f"**{len(idle)} of {len(vendors)} vendors** have no pair this quarter.")
    with st.expander("Which vendors"):
        st.dataframe(idle[["VendorID", "VendorName", "VendorTier", "TotalSpend"]]
                     .sort_values("TotalSpend", ascending=False), width="stretch", hide_index=True)

    st.subheader("The queue")
    c1, c2, c3 = st.columns(3)
    pick_buyer = c1.multiselect("Buyer", sorted(queue["BuyerName"].unique()))
    pick_action = c2.multiselect("Action", sorted(queue["ActionCode"].unique()))
    only_tq = c3.checkbox("Only this quarter", value=True)
    view = queue
    if pick_buyer:
        view = view[view["BuyerName"].isin(pick_buyer)]
    if pick_action:
        view = view[view["ActionCode"].isin(pick_action)]
    if only_tq:
        view = view[view["IsThisQuarter"] == 1]
    st.caption(f"{len(view)} of {len(queue)} pairs, in SQL's PriorityRank order")
    st.dataframe(view[["PriorityRank", "BuyerRank", "BuyerName", "PartNumber", "Category", "VendorID",
                       "VendorName", "IsSingleSource", "LeverageScore", "ErosionCapturePct", "TTMSpend",
                       "AnnualOpportunity", "ActionCode", "RecommendedAction"]],
                 width="stretch", hide_index=True)
    st.caption("RecommendedAction is SQL's text, shown verbatim. For RENEGOTIATE it states each "
               "pair's own position: whether a second supplier is qualified, and whether Lumen is "
               "at least 8% of the supplier's revenue.")

    st.subheader("Where the spend with no agreement sits")
    st.dataframe(data["category"].rename(columns={"GroupValue": "Category"})
                 [["Category", "TotalSpend", "SpendSharePct", "MaverickPct", "SingleSourcePct"]]
                 .sort_values("MaverickPct", ascending=False), width="stretch", hide_index=True)
    st.caption("Trailing 12 months, duplicate PO lines excluded - so TotalSpend sums to less than "
               "the headline spend on the first screen, which counts them. Not rogue buying: "
               "contracts quietly expiring, and purchasing carrying on at drifting spot prices.")

    put = queue[queue["ActionCode"] == "PUT_ON_CONTRACT"]
    decision(
        (f"Recommendation 1: work the quarter's {len(tq)} pairs, starting with the "
         f"{int(put['IsThisQuarter'].sum())} PUT_ON_CONTRACT", "category managers",
         f"{len(tq)} pairs carry {usd_md(tq_value)} a year - {share:.1f}% of the "
         f"{usd_md(gap_total)} gap. This is what the sourcing team does on Monday."),
        ("Recommendation 2: re-paper the lapsed agreements before negotiating anything on them",
         "sourcing operations",
         f"{now['MaverickSpendPct']:.2f}% of spend has no agreement in force, against a "
         f"{targets.at['MaverickSpendPct', 'TargetValue']:.0f}% target. PUT_ON_CONTRACT "
         f"carries {usd_md(put['AnnualOpportunity'].sum())} across {len(put)} pairs that cannot "
         "be negotiated until there is something to negotiate against."),
    )


# =============================================================================
# 6. CAN THE DATA CARRY IT?
# =============================================================================
else:
    st.title("What is wrong with the data, and does the answer survive it?")
    dq = data["dq"]
    st.markdown(
        f"{len(dq)} behavioural checks run. **{int((dq['Anomalies'] == 0).sum())} of them found "
        "nothing and are reported anyway** - a summary listing only what was found cannot "
        "distinguish *we looked and it was clean* from *we never looked*."
    )
    st.dataframe(dq, width="stretch", hide_index=True)

    st.subheader("The gate")
    # usp_RunDataQualityChecks (sql/04): exposure over all landed cost, as the
    # procedure computes it; the README quotes the result.
    exposure = float(dq.loc[dq["CountsTowardExposure"] == 1, "AmountAtRisk"].sum())
    landed = float(data["po"]["LandedCost"].sum())
    misstated = round(100 * exposure / landed, 3)
    worst = dq.loc[dq["AmountAtRisk"].idxmax()]
    c1, c2, c3 = st.columns(3)
    metric_card(c1, "Spend mis-stated", f"{misstated:.3f}%", "Grey",
                f"{usd(exposure, 2)} of {usd(landed, 2)} landed cost")
    metric_card(c2, "Gate tolerance", f"{GATE_TOLERANCE_PCT:.3f}%", "Grey",
                f"A margin of {GATE_TOLERANCE_PCT - misstated:.3f} points")
    metric_card(c3, "Largest defect by value", usd(worst["AmountAtRisk"]), "Grey",
                f"{worst['AnomalyType']} - outside the gate on purpose")
    st.markdown(
        "The gate does not count defective rows. It asks how much **spend** is mis-stated (only "
        "checks flagged CountsTowardExposure add to it) and dbo.usp_RunDataQualityChecks raises "
        "above the tolerance. That margin is luck, not comfort.\n\nThe largest defect by value is "
        "outside on purpose: it does not mis-state how much was spent, it mis-states whether the "
        "PPV computed on that spend means anything. A ratio problem, not a total problem."
    )

    amb = dq[dq["AnomalyType"] == "AMBIGUOUS_CONTRACT"].iloc[0]
    # Each ambiguous line carries the agreement SQL resolved it to (fn_POLineCost:
    # latest start, then highest key); counting them is counting SQL's rows.
    po = data["po"]
    overlapping = po.loc[po["AgreementsInForce"] > 1, "AgreementNo"].nunique()
    decision((
        f"Recommendation 8: resolve the {overlapping} overlapping agreements", "sourcing operations",
        f"{int(amb['Anomalies'])} PO lines have two agreements in force at once, so the "
        f"contracted price - and therefore PPV - is ambiguous on {usd_md(amb['AmountAtRisk'])} "
        "of spend. Resolving them is what makes the variance figure mean something.",
    ))
