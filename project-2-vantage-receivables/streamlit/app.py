"""
================================================================================
Project 2 - Vantage Wholesale Supply: Receivables Performance
App:     streamlit/app.py
Purpose: The web front end to the receivables control. Five screens, each
         answering one question the case study answers, each ending in a
         decision and the person who makes it.

WHAT THIS IS FOR

    The finance director says collections have gone slack. The collections
    manager says Sales sold longer terms. This app settles it: of the days DSO
    has risen, how many Vantage granted itself, how many customers took, and
    how many were never customers at all -- and then who each collector should
    call first, and who should not be called.

    It is not a chart gallery. Every screen ends in a decision.

NO ANALYTICAL LOGIC LIVES HERE, AND THAT IS DELIBERATE

    Every definition -- countback DSO, Best Possible DSO, the bridge, the
    comparable window, collectable exposure, the priority rank, the action
    code, today's worklist, promise status, the data-quality exposure flag --
    is computed in SQL and arrives in an extract column. This app filters,
    sorts, counts and sums rows SQL already flagged, and draws. The only other
    arithmetic is a difference or ratio of two SQL figures whose result the
    case study quotes, taken from the figures AS PUBLISHED (case study 6.3).

    What the extracts do NOT carry is left out rather than re-derived: RAG
    status (usp_ARScorecard is not exported), the trailing twelve-month
    promise kept-rate, the terms shift-share and the regional billing-lag
    totals. A pandas copy would be a second definition, and the first time it
    disagreed with the workbook nobody could say which was wrong. A card that
    shows a scorecard metric at the reporting date is drawn in the status the
    case study publishes for it (its appendix); every other card colour is
    emphasis, not status, and the sidebar says so.

DATA
    The committed CSVs in erp_extracts/, exported from SQL Server (proved equal
    to SQL by erp_extracts/Validate_Extracts.ps1). Cloud has no SQL Server.

DATA DISCLOSURE
    Vantage Wholesale Supply is fictional and all data is synthetic. No
    confidential data is used and no claim is made about any production system.
================================================================================
"""

from __future__ import annotations

from pathlib import Path

import pandas as pd
import streamlit as st

EXTRACTS = Path(__file__).resolve().parent.parent / "erp_extracts"

st.set_page_config(
    page_title="Vantage Wholesale - Receivables Control",
    page_icon="=",
    layout="wide",
)

RED, AMBER, GREEN, GREY = "#c00000", "#bf8f00", "#548235", "#808080"
RAG = {"Red": RED, "Amber": AMBER, "Green": GREEN, "Grey": GREY}

SCREENS = [
    "Granted, taken, or ours",
    "Can the target be hit?",
    "Who to call first",
    "Cash already banked",
    "Green lights, failing cohorts",
]


@st.cache_data(show_spinner="Loading the committed extracts...")
def load_all() -> dict[str, pd.DataFrame]:
    names = [
        "ar_kpi_monthly", "dso_bridge_monthly", "ar_targets", "monthly_credit_sales",
        "priority_action_queue", "ar_open_items", "cash_application_worklist",
        "data_quality_summary", "promise_status", "billing_lag_monthly",
    ]
    # utf-8-sig: Export-Csv writes a byte-order mark before the first header.
    return {n: pd.read_csv(EXTRACTS / f"{n}.csv", encoding="utf-8-sig") for n in names}


def metric_card(column, label: str, value: str, rag: str, sub: str = "") -> None:
    colour = RAG.get(rag, GREY)
    column.markdown(
        f"""<div style="border-left:6px solid {colour};padding:0.4rem 0 0.4rem 0.9rem;margin-bottom:0.6rem">
        <div style="font-size:0.78rem;color:#555;text-transform:uppercase;letter-spacing:0.04em">{label}</div>
        <div style="font-size:2.0rem;font-weight:700;color:{colour};line-height:1.1">{value}</div>
        <div style="font-size:0.78rem;color:#666">{sub}</div></div>""",
        unsafe_allow_html=True,
    )


def decision(action: str, owner: str, why: str) -> None:
    st.subheader("The decision")
    st.success(f"**{action}**  \nOwner: {owner}.  \n{why}")


def usd(value: float, cents: bool = False) -> str:
    return f"${value:,.2f}" if cents else f"${value:,.0f}"


data = load_all()

# The comparable window is a SQL flag: the ledger opens in January 2024 and
# cannot show a representative ageing profile before late April.
kpi = data["ar_kpi_monthly"].sort_values("AsOfDate")
window = kpi[kpi["IsComparablePeriod"] == 1]
k0, k1 = window.iloc[0], window.iloc[-1]
bridge_all = data["dso_bridge_monthly"].sort_values("AsOfDate")
b1 = bridge_all[bridge_all["AsOfDate"] == k1["AsOfDate"]].iloc[0]
targets = data["ar_targets"].set_index("MetricName")
queue = data["priority_action_queue"].sort_values("PriorityRank")
today = queue[queue["IsTodaysWorklist"] == 1]
as_of = kpi["AsOfDate"].max()

# 2.38 days: unapplied cash over credit sales per day, both as published, and
# rounded before anything is derived from it (case study section 6.3).
unapplied_days = round(float(k1["UnappliedCash"]) / float(b1["SalesPerDay"]), 2)
customer_lateness = float(k1["AvgDaysDelinquent"]) - unapplied_days

st.sidebar.title("Vantage Wholesale Supply")
st.sidebar.caption("Receivables control")
st.sidebar.markdown(f"**Reporting date** {as_of}")
st.sidebar.caption("Source: committed CSV extracts (erp_extracts/), exported from SQL Server")
st.sidebar.caption(
    "Card borders: a scorecard metric at the reporting date carries the red/amber/green "
    "status the case study publishes; on every other card the colour is emphasis, not status."
)
st.sidebar.markdown("---")
st.sidebar.caption(
    "Vantage Wholesale Supply is fictional. All data is synthetic and generated by "
    "the scripts in `sql/`. No confidential data is used and no claim is made "
    "about any production system."
)

page = st.sidebar.radio("Screen", SCREENS)


# =============================================================================
# 1. GRANTED, TAKEN, OR OURS
# =============================================================================
if page == SCREENS[0]:
    st.title("Who granted the days, who took them, and who never had them?")
    st.markdown(
        f"DSO has climbed over the comparable window, {k0['YearMonth']} to {k1['YearMonth']}. "
        "The finance director reads that as collections going slack; the collections "
        "manager reads it as Sales selling longer terms. Both are plausible and both are "
        "self-serving. **Split the number.**"
    )

    c1, c2, c3 = st.columns(3)
    metric_card(c1, f"DSO (countback), {k0['YearMonth']}", f"{k0['DSO_Countback']:.2f} d", "Grey",
                "Start of the comparable window")
    metric_card(c2, f"DSO (countback), {k1['YearMonth']}", f"{k1['DSO_Countback']:.2f} d", "Red",
                f"Target {targets.loc['DSO', 'TargetValue']:.2f}")
    metric_card(c3, "The rise", f"{k1['DSO_Countback'] - k0['DSO_Countback']:+.2f} d", "Red",
                f"More days of sales sitting in receivables than in {k0['YearMonth']}")

    c1, c2, c3 = st.columns(3)
    metric_card(c1, "Granted - longer terms Vantage sold",
                f"{k1['BPDSO_Countback'] - k0['BPDSO_Countback']:+.2f} d", "Amber",
                f"Best Possible DSO {k0['BPDSO_Countback']:.2f} -> {k1['BPDSO_Countback']:.2f}")
    metric_card(c2, "Taken - customers paying later",
                f"{k1['AvgDaysDelinquent'] - k0['AvgDaysDelinquent']:+.2f} d", "Red",
                f"Average days delinquent {k0['AvgDaysDelinquent']:.2f} -> {k1['AvgDaysDelinquent']:.2f}")

    st.info(
        "**Except that a large part of 'taken' is not customers at all.** Cash that has "
        "reached Vantage's bank but has not been matched to an invoice does not reduce a "
        "balance, so every dollar of it makes an invoice look unpaid. It could not even be "
        "expressed until cash receipts and cash applications were modelled as separate events."
    )

    c1, c2, c3 = st.columns(3)
    metric_card(c1, f"Of the {k1['AvgDaysDelinquent']:.2f} days taken: Vantage's own cash",
                f"{unapplied_days:.2f} d", "Red",
                f"{usd(k1['UnappliedCash'], cents=True)} banked, matched to no invoice")
    metric_card(c2, "Genuine customer lateness", f"{customer_lateness:.2f} d", "Amber",
                "Customers paying late, disputed balances included")
    metric_card(c3, "Unapplied cash, % of open AR", f"{k1['UnappliedCashPct']:.2f}%", "Red",
                f"{k0['UnappliedCashPct']:.2f}% in {k0['YearMonth']}")

    st.subheader("Over the comparable window")
    trend = window.set_index("YearMonth")[["DSO_Countback", "BPDSO_Countback", "AvgDaysDelinquent"]]
    st.line_chart(trend.rename(columns={
        "DSO_Countback": "DSO (countback)",
        "BPDSO_Countback": "Granted: Best Possible DSO",
        "AvgDaysDelinquent": "Taken: average days delinquent",
    }), height=300)

    st.subheader(f"The bridge at {as_of}, on classic DSO")
    st.markdown(
        "Built on **classic** DSO (balance / sales per day) because classic partitions "
        "exactly and countback does not: no residual bar, nothing to plug."
    )
    st.dataframe(pd.DataFrame(
        [["Granted - not yet due", b1["GrantedDays"], b1["CurrentAR"], "Sales / terms policy"],
         ["Taken - disputed and past due", b1["DisputeDays"], b1["DisputedPastDueAR"], "Billing / sales"],
         ["Taken - past due, undisputed", b1["LatenessDays"], b1["UndisputedPastDueAR"], "Collections"],
         ["Total (classic DSO)", b1["DSO_Classic"], b1["TotalAR"], ""]],
        columns=["Component", "Days", "Cash", "Owner"]),
        width="stretch", hide_index=True)
    st.caption(
        f"Billing lag ({b1['BillingLagDays']:.2f} d) happens before the DSO clock starts, so "
        f"it extends the bridge rather than sitting in it: despatch to cash is "
        f"{b1['CashCycleDays']:.2f} days, larger than anything Vantage currently reports."
    )

    decision(
        "Apply the cash before making a single call.",
        "Cash application",
        f"{unapplied_days:.2f} days of DSO are Vantage's own money in Vantage's bank. "
        "Clearing them costs no customer conversation. Collections then works the "
        f"{customer_lateness:.2f} days of genuine lateness (screen 3), handing the disputed "
        "part to billing and sales, and the granted days go back to Sales as the commercial "
        "decision they were (screen 2).",
    )


# =============================================================================
# 2. CAN THE TARGET BE HIT?
# =============================================================================
elif page == SCREENS[1]:
    st.title("The 45-day target sits below the floor the terms allow")
    dso_t = targets.loc["DSO"]
    wat = float(b1["WeightedAvgTermsDays"])

    c1, c2, c3 = st.columns(3)
    metric_card(c1, "DSO target", f"{dso_t['TargetValue']:.2f} d", "Grey",
                f"Warning at {dso_t['WarningValue']:.2f}")
    metric_card(c2, "Weighted average terms sold", f"{wat:.2f} d", "Amber",
                f"Twelve months to {as_of}")
    metric_card(c3, "Collection allowance the target leaves",
                f"{dso_t['TargetValue'] - wat:+.2f} d", "Red",
                "Negative: missed even if every customer pays on time")
    st.markdown(
        f"If every customer paid **exactly on the due date**, DSO would settle near "
        f"{wat:.1f} and the {dso_t['TargetValue']:.0f}-day target would still be missed. "
        "A target nobody can reach does not motivate a collections team; it teaches them "
        "the scorecard is theatre."
    )
    # Both series are SQL columns: the month's own terms (monthly_credit_sales) and
    # the trailing twelve-month figure the floor rests on (dso_bridge_monthly).
    sold = data["monthly_credit_sales"].set_index("YearMonth")[["WeightedAvgTermsDays"]]
    sold = sold.rename(columns={"WeightedAvgTermsDays": "Terms sold in the month"})
    trailing = bridge_all.assign(YearMonth=bridge_all["AsOfDate"].str[:7]).set_index("YearMonth")
    sold["Trailing twelve months (the floor)"] = trailing["WeightedAvgTermsDays"]
    sold["DSO target"] = float(dso_t["TargetValue"])
    st.caption("Dollar-weighted terms sold, against the DSO target")
    st.line_chart(sold, height=260)
    st.caption(
        f"The floor is the trailing twelve-month figure ({wat:.2f} at {as_of}), not any single "
        "month: single months vary, and some sold shorter terms than the target. Before "
        "2024-12 the trailing figure covers only the months since the ledger opened. The "
        "rise between the second halves of 2024 and 2025 "
        "was a decision, not an accident: the shift-share decomposition in SQL "
        "(fn_WATShiftShare) attributes it to the same customers being moved onto longer terms, "
        "not to a change in who was buying. It is not in the extracts, so it is not restated here."
    )

    st.subheader("Two DSOs, on purpose")
    c1, c2, c3 = st.columns(3)
    metric_card(c1, "Countback DSO - the headline", f"{k1['DSO_Countback']:.2f} d", "Red",
                "Consumes the balance against the months that generated it")
    metric_card(c2, "Classic DSO - the bridge", f"{k1['DSO_Simple']:.2f} d", "Grey",
                "What a lender or a benchmark quotes")
    metric_card(c3, "Method gap", f"{k1['DSO_MethodGap']:.2f} d", "Grey",
                "Published, so nobody argues across methods")

    st.subheader("Collection effectiveness is flattered by write-offs")
    cei_t = targets.loc["CEI"]
    written_off = float(kpi["WriteOffs"].sum())
    first_writeoff = kpi[kpi["WriteOffs"] > 0]["YearMonth"].iloc[0]
    c1, c2, c3 = st.columns(3)
    metric_card(c1, "CEI (book)", f"{k1['CEI_Book']:.2f}%", "Red",
                f"Target {cei_t['TargetValue']:.2f}, warning {cei_t['WarningValue']:.2f}")
    metric_card(c2, "CEI (cash only)", f"{k1['CEI_Cash']:.2f}%", "Red",
                f"Paper collections gap {k1['PaperCollectionsGap']:.2f} pts - the honest one")
    metric_card(c3, "Written off", usd(written_off), "Red",
                f"All since {first_writeoff}; scored as collections by CEI (book)")
    st.markdown(
        "The textbook index measures whether AR went down - and AR goes down identically "
        "whether a customer wired the money or a controller wrote the balance off. That "
        "makes the standard metric improvable by giving up."
    )
    ends = window.iloc[[0, -1]][["YearMonth", "CEI_Book", "CEI_Cash", "PaperCollectionsGap", "WriteOffs"]]
    st.dataframe(ends, width="stretch", hide_index=True)
    st.line_chart(window.set_index("YearMonth")[["CEI_Book", "CEI_Cash"]], height=240)

    st.subheader("The scorecard, as the extracts carry it")
    board = [("DSO", k1["DSO_Countback"]), ("AvgDaysDelinquent", k1["AvgDaysDelinquent"]),
             ("CEI", k1["CEI_Book"]), ("PctPastDue", k1["PctPastDue"]),
             ("Pct90Plus", k1["Pct90Plus"]), ("UnappliedCashPct", k1["UnappliedCashPct"]),
             ("BillingLagDays", b1["BillingLagDays"])]
    card = targets.loc[[m for m, _ in board], ["TargetValue", "WarningValue", "Direction", "Description"]]
    card.insert(0, "Value", [v for _, v in board])
    st.dataframe(card.reset_index(), width="stretch", hide_index=True)
    st.caption(
        "Target and warning come from ar_targets; red/amber/green is decided in SQL "
        "(usp_ARScorecard), which is not exported, so it is not re-derived here. The card "
        "borders copy the status the case study publishes; they do not compute it. Promise "
        "kept rate, credit utilisation and dispute age are not in the extracts."
    )

    decision(
        "Reset the DSO target to terms plus a collection allowance (about 55 days in the "
        "case study), and measure collections on CEI (cash), not CEI (book).",
        "CFO",
        f"The target sits {wat - float(dso_t['TargetValue']):.2f} days below the terms floor, and "
        f"{usd(written_off)} of write-offs currently score as collections.",
    )


# =============================================================================
# 3. WHO TO CALL FIRST
# =============================================================================
elif page == SCREENS[2]:
    st.title("Who each collector calls first")
    items = data["ar_open_items"]
    open_ar = float(items["OpenBalance"].sum())
    ageing = (items.groupby(["BucketKey", "BucketName"])
                   .agg(Invoices=("InvoiceNo", "size"), Balance=("OpenBalance", "sum"))
                   .reset_index().sort_values("BucketKey"))
    ageing["SharePct"] = (100 * ageing["Balance"] / open_ar).round(1)
    ageing["Balance"] = ageing["Balance"].round(2)

    st.markdown(
        f"Today six collectors work a {len(items):,}-line ageing report from the top down "
        "by balance. A ranked list of every past-due account is still a ledger with a sort "
        "order."
    )
    c1, c2, c3, c4 = st.columns(4)
    metric_card(c1, "Open AR", usd(open_ar, cents=True), "Grey", f"{len(items):,} open invoices")
    metric_card(c2, "Past due share of AR", f"{k1['PctPastDue']:.2f}%", "Red",
                f"Target {targets.loc['PctPastDue', 'TargetValue']:.2f}")
    metric_card(c3, "90+ share of AR", f"{k1['Pct90Plus']:.2f}%", "Amber",
                f"Target {targets.loc['Pct90Plus', 'TargetValue']:.2f}")
    metric_card(c4, "Accounts with a past-due balance",
                f"{int((queue['PastDueBalance'] > 0).sum())}", "Red",
                f"{len(queue)} enter the queue; the rest on a disputed balance within terms")
    st.dataframe(ageing[["BucketName", "Invoices", "Balance", "SharePct"]],
                 width="stretch", hide_index=True)

    st.subheader("The bounded queue")
    st.markdown(
        "The queue is ranked in SQL on **collectable exposure** - ageing-weighted past-due "
        "dollars, less 75% of the ageing-weighted disputed dollars, less cash already banked "
        "against the account - and then bounded to what the team can work. This screen "
        "filters it; **it does not re-rank it**."
    )
    per_collector = today.groupby("CollectorName").size()
    share = 100 * today["CollectableExposure"].sum() / queue["CollectableExposure"].sum()
    c1, c2, c3 = st.columns(3)
    metric_card(c1, "On today's list", f"{len(today)}", "Amber",
                f"{per_collector.min()}-{per_collector.max()} per collector, "
                f"{len(per_collector)} collectors")
    metric_card(c2, "Share of collectable exposure", f"{share:.1f}%", "Green",
                f"{usd(today['CollectableExposure'].sum())} of "
                f"{usd(queue['CollectableExposure'].sum())}")
    apply_cash = queue[queue["ActionCode"] == "APPLY_CASH"]
    metric_card(c3, "APPLY_CASH accounts on today's list",
                f"{int(apply_cash['IsTodaysWorklist'].sum())}", "Green",
                f"of {len(apply_cash)} - routed to cash application instead")

    mix = (queue.groupby("ActionCode")
                .agg(Accounts=("CustomerID", "size"), PastDue=("PastDueBalance", "sum"),
                     CollectableExposure=("CollectableExposure", "sum"),
                     OnTodaysList=("IsTodaysWorklist", "sum"))
                .reset_index().sort_values("CollectableExposure", ascending=False))
    total = pd.DataFrame([["Total", mix["Accounts"].sum(), mix["PastDue"].sum(),
                           mix["CollectableExposure"].sum(), mix["OnTodaysList"].sum()]],
                         columns=mix.columns)
    mix = pd.concat([mix, total], ignore_index=True)
    mix[["PastDue", "CollectableExposure"]] = mix[["PastDue", "CollectableExposure"]].round(2)
    st.dataframe(mix, width="stretch", hide_index=True)
    st.caption(
        "APPLY_CASH takes precedence over every other action code (precedence, not rank: the "
        "banked cash offsets their exposure, so they sit low in PriorityRank). These accounts "
        "have cash already banked covering at least half their past-due balance. Ringing a "
        "customer who has already paid is the "
        "most damaging call a collections team can make. Disputes are discounted, not ranked; "
        "two broken promises in 90 days escalates regardless of balance."
    )

    st.subheader("A collector's worklist")
    c1, c2 = st.columns([2, 1])
    who = c1.selectbox("Collector", ["All collectors"] + sorted(queue["CollectorName"].unique()))
    whole = c2.checkbox("Show the whole queue, not only today's list", value=False)
    view = queue if whole else today
    if who != "All collectors":
        view = view[view["CollectorName"] == who].sort_values("CollectorRank")
    st.caption(f"{len(view)} of {len(queue)} accounts")
    st.dataframe(
        view[["PriorityRank", "CollectorRank", "CollectorName", "CustomerID", "CustomerName",
              "RiskTier", "ActionCode", "PastDueBalance", "Balance90Plus", "DisputedBalance",
              "UnappliedCash", "CollectableExposure", "OldestDaysPastDue", "BrokenPromises90d",
              "CreditHoldFlag", "RecommendedAction"]],
        width="stretch", hide_index=True,
    )

    escalations = int(today[today["ActionCode"] == "ESCALATE"].shape[0])
    decision(
        f"Work today's {len(today)} accounts, starting with the {escalations} escalations.",
        "Collections manager",
        f"It focuses {len(per_collector)} collectors on "
        f"{usd(today['CollectableExposure'].sum())} of the "
        f"{usd(queue['CollectableExposure'].sum())} at stake. The {len(apply_cash)} APPLY_CASH "
        f"accounts go to cash application: {int(apply_cash['IsTodaysWorklist'].sum())} of them "
        f"on today's list at {as_of} (UAT-20 checks this date), and in the whole queue each "
        "one reads 'Do not call'.",
    )


# =============================================================================
# 4. CASH ALREADY BANKED
# =============================================================================
elif page == SCREENS[3]:
    st.title("The cheapest money in the analysis is already in the bank")
    # SQL's order (usp_CashApplicationWorklist): oldest first, then largest.
    cash = data["cash_application_worklist"].sort_values(
        ["OldestUnappliedDays", "UnappliedCash"], ascending=False, kind="mergesort")

    routed = queue.loc[queue["ActionCode"] == "APPLY_CASH", "CustomerID"]
    st.markdown(
        "Collections gets a call list. Cash application gets its own: money customers have "
        "already paid, sitting unmatched against their accounts. **The cash must be applied "
        f"before any collections contact.** The {int(cash['CustomerID'].isin(routed).sum())} "
        "APPLY_CASH accounts among them, whose banked cash covers at least half their past-due "
        "balance, are not called at all (screen 3). The rest can still be on a call list for "
        "what their cash does not cover, so applying it first stops a collector chasing an "
        "invoice that is already paid. No negotiation, no concession, no relationship cost."
    )
    c1, c2, c3 = st.columns(3)
    metric_card(c1, "Accounts holding unmatched cash", f"{len(cash)}", "Red",
                f"{usd(cash['UnappliedCash'].sum(), cents=True)} in all")
    metric_card(c2, "Unmatched cash, in days of DSO", f"{unapplied_days:.2f} d", "Red",
                f"At {usd(b1['SalesPerDay'])} of credit sales a day")
    metric_card(c3, "Past due that would be wrongly chased", usd(cash["WronglyChaseable"].sum()),
                "Red", "If nobody applies the cash first")
    c1, c2, c3 = st.columns(3)
    metric_card(c1, "Receipts with no application at all",
                f"{int(cash['FullyUnappliedReceipts'].sum())}", "Amber",
                f"of {int(cash['UnappliedReceipts'].sum())} with any unapplied remainder")
    metric_card(c2, "Oldest receipt with an unapplied remainder",
                f"{int(cash['OldestUnappliedDays'].max())} days", "Amber", "Oldest first below")
    metric_card(c3, "Receipts with no remittance advice",
                f"{int(cash['NoRemittanceAdvice'].sum())}", "Grey", "Why they were never matched")
    st.dataframe(
        cash[["CustomerID", "CustomerName", "Region", "CollectorName", "UnappliedCash",
              "UnappliedReceipts", "FullyUnappliedReceipts", "NoRemittanceAdvice",
              "OldestUnappliedDays", "PastDueBalance", "WronglyChaseable"]],
        width="stretch", hide_index=True,
    )

    st.subheader("And some of the cash was banked twice")
    dq = data["data_quality_summary"]
    dup = dq.set_index("AnomalyType").loc["DUPLICATE_RECEIPT"]
    exposure = float(dq[dq["CountsTowardExposure"] == 1]["AmountAtRisk"].sum())
    clean = int((dq["Anomalies"] == 0).sum())
    c1, c2, c3 = st.columns(3)
    metric_card(c1, "Duplicate cash receipts", f"{int(dup['Anomalies'])}", "Red",
                f"{usd(dup['AmountAtRisk'], cents=True)} banked twice")
    metric_card(c2, "Open AR mis-stated by", f"{100 * exposure / float(k1['EndAR']):.3f}%", "Red",
                f"of {usd(k1['EndAR'], cents=True)}; the SQL gate fails above 1.000%")
    metric_card(c3, "Checks that found nothing", f"{clean} of {len(dq)}", "Grey",
                "Listed anyway: 'clean' and 'never looked' must differ")
    st.markdown(
        "The data-quality gate **fails**, and the build does not suppress it: a layer nobody "
        "has watched fire is an assertion, not a control. `DUPLICATE_RECEIPT` and "
        "`OVER_APPLIED_CASH` report the **same money** - one is the cause, the other its "
        "effect on the ledger - so only the effect counts toward exposure "
        "(`CountsTowardExposure`), or the same dollars would be charged twice."
    )
    st.dataframe(dq, width="stretch", hide_index=True)

    decision(
        "Apply the unmatched cash before a single call is made, and fix duplicate receipt posting.",
        "Cash application (the unmatched cash) and the Controller (the duplicates)",
        f"Together {usd(float(cash['UnappliedCash'].sum()) + float(dup['AmountAtRisk']))} of "
        "mis-stated or unusable receivables. Cash application is the broken function, and "
        "neither fix needs a customer conversation.",
    )


# =============================================================================
# 5. GREEN LIGHTS, FAILING COHORTS
# =============================================================================
else:
    st.title("Three averages that cannot breach until the whole book does")
    st.markdown(
        "Promise-keeping, billing lag and credit utilisation all read Green for the portfolio. "
        "Each is arithmetically correct and operationally useless alone: by the time a "
        "portfolio average breaches, the cohort that caused it has been failing for a year."
    )

    st.subheader("Promises to pay, by risk tier")
    promises = data["promise_status"]
    tiers = (promises.pivot_table(index="RiskTier", columns="PromiseStatus", values="PromiseNo",
                                  aggfunc="count", fill_value=0).reset_index())
    broken = promises[promises["PromiseStatus"] == "Broken"]
    broken_by_tier = broken.groupby("RiskTier")["PromisedAmount"].sum().round(2)
    tiers["BrokenDollars"] = tiers["RiskTier"].map(broken_by_tier).fillna(0.0)
    tiers = tiers.sort_values("BrokenDollars", ascending=False)
    worst = tiers.iloc[0]
    c1, _ = st.columns(2)
    metric_card(c1, f"Broken promises, {worst['RiskTier']} risk tier", f"{int(worst['Broken'])}",
                "Red", f"{usd(worst['BrokenDollars'])} promised and not paid; "
                f"{len(broken)} broken in the whole register")
    st.dataframe(tiers, width="stretch", hide_index=True,
                 column_config={"BrokenDollars": st.column_config.NumberColumn(format="dollar")})
    st.caption(
        f"All {len(promises):,} promises in the register, promised pay dates "
        f"{promises['PromisedPayDate'].min()} to {promises['PromisedPayDate'].max()}. Status "
        "is computed from cash in SQL (fn_PromiseStatus), never from the register the "
        "collectors maintain. The kept-rate KPI is a trailing twelve-month figure no extract "
        "carries, so it is not recomputed here."
    )

    st.subheader("Billing lag, by region")
    lag = data["billing_lag_monthly"]
    year = as_of[:4]
    southeast = lag[(lag["Region"] == "Southeast") & lag["YearMonth"].str.startswith(year)]
    c1, c2, c3 = st.columns(3)
    metric_card(c1, "Billing lag, portfolio", f"{b1['BillingLagDays']:.2f} d", "Green",
                f"Target {targets.loc['BillingLagDays', 'TargetValue']:.2f}; twelve months, dollar-weighted")
    metric_card(c2, f"Southeast, {southeast.iloc[0]['YearMonth']}",
                f"{southeast.iloc[0]['BillingLagDays']:.2f} d", "Grey", "Despatch to invoice")
    metric_card(c3, f"Southeast, {southeast.iloc[-1]['YearMonth']}",
                f"{southeast.iloc[-1]['BillingLagDays']:.2f} d", "Red", "Drifting all year")
    # a date index gives the chart a time axis; twelve-plus text labels get truncated
    by_region = lag.pivot(index="YearMonth", columns="Region", values="BillingLagDays")
    by_region.index = pd.to_datetime(by_region.index + "-01")
    st.line_chart(by_region, height=260)
    st.caption(
        "Four regions sit flat; one site drifts. Billing lag is, with cash application "
        "(screen 4), one of the two parts of the cash cycle Vantage can shorten without a "
        "customer conversation."
    )

    st.subheader("Credit utilisation, the queue accounts nearest their limit")
    hold = queue[queue["CreditHoldFlag"] == 1]
    util_t = targets.loc["CreditUtilization"]
    c1, c2 = st.columns(2)
    metric_card(c1, "Accounts over their credit limit", f"{len(hold)}", "Red",
                ", ".join(f"{r.CustomerID} at {r.CreditUtilizationPct:.2f}%" for r in hold.itertuples()))
    metric_card(c2, "Credit utilisation target", f"{util_t['TargetValue']:.2f}%", "Grey",
                f"Warning (credit hold) above {util_t['WarningValue']:.2f}%")
    st.dataframe(
        queue.sort_values("CreditUtilizationPct", ascending=False).head(5)[
            ["CustomerID", "CustomerName", "RiskTier", "CreditUtilizationPct", "CreditHoldFlag",
             "ActionCode", "PastDueBalance"]],
        width="stretch", hide_index=True,
    )
    st.caption(
        "Credit hold is a flag beside the collections action, not a replacement for it: as a "
        "competing branch it was unreachable, because every over-limit account already "
        "qualified for escalation."
    )

    decision(
        "Report promise-kept by risk tier, never as a portfolio figure; investigate Southeast "
        "invoice posting.",
        "Collections manager (promises) and Operations (Southeast billing)",
        f"{int(worst['Broken'])} of {len(broken)} broken promises sit in one tier underneath a "
        f"portfolio figure that clears its target, and Southeast billing lag went from "
        f"{southeast.iloc[0]['BillingLagDays']:.2f} to {southeast.iloc[-1]['BillingLagDays']:.2f} "
        "days while four regions held steady.",
    )
