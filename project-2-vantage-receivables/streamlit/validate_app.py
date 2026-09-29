"""
================================================================================
Project 2 - Vantage Wholesale Supply: Receivables Performance
Script:  streamlit/validate_app.py
Purpose: Run every screen of the app headlessly, assert none of them raises,
         and assert the figures it shows are the figures the case study prints.

WHY "IT STARTS" IS NOT A TEST

    `streamlit run` returns HTTP 200 as soon as the server is up. The app script
    does not execute until a browser connects, and then it executes ONE screen
    -- whichever the sidebar defaults to. A KeyError on the fourth screen is
    invisible to any check that looks at the server.

    AppTest runs the script in-process, once per screen, and surfaces the
    exception if there is one. The checks below then read what each screen
    actually rendered: its explanatory text, its tables, its decision, and the
    value on every headline card.

WHY THE EXPECTED FIGURES ARE HARD-CODED STRINGS

    A check that recomputes the expected value from the same CSV the app reads
    cannot fail: both sides move together. Each expected string below is typed
    in as the case study prints it, with the line it comes from, so a changed
    extract, a changed format or a changed rounding rule all FAIL here.
    CS = case_study/Vantage_Receivables_Case_Study.md, RM = README.md.

FIGURES THE APP BUILD CORRECTED (2026-09-29)

    Building this app found figures where the case study and SQL disagreed.
    Each was settled in SQL, then fixed at the source: see
    docs/data_validation_report.md section 3f.
    - 10-13 per collector on today's list (CS line 238, RM line 134);
    - broken promises: 255 in the register, 239 of them in the High risk
      tier worth $607,759 (CS lines 180 and 276, RM line 92);
    - the paper collections gap, now derived from the CEIs as published:
      0.75 at 2024-06 (CS line 193) and 1.65 at 2025-12.
    Accounts with a past-due balance: 262, as RM line 130 prints (the other
    10 of the 272 queued accounts carry only disputed balances).

WHY THE WORDING IS CHECKED TOO

    Some sentences the app once showed were false against its own data
    ("these accounts must not be called" over 97 accounts, 15 of them on
    today's call list). RETIRED lists them per screen; each must not return.
    The default worklist order is checked with hard-coded rows, because a
    pandas re-sort of today's list is exactly the "ledger with a sort order"
    the queue exists to replace, and nothing else on the screen would change.

RUN IT WITH
    python project-2-vantage-receivables/streamlit/validate_app.py
    (from the repository root or from any other directory)

DATA DISCLOSURE
    Vantage Wholesale Supply is fictional and all data is synthetic.
================================================================================
"""

from __future__ import annotations

import re
import sys
from pathlib import Path

import pyarrow as pa  # ships with streamlit; reads the data behind st.line_chart
from streamlit.testing.v1 import AppTest

HERE = Path(__file__).resolve().parent
APP = HERE / "app.py"

passed = 0
failed = 0


def check(name: str, ok: bool, detail: str = "") -> None:
    global passed, failed
    if ok:
        passed += 1
        print(f"PASS  {name}")
    else:
        failed += 1
        print(f"FAIL  {name}   {detail}")


_CARD = re.compile(
    r'<div style="border-left:6px solid (#[0-9a-fA-F]{6})[^"]*">\s*'
    r'<div style="font-size:0\.78rem;color:#555[^"]*">(.*?)</div>\s*'
    r'<div style="font-size:2\.0rem[^"]*">(.*?)</div>\s*'
    r'<div style="font-size:0\.78rem;color:#666">(.*?)</div>',
    re.DOTALL,
)

# The P5 palette the app copies. Typed in here, not imported, so a changed
# colour in the app fails the status checks below.
RED, AMBER, GREEN, GREY = "#c00000", "#bf8f00", "#548235", "#808080"
NAME = {RED: "Red", AMBER: "Amber", GREEN: "Green", GREY: "Grey"}

# Sentences the app once showed that were false against its own data. Each
# must stay gone (r5-review-app-p2, 2026-09-29).
RETIRED = {
    "Granted, taken, or ours": ["What a collector can actually move",
                                "Days of sales now sitting in receivables"],
    "Can the target be hit?": ["The rise through 2025"],
    "Who to call first": ["APPLY_CASH outranks everything", "never to a collector"],
    "Cash already banked": ["must not be called"],
    "Green lights, failing cohorts": ["the only part of the cash cycle"],
}


def cards(at: AppTest) -> dict[str, tuple[str, str, str]]:
    """Every metric card on the screen: label -> (value, sub-line, border colour)."""
    found = {}
    for m in at.markdown:
        for colour, label, value, sub in _CARD.findall(m.value):
            found[label.strip()] = (value.strip(), sub.strip(), colour.lower())
    return found


def check_status(at: AppTest, expected: list[tuple[str, str, str]]) -> None:
    """(card label, colour, source). A card showing a scorecard metric at the
    reporting date is drawn in the status the case study publishes for it."""
    shown = cards(at)
    for label, colour, source in expected:
        got = shown.get(label)
        check(f"  card '{label}' drawn {NAME[colour]}   [{source}]",
              got is not None and got[2] == colour,
              f"drawn {NAME.get(got[2], got[2]) if got else 'no such card'}")


def check_retired(at: AppTest, screen: str) -> None:
    body = text_of(at)  # card sub-lines are markdown, so they are in here too
    for phrase in RETIRED[screen]:
        check(f"  retired claim stays gone: '{phrase}'", phrase.lower() not in body.lower(),
              "the false sentence is back on the screen")


def chart_series(at: AppTest) -> dict[str, dict[str, float]]:
    """Every st.line_chart on the screen, as series name -> {x: y}."""
    series: dict[str, dict[str, float]] = {}
    for chart in at.get("vega_lite_chart"):
        for ds in chart.proto.datasets:
            frame = pa.ipc.open_stream(ds.data.data).read_all().to_pandas()
            x, name, y = frame.columns[:3]
            for _, row in frame.iterrows():
                series.setdefault(str(row[name]), {})[str(row[x])] = row[y]
    return series


def text_of(at: AppTest) -> str:
    return " ".join(
        [m.value for m in at.markdown] + [c.value for c in at.caption]
        + [t.value for t in at.title] + [s.value for s in at.subheader]
        + [i.value for i in at.info] + [s.value for s in at.success]
    )


def table_with(at: AppTest, column: str):
    """The first table on the screen that has this column, or None."""
    for d in at.dataframe:
        if column in d.value.columns:
            return d.value
    return None


def open_screen(label: str) -> AppTest | None:
    at = AppTest.from_file(str(APP), default_timeout=300)
    at.run()
    if not at.exception:
        try:
            at.sidebar.radio[0].set_value(label).run()
        except Exception as exc:  # noqa: BLE001
            check(f"screen renders: {label}", False, f"could not select it: {exc}")
            return None
    if at.exception:
        check(f"screen renders: {label}", False, str(at.exception[0].value)[:180])
        return None
    check(f"screen renders: {label}", True)
    return at


def check_screen(at: AppTest, must_contain: str, min_tables: int) -> None:
    """A screen that renders nothing renders without error. Assert it produced
    the thing it exists for, and that it ends in a decision with an owner."""
    body = text_of(at)
    check(f"  ...says something about '{must_contain}'", must_contain.lower() in body.lower(),
          "the screen rendered but its explanatory text is missing")
    check(f"  ...shows at least {min_tables} table(s)", len(at.dataframe) >= min_tables,
          f"found {len(at.dataframe)}")
    check("  ...ends in a decision with an owner",
          len(at.success) == 1 and "Owner:" in at.success[0].value,
          f"found {len(at.success)} decision block(s)")


def check_cards(at: AppTest, expected: list[tuple[str, str, str, str]]) -> None:
    """(card label, value exactly as printed, text the sub-line must contain, source)."""
    shown = cards(at)
    for label, value, sub, source in expected:
        got = shown.get(label)
        ok = got is not None and got[0] == value and sub in got[1]
        check(f"  card '{label}' = {value}" + (f" ({sub})" if sub else "") + f"   [{source}]",
              ok, f"rendered {got!r}" if got else f"no such card; cards are {sorted(shown)}")


def check_table(frame, name: str, key: str, expected: dict, fmt, source: str) -> None:
    """Each expected row {key value: {column: printed string}} must match the table.
    fmt prints a cell the way the case study does: one function, or one per column."""
    if frame is None:
        check(f"  table '{name}'", False, "not on the screen")
        return
    rows = frame.set_index(key)
    for k, columns in expected.items():
        for column, want in columns.items():
            show = fmt[column] if isinstance(fmt, dict) else fmt
            got = show(rows.at[k, column]) if k in rows.index and column in rows.columns else None
            check(f"  table '{name}': {k} {column} = {want}   [{source}]", got == want,
                  f"rendered {got!r}")


def two(v) -> str:
    return f"{float(v):.2f}"


def one(v) -> str:
    return f"{float(v):.1f}"


def whole(v) -> str:
    return f"{int(v)}"


def dollars(v) -> str:
    return f"${float(v):,.0f}"


print()
print("Vantage Wholesale - running every screen and checking its figures")
print("-" * 78)

# Rule 8 of the build spec: the page text and the code are ASCII only.
for f in (APP, HERE / "validate_app.py", HERE / "requirements.txt"):
    raw = f.read_bytes()
    check(f"ASCII only: {f.name}", all(b < 128 for b in raw),
          f"first non-ASCII byte at offset {next((i for i, b in enumerate(raw) if b > 127), -1)}")

# ---------------------------------------------------------------- sidebar
at = AppTest.from_file(str(APP), default_timeout=300)
at.run()
side = " ".join([m.value for m in at.sidebar.markdown] + [c.value for c in at.sidebar.caption])
check("sidebar: reporting date 2025-12-31   [CS line 4]", "2025-12-31" in side, side[:120])
check("sidebar: source is the committed CSV extracts",
      "Source: committed CSV extracts (erp_extracts/), exported from SQL Server" in side, side[:200])
check("sidebar: fictional / synthetic data disclosure",
      "fictional" in side and "synthetic" in side, side[:200])
print("-" * 78)

# ---------------------------------------------------------------- 1
at = open_screen("Granted, taken, or ours")
if at:
    check_screen(at, "nothing to plug", 1)
    check_cards(at, [
        ("DSO (countback), 2024-06", "55.85 d", "", "CS line 107; RM line 18"),
        ("DSO (countback), 2025-12", "63.07 d", "Target 45.00", "CS lines 107, 402"),
        ("The rise", "+7.22 d", "", "CS line 107; RM line 24"),
        ("Granted - longer terms Vantage sold", "+1.94 d", "44.33 -> 46.27", "CS line 108; RM line 22"),
        ("Taken - customers paying later", "+5.28 d", "11.52 -> 16.80", "CS line 109; RM line 23"),
        ("Of the 16.80 days taken: Vantage's own cash", "2.38 d", "$419,392.13",
         "CS lines 119, 121, 126"),
        ("Genuine customer lateness", "14.42 d", "", "CS line 126; RM line 32"),
        ("Unapplied cash, % of open AR", "3.87%", "1.79% in 2024-06", "CS line 118"),
    ])
    check_table(table_with(at, "Component"), "DSO bridge", "Component", {
        "Granted - not yet due":         {"Days": "45.97", "Cash": "8095793.78"},
        "Taken - disputed and past due": {"Days": "0.88", "Cash": "154864.71"},
        "Taken - past due, undisputed":  {"Days": "14.74", "Cash": "2596423.00"},
        "Total (classic DSO)":           {"Days": "61.59", "Cash": "10847081.49"},
    }, two, "CS lines 70-73")
    body = text_of(at)
    check("  cash cycle: billing lag 1.78 d, despatch to cash 63.37 days   [CS line 82]",
          "(1.78 d)" in body and "63.37 days" in body, "caption missing or changed")
    check_status(at, [
        ("DSO (countback), 2025-12", RED, "CS line 402"),
        ("Unapplied cash, % of open AR", RED, "CS line 408"),
    ])
    check_retired(at, "Granted, taken, or ours")

# ---------------------------------------------------------------- 2
at = open_screen("Can the target be hit?")
if at:
    check_screen(at, "floor the terms allow", 2)
    check_cards(at, [
        ("DSO target", "45.00 d", "Warning at 55.00", "CS lines 154, 402"),
        ("Weighted average terms sold", "46.68 d", "", "CS line 155; RM line 37"),
        ("Collection allowance the target leaves", "-1.68 d", "", "CS line 156"),
        ("Countback DSO - the headline", "63.07 d", "", "CS line 56"),
        ("Classic DSO - the bridge", "61.59 d", "", "CS line 57"),
        ("CEI (book)", "66.91%", "Target 85.00, warning 75.00", "CS lines 191, 404"),
        ("CEI (cash only)", "65.26%", "Paper collections gap 1.65 pts", "CS lines 192-193"),
        ("Written off", "$702,389", "All since 2025-04", "CS line 195"),
    ])
    check("  terms floor: 'settle near 46.7'   [CS line 158; RM line 37]",
          "settle near 46.7 " in text_of(at), "text missing or changed")
    check_table(table_with(at, "CEI_Book"), "CEI at both ends of the window", "YearMonth", {
        "2024-06": {"CEI_Book": "71.80", "CEI_Cash": "71.05", "PaperCollectionsGap": "0.75"},
        "2025-12": {"CEI_Book": "66.91", "CEI_Cash": "65.26", "PaperCollectionsGap": "1.65"},
    }, two, "CS lines 191-192")
    card = {
        "DSO":               {"Value": "63.07", "TargetValue": "45.00", "WarningValue": "55.00"},
        "AvgDaysDelinquent": {"Value": "16.80", "TargetValue": "8.00", "WarningValue": "15.00"},
        "CEI":               {"Value": "66.91", "TargetValue": "85.00", "WarningValue": "75.00"},
        "PctPastDue":        {"Value": "25.36", "TargetValue": "15.00", "WarningValue": "25.00"},
        "Pct90Plus":         {"Value": "3.35", "TargetValue": "3.00", "WarningValue": "6.00"},
        "UnappliedCashPct":  {"Value": "3.87", "TargetValue": "0.50", "WarningValue": "1.50"},
        "BillingLagDays":    {"Value": "1.78", "TargetValue": "2.00", "WarningValue": "4.00"},
    }
    check_table(table_with(at, "TargetValue"), "scorecard", "MetricName", card, two,
                "CS lines 402-409")
    decision = at.success[0].value if at.success else ""
    check("  decision: reset the target to about 55 days   [CS line 274]",
          "about 55 days" in decision and "1.68 days" in decision, decision[:120])
    check_status(at, [
        ("Countback DSO - the headline", RED, "CS line 402"),
        ("CEI (book)", RED, "CS line 404"),
    ])
    # The title says the target sits below the floor. The floor is the trailing
    # twelve-month figure, so the chart must plot it, end on 46.68 and never dip
    # to the target; single months may (2025-05 sold 43.87).
    floor = chart_series(at).get("Trailing twelve months (the floor)", {})
    points = [v for v in floor.values() if v == v]  # drop the NaN before a trailing value exists
    check("  terms chart: the trailing twelve-month floor ends on 46.68   [CS line 155]",
          f"{float(floor.get('2025-12', 'nan')):.2f}" == "46.68", f"series is {sorted(floor)[-3:]}")
    check("  terms chart: the floor never falls to the 45-day target   [CS lines 158-159]",
          len(points) >= 12 and min(points) > 45.0, f"{len(points)} points, min {min(points, default=None)}")
    check("  terms caption: the rise is H2-2024 vs H2-2025   [CS line 136]",
          "between the second halves of 2024 and 2025" in text_of(at), "caption missing or changed")
    check_retired(at, "Can the target be hit?")

# ---------------------------------------------------------------- 3
at = open_screen("Who to call first")
if at:
    check_screen(at, "does not re-rank", 3)
    check_cards(at, [
        ("Open AR", "$10,847,081.49", "2,346 open invoices", "CS line 424; RM lines 230-231"),
        ("Past due share of AR", "25.36%", "Target 15.00", "CS line 406"),
        ("90+ share of AR", "3.35%", "Target 3.00", "CS line 407"),
        ("Accounts with a past-due balance", "262", "272 enter the queue", "RM line 130"),
        ("On today's list", "65", "10-13 per collector, 6 collectors",
         "CS line 275; RM line 134 (CS line 238 says 10 to 12: doc finding F3)"),
        ("Share of collectable exposure", "81.3%", "$640,540 of $788,318", "CS lines 238, 275"),
        ("APPLY_CASH accounts on today's list", "0", "of 51", "CS lines 235, 246 (UAT-20)"),
    ])
    check_status(at, [
        ("Past due share of AR", RED, "CS line 406"),
        ("90+ share of AR", AMBER, "CS line 407"),
    ])
    check_retired(at, "Who to call first")
    check_table(table_with(at, "BucketName"), "ageing", "BucketName", {
        "Current": {"Invoices": "1672", "Balance": "8095793.78", "SharePct": "74.6"},
        "1-30":    {"Invoices": "414", "Balance": "1839519.41", "SharePct": "17.0"},
        "31-60":   {"Invoices": "101", "Balance": "405938.13", "SharePct": "3.7"},
        "61-90":   {"Invoices": "34", "Balance": "142966.39", "SharePct": "1.3"},
        "90+":     {"Invoices": "125", "Balance": "362863.78", "SharePct": "3.3"},
    }, {"Invoices": whole, "Balance": two, "SharePct": one}, "CS lines 419-423")
    actions = {
        "ESCALATE":          (19, "$758,532", "$359,632", 19),
        "COLLECTION_CALL":   (26, "$476,443", "$155,712", 15),
        "STANDARD_DUNNING":  (100, "$634,679", "$144,147", 21),
        "FINAL_NOTICE":      (19, "$356,211", "$80,901", 5),
        "RESOLVE_DISPUTE":   (45, "$195,694", "$36,407", 5),
        "COURTESY_REMINDER": (12, "$26,036", "$6,509", 0),
        "APPLY_CASH":        (51, "$303,692", "$5,011", 0),
        "Total":             (272, "$2,751,288", "$788,318", 65),
    }
    mix = table_with(at, "OnTodaysList")
    check_table(mix, "actions (counts)", "ActionCode",
                {k: {"Accounts": str(v[0]), "OnTodaysList": str(v[3])} for k, v in actions.items()},
                whole, "CS lines 229-236; RM lines 138-144")
    check_table(mix, "actions (dollars)", "ActionCode",
                {k: {"PastDue": v[1], "CollectableExposure": v[2]} for k, v in actions.items()},
                dollars, "CS lines 229-236; RM lines 138-144")
    decision = at.success[0].value if at.success else ""
    check("  decision: start with the 19 escalations; $640,540 of the $788,318   [CS line 275]",
          "19 escalations" in decision and "$640,540 of the $788,318" in decision, decision[:120])
    check("  decision: APPLY_CASH claim is scoped to the date UAT-20 checks   [CS line 246; F-NEW-A]",
          "0 of them on today's list at 2025-12-31 (UAT-20 checks this date)" in decision,
          decision[-200:])

    # The controls RE-CUT the SQL ranking; they never re-rank it.
    work = table_with(at, "CollectorRank")
    check("  worklist: today's list is 65 accounts, none of them APPLY_CASH   [CS lines 238, 246]",
          work is not None and len(work) == 65 and not (work["ActionCode"] == "APPLY_CASH").any(),
          f"{None if work is None else len(work)} rows")
    # Hard-coded, not recomputed: SQL's first three at 2025-12-31, as
    # priority_action_queue.csv carries them. A pandas re-sort of today's list
    # (by balance, by age, by anything) moves them.
    head = None if work is None else list(zip(work["PriorityRank"].head(3), work["CustomerID"].head(3)))
    check("  worklist: today's list opens C0351, C0366, C0302 at PriorityRank 1-3   [extract, 2025-12-31]",
          head == [(1, "C0351"), (2, "C0366"), (3, "C0302")], f"opens {head}")
    check("  worklist: today's list is in SQL's PriorityRank order",
          work is not None and list(work["PriorityRank"]) == sorted(work["PriorityRank"]),
          "today's list was re-sorted")
    at.checkbox[0].check().run()
    work = table_with(at, "CollectorRank")
    check("  worklist: whole queue is 272 accounts in SQL's PriorityRank order   [CS line 236]",
          work is not None and len(work) == 272
          and list(work["PriorityRank"]) == sorted(work["PriorityRank"]),
          f"{None if work is None else len(work)} rows")
    at.checkbox[0].uncheck().run()
    at.selectbox[0].set_value("Priya Natarajan").run()
    work = table_with(at, "CollectorRank")
    check("  worklist: one collector's list is theirs alone, in SQL's CollectorRank order",
          not at.exception and work is not None and len(work) > 0
          and set(work["CollectorName"]) == {"Priya Natarajan"}
          and list(work["CollectorRank"]) == sorted(work["CollectorRank"]),
          "filter or order wrong")

# ---------------------------------------------------------------- 4
at = open_screen("Cash already banked")
if at:
    check_screen(at, "before any collections contact", 2)
    check("  only the 51 APPLY_CASH accounts are 'not called at all'   [CS lines 235, 242]",
          "The 51 APPLY_CASH accounts among them" in text_of(at)
          and "are not called at all" in text_of(at), "text missing or changed")
    check_retired(at, "Cash already banked")
    check_cards(at, [
        ("Accounts holding unmatched cash", "97", "$419,392.13", "CS lines 119, 259"),
        ("Unmatched cash, in days of DSO", "2.38 d", "$176,118", "CS line 121"),
        ("Past due that would be wrongly chased", "$409,640", "", "CS lines 260, 272"),
        ("Receipts with no application at all", "57", "", "CS lines 259, 272"),
        ("Oldest receipt with an unapplied remainder", "679 days", "", "CS lines 260, 272"),
        ("Duplicate cash receipts", "96", "$371,039.99", "CS line 204; RM line 158"),
        ("Open AR mis-stated by", "3.421%", "1.000%", "CS line 211; RM line 156"),
        ("Checks that found nothing", "6 of 12", "", "RM line 50; CS lines 204-209"),
    ])
    check_table(table_with(at, "AnomalyType"), "data quality", "AnomalyType", {
        "DUPLICATE_RECEIPT":       {"Anomalies": "96", "AmountAtRisk": "371039.99"},
        "OVER_APPLIED_CASH":       {"Anomalies": "101", "AmountAtRisk": "371039.99"},
        "UNAPPLIED_CASH_AGED":     {"Anomalies": "65", "AmountAtRisk": "265487.11"},
        "RECEIPT_BEFORE_INVOICE":  {"Anomalies": "27", "AmountAtRisk": "100496.95"},
        "PRE_BILLING":             {"Anomalies": "20", "AmountAtRisk": "76841.74"},
        "DUE_DATE_TERMS_MISMATCH": {"Anomalies": "52", "AmountAtRisk": "12727.40"},
    }, {"Anomalies": whole, "AmountAtRisk": two}, "CS lines 204-209")
    decision = at.success[0].value if at.success else ""
    check("  decision: together $790,432 of mis-stated or unusable receivables   [CS line 282]",
          "$790,432" in decision, decision[:160])

# ---------------------------------------------------------------- 5
at = open_screen("Green lights, failing cohorts")
if at:
    check_screen(at, "cohort", 2)
    check_cards(at, [
        ("Billing lag, portfolio", "1.78 d", "Target 2.00", "CS lines 171, 409; RM line 93"),
        ("Southeast, 2025-01", "1.56 d", "", "CS line 171; RM line 93"),
        ("Southeast, 2025-12", "4.78 d", "", "CS line 171; RM line 93"),
        ("Accounts over their credit limit", "1", "C0297 at 106.10%", "CS lines 172, 415"),
        ("Credit utilisation target", "80.00%", "100.00%", "CS line 411"),
        # SQL-verified, not as the documents print it (see the docstring):
        # CS lines 180, 276 and RM line 92 still say 244 / $625,642.
        ("Broken promises, High risk tier", "239", "$607,759 promised and not paid; "
         "255 broken in the whole register", "SQL fn_PromiseKeptRate 2025-12-31, finding F4"),
    ])
    check_status(at, [("Billing lag, portfolio", GREEN, "CS line 409")])
    # 1.56 d is under the 2.00 target: red or amber would claim a breach.
    se = cards(at).get("Southeast, 2025-01")
    check("  card 'Southeast, 2025-01' not drawn red or amber   [CS line 409 target 2.00]",
          se is not None and se[2] not in (RED, AMBER),
          f"drawn {NAME.get(se[2], se[2]) if se else 'no such card'}")
    check_retired(at, "Green lights, failing cohorts")
    credit = table_with(at, "CreditUtilizationPct")
    above = None if credit is None else int((credit["CreditUtilizationPct"] > 80).sum())
    check("  credit table: three accounts above the 80% target   [CS lines 172, 415]",
          above == 3, f"found {above}")

print("-" * 78)
total = passed + failed
print(f"{passed} of {total} checks passed" + ("" if not failed else f"; {failed} FAILED (see FAIL above)"))
sys.exit(1 if failed else 0)
