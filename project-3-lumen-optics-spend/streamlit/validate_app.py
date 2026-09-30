"""
================================================================================
Project 3 - Lumen Optics Manufacturing: Photonics Spend Scorecard
Script:  streamlit/validate_app.py
Purpose: Run every screen of the app headlessly, assert none of them raises,
         and assert every headline figure renders exactly as the case study
         and README state it.

WHY "IT STARTS" IS NOT A TEST

    `streamlit run` returns HTTP 200 as soon as the server is up. The app script
    does not execute until a browser connects, and then it executes ONE screen
    -- whichever the sidebar defaults to. A KeyError on the fourth screen is
    invisible to any check that looks at the server.

    AppTest runs the script in-process, once per screen, and surfaces the
    exception if there is one. The checks below then ask whether each screen
    produced what it exists for, rather than rendering an empty page.

EVERY EXPECTED VALUE IS TYPED IN, NEVER READ FROM AN EXTRACT

    Each expected figure is copied from the case study or the README, with the
    line it comes from, or from a read-only SQL SELECT where neither quotes it.
    None is computed from the CSVs: a check that took its expectation from the
    same file the app reads would agree with any change to that file.

WHAT ELSE IS CHECKED, BECAUSE IT FAILED SILENTLY BEFORE

    Order: every table carrying PriorityRank must list it strictly increasing,
    before and after the queue's controls are used -- a re-ranked queue once
    passed every count check. Dollars: AppTest returns the markdown source, not
    the render, so a '$' the app forgot to escape (two of them can start LaTeX
    and swallow the amounts) is looked for in the source of every text element
    except the metric cards, which are raw HTML where '$' is literal. Streamlit
    1.64 happens not to render two bare '$' as LaTeX when a space precedes the
    closing one, so an unescaped '$' may look fine today; the check is about
    intent -- every '$' in markdown is escaped -- not about today's render.
    Colour: a card for a scorecard metric must be drawn in the status the case
    study publishes; every other card must be grey. Denominator: a screen that
    fails to render fails all of its checks, so the last line always reads
    "x of" the same total.

    Whether the committed extracts equal what SQL produces is checked by
    erp_extracts/Validate_Extracts.ps1, not here: the deployed app has no SQL
    Server, and this script runs wherever the app does.

RUN IT WITH
    python project-3-lumen-optics-spend/streamlit/validate_app.py
    (from the repository root or any other directory)

DATA DISCLOSURE
    Lumen Optics Manufacturing is fictional and all data is synthetic.
================================================================================
"""

from __future__ import annotations

import re
import sys
from pathlib import Path

import pandas as pd

from streamlit.testing.v1 import AppTest

HERE = Path(__file__).resolve().parent
APP = HERE / "app.py"

# Typed here, not imported from the app: a card must be one of these colours.
COLOUR = {"Red": "#c00000", "Amber": "#bf8f00", "Green": "#548235", "Grey": "#808080"}
NO_STATUS = "no status: nothing computes one, so neutral"

# Citations: CS = case_study/Lumen_Optics_Spend_Case_Study.md, RM = README.md
# (both in project-3-lumen-optics-spend/). SQL = a read-only SELECT against
# LumenSpend on 2026-09-30. The case study prints a minus sign as U+2212; the
# app is ASCII-only, so "-0.19%" is its rendering of the case study's figure.
SCREENS = [
    {
        "label": "Green, and the money is gone",
        "must_contain": "learning curve", "tables": 1, "charts": 1, "vendor_tables": 0,
        "cards": [
            ("Purchase price variance", "-0.19%", "Green", "CS l.18, l.360; RM l.19, l.257"),
            ("Erosion capture", "29.50%", "Red", "CS l.51, l.358; RM l.20, l.255"),
            ("Annual value of the gap", "$4,047,609", "Grey", "CS l.53, l.373; RM l.21, l.263"),
        ],
        "shows": [
            ("Target 80.00%", "CS l.52, l.358 - erosion capture target"),
            ("8.3% of", "CS l.55-56; RM l.23 - share of direct-materials spend"),
            ("$48,586,386.86", "CS l.371; RM l.262 - direct-materials spend, 12 months"),
            ("laser diodes to fall about 8% a year", "CS l.38; RM l.26 - benchmark"),
            ("fibre connectors 7%", "CS l.38; RM l.26 - benchmark"),
            ("optical coatings 5%", "CS l.38; RM l.26 - benchmark"),
            ("optical fibre 4%", "CS l.38 - benchmark"),
            ("PPV reads -0.19%", "CS l.233 - Recommendation 3 basis"),
            ("Recommendation 3: add erosion capture to the monthly pack", "CS l.233"),
        ],
        "row_sources": {"MetricName": "CS appendix l.358-365"},
        "rows": [("MetricName", key, dict(zip(("Value", "TargetValue", "WarningValue"), vals)))
                 for key, vals in [
                     ("ErosionCapturePct", (29.50, 80.00, 60.00)),
                     ("MaverickSpendPct", (27.09, 5.00, 12.00)),
                     ("PPVPct", (-0.19, 0.00, 2.00)),
                     ("AcceptanceRatePct", (99.11, 98.00, 95.00)),
                     ("OnTimeDeliveryPct", (80.62, 95.00, 90.00)),
                     ("ExpediteSpendPct", (0.23, 1.00, 3.00)),
                     ("SingleSourceSpendPct", (27.17, 20.00, 35.00)),
                     ("Top5VendorSharePct", (27.20, 50.00, 70.00))]],
    },
    {
        "label": "Where the gap sits",
        "must_contain": "LeverageScore", "tables": 3, "charts": 1, "vendor_tables": 1,
        "cards": [
            ("Vendor-part pairs assessed", "185", "Grey", "CS l.103, l.377"),
            ("Sole-source pairs", "33", "Grey", "CS l.118; RM l.35"),
            # SQL: SUM(AnnualOpportunity) / COUNT(*) over dbo.vw_RenegotiationQueue,
            # IsSingleSource = 1 -> 44,574.31 (CS l.214: "per pair ... sole-source")
            ("Gap per sole-source pair", "$44,574", "Grey", "SQL; CS l.214"),
            ("Average leverage score", "38.1", "Grey", "CS l.118; RM l.35 - sole source"),
        ],
        "shows": [
            ("carry $1,470,952", "CS l.118, l.374; RM l.35 - opportunity on sole-source parts"),
            ("36% of the gap", "CS l.130 - 36% of the total opportunity"),
            ("against $16,952 where an alternative exists", "SQL: same, IsSingleSource = 0 -> 16,951.69"),
            ("against 72.8", "CS l.117; RM l.34 - average leverage, 2+ suppliers"),
            ("capture 8.91% of the expected erosion against 38.26%",
             "CS l.117-120; RM l.34-35; SQL: sql/05 l.300 weighting by IsSingleSource"),
            ("at least $40,000 a year of opportunity, can be requalified within 8 months",
             "CS l.235; sql/05 DUAL_SOURCE rule"),
            ("8 pairs, $515,330 a year, all 8 flagged for this quarter", "CS l.208, l.235"),
            ("Recommendation 5: qualify a second source on the 8 DUAL_SOURCE pairs", "CS l.235"),
        ],
        "row_sources": {"Sourcing": "CS l.117-118; RM l.34-35; per pair: SQL"},
        "rows": [
            ("Sourcing", "Parts with 2+ qualified suppliers",
             {"Pairs": 152, "Opportunity": 2576657, "OpportunityPerPair": 16952, "AvgLeverage": 72.8}),
            ("Sourcing", "Parts with one qualified supplier",
             {"Pairs": 33, "Opportunity": 1470952, "OpportunityPerPair": 44574, "AvgLeverage": 38.1}),
        ],
        "ranked": 1,
        # SQL: TOP 3 ... WHERE IsSingleSource = 1 ORDER BY PriorityRank -> ranks 3, 5, 6
        "first": ("QualificationMonths", "PartNumber", ["LOM-0096", "LOM-0057", "LOM-0082"]),
    },
    {
        "label": "Price is not cost",
        "must_contain": "cost per accepted unit", "tables": 1, "charts": 1, "vendor_tables": 1,
        "cards": [
            ("Material paid for and unusable", "$1,571,168", "Grey", "CS l.166, l.234, l.376; RM l.114"),
            ("Of which VEN-004", "$289,024", "Grey", "CS l.159, l.234; RM l.113"),
        ],
        "shows": [
            ("The cheapest price was the dearest usable part - until the price fell further",
             "CS l.134; RM l.99 - section 4.2 heading"),
            ("18% of it", "CS l.159 - VEN-004's share of unusable material"),
            ("$227.76 in 2023 (the dearest), $192.58 in 2024 and $168.28 in 2025", "CS l.145; RM l.105"),
            ("$204.71 over all history. Average unit price $184.14", "CS l.145; RM l.105"),
            ("acceptance 91.61%", "CS l.145; RM l.105"),
            ("$212.00 in 2023, $204.35 in 2024, $189.42 in 2025; $204.44 over all history",
             "CS l.146; RM l.106"),
            ("Average unit price $197.43; acceptance 98.99%", "CS l.146; RM l.106"),
            ("$223.59 in 2023, no orders in 2024, $193.55 in 2025; $213.94 over all history",
             "CS l.147; RM l.107"),
            ("Average unit price $211.49; acceptance 99.09%", "CS l.147; RM l.107"),
            ("rejected 9.2% of what it shipped that year", "CS 4.2 ('a 9% rejection rate', 2023); SQL LOM-0072 2023 9.22"),
            ("rejects 8.39% on this part against 1.00%", "CS l.137 (8.4% vs 1.0%); 100 - 91.61"),
            ("8.4 times the rate", "CS l.138"),
            ("cheapest per accepted unit on all three", "CS l.152-153; RM l.110-111"),
            ("by $21 a unit", "CS l.153"),
            ("27-cent tie", "CS l.158; RM l.111-112"),
            ("over a recent window", "CS l.157; RM l.112"),
            ("Recommendation 4: keep VEN-004 on laser diodes, and put its rejection rate under "
             "corrective action", "CS l.234"),
            ("rejecting 8.4 times as often as its competitors", "CS l.234"),
        ],
        "row_sources": {},
        "rows": [],
    },
    {
        "label": "Green lights, failing cohorts",
        "must_contain": "cohort", "tables": 1, "charts": 1, "vendor_tables": 1,
        "cards": [
            ("Acceptance rate", "99.11%", "Green", "CS l.184, l.361; RM l.258"),
            ("On-time delivery", "80.62%", "Red", "CS l.185, l.362; RM l.259"),
            ("Expedite spend", "0.23%", "Green", "CS l.186, l.363; RM l.260"),
            ("VEN-004 on Laser Diode", "91.4%", "Grey", "CS l.184, l.189, l.367; RM l.265"),
            ("VEN-009 on time", "59.64%", "Grey", "CS l.236 - over all history"),
            ("Expedite spend at 2023-12-31", "0.12%", "Grey", "CS l.186 - 0.12% -> 0.23%"),
        ],
        "shows": [
            ("Portfolio figures hide the cohorts that are failing", "CS l.180 - section 4.4 heading"),
            ("As a vendor it reads 98.3%", "CS l.188"),
            ("83% of its 2023 orders and 20.8% of 2025's", "CS l.185, l.236"),
            ("It has doubled since, across the whole book", "CS l.186"),
            ("Acceptance is 99.11% company-wide and 91.4% for VEN-004 on Laser Diode", "CS l.237"),
            ("the worst delivery record of the 36 vendors", "CS l.236"),
            ("Recommendation 7: report every aggregate with its cohort cut", "CS l.237"),
            ("Recommendation 6: put VEN-009 on a delivery corrective action", "CS l.236"),
        ],
        "row_sources": {"VendorID": "CS l.236 (59.64%)"},
        "rows": [("VendorID", "VEN-009", {"OnTimePct": 59.64})],
        # SQL: TOP 1 ... FROM fn_VendorScorecard ORDER BY OnTimePct -> VEN-009
        "first": ("AvgDaysLate", "VendorID", ["VEN-009"]),
    },
    {
        "label": "The queue",
        "must_contain": "re-ranks nothing", "tables": 5, "charts": 0, "vendor_tables": 2,
        "cards": [
            ("Pairs in the queue", "185", "Grey", "CS l.199, l.377; RM l.126"),
            ("Workable this quarter", "43", "Grey", "CS l.202, l.378; RM l.129"),
            ("Annual value of this quarter's pairs", "$2,673,946", "Grey", "CS l.231"),
            ("Share of the gap", "66.1%", "Grey", "CS l.202, l.231, l.378; RM l.129"),
        ],
        "shows": [
            ("43 pairs carry $2,673,946 a year - 66.1% of the $4,047,609 gap", "CS l.231, l.373"),
            ("starting with the 20 PUT_ON_CONTRACT", "CS l.231 - Recommendation 1"),
            ("$1,623,978 across 48 pairs", "CS l.177-178, l.232 - PUT_ON_CONTRACT"),
            ("27.09% of spend has no agreement in force", "CS l.170, l.232 - maverick spend"),
            ("12 of 36 vendors", "SQL: Dim_Vendor with no IsThisQuarter pair in vw_RenegotiationQueue"),
            ("it states each pair's own position", "sql/05 RecommendedAction for RENEGOTIATE (per-pair since 2026-09-30)"),
            ("Recommendation 1: work the quarter's 43 pairs", "CS l.231"),
            ("Recommendation 2: re-paper the lapsed agreements", "CS l.232"),
        ],
        "row_sources": {"ActionCode": "CS l.206-210; RM l.133-137",
                        "BuyerID": "SQL: GROUP BY BuyerID over dbo.vw_RenegotiationQueue"},
        "rows": [
            ("ActionCode", "RENEGOTIATE",
             {"Pairs": 54, "Opportunity": 1631840, "Spend12m": 16221031, "ThisQuarter": 15}),
            ("ActionCode", "PUT_ON_CONTRACT",
             {"Pairs": 48, "Opportunity": 1623978, "Spend12m": 11559724, "ThisQuarter": 20}),
            ("ActionCode", "DUAL_SOURCE",
             {"Pairs": 8, "Opportunity": 515330, "Spend12m": 4652024, "ThisQuarter": 8}),
            ("ActionCode", "ACCEPT",
             {"Pairs": 73, "Opportunity": 276461, "Spend12m": 10320429, "ThisQuarter": 0}),
            ("ActionCode", "FIX_QUALITY",
             {"Pairs": 2, "Opportunity": 0, "Spend12m": 1106723, "ThisQuarter": 0}),
        ] + [("BuyerID", buyer, {"QueuePairs": pairs, "ThisQuarter": tq, "AnnualOppOnQuarterPairs": opp})
             for buyer, pairs, tq, opp in [
                 ("BUY-01", 26, 8, 1014574.93), ("BUY-02", 16, 7, 411870.17),
                 ("BUY-03", 51, 8, 524275.51), ("BUY-04", 47, 7, 403375.06),
                 ("BUY-05", 26, 8, 242616.37), ("BUY-06", 19, 5, 77234.17)]],
        "ranked": 1,
        # SQL: TOP 3 ... FROM vw_RenegotiationQueue ORDER BY PriorityRank (all IsThisQuarter = 1)
        "first": ("RecommendedAction", "PartNumber", ["LOM-0104", "LOM-0128", "LOM-0096"]),
    },
    {
        "label": "Can the data carry it?",
        "must_contain": "found nothing", "tables": 1, "charts": 0, "vendor_tables": 0,
        "cards": [
            ("Spend mis-stated", "0.497%", "Grey", "RM l.163, l.236"),
            ("Gate tolerance", "0.500%", "Grey", "RM l.163, l.236"),
            ("Largest defect by value", "$1,495,917", "Grey", "RM l.167; CS l.238 (AMBIGUOUS_CONTRACT)"),
        ],
        "shows": [
            ("A margin of 0.003 points", "RM l.163-164 - by three thousandths of a point"),
            ("ambiguous on $1,495,917 of spend", "CS l.238 - Recommendation 8"),
            ("61 PO lines", "CS l.86, l.238 - lines with two agreements in force"),
            ("resolve the 6 overlapping agreements", "CS l.238; SQL: 6 distinct agreements on the 61"),
            ("10 behavioural checks", "RM l.55 - 10 behavioural checks"),
            ("Recommendation 8", "CS l.238 - decision"),
        ],
        "row_sources": {},
        "rows": [],
    },
]

# Claims the app must NOT make: contradicted by the data (checked 2026-09-29/30)
# or superseded wording.
DISPUTED = [
    ("drives expedite", "the data does not tie expedite fees to late delivery: late lines are "
                        "expedited 4.76% of the time, on-time lines 5.03%"),
    ("driver behind rising expedite", "same: the earlier Recommendation 6 value claim"),
    ("long-lead", "CS l.186: expedite doubled across the book, not on long-lead parts"),
    ("61 overlapping agreements", "CS l.238: 6 agreements over 61 PO lines"),
    ("three largest sole-source", "SQL routes the largest sole-source pairs to PUT_ON_CONTRACT "
                                  "and RENEGOTIATE, not DUAL_SOURCE"),
    ("largest exactly where leverage", "in total the competitive gap is larger ($2,576,657 vs "
                                       "$1,470,952); sole source is larger only per pair"),
    ("undercuts every competitor", "CS l.136: on LOM-0080 in 2023 VEN-004's average unit price "
                                   "ties VEN-007's (SQL: 501.9232 each)"),
    ("seven times", "CS l.138: 8.4 times the rate"),
    ("This quarter's opportunity", "the figure is annual: the annual value of this quarter's pairs"),
]

CARD = re.compile(r'border-left:6px solid (#[0-9a-fA-F]{6});.*?letter-spacing:0.04em">(.*?)</div>'
                  r'\s*<div[^>]*>(.*?)</div>', re.S)
BARE_DOLLAR = re.compile(r"(?<!\\)\$")
TEXT_KINDS = ("markdown", "caption", "info", "success", "warning", "error", "title", "header",
              "subheader")

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


def not_run(count: int, why: str) -> None:
    """Checks that could not run are failures, so the total never shrinks."""
    global failed
    failed += count
    print(f"FAIL  {count} check(s) not run: {why}")


def planned(screen: dict) -> int:
    """How many checks the loop below runs for this screen (render included)."""
    return (6 + (1 if screen.get("ranked") else 0) + (1 if screen.get("first") else 0)
            + len(screen["cards"]) + len(screen["shows"]) + len(screen["rows"]))


def open_screen(label: str) -> tuple[AppTest | None, str]:
    at = AppTest.from_file(str(APP), default_timeout=300)
    at.run()
    if at.exception:
        return None, str(at.exception[0].value)[:180]
    try:
        at.sidebar.radio[0].set_value(label).run()
    except Exception as exc:  # noqa: BLE001
        return None, f"could not select: {exc}"
    if at.exception:
        return None, str(at.exception[0].value)[:180]
    return at, ""


def sources(at: AppTest) -> list[str]:
    """The markdown source of every text element, as the app passed it."""
    return [e.value for kind in TEXT_KINDS for e in at.get(kind)]


def page_text(at: AppTest) -> str:
    """What a reader sees as text; the app's escaped \\$ reads as $."""
    return " ".join(sources(at)).replace("\\$", "$")


def cards(at: AppTest) -> dict[str, tuple[str, str]]:
    """Label -> (value, colour) of every metric card the screen drew."""
    return {label: (value, colour.lower()) for m in at.markdown
            for colour, label, value in CARD.findall(m.value)}


def bare_dollars(at: AppTest) -> list[str]:
    """Text elements, other than metric cards (raw HTML), with an unescaped $."""
    return [s[:90] for s in sources(at) if not CARD.search(s) and BARE_DOLLAR.search(s)]


def frames(at: AppTest, marker: str) -> list:
    return [f.value for f in at.dataframe if marker in f.value.columns]


def rank_ok(at: AppTest, expected_tables: int) -> tuple[bool, str]:
    ranked = frames(at, "PriorityRank")
    bad = [list(f["PriorityRank"].head(6)) for f in ranked
           if not f["PriorityRank"].is_monotonic_increasing or not f["PriorityRank"].is_unique]
    ok = len(ranked) == expected_tables and not bad
    return ok, f"{len(ranked)} ranked table(s), expected {expected_tables}; out of order: {bad}"


def first_ok(at: AppTest, marker: str, key: str, expected: list[str]) -> tuple[bool, str]:
    found = frames(at, marker)
    got = list(found[0][key].head(len(expected))) if len(found) == 1 else None
    return got == expected, f"got {got}, expected {expected}"


def row_matches(at: AppTest, key_col: str, key: str, expected: dict) -> tuple[bool, str]:
    for frame in (f.value for f in at.dataframe):
        if key_col in frame.columns and set(expected) <= set(frame.columns):
            rows = frame[frame[key_col] == key]
            if len(rows) != 1:
                continue
            row = rows.iloc[0]
            wrong = {c: float(row[c]) for c, v in expected.items() if abs(float(row[c]) - v) > 1e-9}
            return (not wrong), f"rendered {wrong}, expected {expected}"
    return False, f"no table has a {key_col} = {key!r} row with columns {sorted(expected)}"


print()
print("Lumen Optics - running every screen and checking its figures against the case study")
print("-" * 78)

PLANNED = sum(planned(s) for s in SCREENS)
all_text = ""
for screen in SCREENS:
    label = screen["label"]
    at, error = open_screen(label)
    check(f"screen renders: {label}", at is not None, error)
    if at is None:
        not_run(planned(screen) - 1, f"'{label}' did not render")
        continue

    text = page_text(at)
    all_text += " " + text
    check(f"  ...and says something about '{screen['must_contain']}'",
          screen["must_contain"].lower() in text.lower(), "its explanatory text is missing")
    check(f"  ...and shows exactly {screen['tables']} table(s)",
          len(at.dataframe) == screen["tables"], f"found {len(at.dataframe)}")
    charts = len(at.get("vega_lite_chart"))
    check(f"  ...and draws exactly {screen['charts']} chart(s)", charts == screen["charts"],
          f"found {charts}")
    named = [list(f.columns) for f in frames(at, "VendorName")]
    unlabelled = [c for c in named if "VendorID" not in c]
    check(f"  ...and shows VendorID beside VendorName in its {screen['vendor_tables']} vendor "
          "table(s) (VEN-024 and VEN-036 share a name)",
          len(named) == screen["vendor_tables"] and not unlabelled,
          f"{len(named)} table(s) with VendorName; without VendorID: {unlabelled}")
    stray = bare_dollars(at)
    check("  ...and escapes every $ in its text (two bare $ can start LaTeX)", not stray,
          f"unescaped $ in: {stray}")
    if screen.get("ranked"):
        ok, detail = rank_ok(at, screen["ranked"])
        check("  ...and lists every PriorityRank table in strictly increasing rank (SQL's order)",
              ok, detail)
    if screen.get("first"):
        marker, key, expected = screen["first"]
        ok, detail = first_ok(at, marker, key, expected)
        check(f"  ...and its first rows are {expected} (SQL)", ok, detail)

    drawn = cards(at)
    for label_, value, rag, source in screen["cards"]:
        got = drawn.get(label_)
        why = NO_STATUS if rag == "Grey" else "the status the case study publishes"
        check(f"  card {label_!r} reads {value!r} in {rag} ({why})  [{source}]",
              got == (value, COLOUR[rag]), f"it reads {got!r}")
    for expected, source in screen["shows"]:
        check(f"  shows {expected!r}  [{source}]", expected in text, "not found in the rendered text")
    for key_col, key, expected in screen["rows"]:
        ok, detail = row_matches(at, key_col, key, expected)
        check(f"  table row {key_col}={key}  [{screen['row_sources'][key_col]}]", ok, detail)

print("-" * 78)

# The sidebar is the same on every screen: read it once.
SIDEBAR_CHECKS = 3
PLANNED += SIDEBAR_CHECKS
at, error = open_screen(SCREENS[0]["label"])
if at is None:
    not_run(SIDEBAR_CHECKS, f"sidebar: {error}")
else:
    side = " ".join([m.value for m in at.sidebar.markdown] + [c.value for c in at.sidebar.caption])
    check("sidebar: reporting date 2025-12-31 (CS l.4), read from the extract",
          "2025-12-31" in side, side[:120])
    check("sidebar: names the source as the committed CSV extracts",
          "Source: committed CSV extracts (erp_extracts/), exported from SQL Server" in side, side[:120])
    check("sidebar: carries the fictional / synthetic data disclosure",
          "fictional" in side and "synthetic" in side, side[:120])

# The controls re-cut the queue SQL ranked. Each state must list the right pairs
# AND keep SQL's order: a filter may drop rows, never reorder them.
QUEUE_STATES = [
    ("listed by default = the 43 this-quarter pairs (CS l.202)", 43, None),
    ("unticking 'Only this quarter' lists all 185 pairs (CS l.377)", 185, ["LOM-0104", "LOM-0128",
                                                                          "LOM-0096"]),
    ("one buyer's quarter = that buyer's flagged pairs (BUY-06, 5 - SQL)", 5, None),
    ("the DUAL_SOURCE quarter = all 8 DUAL_SOURCE pairs (CS l.208, l.235)", 8, None),
]
PLANNED += len(QUEUE_STATES)


def widget(app: AppTest, kind: str, label: str):
    [found] = [w for w in app.get(kind) if w.label == label]
    return found


QUEUE_STEPS = [  # what to do to reach each state, from the one before it
    lambda app: None,
    lambda app: widget(app, "checkbox", "Only this quarter").uncheck().run(),
    lambda app: (widget(app, "checkbox", "Only this quarter").check().run(),
                 widget(app, "multiselect", "Buyer").select("Tomas Varga").run()),
    lambda app: (widget(app, "multiselect", "Buyer").unselect("Tomas Varga").run(),
                 widget(app, "multiselect", "Action").select("DUAL_SOURCE").run()),
]
at, error = open_screen("The queue")
done = 0
if at is not None:
    try:
        for step, (name, rows, first) in zip(QUEUE_STEPS, QUEUE_STATES):
            step(at)
            if at.exception:
                raise RuntimeError(str(at.exception[0].value)[:180])
            found = frames(at, "RecommendedAction")
            listed = len(found[0]) if len(found) == 1 else -1
            ordered, detail = rank_ok(at, 1)
            top_ok, top = (first_ok(at, "RecommendedAction", "PartNumber", first)
                           if first else (True, ""))
            check(f"queue control: {name}, in strictly increasing PriorityRank",
                  listed == rows and ordered and top_ok, f"listed {listed}; {detail} {top}")
            done += 1
    except Exception as exc:  # noqa: BLE001 -- a missing control is a failure, not a crash
        error = f"a control could not be used: {exc}"[:180]
if done < len(QUEUE_STATES):
    not_run(len(QUEUE_STATES) - done, f"queue controls: {error}")

print("-" * 78)
PLANNED += len(DISPUTED)
for phrase, why in DISPUTED:
    check(f"does not repeat a disputed claim: {phrase!r}", phrase.lower() not in all_text.lower(),
          why)

# Static checks on the three deployed files (spec: ASCII-only .py, exact pins).
print("-" * 78)
STATIC_CHECKS = 4
PLANNED += STATIC_CHECKS
for name in ("app.py", "validate_app.py"):
    raw = (HERE / name).read_bytes()
    check(f"static: {name} is ASCII-only", all(b < 128 for b in raw),
          f"first non-ASCII byte at offset {next((i for i, b in enumerate(raw) if b > 127), -1)}")
pins = sorted(line.strip() for line in (HERE / "requirements.txt").read_text().splitlines()
              if line.strip() and not line.lstrip().startswith("#"))
check("static: requirements.txt pins exactly pandas==3.0.6 and streamlit==1.64.0",
      pins == ["pandas==3.0.6", "streamlit==1.64.0"], f"found {pins}")
# The queue's instruction must not claim an alternative where none is qualified:
# SQL printed "a credible alternative exists" on five sole-source pairs until
# 2026-09-30. Read from the extract the app displays.
_q = pd.read_csv(HERE.parent / "erp_extracts" / "renegotiation_queue.csv")
_bad = _q[(_q["IsSingleSource"] == 1) & _q["RecommendedAction"].str.contains("alternative exists", case=False)]
check("data: no sole-source pair's instruction claims a qualified alternative",
      len(_bad) == 0, f"{len(_bad)} rows, e.g. {_bad['PartNumber'].head(3).tolist()}")

print("-" * 78)
total = passed + failed
if total != PLANNED:
    print(f"FAIL  bookkeeping: {total} checks counted, {PLANNED} planned")
    failed += 1
    total = PLANNED + 1
if failed:
    print(f"  {failed} check(s) failed -- see FAIL above.")
print(f"{passed} of {total} checks passed")
sys.exit(1 if failed else 0)
