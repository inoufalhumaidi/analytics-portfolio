"""
================================================================================
Project 4 - Talon Robotics: Payload Deployment Delivery
Script:  streamlit/validate_app.py
Purpose: Run every screen of the app headlessly, assert none of them raises,
         and assert every headline figure renders exactly as the case study
         states it.

WHY "IT STARTS" IS NOT A TEST

    `streamlit run` returns HTTP 200 as soon as the server is up. The app script
    does not execute until a browser connects, and then it executes ONE screen
    -- whichever the sidebar defaults to. A KeyError on the fourth screen is
    invisible to any check that looks at the server.

    AppTest runs the script in-process, once per screen, and surfaces the
    exception if there is one. It also lets the assertions below check that each
    screen actually produced the thing it exists for, rather than rendering an
    empty page successfully.

    A screen that raises does not skip its checks: every one of them runs
    against an empty page and FAILS, so the final 'N of M' has the same M on a
    broken run as on a clean one, and a broken screen cannot shrink the count.

WHY THE EXPECTED FIGURES ARE TYPED IN, NOT READ FROM THE EXTRACTS

    A check that reads its expected value from the same CSV the app reads
    cannot fail: a changed extract moves the app and the expectation together,
    and the page goes on disagreeing with the case study in silence. So every
    headline figure below is the string the case study or README prints, with
    the line it is printed on. If one of these fails, either the extract or the
    case study is wrong -- find out which; do not edit the expectation to match.

    The extracts are read here for five things only, none of them a headline
    figure:
      - the rig-week re-cut: at 180 the rows the app lists must be SQL's own
        IsThisWeek rows, row for row; at 360 they must be SQL's rule
        (CumulativeRigHours <= capacity AND AirframeHours = 0, sql/05
        fn_VerificationQueue) applied here to the CSV -- so a re-ranked queue,
        or a week that books flights, fails;
      - the trend row the gap card reads must agree with the scorecard row;
      - a card that restates a SQL data-quality check must show that check's
        count, not a count re-derived in pandas;
      - the exposure checks the quoted gate result rests on must not have moved
        since it was quoted (if they have, re-run the gate).
    Three SQL files are read as text, never run: the app's quoted constants
    must equal the literals they quote (the gate tolerance in sql/04, the Field
    run price in sql/01, the rig-week capacity in sql/05).

    Not asserted, because the app does not show them: '19 of them must-ship'
    and 'overlap by 17' (CS l.184-185) -- no extract carries either without
    re-deriving SQL's check in pandas. And CS l.230-231's queue-wide "91 items
    need one written first" is a finding, not an expectation: the extract's
    NeedsNewCase sums to 99 (91 RAISE_TEST_LEVEL + 8 VERIFY); the 91 asserted
    below is Rec 5's, for RAISE_TEST_LEVEL (CS l.290).

RUN IT WITH
    python project-4-talon-robotics/streamlit/validate_app.py   (from any cwd)

DATA DISCLOSURE
    Talon Robotics is fictional and all data is synthetic.
================================================================================
"""

from __future__ import annotations

import ast
import re
import sys
from pathlib import Path

import pandas as pd
from streamlit.testing.v1 import AppTest

HERE = Path(__file__).resolve().parent
APP = HERE / "app.py"
EXTRACTS = HERE.parent / "erp_extracts"
SQL = HERE.parent / "sql"

# CS = case_study/Talon_Payload_Delivery_Case_Study.md, RM = README.md; l. = line.
SCREENS = [
    # (screen label, what it must have produced to count as working)
    ("Complete or ready?",         {"min_tables": 1, "must_contain": "one build"}),
    ("Tested is not verified",     {"min_tables": 4, "must_contain": "still applies"}),
    ("Where the gap is",           {"min_tables": 1, "must_contain": "still being designed"}),
    ("The verification queue",     {"min_tables": 2, "must_contain": "does not re-rank"}),
    ("Can the answer be trusted?", {"min_tables": 2, "must_contain": "found nothing"}),
]

# Every metric card, label -> value exactly as the case study or README prints it.
EXPECTED_CARDS = {
    "Complete or ready?": {
        "Work items closed": "93.86%",                                   # CS l.25, RM l.17
        "Requirements with passing evidence": "96.53%",                  # CS l.26, RM l.18
        "Must-ship requirements ready to ship": "51.96%",                # CS l.27, RM l.19
        "The gap": "41.90 points",                                       # CS l.28, RM l.20
        "Must-ship without current, admissible evidence": "135 of 281",  # CS l.286 (Rec 1)
        "Stale verification": "24.48%",                                  # CS l.131, RM l.266
        "Policy compliance": "77.45%",                                   # CS l.192, RM l.267
        "Critical RAID items open": "9",                                 # RM l.269
    },
    "Tested is not verified": {
        "Never tested": "18",                                            # CS l.125
        "Evidence overtaken by change": "95",                            # CS l.130
        "Evidence that never counted": "113",                            # CS l.132, l.166
        "Passed below the required level": "110",                        # CS l.182, RM l.126
        "Passed by their own owner": "20",                               # CS l.184, RM l.128
        "Policy compliance": "77.45%",                                   # CS l.192
    },
    "Where the gap is": {
        "Flight Control Interface": "23.40%",                            # CS l.38, l.144
        "Payload Release Mechanism": "32.26%",                           # CS l.38, l.145
        "Their share of what must ship": "39%",                          # CS l.153, RM l.38
    },
    "The verification queue": {
        "Outstanding requirements": "226",                               # CS l.41, l.222
        "In the first rig week": "69",                                   # CS l.43 "Sixty-nine", l.226; RM l.140
        "Rig hours to clear": "425.0",                                   # CS l.41, l.226 "425"; RM l.271 "425.0"
        "Rig-weeks to clear": "2.4",                                     # CS l.42, l.226; RM l.140, l.271
        "Airframe hours to clear": "272.0",                              # CS l.42, l.227 "272"; RM l.272 "272.0"
        "Flights": "34",                                                 # CS l.42, l.227; RM l.141, l.272
    },
    "Can the answer be trusted?": {
        "Must-ship resting on flagged data": "1.779%",                   # CS l.265, RM l.192, l.247, l.250
        "Gate tolerance": "2.000%",                                      # CS l.265, RM l.192; sql/04 l.301
        "Critical RAID items open": "9",                                 # RM l.269
        "Open RAID items overdue": "57.78%",                             # CS l.292, RM l.270
        "Due before they were raised": "7",                              # CS l.276 "Seven", l.292
    },
}

# A card showing a board metric must carry the status the README board gives
# it (RM l.262-270), so a colour cannot contradict the thresholds beside it.
EXPECTED_COLOURS = {
    "Complete or ready?": {"Must-ship requirements ready to ship": "Red",   # RM l.264
                           "Stale verification": "Amber",                   # RM l.266 (T15/W30, 24.48)
                           "Policy compliance": "Red",                      # RM l.267
                           "Critical RAID items open": "Red"},              # RM l.269
    "Tested is not verified": {"Policy compliance": "Red"},                 # RM l.267
    "Can the answer be trusted?": {"Critical RAID items open": "Red",       # RM l.269
                                   "Open RAID items overdue": "Red"},       # RM l.270
}
COLOUR_NAMES = {"#c00000": "Red", "#bf8f00": "Amber", "#548235": "Green", "#808080": "Grey"}

# Cards that restate a SQL data-quality check (sql/04_data_quality_checks.sql)
# must show that check's own count from data_quality_summary.csv. What is
# tested is the DEFINITION (no pandas re-derivation), not the figure.
SQL_CHECK_CARDS = {
    "Tested is not verified": {"Passed below the required level": "UNDER_LEVELLED_VERIFICATION",
                               "Passed by their own owner": "SELF_VERIFIED_SAFETY"},
    "Can the answer be trusted?": {"Due before they were raised": "RAID_DUE_BEFORE_RAISED"},
}

# The gate result is quoted, not computed (usp_RunDataQualityChecks prints it).
# It rests on the checks with CountsTowardExposure = 1; these are their counts
# when it was quoted. If the extract moves, the quote is stale: re-run the gate.
GATE_EXPOSURE_COUNTS = {"RUN_BEFORE_BUILD": 14, "ORPHAN_DIMENSION_KEY": 0, "REQ_WITHOUT_TESTCASE": 0}

# The app's quoted SQL constants, and the literal each must equal.
SQL_LITERALS = {
    "SQL_GATE_TOLERANCE_PCT": (SQL / "04_data_quality_checks.sql",
                               r"@MaxAffectedMustShipPct\s+DECIMAL\(\d+,\s*\d+\)\s*=\s*([\d.]+)"),
    "SQL_FIELD_HOURS_PER_RUN": (SQL / "01_create_schema.sql", r"\('Field',\s*\d+,\s*([\d.]+)"),
    "SQL_RIG_HOURS_PER_WEEK": (SQL / "05_kpi_views.sql",
                               r"CREATE VIEW dbo\.vw_VerificationQueue AS.*?,\s*([\d.]+)\);"),
}

# Text a screen must carry because a figure it shows is a floor, or because
# the page once said otherwise (reviews r5 and r7, Project 4).
CAVEATS = {
    "The verification queue": [
        "is a floor",                      # CS l.231: a test to be written is priced at one run
        "never divided by rig capacity",   # CS l.233
        "costed, not scheduled",           # CS l.234: the data holds no airframe capacity
    ],
    "Can the answer be trusted?": [
        "It passes",                       # CS l.265: the screen answers its own question
        "Why it passes",                   # review r5 I4: and says why
        "IsOverdue excludes them",         # CS l.277
    ],
}

# Claims the extracts contradict, or figures they no longer carry: none may come back.
FORBIDDEN = {
    "Tested is not verified": ["re-witness the 3"],                   # old Rec 4 (now 2 + 1)
    "The verification queue": ["610", "3.4 rig", "re-ran what was cheap",   # 82 of 226 are Integration-
                               "every hour left is hardware time",     # level; 64 cost no hardware time
                               "Rig contention is the cause"],
    "Can the answer be trusted?": ["The catalogue sentence"],        # catalogue corrected at source
}

# Numbers the case study quotes inside a sentence, as the app must print them.
EXPECTED_TEXT = {
    "Complete or ready?": [
        "report readiness and completion side by side",   # CS l.291 (Rec 6)
        "never closed below 42 points",                   # CS l.291 "at least 42 points"
    ],
    "Tested is not verified": [
        "3.5% of 519",                                    # CS l.125 "18 requirements -- 3.5%", l.32
        "together they are the 113",                      # CS l.185-186, RM l.129
        "Schedule the 110 under-levelled requirements explicitly",   # CS l.290 (Rec 5)
        "62 need a rig slot and 12 a flight",             # CS l.290
        "91 need a test written first",                   # CS l.290
        "Countersign the 2 self-verified requirements a signature alone would fix",  # CS l.289
        "(REQ-0124 and REQ-0002)",                        # CS l.289
        "re-run REQ-0488 under a witness",                # CS l.289
        "Owner: Quality",                                 # CS l.289
    ],
    "Where the gap is": [
        "109 of 281 must-ship",                           # CS l.37
        "carry 172 must-ship",                            # RM l.36
        "between **59% and 83%**",                        # CS l.38-39, RM l.36
        "15 to 23 builds",                                # RM l.36
        "At 45 and 42 builds changed",                    # CS l.287 (Rec 2)
    ],
    "The verification queue": [
        "425 rig hours outstanding against 180 a week, plus 272 airframe hours on one aircraft",  # CS l.288
        "A seventh rig shortens the rig work but not the flights",                                 # CS l.288
        "2 of the 3 here",                                # CS l.248 "two of the three"
    ],
    "Can the answer be trusted?": [
        "11 behavioural checks run",                      # CS l.268 "Eleven"
        "3 of them found nothing",                        # CS l.268 "three found nothing"
    ],
}

# Each screen ends in a decision: the recommendation, and who acts on it.
DECISIONS = {
    "Complete or ready?":         ("Do not ship build 88", "Owner: programme board"),     # CS l.286
    "Tested is not verified":     ("under-levelled", "Owner: test manager"),              # CS l.290
    "Where the gap is":           ("Freeze", "Owner: chief engineer"),                    # CS l.287
    "The verification queue":     ("rig and airframe time as the schedule driver",
                                   "Owner: programme board"),                             # CS l.288
    "Can the answer be trusted?": ("Re-baseline the 7 RAID items", "Owner: PMO"),         # CS l.292
}
# ...and the trust screen's decision must carry the gate's answer (CS l.265).
DECISION_ALSO = {"Can the answer be trusted?": "The gate passes at 1.779% against 2.000%"}

# The board on screen 1: README "Headline figures at build 88", l.262-272.
EXPECTED_BOARD = {
    "ShipReadinessPct": "51.96", "VerificationCoveragePct": "96.53",
    "StaleVerificationPct": "24.48", "PolicyCompliancePct": "77.45",
    "WorkItemCompletionPct": "93.86", "CriticalRAIDOpen": "9",
    "OverdueRAIDPct": "57.78", "RigHoursOutstanding": "425.0", "AirframeHoursOutstanding": "272.0",
}
EXPECTED_BOARD_NOTES = {"RigHoursOutstanding": "2.4 rig-weeks",      # RM l.271
                        "AirframeHoursOutstanding": "34 flights"}    # RM l.272

# CS Finding 1 table, l.118-123: state -> (requirements, of which must-ship).
EXPECTED_STATES = {
    "Current": (293, 146), "Stale - subsystem changed": (88, 54),
    "Stale - requirement changed": (7, 3), "Insufficient evidence": (113, 70),
    "No evidence": (18, 8), "Total baselined": (519, 281),
}

# CS Finding 2 table, l.144-151: name -> (must-ship, current, readiness, builds changed).
EXPECTED_SUBSYSTEMS = {
    "Flight Control Interface":   (47, 11, "23.40", 45),
    "Payload Release Mechanism":  (62, 20, "32.26", 42),
    "Power Management":           (37, 22, "59.46", 15),
    "Command & Telemetry Link":   (43, 26, "60.47", 23),
    "Structural Interface":       (19, 13, "68.42", 15),
    "Ground Control Software":    (28, 20, "71.43", 21),
    "Navigation & Sensor Fusion": (33, 24, "72.73", 17),
    "Diagnostics & Logging":      (12, 10, "83.33", 18),
}

# Queue actions: CS l.248-249 / l.289 (3 witness), l.290 / l.182 (110 raise
# level), l.119 (88 subsystem changed), l.120 (7 requirement changed), l.122 (18).
EXPECTED_ACTIONS = {"INDEPENDENT_WITNESS": 3, "RAISE_TEST_LEVEL": 110, "RERUN": 88,
                    "REVIEW_THEN_RERUN": 7, "VERIFY": 18}
# ...and what the 110 RAISE_TEST_LEVEL items need: CS l.290 (Rec 5).
EXPECTED_RAISE = {"NeedRig": 62, "NeedFlight": 12, "NewTestFirst": 91}

CARD = re.compile(r'border-left:6px solid (#[0-9a-fA-F]{6});.*?letter-spacing:0.04em">(.*?)</div>'
                  r'\s*<div[^>]*>(.*?)</div>', re.S)

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


# Every reader below takes None for a screen that did not render and returns an
# empty page, so that screen's checks run and fail instead of being skipped.
def card_blocks(at: AppTest | None) -> dict[str, tuple[str, str]]:
    found = {}
    for block in ([] if at is None else at.markdown):
        for colour, label, value in CARD.findall(block.value):
            found[label.strip()] = (value.strip(), COLOUR_NAMES.get(colour.lower(), colour))
    return found


def cards(at: AppTest | None) -> dict[str, str]:
    return {label: value for label, (value, _) in card_blocks(at).items()}


def body(at: AppTest | None) -> str:
    if at is None:
        return ""
    return " ".join(
        [m.value for m in at.markdown] + [c.value for c in at.caption]
        + [t.value for t in at.title] + [s.value for s in at.subheader]
        + [i.value for i in at.info] + [s.value for s in at.success]
    )


def decided(at: AppTest | None) -> str:
    return "" if at is None else " ".join(s.value for s in at.success)


def tables(at: AppTest | None) -> list[pd.DataFrame]:
    return [] if at is None else [frame.value for frame in at.dataframe]


def table_with(at: AppTest | None, *columns: str) -> pd.DataFrame | None:
    for frame in tables(at):
        if all(c in frame.columns for c in columns):
            return frame
    return None


def open_screen(label: str) -> AppTest | None:
    at = AppTest.from_file(str(APP), default_timeout=120)
    at.run()
    if not at.exception:
        try:
            at.sidebar.radio[0].set_value(label).run()
        except Exception as exc:  # noqa: BLE001
            check(f"screen renders: {label}", False, f"could not select: {exc}")
            return None
    if at.exception:
        check(f"screen renders: {label}", False, str(at.exception[0].value)[:180])
        return None
    check(f"screen renders: {label}", True)
    return at


def week_rows(at: AppTest | None) -> list[str]:
    """RequirementIDs the app lists as the rig week, in the order it lists them."""
    frame = table_with(at, "RequirementID", "CumulativeRigHours")
    return [] if frame is None else list(frame["RequirementID"])


print()
print("Talon Robotics - running every screen and checking it against the case study")
print("-" * 78)

# --- static checks -----------------------------------------------------------
for path in (APP, Path(__file__).resolve(), HERE / "requirements.txt"):
    text = path.read_bytes()
    bad = [i for i, b in enumerate(text) if b > 127]
    check(f"ASCII only: {path.name}", not bad, f"first non-ASCII byte at offset {bad[:1]}")

tree = ast.parse(APP.read_text(encoding="ascii", errors="replace"))
imported, constants = set(), {}
for node in ast.walk(tree):
    if isinstance(node, ast.Import):
        imported |= {a.name.split(".")[0] for a in node.names}
    elif isinstance(node, ast.ImportFrom):
        imported.add((node.module or "").split(".")[0])
for node in tree.body:  # module-level NAME = literal, and A, B = x, y
    if isinstance(node, ast.Assign) and len(node.targets) == 1:
        target = node.targets[0]
        names = [target] if isinstance(target, ast.Name) else list(getattr(target, "elts", []))
        try:
            values = ast.literal_eval(node.value)
        except ValueError:
            continue
        values = [values] if len(names) == 1 else list(values)
        constants.update({n.id: v for n, v in zip(names, values) if isinstance(n, ast.Name)})
extra = imported - {"__future__", "pathlib", "pandas", "streamlit"}
check("app.py imports only pandas and streamlit (no SQL driver, no chart library)",
      not extra, f"also imports {sorted(extra)}")

pins = [ln.strip() for ln in (HERE / "requirements.txt").read_text().splitlines()
        if ln.strip() and not ln.startswith("#")]
check("requirements.txt pins exactly streamlit==1.64.0 and pandas==3.0.6",
      pins == ["streamlit==1.64.0", "pandas==3.0.6"], f"found {pins}")

for name, (sql_file, pattern) in SQL_LITERALS.items():
    found = re.search(pattern, sql_file.read_text(encoding="utf-8", errors="replace"), re.S)
    literal = float(found.group(1)) if found else None
    check(f"app constant {name} = {constants.get(name)} is the literal in sql/{sql_file.name} ({literal})",
          literal is not None and constants.get(name) == literal, "moved, or not found")

# --- the cross-extract facts the app relies on -------------------------------
kpi = pd.read_csv(EXTRACTS / "readiness_kpi.csv", encoding="utf-8-sig").iloc[0]
trend = pd.read_csv(EXTRACTS / "readiness_trend.csv", encoding="utf-8-sig").set_index("BuildNumber")
at_rc = trend.loc[int(kpi["AsOfBuild"])]
shared = ["WorkItemCompletionPct", "VerificationCoveragePct", "ShipReadinessPct", "StaleVerificationPct"]
drift = [c for c in shared if abs(float(at_rc[c]) - float(kpi[c])) > 0.005]
check("readiness_trend at the as-of build agrees with readiness_kpi (the gap card reads the trend)",
      not drift, f"differ on {drift}")

queue_csv = pd.read_csv(EXTRACTS / "verification_queue.csv", encoding="utf-8-sig").sort_values("PriorityRank")
sql_week = list(queue_csv.loc[queue_csv["IsThisWeek"] == 1, "RequirementID"])  # SQL's own flag


def sql_rule_at(capacity: float) -> list[str]:
    """SQL's IsThisWeek rule (sql/05 fn_VerificationQueue) at another capacity."""
    fits = (queue_csv["CumulativeRigHours"] <= capacity) & (queue_csv["AirframeHours"] == 0)
    return list(queue_csv.loc[fits, "RequirementID"])


dq_csv = pd.read_csv(EXTRACTS / "data_quality_summary.csv", encoding="utf-8-sig").set_index("AnomalyType")
exposure = {k: int(v) for k, v in dq_csv.loc[dq_csv["CountsTowardExposure"] == 1, "Anomalies"].items()}
check("gate: the exposure checks behind the quoted 1.779% have not moved (else re-run the gate)",
      exposure == GATE_EXPOSURE_COUNTS, f"extract now has {exposure}")

print("-" * 78)

# --- every screen ---------------------------------------------------------------
not_rendered = []
for label, expect in SCREENS:
    at = open_screen(label)
    if at is None:
        not_rendered.append(label)

    text = body(at)
    if label == SCREENS[0][0]:
        # The sidebar is the same on every screen, so it is checked once.
        side = "" if at is None else " ".join(
            [m.value for m in at.sidebar.markdown] + [c.value for c in at.sidebar.caption])
        check("  ...sidebar reports as of 2026-09-30, build 88",           # CS l.7
              "2026-09-30" in side and "**Build** 88" in side, f"sidebar reads {side[:120]!r}")
        check("  ...sidebar names the source and discloses synthetic data",
              "Source: committed CSV extracts (erp_extracts/), exported from SQL Server" in side
              and "fictional" in side and "synthetic" in side, "source line or disclosure missing")

    check(f"  ...and says something about '{expect['must_contain']}'",
          expect["must_contain"].lower() in text.lower(),
          "the screen rendered but its explanatory text is missing")
    check(f"  ...and shows at least {expect['min_tables']} table(s)",
          len(tables(at)) >= expect["min_tables"], f"found {len(tables(at))}")

    shown, blocks = cards(at), card_blocks(at)
    for card_label, value in EXPECTED_CARDS[label].items():
        check(f"  ...card '{card_label}' reads {value}", shown.get(card_label) == value,
              f"rendered {shown.get(card_label)!r}")
    for card_label, colour in EXPECTED_COLOURS.get(label, {}).items():
        got = blocks.get(card_label, (None, None))[1]
        check(f"  ...card '{card_label}' is {colour}, as the README board scores it", got == colour,
              f"rendered {got!r}")

    for card_label, anomaly in SQL_CHECK_CARDS.get(label, {}).items():
        want = str(int(dq_csv.loc[anomaly, "Anomalies"]))
        check(f"  ...card '{card_label}' is SQL's {anomaly} count ({want}), not a pandas re-derivation",
              shown.get(card_label) == want, f"rendered {shown.get(card_label)!r}")

    for phrase in EXPECTED_TEXT.get(label, []):
        check(f"  ...says '{phrase}'", phrase in text, "not found in the rendered text")
    for phrase in CAVEATS.get(label, []):
        check(f"  ...carries the caveat '{phrase}'", phrase in text, "not found in the rendered text")
    for phrase in FORBIDDEN.get(label, []):
        check(f"  ...no longer says '{phrase}'", at is not None and phrase not in text,
              "found in the rendered text" if at is not None else "screen did not render")

    action, owner = DECISIONS[label]
    decision_text = decided(at)
    check(f"  ...ends in the decision '{action}' ({owner})",
          action in decision_text and owner in decision_text, f"decision block reads {decision_text[:120]!r}")
    if label in DECISION_ALSO:
        check(f"  ...and the decision says '{DECISION_ALSO[label]}'", DECISION_ALSO[label] in decision_text,
              f"decision block reads {decision_text[:120]!r}")

    if label == "Complete or ready?":
        board = table_with(at, "Metric")
        values = {} if board is None else dict(zip(board["Metric"], board["Value"]))
        notes = {} if board is None or "Note" not in board else dict(zip(board["Metric"], board["Note"]))
        for metric, value in EXPECTED_BOARD.items():
            check(f"  ...board {metric} = {value}", values.get(metric) == value,
                  f"rendered {values.get(metric)!r}")
        for metric, note in EXPECTED_BOARD_NOTES.items():
            check(f"  ...board {metric} notes '{note}'", note in str(notes.get(metric, "")),
                  f"rendered {notes.get(metric)!r}")
        check("  ...board shows no Status column (RAG is not in the extracts)",
              board is not None and "Status" not in board.columns, "a status was derived")

    elif label == "Tested is not verified":
        states = table_with(at, "Verification state")
        got = {} if states is None else {
            r["Verification state"]: (int(r["Requirements"]), int(r["Of which must-ship"]))
            for _, r in states.iterrows()}
        for state, pair in EXPECTED_STATES.items():
            check(f"  ...state '{state}' = {pair[0]} / {pair[1]} must-ship", got.get(state) == pair,
                  f"rendered {got.get(state)}")

    elif label == "Where the gap is":
        sub = table_with(at, "SubsystemName")
        got = {} if sub is None else {
            r["SubsystemName"]: (int(r["MustShip"]), int(r["CurrentMustShip"]),
                                 f"{float(r['ReadinessPct']):.2f}", int(r["BuildsChanged"]))
            for _, r in sub.iterrows()}
        for name, row in EXPECTED_SUBSYSTEMS.items():
            check(f"  ...{name}: {row[0]} must-ship, {row[1]} current, {row[2]}%, {row[3]} builds",
                  got.get(name) == row, f"rendered {got.get(name)}")

    elif label == "The verification queue":
        mix = table_with(at, "ActionCode", "Requirements")
        got = {} if mix is None else dict(zip(mix["ActionCode"], mix["Requirements"].astype(int)))
        for code, n in EXPECTED_ACTIONS.items():
            check(f"  ...action {code}: {n} requirements", got.get(code) == n, f"rendered {got.get(code)}")
        raise_row = {} if mix is None else mix.set_index("ActionCode").loc["RAISE_TEST_LEVEL"].to_dict()
        for column, n in EXPECTED_RAISE.items():
            check(f"  ...RAISE_TEST_LEVEL {column}: {n}", raise_row.get(column) == n,
                  f"rendered {raise_row.get(column)!r}")

        # The re-cut at SQL's capacity must BE SQL's answer, row for row -- not
        # the same count reached another way.
        listed = week_rows(at)
        check(f"  ...re-cut at 180 lists exactly SQL's {len(sql_week)} IsThisWeek rows, in rank order",
              listed == sql_week, f"app lists {len(listed)} rows; first difference at "
              f"{next((i for i, (a, b) in enumerate(zip(listed, sql_week)) if a != b), min(len(listed), len(sql_week)))}")

        # The control must change the answer by SQL's rule, and nothing else.
        moved, error = None, "screen did not render"
        if at is not None:
            try:
                at.slider[0].set_value(360.0).run()
                moved, error = (None, str(at.exception[0].value)[:120]) if at.exception else (at, "")
            except Exception as exc:  # noqa: BLE001
                error = str(exc)[:120]
        after = cards(moved)
        want = sql_rule_at(360.0)
        check("  ...capacity slider moves to 360 rig hours without raising", moved is not None, error)
        check(f"  ...at 360 it lists SQL's rule at 360 ({len(want)} rows: CumulativeRigHours <= 360, "
              "no airframe hours), row for row", week_rows(moved) == want,
              f"app lists {len(week_rows(moved))} rows")
        check(f"  ...at 360 the week card reads {len(want)}", after.get("In the first rig week") == str(len(want)),
              f"rendered {after.get('In the first rig week')!r}")
        check("  ...at 360 rig-weeks are 425.0 / 360 = 1.2",                 # 425.0: RM l.271
              after.get("Rig-weeks to clear") == f"{425.0 / 360:.1f}", f"rendered {after.get('Rig-weeks to clear')!r}")
        check("  ...at 360 the airframe work and the queue do not move (272.0 h, 34 flights, 226)",
              (after.get("Airframe hours to clear"), after.get("Flights"), after.get("Outstanding requirements"))
              == ("272.0", "34", "226"), f"cards after the move: {after}")

print("-" * 78)
total = passed + failed
if not_rendered:
    print(f"  {len(not_rendered)} screen(s) did not render; every check on them was run and FAILED: "
          f"{not_rendered}")
if failed:
    print(f"  At least one check failed -- see FAIL above. {failed} failed.")
print(f"  {passed} of {total} checks passed")
sys.exit(1 if failed else 0)
