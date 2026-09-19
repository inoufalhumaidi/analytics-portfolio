"""
================================================================================
Project 5 - Meridian UAV Services: Predictive Maintenance
Script:  streamlit/validate_app.py
Purpose: Run every screen of the app headlessly and assert none of them raises.

WHY "IT STARTS" IS NOT A TEST

    `streamlit run` returns HTTP 200 as soon as the server is up. The app script
    does not execute until a browser connects, and then it executes ONE screen
    -- whichever the sidebar defaults to. A KeyError on the fourth screen is
    invisible to any check that looks at the server.

    AppTest runs the script in-process, once per screen, and surfaces the
    exception if there is one. It also lets the assertions below check that each
    screen actually produced the thing it exists for, rather than rendering an
    empty page successfully.

RUN IT WITH
    python streamlit/validate_app.py

DATA DISCLOSURE
    Meridian UAV Services is fictional and all data is synthetic.
================================================================================
"""

from __future__ import annotations

import sys
from pathlib import Path

from streamlit.testing.v1 import AppTest

APP = Path(__file__).resolve().parent / "app.py"

SCREENS = [
    # (screen label, what it must have produced to count as working)
    ("The two answers",      {"min_metrics": 0, "min_tables": 1, "must_contain": "unit"}),
    ("Where the gap is",     {"min_metrics": 0, "min_tables": 2, "must_contain": "StressRatio"}),
    ("The queue",            {"min_metrics": 0, "min_tables": 2, "must_contain": "rank"}),
    ("Is the alarm useful?", {"min_metrics": 0, "min_tables": 0, "must_contain": "lead time"}),
    ("Data quality",         {"min_metrics": 0, "min_tables": 2, "must_contain": "found nothing"}),
]

passed = 0
failed = 0


def check(name: str, ok: bool, detail: str = "") -> None:
    global passed, failed
    if ok:
        passed += 1
        print(f"ok    {name}")
    else:
        failed += 1
        print(f"FAIL  {name}   {detail}")


print()
print("Meridian UAV - checking SQL and the extracts agree, then running every screen")
print("-" * 78)

"""
THE TWO SOURCES MUST RETURN THE SAME THING.

data_access falls back from SQL Server to the CSV extracts, which is what makes
the analysis reproducible without a database. That fallback is only honest if
the two paths answer the same question.

They did not. `fleet_kpi` was defined as fn_FleetKPI alone in the SQL form and
as fn_FleetKPI UNION fn_PredictionKPI in the extract, so the dataset returned 5
rows from SQL and 8 from the file. The three missing rows were the prediction
metrics -- which is to say the second finding vanished from the app whenever it
could reach a database, and appeared whenever it could not.

Nothing raised. Both sources loaded, both looked like a scorecard.
"""
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "python"))
import data_access  # noqa: E402

for name, (query, filename) in sorted(data_access.SOURCES.items()):
    if query is None:
        continue  # the extract is the canonical copy for this one, by design
    try:
        from_sql = len(data_access.load(name))
        from_csv = len(data_access.load(name, force_csv=True))
    except Exception as exc:  # noqa: BLE001
        check(f"source parity: {name}", False, str(exc)[:120])
        continue
    check(f"source parity: {name} ({from_sql} rows)", from_sql == from_csv,
          f"SQL returned {from_sql}, the extract has {from_csv}")

print("-" * 78)

for label, expect in SCREENS:
    at = AppTest.from_file(str(APP), default_timeout=300)
    at.run()

    if at.exception:
        check(f"screen renders: {label}", False, str(at.exception[0].value)[:180])
        continue

    # select the screen, then re-run
    try:
        at.sidebar.radio[0].set_value(label).run()
    except Exception as exc:  # noqa: BLE001
        check(f"screen renders: {label}", False, f"could not select: {exc}")
        continue

    if at.exception:
        check(f"screen renders: {label}", False, str(at.exception[0].value)[:180])
        continue

    check(f"screen renders: {label}", True)

    # A screen that renders nothing renders without error. Assert it produced
    # the thing it exists for.
    body = " ".join(
        [m.value for m in at.markdown] + [c.value for c in at.caption]
        + [t.value for t in at.title] + [s.value for s in at.subheader]
        + [i.value for i in at.info]
    )
    check(
        f"  ...and says something about '{expect['must_contain']}'",
        expect["must_contain"].lower() in body.lower(),
        "the screen rendered but its explanatory text is missing",
    )
    check(
        f"  ...and shows at least {expect['min_tables']} table(s)",
        len(at.dataframe) >= expect["min_tables"],
        f"found {len(at.dataframe)}",
    )

print("-" * 78)
print(f"  {passed} passed, {failed} failed")
if failed:
    print("  The app starts and at least one screen does not work.")
    sys.exit(1)
print("  Every screen renders and produces its content.")
print()
