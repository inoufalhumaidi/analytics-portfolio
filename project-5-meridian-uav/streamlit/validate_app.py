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

import datetime
import math
import re
import sys
from collections import Counter
from decimal import Decimal
from pathlib import Path

import numpy as np
import pandas as pd
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
not_run = 0


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

"""
...AND THEY MUST AGREE ON VALUES, NOT ON ROW COUNTS. A CHECK THAT CANNOT RUN
IS NOT A PASS.

This block used to compare len() alone. A redefined feature, a unit error or a
shifted decimal all keep the row count, so all of them passed. Worse, with no
SQL Server BOTH loads read the same CSV: twelve of this script's checks compared
a file with itself, could not fail, and were still counted as passes -- on
exactly the machines (a reviewer's laptop, Streamlit Community Cloud) where
nobody would notice.

Now every value the extract carries is compared with SQL, and without SQL
Server the comparison is reported as NOT RUN, and not counted.
"""

# How closely an extract must match SQL: to 4 decimals, for EVERY dataset.
# sensor_features once needed wider, per-column tolerances because its extract
# rounded features that its SQL form did not. Both now run the same rounded
# query (the rounding is a modelling decision, see data_access.py), so a
# difference beyond the fourth decimal is drift, whichever column it is in.
DEFAULT_TOLERANCE = 5e-5


# Columns a SQL form returns that its extract deliberately does not carry:
# surrogate keys, the as-of date and parameters echoed back, and filter columns.
# Each was checked against the two consumers -- streamlit/app.py and
# python/degradation_model.py -- and neither reads it. ANY OTHER column SQL
# returns and the extract lacks FAILS. That is how component_life lost
# IsDateValid: the model filters on it, so whenever it ran without a database it
# scored 5 invalid lives and its own reconciliation failed -- while this check,
# which only looked at the extract's columns, reported the two sources agreed.
NOT_EXTRACTED = {
    "action_queue": {"AsOfDate"},
    "alert_evaluation": {"AsOfDate", "VibThreshold"},
    "base_scorecard": {"AsOfDate", "BaseKey"},
    "component_life": {"AirframeKey"},
    "component_type": {"ComponentTypeKey"},
    "component_wear": {"AsOfDate", "ComponentTypeKey", "AirframeKey", "BaseKey",
                       "AirframeStatus", "IsInFleetAsOf", "WeibullShape"},
    "dq_findings": {"FindingKey", "EntityKey", "AirframeKey"},
    "threshold_recommendation": {"AsOfDate", "MinPrecisionPct", "Population", "Failures",
                                 "AlertsRaised", "TP", "FP", "FN", "ActionableTP", "LateTP"},
}

# The two shapes a date arrives in as TEXT: ISO from SQL Server, and the
# culture format Export-Csv writes. Anything else that merely parses as a date
# ('10-1', 'Mar') is compared as the text it is.
_DATE_TEXT = re.compile(r"^\d{4}-\d{2}-\d{2}(?:[ T]\d{2}:\d{2}:\d{2}(?:\.\d+)?)?$"
                        r"|^\d{1,2}/\d{1,2}/\d{4}(?: \d{1,2}:\d{2}:\d{2}(?: [AP]M)?)?$")


def _is_date(value) -> bool:
    if isinstance(value, (datetime.date, np.datetime64)):
        return True
    return isinstance(value, str) and bool(_DATE_TEXT.match(value))


def _plain(value):
    """SQL returns Decimal, bool and '' where the CSV has float, 0/1 and blank."""
    if isinstance(value, str):
        return value if value != "" else None
    if isinstance(value, (bool, np.bool_)):
        return int(value)
    if isinstance(value, Decimal):
        return float(value)
    try:
        if pd.isna(value):
            return None
    except (TypeError, ValueError):
        pass
    return value


def _missing(value) -> bool:
    """None, NaN or NaT. pandas 3 turns a None back into NaN inside .map(), so
    testing `is None` alone let a NULL through as the four-letter string 'nan'."""
    if value is None or isinstance(value, str):
        return value is None
    try:
        return bool(pd.isna(value))
    except (TypeError, ValueError):
        return False


def _shown(value):
    """A value as a reader should see it in a FAIL line."""
    if value is None or (not isinstance(value, str) and pd.isna(value)):
        return None
    if isinstance(value, np.generic):
        return value.item()
    if isinstance(value, pd.Timestamp):
        return str(value)
    return value


def _all(values: pd.Series, types) -> bool:
    present = values.dropna()
    return len(present) > 0 and all(isinstance(v, types) and not isinstance(v, bool)
                                    for v in present)


def _column_kind(sql_values: pd.Series, csv_values: pd.Series) -> str:
    """
    Decided from BOTH sources together, so they are never normalised apart --
    except text on one side against numbers on the other, which is TEXT: a code
    like '007' that the CSV reader turned into 7 is a real difference, and
    coercing both to numbers would hide it.
    """
    numbers = (int, float, np.integer, np.floating)
    if (_all(sql_values, str) and _all(csv_values, numbers)) or \
       (_all(sql_values, numbers) and _all(csv_values, str)):
        return "text"
    present = pd.Series(pd.concat([sql_values, csv_values]).dropna().unique(), dtype=object)
    if present.empty:
        return "empty"
    as_number = pd.to_numeric(present, errors="coerce")
    if as_number.notna().all():
        return "integer" if (as_number == as_number.round()).all() else "number"
    if all(_is_date(v) for v in present):
        return "datetime"
    return "text"


def _to_datetimes(values: pd.Series) -> pd.Series:
    """Parse each DISTINCT value once: ~700 dates, not 80,000 cells."""
    text = pd.Series([None if _missing(v) else str(v) for v in values], dtype=object)
    distinct = text.dropna().unique()
    parsed = pd.to_datetime(pd.Series(distinct, dtype=object), format="mixed")
    return pd.to_datetime(text.map(dict(zip(distinct, parsed)))).dt.floor("s")


def _comparable(frame: pd.DataFrame, kinds: dict[str, str]) -> pd.DataFrame:
    """Every column in one comparable form; missing is None/NaN/NaT, never 'nan'."""
    out = pd.DataFrame(index=range(len(frame)))
    for column, kind in kinds.items():
        values = frame[column].reset_index(drop=True).astype(object).map(_plain)
        if kind in ("integer", "number"):
            out[column] = pd.to_numeric(values).astype(float)
        elif kind == "datetime":
            out[column] = _to_datetimes(values)
        else:
            out[column] = pd.Series([None if _missing(v) else str(v) for v in values], dtype=object)
    return out


def frames_agree(name: str, sql: pd.DataFrame, csv: pd.DataFrame) -> tuple[bool, str]:
    """
    Every column the extract carries must exist in SQL, arrive with the same kind
    of type, and hold the same rows with the same values.

    Rows are matched by IDENTITY -- every text, date and whole-number column,
    compared exactly as a multiset -- and only then are the decimal columns
    compared, within their tolerance. The first version lined rows up by
    sorting, so one changed key shifted every row after it and a FAIL line
    reported the wrong column and thousands of phantom differences.
    """
    unknown = [c for c in csv.columns if c not in sql.columns]
    if unknown:
        return False, f"the extract has columns the SQL form does not: {unknown}"
    missing = [c for c in sql.columns
               if c not in csv.columns and c not in NOT_EXTRACTED.get(name, set())]
    if missing:
        return False, (f"SQL returns columns the extract does not carry: {missing} -- add them "
                       "to Export_Extracts.ps1, or list them in NOT_EXTRACTED with the reason")
    columns = list(csv.columns)

    # The app is handed whichever frame load() returns, so a column that is
    # numeric in the extract must be numeric from SQL too. Decimal objects in an
    # object column made Altair draw every measure as a category in SQL mode.
    mistyped = [c for c in columns if pd.api.types.is_numeric_dtype(csv[c])
                and sql[c].notna().any() and not pd.api.types.is_numeric_dtype(sql[c])]
    if mistyped:
        return False, (f"numeric in the extract but {sql[mistyped[0]].dtype} from SQL: "
                       f"{mistyped[:4]} -- the app gets different types from each source")
    if len(sql) != len(csv):
        return False, f"SQL returned {len(sql):,} rows, the extract has {len(csv):,}"

    kinds = {c: _column_kind(sql[c].astype(object).map(_plain), csv[c].astype(object).map(_plain))
             for c in columns}
    a, b = _comparable(sql[columns], kinds), _comparable(csv[columns], kinds)
    keys = [c for c in columns if kinds[c] != "number"]
    decimals = [c for c in columns if kinds[c] == "number"]

    def identity(frame):
        return Counter(tuple(_shown(v) for v in row)
                       for row in frame[keys].itertuples(index=False, name=None))

    in_sql, in_csv = identity(a), identity(b)
    if in_sql != in_csv:
        only_sql = sorted((in_sql - in_csv).elements(), key=repr)
        only_csv = sorted((in_csv - in_sql).elements(), key=repr)
        detail = (f"{len(only_sql)} row(s) in SQL have no match in the extract, "
                  f"{len(only_csv)} in the extract none in SQL")
        if only_sql and only_csv:
            diffs = [f"{keys[i]} SQL {only_sql[0][i]!r} vs extract {only_csv[0][i]!r}"
                     for i in range(len(keys)) if only_sql[0][i] != only_csv[0][i]]
            detail += "; e.g. " + "; ".join(diffs[:3])
        return False, detail

    # Within a group of rows sharing every key, sort the decimals on their
    # tolerance grid FIRST: two values that agree within the tolerance may
    # still differ in their last raw digits, and sorting on those digits could
    # order the group differently on each side. The raw values break any tie
    # that remains.
    def ordered(frame):
        keyed = frame.copy()
        grid = []
        for column in decimals:
            keyed["~" + column] = keyed[column].round(round(-math.log10(2 * DEFAULT_TOLERANCE)))
            grid.append("~" + column)
        keyed = keyed.sort_values(keys + grid + decimals, kind="mergesort", na_position="last")
        return keyed[columns].reset_index(drop=True)

    a, b = ordered(a), ordered(b)
    for column in decimals:
        tolerance = DEFAULT_TOLERANCE
        x, y = a[column], b[column]
        bad = ~((x - y).abs().le(tolerance + 1e-9) | (x.isna() & y.isna()))
        if bad.any():
            i = int(bad.to_numpy().argmax())
            row = ", ".join(f"{k}={_shown(a.at[i, k])!r}" for k in keys[:3])
            return False, (f"{int(bad.sum()):,} value(s) in {column!r} differ by more than "
                           f"{tolerance:g}; first at {row}: SQL {_shown(x.iloc[i])!r} "
                           f"vs extract {_shown(y.iloc[i])!r}")
    return True, f"{len(a):,} rows, {len(a) * len(columns):,} values"


parity_sets = [(n, q) for n, (q, _f) in sorted(data_access.SOURCES.items()) if q is not None]
if not data_access.sql_available():
    not_run = len(parity_sets)
    print(f"NOT RUN  source parity, {not_run} datasets: {data_access.describe_source()}.")
    print("         There is no SQL Server to compare the extracts against, so these")
    print("         checks are reported as not run -- not counted as passes.")
for name, _query in ([] if not_run else parity_sets):
    try:
        from_sql = data_access.load(name)
        from_csv = data_access.load(name, force_csv=True)
    except Exception as exc:  # noqa: BLE001
        check(f"source parity: {name}", False, str(exc)[:120])
        continue
    ok, detail = frames_agree(name, from_sql, from_csv)
    check(f"source parity: {name} ({detail})" if ok else f"source parity: {name}", ok, detail)

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
print(f"  {passed} passed, {failed} failed"
      + (f", {not_run} not run (no SQL Server to compare against)" if not_run else ""))
if failed:
    print("  At least one check failed -- see FAIL above.")
    sys.exit(1)
print("  Every screen renders and produces its content."
      + (" SQL/CSV parity was NOT checked." if not_run else " SQL and the extracts agree."))
print()
