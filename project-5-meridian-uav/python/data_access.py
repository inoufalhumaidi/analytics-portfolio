"""
================================================================================
Project 5 - Meridian UAV Services: Predictive Maintenance
Module:  data_access.py
Purpose: One place that knows where the data comes from.

SQL FIRST, CSV FALLBACK -- AND THE FALLBACK IS ANNOUNCED

    The model reads SQL Server when it can, because the feature definitions
    (rolling windows, baselines, exposure) live in vw_SensorFeatures and should
    not be reimplemented in pandas where the two can silently drift apart.

    When SQL Server is not available it falls back to data_exports/, which was
    written from those same views. That keeps the analysis reproducible for a
    reviewer with no database.

    The fallback PRINTS that it happened. A silent fallback is how an analysis
    comes to be run against a stale extract for three weeks without anyone
    noticing -- the code works, the numbers appear, and nothing says the source
    changed.

DATA DISCLOSURE
    Meridian UAV Services is fictional and all data is synthetic. No
    confidential data is used and no claim is made about any production system.
================================================================================
"""

from __future__ import annotations

import re
import sys
from decimal import Decimal
from pathlib import Path

import pandas as pd

PROJECT_ROOT = Path(__file__).resolve().parent.parent
EXPORT_DIR = PROJECT_ROOT / "data_exports"

SERVER = r"localhost\TEW_SQLEXPRESS"
DATABASE = "MeridianUAV"

# Every dataset this project reads, with the SQL that defines it and the file
# that stands in for it. Keeping the pair together is what stops the two
# drifting: a query changed without its extract is visible here.
SOURCES: dict[str, tuple[str, str]] = {
    "sensor_features": (
        "SELECT * FROM dbo.vw_SensorFeatures",
        "sensor_features.csv.gz",
    ),
    "component_life": (
        "SELECT * FROM dbo.vw_ComponentLifeHistory",
        "component_life_history.csv",
    ),
    "component_wear": (
        "SELECT * FROM dbo.fn_ComponentWear('2026-09-30') WHERE AirframeStatus = 'Active'",
        "component_wear.csv",
    ),
    "component_type": (
        "SELECT * FROM dbo.Dim_ComponentType",
        "dim_component_type.csv",
    ),
    "service_interval": (
        "SELECT * FROM dbo.Ref_ServiceInterval",
        "ref_service_interval.csv",
    ),
    "threshold_sweep": (
        None,  # SQL form needs a loop; the extract is the canonical copy
        "threshold_sweep.csv",
    ),
    "alert_evaluation": (
        "SELECT * FROM dbo.fn_AlertEvaluation('2026-09-30', 3.40)",
        "alert_evaluation.csv",
    ),
    # BOTH boards, matching what Export_Extracts.ps1 writes. The SQL form
    # originally called fn_FleetKPI alone, so this dataset returned 5 rows from
    # SQL and 8 from the extract -- the same name resolving to two different
    # answers depending on whether a database happened to be reachable. The
    # three missing rows were the prediction metrics, which is to say the second
    # finding disappeared whenever the app could reach SQL.
    "fleet_kpi": (
        """SELECT MetricName, MetricValue, Numerator, Denominator, TargetValue, WarningValue,
                  Direction, Unit, RAGStatus, [Description] FROM dbo.fn_FleetKPI('2026-09-30')
           UNION ALL
           SELECT MetricName, MetricValue, Numerator, Denominator, TargetValue, WarningValue,
                  Direction, Unit, RAGStatus, [Description] FROM dbo.fn_PredictionKPI('2026-09-30', 3.400)""",
        "fleet_kpi.csv",
    ),
    "threshold_recommendation": (
        "SELECT * FROM dbo.fn_ThresholdRecommendation('2026-09-30', 75.00)",
        "threshold_recommendation.csv",
    ),
    "base_scorecard": (
        "SELECT * FROM dbo.fn_BaseScorecard('2026-09-30')",
        "base_scorecard.csv",
    ),
    "action_queue": (
        "SELECT * FROM dbo.fn_ActionQueue('2026-09-30', 120.0)",
        "action_queue.csv",
    ),
    "reporting": (
        "SELECT * FROM dbo.Ref_Reporting",
        "ref_reporting.csv",
    ),
    "dq_findings": (
        "SELECT * FROM dbo.DQ_Findings",
        "dq_findings.csv",
    ),
}

# "reason" is the full technical cause, for the log. "summary" is the same fact
# in words a viewer of the app can read -- see describe_source().
_connection_state: dict[str, object] = {
    "checked": False, "available": False, "reason": "", "summary": "",
}


def _try_connect():
    """Open a connection, or explain why not. Checked once per process."""
    if _connection_state["checked"]:
        return _connection_state["available"]

    _connection_state["checked"] = True
    available = _probe()
    if not available:
        # The one place the FULL technical reason is printed. load() repeats a
        # short form per dataset and describe_source() gives it in plain words.
        print(f"  SQL Server not used: {_connection_state['reason']}")
    return available


def _probe():
    try:
        import pyodbc
    except ImportError as exc:
        # ImportError covers two different causes: the package is absent, or it
        # is present but its native ODBC library is not (libodbc.so.2 on Linux,
        # which the pyodbc wheel does not bundle). "Not installed" was only ever
        # true of the first, so the log says what is actually known.
        _connection_state["reason"] = f"pyodbc could not be imported ({exc})"
        _connection_state["summary"] = "no SQL Server client in this environment"
        return False

    # Newest driver first: "ODBC Driver N for SQL Server" by N descending, then
    # Native Client, then the legacy "SQL Server" driver. This comment used to
    # claim that order over a plain reverse-alphabetical sort, which picked the
    # legacy driver ahead of Driver 18 -- and the legacy driver returns DATE
    # columns as text where the current ones return dates.
    def rank(name: str):
        match = re.fullmatch(r"ODBC Driver (\d+) for SQL Server", name)
        if match:
            return (0, -int(match.group(1)), name)
        return (1 if "Native Client" in name else 2, 0, name)

    drivers = [d for d in pyodbc.drivers() if "SQL Server" in d]
    failures = []
    for driver in sorted(drivers, key=rank):
        try:
            conn = pyodbc.connect(
                f"DRIVER={{{driver}}};SERVER={SERVER};DATABASE={DATABASE};"
                "Trusted_Connection=yes;TrustServerCertificate=yes",
                timeout=5,
            )
            conn.close()
            _connection_state["available"] = True
            _connection_state["driver"] = driver
            return True
        except Exception as exc:  # noqa: BLE001 - any driver failure is the same to us
            failures.append(f"{driver}: {exc}")
            _connection_state["reason"] = " | ".join(failures)   # every driver's, not the last
            # Neutral on purpose: this branch also catches a server that answered
            # and refused the login or the database (e.g. sql/01 not yet run), so
            # "not reachable" would send the reader to debug the network.
            _connection_state["summary"] = f"could not connect to SQL Server {SERVER}/{DATABASE}"
    if not drivers:
        _connection_state["reason"] = "no ODBC driver for SQL Server is installed"
        _connection_state["summary"] = "no SQL Server client in this environment"
    return False


def sql_available() -> bool:
    """True when SQL Server answered. Probed once per process, then cached."""
    return bool(_try_connect())


def load(name: str, *, force_csv: bool = False) -> pd.DataFrame:
    """
    Load one dataset by name.

    Tries SQL Server, falls back to the CSV extract, and says which it used.
    """
    if name not in SOURCES:
        raise KeyError(
            f"Unknown dataset {name!r}. Known: {', '.join(sorted(SOURCES))}. "
            "Adding one means adding BOTH its query and its extract file, so the "
            "two cannot drift apart."
        )

    query, filename = SOURCES[name]

    if query is not None and not force_csv and _try_connect():
        import pyodbc

        driver = _connection_state["driver"]
        conn = pyodbc.connect(
            f"DRIVER={{{driver}}};SERVER={SERVER};DATABASE={DATABASE};"
            "Trusted_Connection=yes;TrustServerCertificate=yes"
        )
        try:
            # Built from a cursor rather than pd.read_sql, which warns that a
            # raw DBAPI2 connection is untested and pushes towards SQLAlchemy.
            # Nothing here needs an ORM, and a dependency added to silence a
            # warning is a dependency nobody can later justify.
            cursor = conn.cursor()
            cursor.execute(query)
            columns = [c[0] for c in cursor.description]
            rows = [tuple(r) for r in cursor.fetchall()]
        finally:
            conn.close()
        frame = pd.DataFrame.from_records(rows, columns=columns)
        # SQL DECIMAL columns arrive as decimal.Decimal objects in an object
        # column, while the extract gives float64. Handed to the app as they
        # were, Altair could not infer a type for them and drew every measure on
        # a CATEGORICAL axis in SQL mode -- bar heights stopped encoding values,
        # and only on a machine with a database. Converting here gives the app
        # the same types from either source; validate_app.py asserts it.
        for column in frame.columns:
            present = frame[column].dropna()
            if len(present) and all(isinstance(v, Decimal) for v in present):
                frame[column] = frame[column].astype(float)
        print(f"  {name:<20} {len(frame):>7,} rows   from SQL Server")
        return frame

    path = EXPORT_DIR / filename
    if not path.exists():
        raise FileNotFoundError(
            f"{path} is missing and SQL Server is unavailable "
            f"({_connection_state['reason']}). Run data_exports/Export_Extracts.ps1."
        )

    frame = pd.read_csv(path)  # pandas decompresses .gz by extension

    # Say WHICH of the three reasons applies. The first version reported
    # "SQL unavailable" for the one dataset that has no SQL form at all,
    # quoting a stale error from a driver probe that had already succeeded
    # on a different driver -- a message that was wrong twice over.
    if query is None:
        reason = "no SQL form; the extract is the canonical copy"
    elif force_csv:
        reason = "extract requested explicitly"
    else:
        reason = f"SQL unavailable: {_connection_state['reason'][:60]}"
    print(f"  {name:<20} {len(frame):>7,} rows   from {filename}  ({reason})")
    return frame


def describe_source() -> str:
    """
    One line naming the source, for the app's sidebar and the model's report.

    It names the source by its path WITHIN THE REPOSITORY and gives the reason
    in plain words. It used to print the absolute path and the raw exception,
    and the app shows this line to every viewer: on Streamlit Community Cloud
    that read "CSV extracts in /mount/src/analytics-portfolio/... (SQL
    unavailable: pyodbc not installed (No module named 'pyodbc'))" -- a server
    mount path and a Python error, which reads as a malfunction rather than a
    disclosure. The fallback is still announced, and the full technical reason
    is printed once, when the fallback is decided, for whoever reads the log.
    """
    if not _connection_state["checked"]:
        _try_connect()
    if _connection_state["available"]:
        return f"SQL Server {SERVER}/{DATABASE} via {_connection_state['driver']}"
    where = EXPORT_DIR.relative_to(PROJECT_ROOT.parent).as_posix()
    return f"CSV extracts in {where} ({_connection_state['summary']})"


if __name__ == "__main__":
    print(f"Source: {describe_source()}")
    for key in SOURCES:
        try:
            load(key)
        except Exception as exc:  # noqa: BLE001
            print(f"  {key:<20} FAILED: {exc}", file=sys.stderr)
