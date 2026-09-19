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

import sys
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

_connection_state: dict[str, object] = {"checked": False, "available": False, "reason": ""}


def _try_connect():
    """Open a connection, or explain why not. Checked once per process."""
    if _connection_state["checked"]:
        return _connection_state["available"]

    _connection_state["checked"] = True
    try:
        import pyodbc
    except ImportError as exc:
        _connection_state["reason"] = f"pyodbc not installed ({exc})"
        return False

    # Newest driver first. An older driver still works; it is only the
    # connection string that differs.
    drivers = [d for d in pyodbc.drivers() if "SQL Server" in d]
    preferred = sorted(drivers, reverse=True)
    for driver in preferred:
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
            _connection_state["reason"] = f"{driver}: {exc}"
    if not drivers:
        _connection_state["reason"] = "no ODBC driver for SQL Server is installed"
    return False


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
    if not _connection_state["checked"]:
        _try_connect()
    if _connection_state["available"]:
        return f"SQL Server {SERVER}/{DATABASE} via {_connection_state['driver']}"
    return f"CSV extracts in {EXPORT_DIR} (SQL unavailable: {_connection_state['reason']})"


if __name__ == "__main__":
    print(f"Source: {describe_source()}")
    for key in SOURCES:
        try:
            load(key)
        except Exception as exc:  # noqa: BLE001
            print(f"  {key:<20} FAILED: {exc}", file=sys.stderr)
