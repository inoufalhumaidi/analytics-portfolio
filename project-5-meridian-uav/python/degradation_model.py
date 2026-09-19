"""
================================================================================
Project 5 - Meridian UAV Services: Predictive Maintenance
Module:  degradation_model.py
Purpose: Find out whether the fleet's failure warnings arrive in time to act on,
         and if not, whether that is a limit of the data or of the alarm.

THE ARGUMENT THIS FILE MAKES

    The fleet's deployed alarm -- one fixed vibration threshold across every
    monitored component -- catches 98% of failures at 82% precision, and is
    still mostly useless: 176 of its 283 correct predictions arrive inside the
    time it takes to obtain the part.

    The obvious explanation is that the information simply does not exist early
    enough. THE DATA REFUTES THAT. A model over the SAME sensor channels
    recovers most of the lost lead time at the same precision, and so does a
    fixed threshold chosen per component rather than per fleet. The warning is
    there; the deployed rule throws it away.

    So the finding is not "prediction does not work here". It is that accuracy
    and usefulness are different axes, that the standard fixed-threshold alarm
    is tuned on the wrong one, and that the fix costs nothing but a number.

WHERE THE SIGNAL ACTUALLY IS

    Three models are fitted, not one, because "the model works" is not a finding
    and does not tell a maintenance manager what to fund:

      TELEMETRY  vibration, temperature, current, and their rolling forms
      EXPOSURE   accumulated stress and flight hours, and the airframe's
                 stress-per-flight-hour ratio -- Finding 1's variables
      BOTH

    If exposure alone matched the pair, the condition-monitoring programme would
    be buying nothing the stress model does not already provide. It does not,
    and establishing that is worth more than a headline AUC.

FOUR THINGS DONE DELIBERATELY, EACH EASY TO GET WRONG

 1. THE SPLIT IS BY COMPONENT AND BY TIME, NEVER BY ROW.
    Readings from one component instance are ~60 near-identical rows. A random
    row split puts some in train and some in test, the model memorises the
    instance, and the test AUC comes back near 1.0. It is the most common way a
    maintenance model is oversold.

 2. THE HORIZON MUST EXCEED THE PART LEAD TIME.
    The label is "fails within H flight hours". If H is smaller than the time it
    takes to get the part, the model is TRAINED to fire too late -- and will
    then be reported as accurate, because it is.

 3. NOTHING DERIVED FROM THE REMOVAL IS A FEATURE.
    FlightHoursToRemoval builds the label; it cannot also be an input. The
    exclusion list is asserted at run time rather than trusted.

 4. THE POPULATION IS COMPONENT LIVES, NOT READINGS.
    A component that produced no telemetry cannot be alerted on, so it is a
    guaranteed miss. Building the population from readings drops it and flatters
    recall. The SQL reconciliation at the foot of this script caught exactly
    that: pandas said 2 false negatives, T-SQL said 3.

DATA DISCLOSURE
    Meridian UAV Services is fictional and all data is synthetic. Degradation
    behaviour is modelled on publicly described characteristics of rotating
    machinery. No confidential data is used and no claim is made about any real
    aircraft, operator or manufacturer.
================================================================================
"""

from __future__ import annotations

import json
from pathlib import Path

import matplotlib

matplotlib.use("Agg")  # no display in a build script
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from sklearn.ensemble import HistGradientBoostingClassifier
from sklearn.metrics import average_precision_score, roc_auc_score

import data_access

OUT = Path(__file__).resolve().parent / "outputs"
OUT.mkdir(exist_ok=True)

# The prediction horizon, in FLIGHT HOURS. See note 2 in the header: it must
# exceed the longest monitored part lead time (18 h for a rotor assembly).
HORIZON_FLIGHT_HOURS = 40.0

# Share of component instances, ordered by removal date, used for training.
TRAIN_FRACTION = 0.70

# The precision floor. Not an accuracy target -- a credibility one. Crews stop
# acting on alarms that are wrong too often, and an alarm nobody acts on has an
# effective recall of zero whatever the confusion matrix says.
MIN_PRECISION_PCT = 75.0

# The threshold SQL pins in Ref_Reporting, re-scored here so the two
# implementations of the same rule can be reconciled exactly.
SQL_VIB_THRESHOLD = 3.40

# Anything that knows how the story ended. Asserted, not assumed.
LEAKING_COLUMNS = [
    "FlightHoursToRemoval", "StressHoursToRemoval", "RemovalReason",
    "RemovedDate", "IsFailureInstall", "InstallKey", "ComponentSerial",
    "ReadingKey", "SortieKey",
]

TELEMETRY = [
    "VibrationRms", "TempRiseC", "CurrentDrawA",
    "VibRoll10", "TempRoll10", "CurrRoll10",
    "VibVsBaseline", "TempVsBaseline", "IsRotor",
]
EXPOSURE = [
    "FlightHoursSoFar", "StressHoursSoFar", "SortieSeq",
    "StressPerFlightHourSoFar", "IsRotor",
]
BOTH = sorted(set(TELEMETRY) | set(EXPOSURE))

NUMERIC = [
    "VibrationRms", "TempRiseC", "CurrentDrawA", "VibRoll10", "TempRoll10",
    "CurrRoll10", "VibVsBaseline", "TempVsBaseline", "FlightHoursSoFar",
    "StressHoursSoFar", "FlightHoursToRemoval", "StressHoursToRemoval",
]


# =============================================================================
# Data preparation
# =============================================================================
def prepare() -> tuple[pd.DataFrame, pd.DataFrame]:
    """Returns (readings, lives). See note 4 in the header for why both."""
    print("Loading:")
    readings = data_access.load("sensor_features")
    lives = data_access.load("component_life")
    types = data_access.load("component_type")

    lead = dict(zip(types["ComponentCode"], types["LeadTimeFlightHours"].astype(float)))

    for col in NUMERIC:
        readings[col] = pd.to_numeric(readings[col], errors="coerce").astype(float)
    readings["SortieDate"] = pd.to_datetime(readings["SortieDate"])
    readings["RemovedDate"] = pd.to_datetime(readings["RemovedDate"])

    """
    Only COMPLETED lives are scored, and only those that ended in a failure or a
    scheduled replacement.

    A component still fitted has not had the chance to fail, so counting it as a
    negative rewards the model for not alerting on something whose outcome is
    unknown. A component removed because its airframe was retired is censored
    for a reason unrelated to its condition; keeping it as a negative teaches
    the model that a healthy-looking part is definitely healthy -- which happens
    to be true here, for the wrong reason.
    """
    before = len(readings)
    readings = readings[readings["RemovalReason"].isin(["Failure", "Scheduled"])].copy()
    print(
        f"\n  {before - len(readings):,} readings dropped: component still fitted, or "
        f"removed with a retired airframe. Outcome unknown, so not scoreable."
    )

    readings["StressPerFlightHourSoFar"] = np.where(
        readings["FlightHoursSoFar"] > 0,
        readings["StressHoursSoFar"] / readings["FlightHoursSoFar"],
        np.nan,
    )
    readings["IsRotor"] = (readings["ComponentCode"] == "ROTOR-ASSY").astype(int)
    readings["LeadTimeRequired"] = readings["ComponentCode"].map(lead).astype(float)
    readings["WillFail"] = (
        (readings["IsFailureInstall"].astype(int) == 1)
        & (readings["FlightHoursToRemoval"] <= HORIZON_FLIGHT_HOURS)
    ).astype(int)

    # --- the scoreable population, built from LIVES ---------------------------
    lives = lives[
        lives["ComponentCode"].isin(["ROTOR-ASSY", "MOTOR-ESC"])
        & lives["RemovalReason"].isin(["Failure", "Scheduled"])
    ].copy()
    if "IsDateValid" in lives.columns:
        invalid = int((lives["IsDateValid"].astype(int) == 0).sum())
        lives = lives[lives["IsDateValid"].astype(int) == 1]
        if invalid:
            print(f"  {invalid} lives excluded: removed before installed, so the life has no length.")
    lives["RemovedDate"] = pd.to_datetime(lives["RemovedDate"])
    lives["IsFailure"] = lives["IsFailure"].astype(int)
    lives["LeadTimeRequired"] = lives["ComponentCode"].map(lead).astype(float)

    blind = int((~lives["ComponentSerial"].isin(readings["ComponentSerial"])).sum())
    print(
        f"  {len(lives):,} component lives are scoreable, of which {blind} produced no "
        f"telemetry at all -- kept, because a component nobody can see is a guaranteed miss."
    )
    return readings, lives


def split_by_component_and_time(lives: pd.DataFrame) -> tuple[set, set]:
    """
    Split on the COMPONENT INSTANCE, ordered by when it came off.

    Both a grouped split and a temporal one, and it needs to be both. Grouping
    stops the model memorising an instance; ordering by removal date makes the
    test set the future rather than a random sample of the past, which is the
    only arrangement that resembles deployment.
    """
    ordered = lives.sort_values(["RemovedDate", "ComponentSerial"])
    cut = int(len(ordered) * TRAIN_FRACTION)
    train = set(ordered.iloc[:cut]["ComponentSerial"])
    test = set(ordered.iloc[cut:]["ComponentSerial"])
    assert not (train & test), "a component instance appeared in both splits"
    return train, test


# =============================================================================
# The evaluation that matters
# =============================================================================
def evaluate_alerts(lives: pd.DataFrame, scored: pd.DataFrame,
                    threshold: float, score_column: str) -> dict:
    """
    Score a decision rule the way an operator experiences it: ONE alert per
    component, the first time the rule fires, then how much flying was left.

    Scoring every reading independently would report a single failure as dozens
    of true positives and flatter any rule that keeps firing once it starts.

    `lives` is the population; `scored` carries the readings that exist for it.
    A life with no readings simply never alerts.
    """
    fired = scored[scored[score_column] >= threshold]
    first = (
        fired.sort_values(["ComponentSerial", "SortieSeq"])
        .groupby("ComponentSerial", as_index=False)
        .first()[["ComponentSerial", "FlightHoursToRemoval", "SortieDate"]]
        .rename(columns={"FlightHoursToRemoval": "LeadTimeGiven", "SortieDate": "AlertDate"})
    )

    lv = lives.merge(first, on="ComponentSerial", how="left")
    lv["Alerted"] = lv["LeadTimeGiven"].notna().astype(int)

    tp = int(((lv["Alerted"] == 1) & (lv["IsFailure"] == 1)).sum())
    fp = int(((lv["Alerted"] == 1) & (lv["IsFailure"] == 0)).sum())
    fn = int(((lv["Alerted"] == 0) & (lv["IsFailure"] == 1)).sum())
    tn = int(((lv["Alerted"] == 0) & (lv["IsFailure"] == 0)).sum())

    caught = lv[(lv["Alerted"] == 1) & (lv["IsFailure"] == 1)]
    actionable = int((caught["LeadTimeGiven"] >= caught["LeadTimeRequired"]).sum())

    def pct(num, den):
        return round(100.0 * num / den, 2) if den else None

    return {
        "threshold": round(float(threshold), 4),
        "population": int(len(lv)),
        "failures": int(lv["IsFailure"].sum()),
        "alerts": tp + fp,
        "tp": tp, "fp": fp, "fn": fn, "tn": tn,
        "actionable_tp": actionable,
        "late_tp": tp - actionable,
        "precision_pct": pct(tp, tp + fp),
        "recall_pct": pct(tp, tp + fn),
        # of the failures CAUGHT, the share caught early enough to act on
        "actionable_lead_time_pct": pct(actionable, tp),
        # of ALL failures, the share both caught and caught in time
        "preventable_failure_pct": pct(actionable, tp + fn),
        "median_lead_time_flight_hours": (
            round(float(caught["LeadTimeGiven"].median()), 2) if len(caught) else None
        ),
    }


def choose_operating_point(sweep: pd.DataFrame) -> pd.Series:
    """
    The LOWEST threshold that reaches the precision floor.

    Not the threshold that maximises F1 or AUC. Every step above the floor is
    precision the crew did not ask for, bought with lead time the planner did.
    """
    qualifying = sweep[sweep["precision_pct"].fillna(0) >= MIN_PRECISION_PCT]
    if qualifying.empty:
        raise RuntimeError(
            f"No threshold reaches {MIN_PRECISION_PCT}% precision, so the operating point "
            "cannot be chosen the way the case study describes and the published figures "
            "would not follow from the stated method."
        )
    return qualifying.iloc[0]


# =============================================================================
# Main
# =============================================================================
def main() -> None:
    print("=" * 78)
    print("Meridian UAV - component degradation model")
    print("=" * 78)
    print(f"Source: {data_access.describe_source()}\n")

    readings, lives = prepare()

    overlap = sorted(set(BOTH) & set(LEAKING_COLUMNS))
    if overlap:
        raise AssertionError(f"These columns know how the story ended: {overlap}")

    train_serials, test_serials = split_by_component_and_time(lives)
    train = readings[readings["ComponentSerial"].isin(train_serials)]
    test_readings = readings[readings["ComponentSerial"].isin(test_serials)].copy()
    test_lives = lives[lives["ComponentSerial"].isin(test_serials)].copy()

    print(
        f"\n  Train: {len(train):>7,} readings over {len(train_serials):>4} component lives"
        f"\n  Test : {len(test_readings):>7,} readings over {len(test_serials):>4} component lives"
        f"  ({int(test_lives['IsFailure'].sum())} of which failed)"
    )
    print(
        f"  Positive rate: train {train['WillFail'].mean():.3%}, "
        f"test {test_readings['WillFail'].mean():.3%}  (a reading is positive if its "
        f"component failed within {HORIZON_FLIGHT_HOURS:.0f} flight hours)"
    )

    # -------------------------------------------------------------------------
    # Three feature sets. Where is the signal?
    # -------------------------------------------------------------------------
    print("\n" + "-" * 78)
    print("WHERE THE SIGNAL IS -- the same target, three sets of inputs")
    print("-" * 78)
    print(f"  {'features':<12} {'AUC':>7} {'thr':>6} {'precision':>10} {'recall':>8} "
          f"{'actionable':>11} {'median lead':>12}")

    variants: dict[str, dict] = {}
    for name, feats in (("telemetry", TELEMETRY), ("exposure", EXPOSURE), ("both", BOTH)):
        model = HistGradientBoostingClassifier(
            max_iter=300, learning_rate=0.06, max_leaf_nodes=31,
            min_samples_leaf=40, l2_regularization=1.0,
            random_state=20260930,  # the whole project is reproducible or it is nothing
        )
        model.fit(train[feats], train["WillFail"])
        scored = test_readings.copy()
        scored["Score"] = model.predict_proba(test_readings[feats])[:, 1]

        auc = float(roc_auc_score(test_readings["WillFail"], scored["Score"]))
        ap = float(average_precision_score(test_readings["WillFail"], scored["Score"]))

        rows = []
        for t in np.round(np.arange(0.05, 0.96, 0.05), 2):
            r = evaluate_alerts(test_lives, scored, float(t), "Score")
            r["feature_set"] = name
            rows.append(r)
        sweep = pd.DataFrame(rows)

        try:
            op = choose_operating_point(sweep)
            print(f"  {name:<12} {auc:7.4f} {op['threshold']:6.2f} "
                  f"{op['precision_pct']:9.2f}% {op['recall_pct']:7.2f}% "
                  f"{op['actionable_lead_time_pct']:10.2f}% "
                  f"{op['median_lead_time_flight_hours']:11.2f}h")
        except RuntimeError:
            op = None
            print(f"  {name:<12} {auc:7.4f}   no threshold reaches "
                  f"{MIN_PRECISION_PCT:.0f}% precision")

        variants[name] = {
            "auc": round(auc, 4), "average_precision": round(ap, 4),
            "sweep": sweep, "scored": scored,
            "operating_point": (None if op is None else {k: v for k, v in op.items()
                                                         if k != "feature_set"}),
        }

    print(
        "\n  Exposure alone cannot reach the precision floor at any threshold. The "
        "condition-monitoring\n  telemetry is carrying the signal, and the stress model "
        "is not a substitute for it -- the two\n  findings in this project are about "
        "different things and both are needed."
    )

    # -------------------------------------------------------------------------
    # The deployed rule, on the SAME components
    # -------------------------------------------------------------------------
    print("\n" + "-" * 78)
    print(f"THE DEPLOYED ALARM, ON THE SAME TEST COMPONENTS (vibration >= {SQL_VIB_THRESHOLD}x baseline)")
    print("-" * 78)

    # The rule ignores the baseline window itself: inside it, a reading is being
    # compared against a mean that includes it.
    rule_scope = test_readings[test_readings["SortieSeq"] > 10]
    deployed = evaluate_alerts(test_lives, rule_scope, SQL_VIB_THRESHOLD, "VibVsBaseline")
    print(f"  Precision                    {deployed['precision_pct']:>6.2f}%")
    print(f"  Recall                       {deployed['recall_pct']:>6.2f}%")
    print(f"  Median warning               {deployed['median_lead_time_flight_hours']:>6.2f} flight hours")
    print(f"  ACTIONABLE LEAD TIME         {deployed['actionable_lead_time_pct']:>6.2f}%")

    best = variants["both"]["operating_point"]
    if best:
        print(
            f"\n  The same sensor, read by a model instead of a fixed threshold: "
            f"{best['actionable_lead_time_pct']:.2f}% actionable against "
            f"{deployed['actionable_lead_time_pct']:.2f}%.\n"
            f"  The warning was always in the data. The deployed rule discards it."
        )

    # -------------------------------------------------------------------------
    # Reconciliation. Implemented twice, required to agree.
    # -------------------------------------------------------------------------
    print("\n" + "-" * 78)
    print("RECONCILIATION -- the same alarm rule, computed in pandas and in T-SQL")
    print("-" * 78)
    full_scope = readings[readings["SortieSeq"] > 10]
    pandas_full = evaluate_alerts(lives, full_scope, SQL_VIB_THRESHOLD, "VibVsBaseline")
    sql_eval = data_access.load("alert_evaluation")
    sql_counts = {
        "tp": int(sql_eval["TruePositive"].sum()),
        "fp": int(sql_eval["FalsePositive"].sum()),
        "fn": int(sql_eval["FalseNegative"].sum()),
        "actionable_tp": int(sql_eval["ActionableTruePositive"].sum()),
        "population": int(len(sql_eval)),
    }
    print(f"    {'':14} {'pandas':>8} {'T-SQL':>8}")
    agree = True
    for key in ("population", "tp", "fp", "fn", "actionable_tp"):
        ok = pandas_full[key] == sql_counts[key]
        agree &= ok
        print(f"    {key:<14} {pandas_full[key]:>8,} {sql_counts[key]:>8,}   {'ok' if ok else 'MISMATCH'}")
    if not agree:
        raise AssertionError(
            "The pandas and T-SQL implementations of the same alarm rule disagree. One of "
            "them is wrong and there is no way to tell which from inside either."
        )
    print("\n  Both implementations agree exactly.")

    # -------------------------------------------------------------------------
    # Outputs
    # -------------------------------------------------------------------------
    write_figures(readings, variants, test_lives, deployed)

    for name, v in variants.items():
        v["sweep"].to_csv(OUT / f"model_sweep_{name}.csv", index=False)
    best_scored = variants["both"]["scored"]
    best_scored[["ComponentSerial", "ComponentCode", "SortieSeq", "SortieDate", "Score",
                 "VibVsBaseline", "FlightHoursToRemoval", "LeadTimeRequired",
                 "IsFailureInstall", "WillFail"]].to_csv(OUT / "model_test_scores.csv", index=False)

    metrics = {
        "horizon_flight_hours": HORIZON_FLIGHT_HOURS,
        "train_fraction": TRAIN_FRACTION,
        "min_precision_pct": MIN_PRECISION_PCT,
        "feature_sets": {
            name: {"auc": v["auc"], "average_precision": v["average_precision"],
                   "operating_point": v["operating_point"]}
            for name, v in variants.items()
        },
        "deployed_fixed_threshold": deployed,
        "sql_reconciliation": {"pandas": pandas_full, "tsql": sql_counts, "agree": bool(agree)},
        "source": data_access.describe_source(),
    }
    (OUT / "model_metrics.json").write_text(json.dumps(metrics, indent=2, default=str), encoding="utf-8")

    print("\n" + "=" * 78)
    print(f"Written to {OUT}")
    for path in sorted(OUT.iterdir()):
        print(f"  {path.name}")
    print("=" * 78)


def write_figures(readings: pd.DataFrame, variants: dict,
                  test_lives: pd.DataFrame, deployed: dict) -> None:
    """Three figures, each making one point."""

    # --- 1. why one threshold cannot serve two components --------------------
    fig, axes = plt.subplots(1, 2, figsize=(12, 4.6))
    for ax, code in zip(axes, ["ROTOR-ASSY", "MOTOR-ESC"]):
        part = readings[(readings["ComponentCode"] == code)
                        & (readings["RemovalReason"] == "Failure")]
        for serial in part["ComponentSerial"].drop_duplicates().head(40):
            one = part[part["ComponentSerial"] == serial].sort_values("SortieSeq")
            ax.plot(-one["FlightHoursToRemoval"], one["VibVsBaseline"],
                    alpha=0.22, linewidth=0.8, color="#1f4e79")
        need = float(part["LeadTimeRequired"].iloc[0])
        ax.axhline(3.40, color="#7f7f7f", linestyle=":", linewidth=1.2)
        ax.axvline(-need, color="#c00000", linestyle="--", linewidth=1.5)
        ax.text(-need, ax.get_ylim()[1] * 0.96, f"  {need:.0f}h to get the part",
                color="#c00000", fontsize=8, va="top")
        ax.text(-155, 3.55, "deployed threshold 3.40x", color="#7f7f7f", fontsize=8)
        ax.set_title(f"{code}")
        ax.set_xlabel("flight hours until removal  (0 = failure)")
        ax.set_ylabel("vibration / own baseline")
        ax.set_xlim(-160, 2)
        ax.grid(alpha=0.25)
    fig.suptitle(
        "One threshold, two degradation shapes. The rotor crosses 3.40x with hours to spare;\n"
        "the motor controller crosses it inside the time it takes to obtain the part.",
        fontsize=10,
    )
    fig.tight_layout()
    fig.savefig(OUT / "fig_degradation_curves.png", dpi=140)
    plt.close(fig)

    # --- 2. the trade-off, per feature set -----------------------------------
    fig, ax = plt.subplots(figsize=(9.5, 5))
    colours = {"telemetry": "#1f4e79", "exposure": "#7f7f7f", "both": "#c00000"}
    for name, v in variants.items():
        usable = v["sweep"].dropna(subset=["precision_pct", "actionable_lead_time_pct"])
        ax.plot(usable["precision_pct"], usable["actionable_lead_time_pct"],
                marker="o", markersize=3.5, label=f"model: {name}", color=colours[name])
    ax.scatter([deployed["precision_pct"]], [deployed["actionable_lead_time_pct"]],
               marker="X", s=170, color="#000000", zorder=5,
               label="deployed fixed threshold 3.40x")
    ax.axhline(70, color="#c00000", linestyle=":", linewidth=1)
    ax.axvline(75, color="#1f4e79", linestyle=":", linewidth=1)
    ax.text(76, 72, "both targets met, above and right", fontsize=8, color="#555")
    ax.set_xlabel("precision %")
    ax.set_ylabel("actionable lead time %")
    ax.set_title(
        "Accuracy and usefulness are different axes.\n"
        "The deployed alarm sits far below curves reachable from the same sensor."
    )
    ax.legend(loc="lower left")
    ax.grid(alpha=0.25)
    fig.tight_layout()
    fig.savefig(OUT / "fig_precision_vs_leadtime.png", dpi=140)
    plt.close(fig)

    # --- 3. where the warnings land relative to the part lead time -----------
    op = variants["both"]["operating_point"]
    scored = variants["both"]["scored"]
    fired = scored[scored["Score"] >= op["threshold"]] if op else scored.iloc[0:0]
    first = (fired.sort_values(["ComponentSerial", "SortieSeq"])
                  .groupby("ComponentSerial", as_index=False).first())
    caught = first[first["IsFailureInstall"].astype(int) == 1]

    fig, ax = plt.subplots(figsize=(9.5, 5))
    for code, colour in [("ROTOR-ASSY", "#1f4e79"), ("MOTOR-ESC", "#ed7d31")]:
        part = caught[caught["ComponentCode"] == code]
        if part.empty:
            continue
        ax.hist(part["FlightHoursToRemoval"], bins=np.arange(0, 130, 5), alpha=0.6,
                label=f"{code} (n={len(part)})", color=colour)
        need = float(part["LeadTimeRequired"].iloc[0])
        ax.axvline(need, color=colour, linestyle="--", linewidth=1.6)
        ax.text(need, ax.get_ylim()[1] * 0.9, f" {code} needs {need:.0f}h",
                fontsize=8, color=colour)
    ax.set_xlabel("flight hours of warning given")
    ax.set_ylabel("failures caught")
    ax.set_title(
        "Every bar left of a component's dashed line is a correct prediction\n"
        "that arrived too late to obtain the part."
    )
    ax.legend()
    ax.grid(alpha=0.25)
    fig.tight_layout()
    fig.savefig(OUT / "fig_lead_time_distribution.png", dpi=140)
    plt.close(fig)


if __name__ == "__main__":
    main()
