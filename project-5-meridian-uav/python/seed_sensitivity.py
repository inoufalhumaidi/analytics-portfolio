"""
================================================================================
Project 5 - Meridian UAV Services: Predictive Maintenance
Script:  python/seed_sensitivity.py
Purpose: Show how much the model comparison moves with the random seed.

WHY THIS EXISTS

    degradation_model.py trains every feature set with one fixed seed, so the
    project is reproducible. But one seed is one draw: HistGradientBoosting
    holds out a random slice of the training readings for early stopping, and
    a different slice gives a different model. A published "92.22% actionable"
    from a single draw says nothing about how far the next draw would land.

    This retrains each feature set under ten seeds -- the published one and
    nine others -- with everything else identical, and reports the spread. The
    finding that matters is not any single figure but whether the comparison
    holds under every seed: telemetry clears the precision floor, exposure
    never does.

    It reuses degradation_model.py's own preparation, split, alert evaluation
    and operating-point choice, so there is no second copy of any of them.

RUN IT WITH
    python python/seed_sensitivity.py

DATA DISCLOSURE
    Meridian UAV Services is fictional and all data is synthetic.
================================================================================
"""

from __future__ import annotations

import sys

import numpy as np
import pandas as pd
from sklearn.ensemble import HistGradientBoostingClassifier
from sklearn.metrics import roc_auc_score

import degradation_model as dm

PUBLISHED_SEED = 20260930
SEEDS = [PUBLISHED_SEED] + [11, 23, 37, 41, 53, 67, 79, 83, 97]


def run(train: pd.DataFrame, test_readings: pd.DataFrame, test_lives: pd.DataFrame,
        feats: list[str], seed: int) -> dict:
    """The same model and the same operating-point rule as degradation_model.main(), one seed."""
    model = HistGradientBoostingClassifier(
        max_iter=300, learning_rate=0.06, max_leaf_nodes=31,
        min_samples_leaf=40, l2_regularization=1.0, random_state=seed)
    model.fit(train[feats], train["WillFail"])
    scored = test_readings.copy()
    scored["Score"] = model.predict_proba(test_readings[feats])[:, 1]
    out = {"seed": seed, "auc": float(roc_auc_score(test_readings["WillFail"], scored["Score"]))}
    sweep = pd.DataFrame([dm.evaluate_alerts(test_lives, scored, float(t), "Score")
                          for t in np.round(np.arange(0.05, 0.96, 0.05), 2)])
    try:
        op = dm.choose_operating_point(sweep)
        out.update(reaches_floor=True, precision=float(op["precision_pct"]),
                   actionable=float(op["actionable_lead_time_pct"]),
                   lead=float(op["median_lead_time_flight_hours"]))
    except RuntimeError:
        out.update(reaches_floor=False, precision=np.nan, actionable=np.nan, lead=np.nan)
    return out


def main() -> int:
    readings, lives = dm.prepare()
    train_serials, test_serials = dm.split_by_component_and_time(lives)
    train = readings[readings["ComponentSerial"].isin(train_serials)]
    test_readings = readings[readings["ComponentSerial"].isin(test_serials)].copy()
    test_lives = lives[lives["ComponentSerial"].isin(test_serials)].copy()

    print("\n" + "=" * 78)
    print(f"SEED SENSITIVITY -- each feature set retrained under {len(SEEDS)} seeds")
    print("=" * 78)
    results = {}
    for name, feats in (("telemetry", dm.TELEMETRY), ("exposure", dm.EXPOSURE), ("both", dm.BOTH)):
        rows = pd.DataFrame([run(train, test_readings, test_lives, feats, s) for s in SEEDS])
        results[name] = rows
        reached = int(rows["reaches_floor"].sum())
        line = (f"  {name:<10} AUC {rows['auc'].min():.4f}-{rows['auc'].max():.4f}   "
                f"reaches the {dm.MIN_PRECISION_PCT:.0f}% precision floor under {reached} of {len(SEEDS)} seeds")
        if reached:
            ok = rows[rows["reaches_floor"]]
            line += (f"\n  {'':<10} actionable {ok['actionable'].min():.2f}%-{ok['actionable'].max():.2f}% "
                     f"(median {ok['actionable'].median():.2f}%), precision {ok['precision'].min():.2f}%-"
                     f"{ok['precision'].max():.2f}%, median warning {ok['lead'].min():.2f}-{ok['lead'].max():.2f} h")
        print(line)

    # The published seed must reproduce the published figures, or this study
    # is measuring something other than what the case study reports.
    pub = results["telemetry"].iloc[0]
    failures = []
    if not (pub["reaches_floor"] and round(pub["actionable"], 2) == 92.22 and round(pub["precision"], 2) == 75.63):
        failures.append(f"published seed gives telemetry {pub['actionable']:.2f}% actionable at "
                        f"{pub['precision']:.2f}% precision, not the published 92.22% at 75.63%")
    if results["exposure"]["reaches_floor"].any():
        failures.append("exposure reaches the precision floor under some seed: "
                        "the case study's 'never reaches the floor' does not hold")
    if not results["telemetry"]["reaches_floor"].all():
        failures.append("telemetry misses the precision floor under some seed")

    print("-" * 78)
    if failures:
        for f in failures:
            print("FAIL ", f)
        return 1
    print("The comparison holds under every seed: telemetry always clears the precision floor,")
    print("exposure never does. The published seed reproduces the published figures.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
