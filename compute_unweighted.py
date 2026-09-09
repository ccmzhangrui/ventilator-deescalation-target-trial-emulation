"""Compute crude (unweighted) outcome counts from the production analysable datasets.

These are the *observed* event counts in each strategy arm, before any IPCW
weighting.  They are reported in the manuscript as "28-day deaths (unweighted)"
and in the primary-outcome paragraph.  They must NOT be back-derived from the
weighted risks (weighted risk x n), which is what an earlier version of the
build script did; that produced non-integer-consistent values.

Outputs: results/unweighted_counts.json
"""

import json
from pathlib import Path

import pandas as pd
from lifelines.statistics import logrank_test

RES = Path(__file__).resolve().parent / "results"

MIMIC_OUT = "death_within_hosp_28days"          # 28-day in-hospital mortality
EICU_OUT = "died"                               # eICU: in-hospital mortality


def _counts(path, outcome_col, arm_col="strategy"):
    d = pd.read_parquet(path)
    out = {}
    for arm in ("early", "deferred"):
        sub = d[d[arm_col] == arm]
        out[arm] = {
            "n": int(len(sub)),
            "events": int(sub[outcome_col].sum()),
            "risk": float(sub[outcome_col].mean()),
        }
    rd = out["early"]["risk"] - out["deferred"]["risk"]
    rr = (out["early"]["risk"] / out["deferred"]["risk"]
          if out["deferred"]["risk"] > 0 else float("nan"))
    out["rd_unweighted"] = rd
    out["rr_unweighted"] = rr
    return out


def main():
    res = {}
    p_mimic = RES / "mimic_analyzable.parquet"
    if p_mimic.exists():
        res["mimic"] = _counts(p_mimic, MIMIC_OUT)
        # Unweighted log-rank test (time-to-death within 28 days, censored at 28 d).
        d = pd.read_parquet(p_mimic)
        e = d[d["strategy"] == "early"]
        f = d[d["strategy"] == "deferred"]
        if {"surv_time", "event"}.issubset(d.columns):
            lr = logrank_test(e["surv_time"].values, f["surv_time"].values,
                              e["event"].values, f["event"].values)
            res["mimic"]["logrank_p"] = float(lr.p_value)
    p_eicu = RES / "eicu_analyzable.parquet"
    if p_eicu.exists():
        d = pd.read_parquet(p_eicu)
        col = EICU_OUT if EICU_OUT in d.columns else MIMIC_OUT
        res["eicu"] = _counts(p_eicu, col)
    (RES / "unweighted_counts.json").write_text(
        json.dumps(res, indent=2, ensure_ascii=False), encoding="utf-8")
    print(json.dumps(res, indent=2, ensure_ascii=False))


if __name__ == "__main__":
    main()
