#!/usr/bin/env python3
"""
eICU-CRD 2.0 TTE de-escalation analysis — external validation cohort.

Replicates the MIMIC-IV design (run_analysis.py):
  - Invasive MV > 24 h, alive at 48-h landmark (from vent start)
  - Day-2 (24-48 h) FiO2 <= 50% and PEEP <= 10 cmH2O (window medians)
  - Day-2 mode classifiable (treatment labels; fallback PS/PC inference)
  - Strategy: early = mode step-down by Day 3 (48-72 h) or extubation <= 72 h
  - CCW-style stabilised IPCW (logistic PS, truncate 1-99 pct)
  - Outcome: in-hospital mortality (hospitaldischargestatus == Expired)
Outputs: results/eicu_validation.json, eicu_flow_counts.csv,
         eicu_balance.csv, eicu_table1.csv, eicu_analyzable.parquet
"""
import json, time
from pathlib import Path
import numpy as np
import pandas as pd

BASE = Path("/Users/zhangrui/Desktop/tte/TTE_Deescalation_v2")
CACHE = BASE / "eicu_cache"
RES = BASE / "results"
RES.mkdir(exist_ok=True)

MIN_H = 60          # minutes per hour
DAY2_LO, DAY2_HI = 24 * MIN_H, 48 * MIN_H
DAY3_LO, DAY3_HI = 48 * MIN_H, 72 * MIN_H
GRACE_HI = 72 * MIN_H

VENT_PROOF = {"fio2", "peep", "vent_rate", "ps", "pc", "tv_set", "pplat",
              "total_rr", "map_airway", "peakp"}
MODE_LEVEL = {"controlled": 2, "assisted": 1}


def log(msg):
    print(f"[{time.strftime('%H:%M:%S')}] {msg}", flush=True)


def wmean(w, x):
    return float(np.sum(w * x) / np.sum(w))


# ─────────────────────── Vent timeline per stay ──────────────────────
def build_vent_timeline(resp):
    """Per-stay vent start / last vent evidence / Day-2 / Day-3 settings."""
    rp = resp[resp["var"].isin(VENT_PROOF)]
    g = rp.groupby("stay_id")["offset_min"]
    tl = pd.DataFrame({"vent_start": g.min(), "last_vent": g.max()})
    return tl


def window_stats(resp, tl, lo, hi, suffix):
    """Median settings in [vent_start+lo, vent_start+hi)."""
    r = resp.merge(tl[["vent_start"]], left_on="stay_id", right_index=True)
    r = r[(r["offset_min"] >= r["vent_start"] + lo) &
          (r["offset_min"] < r["vent_start"] + hi)]
    pv = (r.groupby(["stay_id", "var"])["value"].median().unstack())
    pv.columns = [f"{c}{suffix}" for c in pv.columns]
    return pv


def mode_from_treat(treat, tl, lo, hi):
    """Most definitive mode label in window from treatment events
    (weaning counts as assisted, consistent with MIMIC SBT/SPONT)."""
    t = treat[treat["event"].isin(["controlled", "assisted", "weaning"])].copy()
    t["event"] = t["event"].replace({"weaning": "assisted"})
    t = t.merge(tl[["vent_start"]], left_on="stay_id", right_index=True)
    t = t[(t["offset_min"] >= t["vent_start"] + lo) &
          (t["offset_min"] < t["vent_start"] + hi)]
    # if both appear, take the latest event in window
    t = t.sort_values("offset_min").groupby("stay_id").last()
    return t["event"]


def mode_evidence_wide(treat, resp, tl, hi, lo=0):
    """Latest mode evidence in [vent_start+lo, vent_start+hi) — wide fallback.

    Priority per stay (latest evidence wins):
      treatment label (controlled/assisted/weaning) > PS>0 (assisted)
      > PC charted (controlled) > Vent Rate charted (controlled).
    """
    # treatment events
    t = treat[treat["event"].isin(["controlled", "assisted", "weaning"])].copy()
    t["event"] = t["event"].replace({"weaning": "assisted"})
    t = t.merge(tl[["vent_start"]], left_on="stay_id", right_index=True)
    t = t[(t["offset_min"] >= t["vent_start"] + lo) &
          (t["offset_min"] < t["vent_start"] + hi)]
    t = t.sort_values("offset_min").groupby("stay_id").last()
    mode = t["event"].copy()

    # respiratory inference
    r = resp[resp["var"].isin(["ps", "pc", "vent_rate"])].merge(
        tl[["vent_start"]], left_on="stay_id", right_index=True)
    r = r[(r["offset_min"] >= r["vent_start"] + lo) &
          (r["offset_min"] < r["vent_start"] + hi)]
    r = r[(r["var"] != "ps") | (r["value"] > 0)]
    r["mode"] = np.where(r["var"] == "ps", "assisted", "controlled")
    r = r.sort_values("offset_min").groupby("stay_id").last()
    mode = mode.reindex(
        mode.index.union(r.index)).fillna(r["mode"])
    return mode


def mode_infer(win_stats, treat_mode):
    """Combine treatment label with PS/PC inference (strict window)."""
    mode = treat_mode.copy()
    ps = win_stats.get("ps")
    pc = win_stats.get("pc")
    vr = win_stats.get("vent_rate")
    inferred = pd.Series(index=win_stats.index, dtype=object)
    if ps is not None:
        inferred[ps.notna() & (ps > 0)] = "assisted"
    if pc is not None:
        inferred[inferred.isna() & pc.notna()] = "controlled"
    if vr is not None:
        inferred[inferred.isna() & vr.notna()] = "controlled"
    mode = mode.reindex(win_stats.index).fillna(inferred)
    return mode


def niv_flag(treat):
    niv = treat[treat["event"] == "niv"].groupby("stay_id").size()
    return niv


# ───────────────────────────── Main ──────────────────────────────────
def main():
    t0 = time.time()
    log("loading caches")
    base = pd.read_parquet(CACHE / "base.parquet")
    resp = pd.read_parquet(CACHE / "resp.parquet")
    treat = pd.read_parquet(CACHE / "treat.parquet")
    labs = pd.read_parquet(CACHE / "labs.parquet")
    drugs = pd.read_parquet(CACHE / "drugs.parquet")
    hist = pd.read_parquet(CACHE / "history.parquet")

    flow = {}
    flow["first_stay_adults"] = int(len(base))

    # vent timeline
    tl = build_vent_timeline(resp)
    niv = niv_flag(treat)
    # exclude stays whose ONLY vent evidence is NIV treatment with no resp
    # settings (they would not appear in tl anyway)
    df = base.merge(tl, left_on="stay_id", right_index=True, how="inner")
    flow["with_vent_settings"] = int(len(df))

    # vent duration > 24h (last - first > 1440 min)
    df["vent_dur_min"] = df["last_vent"] - df["vent_start"]
    df = df[df["vent_dur_min"] > DAY2_LO]
    flow["vent_gt_24h"] = int(len(df))

    # alive at 48-h landmark
    df["died"] = (df["hospitaldischargestatus"] == "Expired").astype(int)
    df["unit_died"] = (df["unitdischargestatus"] == "Expired").astype(int)
    df["death_offset"] = np.where(
        df["unit_died"] == 1, df["unitdischargeoffset"], np.nan)
    df = df[~((df["unit_died"] == 1) &
              (df["death_offset"] <= df["vent_start"] + DAY2_HI))]
    flow["alive_at_48h"] = int(len(df))

    # Day-2 / Day-3 window settings
    d2 = window_stats(resp, tl, DAY2_LO, DAY2_HI, "_d2")
    d3 = window_stats(resp, tl, DAY3_LO, DAY3_HI, "_d3")
    df = df.merge(d2, left_on="stay_id", right_index=True, how="left")

    # has Day-2 settings
    df = df[df["fio2_d2"].notna() & df["peep_d2"].notna()]
    flow["has_d2_settings"] = int(len(df))

    # FiO2 <= 50, PEEP <= 10
    df = df[(df["fio2_d2"] <= 50) & (df["peep_d2"] <= 10)]
    flow["fio2_peep_ok"] = int(len(df))

    # mode classification Day-2 / Day-3
    # strict window first, then wide fallback (latest evidence up to window end)
    d2s = d2.reindex(df["stay_id"])
    d3s = d3.reindex(df["stay_id"])
    tm2 = mode_from_treat(treat, tl, DAY2_LO, DAY2_HI).reindex(df["stay_id"])
    tm3 = mode_from_treat(treat, tl, DAY3_LO, DAY3_HI).reindex(df["stay_id"])
    mode_d2 = mode_infer(d2s, tm2)
    mode_d3 = mode_infer(d3s, tm3)
    # wide fallback: latest evidence in [start, 48h) for D2 / [48h, 84h) for D3
    wide_d2 = mode_evidence_wide(treat, resp, tl, hi=DAY2_HI)
    wide_d3 = mode_evidence_wide(treat, resp, tl, lo=DAY3_LO,
                                 hi=DAY3_HI + 12 * MIN_H)
    mode_d2 = mode_d2.fillna(wide_d2.reindex(mode_d2.index))
    mode_d3 = mode_d3.fillna(wide_d3.reindex(mode_d3.index))
    df["mode_d2"] = mode_d2.values
    df["mode_d3"] = mode_d3.reindex(df["stay_id"]).values

    df = df[df["mode_d2"].isin(["controlled", "assisted"])]
    flow["mode_d2_classifiable"] = int(len(df))

    # extubation <= 72 h: last vent evidence before vent_start+4320
    df["extub_le_72h"] = (df["last_vent"] <= df["vent_start"] + GRACE_HI)
    # strict extubation: unit stay continues >= 6 h after last vent evidence
    # (excludes ICU-to-ICU transfers misread as extubation)
    df["extub_le_72h_strict"] = (
        df["extub_le_72h"] &
        (df["unitdischargeoffset"] >= df["last_vent"] + 6 * MIN_H))

    # grace-period death (landmark .. +24 h)
    df["grace_death"] = ((df["unit_died"] == 1) &
                         (df["death_offset"] > df["vent_start"] + DAY2_HI) &
                         (df["death_offset"] <= df["vent_start"] + GRACE_HI))

    # strategy
    def classify(r, extub_col="extub_le_72h"):
        if r["grace_death"]:
            return "grace_death"
        m2 = MODE_LEVEL.get(r["mode_d2"])
        m3 = MODE_LEVEL.get(r["mode_d3"]) if pd.notna(r["mode_d3"]) else None
        if m3 is not None:
            return "early" if m3 < m2 else "deferred"
        # D3 unknown
        if r[extub_col]:
            return "early"
        return "unascertainable"

    df["strategy"] = df.apply(classify, axis=1)
    df["strategy_strict"] = df.apply(
        lambda r: classify(r, "extub_le_72h_strict"), axis=1)
    flow["eligible"] = int(len(df))
    strat_counts = df["strategy"].value_counts().to_dict()
    log(f"strategy counts: {strat_counts}")
    log(f"flow: {flow}")
    log(f"mode_d2 dist: {df['mode_d2'].value_counts(dropna=False).to_dict()}")
    log(f"mode_d3 dist: {df['mode_d3'].value_counts(dropna=False).to_dict()}")
    log(f"extub<=72h: {df['extub_le_72h'].sum()}, grace_death: "
        f"{df['grace_death'].sum()}")

    # ── covariates ────────────────────────────────────────────────────
    log("building covariates")
    # labs in Day-2 window (relative to vent start)
    lb = labs.merge(df[["stay_id", "vent_start"]], on="stay_id")
    lb = lb[(lb["offset_min"] >= lb["vent_start"]) &
            (lb["offset_min"] < lb["vent_start"] + DAY2_HI)]
    lab_agg = (lb.groupby(["stay_id", "var"])["value"]
                 .agg(["median", "max"]).unstack())
    lab_flat = pd.DataFrame(index=df["stay_id"].unique())
    for v in ["lactate", "pao2", "ph", "bicarb", "creatinine", "platelets",
              "wbc", "tbil", "albumin", "hgb"]:
        if ("median", v) in lab_agg.columns:
            lab_flat[f"{v}_d2"] = lab_agg[("median", v)]
    lab_flat.index.name = "stay_id"
    df = df.merge(lab_flat, left_on="stay_id", right_index=True, how="left")

    # P/F ratio Day-2
    df["pf_d2"] = df["pao2_d2"] / (df["fio2_d2"] / 100.0)

    # drugs in Day-2 window
    dg = drugs.merge(df[["stay_id", "vent_start"]], on="stay_id")
    dg = dg[(dg["offset_min"] >= dg["vent_start"]) &
            (dg["offset_min"] < dg["vent_start"] + DAY2_HI)]
    for name in ["norepi", "vasopressin", "epinephrine", "phenylephrine",
                 "dopamine", "propofol", "fentanyl", "midazolam",
                 "dexmedetomidine"]:
        s = dg[dg["drug"] == name].groupby("stay_id").size()
        df[f"drug_{name}_d2"] = df["stay_id"].map(s).fillna(0).astype(int)
    df["any_vasopressor_d2"] = (df[["drug_norepi_d2", "drug_vasopressin_d2",
                                    "drug_epinephrine_d2",
                                    "drug_phenylephrine_d2",
                                    "drug_dopamine_d2"]].sum(axis=1) > 0
                                ).astype(int)

    # history
    df = df.merge(hist, on="stay_id", how="left")
    for k in ["copd", "chf", "diabetes", "renal", "malignancy",
              "hypertension", "cirrhosis", "cad"]:
        df[k] = df[k].fillna(False).astype(int)

    # demographics
    df["male"] = (df["gender"] == "Male").astype(int)
    df["eth_white"] = (df["ethnicity"] == "Caucasian").astype(int)
    df["eth_black"] = (df["ethnicity"] == "African American").astype(int)
    df["eth_hispanic"] = (df["ethnicity"] == "Hispanic").astype(int)
    df["adm_ed"] = (df["hospitaladmitsource"] == "Emergency Department"
                    ).astype(int)
    df["mode_d2_controlled"] = (df["mode_d2"] == "controlled").astype(int)

    df["log_apache"] = np.log1p(
        pd.to_numeric(df["apachescore"], errors="coerce")
          .mask(lambda s: s < 0))

    covariates = ["age_num", "male", "eth_white", "eth_black", "eth_hispanic",
                  "admissionweight", "adm_ed", "log_apache",
                  "fio2_d2", "peep_d2", "vent_rate_d2", "sao2_d2",
                  "mode_d2_controlled", "any_vasopressor_d2",
                  "propofol_any", "fentanyl_any",
                  "lactate_d2", "creatinine_d2", "platelets_d2", "wbc_d2",
                  "copd", "chf", "diabetes", "renal", "malignancy"]
    df["propofol_any"] = df["drug_propofol_d2"]
    df["fentanyl_any"] = df["drug_fentanyl_d2"]

    # impute median for missing covariates (keep all eligible patients)
    cov_df = df.set_index("stay_id")[covariates].copy()
    cov_df = cov_df.replace([np.inf, -np.inf], np.nan)
    for c in covariates:
        med = cov_df[c].median()
        cov_df[c] = cov_df[c].fillna(med if pd.notna(med) else 0)

    # ── analysable set + weights ─────────────────────────────────────
    from sklearn.linear_model import LogisticRegression

    def run_estimate(strat_col, n_boot=2000, seed=20260818):
        an = df[df[strat_col].isin(["early", "deferred"])].copy()
        n_early = int((an[strat_col] == "early").sum())
        n_def = int((an[strat_col] == "deferred").sum())
        log(f"[{strat_col}] analysable: {len(an):,} "
            f"(early {n_early:,} / deferred {n_def:,})")

        y = (an[strat_col] == "early").astype(int).values
        X = cov_df.loc[an["stay_id"]].values
        p_early = n_early / len(an)
        model = LogisticRegression(max_iter=1000, C=1.0, solver="lbfgs")
        model.fit(X, y)
        ps = np.clip(model.predict_proba(X)[:, 1], 0.01, 0.99)
        w = np.where(y == 1, p_early / ps, (1 - p_early) / (1 - ps))
        lo, hi = np.percentile(w, [1, 99])
        an["ps"] = ps
        an["weight"] = np.clip(w, lo, hi)

        we = an["weight"].values
        ye = y.astype(bool)
        mort_e = wmean(we[ye], an["died"].values[ye])
        mort_d = wmean(we[~ye], an["died"].values[~ye])
        rd = mort_e - mort_d
        rr = mort_e / mort_d
        icu_e = wmean(we[ye], an["unit_died"].values[ye])
        icu_d = wmean(we[~ye], an["unit_died"].values[~ye])

        rng = np.random.default_rng(seed)
        n = len(an)
        rds, rrs = [], []
        da = an["died"].values
        for b in range(n_boot):
            idx = rng.integers(0, n, n)
            yb = y[idx]
            if yb.sum() < 10 or (1 - yb).sum() < 10:
                continue
            mb = LogisticRegression(max_iter=500, C=1.0, solver="lbfgs")
            mb.fit(X[idx], yb)
            psb = np.clip(mb.predict_proba(X[idx])[:, 1], 0.01, 0.99)
            pe = yb.mean()
            wb = np.where(yb == 1, pe / psb, (1 - pe) / (1 - psb))
            lob, hib = np.percentile(wb, [1, 99])
            wb = np.clip(wb, lob, hib)
            db = da[idx]
            me = wmean(wb[yb == 1], db[yb == 1])
            md = wmean(wb[yb == 0], db[yb == 0])
            rds.append(me - md)
            rrs.append(me / md if md > 0 else np.nan)
            if b % 500 == 0:
                log(f"  boot {b}")
        rd_lo, rd_hi = np.percentile(rds, [2.5, 97.5])
        rr_lo, rr_hi = np.nanpercentile(rrs, [2.5, 97.5])

        # balance
        rows = []
        for j, c in enumerate(covariates):
            x = cov_df.loc[an["stay_id"], c].values
            smd_u = ((x[ye].mean() - x[~ye].mean()) /
                     np.sqrt((x[ye].var(ddof=1) + x[~ye].var(ddof=1)) / 2
                             + 1e-12))
            me_w = wmean(we[ye], x[ye]); md_w = wmean(we[~ye], x[~ye])
            ve_w = np.sum(we[ye] * (x[ye] - me_w) ** 2) / np.sum(we[ye])
            vd_w = np.sum(we[~ye] * (x[~ye] - md_w) ** 2) / np.sum(we[~ye])
            smd_w = (me_w - md_w) / np.sqrt((ve_w + vd_w) / 2 + 1e-12)
            rows.append({"covariate": c,
                         "smd_unweighted": round(abs(float(smd_u)), 4),
                         "smd_weighted": round(abs(float(smd_w)), 4)})
        bal = pd.DataFrame(rows)

        # table 1
        t1 = []
        for c in covariates:
            x = cov_df.loc[an["stay_id"], c]
            t1.append({"covariate": c,
                       "early_mean": round(float(x[ye].mean()), 3),
                       "deferred_mean": round(float(x[~ye].mean()), 3)})
        t1 = pd.DataFrame(t1)

        return {
            "an": an, "bal": bal, "t1": t1,
            "n": int(len(an)), "n_early": n_early, "n_def": n_def,
            "p_early": round(p_early, 4),
            "weight_mean": round(float(an["weight"].mean()), 4),
            "weight_sd": round(float(an["weight"].std()), 4),
            "weight_max": round(float(an["weight"].max()), 4),
            "mort_e": mort_e, "mort_d": mort_d, "rd": rd, "rr": rr,
            "rd_lo": float(rd_lo), "rd_hi": float(rd_hi),
            "rr_lo": float(rr_lo), "rr_hi": float(rr_hi),
            "icu_e": icu_e, "icu_d": icu_d,
            "max_smd": round(float(bal["smd_weighted"].max()), 4),
            "n_boot": len(rds),
        }

    main_est = run_estimate("strategy")
    an = main_est["an"]
    bal = main_est["bal"]
    strict_est = run_estimate("strategy_strict", n_boot=1000,
                              seed=20260819)
    log("strict-extubation SA: RD "
        f"{strict_est['rd']*100:.1f} pp "
        f"({strict_est['rd_lo']*100:.1f} to {strict_est['rd_hi']*100:.1f})")

    # ── save ──────────────────────────────────────────────────────────
    bal.to_csv(RES / "eicu_balance.csv", index=False)
    main_est["t1"].to_csv(RES / "eicu_table1.csv", index=False)
    pd.DataFrame([flow]).to_csv(RES / "eicu_flow_counts.csv", index=False)
    an.to_parquet(RES / "eicu_analyzable.parquet", index=False)

    out = {
        "database": "eICU-CRD v2.0",
        "design": "external validation, in-hospital mortality",
        "flow": flow,
        "strategy_counts": {k: int(v) for k, v in strat_counts.items()},
        "n_analyzable": main_est["n"],
        "n_early": main_est["n_early"],
        "n_deferred": main_est["n_def"],
        "weight_diag": {
            "p_early": main_est["p_early"],
            "weight_mean": main_est["weight_mean"],
            "weight_sd": main_est["weight_sd"],
            "weight_max": main_est["weight_max"],
        },
        "primary_hospital_mortality": {
            "mort_early": round(main_est["mort_e"], 4),
            "mort_deferred": round(main_est["mort_d"], 4),
            "rd": round(main_est["rd"], 4),
            "rd_ci_lo": round(main_est["rd_lo"], 4),
            "rd_ci_hi": round(main_est["rd_hi"], 4),
            "rr": round(main_est["rr"], 4),
            "rr_ci_lo": round(main_est["rr_lo"], 4),
            "rr_ci_hi": round(main_est["rr_hi"], 4),
        },
        "icu_mortality": {
            "mort_early": round(main_est["icu_e"], 4),
            "mort_deferred": round(main_est["icu_d"], 4),
        },
        "sa_strict_extubation": {
            "n": strict_est["n"],
            "n_early": strict_est["n_early"],
            "n_deferred": strict_est["n_def"],
            "mort_early": round(strict_est["mort_e"], 4),
            "mort_deferred": round(strict_est["mort_d"], 4),
            "rd": round(strict_est["rd"], 4),
            "rd_ci_lo": round(strict_est["rd_lo"], 4),
            "rd_ci_hi": round(strict_est["rd_hi"], 4),
            "rr": round(strict_est["rr"], 4),
            "rr_ci_lo": round(strict_est["rr_lo"], 4),
            "rr_ci_hi": round(strict_est["rr_hi"], 4),
        },
        "balance_max_smd_weighted": main_est["max_smd"],
        "n_boot": main_est["n_boot"],
        "runtime_sec": round(time.time() - t0, 1),
    }
    with open(RES / "eicu_validation.json", "w") as f:
        json.dump(out, f, indent=2)
    log("DONE")
    print(json.dumps(out, indent=2))


if __name__ == "__main__":
    main()
