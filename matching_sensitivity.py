#!/usr/bin/env python3
"""
Matching-specification sensitivity analysis for the TTE de-escalation study.
=============================================================================
Explores alternative propensity-model covariate specifications requested by the PI:
  S1: weight -> BMI (continuous)
  S2: BMI stratified + PaO2/FiO2 stratified (Berlin categories)  [PI's core suggestion]
  S3: S2 + reconstructed APACHE II (components available in MIMIC-IV flat data;
      missing pH/Na/K/Hct/chronic-health scored as 0 -> labelled "approximation")
  S4: S2 + measured Day-1 severity (worst GCS, max norepinephrine, lactate)
  S5: propensity overlap restriction ps in [0.10, 0.90]  ("fewer, cleaner matches")
All specs use the identical clone-censor-weight IPCW framework as the primary
analysis (run_analysis.py), stabilised weights truncated at 1st-99th pct,
bootstrap 800. Outputs: results/matching_sensitivity.json + .csv
"""

import warnings, json, sys, time
from pathlib import Path
import numpy as np
import pandas as pd
from sklearn.linear_model import LogisticRegression

warnings.filterwarnings("ignore")
sys.path.insert(0, "/Users/zhangrui/Desktop/tte/TTE_Deescalation_v2")
import run_analysis as ra

OUT_DIR = Path("/Users/zhangrui/Desktop/tte/TTE_Deescalation_v2/results")
RNG_SEED = 20260816
N_BOOT = 800

t0 = time.time()

# ─── Load & build base dataset (identical to primary analysis) ──────────
df = ra.load_excel()
eligible, flow = ra.screen_eligibility(df)
eligible = ra.classify_strategies(eligible)
eligible = ra.compute_outcomes(eligible)
# Build baseline covariates exactly as the primary analysis does
eligible, _cov_cols_primary, _cov_df_primary = ra.build_covariates(eligible)
analyzable = eligible[eligible["strategy"].isin(["early", "deferred"])].copy()
print(f"[{time.time()-t0:.0f}s] analyzable n={len(analyzable)}")

# ─── Helper ─────────────────────────────────────────────────────────────
def sf(v):
    try:
        f = float(v)
        return f if f == f else np.nan
    except (TypeError, ValueError):
        return np.nan

# ─── New covariates ─────────────────────────────────────────────────────
# BMI
analyzable["bmi"] = analyzable["weight"].apply(sf) / (analyzable["height"].apply(sf) / 100.0) ** 2
# BMI category (WHO), missing kept as own category
bmi_bins = [0, 18.5, 25, 30, np.inf]
bmi_labels = ["<18.5", "18.5-25", "25-30", ">=30"]
analyzable["bmi_cat"] = pd.cut(analyzable["bmi"], bins=bmi_bins, labels=bmi_labels)
analyzable["bmi_cat"] = analyzable["bmi_cat"].astype("object").where(analyzable["bmi"].notna(), "missing")

# Day-2 PaO2/FiO2 mean (same as primary)
pcols = [c for c in analyzable.columns if "氧合指数_PaO2/FiO2_Day2" in str(c)]
analyzable["pf_d2"] = analyzable[pcols].apply(lambda r: np.nanmean([sf(v) for v in r.values]), axis=1)
pf_bins = [0, 100, 200, 300, np.inf]
pf_labels = ["<100", "100-200", "200-300", ">=300"]
analyzable["pf_cat"] = pd.cut(analyzable["pf_d2"], bins=pf_bins, labels=pf_labels)
analyzable["pf_cat"] = analyzable["pf_cat"].astype("object").where(analyzable["pf_d2"].notna(), "missing")

# ─── Reconstructed APACHE II (approximation, Day-1 worst values) ────────
def worst_across(day1_cols):
    """Worst (min for these physiological derangements) across Day-1 timepoints."""
    def f(row):
        vals = [sf(v) for v in row.values]
        vals = [v for v in vals if not np.isnan(v)]
        return min(vals) if vals else np.nan
    return f

def apache2_score(age, gcs, temp, hr, map_, rr, pf, cr, wbc):
    """APACHE II approximation: physiological component from available data.
    pH/Na/K/Hct + chronic health not available -> scored 0 (under-estimate)."""
    pts = 0
    # Temperature
    if not np.isnan(temp):
        if temp >= 41: pts += 4
        elif temp >= 39: pts += 3
        elif temp >= 38.5: pts += 1
        elif temp >= 36: pts += 0
        elif temp >= 34: pts += 1
        elif temp >= 32: pts += 2
        elif temp >= 30: pts += 3
        else: pts += 4
    # MAP
    if not np.isnan(map_):
        if map_ >= 160: pts += 4
        elif map_ >= 130: pts += 3
        elif map_ >= 110: pts += 2
        elif map_ >= 70: pts += 0
        elif map_ >= 50: pts += 2
        else: pts += 4
    # HR
    if not np.isnan(hr):
        if hr >= 180: pts += 4
        elif hr >= 140: pts += 3
        elif hr >= 110: pts += 2
        elif hr >= 70: pts += 0
        elif hr >= 55: pts += 2
        elif hr >= 40: pts += 3
        else: pts += 4
    # RR
    if not np.isnan(rr):
        if rr >= 50: pts += 4
        elif rr >= 35: pts += 3
        elif rr >= 25: pts += 1
        elif rr >= 12: pts += 0
        elif rr >= 10: pts += 1
        elif rr >= 6: pts += 2
        else: pts += 4
    # Oxygenation (P/F approximation)
    if not np.isnan(pf):
        if pf >= 300: pts += 0
        elif pf >= 200: pts += 2
        elif pf >= 100: pts += 3
        else: pts += 4
    # Creatinine
    if not np.isnan(cr):
        if cr >= 350: pts += 4      # μmol/L (3.5 mg/dL = 309; use 350 as approx)
        elif cr >= 177: pts += 3
        elif cr >= 133: pts += 2
        elif cr >= 53: pts += 0
        else: pts += 2
    # WBC (×10^9/L)
    if not np.isnan(wbc):
        if wbc >= 40: pts += 4
        elif wbc >= 20: pts += 2
        elif wbc >= 15: pts += 1
        elif wbc >= 3: pts += 0
        elif wbc >= 1: pts += 2
        else: pts += 4
    # GCS
    if not np.isnan(gcs):
        pts += (15 - min(15, int(gcs)))
    # Age
    if not np.isnan(age):
        if age >= 75: pts += 6
        elif age >= 65: pts += 5
        elif age >= 55: pts += 3
        elif age >= 45: pts += 2
    return pts

day1_t = [c for c in analyzable.columns if "体温(℃)_Day1" in str(c)]
day1_hr = [c for c in analyzable.columns if "心率(次/min)_Day1" in str(c)]
day1_map = [c for c in analyzable.columns if "有创血压_平均压(mmHg)_Day1" in str(c)]
day1_rr = [c for c in analyzable.columns if "呼吸频率(次/min)_Day1" in str(c)]
day1_gcs = [c for c in analyzable.columns if "Glasgow评分_Day1" in str(c)]
day1_pf = [c for c in analyzable.columns if "氧合指数_PaO2/FiO2_Day1" in str(c)]
day1_ne = [c for c in analyzable.columns if "去甲肾上腺素" in str(c) and "Day1" in str(c)]
day1_lact = [c for c in analyzable.columns if "乳酸_Lactate(mmol/L)_Day1" in str(c)]

analyzable["d1_temp"] = analyzable[day1_t].apply(worst_across(day1_t), axis=1)
analyzable["d1_hr"] = analyzable[day1_hr].apply(worst_across(day1_hr), axis=1)
analyzable["d1_map"] = analyzable[day1_map].apply(worst_across(day1_map), axis=1)
analyzable["d1_rr"] = analyzable[day1_rr].apply(worst_across(day1_rr), axis=1)
analyzable["d1_gcs"] = analyzable[day1_gcs].apply(worst_across(day1_gcs), axis=1)
analyzable["d1_pf"] = analyzable[day1_pf].apply(lambda r: np.nanmean([sf(v) for v in r.values]), axis=1)
analyzable["d1_ne"] = analyzable[day1_ne].apply(lambda r: np.nanmax([sf(v) for v in r.values]), axis=1)
analyzable["d1_lact"] = analyzable[day1_lact].apply(lambda r: np.nanmax([sf(v) for v in r.values]), axis=1)

analyzable["apache2_approx"] = analyzable.apply(
    lambda r: apache2_score(r["age_num"] if "age_num" in analyzable.columns else sf(r["age"]),
                            r["d1_gcs"], r["d1_temp"], r["d1_hr"], r["d1_map"],
                            r["d1_rr"], r["d1_pf"],
                            sf(r["肌酐_Cr(μmol/L)_Day2"]) if "肌酐_Cr(μmol/L)_Day2" in analyzable.columns else np.nan,
                            sf(r["白细胞_WBC(×10^9/L)_Day2"]) if "白细胞_WBC(×10^9/L)_Day2" in analyzable.columns else np.nan),
    axis=1
)

# ─── Base covariates (primary specification, minus weight so we can swap) ─
BASE_COLS = [
    "age_num", "male",
    "sofa_total", "sofa_respiration", "sofa_cardiovascular", "sofa_cns", "sofa_renal",
    "d2_fio2", "d2_peep", "d2_tv", "d2_pip", "d2_pplat",
    "pao2_fio2_d2_mean", "d2_cr", "d2_wbc", "d2_plt", "d2_lactate",
    "sepsis3_flag", "ards_flag", "crrt_flag",
]

# ─── Specification builders ─────────────────────────────────────────────
def median_fill(an, cols):
    out = an[cols].copy()
    for c in out.columns:
        med = out[c].median()
        if pd.isna(med):
            med = 0
        out[c] = out[c].fillna(med)
    return out

def spec_M0(an):
    """Reproduce primary spec exactly (weight as continuous)."""
    cols = ["weight_num"] + BASE_COLS
    return median_fill(an, cols), cols

def spec_S1(an):
    """BMI continuous replaces weight."""
    cols = ["bmi"] + [c for c in BASE_COLS if c != "weight_num"]
    return median_fill(an, cols), cols

def spec_S2(an):
    """BMI categories + P/F categories (missing as own level)."""
    cont_cols = [c for c in BASE_COLS if c not in ("weight_num", "pao2_fio2_d2_mean")]
    X = median_fill(an, cont_cols)
    X["bmi_cat"] = an["bmi_cat"].values
    X["pf_cat"] = an["pf_cat"].values
    X = pd.get_dummies(X, columns=["bmi_cat", "pf_cat"], drop_first=False, dtype=float)
    return X, list(X.columns)

def spec_S3(an):
    """S2 + APACHE-II approximation."""
    X, _ = spec_S2(an)
    X["apache2_approx"] = an["apache2_approx"].fillna(an["apache2_approx"].median())
    return X, list(X.columns)

def spec_S4(an):
    """S2 + measured Day-1 severity (GCS, NE, lactate)."""
    X, _ = spec_S2(an)
    for c in ["d1_gcs", "d1_ne", "d1_lact"]:
        X[c] = an[c].fillna(an[c].median())
    return X, list(X.columns)

def spec_S5(an):
    """Primary spec but restrict to ps in [0.10, 0.90] (overlap trimming)."""
    return spec_M0(an)

# ─── Generic IPCW fit + effect + balance ────────────────────────────────
def run_spec(an, X, cols, overlap_trim=False, n_boot=N_BOOT):
    y = (an["strategy"] == "early").astype(int).values
    mask = np.ones(len(an), dtype=bool)
    if overlap_trim:
        # fit model first to determine ps
        m0 = LogisticRegression(max_iter=1000, penalty="l2", C=1.0, solver="lbfgs")
        m0.fit(X.values, y)
        ps0 = m0.predict_proba(X.values)[:, 1]
        mask = (ps0 >= 0.10) & (ps0 <= 0.90)

    sub = an[mask].copy()
    Xsub = X.loc[sub.index]
    ysub = (sub["strategy"] == "early").astype(int).values
    n_e = int((sub["strategy"] == "early").sum())
    n_d = int((sub["strategy"] == "deferred").sum())
    p_e = n_e / len(sub)

    model = LogisticRegression(max_iter=1000, penalty="l2", C=1.0, solver="lbfgs")
    model.fit(Xsub.values, ysub)
    ps = np.clip(model.predict_proba(Xsub.values)[:, 1], 0.01, 0.99)
    w = np.where(ysub == 1, p_e / ps, (1 - p_e) / (1 - ps))
    lo, hi = np.percentile(w, [1, 99])
    w = np.clip(w, lo, hi)
    sub["ps"] = ps
    sub["weight_trunc"] = w

    # Effects
    eff = ra.estimate_effects(sub)
    # Balance
    bal = ra.check_balance(sub, cols, Xsub)
    smd_w = bal["smd_weighted"].max() if len(bal) else np.nan
    n_smd_gt10 = int((bal["smd_weighted"] > 0.10).sum()) if len(bal) else np.nan

    # Bootstrap CIs (refit propensity within resample, same as primary)
    idx_e = np.where(sub["strategy"].values == "early")[0]
    idx_d = np.where(sub["strategy"].values == "deferred")[0]
    rng = np.random.RandomState(RNG_SEED)
    died = sub["died_28d"].values
    surv = sub["surv_time"].values
    strat = sub["strategy"].values
    rds, rrs = [], []
    for b in range(n_boot):
        bi = np.concatenate([rng.choice(idx_e, len(idx_e), replace=True),
                             rng.choice(idx_d, len(idx_d), replace=True)])
        by = (strat[bi] == "early").astype(int)
        bp = np.mean(by)
        try:
            m = LogisticRegression(max_iter=500, penalty="l2", C=1.0, solver="lbfgs")
            m.fit(Xsub.values[bi], by)
            bps = np.clip(m.predict_proba(Xsub.values[bi])[:, 1], 0.01, 0.99)
        except Exception:
            bps = np.full(len(bi), bp)
        bw = np.where(by == 1, bp / bps, (1 - bp) / (1 - bps))
        bwl, bwh = np.percentile(bw, [1, 99])
        bw = np.clip(bw, bwl, bwh)
        we = bw[by == 1]; wd = bw[by == 0]
        de = died[bi][by == 1]; dd = died[bi][by == 0]
        me = np.sum(we * de) / np.sum(we) if we.sum() > 0 else 0
        md = np.sum(wd * dd) / np.sum(wd) if wd.sum() > 0 else 0
        rds.append(me - md)
        rrs.append(me / md if md > 0 else np.nan)

    return {
        "n": int(len(sub)), "n_early": n_e, "n_deferred": n_d,
        "rd": eff["rd"], "rr": eff["rr"],
        "mort_early": eff["mort_early"], "mort_deferred": eff["mort_deferred"],
        "rd_ci_lo": round(float(np.nanpercentile(rds, 2.5)), 4),
        "rd_ci_hi": round(float(np.nanpercentile(rds, 97.5)), 4),
        "rr_ci_lo": round(float(np.nanpercentile(rrs, 2.5)), 4),
        "rr_ci_hi": round(float(np.nanpercentile(rrs, 97.5)), 4),
        "smd_max": round(float(smd_w), 4), "smd_gt_0_10": int(n_smd_gt10),
        "n_covariates": len(cols),
    }

# ─── Run all specifications ─────────────────────────────────────────────
results = {}
specs = {
    "M0_reproduce":        (spec_M0, dict()),
    "S1_bmi_continuous":   (spec_S1, dict()),
    "S2_bmi_cat_pf_cat":   (spec_S2, dict()),
    "S3_bmi_pf_cat_apache":(spec_S3, dict()),
    "S4_bmi_pf_cat_severity": (spec_S4, dict()),
    "S5_overlap_trim_010_090": (spec_S5, dict(overlap_trim=True)),
}
for name, (builder, kw) in specs.items():
    print(f"\n>>> {name} ...")
    X, cols = builder(analyzable)
    r = run_spec(analyzable, X, cols, **kw)
    results[name] = r
    print(f"    n={r['n']} (E={r['n_early']}/D={r['n_deferred']}) RD={r['rd']} "
          f"[{r['rd_ci_lo']},{r['rd_ci_hi']}] RR={r['rr']} SMDmax={r['smd_max']}")

# ─── Outputs ────────────────────────────────────────────────────────────
out = {"primary_ref": {"rd": -0.045, "rd_ci": [-0.0732, -0.0165]}, "specs": results}
with open(OUT_DIR / "matching_sensitivity.json", "w") as f:
    json.dump(out, f, indent=2, ensure_ascii=False)

rows = []
for name, r in results.items():
    rows.append({"spec": name, **r})
pd.DataFrame(rows).to_csv(OUT_DIR / "matching_sensitivity.csv", index=False)

print(f"\n[{time.time()-t0:.0f}s] Done. -> results/matching_sensitivity.json/.csv")
print(json.dumps(out, indent=2, ensure_ascii=False))
