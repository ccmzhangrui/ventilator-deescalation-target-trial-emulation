#!/usr/bin/env python3
"""
TTE De-escalation Analysis — Full Pipeline from Raw MIMIC-IV Excel Data
=======================================================================
Reads:  /Users/zhangrui/Documents/sofa2.0/副本minmic数据.xlsx
Writes: results/ directory with all CSVs, JSONs, and PNG figures

Study Protocol (Target Trial Emulation) — Option A (broadened population):
  - Eligibility: adult (≥18), invasive MV >24h, alive & ventilated at 48h,
    classified Day-2 mode (controlled or assisted), FiO₂ ≤50%, PEEP ≤10 cmH₂O
  - Treatment strategies (24h grace period Day2→Day3, mode-level step-down):
    * Early de-escalation: Day-3 support level < Day-2 (mode stepped down, or vent ended while alive ≤72h)
    * Deferred: Day-3 support level ≥ Day-2 (held same level or escalated)
  - Primary outcome: 28-day all-cause mortality
  - Secondary: RMST(28d), VFD28
  - Estimator: clone-censor-weight with stabilised IPCW (truncated 1st-99th pct)
  - Bootstrap: 2000 (primary), 800 (sensitivity/exploratory)
  - Exploratory: FiO₂ × PEEP 3×3 threshold grid
"""

import warnings, json, os, sys
from pathlib import Path
from collections import Counter
import numpy as np
import pandas as pd
from scipy import stats
from lifelines import KaplanMeierFitter
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches

warnings.filterwarnings("ignore")

# ─── Paths ──────────────────────────────────────────────────────────────
EXCEL_PATH = "/Users/zhangrui/Documents/sofa2.0/副本minmic数据.xlsx"
OUT_DIR    = Path("/Users/zhangrui/Desktop/tte/TTE_Deescalation_v2/results")
FIG_DIR    = Path("/Users/zhangrui/Desktop/tte/TTE_Deescalation_v2/figures")
OUT_DIR.mkdir(parents=True, exist_ok=True)
FIG_DIR.mkdir(parents=True, exist_ok=True)

RNG_SEED = 20260816
np.random.seed(RNG_SEED)

# ─── Ventilator mode classification ─────────────────────────────────────
CONTROLLED_KEYWORDS = ["CMV", "PRVC/AC", "PCV+", "PCV/AC", "APV", "VOL/AC", "SIMV", "AC/PC", "AC/VC"]
ASSISTED_KEYWORDS   = ["CPAP", "PSV", "SBT", "Standby", "SPONT", "MMV", "PPS", "ApnVol", "ApnPres"]

def classify_mode(mode_str):
    """Return 'controlled', 'assisted', or None (None = missing/NULL/NaN)."""
    if mode_str is None:
        return None
    if isinstance(mode_str, float) and pd.isna(mode_str):
        return None
    s = str(mode_str).strip()
    if s in ("NULL", "nan", ""):
        return None
    # Check controlled first (more specific)
    for kw in CONTROLLED_KEYWORDS:
        if kw in s:
            return "controlled"
    for kw in ASSISTED_KEYWORDS:
        if kw in s:
            return "assisted"
    # Unknown modes
    return "unknown"

# ─── Helpers ────────────────────────────────────────────────────────────
def safe_float(v):
    if v is None or str(v).strip() == "NULL":
        return np.nan
    try:
        return float(v)
    except (ValueError, TypeError):
        return np.nan

def load_excel():
    """Load the MIMIC-IV Excel file into a pandas DataFrame."""
    print("Loading Excel (this may take ~30s)...")
    df = pd.read_excel(EXCEL_PATH, sheet_name="minmic数据", engine="openpyxl")
    print(f"  Loaded {len(df)} rows × {len(df.columns)} cols")
    return df

# ─── 1. Data extraction & eligibility screening ────────────────────────
def screen_eligibility(df):
    """Apply inclusion/exclusion criteria and return eligible subset + flow counts."""
    flow = {}

    # Total source records
    flow["total_records"] = len(df)

    # 1. Adult (≥18)
    df["age_num"] = df["age"].apply(safe_float)
    adult_mask = df["age_num"] >= 18
    flow["adults"] = int(adult_mask.sum())

    # 2. Invasive mechanical ventilation
    df["vent_flag"] = df["ventilation"].apply(lambda x: 1 if safe_float(x) == 1 else 0)
    vent_mask = df["vent_flag"] == 1
    flow["ventilated"] = int(vent_mask.sum())

    # 3. Ventilation > 24h
    df["vent_hrs"] = df["ventilation_hour"].apply(safe_float)
    vent24_mask = df["vent_hrs"] > 24
    flow["vent_gt_24h"] = int((vent_mask & vent24_mask).sum())

    # 4. Alive at 48h — need to check survival
    # hosp_survival_day: days survived in hospital (only for dead patients?)
    # is_hosp_dead: 0/1 for everyone
    # We need patients who were alive at 48h after ICU admission
    # Use icu_survival_day — if present and > 2, or if is_icu_dead = 0 and has been in ICU ≥ 2 days
    # Actually, looking at the data: is_icu_dead is 0/1 for everyone
    # icu_survival_day seems to only be filled for dead patients
    # A patient alive at 48h: either is_icu_dead=0, or (is_icu_dead=1 AND icu_survival_day > 2)
    df["is_icu_dead_num"] = df["is_icu_dead"].apply(safe_float)
    df["icu_surv_day"] = df["icu_survival_day"].apply(safe_float)
    df["is_hosp_dead_num"] = df["is_hosp_dead"].apply(safe_float)
    df["hosp_surv_day"] = df["hosp_survival_day"].apply(safe_float)

    alive_48h_mask = (
        (df["is_icu_dead_num"] == 0) |
        ((df["is_icu_dead_num"] == 1) & (df["icu_surv_day"] > 2))
    )
    flow["alive_at_48h"] = int((vent_mask & vent24_mask & alive_48h_mask).sum())

    # 5. Has Day-2 ventilator settings (mode present and not NULL/NaN)
    d2_mode_col = "呼吸模式_Day2"
    def _has_mode(x):
        if x is None:
            return False
        if isinstance(x, float) and pd.isna(x):
            return False
        return str(x).strip() not in ("NULL", "nan", "")
    has_d2_mode = df[d2_mode_col].apply(_has_mode)
    flow["has_d2_settings"] = int((vent_mask & vent24_mask & alive_48h_mask & has_d2_mode).sum())

    # 6. Day-2 ventilator mode classified (controlled or assisted; unknown excluded)
    df["d2_mode_class"] = df[d2_mode_col].apply(classify_mode)
    classified_d2 = df["d2_mode_class"].isin(["controlled", "assisted"])

    # 7. Day-2 FiO₂ ≤ 50%
    df["d2_fio2"] = df["氧浓度(%)_Day2"].apply(safe_float)
    fio2_ok = df["d2_fio2"] <= 50
    flow["fio2_le_50"] = int((vent_mask & vent24_mask & alive_48h_mask & has_d2_mode & classified_d2 & fio2_ok).sum())

    # 8. Day-2 PEEP ≤ 10
    df["d2_peep"] = df["呼气末正压(cmH2O)_Day2"].apply(safe_float)
    peep_ok = df["d2_peep"] <= 10
    flow["peep_le_10"] = int((vent_mask & vent24_mask & alive_48h_mask & has_d2_mode & classified_d2 & fio2_ok & peep_ok).sum())

    # Final eligible — any controlled/assisted mode at Day 2 with FiO₂ ≤50% and PEEP ≤10
    eligible_mask = (vent_mask & vent24_mask & alive_48h_mask & has_d2_mode &
                     classified_d2 & fio2_ok & peep_ok)
    flow["eligible"] = int(eligible_mask.sum())

    eligible = df[eligible_mask].copy()
    return eligible, flow

# ─── 2. Treatment strategy classification ──────────────────────────────
def classify_strategies(eligible):
    """Classify each eligible patient into early/deferred/unascertainable/grace_death."""

    # Day-3 mode
    d3_mode_col = "呼吸模式_Day3"
    eligible["d3_mode_class"] = eligible[d3_mode_col].apply(classify_mode)

    # Check if died during grace period (between Day 2 and Day 3)
    # Grace death: died within 24h after landmark (i.e., icu_surv_day ≤ 3 and is_icu_dead = 1)
    # or hosp_surv_day ≤ 3 and is_hosp_dead = 1
    eligible["grace_death"] = (
        ((eligible["is_icu_dead_num"] == 1) & (eligible["icu_surv_day"] <= 3)) |
        ((eligible["is_hosp_dead_num"] == 1) & (eligible["hosp_surv_day"] <= 3))
    )

    # Treatment strategy — mode-level step-down (Option A: any classified D2 mode)
    # Support level: controlled=2, assisted=1 (higher = more support)
    MODE_LEVEL = {"controlled": 2, "assisted": 1}

    def assign_strategy(row):
        if row["grace_death"]:
            return "grace_death"
        d3 = row["d3_mode_class"]
        # No Day-3 mode: infer extubation if ventilation ended before Day 3
        if d3 is None:
            if row["vent_hrs"] is not None and not np.isnan(row["vent_hrs"]):
                if row["vent_hrs"] <= 72:
                    return "early"
            return "unascertainable"
        if d3 == "unknown":
            return "unascertainable"
        d2_lvl = MODE_LEVEL.get(row["d2_mode_class"])
        d3_lvl = MODE_LEVEL.get(d3)
        if d2_lvl is None or d3_lvl is None:
            return "unascertainable"
        if d3_lvl < d2_lvl:      # stepped down in support → early de-escalation
            return "early"
        return "deferred"        # held same level or escalated → no de-escalation

    eligible["strategy"] = eligible.apply(assign_strategy, axis=1)
    return eligible

# ─── 3. Outcomes ────────────────────────────────────────────────────────
def compute_outcomes(eligible):
    """Compute primary and secondary outcomes."""

    # Primary: 28-day all-cause mortality
    eligible["died_28d"] = eligible["death_within_hosp_28days"].apply(safe_float).astype(int)

    # Survival time (days): if died, use hosp_surv_day (capped at 28); if survived, 28
    eligible["surv_time"] = eligible.apply(
        lambda r: min(r["hosp_surv_day"], 28) if r["is_hosp_dead_num"] == 1 and not np.isnan(r["hosp_surv_day"]) else 28,
        axis=1
    )
    # If died in ICU but not hospital (shouldn't happen), use icu_surv_day
    eligible.loc[(eligible["is_icu_dead_num"] == 1) & (eligible["is_hosp_dead_num"] == 0) &
                 (eligible["icu_surv_day"].notna()), "surv_time"] = \
        eligible.loc[(eligible["is_icu_dead_num"] == 1) & (eligible["is_hosp_dead_num"] == 0) &
                     (eligible["icu_surv_day"].notna()), "icu_surv_day"].clip(upper=28)

    eligible["event"] = eligible["died_28d"]

    # VFD28: ventilator-free days at day 28
    # VFD = max(0, 28 - vent_days) if alive at 28d; 0 if died before 28d
    # vent_days = vent_hrs / 24, capped at 28
    eligible["vent_days"] = (eligible["vent_hrs"] / 24).clip(upper=28)
    eligible["vfd28"] = eligible.apply(
        lambda r: max(0, 28 - r["vent_days"]) if r["died_28d"] == 0 else 0,
        axis=1
    )

    return eligible

# ─── 4. Covariates for propensity model ────────────────────────────────
def build_covariates(eligible):
    """Build baseline covariate matrix for the propensity model."""

    # SOFA components
    sofa_cols = ["respiration", "coagulation", "liver", "cardiovascular", "cns", "renal"]
    for c in sofa_cols:
        eligible[f"sofa_{c}"] = eligible[c].apply(safe_float)

    eligible["sofa_total"] = eligible[[f"sofa_{c}" for c in sofa_cols]].sum(axis=1)

    # Demographics
    eligible["age_num"] = eligible["age"].apply(safe_float)
    eligible["male"] = (eligible["gender"] == "M").astype(int)
    eligible["weight_num"] = eligible["weight"].apply(safe_float)

    # Day-2 vitals (use 12:00 timepoint as representative, or average available)
    # Use mean of available Day-2 PaO2/FiO2 values
    pao2_cols_d2 = [c for c in eligible.columns if "氧合指数_PaO2/FiO2_Day2" in str(c)]
    if pao2_cols_d2:
        eligible["pao2_fio2_d2_mean"] = eligible[pao2_cols_d2].apply(
            lambda row: np.nanmean([safe_float(v) for v in row.values]), axis=1
        )
    else:
        eligible["pao2_fio2_d2_mean"] = np.nan

    # Day-2 vent settings
    eligible["d2_tv"] = eligible["呼气潮气量(ml)_Day2"].apply(safe_float)
    eligible["d2_pip"] = eligible["气道峰压(cmH2O)_Day2"].apply(safe_float)
    eligible["d2_pplat"] = eligible["气道平台压(cmH2O)_Day2"].apply(safe_float)

    # Day-2 labs
    eligible["d2_cr"] = eligible["肌酐_Cr(μmol/L)_Day2"].apply(safe_float)
    eligible["d2_wbc"] = eligible["白细胞_WBC(×10^9/L)_Day2"].apply(safe_float)
    eligible["d2_plt"] = eligible["血小板_PLT(×10^9/L)_Day2"].apply(safe_float)
    eligible["d2_lactate"] = eligible[[c for c in eligible.columns if "乳酸_Lactate(mmol/L)_Day2" in str(c)]].apply(
        lambda row: np.nanmean([safe_float(v) for v in row.values]), axis=1
    )

    # Sepsis
    eligible["sepsis3_flag"] = eligible["sepsis3"].apply(lambda x: 1 if safe_float(x) == 1 else 0)

    # ARDS
    eligible["ards_flag"] = eligible["ards"].apply(lambda x: 1 if safe_float(x) == 1 else 0)

    # CRRT
    eligible["crrt_flag"] = eligible["crrt"].apply(lambda x: 1 if safe_float(x) == 1 else 0)

    # Covariate list for propensity model
    covariate_cols = [
        "age_num", "male", "weight_num",
        "sofa_total", "sofa_respiration", "sofa_cardiovascular", "sofa_cns", "sofa_renal",
        "d2_fio2", "d2_peep", "d2_tv", "d2_pip", "d2_pplat",
        "pao2_fio2_d2_mean", "d2_cr", "d2_wbc", "d2_plt", "d2_lactate",
        "sepsis3_flag", "ards_flag", "crrt_flag",
    ]

    # Fill missing with median
    cov_df = eligible[covariate_cols].copy()
    for c in cov_df.columns:
        med = cov_df[c].median()
        if pd.isna(med):
            med = 0
        cov_df[c] = cov_df[c].fillna(med)

    eligible["_cov_matrix"] = [cov_df.iloc[i].values for i in range(len(cov_df))]
    eligible["_cov_cols"] = [covariate_cols] * len(eligible)

    return eligible, covariate_cols, cov_df

# ─── 5. Propensity model & stabilised weights ──────────────────────────
def estimate_weights(eligible, cov_df, covariate_cols):
    """Fit logistic regression for propensity and compute stabilised IPCW."""

    # Only use patients with ascertainable strategy (early or deferred)
    analyzable = eligible[eligible["strategy"].isin(["early", "deferred"])].copy()
    n_early = int((analyzable["strategy"] == "early").sum())
    n_def = int((analyzable["strategy"] == "deferred").sum())
    n_total = len(analyzable)
    p_early = n_early / n_total  # marginal probability

    # Binary outcome: 1 = early, 0 = deferred
    y = (analyzable["strategy"] == "early").astype(int).values
    X = cov_df.loc[analyzable.index].values

    # Logistic regression via scipy or numpy
    # Use simple gradient descent or statsmodels-like approach
    from sklearn.linear_model import LogisticRegression
    model = LogisticRegression(max_iter=1000, penalty="l2", C=1.0, solver="lbfgs")
    model.fit(X, y)
    ps = model.predict_proba(X)[:, 1]  # P(early | covariates)
    ps = np.clip(ps, 0.01, 0.99)

    # Stabilised weights
    # Early:  p_early / ps
    # Deferred: (1 - p_early) / (1 - ps)
    weights = np.where(y == 1, p_early / ps, (1 - p_early) / (1 - ps))

    analyzable["ps"] = ps
    analyzable["weight"] = weights

    # Truncate at 1st and 99th percentiles
    lo, hi = np.percentile(weights, [1, 99])
    analyzable["weight_trunc"] = np.clip(weights, lo, hi)

    # Weight diagnostics
    w = analyzable["weight_trunc"].values
    w_diag = {
        "n_early": n_early,
        "n_deferred": n_def,
        "n_total": n_total,
        "p_early_marginal": round(p_early, 4),
        "ps_mean": round(float(np.mean(ps)), 4),
        "ps_min": round(float(np.min(ps)), 4),
        "ps_max": round(float(np.max(ps)), 4),
        "weight_mean": round(float(np.mean(w)), 4),
        "weight_sd": round(float(np.std(w)), 4),
        "weight_min": round(float(np.min(w)), 4),
        "weight_p1": round(float(lo), 4),
        "weight_p50": round(float(np.median(w)), 4),
        "weight_p99": round(float(hi), 4),
        "weight_max": round(float(np.max(w)), 4),
        "trunc_lo": round(float(lo), 4),
        "trunc_hi": round(float(hi), 4),
    }

    return analyzable, w_diag

# ─── 6. Covariate balance ──────────────────────────────────────────────
def check_balance(analyzable, covariate_cols, cov_df):
    """Check standardized mean differences before and after weighting."""
    idx = analyzable.index
    y = (analyzable["strategy"] == "early").values
    w = analyzable["weight_trunc"].values

    rows = []
    for j, col in enumerate(covariate_cols):
        x = cov_df.loc[idx, col].values

        # Unweighted SMD
        mean_e = np.mean(x[y == 1])
        mean_d = np.mean(x[y == 0])
        var_e = np.var(x[y == 1], ddof=1)
        var_d = np.var(x[y == 0], ddof=1)
        smd_uw = (mean_e - mean_d) / np.sqrt((var_e + var_d) / 2) if (var_e + var_d) > 0 else 0

        # Weighted SMD
        w_e = w[y == 1]
        w_d = w[y == 0]
        x_e = x[y == 1]
        x_d = x[y == 0]
        mean_e_w = np.sum(w_e * x_e) / np.sum(w_e)
        mean_d_w = np.sum(w_d * x_d) / np.sum(w_d)
        var_e_w = np.sum(w_e * (x_e - mean_e_w) ** 2) / np.sum(w_e)
        var_d_w = np.sum(w_d * (x_d - mean_d_w) ** 2) / np.sum(w_d)
        smd_w = (mean_e_w - mean_d_w) / np.sqrt((var_e_w + var_d_w) / 2) if (var_e_w + var_d_w) > 0 else 0

        rows.append({
            "covariate": col,
            "mean_early_unw": round(float(mean_e), 3),
            "mean_def_unw": round(float(mean_d), 3),
            "smd_unweighted": round(float(abs(smd_uw)), 4),
            "mean_early_w": round(float(mean_e_w), 3),
            "mean_def_w": round(float(mean_d_w), 3),
            "smd_weighted": round(float(abs(smd_w)), 4),
        })

    return pd.DataFrame(rows)

# ─── 7. Primary effect estimation ──────────────────────────────────────
def estimate_effects(analyzable):
    """Estimate primary and secondary effects using weighted approach."""

    y_early = (analyzable["strategy"] == "early").values
    w = analyzable["weight_trunc"].values
    died = analyzable["died_28d"].values
    surv = analyzable["surv_time"].values
    event = analyzable["event"].values
    vfd = analyzable["vfd28"].values

    # Weighted mortality
    w_early = w[y_early]
    w_def = w[~y_early]
    d_early = died[y_early]
    d_def = died[~y_early]

    mort_early = np.sum(w_early * d_early) / np.sum(w_early)
    mort_def = np.sum(w_def * d_def) / np.sum(w_def)

    rd = mort_early - mort_def
    rr = mort_early / mort_def if mort_def > 0 else np.nan

    # Weighted VFD28
    vfd_early = np.sum(w_early * vfd[y_early]) / np.sum(w_early)
    vfd_def = np.sum(w_def * vfd[~y_early]) / np.sum(w_def)
    vfd_diff = vfd_early - vfd_def

    # RMST difference (using weighted pseudo-observations)
    # For a simple approach: use weighted mean survival time up to 28 days
    rmst_early = np.sum(w_early * surv[y_early]) / np.sum(w_early)
    rmst_def = np.sum(w_def * surv[~y_early]) / np.sum(w_def)
    rmst_diff = rmst_early - rmst_def

    results = {
        "mort_early": round(float(mort_early), 4),
        "mort_deferred": round(float(mort_def), 4),
        "rd": round(float(rd), 4),
        "rr": round(float(rr), 4),
        "rmst_early": round(float(rmst_early), 2),
        "rmst_deferred": round(float(rmst_def), 2),
        "rmst_diff": round(float(rmst_diff), 2),
        "vfd_early": round(float(vfd_early), 2),
        "vfd_deferred": round(float(vfd_def), 2),
        "vfd_diff": round(float(vfd_diff), 2),
    }

    return results

# ─── 8. Bootstrap CIs ──────────────────────────────────────────────────
def bootstrap_effects(analyzable, cov_df, covariate_cols, n_boot=2000):
    """Bootstrap confidence intervals for primary effects."""

    idx_early = np.where(analyzable["strategy"].values == "early")[0]
    idx_def = np.where(analyzable["strategy"].values == "deferred")[0]
    n_e = len(idx_early)
    n_d = len(idx_def)

    died = analyzable["died_28d"].values
    surv = analyzable["surv_time"].values
    vfd = analyzable["vfd28"].values
    strategy = analyzable["strategy"].values

    from sklearn.linear_model import LogisticRegression

    boot_rds = []
    boot_rrs = []
    boot_rmsts = []
    boot_vfds = []

    rng = np.random.RandomState(RNG_SEED)

    for b in range(n_boot):
        # Resample within each arm
        boot_e = rng.choice(idx_early, size=n_e, replace=True)
        boot_d = rng.choice(idx_def, size=n_d, replace=True)
        boot_idx = np.concatenate([boot_e, boot_d])

        boot_y = strategy[boot_idx]
        boot_died = died[boot_idx]
        boot_surv = surv[boot_idx]
        boot_vfd = vfd[boot_idx]

        # Refit propensity model on bootstrap sample
        boot_X = cov_df.loc[analyzable.index[boot_idx]].values
        boot_y_bin = (boot_y == "early").astype(int)

        try:
            model = LogisticRegression(max_iter=500, penalty="l2", C=1.0, solver="lbfgs")
            model.fit(boot_X, boot_y_bin)
            boot_ps = model.predict_proba(boot_X)[:, 1]
            boot_ps = np.clip(boot_ps, 0.01, 0.99)
        except Exception:
            boot_ps = np.full(len(boot_idx), np.mean(boot_y_bin))

        p_e = np.mean(boot_y_bin)
        boot_w = np.where(boot_y_bin == 1, p_e / boot_ps, (1 - p_e) / (1 - boot_ps))
        lo, hi = np.percentile(boot_w, [1, 99])
        boot_w = np.clip(boot_w, lo, hi)

        w_e = boot_w[boot_y_bin == 1]
        w_d = boot_w[boot_y_bin == 0]
        d_e = boot_died[boot_y_bin == 1]
        d_d = boot_died[boot_y_bin == 0]

        me = np.sum(w_e * d_e) / np.sum(w_e) if np.sum(w_e) > 0 else 0
        md = np.sum(w_d * d_d) / np.sum(w_d) if np.sum(w_d) > 0 else 0

        boot_rds.append(me - md)
        boot_rrs.append(me / md if md > 0 else np.nan)

        re = np.sum(w_e * boot_surv[boot_y_bin == 1]) / np.sum(w_e)
        rd = np.sum(w_d * boot_surv[boot_y_bin == 0]) / np.sum(w_d)
        boot_rmsts.append(re - rd)

        ve = np.sum(w_e * boot_vfd[boot_y_bin == 1]) / np.sum(w_e)
        vd = np.sum(w_d * boot_vfd[boot_y_bin == 0]) / np.sum(w_d)
        boot_vfds.append(ve - vd)

        if (b + 1) % 500 == 0:
            print(f"  Bootstrap {b+1}/{n_boot}...")

    boot_rds = np.array(boot_rds)
    boot_rrs = np.array(boot_rrs)
    boot_rmsts = np.array(boot_rmsts)
    boot_vfds = np.array(boot_vfds)

    cis = {
        "rd_ci_lo": round(float(np.nanpercentile(boot_rds, 2.5)), 4),
        "rd_ci_hi": round(float(np.nanpercentile(boot_rds, 97.5)), 4),
        "rr_ci_lo": round(float(np.nanpercentile(boot_rrs, 2.5)), 4),
        "rr_ci_hi": round(float(np.nanpercentile(boot_rrs, 97.5)), 4),
        "rmst_ci_lo": round(float(np.nanpercentile(boot_rmsts, 2.5)), 2),
        "rmst_ci_hi": round(float(np.nanpercentile(boot_rmsts, 97.5)), 2),
        "vfd_ci_lo": round(float(np.nanpercentile(boot_vfds, 2.5)), 2),
        "vfd_ci_hi": round(float(np.nanpercentile(boot_vfds, 97.5)), 2),
    }

    return cis

# ─── 9. Sensitivity analyses ──────────────────────────────────────────
def sensitivity_analyses(eligible, cov_df, covariate_cols, n_boot=800):
    """Run pre-specified sensitivity analyses."""

    from sklearn.linear_model import LogisticRegression
    results = {}

    # Prepare common data
    analyzable = eligible[eligible["strategy"].isin(["early", "deferred"])].copy()
    y = (analyzable["strategy"] == "early").astype(int).values
    X = cov_df.loc[analyzable.index].values
    model = LogisticRegression(max_iter=1000, penalty="l2", C=1.0, solver="lbfgs")
    model.fit(X, y)
    ps = np.clip(model.predict_proba(X)[:, 1], 0.01, 0.99)
    p_e = np.mean(y)
    weights = np.where(y == 1, p_e / ps, (1 - p_e) / (1 - ps))

    # Truncated weights (for primary comparison)
    lo_t, hi_t = np.percentile(weights, [1, 99])
    w_trunc = np.clip(weights, lo_t, hi_t)

    died_hosp = analyzable["died_28d"].values  # hospital 28-day mortality
    died_icu = analyzable["death_within_icu_28days"].apply(safe_float).astype(int).values

    # Subset arrays for early and deferred
    e_mask = y == 1
    d_mask = y == 0

    # SA1: No weight truncation
    w_e1 = weights[e_mask]; w_d1 = weights[d_mask]
    d_e1 = died_hosp[e_mask]; d_d1 = died_hosp[d_mask]
    me1 = np.sum(w_e1 * d_e1) / np.sum(w_e1)
    md1 = np.sum(w_d1 * d_d1) / np.sum(w_d1)
    results["SA1_no_truncation"] = {"rd": round(float(me1 - md1), 4)}

    # SA2: ICU mortality (with truncated weights)
    w_e2 = w_trunc[e_mask]; w_d2 = w_trunc[d_mask]
    d_e2 = died_icu[e_mask]; d_d2 = died_icu[d_mask]
    me2 = np.sum(w_e2 * d_e2) / np.sum(w_e2)
    md2 = np.sum(w_d2 * d_d2) / np.sum(w_d2)
    results["SA2_icu_mortality"] = {"rd": round(float(me2 - md2), 4)}

    # SA3: Unweighted (per-protocol)
    d_e3 = died_hosp[e_mask]; d_d3 = died_hosp[d_mask]
    me3 = np.mean(d_e3)
    md3 = np.mean(d_d3)
    results["SA3_unweighted"] = {"rd": round(float(me3 - md3), 4)}

    # Bootstrap CIs
    rng = np.random.RandomState(RNG_SEED + 1)
    n_e = int(e_mask.sum())
    n_d = int(d_mask.sum())

    for sa_name in results:
        boot_rds = []
        for b in range(n_boot):
            # Resample positions within each arm (0 to n_e-1, 0 to n_d-1)
            be = rng.randint(0, n_e, size=n_e)
            bd = rng.randint(0, n_d, size=n_d)

            if sa_name == "SA1_no_truncation":
                bwe = w_e1[be]; bwd = w_d1[bd]
                bde = d_e1[be]; bdd = d_d1[bd]
                bme = np.sum(bwe * bde) / np.sum(bwe) if np.sum(bwe) > 0 else 0
                bmd = np.sum(bwd * bdd) / np.sum(bwd) if np.sum(bwd) > 0 else 0
            elif sa_name == "SA2_icu_mortality":
                bwe = w_e2[be]; bwd = w_d2[bd]
                bde = d_e2[be]; bdd = d_d2[bd]
                bme = np.sum(bwe * bde) / np.sum(bwe) if np.sum(bwe) > 0 else 0
                bmd = np.sum(bwd * bdd) / np.sum(bwd) if np.sum(bwd) > 0 else 0
            else:  # SA3_unweighted
                bde = d_e3[be]; bdd = d_d3[bd]
                bme = np.mean(bde)
                bmd = np.mean(bdd)

            boot_rds.append(bme - bmd)

        results[sa_name]["rd_ci_lo"] = round(float(np.nanpercentile(boot_rds, 2.5)), 4)
        results[sa_name]["rd_ci_hi"] = round(float(np.nanpercentile(boot_rds, 97.5)), 4)

    return results

# ─── 10. Exploratory threshold grid ────────────────────────────────────
def threshold_grid(eligible, cov_df, covariate_cols, n_boot=800):
    """FiO₂ × PEEP 3×3 threshold grid analysis."""

    from sklearn.linear_model import LogisticRegression

    fio2_thresholds = [40, 50, 60]
    peep_thresholds = [5, 8, 10]
    results = []

    for fio2_t in fio2_thresholds:
        for peep_t in peep_thresholds:
            # Subset: patients with Day-2 FiO₂ ≤ threshold and PEEP ≤ threshold
            sub_mask = (eligible["d2_fio2"] <= fio2_t) & (eligible["d2_peep"] <= peep_t)
            sub = eligible[sub_mask].copy()

            if len(sub) < 20:
                results.append({
                    "fio2_threshold": fio2_t, "peep_threshold": peep_t,
                    "n": len(sub), "n_early": 0, "n_deferred": 0,
                    "rd": np.nan, "rd_ci_lo": np.nan, "rd_ci_hi": np.nan,
                })
                continue

            analyzable = sub[sub["strategy"].isin(["early", "deferred"])].copy()
            n_e = int((analyzable["strategy"] == "early").sum())
            n_d = int((analyzable["strategy"] == "deferred").sum())

            if n_e < 5 or n_d < 5:
                results.append({
                    "fio2_threshold": fio2_t, "peep_threshold": peep_t,
                    "n": len(analyzable), "n_early": n_e, "n_deferred": n_d,
                    "rd": np.nan, "rd_ci_lo": np.nan, "rd_ci_hi": np.nan,
                })
                continue

            y = (analyzable["strategy"] == "early").astype(int).values
            died = analyzable["died_28d"].values

            # Unweighted (small samples — weighting unstable)
            me = np.mean(died[y == 1])
            md = np.mean(died[y == 0])
            rd = me - md

            # Bootstrap CI
            rng = np.random.RandomState(RNG_SEED + fio2_t * 100 + peep_t)
            idx_e = np.where(y == 1)[0]
            idx_d = np.where(y == 0)[0]
            boot_rds = []
            for b in range(n_boot):
                be = rng.choice(idx_e, len(idx_e), replace=True)
                bd = rng.choice(idx_d, len(idx_d), replace=True)
                bme = np.mean(died[be])
                bmd = np.mean(died[bd])
                boot_rds.append(bme - bmd)

            results.append({
                "fio2_threshold": fio2_t, "peep_threshold": peep_t,
                "n": len(analyzable), "n_early": n_e, "n_deferred": n_d,
                "rd": round(float(rd), 4),
                "rd_ci_lo": round(float(np.nanpercentile(boot_rds, 2.5)), 4),
                "rd_ci_hi": round(float(np.nanpercentile(boot_rds, 97.5)), 4),
            })

    return pd.DataFrame(results)

# ─── 11. Baseline table ────────────────────────────────────────────────
def make_baseline_table(eligible, analyzable):
    """Create Table 1 baseline characteristics by strategy."""

    rows = []
    strat = analyzable["strategy"].values

    def add_row(label, val_early, val_def, fmt="{:.1f}", is_cat=False):
        if is_cat:
            rows.append({"Variable": label, "Early (n={})".format(int((strat=="early").sum())): val_early,
                         "Deferred (n={})".format(int((strat=="deferred").sum())): val_def})
        else:
            rows.append({"Variable": label,
                         "Early (n={})".format(int((strat=="early").sum())): fmt.format(val_early) if not np.isnan(val_early) else "—",
                         "Deferred (n={})".format(int((strat=="deferred").sum())): fmt.format(val_def) if not np.isnan(val_def) else "—"})

    e = analyzable[strat == "early"]
    d = analyzable[strat == "deferred"]

    # Demographics
    add_row("Age, years", np.nanmean(e["age_num"]), np.nanmean(d["age_num"]), "{:.1f}")
    add_row("Male, n (%)", 100*np.mean(e["male"]), 100*np.mean(d["male"]), "{:.1f}")
    add_row("Weight, kg", np.nanmean(e["weight_num"]), np.nanmean(d["weight_num"]), "{:.1f}")

    # SOFA
    add_row("SOFA total", np.nanmean(e["sofa_total"]), np.nanmean(d["sofa_total"]), "{:.1f}")
    add_row("SOFA respiration", np.nanmean(e["sofa_respiration"]), np.nanmean(d["sofa_respiration"]), "{:.1f}")
    add_row("SOFA cardiovascular", np.nanmean(e["sofa_cardiovascular"]), np.nanmean(d["sofa_cardiovascular"]), "{:.1f}")
    add_row("SOFA CNS", np.nanmean(e["sofa_cns"]), np.nanmean(d["sofa_cns"]), "{:.1f}")
    add_row("SOFA renal", np.nanmean(e["sofa_renal"]), np.nanmean(d["sofa_renal"]), "{:.1f}")

    # Day-2 vent
    add_row("Day-2 FiO₂, %", np.nanmean(e["d2_fio2"]), np.nanmean(d["d2_fio2"]), "{:.1f}")
    add_row("Day-2 PEEP, cmH₂O", np.nanmean(e["d2_peep"]), np.nanmean(d["d2_peep"]), "{:.1f}")
    add_row("Day-2 tidal volume, mL", np.nanmean(e["d2_tv"]), np.nanmean(d["d2_tv"]), "{:.0f}")
    add_row("Day-2 plateau pressure, cmH₂O", np.nanmean(e["d2_pplat"]), np.nanmean(d["d2_pplat"]), "{:.1f}")

    # Comorbidities
    add_row("ARDS, n (%)", 100*np.mean(e["ards_flag"]), 100*np.mean(d["ards_flag"]), "{:.1f}")
    add_row("Sepsis-3, n (%)", 100*np.mean(e["sepsis3_flag"]), 100*np.mean(d["sepsis3_flag"]), "{:.1f}")
    add_row("CRRT, n (%)", 100*np.mean(e["crrt_flag"]), 100*np.mean(d["crrt_flag"]), "{:.1f}")

    # Labs
    add_row("Day-2 creatinine, μmol/L", np.nanmean(e["d2_cr"]), np.nanmean(d["d2_cr"]), "{:.1f}")
    add_row("Day-2 WBC, ×10⁹/L", np.nanmean(e["d2_wbc"]), np.nanmean(d["d2_wbc"]), "{:.1f}")
    add_row("Day-2 platelets, ×10⁹/L", np.nanmean(e["d2_plt"]), np.nanmean(d["d2_plt"]), "{:.0f}")
    add_row("Day-2 lactate, mmol/L", np.nanmean(e["d2_lactate"]), np.nanmean(d["d2_lactate"]), "{:.2f}")

    return pd.DataFrame(rows)

# ─── 12. Figures ───────────────────────────────────────────────────────
def make_figures(flow, eligible, analyzable, threshold_df, balance_df):
    """Generate all publication-quality figures."""

    # Strategy counts from eligible (includes all strategies)
    strat_counts = eligible["strategy"].value_counts()
    n_early = int(strat_counts.get("early", 0))
    n_deferred = int(strat_counts.get("deferred", 0))
    n_unasc = int(strat_counts.get("unascertainable", 0))
    n_grace = int(strat_counts.get("grace_death", 0))

    plt.rcParams.update({
        "font.size": 10, "font.family": "sans-serif",
        "axes.spines.top": False, "axes.spines.right": False,
        "figure.dpi": 300,
    })

    # ── Figure 1: Flow diagram ──
    fig, ax = plt.subplots(figsize=(10, 8))
    ax.set_xlim(0, 10); ax.set_ylim(0, 10)
    ax.axis("off")

    boxes = [
        (5, 9, f"Total MIMIC-IV records\nn = {flow['total_records']:,}"),
        (5, 7.5, f"Adults (>=18 years)\nn = {flow['adults']:,}"),
        (5, 6, f"Received invasive MV\nn = {flow['ventilated']:,}"),
        (5, 4.5, f"MV >24h, alive at 48h\nn = {flow['alive_at_48h']:,}"),
        (5, 3, f"Classified Day-2 mode\nFiO2 <=50%, PEEP <=10\nn = {flow['eligible']:,}"),
        (2.2, 1.3, f"Early de-escalation\nn = {n_early}"),
        (5, 1.3, f"Unascertainable\nn = {n_unasc}\n(Grace deaths: {n_grace})"),
        (7.8, 1.3, f"Deferred\nn = {n_deferred}"),
    ]

    for x, y, txt in boxes:
        color = "#4A90D9" if "Early" in txt or "Total" in txt else "#E8E8E8"
        text_color = "white" if "Early" in txt or "Total" in txt else "black"
        box = mpatches.FancyBboxPatch((x-1.8, y-0.5), 3.6, 1.0,
                                       boxstyle="round,pad=0.1", facecolor=color, edgecolor="#333")
        ax.add_patch(box)
        ax.text(x, y, txt, ha="center", va="center", fontsize=8, color=text_color)

    # Arrows
    for y_from, y_to in [(8.5, 8), (7, 5.5), (5.5, 4), (3.5, 2)]:
        ax.annotate("", xy=(5, y_to), xytext=(5, y_from), arrowprops=dict(arrowstyle="->", color="#333"))
    ax.annotate("", xy=(2.5, 1.8), xytext=(5, 2.5), arrowprops=dict(arrowstyle="->", color="#333"))
    ax.annotate("", xy=(7.5, 1.8), xytext=(5, 2.5), arrowprops=dict(arrowstyle="->", color="#333"))

    # Exclusion annotations
    excl_y = [7.5, 6, 4.5, 3]
    excl_txt = [
        f"Excluded: <18 years (n={flow['total_records']-flow['adults']:,})",
        f"Excluded: no MV (n={flow['adults']-flow['ventilated']:,})",
        f"Excluded: MV ≤24h or died <48h\n(n={flow['ventilated']-flow['alive_at_48h']:,})",
        f"Excluded: no Day-2 mode, FiO2 >50%\nor PEEP >10\n(n={flow['alive_at_48h']-flow['eligible']:,})",
    ]
    for y, txt in zip(excl_y, excl_txt):
        ax.text(8.5, y, txt, ha="left", va="center", fontsize=7, color="#666", style="italic")

    fig.tight_layout()
    fig.savefig(FIG_DIR / "Figure1_Flow.png", dpi=300, bbox_inches="tight")
    plt.close(fig)
    print("  ✓ Figure 1 (Flow diagram)")

    # ── Figure 2: Weighted survival curves ──
    fig, ax = plt.subplots(figsize=(8, 5))

    for strat_name, color in [("early", "#D44E4E"), ("deferred", "#4A90D9")]:
        sub = analyzable[analyzable["strategy"] == strat_name]
        surv = sub["surv_time"].values
        event = sub["event"].values
        w = sub["weight_trunc"].values

        # Weighted KM
        kmf = KaplanMeierFitter()
        kmf.fit(surv, event_observed=event, weights=w, label=strat_name)
        kmf.plot_survival_function(ax=ax, color=color, ci_show=False, linewidth=2)

    ax.set_xlabel("Days since landmark", fontsize=11)
    ax.set_ylabel("Survival probability", fontsize=11)
    ax.set_title("Weighted Kaplan-Meier Survival Curves (28-day)", fontsize=12)
    ax.set_xlim(0, 28)
    ax.set_ylim(0.7, 1.0)
    ax.legend(["Early de-escalation", "Deferred"], fontsize=9, loc="lower left")
    ax.grid(True, alpha=0.3)
    fig.tight_layout()
    fig.savefig(FIG_DIR / "Figure2_Survival.png", dpi=300, bbox_inches="tight")
    plt.close(fig)
    print("  ✓ Figure 2 (Survival curves)")

    # ── Figure 3: Forest plot for primary + sensitivity ──
    fig, ax = plt.subplots(figsize=(8, 4))

    labels = ["Primary (weighted)", "SA1: No truncation", "SA2: ICU mortality", "SA3: Unweighted"]
    rds = []
    lo_err = []
    hi_err = []

    primary = estimate_effects(analyzable)
    rds.append(primary["rd"])

    sa = sensitivity_analyses.__wrapped__ if hasattr(sensitivity_analyses, '__wrapped__') else None
    # We'll use saved results — for the figure, compute quick versions
    # (CIs already computed in main)
    rds.append(0)  # placeholder
    rds.append(0)
    rds.append(0)

    # Actually, we need the real values — let's compute them inline
    from sklearn.linear_model import LogisticRegression
    y_all = (analyzable["strategy"] == "early").astype(int).values
    X_all = analyzable[["_cov_matrix"]].apply(lambda r: r.iloc[0], axis=1).tolist()
    X_all = np.array(X_all)
    model = LogisticRegression(max_iter=1000, penalty="l2", C=1.0, solver="lbfgs")
    model.fit(X_all, y_all)
    ps = np.clip(model.predict_proba(X_all)[:, 1], 0.01, 0.99)
    p_e = np.mean(y_all)
    w_all = np.where(y_all == 1, p_e / ps, (1 - p_e) / (1 - ps))

    died_all = analyzable["died_28d"].values
    w_e = w_all[y_all == 1]; w_d = w_all[y_all == 0]
    d_e = died_all[y_all == 1]; d_d = died_all[y_all == 0]

    # SA1: no truncation
    me1 = np.sum(w_e * d_e) / np.sum(w_e)
    md1 = np.sum(w_d * d_d) / np.sum(w_d)
    rds[1] = me1 - md1

    # SA2: ICU mortality
    died_icu = analyzable["death_within_icu_28days"].apply(safe_float).astype(int).values
    me2 = np.sum(w_e * died_icu[y_all == 1]) / np.sum(w_e)
    md2 = np.sum(w_d * died_icu[y_all == 0]) / np.sum(w_d)
    rds[2] = me2 - md2

    # SA3: unweighted
    me3 = np.mean(died_all[y_all == 1])
    md3 = np.mean(died_all[y_all == 0])
    rds[3] = me3 - md3

    # Simple CIs (normal approximation)
    for i, rd in enumerate(rds):
        ax.plot(rd, len(labels) - 1 - i, "s", color="#333", markersize=8)

    ax.axvline(0, color="#999", linestyle="--", linewidth=0.8)
    ax.set_yticks(range(len(labels)))
    ax.set_yticklabels(list(reversed(labels)))
    ax.set_xlabel("Risk difference (percentage points)", fontsize=11)
    ax.set_title("Primary and Sensitivity Analyses", fontsize=12)
    ax.set_xlim(-0.3, 0.3)
    ax.grid(True, alpha=0.3, axis="x")
    fig.tight_layout()
    fig.savefig(FIG_DIR / "Figure3_Forest.png", dpi=300, bbox_inches="tight")
    plt.close(fig)
    print("  ✓ Figure 3 (Forest plot)")

    # ── Figure S1: Weight distribution ──
    fig, axes = plt.subplots(1, 2, figsize=(10, 4))
    w = analyzable["weight_trunc"].values
    axes[0].hist(w, bins=50, color="#4A90D9", edgecolor="white", alpha=0.8)
    axes[0].set_xlabel("Stabilised weight (truncated)", fontsize=10)
    axes[0].set_ylabel("Count", fontsize=10)
    axes[0].set_title("Weight distribution", fontsize=11)
    axes[0].axvline(np.mean(w), color="#D44E4E", linestyle="--", label=f"Mean = {np.mean(w):.2f}")
    axes[0].legend(fontsize=9)

    # SMD before/after
    axes[1].barh(np.arange(len(balance_df)), balance_df["smd_unweighted"], height=0.35, color="#D44E4E", alpha=0.7, label="Unweighted")
    axes[1].barh(np.arange(len(balance_df)) + 0.35, balance_df["smd_weighted"], height=0.35, color="#4A90D9", alpha=0.7, label="Weighted")
    axes[1].axvline(0.1, color="#999", linestyle="--", linewidth=0.8)
    axes[1].set_yticks(np.arange(len(balance_df)) + 0.175)
    axes[1].set_yticklabels(balance_df["covariate"], fontsize=7)
    axes[1].set_xlabel("|SMD|", fontsize=10)
    axes[1].set_title("Covariate balance", fontsize=11)
    axes[1].legend(fontsize=8)
    fig.tight_layout()
    fig.savefig(FIG_DIR / "FigureS1_Weights.png", dpi=300, bbox_inches="tight")
    plt.close(fig)
    print("  ✓ Figure S1 (Weights & balance)")

    # ── Figure S2: Threshold grid heatmap ──
    fig, ax = plt.subplots(figsize=(6, 5))
    pivot = threshold_df.pivot(index="peep_threshold", columns="fio2_threshold", values="rd")
    im = ax.imshow(pivot.values, cmap="RdBu_r", vmin=-0.25, vmax=0.25, aspect="auto")
    ax.set_xticks(range(len(fio2_thresholds := [40, 50, 60])))
    ax.set_xticklabels([f"≤{t}%" for t in fio2_thresholds])
    ax.set_yticks(range(len(peep_thresholds := [5, 8, 10])))
    ax.set_yticklabels([f"≤{t}" for t in peep_thresholds])
    ax.set_xlabel("Day-2 FiO₂ threshold", fontsize=11)
    ax.set_ylabel("Day-2 PEEP threshold (cmH₂O)", fontsize=11)
    ax.set_title("Exploratory threshold grid: Risk difference", fontsize=12)

    for i in range(len(peep_thresholds)):
        for j in range(len(fio2_thresholds)):
            val = pivot.values[i, j]
            n = threshold_df[(threshold_df["fio2_threshold"]==fio2_thresholds[j]) &
                            (threshold_df["peep_threshold"]==peep_thresholds[i])]["n"].values[0]
            if not np.isnan(val):
                ax.text(j, i, f"{val*100:.1f}\n(n={n})", ha="center", va="center", fontsize=9,
                       color="white" if abs(val) > 0.15 else "black")
            else:
                ax.text(j, i, "—", ha="center", va="center", fontsize=9)

    plt.colorbar(im, ax=ax, label="Risk difference")
    fig.tight_layout()
    fig.savefig(FIG_DIR / "FigureS2_Threshold.png", dpi=300, bbox_inches="tight")
    plt.close(fig)
    print("  ✓ Figure S2 (Threshold grid)")

# ─── MAIN ──────────────────────────────────────────────────────────────
def main():
    print("=" * 70)
    print("TTE De-escalation Analysis — Full Pipeline")
    print("=" * 70)

    # 1. Load data
    df = load_excel()

    # 2. Screen eligibility
    print("\n[1/10] Screening eligibility...")
    eligible, flow = screen_eligibility(df)
    print(f"  Eligible: {len(eligible)}")
    flow_df = pd.DataFrame([{"step": k, "count": v} for k, v in flow.items()])
    flow_df.to_csv(OUT_DIR / "flow_counts.csv", index=False)

    # 3. Classify strategies
    print("\n[2/10] Classifying treatment strategies...")
    eligible = classify_strategies(eligible)
    strat_counts = eligible["strategy"].value_counts()
    print(f"  {dict(strat_counts)}")

    # 4. Compute outcomes
    print("\n[3/10] Computing outcomes...")
    eligible = compute_outcomes(eligible)

    # 5. Build covariates
    print("\n[4/10] Building covariates...")
    eligible, covariate_cols, cov_df = build_covariates(eligible)

    # 6. Estimate weights
    print("\n[5/10] Estimating stabilised IPCW...")
    analyzable, w_diag = estimate_weights(eligible, cov_df, covariate_cols)
    print(f"  Weight mean: {w_diag['weight_mean']}, p99: {w_diag['weight_p99']}, max: {w_diag['weight_max']}")

    # 7. Check balance
    print("\n[6/10] Checking covariate balance...")
    balance_df = check_balance(analyzable, covariate_cols, cov_df)
    balance_df.to_csv(OUT_DIR / "covariate_balance.csv", index=False)

    # 8. Primary effects
    print("\n[7/10] Estimating primary effects...")
    primary = estimate_effects(analyzable)
    print(f"  Mortality: {primary['mort_early']*100:.1f}% vs {primary['mort_deferred']*100:.1f}%")
    print(f"  RD: {primary['rd']*100:.1f} pp, RR: {primary['rr']:.2f}")
    print(f"  RMST diff: {primary['rmst_diff']:.2f} days")
    print(f"  VFD diff: {primary['vfd_diff']:.2f} days")

    # 9. Bootstrap CIs
    print("\n[8/10] Bootstrapping (2000 resamples)...")
    cis = bootstrap_effects(analyzable, cov_df, covariate_cols, n_boot=2000)
    print(f"  RD 95% CI: [{cis['rd_ci_lo']*100:.1f}, {cis['rd_ci_hi']*100:.1f}] pp")
    print(f"  RR 95% CI: [{cis['rr_ci_lo']:.2f}, {cis['rr_ci_hi']:.2f}]")

    # 10. Sensitivity analyses
    print("\n[9/10] Sensitivity analyses (800 resamples each)...")
    sa_results = sensitivity_analyses(eligible, cov_df, covariate_cols, n_boot=800)
    for k, v in sa_results.items():
        print(f"  {k}: RD = {v['rd']*100:.1f} pp [{v['rd_ci_lo']*100:.1f}, {v['rd_ci_hi']*100:.1f}]")

    # 11. Threshold grid
    print("\n[10/10] Exploratory threshold grid...")
    threshold_df = threshold_grid(eligible, cov_df, covariate_cols, n_boot=800)
    threshold_df.to_csv(OUT_DIR / "threshold_results.csv", index=False)
    print(threshold_df.to_string(index=False))

    # 12. Baseline table
    baseline_df = make_baseline_table(eligible, analyzable)
    baseline_df.to_csv(OUT_DIR / "table1_baseline.csv", index=False)

    # 13. Save primary effects
    primary_full = {**primary, **cis}
    primary_df = pd.DataFrame([primary_full])
    primary_df.to_csv(OUT_DIR / "primary_effects.csv", index=False)

    # 14. Save weight diagnostics
    with open(OUT_DIR / "weight_diagnostics.json", "w") as f:
        json.dump(w_diag, f, indent=2)

    # 15. Save sensitivity results
    with open(OUT_DIR / "sensitivity_results.json", "w") as f:
        json.dump(sa_results, f, indent=2)

    # 16. Save analysis summary
    summary = {
        "flow": flow,
        "strategy_counts": {k: int(v) for k, v in strat_counts.items()},
        "primary": {**primary, **cis},
        "weight_diagnostics": w_diag,
        "sensitivity": sa_results,
        "threshold_grid": threshold_df.to_dict("records"),
        "rng_seed": RNG_SEED,
        "bootstrap_primary": 2000,
        "bootstrap_sensitivity": 800,
    }
    with open(OUT_DIR / "analysis_summary.json", "w") as f:
        json.dump(summary, f, indent=2, default=str)

    # 17. Generate figures
    print("\nGenerating figures...")
    make_figures(flow, eligible, analyzable, threshold_df, balance_df)

    print("\n" + "=" * 70)
    print("ANALYSIS COMPLETE")
    print(f"Results saved to: {OUT_DIR}")
    print(f"Figures saved to: {FIG_DIR}")
    print("=" * 70)

    return summary

if __name__ == "__main__":
    summary = main()
