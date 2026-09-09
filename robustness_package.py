"""
Robustness package for the TTE De-escalation analysis (Post-hoc robustness addendum,
analyses pre-specified before execution of this script; all use the locked analytic
dataset from run_analysis.py, unchanged).

Contents
--------
SA4 : Extended covariate adjustment (+15 pre-landmark confounder dimensions:
      sedation depth [RASS], vasopressor/sedative/analgesic infusion doses,
      haemodynamics [MAP, HR, RR, SpO2], fluid balance, additional Day-2 labs,
      baseline inflammatory markers, BMI, full SOFA subscores)
SA5 : Alternative weight truncation (5th-95th percentile instead of 1st-99th)
AIPW: Augmented inverse-probability-weighted (doubly robust) estimator for the
      primary RD, with bootstrap CI
E-value: minimum strength of association (risk ratio scale) an unmeasured
      confounder would need with both treatment and outcome to explain away
      the observed effect
Subgroups: weighted RD within Day-2 mode (controlled vs assisted), age, SOFA,
      sepsis-3, ARDS strata

All bootstrap resamples refit the propensity (and outcome, for AIPW) models.
"""
import json
import sys
import numpy as np
import pandas as pd
from pathlib import Path
from sklearn.linear_model import LogisticRegression

sys.path.insert(0, "/Users/zhangrui/Desktop/tte/TTE_Deescalation_v2")
import run_analysis as ra
safe_float = ra.safe_float

OUT = ra.OUT_DIR
N_BOOT_SA = 800
N_BOOT_SUB = 400


# ── helpers ──────────────────────────────────────────────────────────────
def _timepoint_mean(row, safe=ra.safe_float):
    vals = [safe(v) for v in row.values]
    vals = [v for v in vals if v is not None and not (isinstance(v, float) and np.isnan(v))]
    return float(np.mean(vals)) if vals else np.nan


def cols_with(eligible, key):
    return [c for c in eligible.columns if key in str(c)]


def build_extended_covariates(eligible, base_cols):
    """Add ~15 pre-landmark confounder dimensions to the base covariate set."""
    ext = eligible.copy()

    # Sedation depth (RASS, mean of Day-2 timepoints)
    rass_cols = cols_with(ext, "RASS评分_Day2")
    ext["d2_rass"] = ext[rass_cols].apply(_timepoint_mean, axis=1) if rass_cols else np.nan

    # Infusion doses (Day-2 mean): norepinephrine, propofol, fentanyl
    ext["d2_norepi"] = ext[cols_with(ext, "去甲肾上腺素")].apply(_timepoint_mean, axis=1)
    ext["d2_propofol"] = ext[cols_with(ext, "丙泊酚")].apply(_timepoint_mean, axis=1)
    ext["d2_fentanyl"] = ext[cols_with(ext, "芬太尼")].apply(_timepoint_mean, axis=1)

    # Haemodynamics (Day-2 mean)
    ext["d2_map"] = ext[cols_with(ext, "有创血压_平均压")].apply(_timepoint_mean, axis=1)
    ext["d2_hr"] = ext[cols_with(ext, "心率(次/min)_Day2")].apply(_timepoint_mean, axis=1)
    ext["d2_rr"] = ext[cols_with(ext, "呼吸频率(次/min)_Day2")].apply(_timepoint_mean, axis=1)
    ext["d2_spo2"] = ext[cols_with(ext, "脉搏氧饱和度")].apply(_timepoint_mean, axis=1)

    # Fluid balance Day-2 (L)
    ext["d2_fluid_balance"] = ext["总平衡(ml)_Day2"].apply(safe_float) / 1000.0

    # Additional Day-2 labs
    lab_map = {
        "d2_bun": "尿素氮_BUN(mmol/L)_Day2",
        "d2_alb": "白蛋白_ALB(g/L)_Day2",
        "d2_tbil": "总胆红素_TBIL(μmol/L)_Day2",
        "d2_hb": "血红蛋白_Hb(g/L)_Day2",
        "d2_lym": "淋巴细胞_LYM(×10^9/L)_Day2",
    }
    for new, src in lab_map.items():
        ext[new] = ext[src].apply(safe_float)

    # Baseline inflammatory markers
    ext["base_pct"] = ext["PCT_基线(ng/mL)"].apply(safe_float)
    ext["base_il6"] = ext["IL-6_基线(pg/mL)"].apply(safe_float)

    # BMI
    h = ext["height"].apply(safe_float) / 100.0
    w = ext["weight"].apply(safe_float)
    ext["bmi"] = np.where((h > 0.8) & (h < 2.3) & (w > 20) & (w < 400), w / h**2, np.nan)

    # Remaining SOFA subscores (not in base model)
    ext["sofa_coagulation"] = ext["coagulation"].apply(safe_float)
    ext["sofa_liver"] = ext["liver"].apply(safe_float)

    new_cols = [
        "d2_rass", "d2_norepi", "d2_propofol", "d2_fentanyl",
        "d2_map", "d2_hr", "d2_rr", "d2_spo2", "d2_fluid_balance",
        "d2_bun", "d2_alb", "d2_tbil", "d2_hb", "d2_lym",
        "base_pct", "base_il6", "bmi", "sofa_coagulation", "sofa_liver",
    ]

    ext_cols = base_cols + new_cols
    cov = ext[ext_cols].copy()
    for c in cov.columns:
        cov[c] = pd.to_numeric(cov[c], errors="coerce")
        med = cov[c].median()
        cov[c] = cov[c].fillna(med if not pd.isna(med) else 0)
    return ext, ext_cols, cov


def weighted_rd(y, died, w):
    """Weighted risk difference."""
    me = np.sum(w[y == 1] * died[y == 1]) / np.sum(w[y == 1])
    md = np.sum(w[y == 0] * died[y == 0]) / np.sum(w[y == 0])
    return me - md, me, md


def fit_ps_weights(X, y, trunc=(1, 99)):
    model = LogisticRegression(max_iter=1000, penalty="l2", C=1.0, solver="lbfgs")
    model.fit(X, y)
    ps = np.clip(model.predict_proba(X)[:, 1], 0.01, 0.99)
    p_e = np.mean(y)
    w = np.where(y == 1, p_e / ps, (1 - p_e) / (1 - ps))
    lo, hi = np.percentile(w, trunc)
    return np.clip(w, lo, hi), ps


def aipw_rd(X, y, died):
    """Doubly robust AIPW estimate of RD (outcome model fit separately per arm)."""
    m1 = LogisticRegression(max_iter=1000, penalty="l2", C=1.0, solver="lbfgs")
    m1.fit(X[y == 1], died[y == 1])
    m0 = LogisticRegression(max_iter=1000, penalty="l2", C=1.0, solver="lbfgs")
    m0.fit(X[y == 0], died[y == 0])
    mu1 = m1.predict_proba(X)[:, 1]
    mu0 = m0.predict_proba(X)[:, 1]

    ps_m = LogisticRegression(max_iter=1000, penalty="l2", C=1.0, solver="lbfgs")
    ps_m.fit(X, y)
    ps = np.clip(ps_m.predict_proba(X)[:, 1], 0.01, 0.99)

    # AIPW means (standardized to the full sample)
    a1 = np.mean(mu1 + y * (died - mu1) / ps)
    a0 = np.mean(mu0 + (1 - y) * (died - mu0) / (1 - ps))
    return a1 - a0, a1, a0


def e_value_rr(rr):
    """E-value on the RR scale for a protective effect (rr < 1)."""
    r = 1.0 / rr  # convert to harmful scale > 1
    return r + np.sqrt(r * (r - 1))


def boot_ci_generic(stat_fn, n_boot, seed):
    rng = np.random.RandomState(seed)
    stats = []
    for _ in range(n_boot):
        try:
            s = stat_fn(rng)
            if s is not None and np.isfinite(s):
                stats.append(s)
        except Exception:
            continue
    stats = np.array(stats)
    if len(stats) < n_boot * 0.9:
        return None
    return (float(np.percentile(stats, 2.5)), float(np.percentile(stats, 97.5)))


# ── main ─────────────────────────────────────────────────────────────────
def main():
    print("=" * 70)
    print("Robustness package (SA4/SA5/AIPW/E-value/subgroups)")
    print("=" * 70)

    # Rebuild the locked analytic dataset
    df = ra.load_excel()
    eligible, flow = ra.screen_eligibility(df)
    eligible = ra.classify_strategies(eligible)
    eligible = ra.compute_outcomes(eligible)
    eligible, base_cols, base_cov = ra.build_covariates(eligible)

    analyzable = eligible[eligible["strategy"].isin(["early", "deferred"])].copy()
    y = (analyzable["strategy"] == "early").astype(int).values
    died = analyzable["died_28d"].values.astype(int)
    n = len(analyzable)
    print(f"Analyzable n = {n} (early {y.sum()}, deferred {n - y.sum()})")

    results = {}

    # ── SA4: extended covariate adjustment ───────────────────────────────
    print("\n[1/5] SA4: extended covariate adjustment...")
    ext, ext_cols, ext_cov = build_extended_covariates(eligible, base_cols)
    X_ext = ext_cov.loc[analyzable.index].values
    w_ext, _ = fit_ps_weights(X_ext, y, trunc=(1, 99))
    rd4, me4, md4 = weighted_rd(y, died, w_ext)
    # balance check on extended set
    bal = []
    for j, c in enumerate(ext_cols):
        x = X_ext[:, j]
        we, wd = w_ext[y == 1], w_ext[y == 0]
        mew = np.sum(we * x[y == 1]) / np.sum(we)
        mdw = np.sum(wd * x[y == 0]) / np.sum(wd)
        vw = np.sum(we * (x[y == 1] - mew) ** 2) / np.sum(we)
        vd = np.sum(wd * (x[y == 0] - mdw) ** 2) / np.sum(wd)
        denom = np.sqrt((vw + vd) / 2)
        bal.append(abs(mew - mdw) / denom if denom > 0 else 0.0)
    smd4_max = float(np.max(bal))
    smd4_gt10 = int(np.sum(np.array(bal) > 0.10))

    idx_pos = np.arange(n)
    def stat_sa4(rng):
        be = rng.choice(idx_pos[y == 1], size=int(y.sum()), replace=True)
        bd = rng.choice(idx_pos[y == 0], size=int(n - y.sum()), replace=True)
        bidx = np.concatenate([be, bd])
        try:
            wb, _ = fit_ps_weights(X_ext[bidx], y[bidx], trunc=(1, 99))
        except Exception:
            return None
        r, _, _ = weighted_rd(y[bidx], died[bidx], wb)
        return r
    ci4 = boot_ci_generic(stat_sa4, N_BOOT_SA, ra.RNG_SEED + 11)
    results["SA4_extended_covariates"] = {
        "n_covariates": len(ext_cols),
        "rd": round(float(rd4), 4), "mort_early": round(float(me4), 4),
        "mort_deferred": round(float(md4), 4),
        "rd_ci_lo": round(ci4[0], 4), "rd_ci_hi": round(ci4[1], 4),
        "smd_max": round(smd4_max, 4), "smd_gt_0_10": smd4_gt10,
    }
    print(f"  RD = {rd4*100:.1f} pp [{ci4[0]*100:.1f}, {ci4[1]*100:.1f}], "
          f"max SMD = {smd4_max:.3f} ({len(ext_cols)} covariates)")

    # ── SA5: 5th-95th percentile truncation ──────────────────────────────
    print("\n[2/5] SA5: alternative truncation (5-95th pct)...")
    X_base = base_cov.loc[analyzable.index].values
    w5, _ = fit_ps_weights(X_base, y, trunc=(5, 95))
    rd5, me5, md5 = weighted_rd(y, died, w5)
    def stat_sa5(rng):
        be = rng.choice(idx_pos[y == 1], size=int(y.sum()), replace=True)
        bd = rng.choice(idx_pos[y == 0], size=int(n - y.sum()), replace=True)
        bidx = np.concatenate([be, bd])
        try:
            wb, _ = fit_ps_weights(X_base[bidx], y[bidx], trunc=(5, 95))
        except Exception:
            return None
        r, _, _ = weighted_rd(y[bidx], died[bidx], wb)
        return r
    ci5 = boot_ci_generic(stat_sa5, N_BOOT_SA, ra.RNG_SEED + 12)
    results["SA5_alt_truncation"] = {
        "rd": round(float(rd5), 4),
        "rd_ci_lo": round(ci5[0], 4), "rd_ci_hi": round(ci5[1], 4),
    }
    print(f"  RD = {rd5*100:.1f} pp [{ci5[0]*100:.1f}, {ci5[1]*100:.1f}]")

    # ── AIPW doubly robust ───────────────────────────────────────────────
    print("\n[3/5] AIPW doubly robust estimator (base covariates)...")
    a_rd, a1, a0 = aipw_rd(X_base, y, died)
    def stat_aipw(rng):
        be = rng.choice(idx_pos[y == 1], size=int(y.sum()), replace=True)
        bd = rng.choice(idx_pos[y == 0], size=int(n - y.sum()), replace=True)
        bidx = np.concatenate([be, bd])
        try:
            r, _, _ = aipw_rd(X_base[bidx], y[bidx], died[bidx])
        except Exception:
            return None
        return r
    ci_a = boot_ci_generic(stat_aipw, N_BOOT_SA, ra.RNG_SEED + 13)
    results["AIPW_doubly_robust"] = {
        "rd": round(float(a_rd), 4),
        "mort_early": round(float(a1), 4), "mort_deferred": round(float(a0), 4),
        "rd_ci_lo": round(ci_a[0], 4), "rd_ci_hi": round(ci_a[1], 4),
    }
    print(f"  RD = {a_rd*100:.1f} pp [{ci_a[0]*100:.1f}, {ci_a[1]*100:.1f}]")

    # ── E-value from locked primary RR ───────────────────────────────────
    print("\n[4/5] E-value...")
    with open(OUT / "analysis_summary.json") as f:
        locked = json.load(f)
    P = locked["primary"]
    rr = P["rr"]; rr_lo = P["rr_ci_hi"]  # CI bound closest to null for protective RR
    ev_point = e_value_rr(rr)
    ev_ci = e_value_rr(rr_lo)
    results["e_value"] = {
        "rr": rr, "e_value_point": round(float(ev_point), 2),
        "rr_ci_bound_nearest_null": rr_lo,
        "e_value_ci_bound": round(float(ev_ci), 2),
    }
    print(f"  RR = {rr:.2f} -> E-value = {ev_point:.2f} (CI bound {ev_ci:.2f})")

    # ── Subgroup consistency ─────────────────────────────────────────────
    print("\n[5/5] Subgroup analyses...")
    sofa_med = float(np.median(base_cov.loc[analyzable.index, "sofa_total"]))
    sub_defs = {
        "Day-2 mode: controlled": analyzable["d2_mode_class"] == "controlled",
        "Day-2 mode: assisted": analyzable["d2_mode_class"] == "assisted",
        "Age >= 65": base_cov.loc[analyzable.index, "age_num"] >= 65,
        "Age < 65": base_cov.loc[analyzable.index, "age_num"] < 65,
        "SOFA >= median": base_cov.loc[analyzable.index, "sofa_total"] >= sofa_med,
        "SOFA < median": base_cov.loc[analyzable.index, "sofa_total"] < sofa_med,
        "Sepsis-3": base_cov.loc[analyzable.index, "sepsis3_flag"] == 1,
        "No sepsis": base_cov.loc[analyzable.index, "sepsis3_flag"] == 0,
        "ARDS": base_cov.loc[analyzable.index, "ards_flag"] == 1,
        "No ARDS": base_cov.loc[analyzable.index, "ards_flag"] == 0,
    }
    sub_rows = []
    for name, mask in sub_defs.items():
        m = mask.values.astype(bool)
        if m.sum() < 50 or y[m].sum() < 20 or (1 - y[m]).sum() < 20:
            print(f"  {name}: too small (n={m.sum()}), skipped")
            continue
        Xs, ys, ds = X_base[m], y[m], died[m]
        try:
            ws, _ = fit_ps_weights(Xs, ys, trunc=(1, 99))
        except Exception:
            continue
        rds, me_s, md_s = weighted_rd(ys, ds, ws)
        pos = np.arange(m.sum())
        def stat_sub(rng, Xs=Xs, ys=ys, ds=ds, pos=pos):
            be = rng.choice(pos[ys == 1], size=int(ys.sum()), replace=True)
            bd = rng.choice(pos[ys == 0], size=int(len(ys) - ys.sum()), replace=True)
            bidx = np.concatenate([be, bd])
            try:
                wb, _ = fit_ps_weights(Xs[bidx], ys[bidx], trunc=(1, 99))
            except Exception:
                return None
            r, _, _ = weighted_rd(ys[bidx], ds[bidx], wb)
            return r
        ci_s = boot_ci_generic(stat_sub, N_BOOT_SUB, ra.RNG_SEED + 20)
        sub_rows.append({
            "subgroup": name, "n": int(m.sum()), "n_early": int(ys.sum()),
            "mort_early": round(float(me_s), 4), "mort_deferred": round(float(md_s), 4),
            "rd": round(float(rds), 4),
            "rd_ci_lo": round(ci_s[0], 4) if ci_s else np.nan,
            "rd_ci_hi": round(ci_s[1], 4) if ci_s else np.nan,
        })
        print(f"  {name}: n={m.sum()}, RD = {rds*100:.1f} pp "
              f"[{ci_s[0]*100:.1f}, {ci_s[1]*100:.1f}]" if ci_s else f"  {name}: CI failed")
    sub_df = pd.DataFrame(sub_rows)
    sub_df.to_csv(OUT / "subgroup_results.csv", index=False)

    results["subgroups"] = sub_rows

    with open(OUT / "robustness_results.json", "w") as f:
        json.dump(results, f, indent=2)
    print("\nSaved:", OUT / "robustness_results.json")
    print("Saved:", OUT / "subgroup_results.csv")
    print("=" * 70)
    return results


if __name__ == "__main__":
    main()
