#!/usr/bin/env python3
"""
Ventilation Trajectory for TTE De-escalation Manuscript (Figure S4)

Builds the "per-protocol verification" figure showing that the early vs
deferred strategies were actually delivered as assigned in the post-landmark
72 hours. Uses ONLY columns present in the work Excel — no fabricated data.

Time resolution (per Option A, "use existing data, be honest"):
  - 6-hourly vital snapshots, Day 1 00:00 → Day 2 18:00 (8 timepoints)
      Respiratory parameters: RR, SpO2, P/F ratio, lactate, norepinephrine, RASS
  - Daily ventilator aggregate, Day 2 and Day 3 (2 timepoints)
      Mode (categorical), FiO2, PEEP, peak/plateau pressure, tidal volume

NOT included (Excel does not contain):
  - Hourly minute-by-minute FiO2/PEEP/mode (raw chartevents not abstracted)
  - Day 4 / Day 5 ventilator settings
  - Waveform data (MIMIC-IV main DB has no 100+ Hz waveform)
  - Diaphragm EMG / ultrasound
  - Patient-ventilator synchrony indices

This script is an addition to the locked run_analysis.py pipeline and only
deals with the analyzable cohort (1,902 patients: 776 early / 1,126 deferred).
"""
import json, os, sys, warnings
from pathlib import Path

import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches

warnings.filterwarnings("ignore")

EXCEL_PATH = "/Users/zhangrui/Documents/sofa2.0/副本minmic数据.xlsx"
OUT_DIR    = Path("/Users/zhangrui/Desktop/tte/TTE_Deescalation_v2/results")
FIG_DIR    = Path("/Users/zhangrui/Desktop/tte/TTE_Deescalation_v2/figures")
OUT_DIR.mkdir(parents=True, exist_ok=True)
FIG_DIR.mkdir(parents=True, exist_ok=True)

# ─────────────────────────────────────────────────────────────────────────
# Reuse screen_eligibility & classify_strategies from run_analysis
# ─────────────────────────────────────────────────────────────────────────
sys.path.insert(0, str(Path(__file__).parent))
from run_analysis import (
    load_excel, screen_eligibility, classify_strategies,
    build_covariates, CONTROLLED_KEYWORDS, ASSISTED_KEYWORDS
)

# ─────────────────────────────────────────────────────────────────────────
# 6-hourly vital trajectory columns (Day 1 → Day 2 18:00)
# ─────────────────────────────────────────────────────────────────────────
TIME_LABELS_6H = [
    "Day 1\n00:00", "Day 1\n06:00", "Day 1\n12:00", "Day 1\n18:00",
    "Day 2\n00:00", "Day 2\n06:00", "Day 2\n12:00", "Day 2\n18:00",
]
TIME_HOURS_6H  = [0, 6, 12, 18, 24, 30, 36, 42]   # hours from start

# Map timepoint index (0..7) → column substring suffix
TIME_COL_SUFFIX_6H = [
    "Day1_00:00", "Day1_06:00", "Day1_12:00", "Day1_18:00",
    "Day2_00:00", "Day2_06:00", "Day2_12:00", "Day2_18:00",
]

# Variables to extract (only respiratory-relevant ones)
VITAL_VARS = [
    ("呼吸频率(次/min)",         "RR",                  "次/min"),
    ("脉搏氧饱和度(%)",           "SpO2",                "%"),
    ("氧合指数_PaO2/FiO2",        "P/F ratio",           "mmHg"),
    ("乳酸_Lactate(mmol/L)",      "Lactate",             "mmol/L"),
    ("持续注射泵入(去甲肾上腺素)(ug/kg/min)", "Norepinephrine", "μg/kg/min"),
    ("RASS评分",                  "RASS",                "score"),
]

# ─────────────────────────────────────────────────────────────────────────
# Daily vent trajectory (Day 2 + Day 3)
# ─────────────────────────────────────────────────────────────────────────
VENT_VARS = [
    ("氧浓度(%)_Day2",        "FiO2_Day2",   "%"),
    ("氧浓度(%)_Day3",        "FiO2_Day3",   "%"),
    ("呼气末正压(cmH2O)_Day2", "PEEP_Day2",   "cmH2O"),
    ("呼气末正压(cmH2O)_Day3", "PEEP_Day3",   "cmH2O"),
    ("气道峰压(cmH2O)_Day2",    "PeakP_Day2",  "cmH2O"),
    ("气道峰压(cmH2O)_Day3",    "PeakP_Day3",  "cmH2O"),
    ("气道平台压(cmH2O)_Day2",  "Pplat_Day2",  "cmH2O"),
    ("气道平台压(cmH2O)_Day3",  "Pplat_Day3",  "cmH2O"),
    ("呼气潮气量(ml)_Day2",      "TV_Day2",     "mL"),
    ("呼气潮气量(ml)_Day3",      "TV_Day3",     "mL"),
]

MODE_DAY2 = "呼吸模式_Day2"
MODE_DAY3 = "呼吸模式_Day3"


def find_column(df, substring):
    """Return the actual column name(s) in df whose name contains substring.
       Returns the first match if multiple (Day_00:00 etc. are handled separately).
    """
    matches = [c for c in df.columns if substring in str(c)]
    return matches


def extract_vitals_6h(df):
    """Build a long DataFrame: stay_id × timepoint × variable × value."""
    rows = []
    for ti, suffix in enumerate(TIME_COL_SUFFIX_6H):
        for prefix, label, unit in VITAL_VARS:
            col_full = f"{prefix}_{suffix}"
            if col_full not in df.columns:
                continue
            tmp = df[["stay_id", col_full]].copy()
            tmp["timepoint"]  = ti
            tmp["t_label"]    = TIME_LABELS_6H[ti]
            tmp["t_hours"]    = TIME_HOURS_6H[ti]
            tmp["variable"]   = label
            tmp["unit"]       = unit
            tmp["value"]      = pd.to_numeric(tmp[col_full], errors="coerce")
            tmp = tmp.dropna(subset=["value"])
            rows.append(tmp[["stay_id", "timepoint", "t_label", "t_hours",
                             "variable", "unit", "value"]])
    if not rows:
        return pd.DataFrame()
    return pd.concat(rows, ignore_index=True)


def extract_vent_daily(df):
    """Build a long DataFrame for Day 2 and Day 3 vent aggregates."""
    rows = []
    for col, label, unit in VENT_VARS:
        if col not in df.columns:
            continue
        day = "Day3" if "Day3" in col else "Day2"
        tmp = df[["stay_id", col]].copy()
        tmp["day"]      = day
        tmp["variable"] = label
        tmp["unit"]     = unit
        tmp["value"]    = pd.to_numeric(tmp[col], errors="coerce")
        tmp = tmp.dropna(subset=["value"])
        rows.append(tmp[["stay_id", "day", "variable", "unit", "value"]])
    if not rows:
        return pd.DataFrame()
    return pd.concat(rows, ignore_index=True)


def classify_mode(mode_str):
    """Same classifier as run_analysis.classify_mode."""
    if mode_str is None or (isinstance(mode_str, float) and pd.isna(mode_str)):
        return None
    s = str(mode_str).strip()
    if s in ("NULL", "nan", ""):
        return None
    for kw in CONTROLLED_KEYWORDS:
        if kw in s:
            return "controlled"
    for kw in ASSISTED_KEYWORDS:
        if kw in s:
            return "assisted"
    return "unknown"


def main():
    print("=" * 70)
    print("Ventilation Trajectory Script — Per-Protocol Verification (Fig S4)")
    print("=" * 70)

    # 1. Load + screen
    print("\n[1] Loading Excel and reconstructing analysable cohort...")
    df = load_excel()
    eligible, flow = screen_eligibility(df)
    print(f"  eligible: {len(eligible)}")
    analyzable = classify_strategies(eligible)
    analyzable = analyzable[analyzable["strategy"].isin(["early", "deferred"])].copy()
    analyzable = analyzable.reset_index(drop=True)
    n_e = int((analyzable["strategy"] == "early").sum())
    n_d = int((analyzable["strategy"] == "deferred").sum())
    print(f"  analyzable: {len(analyzable)}  (early={n_e}, deferred={n_d})")

    # 2. Extract 6-h vitals (filter to analyzable)
    print("\n[2] Extracting 6-hourly respiratory vitals (D1 → D2 18:00)...")
    df_anal = analyzable.copy()
    df_anal = df_anal[["stay_id", "strategy"]].merge(df, on="stay_id", how="left")
    vital_long = extract_vitals_6h(df_anal)
    vital_long = vital_long.merge(df_anal[["stay_id", "strategy"]], on="stay_id", how="left")
    print(f"  vital_long rows: {len(vital_long)}")
    print(f"  variables: {sorted(vital_long['variable'].unique().tolist())}")
    print(f"  timepoints: {sorted(vital_long['t_label'].unique().tolist())}")

    # 3. Extract daily vent D2 + D3
    print("\n[3] Extracting daily vent aggregates (D2 + D3)...")
    vent_long = extract_vent_daily(df_anal)
    vent_long = vent_long.merge(df_anal[["stay_id", "strategy"]], on="stay_id", how="left")
    print(f"  vent_long rows: {len(vent_long)}")

    # 4. Per-strategy summaries
    print("\n[4] Computing per-strategy summaries...")
    vital_summary = (vital_long
                     .groupby(["strategy", "variable", "t_label", "t_hours"])
                     .agg(n=("value", "count"),
                          median=("value", "median"),
                          q1=("value", lambda x: np.nanpercentile(x, 25)),
                          q3=("value", lambda x: np.nanpercentile(x, 75)),
                          mean=("value", "mean"))
                     .reset_index())
    vent_summary = (vent_long
                    .groupby(["strategy", "variable", "day"])
                    .agg(n=("value", "count"),
                         median=("value", "median"),
                         q1=("value", lambda x: np.nanpercentile(x, 25)),
                         q3=("value", lambda x: np.nanpercentile(x, 75)),
                         mean=("value", "mean"))
                    .reset_index())

    # 5. D2 → D3 deltas (per-protocol metrics)
    print("\n[5] Computing Day2 → Day3 deltas (per-protocol)...")
    deltas = []
    for strat in ["early", "deferred"]:
        sub_e = analyzable[analyzable["strategy"] == strat]
        d2_fio2 = pd.to_numeric(sub_e["氧浓度(%)_Day2"], errors="coerce").dropna()
        d3_fio2 = pd.to_numeric(sub_e["氧浓度(%)_Day3"], errors="coerce").dropna()
        d2_peep = pd.to_numeric(sub_e["呼气末正压(cmH2O)_Day2"], errors="coerce").dropna()
        d3_peep = pd.to_numeric(sub_e["呼气末正压(cmH2O)_Day3"], errors="coerce").dropna()
        d2_pip = pd.to_numeric(sub_e["气道峰压(cmH2O)_Day2"], errors="coerce").dropna()
        d3_pip = pd.to_numeric(sub_e["气道峰压(cmH2O)_Day3"], errors="coerce").dropna()
        d2_pplat = pd.to_numeric(sub_e["气道平台压(cmH2O)_Day2"], errors="coerce").dropna()
        d3_pplat = pd.to_numeric(sub_e["气道平台压(cmH2O)_Day3"], errors="coerce").dropna()
        d2_tv = pd.to_numeric(sub_e["呼气潮气量(ml)_Day2"], errors="coerce").dropna()
        d3_tv = pd.to_numeric(sub_e["呼气潮气量(ml)_Day3"], errors="coerce").dropna()

        deltas.append({
            "strategy": strat,
            "FiO2_Day2_median":   round(float(d2_fio2.median()), 2),
            "FiO2_Day3_median":   round(float(d3_fio2.median()), 2),
            "FiO2_delta_D3_minus_D2": round(float(d3_fio2.median() - d2_fio2.median()), 2),
            "PEEP_Day2_median":   round(float(d2_peep.median()), 2),
            "PEEP_Day3_median":   round(float(d3_peep.median()), 2),
            "PEEP_delta":         round(float(d3_peep.median() - d2_peep.median()), 2),
            "PeakP_Day2_median":  round(float(d2_pip.median()), 2),
            "PeakP_Day3_median":  round(float(d3_pip.median()), 2),
            "PeakP_delta":        round(float(d3_pip.median() - d2_pip.median()), 2),
            "Pplat_Day2_median":  round(float(d2_pplat.median()), 2),
            "Pplat_Day3_median":  round(float(d3_pplat.median()), 2),
            "Pplat_delta":        round(float(d3_pplat.median() - d2_pplat.median()), 2),
            "TV_Day2_median":     round(float(d2_tv.median()), 0),
            "TV_Day3_median":     round(float(d3_tv.median()), 0),
            "TV_delta":           round(float(d3_tv.median() - d2_tv.median()), 0),
            "n_FiO2_Day2":        int(len(d2_fio2)),
            "n_FiO2_Day3":        int(len(d3_fio2)),
            "n_Peep_Day2":        int(len(d2_peep)),
            "n_Peep_Day3":        int(len(d3_peep)),
        })
    delta_df = pd.DataFrame(deltas)

    # 6. Mode distribution D2 vs D3 (per group, as % of analyzable in the group)
    print("\n[6] Mode distribution at Day 2 vs Day 3...")
    def mode_dist(strat):
        sub = analyzable[analyzable["strategy"] == strat]
        # value_counts with dropna=False returns NaN-keyed entries for None
        d2_dict = pd.Series([classify_mode(v) for v in sub[MODE_DAY2]]).value_counts(dropna=False).to_dict()
        d3_dict = pd.Series([classify_mode(v) for v in sub[MODE_DAY3]]).value_counts(dropna=False).to_dict()
        # Convert NaN key to "missing" so we can look it up reliably
        def get_count(d, key):
            if key == "missing":
                # sum NaN-keyed + (any None-keyed, which pandas doesn't usually preserve)
                return int(sum(v for k, v in d.items()
                               if k is None or (isinstance(k, float) and np.isnan(k))))
            return int(d.get(key, 0))
        out = {}
        n_grp = len(sub)
        for cat in ["controlled", "assisted", "unknown", "missing"]:
            out[f"D2_{cat}"]  = get_count(d2_dict, cat)
            out[f"D3_{cat}"]  = get_count(d3_dict, cat)
            out[f"D2pct_{cat}"] = round(100 * out[f"D2_{cat}"] / n_grp, 1)
            out[f"D3pct_{cat}"] = round(100 * out[f"D3_{cat}"] / n_grp, 1)
        out["n_group"] = n_grp
        return out
    mode_summary = {
        "early":    mode_dist("early"),
        "deferred": mode_dist("deferred"),
    }
    print(f"  early: D2 controlled={mode_summary['early']['D2_controlled']} "
          f"({mode_summary['early']['D2pct_controlled']}%) → "
          f"D3 controlled={mode_summary['early']['D3_controlled']} "
          f"({mode_summary['early']['D3pct_controlled']}%), "
          f"D3 extubated={mode_summary['early']['D3_missing']} "
          f"({mode_summary['early']['D3pct_missing']}%)")
    print(f"  deferred: D2 controlled={mode_summary['deferred']['D2_controlled']} "
          f"({mode_summary['deferred']['D2pct_controlled']}%) → "
          f"D3 controlled={mode_summary['deferred']['D3_controlled']} "
          f"({mode_summary['deferred']['D3pct_controlled']}%), "
          f"D3 extubated={mode_summary['deferred']['D3_missing']} "
          f"({mode_summary['deferred']['D3pct_missing']}%)")

    # 7. Save summaries
    summary = {
        "time_resolution": {
            "vitals_6h": "Day 1 00:00 → Day 2 18:00 (8 timepoints)",
            "vent_daily": "Day 2 + Day 3 (2 timepoints, daily aggregate)",
            "no_waveform": "MIMIC-IV main DB has no 100+ Hz waveform",
            "no_day4_day5": "Excel does not contain Day 4/5 vent settings",
            "no_diaphragm": "no EMG / diaphragm ultrasound available",
            "no_synchrony": "no waveform-derived synchrony indices available",
        },
        "n_analyzable": len(analyzable),
        "n_early": n_e,
        "n_deferred": n_d,
        "deltas_D3_minus_D2_by_group": delta_df.to_dict("records"),
        "mode_distribution": mode_summary,
    }
    with open(OUT_DIR / "ventilation_trajectory.json", "w") as f:
        json.dump(summary, f, indent=2, default=str)
    print("\n  ✓ Saved results/ventilation_trajectory.json")

    # Save summaries as CSV for supplement table
    vital_summary.to_csv(OUT_DIR / "vital_trajectory_summary.csv", index=False)
    vent_summary.to_csv(OUT_DIR / "vent_daily_trajectory_summary.csv", index=False)
    delta_df.to_csv(OUT_DIR / "vent_daily_deltas.csv", index=False)
    print("  ✓ Saved CSVs")

    # 8. Plot Figure S4 (a single 2-panel figure)
    print("\n[7] Plotting Figure S4...")

    plt.rcParams.update({
        "font.size": 9, "font.family": "sans-serif",
        "axes.spines.top": False, "axes.spines.right": False,
        "axes.linewidth": 0.8, "xtick.major.width": 0.8, "ytick.major.width": 0.8,
        "figure.dpi": 300,
    })

    # — Panel A: 6-h vitals (P/F ratio + RR + SpO2)
    fig, axes = plt.subplots(2, 2, figsize=(11, 8))

    color_e = "#1F77B4"   # muted blue
    color_d = "#D6604C"   # muted red

    subplot_specs = [
        (axes[0, 0], "P/F ratio",     "mmHg"),
        (axes[0, 1], "SpO2",          "%"),
        (axes[1, 0], "RR",            "次/min"),
        (axes[1, 1], "Norepinephrine", "μg/kg/min"),
    ]
    for ax, var, ylabel in subplot_specs:
        for strat, color, marker in [
            ("early",    color_e, "o"),
            ("deferred", color_d, "s"),
        ]:
            s = vital_summary[vital_summary["variable"] == var]
            s = s[s["strategy"] == strat].sort_values("t_hours")
            ax.plot(s["t_hours"], s["median"], marker=marker, color=color,
                    linewidth=2, markersize=5, label=("Early" if strat == "early" else "Deferred"))
            ax.fill_between(s["t_hours"], s["q1"], s["q3"],
                             color=color, alpha=0.12, linewidth=0)
        ax.set_ylabel(ylabel, fontsize=9)
        ax.set_xticks(TIME_HOURS_6H)
        ax.set_xticklabels(TIME_LABELS_6H, fontsize=7)
        ax.set_title(var, fontsize=10, loc="left")
        ax.grid(True, axis="y", linestyle=":", alpha=0.3)
        # Only legend on the first plot
        if var == "P/F ratio":
            ax.legend(loc="upper right", fontsize=8, frameon=False)

    axes[0, 0].axvline(24, color="#666", linestyle="--", alpha=0.4, linewidth=0.7)
    axes[0, 0].text(24.3, axes[0, 0].get_ylim()[1]*0.96, "Day 2\nstart",
                    fontsize=7, color="#666", alpha=0.7, va="top")
    fig.suptitle("Figure S4A: 6-hourly respiratory trajectory, Day 1 → Day 2 18:00 (n=1,902)",
                 fontsize=11, y=1.00)
    fig.tight_layout()
    fig.savefig(FIG_DIR / "FigureS4A_VitalsTrajectory.png", dpi=300, bbox_inches="tight")
    plt.close(fig)
    print("  ✓ FigureS4A_VitalsTrajectory.png")

    # — Panel B: Daily vent D2 → D3 (FiO2, PEEP) + stacked mode bars
    fig, axes = plt.subplots(1, 3, figsize=(12, 4.2))

    # FiO2 line panel
    ax = axes[0]
    for strat, color in [("early", color_e), ("deferred", color_d)]:
        s = vent_summary[(vent_summary["strategy"] == strat) &
                          (vent_summary["variable"].isin(["FiO2_Day2", "FiO2_Day3"]))]
        x = [0, 1]
        y = [s[s["variable"]=="FiO2_Day2"]["median"].values[0],
             s[s["variable"]=="FiO2_Day3"]["median"].values[0]]
        ax.plot(x, y, "-o", color=color, linewidth=2, markersize=7,
                label=("Early" if strat == "early" else "Deferred"))
    ax.set_xticks([0, 1]); ax.set_xticklabels(["Day 2", "Day 3"])
    ax.set_ylabel("FiO₂, %", fontsize=10)
    ax.set_title("FiO₂", fontsize=10, loc="left")
    ax.legend(fontsize=9, frameon=False, loc="lower left")
    ax.set_xlim(-0.3, 1.3)
    ax.grid(True, axis="y", linestyle=":", alpha=0.3)

    # PEEP line panel
    ax = axes[1]
    for strat, color in [("early", color_e), ("deferred", color_d)]:
        s = vent_summary[(vent_summary["strategy"] == strat) &
                          (vent_summary["variable"].isin(["PEEP_Day2", "PEEP_Day3"]))]
        x = [0, 1]
        y = [s[s["variable"]=="PEEP_Day2"]["median"].values[0],
             s[s["variable"]=="PEEP_Day3"]["median"].values[0]]
        ax.plot(x, y, "-o", color=color, linewidth=2, markersize=7,
                label=("Early" if strat == "early" else "Deferred"))
    ax.set_xticks([0, 1]); ax.set_xticklabels(["Day 2", "Day 3"])
    ax.set_ylabel("PEEP, cmH₂O", fontsize=10)
    ax.set_title("PEEP", fontsize=10, loc="left")
    ax.legend(fontset=8, frameon=False, loc="lower left") if False else ax.legend(fontsize=8, frameon=False, loc="lower left")
    ax.set_xlim(-0.3, 1.3)
    ax.grid(True, axis="y", linestyle=":", alpha=0.3)

    # Mode distribution stacked bar panel
    ax = axes[2]
    cats = ["controlled", "assisted", "unknown", "missing"]
    cat_labels = ["Controlled", "Assisted", "Unknown", "Extubated/\nno mode"]
    cat_colors = ["#8C564B", "#9467BD", "#B0B0B0", "#2CA02C"]
    bar_width = 0.32
    x_pos = np.array([0, 1])  # Day 2, Day 3
    for gi, strat in enumerate(["early", "deferred"]):
        ms = mode_summary[strat]
        bottom = np.zeros(2)  # day 2, day 3
        days = ["D2", "D3"]
        offsets = [-bar_width/2, bar_width/2]
        for ci, cat in enumerate(cats):
            vals = [ms[f"{d}_{cat}"] for d in days]
            ax.bar(x_pos + offsets[gi], vals, bar_width, bottom=bottom,
                   color=cat_colors[ci], edgecolor="white", linewidth=0.5,
                   label=cat_labels[ci] if gi == 0 else None)
            bottom = bottom + np.array(vals)
    ax.set_xticks(x_pos)
    ax.set_xticklabels(["Day 2", "Day 3"])
    ax.set_ylabel("Count", fontsize=10)
    ax.set_title("Ventilator mode", fontsize=10, loc="left")
    _yl = ax.get_ylim()
    ax.set_ylim(_yl[0], _yl[1] * 1.38)
    ax.legend(fontsize=8.5, frameon=False, loc="upper left", ncol=1)
    # Add text labels for counts (only for major categories)
    for gi, strat in enumerate(["early", "deferred"]):
        ms = mode_summary[strat]
        for d_i, d in enumerate(["D2", "D3"]):
            ax.text(d_i + (-bar_width/2 if gi==0 else bar_width/2),
                    ms[f"{d}_controlled"] + ms[f"{d}_assisted"] + 5,
                    "n="+str(ms[f"{d}_controlled"]+ms[f"{d}_assisted"]),
                    ha="center", fontsize=7, color="#333")
    ax.set_xlim(-0.5, 1.5)

    fig.suptitle("Figure S4B: Daily ventilator trajectory, Day 2 → Day 3",
                 fontsize=11, y=1.02)
    fig.tight_layout()
    fig.savefig(FIG_DIR / "FigureS4B_VentDaily.png", dpi=300, bbox_inches="tight")
    plt.close(fig)
    print("  ✓ FigureS4B_VentDaily.png")

    # Combined Figure S4 (concatenated vertical, for manuscript)
    fig_combined = plt.figure(figsize=(7.4, 8.8))
    gs = fig_combined.add_gridspec(2, 1, height_ratios=[1.5, 1], hspace=0.35)
    # Top: Panel A grid (2x2)
    gs_top = gs[0].subgridspec(2, 2, hspace=0.45, wspace=0.3)
    axes_top = gs_top.subplots()
    # Map variable to its axes_top slot
    panel_a_axes = {
        "P/F ratio":      axes_top[0, 0],
        "SpO2":            axes_top[0, 1],
        "RR":              axes_top[1, 0],
        "Norepinephrine":  axes_top[1, 1],
    }
    panel_a_ylabels = {
        "P/F ratio": "mmHg",
        "SpO2":      "%",
        "RR":        "breaths/min",
        "Norepinephrine": "μg/kg/min",
    }
    for var, ax in panel_a_axes.items():
        ylabel = panel_a_ylabels[var]
        for strat, color, marker in [
            ("early", color_e, "o"),
            ("deferred", color_d, "s"),
        ]:
            s = vital_summary[vital_summary["variable"] == var]
            s = s[s["strategy"] == strat].sort_values("t_hours")
            ax.plot(s["t_hours"], s["median"], marker=marker, color=color,
                    linewidth=2, markersize=5, label=("Early" if strat == "early" else "Deferred"))
            ax.fill_between(s["t_hours"], s["q1"], s["q3"],
                             color=color, alpha=0.12, linewidth=0)
        ax.set_ylabel(ylabel, fontsize=10)
        ax.set_xticks(TIME_HOURS_6H[::2])
        ax.set_xticklabels([TIME_LABELS_6H[i] for i in range(0, len(TIME_LABELS_6H), 2)],
                           fontsize=8)
        ax.set_title(var, fontsize=10.5, loc="left")
        ax.grid(True, axis="y", linestyle=":", alpha=0.3)
        if var == "P/F ratio":
            ax.legend(loc="upper right", fontsize=9, frameon=False)

    # Day 2 divider
    axes_top[0, 0].axvline(24, color="#666", linestyle="--", alpha=0.4, linewidth=0.7)
    ylim_top = axes_top[0, 0].get_ylim()
    axes_top[0, 0].text(24.3, ylim_top[1] - 0.04 * (ylim_top[1] - ylim_top[0]),
                        "Day 2", fontsize=8, color="#666", alpha=0.85,
                        va="top", ha="left")

    # Bottom: Panel B (3 subplots horizontal)
    gs_bot = gs[1].subgridspec(1, 3, wspace=0.35)
    axes_bot = gs_bot.subplots()

    ax = axes_bot[0]
    for strat, color, dx, ls, mk, va, yo in [
            ("early", color_e, -0.08, "-", "o", "bottom", 10),
            ("deferred", color_d, 0.08, "--", "s", "top", -11)]:
        s = vent_summary[(vent_summary["strategy"] == strat) &
                          (vent_summary["variable"].isin(["FiO2_Day2", "FiO2_Day3"]))]
        x = [0 + dx, 1 + dx]
        y = [s[s["variable"]=="FiO2_Day2"]["median"].values[0],
             s[s["variable"]=="FiO2_Day3"]["median"].values[0]]
        ns = [int(s[s["variable"]=="FiO2_Day2"]["n"].values[0]),
              int(s[s["variable"]=="FiO2_Day3"]["n"].values[0])]
        ax.plot(x, y, ls, marker=mk, color=color, linewidth=2, markersize=7,
                label=("Early" if strat == "early" else "Deferred"))
        for xi, yi, ni in zip(x, y, ns):
            ax.annotate(f"n={ni:,}", (xi, yi), textcoords="offset points",
                        xytext=(0, yo), ha="center", va=va, fontsize=6.5,
                        color=color, clip_on=False, annotation_clip=False)
    ax.set_xticks([0, 1]); ax.set_xticklabels(["Day 2", "Day 3"])
    ax.set_ylabel("FiO₂, %", fontsize=10)
    ax.set_title("FiO₂", fontsize=10, loc="left")
    ax.set_ylim(35, 45)
    ax.legend(fontsize=9, frameon=False, loc="lower left")
    ax.set_xlim(-0.42, 1.42); ax.grid(True, axis="y", linestyle=":", alpha=0.3)

    ax = axes_bot[1]
    for strat, color, dx, ls, mk, va, yo in [
            ("early", color_e, -0.08, "-", "o", "bottom", 10),
            ("deferred", color_d, 0.08, "--", "s", "top", -11)]:
        s = vent_summary[(vent_summary["strategy"] == strat) &
                          (vent_summary["variable"].isin(["PEEP_Day2", "PEEP_Day3"]))]
        x = [0 + dx, 1 + dx]
        y = [s[s["variable"]=="PEEP_Day2"]["median"].values[0],
             s[s["variable"]=="PEEP_Day3"]["median"].values[0]]
        ns = [int(s[s["variable"]=="PEEP_Day2"]["n"].values[0]),
              int(s[s["variable"]=="PEEP_Day3"]["n"].values[0])]
        ax.plot(x, y, ls, marker=mk, color=color, linewidth=2, markersize=7,
                label=("Early" if strat == "early" else "Deferred"))
        for xi, yi, ni in zip(x, y, ns):
            ax.annotate(f"n={ni:,}", (xi, yi), textcoords="offset points",
                        xytext=(0, yo), ha="center", va=va, fontsize=6.5,
                        color=color, clip_on=False, annotation_clip=False)
    ax.set_xticks([0, 1]); ax.set_xticklabels(["Day 2", "Day 3"])
    ax.set_ylabel("PEEP, cmH₂O", fontsize=10)
    ax.set_title("PEEP", fontsize=10, loc="left")
    ax.set_ylim(3, 7)
    ax.legend(fontsize=9, frameon=False, loc="lower left")
    ax.set_xlim(-0.42, 1.42); ax.grid(True, axis="y", linestyle=":", alpha=0.3)

    ax = axes_bot[2]
    for gi, strat in enumerate(["early", "deferred"]):
        ms = mode_summary[strat]
        bottom = np.zeros(2)
        days = ["D2", "D3"]
        offsets = [-bar_width/2, bar_width/2]
        for ci, cat in enumerate(cats):
            vals = [ms[f"{d}_{cat}"] for d in days]
            ax.bar(x_pos + offsets[gi], vals, bar_width, bottom=bottom,
                   color=cat_colors[ci], edgecolor="white", linewidth=0.5,
                   label=cat_labels[ci] if gi == 0 else None)
            bottom = bottom + np.array(vals)
    ax.set_xticks(x_pos)
    ax.set_xticklabels(["Day 2", "Day 3"])
    ax.set_ylabel("Count", fontsize=10)
    ax.set_title("Ventilator mode", fontsize=10, loc="left")
    _yl = ax.get_ylim()
    ax.set_ylim(_yl[0], _yl[1] * 1.95)
    ax.legend(fontsize=7, frameon=False, loc="upper center", ncol=2,
              handlelength=1.0, columnspacing=0.8, labelspacing=0.3,
              borderaxespad=0.1)
    # Label every bar with arm name + total (arms otherwise indistinguishable)
    for gi, strat in enumerate(["early", "deferred"]):
        ms = mode_summary[strat]
        for d_i, d in enumerate(["D2", "D3"]):
            bx = d_i + (-bar_width/2 if gi == 0 else bar_width/2)
            total = sum(ms[f"{d}_{c}"] for c in cats)
            ax.text(bx, total + 20,
                    ("Early\n" if gi == 0 else "Deferred\n") + f"{total:,}",
                    ha="center", va="bottom", fontsize=7.5, color="#333",
                    fontweight="bold")
    ax.set_xlim(-0.5, 1.5)

    fig_combined.canvas.draw()
    _pA = axes_top[0, 0].get_position()
    _pB = axes_bot[0].get_position()
    fig_combined.text(_pA.x0 - 0.055, _pA.y1 + 0.004, "(A)", fontsize=13,
                      fontweight="bold", ha="left", va="bottom")
    fig_combined.text(_pB.x0 - 0.055, _pB.y1 + 0.004, "(B)", fontsize=13,
                      fontweight="bold", ha="left", va="bottom")
    fig_combined.suptitle(
        "Figure S4: Per-protocol verification of de-escalation strategy delivery",
        fontsize=12, y=1.00
    )
    fig_combined.text(0.5, 0.005,
        "(A) 6-hourly respiratory trajectory through Day 2 18:00  |  "
        "(B) Daily ventilator aggregates, Day 2 → Day 3  |  "
        "Median [IQR], n = 1,902.  "
        "Note: FiO₂/PEEP medians coincide between arms at Day 2 (selection criterion) "
        "and Day 3 (driven by FiO₂/PEEP being held in the surviving minority of the early arm "
        "who remained ventilated). Strategy divergence is captured in mode distribution (right panel).",
        ha="center", fontsize=8.5, color="#666", style="italic", wrap=True)
    fig_combined.savefig(FIG_DIR / "FigureS4_Trajectory.png", dpi=300, bbox_inches="tight")
    plt.close(fig_combined)
    print("  ✓ FigureS4_Trajectory.png (combined)")

    # Print final summary stats
    print("\n" + "=" * 70)
    print("D2 → D3 DELTAS (median)")
    print("=" * 70)
    for _, r in delta_df.iterrows():
        print(f"  {r['strategy']:>10}: "
              f"ΔFiO₂ = {r['FiO2_delta_D3_minus_D2']:+.1f}%, "
              f"ΔPEEP = {r['PEEP_delta']:+.1f} cmH₂O, "
              f"ΔPeak = {r['PeakP_delta']:+.1f} cmH₂O, "
              f"ΔPplat = {r['Pplat_delta']:+.1f} cmH₂O")
    print("\nMode distribution D2 → D3 (% of analyzable in group):")
    for strat in ["early", "deferred"]:
        ms = mode_summary[strat]
        total_d3 = ms["D3_controlled"] + ms["D3_assisted"] + ms["D3_unknown"] + ms["D3_missing"]
        print(f"  {strat}: D2 controlled {ms['D2_controlled']}/{total_d3} "
              f"({100*ms['D2_controlled']/(len(analyzable[analyzable['strategy']==strat])):.1f}%) "
              f"→ D3 controlled {ms['D3_controlled']}/{total_d3} "
              f"({ms['D3pct_controlled']}%), "
              f"extubated/no-mode at D3: {ms['D3pct_missing']}%")
    print("\nDONE.")


if __name__ == "__main__":
    main()
