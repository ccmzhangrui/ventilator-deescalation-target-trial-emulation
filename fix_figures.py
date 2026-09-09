"""
Fix overlap issues in figures:
  - Figure 2: row labels overlapping with "N at risk" header
  - Figure 3: legend overlapping with data, "0 (no difference)" overlapping dashed line
  - Figure S1: Max=2.17 arrow overlapping with histogram
"""
from pathlib import Path
import json
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch, Rectangle
from matplotlib.lines import Line2D
from matplotlib import gridspec
from lifelines import KaplanMeierFitter
from lifelines.statistics import logrank_test

import sys
sys.path.insert(0, "/Users/zhangrui/Desktop/tte/TTE_Deescalation_v2")
from run_analysis import (load_excel, screen_eligibility, classify_strategies,
                          compute_outcomes, build_covariates, estimate_weights)

OUT_DIR = Path("/Users/zhangrui/Desktop/tte/TTE_Deescalation_v2/results")
FIG_DIR = Path("/Users/zhangrui/Desktop/tte/TTE_Deescalation_v2/figures")

# Academic palette
C_EARLY = "#C0392B"
C_DEFER = "#2E86AB"
C_TEXT = "#222222"
C_GRID = "#DDDDDD"
C_INCL = "#FFF6E0"
C_EXCL = "#F5E6E6"

plt.rcParams.update({
    "font.family": "DejaVu Sans",
    "font.size": 10,
    "axes.edgecolor": "#333333",
    "axes.labelcolor": C_TEXT,
    "axes.titlecolor": C_TEXT,
    "xtick.color": C_TEXT,
    "ytick.color": C_TEXT,
    "axes.spines.top": False,
    "axes.spines.right": False,
    "savefig.dpi": 350,
    "figure.dpi": 110,
})


def fig1_flow(flow, n_early, n_deferred, n_unasc, n_grace):
    fig, ax = plt.subplots(figsize=(11, 9.5))
    ax.set_xlim(0, 14)
    ax.set_ylim(0, 11)
    ax.axis("off")

    ax.text(7, 10.5, "Figure 1.  Study flow diagram",
            ha="center", va="center", fontsize=14, fontweight="bold", color=C_TEXT)
    ax.text(7, 10.1, "MIMIC-IV v2.2  |  Adults receiving invasive mechanical ventilation",
            ha="center", va="center", fontsize=10, color="#555555", style="italic")

    box_w = 4.0
    box_h = 0.85

    steps = [
        (8.7, "Total MIMIC-IV ICU admissions", "", C_EARLY, C_EARLY, "white",
         f"n = {flow['total_records']:,}"),
        (7.5, "Adults (≥18 years)", "", C_INCL, "#888888", C_TEXT,
         f"n = {flow['adults']:,}"),
        (6.3, "Received invasive mechanical ventilation", "", C_INCL, "#888888", C_TEXT,
         f"n = {flow['ventilated']:,}"),
        (5.1, "MV duration >24 h & alive at 48 h", "", C_INCL, "#888888", C_TEXT,
         f"n = {flow['alive_at_48h']:,}"),
        (3.9, "Day-2 classified mode (controlled or assisted) +",
         "FiO₂ ≤50 % & PEEP ≤10 cmH₂O",
         "#FFE5B4", "#888888", C_TEXT, f"n = {flow['eligible']:,}"),
    ]

    for y, lbl, sub, fc, ec, tc, count in steps:
        if sub:
            txt = f"{lbl}\n{sub}\n{count}"
        else:
            txt = f"{lbl}\n{count}"
        box = FancyBboxPatch((7 - box_w/2, y - box_h/2), box_w, box_h,
                             boxstyle="round,pad=0.04,rounding_size=0.12",
                             facecolor=fc, edgecolor=ec, linewidth=1.2)
        ax.add_patch(box)
        ax.text(7, y, txt, ha="center", va="center", fontsize=9.5,
                color=tc, fontweight="bold" if fc == C_EARLY else "normal")

    for y_top, y_bot in [(8.28, 7.92), (7.08, 6.72), (5.88, 5.52), (4.68, 4.32)]:
        ax.annotate("", xy=(7, y_bot), xytext=(7, y_top),
                    arrowprops=dict(arrowstyle="-|>", color="#444", lw=1.4))

    excl_data = [
        (7.5, f"<18 years\nn = {flow['total_records']-flow['adults']:,}"),
        (6.3, f"No invasive MV\nn = {flow['adults']-flow['ventilated']:,}"),
        (5.1, f"MV ≤24 h or\ndied within 48 h\nn = {flow['ventilated']-flow['alive_at_48h']:,}"),
        (3.9, f"No Day-2 settings,\nnon-controlled mode,\nFiO₂ >50 % or PEEP >10\nn = {flow['alive_at_48h']-flow['eligible']:,}"),
    ]
    for y, txt in excl_data:
        ax.text(11.6, y, f"Excluded:\n{txt}", ha="left", va="center",
                fontsize=8.5, color="#9E2A2A", style="italic",
                bbox=dict(boxstyle="round,pad=0.25", facecolor=C_EXCL,
                          edgecolor="#C0392B", linewidth=0.7))

    for y, _ in excl_data:
        ax.plot([9.05, 11.45], [y, y], color="#C0392B", lw=0.8, ls=":")

    arm_y = 2.0
    arms = [
        (2.5, "Early de-escalation", n_early, C_EARLY, "white"),
        (7.0, "Unascertainable", n_unasc, "#999999", "white"),
        (11.5, "Deferred", n_deferred, C_DEFER, "white"),
    ]
    for x, lbl, n, fc, tc in arms:
        arm_box_w = 3.2
        arm_box_h = 1.0
        box = FancyBboxPatch((x - arm_box_w/2, arm_y - arm_box_h/2), arm_box_w, arm_box_h,
                             boxstyle="round,pad=0.04,rounding_size=0.12",
                             facecolor=fc, edgecolor=fc, linewidth=1.2)
        ax.add_patch(box)
        ax.text(x, arm_y + 0.18, lbl, ha="center", va="center",
                fontsize=10.5, color=tc, fontweight="bold")
        ax.text(x, arm_y - 0.25, f"n = {n}", ha="center", va="center",
                fontsize=10, color=tc)
        if "Unasc" in lbl:
            ax.text(x, arm_y - 0.6, f"(grace deaths: {n_grace})",
                    ha="center", va="center", fontsize=7.5, color="#666",
                    style="italic")

    for x_target in [2.5, 7.0, 11.5]:
        ax.annotate("", xy=(x_target, arm_y + 0.5), xytext=(7, 3.45),
                    arrowprops=dict(arrowstyle="-|>", color="#444", lw=1.2))

    ax.text(7, 0.5,
            f"Analysed cohort: early (n={n_early}) + deferred (n={n_deferred}) = {n_early+n_deferred}; "
            f"unascertainable excluded from primary analysis (grace-period deaths: {n_grace}; "
            f"missing Day-3 settings: {n_unasc - n_grace})",
            ha="center", va="center", fontsize=8, color="#555", style="italic",
            wrap=True)

    plt.savefig(FIG_DIR / "Figure1_Flow.png", dpi=350, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    print("  ✓ Figure 1 (Flow diagram)")


def fig2_survival(analyzable, p_value, primary=None):
    """
    NEJM/Lancet-style weighted Kaplan-Meier figure:
      - Step lines with light 95% CI shading
      - No arrows; clean inline end-of-curve labels
      - Risk-ratio + risk-difference + p-value callout (top-left), numbers
        read from analysis_summary.json so the figure stays in lock-step
        with the manuscript.
      - Number-at-risk table aligned to the x-axis, row labels on the left.
    """
    if primary is None:
        with open(OUT_DIR / "analysis_summary.json") as _f:
            primary = json.load(_f)["primary"]

    fig = plt.figure(figsize=(7.2, 4.6), facecolor="white")
    gs = gridspec.GridSpec(
        2, 1, height_ratios=[2.45, 0.95],
        left=0.20, right=0.965, top=0.90, bottom=0.13, hspace=0.10
    )

    # ── Top: KM curves with 95% CI bands ──
    ax = fig.add_subplot(gs[0])

    n_early = int((analyzable["strategy"] == "early").sum())
    n_deferred = int((analyzable["strategy"] == "deferred").sum())
    mort_early = analyzable[analyzable["strategy"] == "early"]["died_28d"].mean()
    mort_deferred = analyzable[analyzable["strategy"] == "deferred"]["died_28d"].mean()

    for strat_name, color, label in [
        ("early", C_EARLY, "Early de-escalation"),
        ("deferred", C_DEFER, "Deferred"),
    ]:
        sub = analyzable[analyzable["strategy"] == strat_name]
        kmf = KaplanMeierFitter()
        kmf.fit(sub["surv_time"].values, event_observed=sub["event"].values,
                weights=sub["weight_trunc"].values, label=label)
        # suppress the auto xlabel that lifelines otherwise adds ("timeline")
        kmf.plot_survival_function(ax=ax, color=color, ci_show=True,
                                   linewidth=2.2, ci_alpha=0.10)

    # explicit xlabel suppression (lifelines default = "timeline")
    ax.set_xlabel("")
    ax.xaxis.label.set_visible(False)
    ax.tick_params(axis="x", which="both", bottom=False, labelbottom=False)

    # clean inline end-of-curve labels (no arrows, NEJM style)
    y_top, y_bot = 1.0 - mort_early, 1.0 - mort_deferred
    ax.text(28.6, y_top, f"Early: {mort_early*100:.1f}%",
            color=C_EARLY, fontsize=10.5, fontweight="bold",
            ha="left", va="center")
    ax.text(28.6, y_bot, f"Deferred: {mort_deferred*100:.1f}%",
            color=C_DEFER, fontsize=10.5, fontweight="bold",
            ha="left", va="center")

    # stat callout — all numbers come from analysis_summary.json
    p_text = f"p = {p_value:.3f}" if p_value >= 0.001 else "p < 0.001"
    box_lines = [
        f"RR {primary['rr']:.2f}  (95% CI {primary['rr_ci_lo']:.2f}–{primary['rr_ci_hi']:.2f})",
        f"RD {primary['rd']*100:+.1f} pp  (95% CI {primary['rd_ci_lo']*100:+.1f} to {primary['rd_ci_hi']*100:+.1f})",
        f"Log-rank {p_text}",
    ]
    ax.text(0.985, 0.97, "\n".join(box_lines),
            transform=ax.transAxes, fontsize=9, ha="right", va="top",
            linespacing=1.55,
            bbox=dict(boxstyle="round,pad=0.5", facecolor="white",
                      edgecolor="#3F4A5A", linewidth=0.9, alpha=0.95))

    # Lancet-style axes
    for spine in ("top", "right"):
        ax.spines[spine].set_visible(False)
    for spine in ("left", "bottom"):
        ax.spines[spine].set_color("#3F4A5A")
        ax.spines[spine].set_linewidth(0.8)
    ax.set_xlim(0, 28)
    ax.set_ylim(0.80, 1.02)
    ax.set_xticks([0, 7, 14, 21, 28])
    ax.set_yticks([0.80, 0.85, 0.90, 0.95, 1.00])
    ax.set_yticklabels(["0.80", "0.85", "0.90", "0.95", "1.00"])
    ax.tick_params(axis="y", length=0, labelsize=10, colors="#3F4A5A")
    ax.set_ylabel("Survival probability", fontsize=11, fontweight="bold",
                  color="#3F4A5A", labelpad=8)
    ax.set_title("Figure 2.  Weighted Kaplan-Meier survival curves (28-day mortality)",
                 fontsize=12.5, fontweight="bold", color="#1A2230", pad=12,
                 loc="left")
    ax.grid(True, axis="y", alpha=0.18, color=C_GRID, linewidth=0.6)
    legend_handles = [
        Line2D([0], [0], color=C_EARLY, lw=2.2, label="Early de-escalation"),
        Line2D([0], [0], color=C_DEFER, lw=2.2, label="Deferred"),
    ]
    ax.legend(handles=legend_handles, loc="lower left",
              bbox_to_anchor=(0.025, 0.02),
              frameon=True, edgecolor="#3F4A5A", fontsize=9.5, ncol=1)

    # ── Bottom: Number-at-risk table ──
    # Use plain figure coordinates (no subplot) so nothing collides
    # with the top axes or with the row labels at the left.
    time_points = [0, 7, 14, 21, 28]

    def n_at_risk(group_df, t):
        return int((group_df["surv_time"] >= t).sum())

    n_risk_early = [n_at_risk(analyzable[analyzable["strategy"] == "early"], t)
                    for t in time_points]
    n_risk_deferred = [n_at_risk(analyzable[analyzable["strategy"] == "deferred"], t)
                       for t in time_points]

    # x-positions of the time points, mapped from data space to figure space
    # using the same axes geometry as the top subplot
    # We work in axes-fraction coordinates inside a transparent helper axes
    helper = fig.add_subplot(gs[1])
    helper.set_xlim(0, 28)
    helper.set_ylim(-1.0, 1.0)
    helper.axis("off")

    # "No. at risk" label sits high above the first row, well clear of row labels
    helper.text(0, 0.85, "No. at risk", fontsize=10.5, fontweight="bold",
                color="#1A2230", ha="left", va="center")

    # Early row — numbers below the label, with thin tick connectors
    for t, n in zip(time_points, n_risk_early):
        helper.vlines(t, 0.45, 0.65, color=C_EARLY, lw=1.1)
        helper.text(t, 0.20, f"{n:,}", ha="center", va="center",
                    fontsize=11, color=C_EARLY, fontweight="bold")
    # Deferred row
    for t, n in zip(time_points, n_risk_deferred):
        helper.vlines(t, -0.20, -0.40, color=C_DEFER, lw=1.1)
        helper.text(t, -0.55, f"{n:,}", ha="center", va="center",
                    fontsize=11, color=C_DEFER, fontweight="bold")
    # Day tick labels
    for t in time_points:
        helper.text(t, -0.95, str(t), ha="center", va="center",
                    fontsize=10, color="#3F4A5A")
    helper.text(14, -1.30, "Days since landmark", ha="center", va="center",
                fontsize=10.5, color="#3F4A5A", fontweight="bold")

    # Row labels — figure-level so they sit fully outside any subplot
    fig.text(0.130, 0.232, "Early", color=C_EARLY,
             fontsize=10.5, fontweight="bold", ha="right", va="center")
    fig.text(0.130, 0.182, "Deferred", color=C_DEFER,
             fontsize=10.5, fontweight="bold", ha="right", va="center")

    plt.savefig(FIG_DIR / "Figure2_Survival.png", dpi=350,
                bbox_inches="tight", facecolor="white")
    plt.close(fig)
    print("  ✓ Figure 2 (NEJM/Lancet style: step lines + CI bands + end labels, no arrows)")


def fig3_forest(analyzable, summ):
    """
    Fix: legend overlaps with first data point; "0 (no difference)" text overlaps dashed line.
    Solution: legend moved to top-right inside the plot area (above data), and '0' annotation
              placed above x-axis label with adequate spacing.
    """
    fig, ax = plt.subplots(figsize=(7.2, 4.8))

    primary = summ["primary"]
    sa = summ["sensitivity"]

    # Robustness package (SA4/SA5/AIPW) — optional, from robustness_results.json
    rob_path = OUT_DIR / "robustness_results.json"
    rob = json.load(open(rob_path)) if rob_path.exists() else {}

    labels = [
        "Primary (weighted IPCW)",
        "SA1: No weight truncation",
        "SA2: ICU mortality",
        "SA3: Unweighted",
    ]
    rds = [
        primary["rd"] * 100,
        sa["SA1_no_truncation"]["rd"] * 100,
        sa["SA2_icu_mortality"]["rd"] * 100,
        sa["SA3_unweighted"]["rd"] * 100,
    ]
    lo = [
        primary["rd_ci_lo"] * 100,
        sa["SA1_no_truncation"]["rd_ci_lo"] * 100,
        sa["SA2_icu_mortality"]["rd_ci_lo"] * 100,
        sa["SA3_unweighted"]["rd_ci_lo"] * 100,
    ]
    hi = [
        primary["rd_ci_hi"] * 100,
        sa["SA1_no_truncation"]["rd_ci_hi"] * 100,
        sa["SA2_icu_mortality"]["rd_ci_hi"] * 100,
        sa["SA3_unweighted"]["rd_ci_hi"] * 100,
    ]
    if "SA4_extended_covariates" in rob:
        labels.append("SA4: Extended covariates")
        rds.append(rob["SA4_extended_covariates"]["rd"] * 100)
        lo.append(rob["SA4_extended_covariates"]["rd_ci_lo"] * 100)
        hi.append(rob["SA4_extended_covariates"]["rd_ci_hi"] * 100)
    if "SA5_alt_truncation" in rob:
        labels.append("SA5: Truncation 5th\u201395th pct")
        rds.append(rob["SA5_alt_truncation"]["rd"] * 100)
        lo.append(rob["SA5_alt_truncation"]["rd_ci_lo"] * 100)
        hi.append(rob["SA5_alt_truncation"]["rd_ci_hi"] * 100)
    if "AIPW_doubly_robust" in rob:
        labels.append("AIPW doubly robust")
        rds.append(rob["AIPW_doubly_robust"]["rd"] * 100)
        lo.append(rob["AIPW_doubly_robust"]["rd_ci_lo"] * 100)
        hi.append(rob["AIPW_doubly_robust"]["rd_ci_hi"] * 100)

    y_pos = np.arange(len(labels))[::-1]
    xerr_lo = [r - l for r, l in zip(rds, lo)]
    xerr_hi = [h - r for r, h in zip(rds, hi)]

    ax.errorbar(rds, y_pos, xerr=[xerr_lo, xerr_hi],
                fmt="none", ecolor="#444", elinewidth=1.4, capsize=4, capthick=1.4)

    for y, rd, l, h in zip(y_pos, rds, lo, hi):
        ax.scatter(rd, y, s=130, marker="s", color="#2C3E50",
                   edgecolor="white", linewidth=1.2, zorder=5)
        ax.text(h + 0.6, y, f"{rd:+.1f}  [{l:+.1f}, {h:+.1f}]",
                va="center", ha="left", fontsize=9, color="#222",
                fontweight="bold")

    ax.axvline(0, color="#888", linestyle="--", linewidth=0.9, alpha=0.7)

    ax.set_yticks(y_pos)
    ax.set_yticklabels(labels, fontsize=9.5)

    all_vals = lo + hi + rds
    x_lo = min(all_vals) - 4
    x_hi = max(all_vals) + 7
    ax.set_xlim(x_lo, x_hi)
    # Add extra top room for title + legend
    ax.set_ylim(-1.2, len(labels) - 0.2 + 0.4)
    ax.set_xlabel("Risk difference (percentage points)", fontsize=11, fontweight="bold")

    fig.text(0.18, 0.015, "← Favours early", fontsize=9.5, color=C_EARLY,
             fontweight="bold", ha="left", va="bottom")
    fig.text(0.92, 0.015, "Favours deferred →", fontsize=9.5, color=C_DEFER,
             fontweight="bold", ha="right", va="bottom")

    fig.suptitle("Figure 3.  Primary & sensitivity analyses — risk difference (95% CI)",
                 fontsize=12, fontweight="bold", y=0.97)
    ax.grid(True, alpha=0.2, axis="x", color=C_GRID)

    # Legend placed in top-right INSIDE plot area, above the highest row
    legend_elements = [
        Line2D([0], [0], marker="s", color="w", markerfacecolor="#2C3E50",
               markersize=9, label="Point estimate (RD)"),
        Line2D([0], [0], color="#444", lw=1.4, label="95% bootstrap CI"),
    ]
    ax.legend(handles=legend_elements, loc="upper right",
              fontsize=9.5, frameon=True, edgecolor="#444",
              bbox_to_anchor=(0.98, 1.05))

    plt.subplots_adjust(top=0.82, bottom=0.20, left=0.34, right=0.97)
    plt.savefig(FIG_DIR / "Figure3_Forest.png", dpi=350, facecolor="white",
                bbox_inches="tight")
    plt.close(fig)
    print("  ✓ Figure 3 (Forest plot) [legend at top-right inside plot]")


def figS1_weights(analyzable, balance_df, w_diag):
    """
    Fix: Max weight annotation arrow goes outside the plot and overlaps with histogram bars.
    Solution: Place Max annotation to upper-left of the peak (away from bars), or remove the annotation
              and let the legend carry the mean/median info. Keep the max as a text label with no arrow.
    """
    fig, axes = plt.subplots(1, 2, figsize=(7.4, 3.8))

    w = analyzable["weight_trunc"].values
    axes[0].hist(w, bins=40, color=C_DEFER, edgecolor="white",
                 linewidth=0.8, alpha=0.85)
    axes[0].axvline(np.mean(w), color=C_EARLY, linestyle="--", linewidth=1.6,
                    label=f"Mean = {np.mean(w):.2f}")
    axes[0].axvline(np.median(w), color="#555", linestyle=":", linewidth=1.2,
                    label=f"Median = {np.median(w):.2f}")
    # Max annotation - place inside upper area, above the tallest bar, with a small horizontal line marker
    axes[0].axvline(w_diag['weight_max'], color="#444", linestyle="-", linewidth=1.0, alpha=0.7)
    axes[0].annotate(
        f"Max = {w_diag['weight_max']:.2f}",
        xy=(w_diag['weight_max'], axes[0].get_ylim()[1] * 0.55),
        xytext=(0.04, 0.93), textcoords="axes fraction",
        fontsize=9, color="#444", fontweight="bold",
        ha="left", va="top",
        arrowprops=dict(arrowstyle="-", color="#888", lw=0.8, ls="--"))

    axes[0].set_xlabel("Stabilised inverse-probability weight (truncated at 1st–99th percentile)",
                       fontsize=10, fontweight="bold")
    axes[0].set_ylabel("Count (patients)", fontsize=10, fontweight="bold")
    axes[0].set_title("A. Weight distribution", fontsize=11.5, fontweight="bold", loc="left")
    axes[0].legend(fontsize=9, loc="upper left", frameon=True, bbox_to_anchor=(0.02, 0.95))
    axes[0].grid(True, alpha=0.25, color=C_GRID)

    df = balance_df.copy().sort_values("smd_unweighted", ascending=True).reset_index(drop=True)
    y_pos = np.arange(len(df))

    bar_h = 0.38
    axes[1].barh(y_pos + bar_h/2, df["smd_unweighted"], height=bar_h,
                 color="#D98880", alpha=0.85, label="Unweighted")
    axes[1].barh(y_pos - bar_h/2, df["smd_weighted"], height=bar_h,
                 color=C_DEFER, alpha=0.85, label="Weighted (IPCW)")
    axes[1].axvline(0.1, color="#888", linestyle="--", linewidth=1.0)
    axes[1].text(0.105, -1.4, "SMD = 0.10 threshold (acceptable balance)",
                 fontsize=8, color="#888", style="italic", ha="left", va="top")

    axes[1].set_yticks(y_pos)
    axes[1].set_yticklabels(df["covariate"], fontsize=8.5)
    axes[1].set_xlabel("|Standardised mean difference|", fontsize=10, fontweight="bold")
    axes[1].set_title("B. Covariate balance before vs after weighting",
                      fontsize=11.5, fontweight="bold", loc="left")
    axes[1].legend(fontsize=9, loc="lower right", frameon=True)
    axes[1].grid(True, alpha=0.25, axis="x", color=C_GRID)
    axes[1].set_xlim(0, max(df["smd_unweighted"].max() * 1.05, 0.35))

    fig.suptitle("Figure S1.  Stabilised weight diagnostics and covariate balance",
                 fontsize=12.5, fontweight="bold", y=1.02)
    plt.tight_layout()
    plt.savefig(FIG_DIR / "FigureS1_Weights.png", dpi=350, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    print("  ✓ Figure S1 (Weight + balance) [Max label repositioned]")


def figS2_threshold(threshold_df):
    fig, ax = plt.subplots(figsize=(6.8, 5.0))

    fio2_thresholds = [40, 50, 60]
    peep_thresholds = [5, 8, 10]

    pivot_rd = threshold_df.pivot(index="peep_threshold", columns="fio2_threshold", values="rd")
    pivot_lo = threshold_df.pivot(index="peep_threshold", columns="fio2_threshold", values="rd_ci_lo")
    pivot_hi = threshold_df.pivot(index="peep_threshold", columns="fio2_threshold", values="rd_ci_hi")
    pivot_n = threshold_df.pivot(index="peep_threshold", columns="fio2_threshold", values="n")

    im = ax.imshow(pivot_rd.values * 100, cmap="RdBu_r", vmin=-25, vmax=25, aspect="auto")

    for i in range(len(peep_thresholds)):
        for j in range(len(fio2_thresholds)):
            rd = pivot_rd.values[i, j] * 100
            lo = pivot_lo.values[i, j] * 100
            hi = pivot_hi.values[i, j] * 100
            n = int(pivot_n.values[i, j])

            sig = "★" if (lo > 0 or hi < 0) else ""

            color = "white" if abs(rd) > 12 else "#222"

            txt = f"{rd:+.1f}  [{lo:+.1f}, {hi:+.1f}]\n(n={n}){sig}"
            ax.text(j, i, txt, ha="center", va="center", fontsize=9,
                    color=color, fontweight="bold" if sig else "normal")

    ax.set_xticks(range(len(fio2_thresholds)))
    ax.set_xticklabels([f"FiO₂ ≤{t}%" for t in fio2_thresholds], fontsize=10)
    ax.set_yticks(range(len(peep_thresholds)))
    ax.set_yticklabels([f"PEEP ≤{t}" for t in peep_thresholds], fontsize=10)
    ax.set_xlabel("Day-2 FiO₂ threshold", fontsize=11, fontweight="bold", labelpad=8)
    ax.set_ylabel("Day-2 PEEP threshold (cmH₂O)", fontsize=11, fontweight="bold", labelpad=8)
    ax.set_title("Figure S2.  Exploratory threshold grid — risk difference (pp, 95% CI)",
                 fontsize=12, fontweight="bold", pad=10)

    cbar = plt.colorbar(im, ax=ax, label="Risk difference (percentage points)")
    cbar.ax.axhline(0, color="white", lw=1.2)
    ax.text(2.4, -0.85, "★ = 95% CI excludes 0 (statistically significant)",
            fontsize=8.5, color="#444", style="italic", ha="right")

    plt.tight_layout()
    plt.savefig(FIG_DIR / "FigureS2_Threshold.png", dpi=350, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    print("  ✓ Figure S2 (Threshold heatmap)")


def logrank_pvalue(analyzable):
    early = analyzable[analyzable["strategy"] == "early"]
    deferred = analyzable[analyzable["strategy"] == "deferred"]
    res = logrank_test(
        early["surv_time"].values,
        deferred["surv_time"].values,
        early["event"].values,
        deferred["event"].values,
    )
    return res.p_value


def figS3_subgroup(sub_df):
    """Forest plot of subgroup weighted RDs."""
    d = sub_df.copy()
    d = d[~d["subgroup"].str.startswith(("SOFA", "No ARDS"))]  # drop duplicates/report once
    labels = [f"{r.subgroup}  (n={int(r.n)})".replace(">=", "≥")
              for r in d.itertuples()]
    rds = d["rd"].values * 100
    lo = d["rd_ci_lo"].values * 100
    hi = d["rd_ci_hi"].values * 100

    fig, ax = plt.subplots(figsize=(7.2, 0.75 * len(d) + 1.8))
    y_pos = np.arange(len(d))[::-1]
    xerr = [[r - l for r, l in zip(rds, lo)], [h - r for r, h in zip(rds, hi)]]
    ax.errorbar(rds, y_pos, xerr=xerr, fmt="none",
                ecolor="#444", elinewidth=1.4, capsize=4, capthick=1.4)
    for y, rd, l, h in zip(y_pos, rds, lo, hi):
        sig = "*" if (l > 0 or h < 0) else ""
        ax.scatter(rd, y, s=120, marker="s", color="#2C3E50",
                   edgecolor="white", linewidth=1.2, zorder=5)
        ax.text(h + 0.5, y, f"{rd:+.1f}  [{l:+.1f}, {h:+.1f}]{sig}",
                va="center", ha="left", fontsize=9, color="#222", fontweight="bold")
    ax.axvline(0, color="#888", linestyle="--", linewidth=0.9, alpha=0.7)
    ax.set_yticks(y_pos)
    ax.set_yticklabels(labels, fontsize=10)
    all_vals = list(lo) + list(hi) + list(rds)
    ax.set_xlim(min(all_vals) - 4, max(all_vals) + 8)
    ax.set_ylim(-1.2, len(d) - 0.2 + 0.6)
    ax.set_xlabel("Risk difference (percentage points)", fontsize=11, fontweight="bold")
    fig.text(0.36, 0.015, "\u2190 Favours early", fontsize=9.5, color=C_EARLY,
             fontweight="bold", ha="left", va="bottom")
    fig.text(0.95, 0.015, "Favours deferred \u2192", fontsize=9.5, color=C_DEFER,
             fontweight="bold", ha="right", va="bottom")
    fig.suptitle("Figure S3.  Subgroup analyses \u2014 weighted risk difference (95% CI)",
                 fontsize=12, fontweight="bold", y=0.97)
    fig.text(0.97, 0.012, "* 95% CI excludes zero", ha="right", va="bottom",
             fontsize=8, color="#555", style="italic")
    ax.grid(True, alpha=0.2, axis="x", color=C_GRID)
    plt.subplots_adjust(top=0.86, bottom=0.16, left=0.34, right=0.97)
    plt.savefig(FIG_DIR / "FigureS3_Subgroup.png", dpi=350, facecolor="white",
                bbox_inches="tight")
    plt.close(fig)
    print("  \u2713 Figure S3 (Subgroup forest)")


def main():
    print("=" * 70)
    print("Fixing figures: removing text/figure overlap")
    print("=" * 70)

    with open(OUT_DIR / "analysis_summary.json") as f:
        summ = json.load(f)

    print("\n[1/4] Loading analyzable cohort...")
    cache = OUT_DIR / "mimic_analyzable.parquet"
    if cache.exists():
        analyzable = pd.read_parquet(cache)
        have_flow = False
        print(f"  Loaded from cache: {cache.name} ({len(analyzable):,} rows)")
    else:
        df = load_excel()
        eligible, flow = screen_eligibility(df)
        eligible = classify_strategies(eligible)
        eligible = compute_outcomes(eligible)
        eligible, covariate_cols, cov_df = build_covariates(eligible)
        analyzable, w_diag = estimate_weights(eligible, cov_df, covariate_cols)
        analyzable.to_parquet(cache, index=False)
        have_flow = True
        print(f"  Cached analyzable cohort -> {cache.name}")
    print(f"  Early: {(analyzable['strategy']=='early').sum()}, "
          f"Deferred: {(analyzable['strategy']=='deferred').sum()}")

    print("\n[2/4] Computing log-rank p-value...")
    p_logrank = logrank_pvalue(analyzable)
    print(f"  Log-rank p = {p_logrank:.4f}")

    sc = summ["strategy_counts"]
    print("\n[3/4] Generating figures...")
    if have_flow:
        fig1_flow(flow, sc["early"], sc["deferred"], sc["unascertainable"], sc["grace_death"])
    else:
        print("  - Figure 1 skipped (canonical generator: make_flow_figures.py)")
    fig2_survival(analyzable, p_logrank)
    fig3_forest(analyzable, summ)

    print("\n[4/4] Generating supplementary figures...")
    balance_df = pd.read_csv(OUT_DIR / "covariate_balance.csv")
    threshold_df = pd.read_csv(OUT_DIR / "threshold_results.csv")
    figS1_weights(analyzable, balance_df, summ["weight_diagnostics"])
    figS2_threshold(threshold_df)
    subg_path = OUT_DIR / "subgroup_results.csv"
    if subg_path.exists():
        figS3_subgroup(pd.read_csv(subg_path))

    print("\n" + "=" * 70)
    print("All figures regenerated (incl. robustness/subgroup)")
    print("=" * 70)


if __name__ == "__main__":
    main()