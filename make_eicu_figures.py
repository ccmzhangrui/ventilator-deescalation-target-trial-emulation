#!/usr/bin/env python3
"""
External validation figures:
  Figure 4  — Two-database comparison (MIMIC-IV derivation vs eICU-CRD validation)
  (Figure S5 eICU flow is produced by make_flow_figures.py, v4 layout)
Reads: results/analysis_summary.json, results/eicu_validation.json
"""
import json
from pathlib import Path
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch
from matplotlib.lines import Line2D

BASE = Path("/Users/zhangrui/Desktop/tte/TTE_Deescalation_v2")
RES = BASE / "results"
FIG = BASE / "figures"

C_EARLY = "#C0392B"
C_DEFER = "#2E86AB"
C_TEXT = "#222222"
C_MIMIC = "#1F4E79"
C_EICU = "#B4540A"

plt.rcParams.update({
    "font.family": "DejaVu Sans",
    "font.size": 10,
    "axes.edgecolor": "#333333",
    "axes.spines.top": False,
    "axes.spines.right": False,
    "savefig.dpi": 350,
    "figure.dpi": 110,
})

A = json.load(open(RES / "analysis_summary.json"))
E = json.load(open(RES / "eicu_validation.json"))

mp = A["primary"]
ep = E["primary_hospital_mortality"]

# ────────────────────────── Figure 4 ──────────────────────────────
fig = plt.figure(figsize=(7.4, 4.0))
gs = fig.add_gridspec(1, 2, width_ratios=[0.80, 1.45], wspace=0.50,
                      left=0.08, right=0.97, top=0.86, bottom=0.16)

# Panel A: weighted mortality bars both databases
axA = fig.add_subplot(gs[0])
groups = ["MIMIC-IV\n28-day in-hospital",
          "eICU-CRD\nin-hospital"]
early_v = [mp["mort_early"] * 100, ep["mort_early"] * 100]
defer_v = [mp["mort_deferred"] * 100, ep["mort_deferred"] * 100]
x = np.arange(2)
w = 0.34
b1 = axA.bar(x - w / 2, early_v, w, color=C_EARLY, alpha=0.88,
             label="Early")
b2 = axA.bar(x + w / 2, defer_v, w, color=C_DEFER, alpha=0.88,
             label="Deferred")
for xi, v in zip(x - w / 2, early_v):
    axA.text(xi, v + 1.0, f"{v:.1f}%", ha="center", fontsize=10.5,
             fontweight="bold", color=C_EARLY)
for xi, v in zip(x + w / 2, defer_v):
    axA.text(xi, v + 1.0, f"{v:.1f}%", ha="center", fontsize=10.5,
             fontweight="bold", color=C_DEFER)
axA.set_xticks(x)
axA.set_xticklabels(groups, fontsize=9.5)
axA.set_ylabel("Weighted mortality (%)", fontsize=10.5)
axA.set_ylim(0, max(defer_v) * 1.45)  # headroom for top legend
axA.legend(frameon=False, fontsize=9.5, loc="upper center", ncol=2,
           bbox_to_anchor=(0.5, 1.02), handletextpad=0.4, columnspacing=1.0)
axA.set_title("A  Mortality (%)", fontsize=11.5, fontweight="bold", loc="left")
axA.grid(axis="y", color="#DDDDDD", lw=0.6, alpha=0.7)
axA.set_axisbelow(True)

# Panel B: RD forest (two databases)
axB = fig.add_subplot(gs[1])
rds = [mp["rd"] * 100, ep["rd"] * 100]
lo = [mp["rd_ci_lo"] * 100, ep["rd_ci_lo"] * 100]
hi = [mp["rd_ci_hi"] * 100, ep["rd_ci_hi"] * 100]
labels = ["MIMIC-IV (n=1,902)", "eICU-CRD (n=10,957)"]
colors = [C_MIMIC, C_EICU]
ypos = [1, 0]
for yv, rd, l, h, c, lab in zip(ypos, rds, lo, hi, colors, labels):
    axB.plot([l, h], [yv, yv], color=c, lw=2.2, solid_capstyle="round")
    axB.plot(rd, yv, "s", color=c, markersize=11, markeredgecolor="white",
             markeredgewidth=1.2)
    axB.text(h + 0.35, yv, f"{rd:.1f}  ({l:.1f} to {h:.1f})",
             va="center", ha="left", fontsize=10, color=C_TEXT, fontweight="bold")
axB.axvline(0, color="#555555", lw=1, ls="--", alpha=0.7)
# legend at top-right of panel B: database name + n, color-coded
from matplotlib.lines import Line2D as _L2D
_legend_handles = [_L2D([0], [0], marker="s", color="w", markerfacecolor=c, markersize=9,
                        markeredgecolor="white", label=lab)
                  for c, lab in zip(colors, labels)]
axB.legend(handles=_legend_handles, loc="upper right", fontsize=9.5, frameon=True,
           edgecolor="#888", facecolor="white", framealpha=0.95,
           handletextpad=0.5, borderpad=0.4)
axB.set_yticks(ypos)
axB.set_yticklabels([""] * len(ypos))   # suppress left-side row labels (overlaps panel A)
axB.tick_params(axis="y", which="both", left=False, right=False)  # no y-tick labels at all
axB.set_xlabel("Risk difference, percentage points (95% CI)", fontsize=10.5)
axB.set_xlim(min(lo) - 1.5, max(hi) + 5.5)   # right padding for the value text
axB.set_ylim(-0.6, 1.6)
axB.set_title("B  Early vs deferred: risk difference", fontsize=11.5,
              fontweight="bold", loc="left")
axB.grid(axis="x", color="#DDDDDD", lw=0.6, alpha=0.7)
axB.set_axisbelow(True)
axB.text(0.02, -0.32,
         "Negative values favour early de-escalation. Outcome windows differ: "
         "28-day in-hospital\nmortality (MIMIC-IV) vs in-hospital mortality "
         "(eICU-CRD); estimates are therefore compared\ndirectionally, not pooled.",
         transform=axB.transAxes, fontsize=8.2, color="#666666", va="top")

fig.suptitle("Figure 4.  External validation of early de-escalation benefit "
             "in eICU-CRD v2.0", fontsize=13.5, fontweight="bold",
             color=C_TEXT, x=0.08, ha="left")
fig.savefig(FIG / "Figure4_ExternalValidation.png",
            bbox_inches="tight", facecolor="white")
plt.close(fig)
print("Figure 4 saved")
