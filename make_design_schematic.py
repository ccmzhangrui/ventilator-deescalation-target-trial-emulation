#!/usr/bin/env python3
"""
Figure S6 — Schematic of the clone-censor-weight (CCW) target trial emulation design.
NEJM-style timeline + cloning diagram, matching the pedagogy of Gilding & Longo
(Eur Respir J 2026, Figure 1C).

Design rules (top-journal print legibility):
  - Boxes sized generously for LARGE fonts (nominal 13-17 pt -> ~7-9 pt at 15 cm print width)
  - Text pre-wrapped to short lines; renderer-measured auto-shrink loop remains
    as a safety net so text can NEVER overflow its box.

Output: figures/FigureS6_CCW_Design.png
"""
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch, Rectangle
from pathlib import Path

BASE = Path("/Users/zhangrui/Desktop/tte/TTE_Deescalation_v2")
OUT = BASE / "figures" / "FigureS6_CCW_Design.png"

INK      = "#1a1a1a"
GREY     = "#8a8a8a"
PRE_BG   = "#e8e8e8"
GRACE_BG = "#fdf0d5"
GRACE_ED = "#e0a458"
FU_BG    = "#eaf3ee"
EARLY    = "#C0392B"
DEFER    = "#2E86AB"
GREEN    = "#2e6b46"
BOX_FC   = "#ffffff"

fig, ax = plt.subplots(figsize=(11.6, 7.8), dpi=300)
ax.set_xlim(0, 116)
ax.set_ylim(0, 78)
ax.axis("off")

MIN_FS = 6.0
INNER_MARGIN_X = 1.0
INNER_MARGIN_Y = 0.5


def _fits(t, x, y, w, h):
    fig.canvas.draw()
    rend = fig.canvas.get_renderer()
    bb = t.get_window_extent(renderer=rend)
    (x0d, y0d) = ax.transData.transform((x + INNER_MARGIN_X, y + INNER_MARGIN_Y))
    (x1d, y1d) = ax.transData.transform((x + w - INNER_MARGIN_X, y + h - INNER_MARGIN_Y))
    return bb.width <= (x1d - x0d) and bb.height <= (y1d - y0d)


def box(x, y, w, h, text, fc=BOX_FC, ec=INK, fs=13.0, bold=False, tc=INK, lw=1.4):
    b = FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.55,rounding_size=0.9",
                       fc=fc, ec=ec, lw=lw, zorder=3)
    ax.add_patch(b)
    cur = fs
    t = ax.text(x + w/2, y + h/2, text, ha="center", va="center", fontsize=cur,
                color=tc, fontweight="bold" if bold else "normal", zorder=4,
                linespacing=1.25)
    while not _fits(t, x, y, w, h) and cur > MIN_FS:
        cur -= 0.2
        t.set_fontsize(cur)
    return t


def arrow(x1, y1, x2, y2, color=INK, lw=1.6, style="-|>", ms=13, ls="-"):
    a = FancyArrowPatch((x1, y1), (x2, y2), arrowstyle=style, mutation_scale=ms,
                        color=color, lw=lw, zorder=2, linestyle=ls)
    ax.add_patch(a)


# ══════════════ PANEL A — timeline ══════════════
ax.text(1.5, 75.6, "A", fontsize=16, fontweight="bold", color=INK, va="center")
ax.text(5.0, 75.6, "Emulated trial timeline", fontsize=14.5, fontweight="bold",
        color=INK, va="center")

tly = 63.0
x0, x_land, x_grace, x_end = 10, 46, 60, 106

ax.add_patch(Rectangle((x0, tly-3.4), x_land-x0, 6.8, fc=PRE_BG, ec="none", zorder=1))
ax.add_patch(Rectangle((x_land, tly-3.4), x_grace-x_land, 6.8, fc=GRACE_BG, ec="none", zorder=1))
ax.add_patch(Rectangle((x_grace, tly-3.4), x_end-x_grace, 6.8, fc=FU_BG, ec="none", zorder=1))

arrow(x0, tly, x_end+2.5, tly, lw=2.0)

for x, lab in [(x0, "ICU admission\nMV initiated (Day 0)"),
               (x_land, "48-h landmark\nTIME ZERO (Day 2)"),
               (x_grace, "Grace window\nends (Day 3)"),
               (x_end, "Day 28\n(end of follow-up)")]:
    ax.plot([x, x], [tly-0.9, tly+0.9], color=INK, lw=1.6, zorder=3)
    ax.text(x, tly-6.6, lab, ha="center", va="top", fontsize=11, color=INK,
            linespacing=1.35,
            fontweight="bold" if "TIME ZERO" in lab else "normal")

ax.text((x0+x_land)/2, tly+5.2, "Pre-landmark window (Day 1\u20132)\n"
        "eligibility + baseline covariates",
        ha="center", va="bottom", fontsize=11, color="#555555", linespacing=1.35)
ax.text((x_land+x_grace)/2, tly+5.2, "Grace window (24 h)\nstrategy implemented",
        ha="center", va="bottom", fontsize=11, color="#9a6a1f", linespacing=1.35)
ax.text((x_grace+x_end)/2, tly+5.2, "Follow-up to death, discharge or Day 28\n"
        "28-day in-hospital mortality \u00b7 RMST \u00b7 VFD28",
        ha="center", va="bottom", fontsize=11, color=GREEN, linespacing=1.35)

# time-zero anchor box (centred under the landmark tick; 4 short lines)
box(x_land-22, 41.9, 44, 9.3,
    "Three anchors aligned at time zero:\n"
    "eligibility \u2022 strategy assignment \u2022\n"
    "follow-up start (prevents\n"
    "immortal time bias by design)",
    fc="#fff7e6", ec=GRACE_ED, fs=12, lw=1.6)

# ══════════════ PANEL B — clone, censor, weight ══════════════
ax.text(1.5, 37.0, "B", fontsize=16, fontweight="bold", color=INK, va="center")
ax.text(5.0, 37.0, "Clone \u2013 censor \u2013 weight at time zero", fontsize=14.5,
        fontweight="bold", color=INK, va="center")

# lanes: early [22.4, 35.2] centre 28.8; deferred [8.9, 21.7] centre 15.3
box(2.5, 16.2, 19.5, 11.2, "Each eligible\npatient at the\n48-h landmark",
    fs=13, bold=True)

arrow(22.5, 24.6, 30.0, 27.6, lw=1.8)
arrow(22.5, 19.0, 30.0, 16.6, lw=1.8)
ax.text(25.3, 22.4, "cloned", fontsize=10, color=GREY, style="italic",
        ha="center", va="center",
        bbox=dict(fc="white", ec="none", pad=1.2), zorder=5)

# clone boxes
box(30.6, 22.4, 32.0, 12.8, "Clone assigned to\nEARLY de-escalation\n"
    "(step-down by Day 3 or\nextubation \u226472 h)",
    ec=EARLY, fs=17, tc=EARLY, lw=1.8)
box(30.6, 8.9, 32.0, 12.8, "Clone assigned to\nDEFERRED de-escalation\n"
    "(support maintained\nor increased)",
    ec=DEFER, fs=17, tc=DEFER, lw=1.8)

arrow(63.1, 28.8, 68.2, 28.8, lw=1.8)
arrow(63.1, 15.3, 68.2, 15.3, lw=1.8)

# censor boxes
box(68.7, 22.4, 31.0, 12.8,
    "Censored when data\ndeviate: no step-down\nby Day 3 \u2192 early\nclone censored",
    fc="#fbeeee", ec=EARLY, fs=17, lw=1.4)
box(68.7, 8.9, 31.0, 12.8,
    "Censored when data\ndeviate: step-down /\nextubation by Day 3 \u2192\n"
    "deferred clone censored",
    fc="#e9f2f7", ec=DEFER, fs=17, lw=1.4)

# outcome box
box(101.8, 10.2, 13.2, 14.6,
    "Weighted\nper-protocol\ncontrast:\nRD + RR\n(95% CI by\nbootstrap)",
    fc="#eaf3ee", ec=GREEN, fs=13, lw=1.8)
arrow(100.2, 26.5, 101.3, 23.5, color=GREEN, lw=1.8)
arrow(100.2, 12.5, 101.3, 14.0, color=GREEN, lw=1.8)

# weighting box (bottom, 3 lines of fine print)
box(52.0, 0.9, 58.0, 7.0,
    "Uncensored clones reweighted by stabilised inverse-probability-\n"
    "of-censoring weights (21 pre-landmark covariates; truncated at\n"
    "1st\u201399th percentile) \u2192 emulates random assignment",
    fc="#f2f2f2", ec=INK, fs=10.5, lw=1.6)

# dotted connectors: censor boxes -> shared right channel -> weighting box
arrow(99.7, 23.0, 101.5, 23.0, color=GREY, lw=1.2, style="-", ls=(0, (3, 2)))
arrow(99.7, 11.5, 101.5, 11.5, color=GREY, lw=1.2, style="-", ls=(0, (3, 2)))
arrow(101.5, 23.0, 101.5, 7.9, color=GREY, lw=1.2, style="-", ls=(0, (3, 2)))

# weighting -> outcome arrow
arrow(106.5, 7.9, 106.5, 10.2, color=GREEN, lw=1.8)

plt.tight_layout(pad=0.4)
plt.savefig(OUT, dpi=300, bbox_inches="tight", facecolor="white")
print(f"saved: {OUT}")
