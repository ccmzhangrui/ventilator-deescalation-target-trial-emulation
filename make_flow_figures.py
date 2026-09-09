#!/usr/bin/env python3
"""
CONSORT/NEJM-style flow diagrams (top-journal layout, v3):
  - Axes-geometry measured in PIXELS (axes margins no longer corrupt autofit)
  - Pre-wrapped labels (<= ~28 chars/line) so fonts stay uniform across boxes
  - Main column / exclusion column grid, elbow connectors, bold n-counts
Outputs:
  figures/Figure1_Flow.png       (MIMIC-IV, derivation)
  figures/FigureS5_eICU_Flow.png (eICU-CRD, validation)
All counts are derived from the locked JSONs — single source of truth.
"""
import json
from pathlib import Path
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch

BASE = Path("/Users/zhangrui/Desktop/tte/TTE_Deescalation_v2")
RES = BASE / "results"
FIG = BASE / "figures"

# ── refined journal palette ──
C_MAIN_EC = "#33404D"
C_MAIN_FC = "#FFFFFF"
C_KEY_FC  = "#EAF1F8"
C_KEY_EC  = "#1F4E79"
C_EXCL_FC = "#F7F5F1"
C_EXCL_EC = "#CFC8BC"
C_EXCL_TX = "#6B6257"
C_ARROW   = "#55606C"
C_EARLY   = "#C0392B"
C_DEFER   = "#2E86AB"
C_FINAL_FC = "#EAF3E7"
C_FINAL_EC = "#4E7C3A"
C_FINAL_TX = "#2F5023"

plt.rcParams.update({
    "font.family": "DejaVu Sans",
    "savefig.dpi": 350,
    "figure.dpi": 110,
})


def draw_flow(steps, split, out_path, figsize=(10.6, 17.5), xlim=14, top=15.5):
    fig, ax = plt.subplots(figsize=figsize)
    ax.set_xlim(0, xlim)
    ax.set_ylim(0, top + 0.6)
    ax.axis("off")

    MX, MW = 4.10, 6.6          # main column centre / width  (0.80 – 7.40)
    XX, XW = 10.90, 5.9         # exclusion column centre / width (7.95 – 13.85)
    m_r = MX + MW / 2
    x_l = XX - XW / 2

    # TRUE inches per data unit, measured from the rendered axes (pixel-exact)
    fig.canvas.draw()
    _rend = fig.canvas.get_renderer()
    UNIT_IN = ax.get_window_extent(renderer=_rend).width / fig.dpi / xlim

    def autofit(texts, max_w_in, fs, floor):
        """Renderer-measured shrink-to-fit: guarantees containment."""
        fig.canvas.draw()
        rend = fig.canvas.get_renderer()
        dpi = fig.dpi
        cur = fs
        while cur > floor:
            widest = max(t.get_window_extent(renderer=rend).width for t in texts)
            if widest <= max_w_in * dpi:
                break
            cur -= 0.2
            for t in texts:
                t.set_fontsize(cur)
        return cur

    def box(cx, w, y, h, lines, fc, ec, fs=16, tc="#1F2937", kind="main",
            floor=12.0, lw=1.4, margin=0.75):
        bb = FancyBboxPatch((cx - w / 2, y - h / 2), w, h,
                            boxstyle="round,pad=0.045,rounding_size=0.10",
                            fc=fc, ec=ec, lw=lw, mutation_aspect=1)
        ax.add_patch(bb)
        n = len(lines)
        lh = h * 0.76 / max(n, 1)
        y0 = y + (n - 1) * lh / 2
        texts = []
        for i, ln in enumerate(lines):
            if kind == "main":
                bold = i == n - 1 and ln.startswith("n =")
                fsi = fs + 1.0 if bold else fs
                color = tc
            else:  # exclusion
                bold = (i == 0 and ln.startswith("Excluded")) or \
                       (i == n - 1 and ln.startswith("n ="))
                fsi = fs
                color = tc
            texts.append(ax.text(cx, y0 - i * lh, ln, ha="center", va="center",
                                 fontsize=fsi, color=color,
                                 fontweight="bold" if bold else "normal"))
        autofit(texts, (w - margin) * UNIT_IN, fs, floor)

    def varrow(x, y1, y2):
        ax.annotate("", xy=(x, y2), xytext=(x, y1),
                    arrowprops=dict(arrowstyle="-|>", color=C_ARROW, lw=1.6,
                                    mutation_scale=16))

    def hconn(y, x1, x2):
        ax.plot([x1, x2], [y, y], color=C_EXCL_EC, lw=1.2,
                solid_capstyle="butt", zorder=1)

    # ── main steps ──
    dy = 2.05
    ys = [top - 0.70 - i * dy for i in range(len(steps))]
    for i, st in enumerate(steps):
        y = ys[i]
        n_lines = len(st["lines"])
        h = 0.56 + 0.36 * n_lines
        key = st.get("key")
        box(MX, MW, y, h, st["lines"],
            fc=C_KEY_FC if key else C_MAIN_FC,
            ec=C_KEY_EC if key else C_MAIN_EC,
            fs=16, lw=1.7 if key else 1.4)
        if st.get("excl"):
            el = st["excl"]
            eh = 0.44 + 0.29 * len(el)
            box(XX, XW, y, eh, el, fc=C_EXCL_FC, ec=C_EXCL_EC,
                fs=14.5, tc=C_EXCL_TX, kind="excl", floor=11.5)
            hconn(y, m_r + 0.06, x_l - 0.06)
        if i > 0:
            py = ys[i - 1]
            ph = 0.46 + 0.315 * len(steps[i - 1]["lines"])
            varrow(MX, py - ph / 2 - 0.03, y + h / 2 - 0.10)

    # ── strategy split ──
    last_y = ys[-1]
    last_h = 0.46 + 0.315 * len(steps[-1]["lines"])
    bus_y = last_y - last_h / 2 - 0.44
    arm_y = bus_y - 0.72
    AW = 3.00
    ax1_x, ax2_x = 2.25, 6.15
    # bus
    ax.plot([MX, MX], [last_y - last_h / 2, bus_y], color=C_ARROW, lw=1.6)
    ax.plot([ax1_x, ax2_x], [bus_y, bus_y], color=C_ARROW, lw=1.6)
    for xx in (ax1_x, ax2_x):
        ax.annotate("", xy=(xx, arm_y + 0.50), xytext=(xx, bus_y),
                    arrowprops=dict(arrowstyle="-|>", color=C_ARROW, lw=1.6,
                                    mutation_scale=16))
    # arm boxes: 3 stacked lines (name / "de-escalation" / n)
    for ax_x, color, word, nn in (
            (ax1_x, C_EARLY, "Early", split["left_n"]),
            (ax2_x, C_DEFER, "Deferred", split["right_n"])):
        bb = FancyBboxPatch((ax_x - AW / 2, arm_y - 0.50), AW, 1.00,
                            boxstyle="round,pad=0.045,rounding_size=0.10",
                            fc=color, ec=color, lw=1.4)
        ax.add_patch(bb)
        t_a = ax.text(ax_x, arm_y + 0.27, word, ha="center", va="center",
                      fontsize=16, color="white", fontweight="bold")
        t_b = ax.text(ax_x, arm_y, "de-escalation", ha="center", va="center",
                      fontsize=13.5, color="white")
        t_c = ax.text(ax_x, arm_y - 0.28, f"n = {nn:,}", ha="center",
                      va="center", fontsize=14.5, color="white",
                      fontweight="bold")
        autofit([t_a, t_b, t_c], (AW - 0.45) * UNIT_IN, 13.5, 10.5)

    # exclusion for split: elbow connector (bus → box top centre)
    if split.get("excl_lines"):
        el = split["excl_lines"]
        eh = 0.44 + 0.29 * len(el)
        box(XX, XW, arm_y, eh, el, fc=C_EXCL_FC, ec=C_EXCL_EC,
            fs=14.5, tc=C_EXCL_TX, kind="excl", floor=11.0)
        ax.plot([ax2_x + 0.30, XX], [bus_y, bus_y], color=C_EXCL_EC, lw=1.2,
                solid_capstyle="butt", zorder=1)
        ax.plot([XX, XX], [bus_y, arm_y + eh / 2 + 0.03], color=C_EXCL_EC,
                lw=1.2, solid_capstyle="butt", zorder=1)

    # final analysable box under the two arms
    fy = arm_y - 1.24
    ax.plot([ax1_x, ax1_x], [arm_y - 0.50, fy + 0.74], color=C_ARROW, lw=1.4)
    ax.plot([ax2_x, ax2_x], [arm_y - 0.50, fy + 0.74], color=C_ARROW, lw=1.4)
    ax.plot([ax1_x, ax2_x], [fy + 0.74, fy + 0.74], color=C_ARROW, lw=1.4)
    ax.annotate("", xy=(MX, fy + 0.37), xytext=(MX, fy + 0.74),
                arrowprops=dict(arrowstyle="-|>", color=C_ARROW, lw=1.6,
                                mutation_scale=16))
    bb = FancyBboxPatch((MX - MW / 2, fy - 0.37), MW, 0.74,
                        boxstyle="round,pad=0.045,rounding_size=0.10",
                        fc=C_FINAL_FC, ec=C_FINAL_EC, lw=1.7)
    ax.add_patch(bb)
    tf1 = ax.text(MX, fy + 0.12, split["final_label"], ha="center", va="center",
                  fontsize=16, color="#1F2937")
    tf2 = ax.text(MX, fy - 0.15, f"n = {split['final_n']:,}", ha="center",
                  va="center", fontsize=17, color=C_FINAL_TX, fontweight="bold")
    autofit([tf1, tf2], (MW - 0.75) * UNIT_IN, 16, 12.0)

    fig.savefig(out_path, bbox_inches="tight", facecolor="white",
                pad_inches=0.15)
    plt.close(fig)
    print(f"saved: {out_path}")


# ───────────────────── Figure 1: MIMIC-IV ─────────────────────
A = json.load(open(RES / "analysis_summary.json"))
F = A["flow"]
S = A["strategy_counts"]
n_an = S["early"] + S["deferred"]

steps_mimic = [
    {"lines": ["ICU records in MIMIC-IV v2.2",
               f"n = {F['total_records']:,}"], "excl": None},
    {"lines": ["Received invasive mechanical",
               "ventilation",
               f"n = {F['ventilated']:,}"],
     "excl": ["Excluded:", "No invasive ventilation",
              f"n = {F['total_records'] - F['ventilated']:,}"]},
    {"lines": ["Mechanical ventilation > 24 h",
               f"n = {F['vent_gt_24h']:,}"],
     "excl": ["Excluded:", "Ventilated \u226424 h",
              f"n = {F['ventilated'] - F['vent_gt_24h']:,}"]},
    {"lines": ["Alive at 48-h landmark",
               f"n = {F['alive_at_48h']:,}"],
     "excl": ["Excluded:", "Died within 48 h",
              f"n = {F['vent_gt_24h'] - F['alive_at_48h']:,}"]},
    {"lines": ["Day-2 ventilator settings",
               "recorded",
               f"n = {F['has_d2_settings']:,}"],
     "excl": ["Excluded:", "No Day-2 settings",
              f"n = {F['alive_at_48h'] - F['has_d2_settings']:,}"]},
    {"lines": ["Eligible: Day-2 mode",
               "classifiable and FiO\u2082 \u226450%",
               "with PEEP \u226410 cmH\u2082O",
               f"n = {F['eligible']:,}"],
     "key": True,
     "excl": ["Excluded:", "FiO\u2082 >50%, PEEP >10 cmH\u2082O,",
              "or mode unclassifiable",
              f"n = {F['has_d2_settings'] - F['eligible']:,}"]},
]
split_mimic = {
    "left_n": S["early"],
    "right_n": S["deferred"],
    "excl_lines": ["Excluded from primary analysis:",
                   f"Day-3 unascertainable, n = {S['unascertainable']:,}",
                   f"Grace-period death, n = {S['grace_death']:,}"],
    "final_label": "Analysable cohort (CCW-weighted)",
    "final_n": n_an,
}
draw_flow(steps_mimic, split_mimic, FIG / "Figure1_Flow.png")

# ───────────────────── Figure S5: eICU-CRD ─────────────────────
E = json.load(open(RES / "eicu_validation.json"))
EF = E["flow"]
ES = E["strategy_counts"]

steps_eicu = [
    {"lines": ["First ICU stays of adults",
               "in eICU-CRD v2.0",
               f"n = {EF['first_stay_adults']:,}"], "excl": None},
    {"lines": ["Invasive ventilator settings",
               "charted",
               f"n = {EF['with_vent_settings']:,}"],
     "excl": ["Excluded:", "No ventilator settings",
              f"n = {EF['first_stay_adults'] - EF['with_vent_settings']:,}"]},
    {"lines": ["Mechanical ventilation > 24 h",
               f"n = {EF['vent_gt_24h']:,}"],
     "excl": ["Excluded:", "Ventilated \u226424 h",
              f"n = {EF['with_vent_settings'] - EF['vent_gt_24h']:,}"]},
    {"lines": ["Alive at 48-h landmark",
               f"n = {EF['alive_at_48h']:,}"],
     "excl": ["Excluded:", "Died within 48 h",
              f"n = {EF['vent_gt_24h'] - EF['alive_at_48h']:,}"]},
    {"lines": ["Day-2 FiO\u2082 / PEEP charted",
               f"n = {EF['has_d2_settings']:,}"],
     "excl": ["Excluded:", "Not charted at Day 2",
              f"n = {EF['alive_at_48h'] - EF['has_d2_settings']:,}"]},
    {"lines": ["Day-2 FiO\u2082 \u226450% and",
               "PEEP \u226410 cmH\u2082O",
               f"n = {EF['fio2_peep_ok']:,}"],
     "excl": ["Excluded:", "Above oxygenation thresholds",
              f"n = {EF['has_d2_settings'] - EF['fio2_peep_ok']:,}"]},
    {"lines": ["Eligible: Day-2 mode",
               "classifiable (treatment order",
               "or PS/PC inference)",
               f"n = {EF['eligible']:,}"],
     "key": True,
     "excl": ["Excluded:", "Mode unclassifiable",
              f"n = {EF['fio2_peep_ok'] - EF['mode_d2_classifiable']:,}"]},
]
split_eicu = {
    "left_n": ES["early"],
    "right_n": ES["deferred"],
    "excl_lines": ["Excluded from primary analysis:",
                   f"Day-3 unascertainable, n = {ES['unascertainable']:,}",
                   f"Grace-period death, n = {ES['grace_death']:,}"],
    "final_label": "Analysable cohort (CCW-weighted)",
    "final_n": E["n_analyzable"],
}
draw_flow(steps_eicu, split_eicu, FIG / "FigureS5_eICU_Flow.png",
          figsize=(10.6, 19.5), top=17.5)
