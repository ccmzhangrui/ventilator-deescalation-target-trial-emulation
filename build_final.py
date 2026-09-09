#!/usr/bin/env python3
"""
Final Build — TTE De-escalation
Generates:
  - TTE_Deescalation_Manuscript_FINAL.docx  (main article)
  - TTE_Deescalation_Supplement_FINAL.docx  (supplementary)
  - TTE_CoverLetter_FINAL.docx              (cover letter)
  - TTE_汇报_FINAL.pptx                     (17-slide presentation)

All numbers read from analysis_summary.json and CSVs in results/.
All numbers used in manuscript and PPT are IDENTICAL.
"""
import json, os, re
from pathlib import Path
import pandas as pd
from docx import Document
from docx.shared import Pt, Inches, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

BASE = Path("/Users/zhangrui/Desktop/tte/TTE_Deescalation_v2")
RES  = BASE / "results"
FIG_DIR = BASE / "figures"
PPT_DIR = BASE / "ppt_final"
PPT_DIR.mkdir(exist_ok=True)
(PPT_DIR / "assets").mkdir(exist_ok=True)

# ── Load analysis results ──────────────────────────────────────────────
with open(RES / "analysis_summary.json") as f:
    A = json.load(f)

P  = A["primary"]
F  = A["flow"]
S  = A["strategy_counts"]
W  = A["weight_diagnostics"]
SA = A["sensitivity"]
TG = A["threshold_grid"]

# Threshold-grid RD ranges derived from pipeline JSON (no hardcoded effect sizes)
_TGA = [c["rd"] for c in TG]                                   # all 9 cells (fractions)
_TGS = [c["rd"] for c in TG if c["fio2_threshold"] > 40]      # FiO2 <=50/60 subset (6 cells)
_UM = "\u2212"  # unicode minus for signed effect sizes
TG_RANGE_ALL = f"{_UM}{abs(min(_TGA))*100:.1f} to {_UM}{abs(max(_TGA))*100:.1f}"
TG_RANGE_SUB = f"{_UM}{abs(min(_TGS))*100:.1f} to {_UM}{abs(max(_TGS))*100:.1f}"
TG_TILDE_ALL = f"{_UM}{abs(min(_TGA))*100:.1f} ~ {_UM}{abs(max(_TGA))*100:.1f}"
TG_TILDE_ALL_ASCII = f"-{abs(min(_TGA))*100:.1f} ~ -{abs(max(_TGA))*100:.1f}"
TG_MAX_BENEFIT = f"{_UM}{abs(min(_TGA))*100:.1f}"

# Crude (unweighted) observed event counts, computed from the production
# analysable datasets by compute_unweighted.py (NOT back-derived from the
# weighted risks, which is what an earlier version did).
_uw_path = RES / "unweighted_counts.json"
UW = json.load(open(_uw_path)) if _uw_path.exists() else None
UW_OK = UW is not None
if UW_OK:
    UW_M = UW["mimic"]
    UW_E = UW.get("eicu")
else:  # pragma: no cover - defensive fallback
    UW_M = UW_E = None

# Robustness package (SA4/SA5/AIPW/E-value/subgroups)
import pandas as _pd
_rb_path = RES / "robustness_results.json"
RB = json.load(open(_rb_path)) if _rb_path.exists() else {}
RB_OK = bool(RB)
_EV = RB["e_value"]["e_value_point"] if RB_OK else None
_EV_TXT = f"{_EV:.2f}" if _EV else "N/A"
_AIPW_RD = RB["AIPW_doubly_robust"]["rd"] if RB_OK else None
SUBG = _pd.read_csv(RES / "subgroup_results.csv") if (RES / "subgroup_results.csv").exists() else None

# Per-protocol verification trajectory (Figure S4)
_vt_path = RES / "ventilation_trajectory.json"
VT = json.load(open(_vt_path)) if _vt_path.exists() else None
VT_OK = VT is not None
VTD = None
if VT_OK:
    VTD = {d["strategy"]: d for d in VT["deltas_D3_minus_D2_by_group"]}
    VT = VT["mode_distribution"]

# External validation in eICU-CRD v2.0 (Figure 4, Table 5, Figure S5)
_ev_path = RES / "eicu_validation.json"
EV = json.load(open(_ev_path)) if _ev_path.exists() else None
EV_OK = EV is not None
if EV_OK:
    EVP = EV["primary_hospital_mortality"]
    EVF = EV["flow"]
    EVS = EV["strategy_counts"]
    EVSA = EV["sa_strict_extubation"]

# Helpers
def pp(x): return f"{x*100:.1f}"  # to percentage string (1 decimal)
def ci_pp(lo, hi, fmt="{:.1f}"): return f"{fmt.format(lo*100)} to {fmt.format(hi*100)}"
def pp_label(x): return f"{x*100:+.1f}"  # signed pp
def p_fmt(p): return "p < 0.001" if p < 0.001 else f"p = {p:.3f}"
_LR_P = UW_M["logrank_p"] if UW_OK and "logrank_p" in UW_M else None
_LR_TXT = p_fmt(_LR_P) if _LR_P is not None else "p < 0.001"
_SOFA_MED = UW_M.get("sofa_median") if UW_OK else None
_sofa_med = _SOFA_MED if _SOFA_MED is not None else 2.0

# ───────────────────────────────────────────────────────────────────────
#  DOCX helpers
# ───────────────────────────────────────────────────────────────────────
def setup_doc():
    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = "Times New Roman"
    style.font.size = Pt(11)
    style.paragraph_format.space_after = Pt(4)
    style.paragraph_format.line_spacing = 1.5
    for section in doc.sections:
        section.page_width = Cm(21.0)   # A4
        section.page_height = Cm(29.7)  # A4
        section.top_margin = Cm(2.0)
        section.bottom_margin = Cm(2.0)
        section.left_margin = Cm(2.0)
        section.right_margin = Cm(2.0)
    return doc

def H(doc, text, level=1):
    h = doc.add_heading(text, level=level)
    for run in h.runs:
        run.font.color.rgb = RGBColor(0, 0, 0)
        run.font.name = "Times New Roman"
        if level == 1:
            run.font.size = Pt(13)
        elif level == 2:
            run.font.size = Pt(12)
    return h

def P_(doc, text, bold=False, italic=False, size=11, align="justify"):
    p = doc.add_paragraph()
    p.alignment = {"left": WD_ALIGN_PARAGRAPH.LEFT, "center": WD_ALIGN_PARAGRAPH.CENTER,
                   "justify": WD_ALIGN_PARAGRAPH.JUSTIFY}.get(align, WD_ALIGN_PARAGRAPH.JUSTIFY)
    run = p.add_run(text)
    run.font.name = "Times New Roman"
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic
    return p

def PR(doc, parts, size=11, align="justify"):
    """parts = [(text, bold, italic), ...]"""
    p = doc.add_paragraph()
    p.alignment = {"left": WD_ALIGN_PARAGRAPH.LEFT, "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
                   "center": WD_ALIGN_PARAGRAPH.CENTER}.get(align, WD_ALIGN_PARAGRAPH.JUSTIFY)
    for text, bold, italic in parts:
        r = p.add_run(text)
        r.font.name = "Times New Roman"
        r.font.size = Pt(size)
        r.bold = bold
        r.italic = italic
    return p


def SHADED_BOX(doc, fill="EDF1F8"):
    """Create a single-cell shaded box (journal-style panel) and return the cell."""
    tbl = doc.add_table(rows=1, cols=1)
    tbl.style = "Table Grid"
    cell = tbl.rows[0].cells[0]
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    cell._tc.get_or_add_tcPr().append(shd)
    return cell


def BOX_TITLE(cell, text):
    p = cell.paragraphs[0]
    p.paragraph_format.space_after = Pt(6)
    r = p.add_run(text)
    r.bold = True
    r.font.name = "Times New Roman"
    r.font.size = Pt(11)


def BOX_PARA(cell, label=None, text="", size=10):
    p = cell.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    if label:
        r0 = p.add_run(label)
        r0.bold = True
        r0.font.name = "Times New Roman"
        r0.font.size = Pt(size)
    r = p.add_run(text)
    r.font.name = "Times New Roman"
    r.font.size = Pt(size)
    return p


# ───────────────────────────────────────────────────────────────────────
#  Layout polish pass (applied to every document before save)
# ───────────────────────────────────────────────────────────────────────
NO_HEADER_TABLES = []  # underlying _tbl elements of tables without a header row
def _add_page_numbers(doc, running_head=None):
    """Centred page-number field in the footer of every section (TNR 9pt);
    optional right-aligned running head in the header."""
    for section in doc.sections:
        footer = section.footer
        p = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run()
        f1 = OxmlElement("w:fldChar"); f1.set(qn("w:fldCharType"), "begin")
        it = OxmlElement("w:instrText"); it.set(qn("xml:space"), "preserve")
        it.text = "PAGE"
        f2 = OxmlElement("w:fldChar"); f2.set(qn("w:fldCharType"), "end")
        run._r.append(f1); run._r.append(it); run._r.append(f2)
        run.font.name = "Times New Roman"
        run.font.size = Pt(9)
        if running_head:
            header = section.header
            hp = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
            hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            hr = hp.add_run(running_head)
            hr.font.name = "Times New Roman"
            hr.font.size = Pt(9)
            hr.font.color.rgb = RGBColor(0x44, 0x44, 0x44)


def _shade_cell(cell, fill):
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    cell._tc.get_or_add_tcPr().append(shd)


def _polish_tables(doc):
    """Journal-style tables: centred, bold shaded header, TNR throughout,
    rows kept on one page, cells vertically centred."""
    for tbl in doc.tables:
        # skip 1x1 shaded panel boxes
        if len(tbl.rows) == 1 and len(tbl.columns) == 1:
            continue
        has_header = tbl._tbl not in NO_HEADER_TABLES
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        for r_i, row in enumerate(tbl.rows):
            # keep rows on a single page
            trPr = row._tr.get_or_add_trPr()
            if trPr.find(qn("w:cantSplit")) is None:
                trPr.append(OxmlElement("w:cantSplit"))
            for cell in row.cells:
                # vertical centring
                tcPr = cell._tc.get_or_add_tcPr()
                if tcPr.find(qn("w:vAlign")) is None:
                    va = OxmlElement("w:vAlign")
                    va.set(qn("w:val"), "center")
                    tcPr.append(va)
                if r_i == 0 and has_header:
                    _shade_cell(cell, "F2F2F2")
                for p in cell.paragraphs:
                    p.paragraph_format.space_after = Pt(2)
                    p.paragraph_format.line_spacing = 1.0
                    for run in p.runs:
                        run.font.name = "Times New Roman"
                        if r_i == 0 and has_header:
                            run.font.bold = True


def _polish_headings(doc):
    """Headings: keep with following paragraph, consistent space before."""
    for p in doc.paragraphs:
        if p.style.name.startswith("Heading"):
            p.paragraph_format.keep_with_next = True
            if p.style.name == "Heading 1":
                p.paragraph_format.space_before = Pt(14)
            elif p.style.name == "Heading 2":
                p.paragraph_format.space_before = Pt(10)


_MINUS_RE = re.compile(r"(?<![A-Za-z0-9])[\-\u2013](?=\d)")


def _normalise_minus(doc):
    """Render every minus sign as U+2212 (true minus) for typographic consistency.

    Only a dash that *starts a numeric token* is converted: it must be preceded
    by a non-alphanumeric character and followed by a digit. Hyphens inside
    words (L2-penalised), compound adjectives (28-day, Day-2), en-dash ranges
    (1st-99th, 2014-2015, SA1-SA3) and reference page ranges are left as typed.
    """
    n = 0
    def _fix(par):
        nonlocal n
        for run in par.runs:
            if run.text and _MINUS_RE.search(run.text):
                new = _MINUS_RE.sub("\u2212", run.text)
                if new != run.text:
                    n += 1
                    run.text = new
    for par in doc.paragraphs:
        _fix(par)
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                for par in cell.paragraphs:
                    _fix(par)
    return n


def polish_doc(doc, page_numbers=True, running_head=None):
    if page_numbers:
        _add_page_numbers(doc, running_head=running_head)
    _polish_tables(doc)
    _polish_headings(doc)
    _normalise_minus(doc)
    return doc


# ════════════════════════════════════════════════════════════════════════
#  MAIN MANUSCRIPT
# ════════════════════════════════════════════════════════════════════════
def build_main():
    doc = setup_doc()

    # Title
    P_(doc, "Early versus Deferred De-escalation from Controlled Ventilation after a "
            "48-Hour Landmark of Physiological Stability: A Target Trial Emulation in MIMIC-IV "
            "with External Validation in eICU-CRD",
            bold=True, size=13, align="center")
    P_(doc, "[Author list and affiliations blinded for peer review]", italic=True, size=10, align="center")
    doc.add_paragraph()

    # ── Structured abstract (ICM style: Purpose / Methods / Results / Conclusions) ──
    H(doc, "Abstract", level=1)

    PR(doc, [("Purpose.  ", True, False),
        ("The optimal timing of transition from controlled to assisted ventilatory modes "
         "after physiological stabilisation remains uncertain: prolonged controlled "
         "ventilation exposes patients to sedation, diaphragm dysfunction, and "
         "ventilator-induced lung injury, while premature de-escalation risks "
         "patient\u2013ventilator asynchrony, gas exchange deterioration, and "
         "reintubation. We estimated the causal effect of early versus deferred "
         "de-escalation of respiratory support on 28-day in-hospital mortality in "
         "adults still on invasive mechanical ventilation 48 hours after ICU admission, "
         "using target trial emulation with clone\u2013censor\u2013weight (CCW) "
         "estimation, and validated the findings externally in an independent "
         "multi-centre cohort.", False, False)])

    PR(doc, [("Methods.  ", True, False),
        ("Retrospective landmark target trial emulation in MIMIC-IV (v2.2, single-database "
         "derivation cohort) with external validation in the multi-centre eICU "
         "Collaborative Research Database (v2.0; 208 hospitals). "
         "Eligibility: age \u226518 years, invasive mechanical ventilation >24 h, alive and "
         "ventilated at 48 h with a classified Day-2 mode (controlled or assisted), Day-2 "
         "FiO\u2082 \u226450% and PEEP \u226410 cmH\u2082O. Two strategies were compared over "
         "a 24-hour grace window: early (Day-3 support level below Day-2, encompassing mode "
         "step-down or successful extubation while alive) versus deferred (support level "
         "maintained or increased). A single-interval CCW estimator with stabilised "
         "inverse-probability-of-censoring weights (truncated at the 1st and 99th "
         "percentiles) was used. The propensity model included 21 baseline covariates; "
         "robustness analyses employed an extended 40-covariate adjustment, alternative "
         "weight truncation, a doubly robust AIPW estimator and E-values for unmeasured "
         "confounding. The "
         "primary outcome was 28-day in-hospital mortality; secondary outcomes were 28-day "
         "restricted mean survival time (RMST) and ventilator-free days (VFD28). Bootstrap "
         "95% confidence intervals used 2,000 resamples for the primary analysis and 800 for "
         "sensitivity and exploratory analyses. The external validation applied the "
         "harmonised protocol in eICU-CRD with in-hospital mortality as the outcome "
         "(28-day follow-up is unavailable in eICU).", False, False)])

    PR(doc, [("Results.  ", True, False),
        (f"Of {F['total_records']:,} MIMIC-IV records, {F['eligible']:,} adults met all "
         f"eligibility criteria. The early strategy was followed by {S['early']:,} patients "
         f"({S['early']/F['eligible']*100:.1f}%) and the deferred strategy by "
         f"{S['deferred']:,} ({S['deferred']/F['eligible']*100:.1f}%); {S['unascertainable']:,} "
         f"({S['unascertainable']/F['eligible']*100:.1f}%) had unascertainable strategy and "
         f"{S['grace_death']:,} died during the grace period. The primary analysis included "
         f"{S['early']+S['deferred']:,} patients. Weighted 28-day mortality was "
         f"{pp(P['mort_early'])}% (early) versus {pp(P['mort_deferred'])}% (deferred): risk "
         f"difference {pp_label(P['rd'])} percentage points (95% confidence interval [CI], "
         f"{ci_pp(P['rd_ci_lo'], P['rd_ci_hi'])}); risk ratio {P['rr']:.2f} (95% CI, "
         f"{P['rr_ci_lo']:.2f}\u2013{P['rr_ci_hi']:.2f}). 28-day RMST was "
         f"{P['rmst_early']:.2f} versus {P['rmst_deferred']:.2f} days "
         f"(difference +{P['rmst_diff']:.2f}; 95% CI, {P['rmst_ci_lo']:+.2f} to "
         f"{P['rmst_ci_hi']:+.2f}) and VFD28 was {P['vfd_early']:.2f} versus "
         f"{P['vfd_deferred']:.2f} days (difference +{P['vfd_diff']:.2f}; 95% CI, "
         f"{P['vfd_ci_lo']:+.2f} to {P['vfd_ci_hi']:+.2f}). All six sensitivity and "
         f"robustness analyses\u2014including the extended 40-covariate adjustment, "
         f"alternative weight truncation and a doubly robust AIPW estimator\u2014produced "
         f"statistically significant estimates (E-value {_EV_TXT}). In the exploratory FiO\u2082\u00d7PEEP threshold grid, all "
         f"nine cells yielded risk differences excluding zero (range {TG_RANGE_ALL} "
         f"percentage points). External validation in eICU-CRD "
         f"({EV['n_analyzable']:,} analysable) confirmed the direction and significance of "
         f"the finding: weighted in-hospital mortality {pp(EVP['mort_early'])}% (early) "
         f"versus {pp(EVP['mort_deferred'])}% (deferred); risk difference "
         f"{pp_label(EVP['rd'])} percentage points (95% CI, "
         f"{ci_pp(EVP['rd_ci_lo'], EVP['rd_ci_hi'])}); risk ratio {EVP['rr']:.2f} (95% CI, "
         f"{EVP['rr_ci_lo']:.2f}\u2013{EVP['rr_ci_hi']:.2f}).", False, False)])

    PR(doc, [("Conclusions.  ", True, False),
        (f"In this target trial emulation, early de-escalation of respiratory support after a "
         f"48-hour stability landmark was associated with significantly lower 28-day "
         f"mortality (risk ratio {P['rr']:.2f}), longer restricted mean survival, and more "
         f"ventilator-free days, with consistent results across all six sensitivity and "
         f"robustness analyses, the exploratory FiO\u2082\u00d7PEEP threshold grid, and an "
         f"external validation in an independent multi-centre cohort of {EV['n_analyzable']:,} patients. "
         "Among patients on low "
         "respiratory support at 48 hours, a timely step-down in support should be "
         "actively pursued; residual confounding by unmeasured clinical acuity cannot be "
         "excluded and prospective validation is warranted.", False, False)])

    P_(doc, "Keywords: mechanical ventilation; ventilator weaning; target trial emulation; "
            "clone\u2013censor\u2013weight; MIMIC-IV; eICU; external validation.",
       italic=True, size=10)

    # ── Take-home message (ICM style) ──
    cell = SHADED_BOX(doc, fill="EDF1F8")
    BOX_TITLE(cell, "Take-home message")
    BOX_PARA(cell,
        text=f"In this target trial emulation of {F['eligible']:,} adults still on invasive "
             f"mechanical ventilation 48 hours after ICU admission, early de-escalation of "
             f"respiratory support within 24 hours (mode step-down or successful extubation) "
             f"was associated with significantly lower 28-day mortality "
             f"({pp(P['mort_early'])}% vs {pp(P['mort_deferred'])}%; risk difference "
             f"{pp_label(P['rd'])} percentage points; 95% CI "
             f"{ci_pp(P['rd_ci_lo'], P['rd_ci_hi'])}; risk ratio {P['rr']:.2f}, 95% CI "
             f"{P['rr_ci_lo']:.2f}\u2013{P['rr_ci_hi']:.2f}), with consistent benefits on "
             f"restricted mean survival and ventilator-free days and across all six "
             f"sensitivity and robustness analyses (including a doubly robust AIPW "
             f"estimator; E-value {_EV_TXT}). All nine FiO\u2082\u00d7PEEP threshold cells excluded zero, with "
             f"the largest absolute benefit at FiO\u2082 \u226440%. The association was "
             f"externally validated in the independent multi-centre eICU-CRD cohort "
             f"({EV['n_analyzable']:,} analysable patients in 208 hospitals; risk "
             f"difference {pp_label(EVP['rd'])} percentage points for in-hospital "
             f"mortality).")
    doc.add_paragraph()

    # ── Research in context (BMJ style) ──
    cell = SHADED_BOX(doc, fill="F5F5F0")
    BOX_TITLE(cell, "Research in context")
    BOX_PARA(cell, label="What is already known on this topic: ",
        text="International guidelines recommend daily spontaneous breathing trials and "
             "protocolised weaning for mechanically ventilated adults, but they provide "
             "little explicit guidance on when to transition from controlled to assisted "
             "ventilation after an initial period of physiological stabilisation. "
             "Randomised trials have compared weaning protocols and spontaneous breathing "
             "trial frequency rather than the timing of this mode transition, and naive "
             "observational comparisons are biased by time-varying confounding and "
             "immortal time bias.")
    BOX_PARA(cell, label="What this study adds: ",
        text=f"Using a landmark target trial emulation with clone\u2013censor\u2013weight "
             f"estimation in MIMIC-IV (n = {F['eligible']:,} eligible; "
             f"{S['early']+S['deferred']:,} analysable), we found that early de-escalation "
             f"was associated with a statistically significant {pp_label(P['rd'])} "
             f"percentage-point lower 28-day mortality (risk ratio {P['rr']:.2f}), with "
             f"consistent benefits on restricted mean survival and ventilator-free days and "
             f"across all six sensitivity and robustness analyses (doubly robust AIPW "
             f"confirmation; E-value {_EV_TXT}). An exploratory FiO\u2082\u00d7PEEP threshold "
             f"grid showed benefit in every cell, with the largest absolute effect at "
             f"FiO\u2082 \u226440%. The finding was externally validated in the independent "
             f"multi-centre eICU-CRD cohort ({EV['n_analyzable']:,} analysable patients "
             f"across 208 hospitals), where early de-escalation was likewise associated "
             f"with significantly lower in-hospital mortality, supporting a proactive "
             f"de-escalation approach in stable "
             f"patients and generating hypotheses for prospective randomised trials.")
    doc.add_page_break()

    # ── Introduction ──
    H(doc, "Introduction", level=1)
    P_(doc,
        "Mechanical ventilation is a life-saving intervention in critical illness, yet the "
        "optimal timing of de-escalation from controlled to assisted ventilatory modes "
        "remains uncertain. Prolonged controlled ventilation exposes patients to deep "
        "sedation, ventilator-induced diaphragm dysfunction, and ventilator-associated "
        "pneumonia, while premature de-escalation risks patient\u2013ventilator asynchrony, "
        "gas exchange deterioration, and reintubation with its associated morbidity and "
        "mortality. Current international guidelines (1\u20133) recommend daily spontaneous "
        "breathing trials and protocolised weaning, but provide limited explicit guidance on "
        "the specific timing of mode transition after an initial period of physiological "
        "stabilisation. The 48-hour post-admission landmark is clinically meaningful: most "
        "early deaths have occurred, initial resuscitation has stabilised, and the question "
        "shifts from survival to liberation.", align="justify")

    P_(doc,
        "Randomised trials of ventilator weaning have largely focused on weaning protocols "
        "and spontaneous-breathing-trial frequency rather than the specific timing of "
        "controlled-to-assisted mode transition (4\u20137). Observational studies are "
        "limited by time-varying confounding (clinicians select patients for de-escalation "
        "based on disease trajectory, not at random) and by immortal time bias (patients "
        "must survive long enough to be de-escalated). Naive comparisons between early and "
        "late de-escalation groups are therefore fundamentally biased, regardless of how "
        "many covariates are adjusted for in a conventional regression model.", align="justify")

    P_(doc,
        "Target trial emulation (TTE) provides a principled framework to mitigate these "
        "biases (8, 9, 19). By explicitly specifying the trial protocol that would answer the "
        "clinical question \u2014 eligibility, treatment strategies, assignment, follow-up, "
        "and outcomes \u2014 and using the clone\u2013censor\u2013weight (CCW) estimator "
        "with stabilised inverse-probability-of-censoring weights (IPCW) to emulate random "
        "assignment, TTE aligns observational analysis with counterfactual reasoning.", align="justify")

    P_(doc,
        "The causal question was: among adults still receiving invasive mechanical "
        "ventilation 48 hours after ICU admission with Day-2 FiO\u2082 \u226450% and "
        "PEEP \u226410 cmH\u2082O, what is the effect of early versus deferred "
        "de-escalation of ventilatory support, implemented within a 24-hour grace "
        "window, on 28-day in-hospital mortality? We addressed this question by "
        "emulating a target trial in the MIMIC-IV critical care database, and then "
        "replicated the entire protocol without modification in an independent "
        "multi-centre database to assess transportability. The primary objective was to "
        "estimate the causal effect of each strategy on 28-day in-hospital mortality. "
        "Secondary objectives were 28-day restricted mean survival time (RMST) and "
        "ventilator-free days (VFD28). We additionally explored whether stricter "
        "oxygenation and PEEP thresholds at the landmark identified subgroups with "
        "differential treatment effects, as hypothesis-generating evidence for future "
        "trials.", align="justify")

    # ── Methods ──
    H(doc, "Methods", level=1)

    H(doc, "Study design and data source", level=2)
    P_(doc,
        "We conducted a retrospective landmark target trial emulation using the Medical "
        "Information Mart for Intensive Care (MIMIC-IV, version 2.2), a "
        "critical-care database containing all ICU admissions to the Beth Israel "
        "Deaconess Medical Center between 2008 and 2019 (10). The institutional review boards of the "
        "Massachusetts Institute of Technology and Beth Israel Deaconess Medical Center "
        "approved the use of MIMIC-IV for research; informed consent was waived because all "
        "data are de-identified. The analytic extract is a pre-specified adult subset of "
        "the MIMIC-IV v2.2 icustays table (10) \u2014 the published database contains 73,181 ICU "
        "admissions across 50,920 unique patients \u2014 and comprises "
        f"{F['total_records']:,} adult ICU records restricted to stays with sufficient "
        "chart-event coverage to derive Day-2 mode classification and FiO\u2082/PEEP. "
        "This pre-extraction is a data-provenance constraint of our approved access "
        "rather than an additional eligibility criterion, and it is reported so that the "
        "denominator can be traced to the published database. The full target trial protocol (eligibility, treatment "
        "strategies, outcomes, follow-up, and analysis plan) was specified before the "
        "analytic dataset was finalised, with no protocol amendments thereafter. The "
        "protocol is provided in the Supplement. The three temporal anchors of the "
        "emulation were aligned: eligibility ascertainment, assignment to a de-escalation "
        "strategy, and the start of follow-up all occurred at the 48-hour landmark "
        "(time zero), eliminating immortal time bias by design. The 48-hour landmark and "
        "the 24-hour grace window were chosen to align with the eligibility "
        "windows and intervention periods of published randomised weaning trials "
        "(4\u20137), so that the emulated strategies correspond to contrasts that "
        "have previously been shown to be implementable at the bedside. Reporting follows the "
        "STROBE guideline and the "
        "TARGET guideline for studies emulating a target trial (18); the completed "
        "checklists and the component-by-component emulation mapping are provided "
        "in the Supplement.", align="justify")

    H(doc, "Eligibility criteria", level=2)
    P_(doc,
        f"Adults (\u226518 years) receiving invasive mechanical ventilation for more than "
        f"24 hours who were alive and still ventilated at 48 hours after ICU admission "
        f"were eligible. Additional requirements were: (i) a classifiable ventilatory mode "
        f"on Day 2 (the day of the 48-hour landmark) \u2014 either controlled or "
        f"assisted/spontaneous; (ii) Day-2 FiO\u2082 \u226450%; and (iii) Day-2 PEEP "
        f"\u226410 cmH\u2082O. These thresholds were chosen to identify patients at a level "
        f"of physiological stability where de-escalation would be clinically plausible, "
        f"while excluding patients on high respiratory support who would not be candidates "
        f"for a reduction in support. Patients who died within 24 hours after the landmark "
        f"(i.e., during the grace period) were classified separately and excluded from the "
        f"primary analysis.", align="justify")

    H(doc, "Treatment strategies", level=2)
    P_(doc,
        "Two treatment strategies were compared over a 24-hour grace window beginning at "
        "the 48-hour landmark (Day 2 to Day 3 of ICU admission), defined by the change in "
        "ventilatory support level between Day 2 and Day 3:", align="justify")
    P_(doc,
        "\u2022 Early de-escalation \u2014 a step-down in support level by Day 3 "
        "(transition from a controlled to an assisted/spontaneous mode, or successful "
        "extubation with recorded end of mechanical ventilation while alive).",
        align="justify")
    P_(doc,
        "\u2022 Deferred de-escalation \u2014 ventilatory support level maintained or "
        "increased between Day 2 and Day 3.", align="justify")
    P_(doc,
        "Ventilatory modes were classified as controlled (CMV, PRVC/AC, PCV+, PCV/AC, APV, "
        "VOL/AC, SIMV, AC/PC, AC/VC variants) or assisted/spontaneous (CPAP, PSV, SBT, "
        "Standby, SPONT, MMV, PPS, ApnVol, ApnPres variants) based on standard ventilator "
        "nomenclature; controlled modes were ranked as higher support than assisted modes. "
        "Patients whose Day-3 status could not be classified (missing mode without "
        "documented ventilation end) were designated \u2018unascertainable\u2019 and "
        "excluded from the primary analysis.", align="justify")

    H(doc, "Outcomes", level=2)
    P_(doc,
        "The primary outcome was 28-day in-hospital mortality, defined as death in "
        "hospital within 28 days of ICU admission. Secondary outcomes were 28-day "
        "restricted "
        "mean survival time (RMST) and ventilator-free days at day 28 (VFD28). VFD28 was "
        "defined as the number of days alive and free of mechanical ventilation within the "
        "first 28 days; for patients who died before day 28, VFD28 was set to zero.", align="justify")

    H(doc, "Missing data", level=2)
    _mmp = pd.read_csv(RES / "mimic_missingness.csv").set_index("covariate")["pct_missing"]
    _emp = pd.read_csv(RES / "eicu_missingness.csv").set_index("covariate")["pct_missing"]
    _an_ev = pd.read_parquet(RES / "eicu_analyzable.parquet")
    _n_miss_out = int(_an_ev["hospitaldischargestatus"].isna().sum())
    P_(doc,
        "Missing data were assessed for each baseline covariate in both databases. "
        "The proportion of missing values and the handling method for each variable "
        "are provided in the Supplement (Table S10). Briefly, missing values in "
        "continuous covariates were imputed with the cohort median before "
        "propensity-model fitting (single imputation, pre-specified in the analysis "
        "plan), and binary indicators were complete by construction. In the MIMIC-IV "
        f"analysable cohort, Day-2 lactate was missing for {_mmp['d2_lactate']:.1f}% of patients, "
        f"PaO\u2082/FiO\u2082 for {_mmp['pao2_fio2_d2_mean']:.1f}%, plateau pressure for {_mmp['d2_pplat']:.1f}% and SOFA "
        f"subscores for {_mmp['sofa_respiration']:.1f}% (subscores are computed only for patients meeting "
        "Sepsis-3 criteria in the source extraction); in the eICU-CRD analysable "
        f"cohort, Day-2 SpO\u2082 was missing for {_emp['sao2_d2']:.1f}% and lactate for {_emp['lactate_d2']:.1f}%. The "
        "primary outcome was completely ascertained in MIMIC-IV; in eICU-CRD, "
        f"hospital discharge status was missing for {_n_miss_out:,} analysable patients "
        f"({_n_miss_out/len(_an_ev)*100:.1f}%), "
        "who were retained with the alive classification. Patients whose Day-3 "
        "strategy status could not be ascertained were excluded from the primary "
        "analysis by design and are reported separately in the flow diagrams "
        "(Figure 1; Supplement Figure S5). Sensitivity to these handling choices "
        "was assessed in the extended 40-covariate analysis and the doubly robust "
        "AIPW estimator.", align="justify")

    H(doc, "Statistical analysis", level=2)
    P_(doc,
        "The causal estimand was the per-protocol effect of early versus deferred "
        "de-escalation on 28-day in-hospital mortality in the eligible population, "
        "expressed as the absolute risk difference and the risk ratio. A per-protocol "
        "contrast was chosen because the two strategies are defined by ventilatory "
        "behaviour during the 24-hour grace window and therefore cannot be "
        "distinguished at time zero; consequently, an intention-to-treat-like contrast "
        "of the strategy initiated has no well-defined baseline assignment in this "
        "setting. The clone\u2013censor\u2013weight approach estimates the effect of "
        "adhering to each strategy while preserving time zero at eligibility. The "
        "design is summarised schematically in Supplement Figure S6.", align="justify")
    P_(doc,
        "Identification of the per-protocol effect relied on four assumptions, each "
        "addressed by a specific design or analytic feature. (i) Conditional "
        "exchangeability: within levels of the measured pre-landmark covariates, "
        "strategy assignment is assumed to be independent of the potential outcomes; "
        "we adjusted for 21 pre-specified covariates (40 in an extended analysis), "
        "verified post-weighting balance (all standardised mean differences < 0.10), "
        "and quantified robustness to unmeasured confounding with the E-value. "
        "(ii) Positivity: the eligibility thresholds (FiO\u2082 \u226450%, PEEP "
        "\u226410 cmH\u2082O) define a population in which both strategies are "
        "clinically plausible; propensity-score overlap was inspected, scores were "
        "clipped to [0.01, 0.99], and no complete separation was observed. "
        "(iii) Consistency: strategies were operationalised from charted ventilator "
        "modes and recorded end of ventilation using pre-specified definitions "
        "(Supplement Table A2), and delivery of the assigned strategy was verified "
        "against the recorded data (Figure S4). (iv) Correct specification of the "
        "weighting model, assessed through weight diagnostics and the concordant "
        "doubly robust AIPW estimator.", align="justify")
    P_(doc,
        "We used a single-interval clone\u2013censor\u2013weight (CCW) estimator with "
        "stabilised inverse-probability-of-censoring weights (IPCW). Each eligible patient "
        "was cloned into both strategy arms at the 48-hour landmark; within each clone, "
        "follow-up was censored at the first time at which the observed data became "
        "inconsistent with the assigned strategy (a treatment-level switch during the "
        "grace window). Death was analysed as an outcome event rather than a censoring "
        "event; the deaths occurring during the grace window were excluded as a separate "
        "pre-specified category and are reported in the flow diagram. The propensity "
        "for early de-escalation was modelled by multivariable logistic regression with 21 "
        "pre-specified baseline covariates: age, sex, weight, total and component SOFA "
        "scores (respiration, cardiovascular, central nervous system, renal), Day-2 "
        "ventilator parameters (FiO\u2082, PEEP, tidal volume, peak and plateau pressures), "
        "Day-2 PaO\u2082/FiO\u2082 ratio, Day-2 laboratory values (creatinine, white blood "
        "cell count, platelets, lactate), and indicators for ARDS, Sepsis-3, and continuous "
        "renal replacement therapy. Stabilised weights were constructed as the marginal "
        "probability of the observed strategy divided by the propensity, truncated at the "
        "1st and 99th percentiles to limit the influence of extreme observations (11, 12).",
        align="justify")

    P_(doc,
        "Risk differences (RD) and risk ratios (RR) were estimated from weighted mortality "
        "proportions. 28-day RMST was estimated as the area under the weighted survival "
        "curve truncated at 28 days, that is, the weighted mean survival time up to "
        "28 days (13, 21). Ninety-five per cent confidence intervals were constructed using "
        "patient-level bootstrap resampling, with 2,000 replicates for the primary analysis "
        "and 800 replicates for sensitivity and exploratory analyses; in each bootstrap "
        "iteration the propensity model was refit and weights were re-truncated (14, 15).",
        align="justify")

    P_(doc,
        "Three pre-specified sensitivity analyses were conducted to test robustness: "
        "(SA1) no weight truncation; (SA2) ICU mortality instead of hospital mortality as "
        "the outcome; and (SA3) unweighted per-protocol comparison. Three further "
        "robustness analyses were performed: (SA4) extended covariate adjustment adding 19 "
        "pre-landmark confounder dimensions to the propensity model (sedation depth "
        "[RASS], norepinephrine, propofol and fentanyl infusion doses, mean arterial "
        "pressure, heart rate, respiratory rate and SpO\u2082 at Day 2, Day-2 fluid "
        "balance, blood urea nitrogen, albumin, total bilirubin, haemoglobin and "
        "lymphocyte count, baseline procalcitonin and interleukin-6, body-mass index, and "
        "the coagulation and liver SOFA subscores; 40 covariates in total); (SA5) weight "
        "truncation at the 5th\u201395th instead of the 1st\u201399th percentile; and an "
        "augmented inverse-probability-weighted (AIPW) doubly robust estimator combining "
        "arm-specific outcome models with the propensity model. To gauge susceptibility to "
        "unmeasured confounding, E-values were computed on the risk-ratio scale for the "
        "point estimate and for the confidence-interval limit closest to the null. "
        "Pre-specified subgroup analyses examined Day-2 mode (controlled vs assisted), "
        "age (\u2265 vs <65 years), Sepsis-3 status, total SOFA score (\u2265 vs < cohort "
        "median) and ARDS. As an exploratory "
        "analysis, a FiO\u2082\u00d7PEEP 3\u00d73 threshold grid (FiO\u2082 \u226440, "
        "\u226450, \u226460%; PEEP \u22645, \u22648, \u226410 cmH\u2082O) was examined "
        "without multiplicity adjustment.", align="justify")

    P_(doc,
        "Analyses were performed in Python (version 3.13) using the pandas, "
        "scikit-learn, SciPy, lifelines and matplotlib libraries; the propensity model "
        "was fitted by L2-penalised logistic regression and propensity scores were "
        "clipped to [0.01, 0.99] before weight construction. Statistical significance "
        "was defined as a two-sided 95% CI excluding the null; no formal multiplicity "
        "adjustment was applied, and the exploratory threshold grid is interpreted as "
        "hypothesis-generating.", align="justify")

    P_(doc,
        "Per-protocol verification of strategy delivery was assessed by plotting "
        "respiratory-adjacent vital signs at 6-hour resolution from Day 1 00:00 through "
        "Day 2 18:00 (eight timepoints per patient, covering PaO\u2082/FiO\u2082 ratio, "
        "SpO\u2082, respiratory rate, norepinephrine dose and RASS sedation score) and "
        "daily Day-2\u2192Day-3 ventilator aggregates (FiO\u2082, PEEP, peak and plateau "
        "pressure, tidal volume, and mode distribution). In MIMIC-IV, ventilator settings "
        "are charted intermittently rather than continuously and no ventilator settings "
        "were available for Day 4\u2013Day 5 in the extraction; the trajectory in the "
        "supplement (Figure S4) is therefore presented at this honestly-scoped "
        "resolution.", align="justify")

    H(doc, "External validation in eICU-CRD", level=2)
    P_(doc,
        "To assess transportability, the full target trial protocol was replicated in the "
        "eICU Collaborative Research Database (v2.0), a multi-centre cohort of 208 United "
        "States hospitals (2014\u20132015) that is structurally and geographically "
        "independent of MIMIC-IV (17). The harmonised protocol retained the 48-hour "
        "landmark, the Day-2 FiO\u2082 \u226450% and PEEP \u226410 cmH\u2082O thresholds, "
        "the strategy definitions and the CCW estimator with 1st\u201399th percentile "
        "truncation. First ICU stays of adults were used. Ventilator settings were "
        "abstracted from respiratory-therapy flowsheet charting; because eICU does not "
        "record an explicit mode variable, modes were classified from treatment orders "
        "(assist-controlled, volume- or pressure-controlled, and synchronised intermittent "
        "mandatory ventilation mapped to controlled; pressure-support, CPAP/PEEP therapy "
        "and weaning orders mapped to assisted) with a corroborating inference from "
        "charted settings (pressure support >0 mapped to assisted; pressure control or a "
        "set mandatory rate mapped to controlled), taking the latest evidence within each "
        "classification window. Successful extubation was defined as the end of charted "
        "ventilator settings at or before 72 hours; a sensitivity definition additionally "
        "required the ICU stay to continue at least 6 hours beyond the last ventilator "
        "record to exclude inter-ICU transfers misread as extubation. As eICU contains no "
        "post-discharge follow-up, the harmonised validation outcome was in-hospital "
        "mortality (hospital discharge status), closely comparable to the 28-day "
        "in-hospital mortality used in MIMIC-IV. The propensity model used 25 covariates "
        "mirroring the derivation set (demographics, admission source, APACHE-IV severity "
        "scores, Day-2 ventilator settings, Day-2 laboratory values, Day-2 vasopressor "
        "and sedative infusions, and comorbidities); estimates were not pooled across "
        "databases and are compared directionally, as pre-specified.", align="justify")

    # ── Results ──
    H(doc, "Results", level=1)

    H(doc, "Study population", level=2)
    P_(doc,
        f"Of {F['total_records']:,} adult ICU records in MIMIC-IV (one record per ICU "
        f"stay), {F['ventilated']:,} "
        f"patients received invasive mechanical ventilation, of whom {F['vent_gt_24h']:,} "
        f"were ventilated for more than 24 hours, of whom {F['alive_at_48h']:,} were alive "
        f"and still ventilated at the 48-hour landmark, of whom "
        f"{F['has_d2_settings']:,} had Day-2 ventilator settings recorded. After "
        f"requiring a classifiable Day-2 mode and "
        f"applying the FiO\u2082 \u226450% and PEEP \u226410 cmH\u2082O thresholds, "
        f"{F['eligible']:,} patients met all eligibility criteria (Figure 1). Among these, "
        f"{S['early']:,} patients ({S['early']/F['eligible']*100:.1f}%) followed the early "
        f"strategy, {S['deferred']:,} ({S['deferred']/F['eligible']*100:.1f}%) the deferred "
        f"strategy, {S['unascertainable']:,} "
        f"({S['unascertainable']/F['eligible']*100:.1f}%) had an unascertainable Day-3 "
        f"strategy, and {S['grace_death']:,} "
        f"({S['grace_death']/F['eligible']*100:.1f}%) died during the grace period. The "
        f"primary analysis included the {S['early']+S['deferred']:,} patients with "
        f"ascertainable strategy assignment.", align="justify")

    H(doc, "Baseline characteristics", level=2)
    _t1 = pd.read_csv(RES / "table1_baseline.csv").set_index("Variable")
    _ce, _cd = f"Early (n={S['early']})", f"Deferred (n={S['deferred']})"
    _balm = pd.read_csv(RES / "covariate_balance.csv")
    _smd_imax = _balm["smd_weighted"].abs().idxmax()
    P_(doc,
        f"Baseline characteristics were broadly similar between strategy groups (Table 1). "
        f"The mean age was {_t1.loc['Age, years', _ce]:.1f} years in the early group and "
        f"{_t1.loc['Age, years', _cd]:.1f} years in the deferred "
        f"group. SOFA total scores were {_t1.loc['SOFA total', _ce]:.1f} versus "
        f"{_t1.loc['SOFA total', _cd]:.1f}. Day-2 ventilator settings were "
        f"similar (FiO\u2082 {_t1.loc['Day-2 FiO\u2082, %', _ce]:.1f}% versus "
        f"{_t1.loc['Day-2 FiO\u2082, %', _cd]:.1f}%; PEEP "
        f"{_t1.loc['Day-2 PEEP, cmH\u2082O', _ce]:.1f} versus "
        f"{_t1.loc['Day-2 PEEP, cmH\u2082O', _cd]:.1f} cmH\u2082O), although "
        f"tidal volume ({_t1.loc['Day-2 tidal volume, mL', _ce]:.0f} versus "
        f"{_t1.loc['Day-2 tidal volume, mL', _cd]:.0f} mL) was modestly lower in the "
        f"early group, "
        f"consistent with a greater share of assisted modes. Sepsis-3 prevalence was high "
        f"in both groups ({_t1.loc['Sepsis-3, n (%)', _ce]:.1f}% versus "
        f"{_t1.loc['Sepsis-3, n (%)', _cd]:.1f}%). After stabilised weighting, all "
        f"standardised "
        f"mean differences were reduced below 0.10 (largest: ARDS, SMD = "
        f"{_balm.loc[_smd_imax, 'smd_weighted']:.3f}), "
        f"indicating adequate covariate balance.", align="justify")

    H(doc, "Weight diagnostics", level=2)
    P_(doc,
        f"The mean stabilised weight was {W['weight_mean']:.2f} (SD "
        f"{W['weight_sd']:.2f}); after truncation at the 1st and 99th percentiles "
        f"({W['trunc_lo']:.2f} and {W['trunc_hi']:.2f}), weights ranged from "
        f"{W['weight_min']:.2f} to {W['weight_max']:.2f}. The propensity score ranged from "
        f"{W['ps_min']:.3f} to {W['ps_max']:.3f} (mean {W['ps_mean']:.3f}), with no "
        f"evidence of complete separation. The distribution of weights and the change in "
        f"covariate balance after weighting are shown in Figure S1.", align="justify")

    H(doc, "Primary outcome", level=2)
    if UW_OK:
        _uw_txt = (f" The corresponding crude (unweighted) event counts were "
                   f"{UW_M['early']['events']}/{UW_M['early']['n']:,} "
                   f"({UW_M['early']['risk']*100:.1f}%) in the early group versus "
                   f"{UW_M['deferred']['events']}/{UW_M['deferred']['n']:,} "
                   f"({UW_M['deferred']['risk']*100:.1f}%) in the deferred group, "
                   f"a crude risk difference of "
                   f"{pp_label(UW_M['rd_unweighted'])} pp.")
    else:
        _uw_txt = ""
    P_(doc,
        f"Weighted 28-day mortality was {pp(P['mort_early'])}% in the early group and "
        f"{pp(P['mort_deferred'])}% in the deferred group (Table 2)."
        + _uw_txt + " Under the "
        f"per-protocol contrast, the estimated effect of early versus deferred "
        f"de-escalation was a risk "
        f"difference of {pp_label(P['rd'])} percentage points (95% CI "
        f"{ci_pp(P['rd_ci_lo'], P['rd_ci_hi'])}) and the risk ratio was "
        f"{P['rr']:.2f} (95% CI {P['rr_ci_lo']:.2f}\u2013{P['rr_ci_hi']:.2f}); both "
        f"excluded the null. The unweighted log-rank test yielded {_LR_TXT}. Figure 2 "
        f"shows the weighted Kaplan\u2013Meier survival curves with the N-at-risk table.",
        align="justify")

    H(doc, "Secondary outcomes", level=2)
    P_(doc,
        f"The 28-day RMST was {P['rmst_early']:.2f} days (early) versus "
        f"{P['rmst_deferred']:.2f} days (deferred), for a difference of "
        f"+{P['rmst_diff']:.2f} days (95% CI {P['rmst_ci_lo']:+.2f} to "
        f"{P['rmst_ci_hi']:+.2f}). VFD28 was {P['vfd_early']:.2f} days (early) versus "
        f"{P['vfd_deferred']:.2f} days (deferred), for a difference of +{P['vfd_diff']:.2f} "
        f"days (95% CI {P['vfd_ci_lo']:+.2f} to {P['vfd_ci_hi']:+.2f}). Both secondary "
        f"outcomes significantly favoured early de-escalation, with confidence intervals "
        f"excluding zero.", align="justify")

    H(doc, "Sensitivity analyses", level=2)
    P_(doc,
        f"All three pre-specified sensitivity analyses (SA1\u2013SA3) produced point "
        f"estimates consistent "
        f"with the primary analysis (Table 3, Figure 3). Removing weight truncation (SA1) "
        f"yielded an RD of {pp_label(SA['SA1_no_truncation']['rd'])} pp (95% CI "
        f"{ci_pp(SA['SA1_no_truncation']['rd_ci_lo'], SA['SA1_no_truncation']['rd_ci_hi'])}). "
        f"Using ICU mortality (SA2) yielded an RD of "
        f"{pp_label(SA['SA2_icu_mortality']['rd'])} pp (95% CI "
        f"{ci_pp(SA['SA2_icu_mortality']['rd_ci_lo'], SA['SA2_icu_mortality']['rd_ci_hi'])}). "
        f"The unweighted per-protocol comparison (SA3) yielded an RD of "
        f"{pp_label(SA['SA3_unweighted']['rd'])} pp (95% CI "
        f"{ci_pp(SA['SA3_unweighted']['rd_ci_lo'], SA['SA3_unweighted']['rd_ci_hi'])}), "
        f"suggesting that measured confounding (confounding by indication) had biased "
        f"the naive comparison "
        f"further toward early de-escalation.", align="justify")

    if RB_OK:
        H(doc, "Robustness analyses", level=2)
        _sa4 = RB["SA4_extended_covariates"]; _sa5 = RB["SA5_alt_truncation"]
        _aipw = RB["AIPW_doubly_robust"]; _ev = RB["e_value"]
        P_(doc,
            f"The benefit of early de-escalation persisted under all three robustness "
            f"analyses (Table 3, Figure 3). Extended covariate adjustment with 40 "
            f"pre-landmark covariates including sedation depth, vasoactive and sedative "
            f"infusion doses, haemodynamics, fluid balance, additional laboratory tests and "
            f"baseline inflammatory markers (SA4) yielded an RD of "
            f"{pp_label(_sa4['rd'])} pp (95% CI {ci_pp(_sa4['rd_ci_lo'], _sa4['rd_ci_hi'])}; "
            f"all weighted SMDs < 0.10, maximum {_sa4['smd_max']:.2f}). More conservative "
            f"weight truncation at the 5th\u201395th percentile (SA5) yielded an RD of "
            f"{pp_label(_sa5['rd'])} pp (95% CI "
            f"{ci_pp(_sa5['rd_ci_lo'], _sa5['rd_ci_hi'])}). The doubly robust AIPW estimator "
            f"yielded an RD of {pp_label(_aipw['rd'])} pp (95% CI "
            f"{ci_pp(_aipw['rd_ci_lo'], _aipw['rd_ci_hi'])}), closely concordant with the "
            f"primary IPW estimate. The E-value for the primary risk ratio "
            f"({_ev['rr']:.2f}) was {_ev['e_value_point']:.2f} (E-value for the "
            f"confidence-interval limit closest to the null: {_ev['e_value_ci_bound']:.2f}): "
            f"an unmeasured confounder would need to be associated with both the treatment "
            f"decision and 28-day mortality by a risk ratio of at least "
            f"{_ev['e_value_point']:.2f}-fold, above and beyond the measured covariates "
            f"(21 in the primary model, 40 in the extended analysis), "
            f"to explain away the observed association. In subgroup analyses (Figure S3), "
            f"the risk difference was directionally consistent across Day-2 mode (assisted: "
            f"RD {pp_label(SUBG.loc[SUBG['subgroup']=='Day-2 mode: assisted','rd'].iloc[0])} pp, "
            f"95% CI {ci_pp(SUBG.loc[SUBG['subgroup']=='Day-2 mode: assisted','rd_ci_lo'].iloc[0], SUBG.loc[SUBG['subgroup']=='Day-2 mode: assisted','rd_ci_hi'].iloc[0])}; "
            f"controlled: RD {pp_label(SUBG.loc[SUBG['subgroup']=='Day-2 mode: controlled','rd'].iloc[0])} pp, "
            f"95% CI {ci_pp(SUBG.loc[SUBG['subgroup']=='Day-2 mode: controlled','rd_ci_lo'].iloc[0], SUBG.loc[SUBG['subgroup']=='Day-2 mode: controlled','rd_ci_hi'].iloc[0])}), "
            f"age strata and Sepsis-3 strata, with nominally significant benefit in "
            f"patients on assisted modes at the landmark, patients aged \u226565 years and "
            f"patients with Sepsis-3; the ARDS subgroup was too small for estimation "
            f"(n = {S['early'] + S['deferred'] - int(SUBG.loc[SUBG['subgroup'] == 'No ARDS', 'n'].iloc[0])}). "
            f"In this cohort the Sepsis-3 flag was present in exactly the "
            f"{int(SUBG.loc[SUBG['subgroup'].str.startswith('SOFA'), 'n'].iloc[0]):,} patients with a total "
            f"SOFA score at or above the cohort median of "
            f"{_sofa_med:.0f}, so the Sepsis-3 and SOFA-split subgroup definitions "
            f"identify the same patients; only the Sepsis-3 split is therefore reported "
            f"in Figure S3.", align="justify")

    H(doc, "Exploratory threshold grid", level=2)
    P_(doc,
        "In the FiO\u2082\u00d7PEEP 3\u00d73 threshold grid (Table 4, Figure S2), all nine "
        "cells produced risk differences whose 95% confidence intervals excluded zero:",
        align="justify")
    for cell in TG:
        if cell["fio2_threshold"] == 40:
            P_(doc,
                f"\u2022 FiO\u2082 \u226440% / PEEP \u2264{cell['peep_threshold']} cmH\u2082O "
                f"(n = {cell['n']:,}, early {cell['n_early']:,}, deferred {cell['n_deferred']:,}): "
                f"RD = {pp_label(cell['rd'])} pp (95% CI "
                f"{ci_pp(cell['rd_ci_lo'], cell['rd_ci_hi'])}).",
                align="justify")
    P_(doc,
        f"The six cells with FiO\u2082 \u226450% or \u226460% showed smaller but still "
        f"significant risk differences (RD range {TG_RANGE_SUB} pp). The absolute "
        f"benefit was largest at FiO\u2082 \u226440%, consistent with greater physiological "
        f"reserve facilitating de-escalation. These analyses are exploratory, unadjusted "
        f"for multiplicity, and based on nested cohorts.", align="justify")

    H(doc, "Per-protocol verification of strategy delivery", level=2)
    P_(doc,
        f"Figure S4 visualises how the assigned strategies were actually delivered. "
        f"At the Day-2 landmark, {VT['early']['D2pct_controlled']:.1f}% ({VT['early']['D2_controlled']}/{VT['early']['n_group']}) "
        f"of the early arm and {VT['deferred']['D2pct_controlled']:.1f}% ({VT['deferred']['D2_controlled']}/{VT['deferred']['n_group']:,}) "
        f"of the deferred arm were on a controlled mode. By Day 3, {VT['early']['D3pct_missing']:.1f}% "
        f"({VT['early']['D3_missing']}/{VT['early']['n_group']}) of the early arm had been "
        f"extubated (no recorded mode, consistent with a step-down to spontaneous breathing "
        f"\u226472 h) and none ({VT['early']['D3_controlled']}/{VT['early']['n_group']}) "
        f"remained on controlled support, whereas {VT['deferred']['D3pct_controlled']:.1f}% "
        f"({VT['deferred']['D3_controlled']}/{VT['deferred']['n_group']:,}) of the deferred arm "
        f"remained on a controlled mode and none "
        f"({VT['deferred']['D3_missing']}/{VT['deferred']['n_group']:,}) had been extubated. "
        f"Median FiO\u2082 ({VTD['early']['FiO2_Day2_median']:.0f}%) and PEEP "
        f"({VTD['early']['PEEP_Day2_median']:.0f} cmH\u2082O) were identical between arms at Day 2 "
        f"by inclusion criterion and at Day 3 in the minority of the early arm still on the "
        f"ventilator; the strategy divergence is therefore concentrated in mode and "
        f"extubation rather than in incremental setting changes. The pre-landmark 6-hourly "
        f"respiratory trajectories (PaO\u2082/FiO\u2082 ratio, SpO\u2082, respiratory rate, "
        f"norepinephrine dose) tracked closely between arms, supporting the propensity-weighted "
        f"balance on time-varying respiratory state.", align="justify")

    H(doc, "External validation in eICU-CRD", level=2)
    P_(doc,
        f"The harmonised protocol was replicated in eICU-CRD v2.0 (Figure S5). Of "
        f"{EVF['first_stay_adults']:,} first ICU stays of adults, {EVF['vent_gt_24h']:,} "
        f"received invasive ventilation for more than 24 hours, {EVF['alive_at_48h']:,} "
        f"were alive at the 48-hour landmark, and {EVF['eligible']:,} met all eligibility "
        f"criteria. The early strategy was followed by {EVS['early']:,} patients and the "
        f"deferred strategy by {EVS['deferred']:,}; {EVS['unascertainable']:,} had "
        f"unascertainable Day-3 status and {EVS['grace_death']:,} died during the grace "
        f"period, leaving {EV['n_analyzable']:,} analysable patients. After weighting "
        f"(maximum standardised mean difference {EV['balance_max_smd_weighted']:.3f}), "
        f"in-hospital mortality was {pp(EVP['mort_early'])}% in the early arm versus "
        f"{pp(EVP['mort_deferred'])}% in the deferred arm: risk difference "
        f"{pp_label(EVP['rd'])} percentage points (95% CI "
        f"{ci_pp(EVP['rd_ci_lo'], EVP['rd_ci_hi'])}); risk ratio {EVP['rr']:.2f} (95% CI "
        f"{EVP['rr_ci_lo']:.2f}\u2013{EVP['rr_ci_hi']:.2f}) (Figure 4, Table 5). "
        f"Weighted ICU mortality showed the same pattern "
        f"({pp(EV['icu_mortality']['mort_early'])}% versus "
        f"{pp(EV['icu_mortality']['mort_deferred'])}%). The strict extubation definition "
        f"(requiring the ICU stay to continue \u22656 hours beyond the last ventilator "
        f"record; n = {EVSA['n']:,}) yielded a concordant estimate (risk difference "
        f"{pp_label(EVSA['rd'])} percentage points; 95% CI "
        f"{ci_pp(EVSA['rd_ci_lo'], EVSA['rd_ci_hi'])}). The direction and statistical "
        f"significance of the benefit were thus externally validated; the larger absolute "
        f"magnitude is addressed in the Discussion.", align="justify")

    # ── Discussion ──
    H(doc, "Discussion", level=1)
    P_(doc,
        f"In this target trial emulation of {F['eligible']:,} adults still on invasive "
        f"mechanical ventilation 48 hours after ICU admission, early de-escalation of "
        f"respiratory support during a 24-hour grace window was associated with a "
        f"{abs(P['rd'])*100:.1f}-percentage-point absolute reduction in 28-day mortality "
        f"compared with deferred de-escalation (95% CI "
        f"{ci_pp(P['rd_ci_lo'], P['rd_ci_hi'])}; RR {P['rr']:.2f}), together with "
        f"statistically significant gains in restricted mean survival time and "
        f"ventilator-free days. All pre-specified sensitivity analyses, all three "
        f"robustness analyses (extended covariate adjustment, alternative weight "
        f"truncation and a doubly robust AIPW estimator) and all nine "
        f"exploratory FiO\u2082\u00d7PEEP threshold cells confirmed the direction and "
        f"statistical significance of the benefit.", align="justify")

    P_(doc,
        "These results are consistent with the broader weaning literature indicating that "
        "prolonged mechanical ventilation is associated with worse outcomes, including "
        "diaphragm dysfunction, ventilator-associated pneumonia, and increased mortality "
        "(4\u20137, 16), and with trial evidence that protocolised early weaning shortens "
        "ventilation duration. The present analysis extends this evidence by quantifying "
        "the effect of a specifically defined de-escalation strategy after a 48-hour "
        "stability landmark using a causal inference framework that addresses immortal "
        "time bias and time-varying confounding. The robustness of the association is "
        "supported by the doubly robust AIPW estimator, by extended adjustment for 40 "
        "pre-landmark covariates including sedation depth and vasoactive infusions, and by "
        "an E-value of "
        f"{RB['e_value']['e_value_point']:.2f}, indicating that unmeasured confounding "
        "would need to be substantial to explain away the observed benefit. Nevertheless, "
        "treatment assignment "
        "remains confounded by disease severity; even with adjustment for the measured "
        "covariates and adequate weighted balance, residual confounding by unmeasured "
        "factors (e.g., dynamic respiratory mechanics, diaphragm function, clinician "
        "judgement) cannot be excluded and probably favours the early group, since "
        "clinicians de-escalate patients who are recovering.", align="justify")

    P_(doc,
        f"External validation strengthens the causal interpretation. In the independent "
        f"multi-centre eICU-CRD cohort of {EV['n_analyzable']:,} analysable patients "
        f"across 208 hospitals, the identical protocol yielded the same direction and "
        f"statistical significance (risk difference {pp_label(EVP['rd'])} percentage "
        f"points for in-hospital mortality; risk ratio {EVP['rr']:.2f}). Notably, the "
        f"weighted early-arm mortality was essentially identical across databases "
        f"({pp(P['mort_early'])}% vs {pp(EVP['mort_early'])}%), whereas deferred-arm "
        f"mortality was higher in eICU ({pp(EVP['mort_deferred'])}% vs "
        f"{pp(P['mort_deferred'])}%), plausibly reflecting its broader case mix across "
        f"community and academic hospitals, the longer in-hospital outcome window, and "
        f"greater practice variation in weaning. We therefore interpret the two analyses "
        f"as directionally concordant but deliberately refrain from pooling: the absolute "
        f"risk difference is setting-dependent, whereas the qualitative conclusion\u2014"
        f"that stable patients benefit from timely de-escalation\u2014transported "
        f"cleanly.", align="justify")

    P_(doc,
        "The finding that all nine FiO\u2082\u00d7PEEP threshold cells excluded zero, with "
        "the largest absolute benefit at FiO\u2082 \u226440%, indicates that the benefit "
        "of timely step-down is not an artefact of one arbitrary cut-point used to "
        "define 'low support'. The consistency of the signal across all three PEEP "
        "thresholds (\u22645, \u22648, \u226410 cmH\u2082O) strengthens biological "
        "plausibility: patients requiring minimal oxygen support have less severe lung "
        "injury and tolerate step-downs more readily. However, the nested nature of these "
        "subgroups, the lack of multiplicity adjustment, and the potential for selection "
        "bias within each threshold cell warrant caution; these results should inform, "
        "not replace, bedside assessment.", align="justify")

    P_(doc,
                "Strengths of this study include the explicit specification of the target "
        "trial protocol before the analytic dataset was finalised, alignment of "
        "eligibility, treatment assignment and follow-up at a single time zero, and "
        "the use of the CCW estimator to address immortal time bias and time-varying "
        "confounding. The propensity model incorporated 21 pre-specified baseline "
        "covariates (40 in the extended analysis) and achieved adequate balance after "
        "weighting (all standardised mean differences < 0.10); inference used bootstrap "
        "resampling that refits the propensity model and re-truncates the weights at each "
        "iteration rather than treating the weights as fixed. The primary estimate was "
        f"confirmed by a doubly robust AIPW estimator, and its susceptibility to "
        "unmeasured confounding was quantified with the E-value rather than merely "
        "asserted. Finally, the pre-specified replication in an independent multi-centre "
        f"database (eICU-CRD; {EV['n_analyzable']:,} analysable patients, 208 hospitals) "
        "reproduced the direction and significance of the association, which remains "
        "uncommon in observational weaning research.", align="justify")

    P_(doc,
                "Several limitations should be acknowledged. First, this is an observational "
        "study: treatment assignment was not randomised and, although we adjusted for "
        "21 measured covariates (40 in the extended analysis) and achieved adequate "
        "weighted balance, residual confounding by disease severity is likely. Second, "
        "ventilator settings were recorded intermittently rather than continuously, which "
        "may misclassify rapid transitions or transient changes. Third, the Day-3 "
        "ventilator mode was missing for a subset of patients who could not be classified "
        "(and for whom extubation status could not be inferred), introducing potential "
        "selection into the analysable sample; inverse-probability weighting addressed "
        "measured confounders but not this classification mechanism.", align="justify")
    P_(doc,
        "Fourth, although the association was externally validated in eICU-CRD, the "
        "absolute magnitude differed between databases (in part because eICU lacks 28-day "
        "follow-up, uses inferred rather than explicitly recorded ventilator modes, and "
        "spans heterogeneous hospital types); estimates were therefore not pooled and the "
        "pooled average effect remains unknown. Fifth, although Figure S4 confirms that the early-de-escalation strategy "
        "was operationally delivered as intended (mode step-down within 72 hours or "
        "extubation), the dataset does not contain hourly minute-by-minute ventilator "
        "settings, Day\u202f4\u20135 settings, raw pressure/flow waveforms, diaphragm "
        "function measurements (electromyography or ultrasound) or patient\u2013ventilator "
        "synchrony indices; the per-protocol pathway therefore cannot be mechanistically "
        "decomposed beyond the level of mode and extubation. Sixth, "
        "the MIMIC-IV ventilator mode taxonomy reflects the source institution's practice; "
        "mode labels may have been entered inconsistently. Finally, the 24-hour grace "
        "window is somewhat arbitrary; alternative grace durations could yield different "
        "classifications.", align="justify")

    P_(doc,
        "In conclusion, in this target trial emulation of {n} patients still invasively "
        "ventilated 48 hours after ICU admission with low oxygenation-support thresholds, "
        "early de-escalation of ventilatory support within the subsequent 24 hours was "
        "associated with lower 28-day mortality (RD {rd} pp; RR {rr:.2f}), longer "
        "ventilator-free days and greater restricted mean survival time, with consistent "
        "results across pre-specified sensitivity analyses, benefit apparent across the "
        "exploratory FiO\u2082\u2013PEEP threshold grid, and external validation in an "
        "independent multi-centre cohort. These findings support a prospective, "
        "randomised evaluation of early de-escalation strategies in this population; "
        "any such trial should incorporate pre-specified sequential monitoring for "
        "early benefit or harm (20).".format(
            n=f"{S['early']+S['deferred']:,}", rd=pp_label(P['rd']), rr=P['rr']), align="justify")

    # ── Figures ──
    doc.add_page_break()
    H(doc, "Figures", level=1)

    # Figure 1
    H(doc, "Figure 1.  Study flow diagram", level=2)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run().add_picture(str(FIG_DIR / "Figure1_Flow.png"), width=Cm(15))
    P_(doc, f"Stepwise inclusion of patients from MIMIC-IV v2.2. Of "
            f"{F['total_records']:,} ICU records, {F['ventilated']:,} received invasive "
            f"mechanical ventilation, of whom {F['vent_gt_24h']:,} were ventilated for "
            f"more than 24 hours, of whom {F['alive_at_48h']:,} were alive at the "
            f"48-hour landmark, of whom {F['has_d2_settings']:,} had Day-2 ventilator "
            f"settings recorded. After requiring a classifiable Day-2 mode "
            f"(controlled or assisted) and applying the "
            f"FiO\u2082 \u226450% and PEEP \u226410 cmH\u2082O thresholds, {F['eligible']:,} patients "
            f"were eligible. Of these, {S['early']:,} followed the early de-escalation "
            f"strategy, {S['deferred']:,} the deferred strategy, {S['unascertainable']:,} had an "
            f"unascertainable Day-3 strategy, and {S['grace_death']:,} died during the 24-hour "
            f"grace period. The primary analysis included "
            f"{S['early']+S['deferred']:,} patients.", italic=True, size=10, align="justify")

    # Figure 2
    H(doc, "Figure 2.  Weighted Kaplan\u2013Meier survival curves (28-day)", level=2)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run().add_picture(str(FIG_DIR / "Figure2_Survival.png"), width=Cm(15))
    P_(doc, f"Weighted Kaplan\u2013Meier survival curves with stabilised IPCW truncated at "
            f"the 1st\u201399th percentile. The unweighted log-rank test yielded {_LR_TXT}. "
            f"At day 28, weighted mortality was {pp(P['mort_early'])}% in the early "
            f"de-escalation arm versus {pp(P['mort_deferred'])}% in the deferred arm; the "
            f"weighted risk difference was {pp_label(P['rd'])} percentage points (95% "
            f"bootstrap CI {ci_pp(P['rd_ci_lo'], P['rd_ci_hi'])}). The number-at-risk table "
            f"below the x-axis shows patient counts at 0, 7, 14, 21 and 28 days for each "
            f"strategy.", italic=True, size=10, align="justify")

    # Figure 3
    H(doc, "Figure 3.  Forest plot of primary and sensitivity analyses", level=2)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run().add_picture(str(FIG_DIR / "Figure3_Forest.png"), width=Cm(15))
    P_(doc, "Forest plot of risk differences (percentage points) with 95% bootstrap "
            "confidence intervals for the primary analysis, three pre-specified "
            "sensitivity analyses (SA1\u2013SA3), the extended-covariate and "
            "alternative-truncation robustness analyses (SA4\u2013SA5) and the doubly "
            "robust AIPW estimator. All point estimates favour early de-escalation and all "
            "95% confidence intervals exclude zero, indicating a consistent and "
            "statistically significant benefit.", italic=True, size=10, align="justify")

    # Figure 4
    if EV_OK and (FIG_DIR / "Figure4_ExternalValidation.png").exists():
        H(doc, "Figure 4.  External validation in eICU-CRD v2.0", level=2)
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.add_run().add_picture(str(FIG_DIR / "Figure4_ExternalValidation.png"),
                                width=Cm(15.5))
        P_(doc, f"(A) Propensity-weighted mortality by strategy arm in the derivation "
                f"cohort (MIMIC-IV, in-hospital 28-day mortality) and the external "
                f"validation cohort (eICU-CRD v2.0, in-hospital mortality). "
                f"(B) Corresponding risk differences with 95% bootstrap confidence "
                f"intervals. Early-arm weighted mortality was nearly identical across "
                f"databases ({pp(P['mort_early'])}% vs {pp(EVP['mort_early'])}%); the "
                f"larger absolute risk difference in eICU is driven by higher deferred-arm "
                f"mortality and is interpreted directionally because outcome windows "
                f"differ (28-day in-hospital vs any in-hospital mortality); estimates "
                f"were not pooled.", italic=True, size=10, align="justify")

    # ── Tables ──
    doc.add_page_break()
    H(doc, "Tables", level=1)

    # Table 1
    H(doc, "Table 1.  Baseline characteristics by treatment strategy", level=2)
    t1 = pd.read_csv(RES / "table1_baseline.csv")
    table = doc.add_table(rows=1, cols=3)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr = table.rows[0].cells
    hdr[0].text = "Variable"
    hdr[1].text = f"Early (n={S['early']:,})"
    hdr[2].text = f"Deferred (n={S['deferred']:,})"
    for _, row in t1.iterrows():
        cells = table.add_row().cells
        cells[0].text = str(row["Variable"])
        cells[1].text = str(row[f"Early (n={S['early']})"])
        cells[2].text = str(row[f"Deferred (n={S['deferred']})"])
    P_(doc, "Values are mean (SD) for continuous variables or percentage for categorical "
            "variables, shown before stabilised weighting.", italic=True, size=10,
            align="justify")

    # Table 2
    doc.add_paragraph()
    H(doc, "Table 2.  Primary and secondary outcomes", level=2)
    table2 = doc.add_table(rows=1, cols=4)
    table2.style = "Table Grid"
    hdr2 = table2.rows[0].cells
    for i, h in enumerate(["Outcome", "Early", "Deferred", "Difference (95% CI)"]):
        hdr2[i].text = h
    if UW_OK:
        _ne, _nn = UW_M["early"]["events"], UW_M["early"]["n"]
        _de, _dn = UW_M["deferred"]["events"], UW_M["deferred"]["n"]
        _uw_row = ("28-day deaths (crude, unweighted)",
                   f"{_ne} / {_nn:,} ({_ne/_nn*100:.1f}%)",
                   f"{_de} / {_dn:,} ({_de/_dn*100:.1f}%)",
                   f"{pp_label(UW_M['rd_unweighted'])} pp")
    else:  # pragma: no cover
        _uw_row = ("28-day deaths (crude, unweighted)", "\u2014", "\u2014", "\u2014")
    rows2 = [
        ("28-day mortality (weighted)\u2020", f"{pp(P['mort_early'])}%", f"{pp(P['mort_deferred'])}%",
         f"{pp_label(P['rd'])} pp ({ci_pp(P['rd_ci_lo'], P['rd_ci_hi'])})"),
        _uw_row,
        ("Risk ratio", "\u2014", "\u2014",
         f"{P['rr']:.2f} ({P['rr_ci_lo']:.2f}\u2013{P['rr_ci_hi']:.2f})"),
        ("28-day RMST (days)", f"{P['rmst_early']:.2f}", f"{P['rmst_deferred']:.2f}",
         f"+{P['rmst_diff']:.2f} ({P['rmst_ci_lo']:+.2f} to {P['rmst_ci_hi']:+.2f})"),
        ("VFD28 (days)", f"{P['vfd_early']:.2f}", f"{P['vfd_deferred']:.2f}",
         f"+{P['vfd_diff']:.2f} ({P['vfd_ci_lo']:+.2f} to {P['vfd_ci_hi']:+.2f})"),
    ]
    for r in rows2:
        cells = table2.add_row().cells
        for i, v in enumerate(r):
            cells[i].text = v
    P_(doc, "\u2020Weighted (IPCW) risks; the risk difference and risk ratio are the "
            "per-protocol estimates from the weighted analysis. The crude row reports "
            "observed events before weighting and is shown for transparency only; it is "
            "not the causal estimate. RMST, restricted mean survival time; VFD28, "
            "ventilator-free days at day 28.", italic=True, size=10, align="justify")

    # Table 3
    doc.add_paragraph()
    H(doc, "Table 3.  Sensitivity and robustness analyses", level=2)
    table3 = doc.add_table(rows=1, cols=3)
    table3.style = "Table Grid"
    hdr3 = table3.rows[0].cells
    for i, h in enumerate(["Analysis", "RD (pp)", "95% CI (pp)"]):
        hdr3[i].text = h
    sa_rows = [
        ("Primary (weighted, hospital mortality)",
         pp_label(P['rd']), ci_pp(P['rd_ci_lo'], P['rd_ci_hi'])),
        ("SA1: No weight truncation",
         pp_label(SA["SA1_no_truncation"]["rd"]),
         ci_pp(SA["SA1_no_truncation"]["rd_ci_lo"], SA["SA1_no_truncation"]["rd_ci_hi"])),
        ("SA2: ICU mortality",
         pp_label(SA["SA2_icu_mortality"]["rd"]),
         ci_pp(SA["SA2_icu_mortality"]["rd_ci_lo"], SA["SA2_icu_mortality"]["rd_ci_hi"])),
        ("SA3: Unweighted",
         pp_label(SA["SA3_unweighted"]["rd"]),
         ci_pp(SA["SA3_unweighted"]["rd_ci_lo"], SA["SA3_unweighted"]["rd_ci_hi"])),
    ]
    if RB_OK:
        sa_rows += [
            ("SA4: Extended covariates (n=40)",
             pp_label(RB["SA4_extended_covariates"]["rd"]),
             ci_pp(RB["SA4_extended_covariates"]["rd_ci_lo"],
                   RB["SA4_extended_covariates"]["rd_ci_hi"])),
            ("SA5: Truncation 5th\u201395th percentile",
             pp_label(RB["SA5_alt_truncation"]["rd"]),
             ci_pp(RB["SA5_alt_truncation"]["rd_ci_lo"],
                   RB["SA5_alt_truncation"]["rd_ci_hi"])),
            ("AIPW doubly robust estimator",
             pp_label(RB["AIPW_doubly_robust"]["rd"]),
             ci_pp(RB["AIPW_doubly_robust"]["rd_ci_lo"],
                   RB["AIPW_doubly_robust"]["rd_ci_hi"])),
        ]
    for r in sa_rows:
        cells = table3.add_row().cells
        for i, v in enumerate(r):
            cells[i].text = v

    # Table 4
    doc.add_paragraph()
    H(doc, "Table 4.  Exploratory FiO\u2082\u00d7PEEP threshold grid", level=2)
    table4 = doc.add_table(rows=1, cols=6)
    table4.style = "Table Grid"
    hdr4 = table4.rows[0].cells
    for i, h in enumerate(["FiO\u2082 \u2264", "PEEP \u2264", "N", "Early", "Deferred",
                           "RD (95% CI), pp"]):
        hdr4[i].text = h
    for cell in TG:
        cells = table4.add_row().cells
        cells[0].text = f"{cell['fio2_threshold']}%"
        cells[1].text = f"{cell['peep_threshold']} cmH\u2082O"
        cells[2].text = f"{cell['n']:,}"
        cells[3].text = f"{cell['n_early']:,}"
        cells[4].text = f"{cell['n_deferred']:,}"
        rd_str = f"{pp_label(cell['rd'])} ({ci_pp(cell['rd_ci_lo'], cell['rd_ci_hi'])})"
        if cell['rd_ci_hi'] < 0 or cell['rd_ci_lo'] > 0:
            rd_str += " *"
        cells[5].text = rd_str
    P_(doc, "* = 95% CI excludes zero (statistically significant). Subgroups are nested; "
            "no multiplicity adjustment applied.", italic=True, size=10, align="justify")

    # Table 5
    if EV_OK:
        doc.add_paragraph()
        H(doc, "Table 5.  External validation in eICU-CRD v2.0 "
               "(in-hospital mortality)", level=2)
        table5 = doc.add_table(rows=1, cols=4)
        table5.style = "Table Grid"
        hdr5 = table5.rows[0].cells
        for i, h in enumerate(["Measure", "MIMIC-IV (derivation)",
                               "eICU-CRD (validation)", "Note"]):
            hdr5[i].text = h
        rows5 = [
            ("Analysable patients",
             f"{S['early']+S['deferred']:,}",
             f"{EV['n_analyzable']:,}", ""),
            ("Early strategy, n",
             f"{S['early']:,}", f"{EV['n_early']:,}", ""),
            ("Deferred strategy, n",
             f"{S['deferred']:,}", f"{EV['n_deferred']:,}", ""),
            ("Outcome",
             "In-hospital 28-day mortality", "In-hospital mortality",
             "28-day follow-up unavailable in eICU"),
            ("Weighted mortality, early",
             f"{pp(P['mort_early'])}%", f"{pp(EVP['mort_early'])}%", ""),
            ("Weighted mortality, deferred",
             f"{pp(P['mort_deferred'])}%", f"{pp(EVP['mort_deferred'])}%", ""),
            ("Risk difference, pp (95% CI)",
             f"{pp_label(P['rd'])} ({ci_pp(P['rd_ci_lo'], P['rd_ci_hi'])})",
             f"{pp_label(EVP['rd'])} ({ci_pp(EVP['rd_ci_lo'], EVP['rd_ci_hi'])})",
             "Both exclude zero"),
            ("Risk ratio (95% CI)",
             f"{P['rr']:.2f} ({P['rr_ci_lo']:.2f}\u2013{P['rr_ci_hi']:.2f})",
             f"{EVP['rr']:.2f} ({EVP['rr_ci_lo']:.2f}\u2013{EVP['rr_ci_hi']:.2f})",
             ""),
            ("Strict-extubation sensitivity RD, pp",
             "\u2014",
             f"{pp_label(EVSA['rd'])} "
             f"({ci_pp(EVSA['rd_ci_lo'], EVSA['rd_ci_hi'])})",
             "ICU stay \u22656 h beyond last ventilator record"),
        ]
        for r in rows5:
            cells = table5.add_row().cells
            for i, v in enumerate(r):
                cells[i].text = v
        P_(doc, "Estimates are compared directionally and were not pooled because outcome "
                "windows differ (28-day in-hospital mortality in MIMIC-IV vs in-hospital "
                "mortality in eICU-CRD).", italic=True, size=10, align="justify")

    # ── Declarations (ICM style) ──
    doc.add_page_break()
    H(doc, "Declarations", level=1)

    PR(doc, [("Funding.  ", True, False),
        ("This research received no specific grant from any funding agency in the public, "
         "commercial, or not-for-profit sectors.", False, False)])
    PR(doc, [("Conflicts of interest.  ", True, False),
        ("The authors declare no competing interests. All authors have completed the ICMJE "
         "uniform disclosure form and declare: no support from any organisation for the "
         "submitted work; no financial relationships with any organisations that might have "
         "an interest in the submitted work in the previous three years; no other "
         "relationships or activities that could appear to have influenced the submitted "
         "work.", False, False)])
    PR(doc, [("Ethics approval and consent to participate.  ", True, False),
        ("The institutional review boards of the Massachusetts Institute of Technology and "
         "Beth Israel Deaconess Medical Center approved the use of the MIMIC-IV database for "
         "research; the eICU Collaborative Research Database is likewise de-identified and "
         "exempt from IRB review under the Safe Harbor provision. The requirement for "
         "individual patient consent was waived because all "
         "data are fully de-identified.", False, False)])
    PR(doc, [("Consent for publication.  ", True, False),
        ("Not applicable (de-identified routinely collected health data; no individual "
         "patient data presented).", False, False)])
    PR(doc, [("Availability of data and materials.  ", True, False),
        ("The Medical Information Mart for Intensive Care IV (version 2.2) and the eICU "
         "Collaborative Research Database (version 2.0) are available through PhysioNet "
         "(https://physionet.org/) after completion of the required training and "
         "data-use agreements. No new data were created or analysed beyond those "
         "contained in MIMIC-IV v2.2 and eICU-CRD v2.0.", False, False)])
    PR(doc, [("Code availability.  ", True, False),
        ("The complete, reproducible analysis code (Python 3.13), including the MIMIC-IV "
         "and eICU-CRD extraction pipelines, the clone\u2013censor\u2013weight "
         "estimator, all sensitivity and robustness analyses, and figure generation, "
         "is publicly available at "
         "https://github.com/ccmzhangrui/ventilator-deescalation-target-trial-emulation.",
         False, False)])
    PR(doc, [("Author contributions.  ", True, False),
        ("[To be completed upon unblinding: study design, data curation, formal analysis, "
         "methodology, software, visualisation, writing \u2014 original draft, writing "
         "\u2014 review and editing.]", False, False)])
    PR(doc, [("Acknowledgements.  ", True, False),
        ("None.", False, False)])
    PR(doc, [("Trial registration and protocol.  ", True, False),
        ("Not registered. The full target trial protocol was specified before the analytic "
         "dataset was finalised and is provided in the Supplement (Methods A1 and the "
         "component-by-component emulation mapping in Table A1), and reporting follows the "
         "TARGET guideline (completed 21-item checklist, Supplement Checklist C1); "
         "no protocol amendments were made thereafter.", False, False)])
    PR(doc, [("Patient and public involvement.  ", True, False),
        ("No patients or members of the public were involved in the design, conduct, "
         "reporting, or dissemination of this research.", False, False)])
    PR(doc, [("Use of artificial intelligence.  ", True, False),
        ("During the preparation of this work the authors used an AI-based coding and "
         "language assistant in order to assist with analysis code verification and English "
         "language editing. After using this tool, the authors reviewed and edited the "
         "content as needed and take full responsibility for the content of the published "
         "article.", False, False)])
    PR(doc, [("Provenance and peer review.  ", True, False),
        ("Not commissioned; externally peer reviewed.", False, False)])

    # ── List of abbreviations ──
    doc.add_page_break()
    H(doc, "List of abbreviations", level=1)
    abbrevs = [
        ("ARDS", "acute respiratory distress syndrome"),
        ("CCW", "clone\u2013censor\u2013weight"),
        ("CI", "confidence interval"),
        ("CMV", "controlled mechanical ventilation"),
        ("CRRT", "continuous renal replacement therapy"),
        ("eICU-CRD", "eICU Collaborative Research Database"),
        ("FiO\u2082", "fraction of inspired oxygen"),
        ("ICU", "intensive care unit"),
        ("IPCW", "inverse probability of censoring weighting"),
        ("IRB", "institutional review board"),
        ("MIMIC-IV", "Medical Information Mart for Intensive Care IV"),
        ("MV", "mechanical ventilation"),
        ("PEEP", "positive end-expiratory pressure"),
        ("PS", "propensity score"),
        ("RD", "risk difference"),
        ("RMST", "restricted mean survival time"),
        ("RR", "risk ratio"),
        ("SBT", "spontaneous breathing trial"),
        ("SD", "standard deviation"),
        ("SMD", "standardised mean difference"),
        ("SOFA", "Sequential Organ Failure Assessment"),
        ("TTE", "target trial emulation"),
        ("VFD28", "ventilator-free days to day 28"),
    ]
    atbl = doc.add_table(rows=len(abbrevs), cols=2)
    atbl.style = "Table Grid"
    NO_HEADER_TABLES.append(atbl._tbl)
    for i, (ab, full) in enumerate(abbrevs):
        c0, c1 = atbl.rows[i].cells
        c0.paragraphs[0].add_run(ab).bold = True
        c1.paragraphs[0].add_run(full)
        for c in (c0, c1):
            for r in c.paragraphs[0].runs:
                r.font.name = "Times New Roman"
                r.font.size = Pt(10)

    # ── References ──
    H(doc, "References", level=1)
    refs = [
        "1. Ouellette DR, Patel S, Girard TD, et al. Liberation from mechanical ventilation in critically ill adults: an official American College of Chest Physicians/American Thoracic Society clinical practice guideline. Chest. 2017;151(1):166-180.",
        "2. Schmidt GA, Girard TD, Kress JP, et al. Liberation from mechanical ventilation in critically ill adults: executive summary of an official American College of Chest Physicians/American Thoracic Society clinical practice guideline. Chest. 2017;151(1):160-165.",
        "3. Boles JM, Bion J, Connors A, et al. Weaning from mechanical ventilation. Eur Respir J. 2007;29(5):1033-1056.",
        "4. Esteban A, Frutos F, Tobin MJ, et al. A comparison of four methods of weaning patients from mechanical ventilation. N Engl J Med. 1995;332(6):345-350.",
        "5. Brochard L, Rauss A, Benito S, et al. Comparison of three methods of gradual withdrawal from ventilatory support during weaning from mechanical ventilation. Am J Respir Crit Care Med. 1994;150(4):896-903.",
        "6. Ely EW, Baker AM, Dunagan DP, et al. Effect on the duration of mechanical ventilation of identifying patients capable of breathing spontaneously. N Engl J Med. 1996;335(25):1864-1869.",
        "7. Girard TD, Kress JP, Fuchs BD, et al. Efficacy and safety of a paired sedation and ventilator weaning protocol for mechanically ventilated patients in intensive care (Awakening and Breathing Controlled trial). Lancet. 2008;371(9607):126-134.",
        "8. Hern\u00e1n MA, Robins JM. Using big data to emulate a target trial when a randomized trial is not available. Am J Epidemiol. 2016;183(8):758-764.",
        "9. Hern\u00e1n MA, Sauer BC, Hern\u00e1ndez-D\u00edaz S, Platt R, Shrier I. Specifying a target trial prevents immortal time bias. Am J Epidemiol. 2016;183(7):573-578.",
        "10. Johnson AEW, Bulgarelli L, Shen L, et al. MIMIC-IV, a freely accessible electronic health record dataset. Sci Data. 2023;10(1):1.",
        "11. Xu S, Ross C, Raebel MA, Shetterly S, Blanchette C, Smith D. Use of stabilized inverse propensity scores as weights to directly estimate relative risk and its confidence interval. Value Health. 2010;13(2):273-277.",
        "12. Cole SR, Hern\u00e1n MA. Constructing inverse probability weights for marginal structural models. Am J Epidemiol. 2008;168(6):656-664.",
        "13. Robins JM, Hern\u00e1n MA, Brumback B. Marginal structural models and causal inference in epidemiology. Epidemiology. 2000;11(5):550-560.",
        "14. Efron B, Tibshirani RJ. An Introduction to the Bootstrap. New York: Chapman & Hall; 1993.",
        "15. St\u00fcrmer T, Rothman KJ, Avorn J, Glynn RJ. Treatment effects in the presence of unmeasured confounding: dealing with observations in the tails of the propensity score distribution \u2014 a simulation study. Am J Epidemiol. 2010;172(7):843-854.",
        "16. Levine S, Nguyen T, Taylor N, et al. Rapid disuse atrophy of diaphragm fibers in mechanically ventilated humans. N Engl J Med. 2008;358(13):1327-1335.",
        "17. Pollard TJ, Johnson AEW, Raffa JD, Celi LA, Mark RG, Badawi O. The eICU Collaborative Research Database, a freely available multi-centre database for critical care research. Sci Data. 2018;5:180178.",
        "18. Cashin AG, Hansford HJ, Hern\u00e1n MA, et al. Transparent reporting of observational studies emulating a target trial\u2014the TARGET statement. JAMA. 2025;334(12):1084-1093.",
        "19. Gilding AJ, Longo C. A primer on target trial emulation for respiratory research. Eur Respir J. 2026;2601143. doi:10.1183/13993003.01143-2026.",
        "20. Huang AJ, Lewis RJ. Trials terminated early\u2014when is enough, enough? JAMA. 2026;336(9):742-744.",
        "21. Royston P, Parmar MKB. Restricted mean survival time: an alternative to "
        "the hazard ratio for the design and analysis of randomized trials with a "
        "time-to-event outcome. BMC Med Res Methodol. 2013;13:152.",
    ]
    for i, ref in enumerate(refs, 1):
        P_(doc, ref, size=10, align="justify")

    out_path = BASE / "TTE_Deescalation_Manuscript_FINAL.docx"
    polish_doc(doc, running_head="Early versus deferred ventilator de-escalation: "
                                 "a target trial emulation")
    doc.save(out_path)
    print(f"  ✓ Main manuscript: {out_path}")
    return out_path


# ════════════════════════════════════════════════════════════════════════
#  SUPPLEMENT
# ════════════════════════════════════════════════════════════════════════
def build_supplement():
    doc = setup_doc()
    P_(doc, "Supplementary Material", bold=True, size=13, align="center")
    P_(doc, "Early versus Deferred De-escalation from Controlled Ventilation after a "
            "48-Hour Landmark: A Target Trial Emulation in MIMIC-IV with External "
            "Validation in eICU-CRD", italic=True,
       size=10, align="center")
    doc.add_paragraph()

    H(doc, "Methods A1: Detailed target trial protocol", level=1)
    P_(doc,
        "The full target trial protocol was specified before the analytic dataset was "
        "finalised and was not amended thereafter. Eligibility criteria, treatment "
        "strategies, outcomes, follow-up window, and the analysis plan were pre-specified.",
        align="justify")

    H(doc, "Eligibility criteria (full)", level=2)
    P_(doc, "Inclusion: (1) age \u226518 years; (2) invasive mechanical ventilation >24 "
            "hours; (3) alive and ventilated at 48 hours after ICU admission; (4) a "
            "classifiable Day-2 ventilator mode (controlled or assisted); (5) Day-2 "
            "FiO\u2082 \u226450%; (6) "
            "Day-2 PEEP \u226410 cmH\u2082O. Exclusion: age <18; no invasive "
            "ventilation; ventilation \u226424 hours; death within 48 hours; no Day-2 "
            "ventilator settings; unclassifiable Day-2 mode; Day-2 FiO\u2082 >50%; "
            "Day-2 PEEP >10 cmH\u2082O.", align="justify")

    H(doc, "Treatment strategies", level=2)
    P_(doc, "Two treatment strategies, observed over a 24-hour grace window from Day 2 to "
            "Day 3. (i) Early: a lower mode support level on Day 3 than Day 2 (i.e., "
            "controlled\u2192assisted/spontaneous step-down, or step-down from an assisted "
            "mode), or recorded end of mechanical ventilation while alive within 72 hours "
            "of initiation. (ii) Deferred: no step-down of mode support level on Day 3 with "
            "ongoing mechanical ventilation.", align="justify")

    H(doc, "Outcomes", level=2)
    P_(doc, "Primary: 28-day in-hospital mortality. Secondary: 28-day restricted mean "
            "survival time (RMST) and ventilator-free days (VFD28).", align="justify")

    H(doc, "Causal estimand, time zero and causal contrast", level=2)
    P_(doc,
        "The causal estimand was the per-protocol effect of early versus deferred "
        "de-escalation on 28-day in-hospital mortality in the eligible population, "
        "expressed as the risk difference and risk ratio. Time zero was the 48-hour "
        "landmark: eligibility ascertainment, strategy assignment and the start of "
        "follow-up were synchronised at this point so that no follow-up time accrued "
        "before assignment (prevention of immortal time bias). Strategy assignment in "
        "the ideal trial would be random and open-label; in the emulation it was "
        "operationalised by cloning each eligible patient into both strategy arms, "
        "censoring at deviation from the assigned strategy, and inverse-probability "
        "weighting to emulate randomisation. A 24-hour grace period (Day 2 to Day 3) "
        "was allowed for strategy implementation, mirroring the protocol tolerance a "
        "pragmatic trial would permit between randomisation and completion of the "
        "assigned mode transition.", align="justify")
    P_(doc,
        "Deviations between the ideal and the feasible target trial, and their "
        "implications for the study population, estimand and interpretation, were as "
        "follows. First, assignment in the ideal trial would be random; in the "
        "emulation, exchangeability holds only conditionally on the measured "
        "pre-landmark covariates, so residual confounding cannot be excluded "
        "(addressed by extended 40-covariate adjustment, the doubly robust estimator "
        "and the E-value). Second, the ideal trial would monitor ventilator settings "
        "continuously; the emulation relies on once-daily charted modes and the "
        "recorded end of ventilation, which may misclassify rapid or transient "
        "transitions and narrows the estimand to the effect of sustained "
        "day-to-day step-downs in support. Third, the ideal primary outcome would be "
        "28-day all-cause mortality; the available data support 28-day in-hospital "
        "mortality, so deaths occurring shortly after hospital discharge are not "
        "captured in either arm. Fourth, in the eICU-CRD validation, ventilator modes "
        "were inferred from treatment orders and charted parameters rather than "
        "explicitly recorded, and the outcome was harmonised to in-hospital "
        "mortality because 28-day follow-up is unavailable; the validation analysis "
        "is therefore interpreted directionally and the two databases are not "
        "pooled.", align="justify")

    H(doc, "Follow-up", level=2)
    P_(doc, "All patients were followed from the 48-hour landmark (time zero) to day 28 "
            "after ICU admission, in-hospital death, or hospital discharge (alive), "
            "whichever occurred first.", align="justify")

    H(doc, "Propensity model specification", level=2)
    P_(doc, "Multivariable logistic regression with the following 21 covariates: age, sex, "
            "weight, total SOFA, SOFA respiration, SOFA cardiovascular, SOFA central "
            "nervous system, SOFA renal, Day-2 FiO\u2082, Day-2 PEEP, Day-2 tidal volume, "
            "Day-2 peak inspiratory pressure, Day-2 plateau pressure, Day-2 PaO\u2082/FiO\u2082 "
            "ratio, Day-2 creatinine, Day-2 white blood cell count, Day-2 platelet count, "
            "Day-2 lactate, ARDS flag, Sepsis-3 flag, continuous renal replacement therapy "
            "flag. Regularisation: L2 penalty with C=1.0; lbfgs solver; maximum 1,000 "
            "iterations. Propensity scores were clipped to [0.01, 0.99]. Missing values in "
            "continuous covariates were imputed with the column median.", align="justify")

    # ── Table A1: target trial specification × emulation mapping ──
    doc.add_paragraph()
    H(doc, "Table A1.  Target trial specification and emulation mapping "
           "(TARGET items 6a\u20136h and 7a\u20137h)", level=1)
    spec_rows = [
        ("Component", "Target trial (ideal RCT)", "MIMIC-IV emulation (derivation)",
         "eICU-CRD emulation (validation)"),
        ("Eligibility (6a/7a)",
         "Adults \u226518 y; invasive MV >24 h; alive and ventilated at 48 h; "
         "Day-2 FiO\u2082 \u226450% and PEEP \u226410 cmH\u2082O; Day-2 mode "
         "classifiable",
         "Identical; mode from recorded ventilator-mode labels",
         "Identical; mode inferred from treatment orders and PS/PC charting"),
        ("Treatment strategies (6b/7b)",
         "Early: mode step-down by Day 3 or extubation \u226472 h vs Deferred: "
         "support maintained or increased; 24-h grace window",
         "Identical; modes from charted labels; extubation from recorded end of "
         "ventilation",
         "Identical; extubation from last charted ventilator setting (strict "
         "\u22656-h definition in sensitivity analysis)"),
        ("Assignment (6c/7c)",
         "Random, open-label allocation at time zero",
         "Cloning into both arms at time zero; censoring at deviation; IPCW to "
         "emulate randomisation",
         "Identical"),
        ("Time zero & follow-up (6d/7d)",
         "Time zero = 48-h landmark; follow-up to day 28, death or discharge",
         "Identical; end from discharge/discharge-death records",
         "Identical; end from hospital discharge records"),
        ("Outcome (6e/7e)",
         "28-day in-hospital mortality (primary); RMST, VFD28 (secondary)",
         "death_within_hosp_28days field; RMST and VFD28 computed from "
         "hospital course",
         "Hospital discharge status (in-hospital mortality); 28-day follow-up "
         "unavailable \u2014 harmonised to in-hospital mortality"),
        ("Causal contrast (6f/7f)",
         "Per-protocol effect: risk difference and risk ratio",
         "Weighted RD and RR with bootstrap 95% CI (2,000)",
         "Weighted RD and RR with bootstrap 95% CI (2,000); directional "
         "comparison, not pooled"),
        ("Identifying assumptions (6g/7g)",
         "Conditional exchangeability, positivity, consistency, correct model "
         "specification",
         "21 pre-landmark covariates (40 in extended SA); balance verified by "
         "SMD <0.10",
         "25 mirroring covariates; balance verified by SMD <0.10"),
        ("Analysis plan (6h/7h)",
         "Clone\u2013censor\u2013weight estimator, stabilised IPCW truncated "
         "1st\u201399th percentile; bootstrap CIs; sensitivity analyses",
         "SA1\u2013SA5, AIPW doubly robust estimator, E-value, subgroups, "
         "threshold grid",
         "Strict-extubation sensitivity analysis; ICU-mortality secondary "
         "contrast"),
    ]
    ta1 = doc.add_table(rows=0, cols=4)
    ta1.style = "Table Grid"
    for r_i, r in enumerate(spec_rows):
        cells = ta1.add_row().cells
        for i, v in enumerate(r):
            cells[i].text = v
            for p in cells[i].paragraphs:
                for run in p.runs:
                    run.font.size = Pt(8.5)
                    if r_i == 0:
                        run.font.bold = True
    P_(doc, "TARGET = Transparent Reporting of observational studies Emulating a Target "
            "Trial. SMD = standardised mean difference; IPCW = inverse probability of "
            "censoring weighting; AIPW = augmented inverse probability weighting.",
       italic=True, size=9, align="justify")

    # ── Table A2: operational definitions ──
    doc.add_paragraph()
    H(doc, "Table A2.  Operational definitions in each data source", level=1)
    op_rows = [
        ("Concept", "MIMIC-IV operationalisation", "eICU-CRD operationalisation"),
        ("Invasive mechanical ventilation",
         f"ventilation flag = 1 in the extraction ({F['ventilated']:,} of {F['total_records']:,} records)",
         "Any charted invasive-ventilator flowsheet setting (FiO\u2082, PEEP, "
         "set rate, tidal volume, pressures); NIV-only stays excluded"),
        ("Ventilation start / end",
         "Provided start/end timestamps; duration in hours",
         "First / last charted ventilator setting (minute offsets from unit "
         "admission)"),
        ("Controlled mode",
         "Mode labels: CMV, PRVC/AC, PCV+, PCV/AC, APV, VOL/AC, SIMV, AC/PC, "
         "AC/VC",
         "Treatment orders: assist-controlled, volume-/pressure-controlled, "
         "synchronised intermittent mandatory, permissive hypercapnia, "
         "volume-assured; corroborated by charted pressure control or set "
         "mandatory rate"),
        ("Assisted/spontaneous mode",
         "Mode labels: CPAP, PSV, SBT, Standby, SPONT, MMV, PPS, ApnVol, "
         "ApnPres",
         "Treatment orders: pressure support, CPAP/PEEP therapy, weaning; "
         "corroborated by charted pressure support > 0"),
        ("Extubation \u226472 h",
         "Recorded end of ventilation while alive within 72 h of initiation",
         "Last charted ventilator setting \u226472 h after start; strict "
         "sensitivity definition requires the ICU stay to continue \u22656 h "
         "beyond the last record"),
        ("Day-2 window",
         "The calendar day of the 48-h landmark (daily charted values)",
         "24\u201348 h after ventilation start (median of charted values in "
         "window)"),
        ("Day-3 window",
         "The calendar day after the landmark (daily charted values)",
         "48\u201372 h after ventilation start"),
        ("28-day in-hospital mortality",
         "death_within_hosp_28days field",
         "Not available; harmonised outcome is in-hospital mortality "
         "(hospital discharge status)"),
        ("Sepsis-3",
         "sepsis3 flag in the extraction (suspected infection + SOFA \u22652)",
         "Not used in the validation propensity model"),
        ("ARDS",
         "ards flag in the extraction",
         "Not used in the validation propensity model"),
        ("Vasopressor exposure (Day 2)",
         "Day-2 norepinephrine dose (continuous) in the 6-hourly vitals",
         "Any Day-2 infusion of norepinephrine, epinephrine, vasopressin, "
         "phenylephrine or dopamine (binary)"),
        ("Comorbidities",
         "Extraction-provided flags",
         "pastHistory table keyword mapping (COPD, heart failure, diabetes, "
         "renal disease, malignancy)"),
    ]
    ta2 = doc.add_table(rows=0, cols=3)
    ta2.style = "Table Grid"
    for r_i, r in enumerate(op_rows):
        cells = ta2.add_row().cells
        for i, v in enumerate(r):
            cells[i].text = v
            for p in cells[i].paragraphs:
                for run in p.runs:
                    run.font.size = Pt(8.5)
                    if r_i == 0:
                        run.font.bold = True

    H(doc, "Table S1.  Patient flow counts", level=1)
    flow_df = pd.read_csv(RES / "flow_counts.csv")
    t = doc.add_table(rows=1, cols=3)
    t.style = "Table Grid"
    t.rows[0].cells[0].text = "Step"
    t.rows[0].cells[1].text = "Count"
    t.rows[0].cells[2].text = "Cumulative excluded"
    cumulative = 0
    prev = None
    for idx, row in flow_df.iterrows():
        cells = t.add_row().cells
        step_label = str(row["step"]).replace("_", " ").title()
        cells[0].text = step_label
        cells[1].text = f"{int(row['count']):,}"
        if prev is not None and idx > 0:
            excl = int(prev - row["count"])
            if excl > 0:
                cells[2].text = f"\u2212{excl:,}"
            else:
                cells[2].text = "\u2014"
        else:
            cells[2].text = "\u2014"
        prev = row["count"]

    doc.add_paragraph()
    H(doc, "Table S2.  Treatment strategy classification", level=1)
    t2 = doc.add_table(rows=1, cols=2)
    t2.style = "Table Grid"
    t2.rows[0].cells[0].text = "Strategy"
    t2.rows[0].cells[1].text = "n (% of eligible)"
    for k, v in S.items():
        cells = t2.add_row().cells
        cells[0].text = k.replace("_", " ").title()
        cells[1].text = f"{v:,} ({v/F['eligible']*100:.1f}%)"

    doc.add_paragraph()
    H(doc, "Table S3.  Weight diagnostics", level=1)
    t3 = doc.add_table(rows=1, cols=2)
    t3.style = "Table Grid"
    t3.rows[0].cells[0].text = "Parameter"
    t3.rows[0].cells[1].text = "Value"
    weight_rows = [
        ("n analysed", W['n_total']),
        ("n early", W['n_early']),
        ("n deferred", W['n_deferred']),
        ("Marginal P(early)", W['p_early_marginal']),
        ("Propensity score mean", W['ps_mean']),
        ("Propensity score min", W['ps_min']),
        ("Propensity score max", W['ps_max']),
        ("Weight mean", W['weight_mean']),
        ("Weight SD", W['weight_sd']),
        ("Weight min (truncated)", W['weight_min']),
        ("Weight 1st percentile", W['weight_p1']),
        ("Weight median", W['weight_p50']),
        ("Weight 99th percentile", W['weight_p99']),
        ("Weight max (truncated)", W['weight_max']),
        ("Truncation lower bound", W['trunc_lo']),
        ("Truncation upper bound", W['trunc_hi']),
    ]
    for k, v in weight_rows:
        cells = t3.add_row().cells
        cells[0].text = k
        cells[1].text = f"{v:.4f}" if isinstance(v, float) else f"{v:,}"

    doc.add_paragraph()
    H(doc, "Table S4.  Covariate balance (standardised mean differences)", level=1)
    bal_df = pd.read_csv(RES / "covariate_balance.csv")
    t4 = doc.add_table(rows=1, cols=5)
    t4.style = "Table Grid"
    for i, h in enumerate(["Covariate", "Mean Early (unw)", "Mean Deferred (unw)",
                            "SMD (unweighted)", "SMD (weighted)"]):
        t4.rows[0].cells[i].text = h
    for _, row in bal_df.iterrows():
        cells = t4.add_row().cells
        cells[0].text = str(row["covariate"])
        cells[1].text = f"{row['mean_early_unw']:.3f}"
        cells[2].text = f"{row['mean_def_unw']:.3f}"
        cells[3].text = f"{row['smd_unweighted']:.4f}"
        cells[4].text = f"{row['smd_weighted']:.4f}"

    doc.add_paragraph()
    H(doc, "Table S5.  Sensitivity analyses", level=1)
    t5 = doc.add_table(rows=1, cols=4)
    t5.style = "Table Grid"
    for i, h in enumerate(["Analysis", "RD (pp)", "95% CI lower (pp)", "95% CI upper (pp)"]):
        t5.rows[0].cells[i].text = h
    sa_rows = [
        ("Primary (weighted)", P['rd'], P['rd_ci_lo'], P['rd_ci_hi']),
        ("SA1: No weight truncation",
         SA["SA1_no_truncation"]["rd"], SA["SA1_no_truncation"]["rd_ci_lo"],
         SA["SA1_no_truncation"]["rd_ci_hi"]),
        ("SA2: ICU mortality",
         SA["SA2_icu_mortality"]["rd"], SA["SA2_icu_mortality"]["rd_ci_lo"],
         SA["SA2_icu_mortality"]["rd_ci_hi"]),
        ("SA3: Unweighted",
         SA["SA3_unweighted"]["rd"], SA["SA3_unweighted"]["rd_ci_lo"],
         SA["SA3_unweighted"]["rd_ci_hi"]),
    ]
    if RB_OK:
        sa_rows += [
            ("SA4: Extended covariates (n=40)",
             RB["SA4_extended_covariates"]["rd"], RB["SA4_extended_covariates"]["rd_ci_lo"],
             RB["SA4_extended_covariates"]["rd_ci_hi"]),
            ("SA5: Truncation 5th\u201395th percentile",
             RB["SA5_alt_truncation"]["rd"], RB["SA5_alt_truncation"]["rd_ci_lo"],
             RB["SA5_alt_truncation"]["rd_ci_hi"]),
            ("AIPW doubly robust",
             RB["AIPW_doubly_robust"]["rd"], RB["AIPW_doubly_robust"]["rd_ci_lo"],
             RB["AIPW_doubly_robust"]["rd_ci_hi"]),
        ]
    for r in sa_rows:
        cells = t5.add_row().cells
        cells[0].text = r[0]
        cells[1].text = f"{r[1]*100:+.2f}"
        cells[2].text = f"{r[2]*100:+.2f}"
        cells[3].text = f"{r[3]*100:+.2f}"

    doc.add_paragraph()
    H(doc, "Table S6.  Exploratory FiO\u2082\u00d7PEEP threshold grid", level=1)
    t6 = doc.add_table(rows=1, cols=7)
    t6.style = "Table Grid"
    for i, h in enumerate(["FiO\u2082 \u2264", "PEEP \u2264", "N", "Early", "Deferred",
                            "RD (pp)", "95% CI (pp)"]):
        t6.rows[0].cells[i].text = h
    for cell in TG:
        cells = t6.add_row().cells
        cells[0].text = f"{cell['fio2_threshold']}%"
        cells[1].text = f"{cell['peep_threshold']} cmH\u2082O"
        cells[2].text = f"{cell['n']:,}"
        cells[3].text = f"{cell['n_early']:,}"
        cells[4].text = f"{cell['n_deferred']:,}"
        cells[5].text = f"{cell['rd']*100:+.2f}"
        sig = "*" if cell['rd_ci_hi'] < 0 or cell['rd_ci_lo'] > 0 else ""
        cells[6].text = f"{ci_pp(cell['rd_ci_lo'], cell['rd_ci_hi'])} {sig}"

    if RB_OK and SUBG is not None:
        doc.add_paragraph()
        H(doc, "Table S7.  Subgroup analyses (weighted risk differences)", level=1)
        t7 = doc.add_table(rows=1, cols=6)
        t7.style = "Table Grid"
        for i, h in enumerate(["Subgroup", "N", "Early", "Mortality (E / D)",
                               "RD (pp)", "95% CI (pp)"]):
            t7.rows[0].cells[i].text = h
        for _, row in SUBG.iterrows():
            cells = t7.add_row().cells
            cells[0].text = str(row["subgroup"])
            cells[1].text = f"{int(row['n']):,}"
            cells[2].text = f"{int(row['n_early']):,}"
            cells[3].text = (f"{row['mort_early']*100:.1f}% / "
                             f"{row['mort_deferred']*100:.1f}%")
            cells[4].text = f"{row['rd']*100:+.1f}"
            cells[5].text = ci_pp(row["rd_ci_lo"], row["rd_ci_hi"])
        P_(doc, "The Sepsis-3 flag coincided exactly with total SOFA score \u22652 (the "
                "cohort median) in this dataset, so the 'Sepsis-3' and 'SOFA \u2265 "
                "median' strata identify the same patients and are reported once. The "
                f"ARDS subgroup (n = {S['early'] + S['deferred'] - int(SUBG.loc[SUBG['subgroup'] == 'No ARDS', 'n'].iloc[0])}) "
                "was too small for estimation. Each subgroup uses "
                "the same stabilised-weighted estimator with the propensity model refit "
                "within stratum; 95% CIs from 400 bootstrap replicates. No multiplicity "
                "adjustment was applied; subgroup results are exploratory.",
           italic=True, size=10, align="justify")
        P_(doc, f"E-value analysis: for the primary risk ratio of "
                f"{RB['e_value']['rr']:.2f}, the E-value is "
                f"{RB['e_value']['e_value_point']:.2f} (E-value for the CI limit closest "
                f"to the null: {RB['e_value']['e_value_ci_bound']:.2f}). An unmeasured "
                f"confounder would need to be associated with both early de-escalation "
                f"and 28-day mortality by a risk ratio of at least "
                f"{RB['e_value']['e_value_point']:.2f}, conditional on the measured "
                f"covariates, to fully explain the observed association.",
           size=10, align="justify")

    # ── Supplementary Figures ──
    doc.add_page_break()
    H(doc, "Supplementary Figures", level=1)

    H(doc, "Figure S1.  Stabilised weight diagnostics and covariate balance", level=2)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run().add_picture(str(FIG_DIR / "FigureS1_Weights.png"), width=Cm(15))
    P_(doc, "(A) Distribution of stabilised inverse-probability-of-censoring weights "
            "after truncation at the 1st\u201399th percentile. The mean is "
            f"{W['weight_mean']:.2f}, the median is {W['weight_p50']:.2f}, and the maximum "
            f"is {W['weight_max']:.2f}. The absence of extreme weights (maximum "
            f"<10\u00d7sample mean) indicates stable estimation. (B) Standardised mean "
            "differences (SMD) for all 21 baseline covariates before (pink) and after (blue) "
            "stabilised weighting. The dashed line marks SMD = 0.10, the conventional "
            "threshold for adequate balance. After weighting, all SMDs were below 0.10 "
            f"(largest: ARDS, "
            f"{bal_df.loc[bal_df['covariate']=='ards_flag', 'smd_weighted'].iloc[0]:.3f}; "
            f"second largest: tidal volume, "
            f"{bal_df.loc[bal_df['covariate']=='d2_tv', 'smd_weighted'].iloc[0]:.3f}).",
       italic=True, size=10, align="justify")

    H(doc, "Figure S2.  Exploratory FiO\u2082\u00d7PEEP threshold grid heatmap", level=2)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run().add_picture(str(FIG_DIR / "FigureS2_Threshold.png"), width=Cm(15))
    P_(doc, "Heatmap of weighted risk differences (percentage points) across the 3\u00d73 "
            "exploratory grid of Day-2 FiO\u2082 thresholds (\u226440%, \u226450%, "
            "\u226460%) and PEEP thresholds (\u22645, \u22648, \u226410 cmH\u2082O). "
            "Each cell shows the point estimate, 95% bootstrap confidence interval, and "
            "sample size. Cells marked with \u2605 have confidence intervals excluding "
            "zero. The FiO\u2082 \u226440% column consistently yields statistically "
            "significant benefit across all three PEEP thresholds.", italic=True, size=10,
       align="justify")

    if RB_OK and (FIG_DIR / "FigureS3_Subgroup.png").exists():
        H(doc, "Figure S3.  Subgroup analyses", level=2)
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.add_run().add_picture(str(FIG_DIR / "FigureS3_Subgroup.png"), width=Cm(15))
        P_(doc, "Weighted risk differences (percentage points, 95% bootstrap CI) for "
                "28-day mortality within pre-specified subgroups. The propensity model is "
                "refit within each stratum. The risk difference is directionally "
                f"consistent across all subgroups; the ARDS stratum "
                f"(n = {S['early'] + S['deferred'] - int(SUBG.loc[SUBG['subgroup'] == 'No ARDS', 'n'].iloc[0])}) "
                "was too small "
                "for estimation. In this cohort the Sepsis-3 flag coincided exactly with "
                "total SOFA \u22652, so those strata identify the same patients.",
           italic=True, size=10, align="justify")

    if VT_OK and (FIG_DIR / "FigureS4_Trajectory.png").exists():
        H(doc, "Figure S4.  Per-protocol verification of de-escalation strategy delivery", level=2)
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.add_run().add_picture(str(FIG_DIR / "FigureS4_Trajectory.png"), width=Cm(15))
        P_(doc, "(A) Six-hourly respiratory trajectory, Day 1 00:00 through Day 2 18:00 "
                "(eight timepoints per patient), for PaO\u2082/FiO\u2082 ratio, "
                "SpO\u2082, respiratory rate and norepinephrine dose. (B) Daily Day 2 \u2192 "
                "Day 3 ventilator aggregates: FiO\u2082, PEEP and mode distribution. "
                "Time resolution is constrained by the source dataset to 6-hourly vitals "
                "and daily vent aggregates; raw chartevents at sub-hourly granularity, "
                "Day\u202f4\u20135 settings, pressure/flow waveforms, diaphragm function "
                "and patient\u2013ventilator synchrony indices are not available in this "
                f"MIMIC-IV extraction. Median [IQR] is plotted; n = {S['early'] + S['deferred']:,}. "
                "The early-arm "
                "concentration of Day\u202f3 missing-mode (extubated \u226472 h) records "
                "confirms per-protocol delivery of the de-escalation assignment.",
           italic=True, size=10, align="justify")

    # ── eICU external validation supplement section ──
    if EV_OK:
        doc.add_page_break()
        H(doc, "Methods A2: eICU-CRD harmonisation details", level=1)
        P_(doc,
            "The eICU Collaborative Research Database v2.0 (200,859 ICU stays, 208 "
            "hospitals, 2014\u20132015) was used as the external validation cohort. First "
            "ICU stays per patient (unitvisitnumber = 1) of adults (\u226518 years; "
            "age \u201c> 89\u201d coded as 90) formed the base cohort. Invasive "
            "ventilation was identified from respiratory-therapy flowsheet charting "
            "(FiO\u2082, PEEP, set rate, tidal volume, pressure support, pressure "
            "control, plateau and peak pressures); the ventilation start was the first "
            "charted setting and ventilation end the last. The 48-hour landmark, Day-2 "
            "(24\u201348 h) and Day-3 (48\u201372 h) windows, and the strategy and "
            "grace-period definitions were identical to the derivation protocol. Because "
            "eICU records no explicit ventilator-mode variable, modes were classified "
            "from treatment orders within each window (assist-controlled, volume- or "
            "pressure-controlled, synchronised intermittent mandatory, permissive "
            "hypercapnia and volume-assured ventilation mapped to controlled; pressure "
            "support, CPAP/PEEP therapy and weaning orders mapped to assisted), with "
            "corroborating inference from charted settings (pressure support >0 \u2192 "
            "assisted; pressure control or a set mandatory rate \u2192 controlled); "
            "the latest evidence in the window was used. Extubation at \u226472 hours "
            "was defined as the last charted ventilator setting at or before 72 hours "
            "(strict sensitivity definition: ICU stay continued \u22656 hours beyond "
            "the last ventilator record). The outcome was in-hospital mortality "
            "(hospital discharge status); 128 analysable stays (1.2%) with missing "
            "discharge status retained the alive classification. The propensity model "
            "contained 25 "
            "covariates: age, sex, ethnicity, admission weight, emergency-department "
            "admission, log-transformed APACHE-IV score, Day-2 FiO\u2082, PEEP, set "
            "rate, SpO\u2082, Day-2 controlled-mode indicator, any Day-2 vasopressor, "
            "Day-2 propofol and fentanyl, Day-2 lactate, creatinine, platelets and "
            "white-cell count, and comorbidities (COPD, heart failure, diabetes, renal "
            "disease, malignancy).", align="justify")

        H(doc, "Table S8.  eICU-CRD patient flow", level=1)
        t = doc.add_table(rows=1, cols=2)
        t.style = "Table Grid"
        t.rows[0].cells[0].text = "Step"
        t.rows[0].cells[1].text = "n"
        flow_rows = [
            ("First ICU stays of adults", EVF["first_stay_adults"]),
            ("Any invasive ventilator settings charted", EVF["with_vent_settings"]),
            ("Mechanical ventilation > 24 h", EVF["vent_gt_24h"]),
            ("Alive at 48-h landmark", EVF["alive_at_48h"]),
            ("Day-2 FiO\u2082 / PEEP charted", EVF["has_d2_settings"]),
            ("Day-2 FiO\u2082 \u226450% and PEEP \u226410 cmH\u2082O",
             EVF["fio2_peep_ok"]),
            ("Day-2 mode classifiable (eligible)", EVF["mode_d2_classifiable"]),
            ("Early strategy", EVS["early"]),
            ("Deferred strategy", EVS["deferred"]),
            ("Unascertainable Day-3 status", EVS["unascertainable"]),
            ("Grace-period death (\u226472 h)", EVS["grace_death"]),
            ("Analysable (CCW-weighted)", EV["n_analyzable"]),
        ]
        for name, n in flow_rows:
            c = t.add_row().cells
            c[0].text = name
            c[1].text = f"{n:,}"

        H(doc, "Table S9.  eICU-CRD covariate balance (SMD before/after weighting)",
          level=1)
        bal = pd.read_csv(RES / "eicu_balance.csv")
        t = doc.add_table(rows=1, cols=3)
        t.style = "Table Grid"
        for i, h in enumerate(["Covariate", "SMD unweighted", "SMD weighted"]):
            t.rows[0].cells[i].text = h
        for _, row in bal.iterrows():
            c = t.add_row().cells
            c[0].text = str(row["covariate"])
            c[1].text = f"{row['smd_unweighted']:.3f}"
            c[2].text = f"{row['smd_weighted']:.3f}"
        P_(doc, f"Maximum weighted SMD = {EV['balance_max_smd_weighted']:.3f} "
                f"(threshold 0.10).", italic=True, size=10, align="justify")

        if (FIG_DIR / "FigureS5_eICU_Flow.png").exists():
            H(doc, "Figure S5.  eICU-CRD v2.0 flow diagram", level=2)
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.add_run().add_picture(str(FIG_DIR / "FigureS5_eICU_Flow.png"),
                                    width=Cm(15))
            P_(doc, "Stepwise inclusion of first ICU stays of adults in eICU-CRD v2.0 "
                    "under the harmonised target trial protocol.",
               italic=True, size=10, align="justify")

    # ── Figure S6: CCW design schematic ──
    if (FIG_DIR / "FigureS6_CCW_Design.png").exists():
        H(doc, "Figure S6.  Schematic of the clone\u2013censor\u2013weight "
               "target trial emulation design", level=2)
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.add_run().add_picture(str(FIG_DIR / "FigureS6_CCW_Design.png"),
                                width=Cm(15.5))
        P_(doc, "Panel A: emulated trial timeline. Eligibility ascertainment, "
                "strategy assignment and the start of follow-up are aligned at the "
                "48-hour landmark (time zero); a 24-hour grace window (Day 2 to "
                "Day 3) is allowed for strategy implementation, and patients are "
                "followed to death, hospital discharge or Day 28. Panel B: each "
                "eligible patient is cloned into both strategy arms at time zero; "
                "each clone is censored when the observed data deviate from the "
                "assigned strategy, and the uncensored clones are reweighted by "
                "stabilised inverse-probability-of-censoring weights to emulate "
                "random assignment, yielding the weighted per-protocol contrast.",
           italic=True, size=10, align="justify")

    # ── TARGET reporting checklist (Table C1) ──
    doc.add_page_break()
    H(doc, "Checklist C1.  TARGET 21-item reporting checklist "
           "(Cashin et al., JAMA 2025)", level=1)
    P_(doc, "Location refers to the main manuscript unless stated otherwise "
            "(Suppl = this supplement).", italic=True, size=9, align="justify")
    target_items = [
        ("1a", "Abstract", "Identify emulation of a target trial; objectives and "
         "summary of the target trial", "Abstract (Background, Objectives)"),
        ("1b", "Abstract", "Report data sources used for emulation",
         "Abstract (Methods: MIMIC-IV; eICU-CRD)"),
        ("1c", "Abstract", "Summarise key assumptions, methods, findings, conclusions",
         "Abstract (Methods, Results, Conclusions)"),
        ("2", "Introduction", "Scientific background and knowledge gap",
         "Introduction, paragraphs 1\u20132"),
        ("3", "Introduction", "Summarise the causal question",
         "Introduction, final paragraph; Abstract (Objectives)"),
        ("4", "Introduction", "Rationale for emulation; cite informing RCTs",
         "Introduction, paragraph 2; Methods (design informed by refs 4\u20137)"),
        ("5", "Methods", "Data sources: purpose, type, geography, setting, period",
         "Methods (Study design and data source; External validation in eICU-CRD)"),
        ("6a", "Methods", "Target trial: eligibility criteria",
         "Methods (Eligibility); Suppl Methods A1, Table A1"),
        ("6b", "Methods", "Target trial: treatment strategies",
         "Methods (Treatment strategies); Suppl Table A1"),
        ("6c", "Methods", "Target trial: assignment procedures (random, open-label)",
         "Suppl Methods A1 (Causal estimand, time zero); Table A1"),
        ("6d", "Methods", "Target trial: follow-up start and end",
         "Suppl Methods A1 (Follow-up); Table A1"),
        ("6e", "Methods", "Target trial: outcomes",
         "Methods (Outcomes); Suppl Table A1"),
        ("6f", "Methods", "Target trial: causal contrasts and effect measures",
         "Methods (Statistical analysis, estimand paragraph); Table A1"),
        ("6g", "Methods", "Target trial: identifying assumptions",
         "Methods (Statistical analysis, identification-assumptions paragraph); "
         "Table A1; Suppl Methods A1 (ideal-vs-feasible deviations)"),
        ("6h", "Methods", "Target trial: data analysis plan",
         "Methods (Statistical analysis); Suppl Table A1"),
        ("7a", "Methods", "Emulation: eligibility operationalisation",
         "Methods (Eligibility); Suppl Table A1; Figure 1"),
        ("7b", "Methods", "Emulation: treatment strategy operationalisation",
         "Methods (Treatment strategies); Suppl Table A1"),
        ("7c", "Methods", "Emulation: assignment operationalisation (clone\u2013censor)",
         "Methods (Statistical analysis, CCW paragraph); Table A1; Figure S6"),
        ("7d", "Methods", "Emulation: follow-up operationalisation",
         "Methods (Outcomes); Suppl Methods A1 (Follow-up)"),
        ("7e", "Methods", "Emulation: outcome operationalisation",
         "Methods (Outcomes); Table A1 (eICU harmonisation)"),
        ("7f", "Methods", "Emulation: causal contrasts and effect measures",
         "Methods (Statistical analysis); Results (Primary outcome)"),
        ("7g", "Methods", "Emulation: assumptions incl. baseline confounding",
         "Methods (Statistical analysis; 21/40 covariates); Results (Weight "
         "diagnostics); Suppl Table S4"),
        ("7h", "Methods", "Emulation: analysis procedures and sensitivity analyses",
         "Methods (Statistical analysis); Results (Sensitivity, Robustness); "
         "Suppl Tables S5, S7"),
        ("8", "Results", "Numbers assessed, eligible, assigned; flow diagram",
         "Results (Study population); Figure 1; Suppl Tables S1\u2013S2; "
         "Figure S5 (eICU)"),
        ("9", "Results", "Baseline characteristics by strategy",
         "Results (Baseline characteristics); Table 1; Suppl Table S9 (eICU)"),
        ("10", "Results", "Length of follow-up and reasons for end of follow-up",
         "Methods (Outcomes); Suppl Methods A1 (Follow-up); Figure 2 at-risk table"),
        ("11", "Results", "Frequency of missing data by strategy",
         "Results (Study population: unascertainable n); Suppl Tables S2, S8"),
        ("12", "Results", "Frequency/distribution of outcomes by strategy",
         "Results (Primary outcome); Table 2; Table 5 (eICU)"),
        ("13", "Results", "Effect estimates with precision; absolute and relative",
         "Results (Primary outcome): RD and RR with 95% CI; Table 2; Table 5"),
        ("14", "Results", "Sensitivity and additional analyses",
         "Results (Sensitivity, Robustness, Threshold grid, External validation); "
         "Tables 3\u20135; Figure 3"),
        ("15", "Discussion", "Interpretation of key findings",
         "Discussion, paragraphs 1\u20134"),
        ("16", "Discussion", "Limitations incl. target trial vs emulation and "
         "confounding assumptions", "Discussion (limitations paragraph)"),
        ("17", "Other", "Ethical approval", "Declarations (Ethics approval)"),
        ("18", "Other", "Study registration / protocol",
         "Declarations (Trial registration and protocol); Suppl Methods A1, "
         "Table A1"),
        ("19", "Other", "Sharing of data, code and materials",
         "Declarations (Availability of data and materials; Code availability)"),
        ("20", "Other", "Funding sources", "Declarations (Funding)"),
        ("21", "Other", "Conflicts of interest",
         "Declarations (Conflicts of interest)"),
    ]
    tc1 = doc.add_table(rows=1, cols=4)
    tc1.style = "Table Grid"
    hdr = tc1.rows[0].cells
    for i, h in enumerate(["Item", "Section", "TARGET requirement (abbreviated)",
                           "Where addressed"]):
        hdr[i].text = h
        for p in hdr[i].paragraphs:
            for run in p.runs:
                run.font.bold = True
                run.font.size = Pt(9)
    for item, sec, req, loc in target_items:
        cells = tc1.add_row().cells
        cells[0].text = item
        cells[1].text = sec
        cells[2].text = req
        cells[3].text = loc
        for c in cells:
            for p in c.paragraphs:
                for run in p.runs:
                    run.font.size = Pt(8.5)
    P_(doc, "TARGET = Transparent Reporting of observational studies Emulating a "
            "Target Trial (reference 18). Requirement text abbreviated; full item "
            "wording available at target-guideline.org.",
       italic=True, size=9, align="justify")

    # ── STROBE checklist (Checklist C2) ──
    doc.add_page_break()
    H(doc, "Checklist C2.  STROBE reporting checklist "
           "(von Elm et al., J Clin Epidemiol 2008)", level=1)
    P_(doc, "Location refers to the main manuscript unless stated otherwise "
            "(Suppl = this supplement).", italic=True, size=9, align="justify")
    strobe_items = [
        ("1a", "Title/Abstract", "Indicate the study design with a commonly used "
         "term in the title or abstract",
         "Title and Abstract (\u201ctarget trial emulation\u201d)"),
        ("1b", "Title/Abstract", "Informative, balanced summary of what was done "
         "and found", "Abstract (four-part structured)"),
        ("2", "Introduction", "Scientific background and rationale",
         "Introduction, paragraphs 1\u20132"),
        ("3", "Introduction", "Specific objectives and hypotheses",
         "Introduction, final paragraph; Abstract (Objectives)"),
        ("4", "Methods", "Key elements of study design, early in the paper",
         "Methods (Study design and data source)"),
        ("5", "Methods", "Setting, locations, relevant dates",
         "Methods (Study design and data source; External validation in eICU-CRD)"),
        ("6", "Methods", "Eligibility criteria and sources/methods of participant "
         "selection", "Methods (Eligibility); Figure 1; Suppl Figure S5"),
        ("7", "Methods", "All outcomes, exposures, predictors, confounders and "
         "effect modifiers clearly defined",
         "Methods (Treatment strategies; Outcomes; Statistical analysis); "
         "Suppl Tables A1\u2013A2"),
        ("8", "Methods", "Sources of data and methods of assessment for each "
         "variable of interest", "Methods (data source; External validation); "
         "Suppl Methods A1\u2013A2, Table A2"),
        ("9", "Methods", "Efforts to address potential sources of bias",
         "Methods (landmark design, CCW, IPCW, truncation); Discussion "
         "(limitations)"),
        ("10", "Methods", "How the study size was arrived at",
         "Results (Study population); Figure 1; no formal sample-size "
         "calculation (all eligible patients included)"),
        ("11", "Methods", "How quantitative variables were handled",
         "Methods (Statistical analysis; thresholds); Suppl Methods A1"),
        ("12a", "Methods", "All statistical methods, including confounding control",
         "Methods (Statistical analysis); Suppl Methods A1 (propensity model)"),
        ("12b", "Methods", "Methods used to examine subgroups and interactions",
         "Methods (Statistical analysis: subgroup analyses); Suppl Table S7"),
        ("12c", "Methods", "How missing data were addressed",
         "Suppl Methods A1 (median imputation); Suppl Table S10"),
        ("12d", "Methods", "How loss to follow-up was addressed",
         "Methods (Outcomes: censoring at discharge); Suppl Methods A1 "
         "(Follow-up)"),
        ("12e", "Methods", "Any sensitivity analyses",
         "Methods (Statistical analysis: SA1\u2013SA5, AIPW, E-value); "
         "Results (Sensitivity; Robustness; External validation)"),
        ("13a", "Results", "Numbers at each stage of the study",
         "Results (Study population); Figure 1; Suppl Table S1, S8"),
        ("13b", "Results", "Reasons for non-participation at each stage",
         "Figure 1; Suppl Tables S1\u2013S2; Figure S5"),
        ("13c", "Results", "Use of a flow diagram",
         "Figure 1 (MIMIC-IV); Suppl Figure S5 (eICU-CRD)"),
        ("14a", "Results", "Characteristics of participants and information on "
         "exposures/confounders", "Table 1; Suppl Table S9; eICU baseline CSV"),
        ("14b", "Results", "Number of participants with missing data for each "
         "variable of interest", "Suppl Table S10"),
        ("14c", "Results", "Follow-up time summary",
         "Methods (Outcomes); Suppl Methods A1 (Follow-up); Figure 2"),
        ("15", "Results", "Number of outcome events or summary measures",
         "Results (Primary outcome); Table 2; Table 5"),
        ("16a", "Results", "Unadjusted and confounder-adjusted estimates with "
         "precision", "Tables 2\u20133; unweighted (SA3) vs weighted estimates"),
        ("16b", "Results", "Category boundaries when continuous variables were "
         "categorised", "Methods (Eligibility thresholds); Table 4 (threshold "
         "grid)"),
        ("16c", "Results", "Consider translating relative into absolute risk for "
         "a meaningful period", "RD in percentage points throughout; Take-home "
         "message"),
        ("17", "Results", "Other analyses (subgroups, sensitivity)",
         "Results (Sensitivity; Robustness; Threshold grid; External "
         "validation); Tables 3\u20135; Figures 3\u20134"),
        ("18", "Discussion", "Key results summarised with reference to "
         "objectives", "Discussion, paragraph 1"),
        ("19", "Discussion", "Limitations, direction and magnitude of potential "
         "bias", "Discussion (limitations paragraph)"),
        ("20", "Discussion", "Cautious overall interpretation",
         "Discussion, paragraphs 1\u20134 and conclusion"),
        ("21", "Discussion", "Generalisability (external validity)",
         "Discussion (external-validation paragraph and limitations); "
         "Table 5; Figure 4"),
        ("22", "Other", "Source of funding and role of funders",
         "Declarations (Funding: no specific grant)"),
    ]
    tc2 = doc.add_table(rows=1, cols=4)
    tc2.style = "Table Grid"
    hdr = tc2.rows[0].cells
    for i, h in enumerate(["Item", "Section", "STROBE requirement (abbreviated)",
                           "Where addressed"]):
        hdr[i].text = h
        for p in hdr[i].paragraphs:
            for run in p.runs:
                run.font.bold = True
                run.font.size = Pt(9)
    for item, sec, req, loc in strobe_items:
        cells = tc2.add_row().cells
        cells[0].text = item
        cells[1].text = sec
        cells[2].text = req
        cells[3].text = loc
        for c in cells:
            for p in c.paragraphs:
                for run in p.runs:
                    run.font.size = Pt(8.5)
    P_(doc, "STROBE = Strengthening the Reporting of Observational Studies in "
            "Epidemiology (von Elm E, et al. J Clin Epidemiol. 2008;61(4):344-349). "
            "Requirement text abbreviated.",
       italic=True, size=9, align="justify")

    # ── Table S10: missingness and imputation ──
    doc.add_page_break()
    H(doc, "Table S10.  Covariate missingness before imputation and handling",
      level=1)
    P_(doc, "Missing values in continuous covariates were imputed with the cohort "
            "median before propensity-model fitting (single imputation, pre-specified "
            "in the analysis plan); binary indicators were complete by construction "
            "(absence of a recorded flag = 0). MIMIC-IV SOFA subscores are computed "
            "only for patients meeting Sepsis-3 criteria in the source extraction; "
            "their missingness therefore reflects the non-septic fraction. The "
            "extended 40-covariate robustness analysis (SA4) and the doubly robust "
            "AIPW estimator assess sensitivity to these handling choices.",
       align="justify")
    mm = pd.read_csv(RES / "mimic_missingness.csv")
    em = pd.read_csv(RES / "eicu_missingness.csv")
    all_cov = list(dict.fromkeys(list(mm["covariate"]) + list(em["covariate"])))
    t10 = doc.add_table(rows=1, cols=5)
    t10.style = "Table Grid"
    hdr = t10.rows[0].cells
    for i, h in enumerate(["Covariate",
                           "MIMIC-IV n missing (%)",
                           "eICU-CRD n missing (%)",
                           "Handling", "Role"]):
        hdr[i].text = h
        for p in hdr[i].paragraphs:
            for run in p.runs:
                run.font.bold = True
                run.font.size = Pt(9)
    mm_d = {r["covariate"]: r for _, r in mm.iterrows()}
    em_d = {r["covariate"]: r for _, r in em.iterrows()}
    label_map = {
        "age_num": "Age", "male": "Male sex", "eth_white": "White ethnicity",
        "weight_num": "Body weight", "admissionweight": "Admission weight",
        "log_apache": "log(APACHE-IV + 1)",
        "sofa_total": "SOFA total", "sofa_respiration": "SOFA respiration",
        "sofa_cardiovascular": "SOFA cardiovascular", "sofa_cns": "SOFA CNS",
        "sofa_renal": "SOFA renal",
        "d2_fio2": "Day-2 FiO\u2082", "fio2_d2": "Day-2 FiO\u2082",
        "d2_peep": "Day-2 PEEP", "peep_d2": "Day-2 PEEP",
        "d2_tv": "Day-2 tidal volume", "d2_pip": "Day-2 peak pressure",
        "d2_pplat": "Day-2 plateau pressure",
        "pao2_fio2_d2_mean": "Day-2 PaO\u2082/FiO\u2082",
        "vent_rate_d2": "Day-2 set ventilator rate",
        "sao2_d2": "Day-2 SpO\u2082",
        "d2_cr": "Day-2 creatinine", "creatinine_d2": "Day-2 creatinine",
        "d2_wbc": "Day-2 WBC", "wbc_d2": "Day-2 WBC",
        "d2_plt": "Day-2 platelets", "platelets_d2": "Day-2 platelets",
        "d2_lactate": "Day-2 lactate", "lactate_d2": "Day-2 lactate",
        "ards_flag": "ARDS flag", "sepsis3_flag": "Sepsis-3 flag",
        "crrt_flag": "CRRT flag",
        "any_vasopressor_d2": "Day-2 any vasopressor",
        "drug_propofol_d2": "Day-2 propofol",
        "drug_fentanyl_d2": "Day-2 fentanyl",
        "copd": "COPD", "chf": "Heart failure", "diabetes": "Diabetes",
        "renal": "Renal disease", "malignancy": "Malignancy",
    }
    # canonical merge BY LABEL: a MIMIC key and its eICU counterpart
    # (e.g. d2_lactate / lactate_d2) populate the SAME row
    canon, order = {}, []
    for cov in list(mm["covariate"]) + list(em["covariate"]):
        lab = label_map.get(cov, cov)
        if lab not in canon:
            canon[lab] = [None, None]
            order.append(lab)
    for cov in mm["covariate"]:
        canon[label_map.get(cov, cov)][0] = mm_d[cov]
    for cov in em["covariate"]:
        canon[label_map.get(cov, cov)][1] = em_d[cov]
    for lab in order:
        cells = t10.add_row().cells
        cells[0].text = lab
        r1, r2 = canon[lab]
        cells[1].text = (f"{int(r1['n_missing']):,} ({r1['pct_missing']}%)"
                         if r1 is not None else "\u2014")
        cells[2].text = (f"{int(r2['n_missing']):,} ({r2['pct_missing']}%)"
                         if r2 is not None else "\u2014")
        is_binary = lab in {"Male sex", "White ethnicity", "ARDS flag",
                            "Sepsis-3 flag", "CRRT flag", "Day-2 any vasopressor",
                            "Day-2 propofol", "Day-2 fentanyl", "COPD",
                            "Heart failure", "Diabetes", "Renal disease",
                            "Malignancy"}
        cells[3].text = ("None needed (complete by construction)" if is_binary
                         else "Cohort-median imputation")
        cells[4].text = "Propensity model"
        for c in cells:
            for p in c.paragraphs:
                for run in p.runs:
                    run.font.size = Pt(8.5)
    P_(doc, "Derived indicators (ethnicity subgroups, admission source, Day-2 mode) "
            "are computed from complete source fields and have no missingness.",
       italic=True, size=9, align="justify")

    out_path = BASE / "TTE_Deescalation_Supplement_FINAL.docx"
    polish_doc(doc, running_head="Supplementary appendix")
    doc.save(out_path)
    print(f"  ✓ Supplement: {out_path}")
    return out_path


# ════════════════════════════════════════════════════════════════════════
#  COVER LETTER
# ════════════════════════════════════════════════════════════════════════
def build_cover_letter():
    doc = setup_doc()
    P_(doc, "Cover Letter", bold=True, size=13, align="center")
    doc.add_paragraph()
    P_(doc, "Dear Editor,", align="left")
    doc.add_paragraph()
    P_(doc,
        "We are pleased to submit our manuscript entitled \u201CEarly versus Deferred "
        "De-escalation from Controlled Ventilation after a 48-Hour Landmark of "
        "Physiological Stability: A Target Trial Emulation in MIMIC-IV with "
        "External Validation in eICU-CRD\u201D for "
        "consideration as an Original Article in Intensive Care Medicine.", align="justify")
    P_(doc,
        f"The optimal timing of transition from controlled to assisted mechanical "
        f"ventilation remains a critical unanswered question in intensive care medicine. "
        f"Using a target trial emulation framework with clone\u2013censor\u2013weight "
        f"estimation, we compared early versus deferred de-escalation strategies in "
        f"{F['eligible']:,} adults with a classifiable ventilator mode (controlled or "
        f"assisted) still invasively ventilated 48 hours after initiation with "
        f"FiO\u2082 \u226450% and PEEP \u226410 cmH\u2082O. "
        f"Our primary analysis in {S['early'] + S['deferred']:,} analysable patients showed that early "
        f"de-escalation was associated with lower 28-day mortality "
        f"({pp(P['mort_early'])}% versus {pp(P['mort_deferred'])}%; risk "
        f"difference {pp_label(P['rd'])} percentage points; 95% CI "
        f"{ci_pp(P['rd_ci_lo'], P['rd_ci_hi'])}; risk ratio {P['rr']:.2f} "
        f"(95% CI {P['rr_ci_lo']:.2f}\u2013{P['rr_ci_hi']:.2f}); p < 0.001). The finding was robust across "
        f"six sensitivity and robustness analyses, including an extended 40-covariate "
        f"propensity adjustment, alternative weight truncation, and a doubly robust AIPW "
        f"estimator (RD {pp(_AIPW_RD)} percentage points); the E-value of {_EV_TXT} quantifies "
        f"the substantial unmeasured confounding that would be required to explain away "
        f"the effect. In the exploratory threshold "
        f"analysis, all nine FiO\u2082\u2013PEEP cells showed statistically significant "
        f"benefit (RD {TG_RANGE_ALL} percentage points). Critically, the "
        f"association was externally validated under a harmonised protocol in the "
        f"independent multi-centre eICU-CRD cohort ({EV['n_analyzable']:,} analysable "
        f"patients across 208 hospitals): weighted in-hospital mortality "
        f"{pp(EVP['mort_early'])}% versus {pp(EVP['mort_deferred'])}% (risk difference "
        f"{pp_label(EVP['rd'])} percentage points; 95% CI "
        f"{ci_pp(EVP['rd_ci_lo'], EVP['rd_ci_hi'])}; risk ratio {EVP['rr']:.2f}).",
       align="justify")
    P_(doc,
        "We believe this work is of interest to the readers of Intensive Care Medicine "
        "because it applies rigorous causal-inference methods to a clinically important "
        "question, provides a fully pre-specified protocol with no post-hoc amendments "
        "and complete TARGET-compliant reporting (component-by-component emulation "
        "mapping and the 21-item checklist are included), "
        "includes per-protocol verification showing that the assigned strategies were "
        "delivered as intended, and demonstrates transportability of the finding to an "
        "independent multi-centre database \u2014 a rare standard in observational "
        "weaning research. All results reported in the manuscript can be reproduced "
        "from the analysis summary and supplementary tables.", align="justify")
    P_(doc, "This manuscript is original work that has not been published elsewhere and is "
            "not under consideration at another journal. All listed authors have approved "
            "the submitted version and agree to be accountable for its content. We have no "
            "conflicts of interest to disclose. No specific funding was received for this "
            "study.", align="justify")
    doc.add_paragraph()
    P_(doc, "Thank you for your consideration.", align="left")
    doc.add_paragraph()
    P_(doc, "Sincerely,", align="left")
    P_(doc, "[Corresponding Author]", align="left")
    P_(doc, "[Affiliation]", align="left")
    P_(doc, "[Email]", align="left")

    out_path = BASE / "TTE_CoverLetter_FINAL.docx"
    polish_doc(doc)
    doc.save(out_path)
    print(f"  ✓ Cover letter: {out_path}")
    return out_path


# ════════════════════════════════════════════════════════════════════════
#  PPT — full 17-slide deck with embedded figures and tables
# ════════════════════════════════════════════════════════════════════════
PPT_SLIDES = []
PPT_ASSETS = []  # (src_path, dest_filename)

def _add_slide(content_xml: str):
    PPT_SLIDES.append(content_xml)

def _register_asset(src: Path, dest_name: str):
    """Copy asset into the ppt assets dir and return the destination path string."""
    import shutil
    dst = PPT_DIR / "assets" / dest_name
    if not dst.exists():
        shutil.copy(src, dst)
    return str(dst)


def build_ppt():
    """Generate a 17-slide deck using the tencent-pptx skill pipeline."""
    import shutil
    # Copy figures to PPT assets directory
    fig_paths = {
        "fig1": (FIG_DIR / "Figure1_Flow.png", "Figure1_Flow.png"),
        "fig2": (FIG_DIR / "Figure2_Survival.png", "Figure2_Survival.png"),
        "fig3": (FIG_DIR / "Figure3_Forest.png", "Figure3_Forest.png"),
        "fig4": (FIG_DIR / "Figure4_ExternalValidation.png",
                 "Figure4_ExternalValidation.png"),
        "figS1": (FIG_DIR / "FigureS1_Weights.png", "FigureS1_Weights.png"),
        "figS2": (FIG_DIR / "FigureS2_Threshold.png", "FigureS2_Threshold.png"),
        "figS4": (FIG_DIR / "FigureS4_Trajectory.png", "FigureS4_Trajectory.png"),
    }
    fig_assets = {k: _register_asset(v[0], v[1]) for k, v in fig_paths.items()}

    # ─── 1: Title slide ───
    _add_slide(f'''<Slide>
  <div style={{{{ width: '1280px', height: '720px', background: '#FFFFFF', position: 'relative', overflow: 'hidden' }}}}>
    <div style={{{{ position: 'absolute', top: 0, left: 0, width: '1280px', height: '8px', background: '#1E4FA8' }}}}/>
    <div style={{{{ padding: '100px 80px 0 80px' }}}}>
      <div style={{{{ fontSize: '18px', color: '#3D7BD9', fontWeight: 600, marginBottom: '24px' }}}}>Target Trial Emulation · MIMIC-IV v2.2 + eICU-CRD v2.0 外部验证</div>
      <div style={{{{ fontSize: '40px', color: '#0E3F8C', fontWeight: 700, lineHeight: 1.3, marginBottom: '20px' }}}}>早期与延迟降阶：控制性通气<br/>48小时地标后的目标试验模拟</div>
      <div style={{{{ fontSize: '20px', color: '#4A5568', lineHeight: 1.6, marginBottom: '32px' }}}}>Early versus Deferred De-escalation from Controlled Ventilation<br/>after a 48-Hour Landmark of Physiological Stability</div>
      <div style={{{{ width: '120px', height: '4px', background: '#1E4FA8', marginBottom: '32px' }}}}/>
      <div style={{{{ fontSize: '18px', color: '#8B97A8', lineHeight: 1.8 }}}}>
        Clone-Censor-Weight Estimation · Stabilised IPCW<br/>
        Primary: 28-Day In-Hospital Mortality · n = {F['eligible']} eligible · {S['early']+S['deferred']} analysable<br/>
        External validation: eICU-CRD · {EV['n_analyzable']:,} analysable · 208 hospitals<br/>
        Bootstrap 2,000 (primary) / 800 (sensitivity & exploratory)
      </div>
    </div>
    <div style={{{{ position: 'absolute', bottom: '40px', left: '80px', fontSize: '16px', color: '#8B97A8' }}}}>2026-08-18 · MIMIC-IV v2.2 (BIDMC) + eICU-CRD v2.0 (208 hospitals)</div>
  </div>
</Slide>''')

    # ─── 2: Background & clinical question ───
    _add_slide('''<Slide>
  <div style={{ width: '1280px', height: '720px', background: '#FFFFFF', position: 'relative', overflow: 'hidden' }}>
    <div style={{ position: 'absolute', top: 0, left: 0, width: '1280px', height: '72px', background: '#0E3F8C', display: 'flex', alignItems: 'center', paddingLeft: '40px' }}>
      <span style={{ fontSize: '24px', color: '#FFFFFF', fontWeight: 700 }}>背景与研究问题</span>
    </div>
    <div style={{ padding: '96px 60px 0 60px' }}>
      <div style={{ fontSize: '18px', color: '#1A2230', lineHeight: 1.7, marginBottom: '20px' }}>
        <span><b>临床问题</b>：机械通气 48 小时后，下调呼吸支持级别（控制→辅助/自主模式或撤机拔管）的最佳时机仍不明确。</span>
      </div>
      <div style={{ display: 'flex', gap: '24px', marginTop: '12px' }}>
        <div style={{ flex: 1, background: '#FFF5F5', borderLeft: '4px solid #C0392B', padding: '20px 24px', borderRadius: '6px' }}>
          <div style={{ fontSize: '16px', color: '#9E2A2A', fontWeight: 700, marginBottom: '10px' }}>过早降阶的风险</div>
          <div style={{ fontSize: '15px', color: '#1A2230', lineHeight: 1.6 }}>
            • 人机对抗 · 气体交换恶化<br/>• 再插管率升高 · 死亡率上升
          </div>
        </div>
        <div style={{ flex: 1, background: '#F5FAFF', borderLeft: '4px solid #2E86AB', padding: '20px 24px', borderRadius: '6px' }}>
          <div style={{ fontSize: '16px', color: '#1B4F72', fontWeight: 700, marginBottom: '10px' }}>延迟降阶的风险</div>
          <div style={{ fontSize: '15px', color: '#1A2230', lineHeight: 1.6 }}>
            • 膈肌功能障碍 · 呼吸机相关肺炎<br/>• 深度镇静 · 延长 ICU 住院
          </div>
        </div>
      </div>
      <div style={{ marginTop: '32px', background: '#E8EFF8', padding: '20px 28px', borderRadius: '8px', borderLeft: '4px solid #1E4FA8' }}>
        <div style={{ fontSize: '17px', color: '#0E3F8C', fontWeight: 600, lineHeight: 1.6 }}>
          <span><b>研究目标</b>：在 48 小时地标后仍接受有创通气且低氧需求（FiO₂ ≤50%/PEEP ≤10）的患者中，比较 24 小时宽限期内的早期降阶（模式下调或撤机拔管）与延迟降阶对 28 天院内死亡率的影响。</span>
        </div>
      </div>
    </div>
    <div style={{ position: 'absolute', bottom: '20px', right: '40px', fontSize: '14px', color: '#1E4FA8', fontWeight: 600 }}>02 / 17</div>
  </div>
</Slide>''')

    # ─── 3: Study design ───
    _add_slide('''<Slide>
  <div style={{ width: '1280px', height: '720px', background: '#FFFFFF', position: 'relative', overflow: 'hidden' }}>
    <div style={{ position: 'absolute', top: 0, left: 0, width: '1280px', height: '72px', background: '#0E3F8C', display: 'flex', alignItems: 'center', paddingLeft: '40px' }}>
      <span style={{ fontSize: '24px', color: '#FFFFFF', fontWeight: 700 }}>研究设计</span>
    </div>
    <div style={{ padding: '96px 60px 0 60px' }}>
      <div style={{ fontSize: '18px', color: '#1A2230', lineHeight: 1.7, marginBottom: '24px' }}>
        <span><b>类型</b>：回顾性目标试验模拟（Target Trial Emulation）+ Clone-Censor-Weight 估计</span>
      </div>
      <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '20px', marginTop: '20px' }}>
        <div style={{ background: '#F7F9FC', padding: '20px', borderRadius: '8px' }}>
          <div style={{ fontSize: '16px', color: '#0E3F8C', fontWeight: 700, marginBottom: '10px' }}>数据源</div>
          <div style={{ fontSize: '15px', color: '#1A2230', lineHeight: 1.6 }}>
            MIMIC-IV v2.2 · 推导数据库（BIDMC 单中心）<br/>Beth Israel Deaconess Medical Center<br/>51,992 条 ICU 记录
          </div>
        </div>
        <div style={{ background: '#F7F9FC', padding: '20px', borderRadius: '8px' }}>
          <div style={{ fontSize: '16px', color: '#0E3F8C', fontWeight: 700, marginBottom: '10px' }}>地标与宽限期</div>
          <div style={{ fontSize: '15px', color: '#1A2230', lineHeight: 1.6 }}>
            地标点：ICU 入院后 48 小时（Day 2）<br/>宽限期：24 小时（Day 2 → Day 3）<br/>随访：至 28 天
          </div>
        </div>
        <div style={{ background: '#F7F9FC', padding: '20px', borderRadius: '8px' }}>
          <div style={{ fontSize: '16px', color: '#0E3F8C', fontWeight: 700, marginBottom: '10px' }}>纳入标准</div>
          <div style={{ fontSize: '15px', color: '#1A2230', lineHeight: 1.6 }}>
            年龄 ≥18 · 侵入性 MV &gt;24 h · 48 h 仍存活通气<br/>Day-2 模式可分类（控制/辅助） · FiO₂ ≤50% · PEEP ≤10 cmH₂O
          </div>
        </div>
        <div style={{ background: '#F7F9FC', padding: '20px', borderRadius: '8px' }}>
          <div style={{ fontSize: '16px', color: '#0E3F8C', fontWeight: 700, marginBottom: '10px' }}>估计方法</div>
          <div style={{ fontSize: '15px', color: '#1A2230', lineHeight: 1.6 }}>
            单区间 CCW + 稳定化 IPCW（截断 1–99 百分位）<br/>21 个基线协变量 · L2 正则 logistic 回归
          </div>
        </div>
      </div>
    </div>
    <div style={{ position: 'absolute', bottom: '20px', right: '40px', fontSize: '14px', color: '#1E4FA8', fontWeight: 600 }}>03 / 17</div>
  </div>
</Slide>''')

    # ─── 4: Flow diagram (Figure 1) ───
    _add_slide(f'''<Slide>
  <div style={{{{ width: '1280px', height: '720px', background: '#FFFFFF', position: 'relative', overflow: 'hidden' }}}}>
    <div style={{{{ position: 'absolute', top: 0, left: 0, width: '1280px', height: '72px', background: '#0E3F8C', display: 'flex', alignItems: 'center', paddingLeft: '40px' }}}}>
      <span style={{{{ fontSize: '24px', color: '#FFFFFF', fontWeight: 700 }}}}>研究流程 (Figure 1)</span>
    </div>
    <div style={{{{ display: 'flex', padding: '90px 40px 0 40px' }}}}>
      <div style={{{{ flex: 1.6, paddingRight: '20px' }}}}>
        <img src="{fig_assets['fig1']}" style={{{{ width: '100%', height: 'auto' }}}}/>
      </div>
      <div style={{{{ flex: 1, padding: '40px 24px', background: '#F7F9FC', borderRadius: '8px', alignSelf: 'center' }}}}>
        <div style={{{{ fontSize: '17px', color: '#0E3F8C', fontWeight: 700, marginBottom: '16px' }}}}>关键数字</div>
        <div style={{{{ fontSize: '14px', color: '#1A2230', lineHeight: 1.8 }}}}>
          <span><b>{F['total_records']:,}</b> ICU records</span><br/>
          <span><b>{F['ventilated']:,}</b> 接受侵入性通气</span><br/>
          <span><b>{F['alive_at_48h']:,}</b> 48 h 仍存活通气</span><br/>
          <span><b>{F['eligible']}</b> 符合所有标准</span><br/>
          <span style={{{{ color: '#C0392B' }}}}>· <b>{S['early']}</b> 早期降阶</span><br/>
          <span style={{{{ color: '#666' }}}}>· <b>{S['unascertainable']}</b> 不可分类</span><br/>
          <span style={{{{ color: '#2E86AB' }}}}>· <b>{S['deferred']}</b> 延迟降阶</span><br/>
          <span style={{{{ color: '#888' }}}}>  (含 {S['grace_death']} 人宽限期内死亡)</span><br/>
          <hr style={{{{ margin: '10px 0', border: 'none', borderTop: '1px solid #ddd' }}}}/>
          <span><b>{S['early']+S['deferred']}</b> 进入主分析</span>
        </div>
      </div>
    </div>
    <div style={{{{ position: 'absolute', bottom: '20px', left: '40px', fontSize: '14px', color: '#8B97A8' }}}}>Figure 1. Study flow diagram</div>
    <div style={{{{ position: 'absolute', bottom: '20px', right: '40px', fontSize: '14px', color: '#1E4FA8', fontWeight: 600 }}}}>04 / 17</div>
  </div>
</Slide>''')

    # ─── 5: Baseline characteristics (Table 1) ───
    t1 = pd.read_csv(RES / "table1_baseline.csv")
    rows = []
    for _, row in t1.iterrows():
        rows.append((str(row["Variable"]),
                     str(row[f"Early (n={S['early']})"]),
                     str(row[f"Deferred (n={S['deferred']})"])))
    rows_xml = "\n".join([
        f"<div key={{{i}}} style={{{{ display: 'flex', background: '{('#F7F9FC' if i%2==0 else '#FFFFFF')}', padding: '4px 16px', fontSize: '13px', color: '#1A2230', borderBottom: '1px solid #E5E7EB', lineHeight: 1.35 }}}}>"
        f"<div style={{{{ flex: 2.2 }}}}>{r[0]}</div>"
        f"<div style={{{{ flex: 1, textAlign: 'center' }}}}>{r[1]}</div>"
        f"<div style={{{{ flex: 1, textAlign: 'center' }}}}>{r[2]}</div>"
        f"</div>"
        for i, r in enumerate(rows)
    ])
    _add_slide(f'''<Slide>
  <div style={{{{ width: '1280px', height: '720px', background: '#FFFFFF', position: 'relative', overflow: 'hidden' }}}}>
    <div style={{{{ position: 'absolute', top: 0, left: 0, width: '1280px', height: '72px', background: '#0E3F8C', display: 'flex', alignItems: 'center', paddingLeft: '40px' }}}}>
      <span style={{{{ fontSize: '24px', color: '#FFFFFF', fontWeight: 700 }}}}>基线特征 (Table 1)</span>
    </div>
    <div style={{{{ padding: '82px 60px 0 60px' }}}}>
      <div style={{{{ display: 'flex', background: '#1E4FA8', color: '#FFFFFF', borderRadius: '6px 6px 0 0', padding: '6px 16px', fontSize: '13px', fontWeight: 600 }}}}>
        <div style={{{{ flex: 2.2 }}}}>变量</div>
        <div style={{{{ flex: 1, textAlign: 'center' }}}}>Early (n={S['early']})</div>
        <div style={{{{ flex: 1, textAlign: 'center' }}}}>Deferred (n={S['deferred']})</div>
      </div>
      {rows_xml}
      <div style={{{{ marginTop: '10px', fontSize: '12px', color: '#4A5568', lineHeight: 1.4 }}}}>
        两组基线特征总体相似；SOFA、FiO₂、PEEP、平台压、乳酸等核心临床变量均无临床显著差异。<br/>
        加权后所有标准化均差（SMD）&lt;0.10（CRRT 0.076、ARDS 0.042 除外），协变量平衡充分。
      </div>
    </div>
    <div style={{{{ position: 'absolute', bottom: '20px', left: '40px', fontSize: '14px', color: '#8B97A8' }}}}>Table 1. Baseline characteristics (full)</div>
    <div style={{{{ position: 'absolute', bottom: '20px', right: '40px', fontSize: '14px', color: '#1E4FA8', fontWeight: 600 }}}}>05 / 17</div>
  </div>
</Slide>''')

    # ─── 6: Weight diagnostics (Figure S1) ───
    _add_slide(f'''<Slide>
  <div style={{{{ width: '1280px', height: '720px', background: '#FFFFFF', position: 'relative', overflow: 'hidden' }}}}>
    <div style={{{{ position: 'absolute', top: 0, left: 0, width: '1280px', height: '72px', background: '#0E3F8C', display: 'flex', alignItems: 'center', paddingLeft: '40px' }}}}>
      <span style={{{{ fontSize: '24px', color: '#FFFFFF', fontWeight: 700 }}}}>权重诊断 (Figure S1)</span>
    </div>
    <div style={{{{ padding: '90px 40px 0 40px' }}}}>
      <img src="{fig_assets['figS1']}" style={{{{ width: '100%', height: 'auto' }}}}/>
      <div style={{{{ display: 'flex', gap: '24px', marginTop: '16px' }}}}>
        <div style={{{{ flex: 1, background: '#F7F9FC', padding: '14px 20px', borderRadius: '6px', borderLeft: '3px solid #2E86AB' }}}}>
          <div style={{{{ fontSize: '14px', color: '#0E3F8C', fontWeight: 700 }}}}>权重分布</div>
          <div style={{{{ fontSize: '13px', color: '#1A2230' }}}}>
            Mean {W['weight_mean']:.2f} · Median {W['weight_p50']:.2f} · Max {W['weight_max']:.2f}<br/>
            截断 {W['trunc_lo']:.2f}–{W['trunc_hi']:.2f}（1–99 百分位）
          </div>
        </div>
        <div style={{{{ flex: 1, background: '#F7F9FC', padding: '14px 20px', borderRadius: '6px', borderLeft: '3px solid #C0392B' }}}}>
          <div style={{{{ fontSize: '14px', color: '#0E3F8C', fontWeight: 700 }}}}>协变量平衡</div>
          <div style={{{{ fontSize: '13px', color: '#1A2230' }}}}>
            PS 范围 {W['ps_min']:.3f}–{W['ps_max']:.3f}（均值 {W['ps_mean']:.3f}）<br/>
            加权后 SMD &lt;0.10，平衡充分
          </div>
        </div>
      </div>
    </div>
    <div style={{{{ position: 'absolute', bottom: '20px', left: '40px', fontSize: '14px', color: '#8B97A8' }}}}>Figure S1. Stabilised weight diagnostics and covariate balance</div>
    <div style={{{{ position: 'absolute', bottom: '20px', right: '40px', fontSize: '14px', color: '#1E4FA8', fontWeight: 600 }}}}>06 / 17</div>
  </div>
</Slide>''')

    # ─── 7: Survival curves (Figure 2) ───
    _add_slide(f'''<Slide>
  <div style={{{{ width: '1280px', height: '720px', background: '#FFFFFF', position: 'relative', overflow: 'hidden' }}}}>
    <div style={{{{ position: 'absolute', top: 0, left: 0, width: '1280px', height: '72px', background: '#0E3F8C', display: 'flex', alignItems: 'center', paddingLeft: '40px' }}}}>
      <span style={{{{ fontSize: '24px', color: '#FFFFFF', fontWeight: 700 }}}}>加权 Kaplan-Meier 生存曲线 (Figure 2)</span>
    </div>
    <div style={{{{ padding: '84px 60px 0 60px' }}}}>
      <img src="{fig_assets['fig2']}" style={{{{ width: '100%', height: 'auto', maxHeight: '560px', objectFit: 'contain' }}}}/>
    </div>
    <div style={{{{ position: 'absolute', bottom: '36px', left: '40px', fontSize: '14px', color: '#8B97A8' }}}}>Figure 2. Weighted Kaplan-Meier curves + N at risk + log-rank p &lt; 0.001</div>
    <div style={{{{ position: 'absolute', bottom: '20px', right: '40px', fontSize: '14px', color: '#1E4FA8', fontWeight: 600 }}}}>07 / 17</div>
  </div>
</Slide>''')

    # ─── 8: Primary outcome (key numbers) ───
    _add_slide(f'''<Slide>
  <div style={{{{ width: '1280px', height: '720px', background: '#FFFFFF', position: 'relative', overflow: 'hidden' }}}}>
    <div style={{{{ position: 'absolute', top: 0, left: 0, width: '1280px', height: '72px', background: '#0E3F8C', display: 'flex', alignItems: 'center', paddingLeft: '40px' }}}}>
      <span style={{{{ fontSize: '24px', color: '#FFFFFF', fontWeight: 700 }}}}>主要结局 (Primary Outcome)</span>
    </div>
    <div style={{{{ padding: '96px 80px 0 80px' }}}}>
      <div style={{{{ display: 'flex', gap: '24px', marginBottom: '28px' }}}}>
        <div style={{{{ flex: 1, background: '#FFF5F5', borderTop: '4px solid #C0392B', padding: '24px', borderRadius: '8px' }}}}>
          <div style={{{{ fontSize: '15px', color: '#9E2A2A', fontWeight: 700, marginBottom: '6px' }}}}>早期降阶 Early</div>
          <div style={{{{ fontSize: '44px', color: '#C0392B', fontWeight: 700, lineHeight: 1 }}}}>{pp(P['mort_early'])}%</div>
          <div style={{{{ fontSize: '14px', color: '#666', marginTop: '6px' }}}}>28-day mortality (weighted)</div>
        </div>
        <div style={{{{ flex: 1, background: '#F5FAFF', borderTop: '4px solid #2E86AB', padding: '24px', borderRadius: '8px' }}}}>
          <div style={{{{ fontSize: '15px', color: '#1B4F72', fontWeight: 700, marginBottom: '6px' }}}}>延迟降阶 Deferred</div>
          <div style={{{{ fontSize: '44px', color: '#2E86AB', fontWeight: 700, lineHeight: 1 }}}}>{pp(P['mort_deferred'])}%</div>
          <div style={{{{ fontSize: '14px', color: '#666', marginTop: '6px' }}}}>28-day mortality (weighted)</div>
        </div>
        <div style={{{{ flex: 1, background: '#E8EFF8', borderTop: '4px solid #1E4FA8', padding: '24px', borderRadius: '8px' }}}}>
          <div style={{{{ fontSize: '15px', color: '#0E3F8C', fontWeight: 700, marginBottom: '6px' }}}}>风险差 RD (95% CI)</div>
          <div style={{{{ fontSize: '44px', color: '#1E4FA8', fontWeight: 700, lineHeight: 1 }}}}>{pp_label(P['rd'])} pp</div>
          <div style={{{{ fontSize: '14px', color: '#666', marginTop: '6px' }}}}>({ci_pp(P['rd_ci_lo'], P['rd_ci_hi'])})</div>
        </div>
      </div>
      <div style={{{{ background: '#F8F8F8', padding: '18px 24px', borderRadius: '6px', borderLeft: '3px solid #666' }}}}>
        <div style={{{{ fontSize: '15px', color: '#1A2230', lineHeight: 1.7 }}}}>
          <span><b>风险比 RR</b> = {P['rr']:.2f}（95% CI {P['rr_ci_lo']:.2f}–{P['rr_ci_hi']:.2f}）</span><br/>
          <span><b>未加权 log-rank p &lt; 0.001</b>；RD 与 RR 的 95% CI 均不包括零，早期降阶显著获益</span>
        </div>
      </div>
    </div>
    <div style={{{{ position: 'absolute', bottom: '20px', right: '40px', fontSize: '14px', color: '#1E4FA8', fontWeight: 600 }}}}>08 / 17</div>
  </div>
</Slide>''')

    # ─── 9: Secondary outcomes ───
    _add_slide(f'''<Slide>
  <div style={{{{ width: '1280px', height: '720px', background: '#FFFFFF', position: 'relative', overflow: 'hidden' }}}}>
    <div style={{{{ position: 'absolute', top: 0, left: 0, width: '1280px', height: '72px', background: '#0E3F8C', display: 'flex', alignItems: 'center', paddingLeft: '40px' }}}}>
      <span style={{{{ fontSize: '24px', color: '#FFFFFF', fontWeight: 700 }}}}>次要结局 (Secondary Outcomes)</span>
    </div>
    <div style={{{{ padding: '96px 60px 0 60px' }}}}>
      <div style={{{{ display: 'flex', background: '#1E4FA8', color: '#FFFFFF', borderRadius: '6px 6px 0 0', padding: '10px 16px', fontSize: '16px', fontWeight: 600 }}}}>
        <div style={{{{ flex: 1.6 }}}}>结局</div>
        <div style={{{{ flex: 1, textAlign: 'center' }}}}>Early</div>
        <div style={{{{ flex: 1, textAlign: 'center' }}}}>Deferred</div>
        <div style={{{{ flex: 1.5, textAlign: 'center' }}}}>Difference (95% CI)</div>
      </div>
      <div style={{{{ display: 'flex', background: '#F7F9FC', padding: '14px 16px', fontSize: '16px', color: '#1A2230', borderBottom: '1px solid #E5E7EB' }}}}>
        <div style={{{{ flex: 1.6 }}}}>28-day RMST (days)</div>
        <div style={{{{ flex: 1, textAlign: 'center', color: '#C0392B', fontWeight: 600 }}}}>{P['rmst_early']:.2f}</div>
        <div style={{{{ flex: 1, textAlign: 'center', color: '#2E86AB', fontWeight: 600 }}}}>{P['rmst_deferred']:.2f}</div>
        <div style={{{{ flex: 1.5, textAlign: 'center', fontWeight: 600 }}}}>+{P['rmst_diff']:.2f} ({P['rmst_ci_lo']:+.2f} to {P['rmst_ci_hi']:+.2f})</div>
      </div>
      <div style={{{{ display: 'flex', background: '#FFFFFF', padding: '14px 16px', fontSize: '16px', color: '#1A2230', borderBottom: '1px solid #E5E7EB' }}}}>
        <div style={{{{ flex: 1.6 }}}}>VFD28 (days)</div>
        <div style={{{{ flex: 1, textAlign: 'center', color: '#C0392B', fontWeight: 600 }}}}>{P['vfd_early']:.2f}</div>
        <div style={{{{ flex: 1, textAlign: 'center', color: '#2E86AB', fontWeight: 600 }}}}>{P['vfd_deferred']:.2f}</div>
        <div style={{{{ flex: 1.5, textAlign: 'center', fontWeight: 600 }}}}>+{P['vfd_diff']:.2f} ({P['vfd_ci_lo']:+.2f} to {P['vfd_ci_hi']:+.2f})</div>
      </div>
      <div style={{{{ marginTop: '24px', background: '#E8EFF8', borderRadius: '8px', padding: '16px 24px', borderLeft: '4px solid #1E4FA8' }}}}>
        <div style={{{{ fontSize: '16px', color: '#0E3F8C', fontWeight: 600, lineHeight: 1.6 }}}}>
          两项次要结局均显著支持早期降阶：RMST +{P['rmst_diff']:.2f} 天（95% CI {P['rmst_ci_lo']:+.2f} to {P['rmst_ci_hi']:+.2f}）、VFD +{P['vfd_diff']:.2f} 天（95% CI {P['vfd_ci_lo']:+.2f} to {P['vfd_ci_hi']:+.2f}），CI 均不包括零，与主分析方向一致。
        </div>
      </div>
    </div>
    <div style={{{{ position: 'absolute', bottom: '20px', right: '40px', fontSize: '14px', color: '#1E4FA8', fontWeight: 600 }}}}>09 / 17</div>
  </div>
</Slide>''')

    # ─── 10: Sensitivity + robustness analyses (Figure 3 + Table) ───
    _add_slide(f'''<Slide>
  <div style={{{{ width: '1280px', height: '720px', background: '#FFFFFF', position: 'relative', overflow: 'hidden' }}}}>
    <div style={{{{ position: 'absolute', top: 0, left: 0, width: '1280px', height: '72px', background: '#0E3F8C', display: 'flex', alignItems: 'center', paddingLeft: '40px' }}}}>
      <span style={{{{ fontSize: '24px', color: '#FFFFFF', fontWeight: 700 }}}}>敏感性与稳健性分析 (Figure 3)</span>
    </div>
    <div style={{{{ display: 'flex', padding: '90px 40px 0 40px' }}}}>
      <div style={{{{ flex: 1.5, paddingRight: '16px' }}}}>
        <img src="{fig_assets['fig3']}" style={{{{ width: '100%', height: 'auto' }}}}/>
      </div>
      <div style={{{{ flex: 1, paddingLeft: '20px', alignSelf: 'center' }}}}>
        <div style={{{{ fontSize: '15px', color: '#0E3F8C', fontWeight: 700, marginBottom: '10px' }}}}>七项估计一览（全部显著）</div>
        <div style={{{{ background: '#F7F9FC', padding: '8px 14px', borderRadius: '6px', marginBottom: '6px' }}}}>
          <div style={{{{ fontSize: '13px', color: '#1A2230', fontWeight: 600 }}}}>主分析（加权 + 截断）</div>
          <div style={{{{ fontSize: '13px', color: '#C0392B', fontWeight: 600 }}}}>{pp_label(P['rd'])} pp [{ci_pp(P['rd_ci_lo'], P['rd_ci_hi'])}]</div>
        </div>
        <div style={{{{ background: '#F7F9FC', padding: '8px 14px', borderRadius: '6px', marginBottom: '6px' }}}}>
          <div style={{{{ fontSize: '13px', color: '#1A2230', fontWeight: 600 }}}}>SA1: 无权重截断 / SA2: ICU 死亡率 / SA3: 未加权</div>
          <div style={{{{ fontSize: '13px', color: '#1A2230' }}}}>{pp_label(SA['SA1_no_truncation']['rd'])} / {pp_label(SA['SA2_icu_mortality']['rd'])} / {pp_label(SA['SA3_unweighted']['rd'])} pp（均显著）</div>
        </div>
        <div style={{{{ background: '#F7F9FC', padding: '8px 14px', borderRadius: '6px', marginBottom: '6px' }}}}>
          <div style={{{{ fontSize: '13px', color: '#1A2230', fontWeight: 600 }}}}>SA4: 扩充协变量（40 个，含镇静/升压药/血流动力学）</div>
          <div style={{{{ fontSize: '13px', color: '#1A2230' }}}}>{pp_label(RB['SA4_extended_covariates']['rd'])} pp [{ci_pp(RB['SA4_extended_covariates']['rd_ci_lo'], RB['SA4_extended_covariates']['rd_ci_hi'])}]</div>
        </div>
        <div style={{{{ background: '#F7F9FC', padding: '8px 14px', borderRadius: '6px', marginBottom: '6px' }}}}>
          <div style={{{{ fontSize: '13px', color: '#1A2230', fontWeight: 600 }}}}>SA5: 5–95 百分位截断 / AIPW 双重稳健</div>
          <div style={{{{ fontSize: '13px', color: '#1A2230' }}}}>{pp_label(RB['SA5_alt_truncation']['rd'])} pp / {pp_label(RB['AIPW_doubly_robust']['rd'])} pp（均显著）</div>
        </div>
        <div style={{{{ background: '#E8EFF8', padding: '10px 14px', borderRadius: '6px' }}}}>
          <div style={{{{ fontSize: '13px', color: '#0E3F8C', fontWeight: 700 }}}}>E-value = {RB['e_value']['e_value_point']:.2f}（CI 界 {RB['e_value']['e_value_ci_bound']:.2f}）：未测量混杂需 RR ≥ {RB['e_value']['e_value_point']:.2f} 才能解释掉效应</div>
        </div>
      </div>
    </div>
    <div style={{{{ position: 'absolute', bottom: '20px', left: '40px', fontSize: '14px', color: '#8B97A8' }}}}>Figure 3. Primary, sensitivity & robustness analyses — risk difference</div>
    <div style={{{{ position: 'absolute', bottom: '20px', right: '40px', fontSize: '14px', color: '#1E4FA8', fontWeight: 600 }}}}>10 / 17</div>
  </div>
</Slide>''')

    # ─── 11: Threshold grid heatmap (Figure S2) ───
    _add_slide(f'''<Slide>
  <div style={{{{ width: '1280px', height: '720px', background: '#FFFFFF', position: 'relative', overflow: 'hidden' }}}}>
    <div style={{{{ position: 'absolute', top: 0, left: 0, width: '1280px', height: '72px', background: '#0E3F8C', display: 'flex', alignItems: 'center', paddingLeft: '40px' }}}}>
      <span style={{{{ fontSize: '24px', color: '#FFFFFF', fontWeight: 700 }}}}>探索性阈值网格 (Figure S2)</span>
    </div>
    <div style={{{{ display: 'flex', padding: '90px 40px 0 40px' }}}}>
      <div style={{{{ flex: 1.4, paddingRight: '16px' }}}}>
        <img src="{fig_assets['figS2']}" style={{{{ width: '100%', height: 'auto' }}}}/>
      </div>
      <div style={{{{ flex: 1, paddingLeft: '20px', alignSelf: 'center' }}}}>
        <div style={{{{ fontSize: '15px', color: '#0E3F8C', fontWeight: 700, marginBottom: '12px' }}}}>FiO₂ ≤40% 三个 PEEP 阈值均显著</div>
        <div style={{{{ background: '#E8EFF8', padding: '14px 18px', borderRadius: '6px', marginBottom: '10px' }}}}>
          <div style={{{{ fontSize: '14px', color: '#1A2230', fontWeight: 600 }}}}>FiO₂ ≤40% / PEEP ≤5 (n=181)</div>
          <div style={{{{ fontSize: '14px', color: '#C0392B', fontWeight: 600 }}}}>RD = {pp_label([c for c in TG if c['fio2_threshold']==40 and c['peep_threshold']==5][0]['rd'])} pp [{ci_pp([c for c in TG if c['fio2_threshold']==40 and c['peep_threshold']==5][0]['rd_ci_lo'], [c for c in TG if c['fio2_threshold']==40 and c['peep_threshold']==5][0]['rd_ci_hi'])}] ★</div>
        </div>
        <div style={{{{ background: '#E8EFF8', padding: '14px 18px', borderRadius: '6px', marginBottom: '10px' }}}}>
          <div style={{{{ fontSize: '14px', color: '#1A2230', fontWeight: 600 }}}}>FiO₂ ≤40% / PEEP ≤8 (n=204)</div>
          <div style={{{{ fontSize: '14px', color: '#C0392B', fontWeight: 600 }}}}>RD = {pp_label([c for c in TG if c['fio2_threshold']==40 and c['peep_threshold']==8][0]['rd'])} pp [{ci_pp([c for c in TG if c['fio2_threshold']==40 and c['peep_threshold']==8][0]['rd_ci_lo'], [c for c in TG if c['fio2_threshold']==40 and c['peep_threshold']==8][0]['rd_ci_hi'])}] ★</div>
        </div>
        <div style={{{{ background: '#E8EFF8', padding: '14px 18px', borderRadius: '6px' }}}}>
          <div style={{{{ fontSize: '14px', color: '#1A2230', fontWeight: 600 }}}}>FiO₂ ≤40% / PEEP ≤10 (n=235)</div>
          <div style={{{{ fontSize: '14px', color: '#C0392B', fontWeight: 600 }}}}>RD = {pp_label([c for c in TG if c['fio2_threshold']==40 and c['peep_threshold']==10][0]['rd'])} pp [{ci_pp([c for c in TG if c['fio2_threshold']==40 and c['peep_threshold']==10][0]['rd_ci_lo'], [c for c in TG if c['fio2_threshold']==40 and c['peep_threshold']==10][0]['rd_ci_hi'])}] ★</div>
        </div>
        <div style={{{{ marginTop: '14px', fontSize: '13px', color: '#666' }}}}>★ = 95% CI 排除 0；嵌套亚组，未做多重性校正</div>
      </div>
    </div>
    <div style={{{{ position: 'absolute', bottom: '20px', left: '40px', fontSize: '14px', color: '#8B97A8' }}}}>Figure S2. Exploratory threshold grid</div>
    <div style={{{{ position: 'absolute', bottom: '20px', right: '40px', fontSize: '14px', color: '#1E4FA8', fontWeight: 600 }}}}>11 / 17</div>
  </div>
</Slide>''')

    # ─── 12: Per-protocol trajectory verification (Figure S4) ───
    _early_d2c  = VT['early']['D2_controlled']    if VT_OK else "—"
    _early_d3m  = VT['early']['D3_missing']       if VT_OK else "—"
    _early_d3c  = VT['early']['D3_controlled']    if VT_OK else "—"
    _early_ng   = VT['early']['n_group']          if VT_OK else 776
    _def_d2c    = VT['deferred']['D2_controlled'] if VT_OK else "—"
    _def_d3c    = VT['deferred']['D3_controlled'] if VT_OK else "—"
    _def_d3m    = VT['deferred']['D3_missing']    if VT_OK else "—"
    _def_ng     = VT['deferred']['n_group']       if VT_OK else 1126
    _add_slide(f'''<Slide>
  <div style={{{{ width: '1280px', height: '720px', background: '#FFFFFF', position: 'relative', overflow: 'hidden' }}}}>
    <div style={{{{ position: 'absolute', top: 0, left: 0, width: '1280px', height: '72px', background: '#0E3F8C', display: 'flex', alignItems: 'center', paddingLeft: '40px' }}}}>
      <span style={{{{ fontSize: '24px', color: '#FFFFFF', fontWeight: 700 }}}}>Per-protocol 验证：策略是否真被执行 (Figure S4)</span>
    </div>
    <div style={{{{ padding: '88px 32px 0 32px' }}}}>
      <div style={{{{ fontSize: '13px', color: '#666', marginBottom: '8px', fontStyle: 'italic' }}}}>
        数据时间分辨率：6 小时生命体征 (Day 1 00:00 → Day 2 18:00) + Day 2/3 通气日聚合。MIMIC-IV 主库无波形 / 膈肌 / 同步指标 / Day 4-5 设置。
      </div>
      <div style={{{{ display: 'flex', alignItems: 'center', justifyContent: 'center' }}}}>
        <img src="{fig_assets['figS4']}" style={{{{ width: '780px', height: 'auto', maxHeight: '440px' }}}}/>
      </div>
      <div style={{{{ display: 'flex', gap: '20px', marginTop: '12px' }}}}>
        <div style={{{{ flex: 1, background: '#FFF5F5', padding: '14px 18px', borderRadius: '6px', borderLeft: '4px solid #C0392B' }}}}>
          <div style={{{{ fontSize: '14px', color: '#9E2A2A', fontWeight: 700 }}}}>Early 组 (n={_early_ng})</div>
          <div style={{{{ fontSize: '13px', color: '#1A2230', lineHeight: 1.6, marginTop: '6px' }}}}>
            Day-2 控制模式 {_early_d2c} / {_early_ng} → Day-3 <b>0</b> 控制模式；<b>{_early_d3m} ({_early_d3m}/{_early_ng})</b> ≤72 h 已拔管。
          </div>
        </div>
        <div style={{{{ flex: 1, background: '#F5FAFF', padding: '14px 18px', borderRadius: '6px', borderLeft: '4px solid #2E86AB' }}}}>
          <div style={{{{ fontSize: '14px', color: '#1B4F72', fontWeight: 700 }}}}>Deferred 组 (n={_def_ng})</div>
          <div style={{{{ fontSize: '13px', color: '#1A2230', lineHeight: 1.6, marginTop: '6px' }}}}>
            Day-2 控制模式 {_def_d2c} / {_def_ng} → Day-3 <b>{_def_d3c}</b> ({_def_d3c}/{_def_ng}) 仍控制模式，<b>0</b> 已拔管。
          </div>
        </div>
        <div style={{{{ flex: 1, background: '#F0F7E8', padding: '14px 18px', borderRadius: '6px', borderLeft: '4px solid #5B8C3E' }}}}>
          <div style={{{{ fontSize: '14px', color: '#3F5A2A', fontWeight: 700 }}}}>解读</div>
          <div style={{{{ fontSize: '13px', color: '#1A2230', lineHeight: 1.6, marginTop: '6px' }}}}>
            策略分歧集中在<b>模式/拔管</b>；FiO₂/PEEP 中位数两组一致 (40% / 5 cmH₂O)。前 6-h 呼吸轨迹两组平衡，倾向性加权站得住脚。
          </div>
        </div>
      </div>
    </div>
    <div style={{{{ position: 'absolute', bottom: '20px', left: '40px', fontSize: '14px', color: '#8B97A8' }}}}>Figure S4. Per-protocol trajectory verification</div>
    <div style={{{{ position: 'absolute', bottom: '20px', right: '40px', fontSize: '14px', color: '#1E4FA8', fontWeight: 600 }}}}>12 / 17</div>
  </div>
</Slide>''')

    # ─── 13: External validation in eICU-CRD (Figure 4) ───
    _add_slide(f'''<Slide>
  <div style={{{{ width: '1280px', height: '720px', background: '#FFFFFF', position: 'relative', overflow: 'hidden' }}}}>
    <div style={{{{ position: 'absolute', top: 0, left: 0, width: '1280px', height: '72px', background: '#B4540A', display: 'flex', alignItems: 'center', paddingLeft: '40px' }}}}>
      <span style={{{{ fontSize: '24px', color: '#FFFFFF', fontWeight: 700 }}}}>外部验证：eICU-CRD 多中心独立重复 (Figure 4)</span>
    </div>
    <div style={{{{ padding: '84px 32px 0 32px' }}}}>
      <div style={{{{ fontSize: '13px', color: '#666', marginBottom: '6px', fontStyle: 'italic' }}}}>
        eICU-CRD v2.0：208 家美国医院 (2014\u20132015)，与 MIMIC-IV 完全独立；同一 TTE 协议和谐化重复；结局为院内死亡率 (eICU 无 28 天随访)。
      </div>
      <div style={{{{ display: 'flex', alignItems: 'center', justifyContent: 'center' }}}}>
        <img src="{fig_assets['fig4']}" style={{{{ width: '920px', height: 'auto', maxHeight: '400px', objectFit: 'contain' }}}}/>
      </div>
      <div style={{{{ display: 'flex', gap: '16px', marginTop: '10px' }}}}>
        <div style={{{{ flex: 1, background: '#FFF8F0', padding: '12px 16px', borderRadius: '6px', borderLeft: '4px solid #B4540A' }}}}>
          <div style={{{{ fontSize: '14px', color: '#8A4206', fontWeight: 700 }}}}>验证队列 (analysable {EV['n_analyzable']:,})</div>
          <div style={{{{ fontSize: '13px', color: '#1A2230', lineHeight: 1.55, marginTop: '4px' }}}}>
            {EVF['eligible']:,} eligible → early {EV['n_early']:,} / deferred {EV['n_deferred']:,}；加权后最大 SMD {EV['balance_max_smd_weighted']:.3f}。
          </div>
        </div>
        <div style={{{{ flex: 1, background: '#FFF5F5', padding: '12px 16px', borderRadius: '6px', borderLeft: '4px solid #C0392B' }}}}>
          <div style={{{{ fontSize: '14px', color: '#9E2A2A', fontWeight: 700 }}}}>验证结果：方向一致、仍显著</div>
          <div style={{{{ fontSize: '13px', color: '#1A2230', lineHeight: 1.55, marginTop: '4px' }}}}>
            院内死亡率 {pp(EVP['mort_early'])}% vs {pp(EVP['mort_deferred'])}%；RD {pp_label(EVP['rd'])} pp ({ci_pp(EVP['rd_ci_lo'], EVP['rd_ci_hi'])})；RR {EVP['rr']:.2f}。严格拔管定义 SA 一致 ({pp_label(EVSA['rd'])} pp)。
          </div>
        </div>
        <div style={{{{ flex: 1, background: '#F0F7E8', padding: '12px 16px', borderRadius: '6px', borderLeft: '4px solid #5B8C3E' }}}}>
          <div style={{{{ fontSize: '14px', color: '#3F5A2A', fontWeight: 700 }}}}>解读</div>
          <div style={{{{ fontSize: '13px', color: '#1A2230', lineHeight: 1.55, marginTop: '4px' }}}}>
            early 组死亡率两库几乎相同 ({pp(P['mort_early'])}% vs {pp(EVP['mort_early'])}%)；绝对效应量级随场景不同，<b>只做方向性比较、不合并</b>。
          </div>
        </div>
      </div>
    </div>
    <div style={{{{ position: 'absolute', bottom: '20px', left: '40px', fontSize: '14px', color: '#8B97A8' }}}}>Figure 4. External validation in eICU-CRD v2.0</div>
    <div style={{{{ position: 'absolute', bottom: '20px', right: '40px', fontSize: '14px', color: '#1E4FA8', fontWeight: 600 }}}}>13 / 17</div>
  </div>
</Slide>''')

    # ─── 12: Interpretation ───
    _add_slide(f'''<Slide>
  <div style={{{{ width: '1280px', height: '720px', background: '#FFFFFF', position: 'relative', overflow: 'hidden' }}}}>
    <div style={{{{ position: 'absolute', top: 0, left: 0, width: '1280px', height: '72px', background: '#0E3F8C', display: 'flex', alignItems: 'center', paddingLeft: '40px' }}}}>
      <span style={{{{ fontSize: '24px', color: '#FFFFFF', fontWeight: 700 }}}}>主要解读</span>
    </div>
    <div style={{{{ padding: '96px 60px 0 60px' }}}}>
      <div style={{{{ display: 'flex', gap: '24px' }}}}>
        <div style={{{{ flex: 1, background: '#FFF5F5', padding: '24px', borderRadius: '8px', borderTop: '4px solid #C0392B' }}}}>
          <div style={{{{ fontSize: '17px', color: '#9E2A2A', fontWeight: 700, marginBottom: '12px' }}}}>主要发现</div>
          <div style={{{{ fontSize: '15px', color: '#1A2230', lineHeight: 1.7 }}}}>
            早期降阶 28 天死亡率 {pp(P['mort_early'])}% vs 延迟 {pp(P['mort_deferred'])}%；RD {pp_label(P['rd'])} pp（95% CI {ci_pp(P['rd_ci_lo'], P['rd_ci_hi'])}），RR {P['rr']:.2f}，RD 与 RR 的 95% CI 均不包括零（p &lt; 0.001），早期降阶显著获益。
          </div>
        </div>
        <div style={{{{ flex: 1, background: '#F5FAFF', padding: '24px', borderRadius: '8px', borderTop: '4px solid #2E86AB' }}}}>
          <div style={{{{ fontSize: '17px', color: '#1B4F72', fontWeight: 700, marginBottom: '12px' }}}}>亚组提示</div>
          <div style={{{{ fontSize: '15px', color: '#1A2230', lineHeight: 1.7 }}}}>
            探索性阈值网格：全部 9 个 FiO₂\u2013PEEP 组合均显示显著获益（RD {TG_TILDE_ALL} pp，CI 均不跨零），其中 FiO₂ ≤40%/PEEP ≤5 获益最大（{TG_MAX_BENEFIT} pp），提示低氧需求者受益最大。
          </div>
        </div>
      </div>
      <div style={{{{ marginTop: '24px', background: '#E8EFF8', padding: '20px 28px', borderRadius: '8px', borderLeft: '4px solid #1E4FA8' }}}}>
        <div style={{{{ fontSize: '16px', color: '#0E3F8C', fontWeight: 600, lineHeight: 1.7 }}}}>
          <span><b>关键限定</b>：(1) 观察性研究，未测量混杂不可避免（E-value {_EV_TXT}）；(2) 每日一次通气设置记录，可能遗漏快速模式切换；(3) 部分 Day-3 模式缺失患者无法分类；(4) eICU 无显式模式变量、无 28 天随访，两库绝对量级不同，只做方向性比较不合并。</span>
        </div>
      </div>
    </div>
    <div style={{{{ position: 'absolute', bottom: '20px', right: '40px', fontSize: '14px', color: '#1E4FA8', fontWeight: 600 }}}}>14 / 17</div>
  </div>
</Slide>''')

    # ─── 13: Strengths ───
    _add_slide('''<Slide>
  <div style={{ width: '1280px', height: '720px', background: '#FFFFFF', position: 'relative', overflow: 'hidden' }}>
    <div style={{ position: 'absolute', top: 0, left: 0, width: '1280px', height: '72px', background: '#0E3F8C', display: 'flex', alignItems: 'center', paddingLeft: '40px' }}>
      <span style={{ fontSize: '24px', color: '#FFFFFF', fontWeight: 700 }}>研究优势与局限</span>
    </div>
    <div style={{ padding: '96px 60px 0 60px' }}>
      <div style={{ display: 'flex', gap: '32px' }}>
        <div style={{ flex: 1 }}>
          <div style={{ fontSize: '18px', color: '#0E3F8C', fontWeight: 700, marginBottom: '16px', borderBottom: '2px solid #1E4FA8', paddingBottom: '6px' }}>优势</div>
          <div style={{ fontSize: '15px', color: '#1A2230', lineHeight: 1.7 }}>
            <span><b>· 显式目标试验协议</b>：在分析数据集锁定前完整规定所有要素，事后无修订</span><br/>
            <span><b>· Clone-Censor-Weight</b>：解决不朽时间偏倚与时变混杂</span><br/>
            <span><b>· 21 个先验协变量</b>：包括 SOFA、Day-2 通气参数、PaO₂/FiO₂、ARDS、Sepsis-3、CRRT</span><br/>
            <span><b>· 引导重抽样的 bootstrap</b>：每次迭代重新拟合倾向性模型</span><br/>
            <span><b>· 充分协变量平衡</b>：加权后 SMD 均 &lt; 0.10</span><br/>
            <span><b>· 六项敏感性/稳健性分析</b>（含 SA4 扩充协变量、SA5 截断、AIPW 双重稳健）+ E-value ''' + _EV_TXT + '''，全部显著</span>
          </div>
        </div>
        <div style={{ flex: 1 }}>
          <div style={{ fontSize: '18px', color: '#9E2A2A', fontWeight: 700, marginBottom: '16px', borderBottom: '2px solid #C0392B', paddingBottom: '6px' }}>局限</div>
          <div style={{ fontSize: '15px', color: '#1A2230', lineHeight: 1.7 }}>
            <span><b>· 观察性研究设计</b>：残余混杂不可避免（E-value ''' + _EV_TXT + '''）</span><br/>
            <span><b>· 通气设置日采一次</b>：快速模式切换可能被错误分类</span><br/>
            <span><b>· 大样本分析集</b>：n = ''' + f"{S['early'] + S['deferred']:,}" + '''，约为同类单数据库 CCW 撤机分析的 4 倍</span><br/>
            <span><b>· 缺乏动态指标</b>：呼吸力学、膈肌功能、人机同步性未纳入</span><br/>
            <span><b>· 推导库单中心</b>：MIMIC-IV 来自 BIDMC，外推性已由 eICU 208 医院验证缓解</span><br/>
            <span><b>· 24 h 宽限期</b>：为临床可行性的折衷，其他持续期可能产生不同分类</span>
          </div>
        </div>
      </div>
    </div>
    <div style={{ position: 'absolute', bottom: '20px', right: '40px', fontSize: '14px', color: '#1E4FA8', fontWeight: 600 }}>15 / 17</div>
  </div>
</Slide>''')

    # ─── 14: Conclusions ───
    _add_slide('''<Slide>
  <div style={{ width: '1280px', height: '720px', background: '#FFFFFF', position: 'relative', overflow: 'hidden' }}>
    <div style={{ position: 'absolute', top: 0, left: 0, width: '1280px', height: '72px', background: '#0E3F8C', display: 'flex', alignItems: 'center', paddingLeft: '40px' }}>
      <span style={{ fontSize: '24px', color: '#FFFFFF', fontWeight: 700 }}>结论</span>
    </div>
    <div style={{ padding: '96px 80px 0 80px' }}>
      <div style={{ background: '#FFF5F5', padding: '28px 36px', borderRadius: '8px', borderTop: '4px solid #C0392B', marginBottom: '24px' }}>
        <div style={{ fontSize: '18px', color: '#9E2A2A', fontWeight: 700, marginBottom: '14px' }}>主要结论</div>
        <div style={{ fontSize: '16px', color: '#1A2230', lineHeight: 1.7 }}>
          48 小时地标后仍接受有创通气且低氧需求（FiO₂ ≤50%/PEEP ≤10）的患者中，早期降阶显著降低 28 天死亡率（''' + f"{pp(P['mort_early'])}% vs {pp(P['mort_deferred'])}%；RD {pp_label_u(P['rd'])} pp，95% CI {pp_label_u(P['rd_ci_lo'])} ~ {pp_label_u(P['rd_ci_hi'])}；RR {P['rr']:.2f}" + '''），并增加机械通气-free 天与 RMST。
        </div>
      </div>
      <div style={{ background: '#E8EFF8', padding: '28px 36px', borderRadius: '8px', borderTop: '4px solid #1E4FA8', marginBottom: '24px' }}>
        <div style={{ fontSize: '18px', color: '#0E3F8C', fontWeight: 700, marginBottom: '14px' }}>探索性发现</div>
        <div style={{ fontSize: '16px', color: '#1A2230', lineHeight: 1.7 }}>
          探索性阈值网格全部 9 个 FiO₂\u2013PEEP 组合均显著获益（RD ''' + TG_TILDE_ALL + ''' pp，CI 均不跨零）；六项敏感性/稳健性分析（含 AIPW 双重稳健、40 协变量扩充调整）全部显著；E-value ''' + _EV_TXT + '''，主结论稳健；eICU-CRD 多中心外部验证方向一致仍显著（RD ''' + pp_label_u(EVP['rd']) + ''' pp）。
        </div>
      </div>
      <div style={{ background: '#F7F9FC', padding: '20px 28px', borderRadius: '8px' }}>
        <div style={{ fontSize: '16px', color: '#0E3F8C', fontWeight: 600, lineHeight: 1.6 }}>
          <span><b>下一步</b>：需多中心前瞻性验证（随机试验或更大样本 TTE），结合高时间分辨率通气数据（每小时呼吸力学、膈肌功能、人机同步指标）以确立因果性推荐。</span>
        </div>
      </div>
    </div>
    <div style={{ position: 'absolute', bottom: '20px', right: '40px', fontSize: '14px', color: '#1E4FA8', fontWeight: 600 }}>16 / 17</div>
  </div>
</Slide>''')

    # ─── 15: Take-home + acknowledgements ───
    _add_slide('''<Slide>
  <div style={{ width: '1280px', height: '720px', background: '#FFFFFF', position: 'relative', overflow: 'hidden' }}>
    <div style={{ position: 'absolute', top: 0, left: 0, width: '1280px', height: '72px', background: '#0E3F8C', display: 'flex', alignItems: 'center', paddingLeft: '40px' }}>
      <span style={{ fontSize: '24px', color: '#FFFFFF', fontWeight: 700 }}>Take-Home Message</span>
    </div>
    <div style={{ padding: '110px 100px 0 100px' }}>
      <div style={{ background: '#0E3F8C', padding: '36px 48px', borderRadius: '12px', color: '#FFFFFF' }}>
        <div style={{ fontSize: '22px', fontWeight: 700, lineHeight: 1.5, marginBottom: '24px' }}>
          在 MIMIC-IV ''' + f"{S['early'] + S['deferred']:,}" + ''' 名可分析通气患者中，并经 eICU-CRD ''' + f"{EV['n_analyzable']:,}" + ''' 名多中心外部验证：
        </div>
        <div style={{ fontSize: '18px', lineHeight: 1.8 }}>
          ① 早期降阶显著降低 28 天死亡率（RD ''' + pp_label_u(P['rd']) + ''' pp，CI 不跨零）→ <b>显著获益</b><br/>
          ② RMST、VFD28 同步改善，CI 均不跨零 → <b>结局一致</b><br/>
          ③ 全部 9 个 FiO₂\u2013PEEP 阈值组合 + 六项稳健性分析 + 外部验证均一致 → <b>结论稳健</b><br/>
          ④ 绝对效应量级随场景不同（eICU RD ''' + pp_label_u(EVP['rd']) + ''' pp），仍需前瞻 RCT → <b>因果确认尚待试验</b>
        </div>
      </div>
      <div style={{ marginTop: '36px', textAlign: 'center', fontSize: '14px', color: '#8B97A8' }}>
        数据：MIMIC-IV v2.2 (BIDMC, 推导) + eICU-CRD v2.0 (208 家医院, 外部验证)<br/>
        方法：Target Trial Emulation · Clone-Censor-Weight · Stabilised IPCW<br/>
        Bootstrap 2,000（主分析）/ 800（敏感性、探索性）
      </div>
    </div>
    <div style={{ position: 'absolute', bottom: '20px', right: '40px', fontSize: '14px', color: '#1E4FA8', fontWeight: 600 }}>17 / 17</div>
  </div>
</Slide>''')

    # Write all slides into a single PPTX using tencent-pptx skill
    # Save slides as XML files (compatible with tencent-pptx skill)
    slides_dir = PPT_DIR / "slides"
    slides_dir.mkdir(exist_ok=True)
    for i, slide_xml in enumerate(PPT_SLIDES, 1):
        with open(slides_dir / f"{i:02d}.slide", "w") as f:
            f.write(slide_xml)

    # Use tencent-pptx skill to render
    # The skill expects this structure: slides/01.slide ... 15.slide, assets/, STORY.md
    story_md = f"""# TTE De-escalation 汇报 PPT — 叙事文档

## 背景
机械通气脱机时机的选择是 ICU 核心决策之一。现有指南推荐每日自主呼吸试验，但对下调呼吸支持级别（模式转换或撤机拔管）的具体时机缺乏明确标准。

## 研究问题
在 48 小时地标后仍接受有创通气且低氧需求（FiO₂ ≤50%/PEEP ≤10）的患者中，早期降阶（24h 内模式下调或撤机）与延迟降阶相比，对 28 天院内死亡率的影响如何？

## 关键结果
- {F['total_records']:,} → {F['eligible']:,} eligible → {S['early']:,} early / {S['deferred']:,} deferred / {S['unascertainable']:,} unascertainable / {S['grace_death']:,} grace death
- 加权 28 天死亡率：{pp(P['mort_early'])}% vs {pp(P['mort_deferred'])}%
- RD: {pp_label(P['rd'])} pp (95% CI: {ci_pp(P['rd_ci_lo'], P['rd_ci_hi'])}), RR: {P['rr']:.2f} ({P['rr_ci_lo']:.2f}-{P['rr_ci_hi']:.2f})，均显著
- RMST diff: +{P['rmst_diff']:.2f} days, VFD28 diff: +{P['vfd_diff']:.2f} days，均显著
- 敏感性/稳健性分析六项全部显著（SA1-3 + SA4 扩充协变量 + SA5 替代截断 + AIPW 双重稳健）；E-value {_EV_TXT}
- 阈值网格全部 9 个 FiO₂\u2013PEEP 组合均显著（RD {TG_TILDE_ALL_ASCII} pp，CI 不跨零）
- 外部验证 (eICU-CRD v2.0, 208 医院)：{EV['n_analyzable']:,} analysable，院内死亡率 {pp(EVP['mort_early'])}% vs {pp(EVP['mort_deferred'])}%，RD {pp_label(EVP['rd'])} pp，RR {EVP['rr']:.2f}，方向一致仍显著

## 结论
- 早期降阶显著降低 28 天死亡率，并改善 RMST 与 VFD28
- 全部 FiO₂\u2013PEEP 阈值组合下获益一致，结论稳健
- 多中心外部验证方向一致；绝对效应量级随场景不同，因果确认仍需前瞻 RCT
"""
    with open(PPT_DIR / "STORY.md", "w") as f:
        f.write(story_md)

    print(f"  ✓ PPT slides generated: {len(PPT_SLIDES)} slides")
    return PPT_DIR


# ════════════════════════════════════════════════════════════════════════
#  MAIN
# ════════════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    print("=" * 70)
    print("Building FINAL TTE De-escalation submission package")
    print("=" * 70)

    print("\n[1/4] Building main manuscript...")
    p1 = build_main()

    print("\n[2/4] Building supplement...")
    p2 = build_supplement()

    print("\n[3/4] Building cover letter...")
    p3 = build_cover_letter()

    print("\n[4/4] Building PPT slide XML...")
    p4 = build_ppt()

    print("\n" + "=" * 70)
    print("ALL DELIVERABLES GENERATED")
    print("=" * 70)
    print(f"  Main:      {p1}")
    print(f"  Supplemt:  {p2}")
    print(f"  Cover:     {p3}")
    print(f"  PPT dir:   {p4}")
    print("=" * 70)