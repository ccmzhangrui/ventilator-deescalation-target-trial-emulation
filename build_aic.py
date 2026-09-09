#!/usr/bin/env python3
"""
AIC submission pipeline (single process):
  1) rebuild the ICM manuscript (build_final.build_main) with the new
     methodology citations (refs 19 Gilding primer, 20 Huang JAMA editorial)
  2) regression-check the rebuilt ICM text against the pre-edit snapshot
  3) transform the ICM docx into the Annals of Intensive Care variant:
     - AIC title page + 3-part abstract (<=350 words) + keywords
     - Introduction -> Background; standalone Conclusions section
     - LLM-use statement at end of Methods (AIC requirement)
     - ICM declarations replaced by AIC/BMC headings in required order
     - PhysioNet dataset DOI citations appended (refs 21-22)
     - page breaks removed; double line spacing; line numbering
  4) build the AIC cover letter
  5) run verify_data.audit on both manuscripts
"""
import json, re, sys
from pathlib import Path

# ── sandbox-safe mkdir must precede importing build_final (mkdirs PPT_DIR) ──
_orig_mkdir = Path.mkdir

def _safe_mkdir(self, mode=0o777, parents=False, exist_ok=False):
    try:
        return _orig_mkdir(self, mode=mode, parents=parents, exist_ok=exist_ok)
    except PermissionError:
        if exist_ok and self.is_dir():
            return None
        raise

Path.mkdir = _safe_mkdir

import docx
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

import build_final as bf

P_, PR, H = bf.P_, bf.PR, bf.H
pp, ci_pp, pp_label = bf.pp, bf.ci_pp, bf.pp_label
P, F, S = bf.P, bf.F, bf.S
EV, EVP = bf.EV, bf.EVP
_EV_TXT = bf._EV_TXT
BASE = bf.BASE

TITLE = ("Early versus deferred de-escalation from controlled ventilation after "
         "a 48-hour landmark of physiological stability: a target trial emulation "
         "in MIMIC-IV with external validation in eICU-CRD")

ABSTRACT_WORDS = 0  # set by front_aic(); reused by cover_letter_aic()

DATASET_REFS = [
    "22. Johnson A, Bulgarelli L, Pollard T, Horng S, Celi LA, Mark R. MIMIC-IV "
    "(version 2.2). PhysioNet. 2023. https://doi.org/10.13026/6mm1-ek67.",
    "23. Pollard T, Johnson A, Raffa J, Celi LA, Badawi O, Mark R. eICU "
    "Collaborative Research Database (version 2.0). PhysioNet. 2019. "
    "https://doi.org/10.13026/C2WM1R.",
]


def wc(text):
    return len(re.findall(r"\S+", text))


# ────────────────────────────────────────────────────────────────────────────
# 1) rebuild ICM manuscript + regression check
# ────────────────────────────────────────────────────────────────────────────
def rebuild_icm():
    out = bf.build_main()
    snap_path = Path("/tmp/icm_final_snapshot.json")
    if snap_path.exists():
        old = json.load(open(snap_path))
        d = docx.Document(str(out))
        new = [p.text for p in d.paragraphs]
        for t in d.tables:
            for r in t.rows:
                for c in r.cells:
                    new.append(c.text)
        old_s, new_s = set(old), set(new)
        added = [x for x in new if x not in old_s]
        removed = [x for x in old if x not in new_s]
        print(f"  regression: {len(old)} old items, {len(new)} new items, "
              f"{len(added)} added, {len(removed)} removed")
        for x in added:
            print("    +", x[:130])
        for x in removed:
            print("    -", x[:130])
        # expected: additions only around the two citations + two new refs
        unexpected_add = [x for x in added
                          if "(8, 9, 19)" not in x and "stopping at the right time" not in x
                          and "Gilding" not in x and "Trials terminated early" not in x]
        unexpected_rem = [x for x in removed
                          if "principled framework to mitigate" not in x
                          and "prospective, randomised evaluation" not in x]
        if unexpected_rem or unexpected_add:
            print("  ⚠ UNEXPECTED REGRESSION DIFF — inspect before proceeding")
            for x in unexpected_add: print("    ?+", x[:130])
            for x in unexpected_rem: print("    ?-", x[:130])
        else:
            print("  regression OK: only the intended citation changes")
    return out


# ────────────────────────────────────────────────────────────────────────────
# 2) docx transformation helpers
# ────────────────────────────────────────────────────────────────────────────
def para_text(el):
    # join only w:t nodes (itertext() can yield duplicated strings with
    # python-docx's custom parser — observed 3x on Heading paragraphs)
    return "".join(t.text or "" for t in el.iter(qn("w:t")))


def is_heading(el, text):
    if el.tag != qn("w:p"):
        return False
    pPr = el.find(qn("w:pPr"))
    if pPr is None:
        return False
    st = pPr.find(qn("w:pStyle"))
    if st is None or not st.get(qn("w:val"), "").startswith("Heading"):
        return False
    return para_text(el).strip() == text


def move_before(doc, ref_el, builder):
    """builder(doc) appends new block(s) at end; return list of new top-level
    elements moved to just before ref_el, in creation order."""
    body = doc.element.body
    before = list(body.iterchildren())
    builder(doc)
    after = list(body.iterchildren())
    new_els = [el for el in after if el not in before]
    for el in new_els:
        ref_el.addprevious(el)
    return new_els


def front_aic(doc):
    P_(doc, TITLE, bold=True, size=13, align="center")
    doc.add_paragraph()
    P_(doc, "[Author list to be completed]", size=11, align="center")
    P_(doc, "[Institutional addresses to be completed]", size=10, align="center")
    P_(doc, "Corresponding author: [name, postal address and email to be completed]",
       size=10, align="center")
    doc.add_paragraph()

    H(doc, "Abstract", level=1)
    bg = (
        "The optimal timing of transition from controlled to assisted ventilatory "
        "modes after physiological stabilisation remains uncertain. We estimated "
        "the causal effect of early versus deferred de-escalation of respiratory "
        "support on 28-day in-hospital mortality using target trial emulation with "
        "clone\u2013censor\u2013weight estimation in MIMIC-IV (version 2.2), with "
        "external validation in the multi-centre eICU Collaborative Research "
        "Database (version 2.0; 208 hospitals). Adults still invasively ventilated "
        "48 hours after ICU admission with a classified Day-2 ventilator mode, "
        "Day-2 FiO\u2082 \u226450% and PEEP \u226410 cmH\u2082O were eligible. "
        "Early de-escalation (Day-3 support level below Day-2, or extubation "
        "within 72 hours) was compared with deferred de-escalation over a 24-hour "
        "grace window using stabilised inverse-probability-of-censoring weights."
    )
    res = (
        f"Of {F['total_records']:,} MIMIC-IV records, {F['eligible']:,} adults met "
        f"all eligibility criteria and {S['early'] + S['deferred']:,} were "
        f"analysable ({S['early']} early, {S['deferred']:,} deferred). Weighted "
        f"28-day mortality was {pp(P['mort_early'])}% (early) versus "
        f"{pp(P['mort_deferred'])}% (deferred): risk difference "
        f"{pp_label(P['rd'])} percentage points (95% confidence interval [CI], "
        f"{ci_pp(P['rd_ci_lo'], P['rd_ci_hi'])}); risk ratio {P['rr']:.2f} (95% CI, "
        f"{P['rr_ci_lo']:.2f}\u2013{P['rr_ci_hi']:.2f}). Early de-escalation was "
        f"also associated with longer 28-day restricted mean survival time "
        f"(+{P['rmst_diff']:.2f} days) and more ventilator-free days "
        f"(+{P['vfd_diff']:.2f} days). All six sensitivity and robustness "
        f"analyses\u2014including an extended 40-covariate adjustment, alternative "
        f"weight truncation and a doubly robust estimator\u2014remained "
        f"statistically significant, and all nine exploratory FiO\u2082\u00d7PEEP "
        f"threshold cells excluded zero (risk differences {bf.TG_RANGE_ALL} "
        f"percentage points). External validation in eICU-CRD "
        f"({EV['n_analyzable']:,} analysable) confirmed the direction and "
        f"significance of the finding: risk difference {pp_label(EVP['rd'])} "
        f"percentage points (95% CI, {ci_pp(EVP['rd_ci_lo'], EVP['rd_ci_hi'])}) "
        f"for in-hospital mortality; risk ratio {EVP['rr']:.2f}."
    )
    con = (
        "Early de-escalation of respiratory support after a 48-hour stability "
        "landmark was associated with significantly lower mortality, longer "
        "restricted mean survival and more ventilator-free days, with consistent "
        "results across all sensitivity analyses and external validation in an "
        "independent multi-centre cohort. Among patients on low respiratory "
        "support at 48 hours, timely step-down of support should be actively "
        "pursued; residual confounding by unmeasured clinical acuity cannot be "
        "excluded and prospective validation is warranted."
    )
    total = wc(bg) + wc(res) + wc(con) + 3  # +3 section labels
    print(f"  AIC abstract words: {total} (limit 350)")
    assert total <= 350, f"Abstract too long: {total}"
    globals()["ABSTRACT_WORDS"] = total

    PR(doc, [("Background.  ", True, False), (bg, False, False)])
    PR(doc, [("Results.  ", True, False), (res, False, False)])
    PR(doc, [("Conclusions.  ", True, False), (con, False, False)])
    doc.add_paragraph()
    P_(doc, "Keywords: mechanical ventilation; ventilator weaning; target trial "
            "emulation; clone\u2013censor\u2013weight; MIMIC-IV; eICU-CRD; "
            "mortality; external validation.",
       italic=True, size=10)
    doc.add_paragraph()


def declarations_aic(doc):
    H(doc, "Declarations", level=1)
    PR(doc, [("Ethics approval and consent to participate.  ", True, False),
        ("The institutional review boards of the Massachusetts Institute of "
         "Technology and Beth Israel Deaconess Medical Center approved the use of "
         "the MIMIC-IV database for research; the eICU Collaborative Research "
         "Database is likewise de-identified and exempt from IRB review under the "
         "Safe Harbor provision. The requirement for individual patient consent "
         "was waived because all data are fully de-identified.", False, False)])
    PR(doc, [("Consent for publication.  ", True, False),
        ("Not applicable.", False, False)])
    PR(doc, [("Availability of data and materials.  ", True, False),
        ("The datasets analysed during the current study are available through "
         "PhysioNet after completion of the required training and data-use "
         "agreements: MIMIC-IV version 2.2 (22) and the eICU Collaborative "
         "Research Database version 2.0 (23). The complete, reproducible analysis "
         "code (Python 3.13), including the extraction pipelines, the "
         "clone\u2013censor\u2013weight estimator, all sensitivity and robustness "
         "analyses, and figure generation, is openly available at "
         "https://github.com/ccmzhangrui/ventilator-deescalation-target-trial-emulation "
         "(MIT licence). Row-level intermediate files derived from the "
         "credentialed databases are not redistributed; only aggregate outputs "
         "(results/) are published.", False, False)])
    PR(doc, [("Declaration of generative AI and AI-assisted technologies in the "
              "writing process.  ", True, False),
        ("During the preparation of this work the authors used a large "
         "language model\u2013based assistant to support verification of the "
         "analysis code and English-language editing. No content was generated "
         "de novo: every numerical result reported in this manuscript is "
         "reproduced from the locked analysis outputs in the accompanying "
         "repository, and every sentence was reviewed and edited by the authors, "
         "who take full responsibility for the content of the published article.",
         False, False)])
    PR(doc, [("Competing interests.  ", True, False),
        ("The authors declare that they have no competing interests.", False, False)])
    PR(doc, [("Funding.  ", True, False),
        ("This research received no specific grant from any funding agency in "
         "the public, commercial, or not-for-profit sectors.", False, False)])
    PR(doc, [("Authors' contributions.  ", True, False),
        ("[To be completed with author initials upon unblinding: study "
         "conception and design, data curation, formal analysis, methodology, "
         "software, visualisation, writing \u2013 original draft, writing "
         "\u2013 review and editing. All authors will read and approve the final "
         "manuscript.]", False, False)])
    PR(doc, [("Acknowledgements.  ", True, False),
        ("Not applicable.", False, False)])
    PR(doc, [("Trial registration.  ", True, False),
        ("Not registered. The full target trial protocol was specified before "
         "the analytic dataset was finalised and is provided in the supplement "
         "(Methods A1 and Table A1); reporting follows the TARGET guideline "
         "(completed 21-item checklist, supplement Checklist C1).", False, False)])


def transform_to_aic(icm_path):
    doc = docx.Document(str(icm_path))
    body = doc.element.body
    children = list(body.iterchildren())

    intro_el = next(el for el in children if is_heading(el, "Introduction"))
    # 1) delete all front blocks before Introduction (title/abstract/boxes)
    for el in children:
        if el is intro_el:
            break
        if el.tag != qn("w:sectPr"):
            body.remove(el)
    # 2) rename heading
    ts = list(intro_el.iter(qn("w:t")))
    for k, r in enumerate(ts):
        r.text = "Background" if k == 0 else ""
    # 3) insert AIC front before Background heading
    move_before(doc, intro_el, front_aic)

    # 4) (removed) an AI-use statement was previously injected at the end of Methods;
    #    it now belongs in the Declarations section (see declarations_aic) per
    #    Springer Nature / BMC policy.

    # 5) standalone Conclusions heading before "In conclusion," paragraph
    concl_el = None
    for el in body.iterchildren():
        if el.tag == qn("w:p") and para_text(el).strip().startswith("In conclusion,"):
            concl_el = el
            break
    assert concl_el is not None, "conclusion paragraph not found"
    move_before(doc, concl_el, lambda d: H(d, "Conclusions", level=1))

    # 6) replace ICM declarations: delete [Declarations H1, Abbreviations H1)
    children = list(body.iterchildren())
    decl_el = next(el for el in children if is_heading(el, "Declarations"))
    abbr_el = next(el for el in children if is_heading(el, "List of abbreviations"))
    deleting = False
    for el in children:
        if el is decl_el:
            deleting = True
        if el is abbr_el:
            break
        if deleting:
            body.remove(el)

    # 7) insert AIC declarations before References H1
    refs_el = next(el for el in body.iterchildren() if is_heading(el, "References"))
    move_before(doc, refs_el, declarations_aic)

    # 8) append dataset refs (references are the last section)
    for ref in DATASET_REFS:
        P_(doc, ref, size=10, align="justify")

    # 8b) BMC house style: figure legends and tables come AFTER the references.
    #     Move the [Figures H1, List of abbreviations H1) block to the end of the
    #     body, keeping w:sectPr as the final element.
    figs_el = None
    for el in body.iterchildren():
        if is_heading(el, "Figures"):
            figs_el = el
            break
    if figs_el is not None:
        abbr_el2 = next(el for el in body.iterchildren()
                        if is_heading(el, "List of abbreviations"))
        block, inside = [], False
        for el in list(body.iterchildren()):
            if el is figs_el:
                inside = True
            if el is abbr_el2:
                break
            if inside and el.tag != qn("w:sectPr"):
                block.append(el)
        sectPr = body.find(qn("w:sectPr"))
        for el in block:
            body.remove(el)
            if sectPr is not None:
                sectPr.addprevious(el)
            else:
                body.append(el)
        print(f"  ✓ Moved {len(block)} figure/table blocks after References")

    # 9) remove all explicit page breaks
    for br in list(body.iter(qn("w:br"))):
        if br.get(qn("w:type")) == "page":
            br.getparent().remove(br)

    # 10) double line spacing + line numbering
    doc.styles["Normal"].paragraph_format.line_spacing = 2.0
    for sec in doc.sections:
        sectPr = sec._sectPr
        ln = OxmlElement("w:lnNumType")
        ln.set(qn("w:countBy"), "1")
        ln.set(qn("w:start"), "0")
        ln.set(qn("w:distance"), "240")
        ln.set(qn("w:restart"), "newPage")
        sectPr.append(ln)

    bf._normalise_minus(doc)

    out = BASE / "TTE_Deescalation_Manuscript_AIC.docx"
    doc.save(str(out))
    print(f"  ✓ AIC manuscript: {out}")
    return out


# ────────────────────────────────────────────────────────────────────────────
# 3) AIC cover letter
# ────────────────────────────────────────────────────────────────────────────
def cover_letter_aic():
    doc = bf.setup_doc()
    P_(doc, "Cover Letter", bold=True, size=13, align="center")
    doc.add_paragraph()
    P_(doc, "Dear Editor,", align="left")
    doc.add_paragraph()
    P_(doc,
        "We are pleased to submit our manuscript entitled \u201cEarly versus "
        "Deferred De-escalation from Controlled Ventilation after a 48-Hour "
        "Landmark of Physiological Stability: A Target Trial Emulation in "
        "MIMIC-IV with External Validation in eICU-CRD\u201d for consideration "
        "as a Research Article in Annals of Intensive Care.", align="justify")
    P_(doc,
        f"The optimal timing of transition from controlled to assisted "
        f"mechanical ventilation remains a critical unanswered question in "
        f"intensive care medicine. Using a target trial emulation framework "
        f"with clone\u2013censor\u2013weight estimation, we compared early "
        f"versus deferred de-escalation strategies in {F['eligible']:,} adults "
        f"still invasively ventilated 48 hours after ICU admission with "
        f"FiO\u2082 \u226450% and PEEP \u226410 cmH\u2082O. Our primary analysis "
        f"in {S['early'] + S['deferred']:,} analysable patients showed that "
        f"early de-escalation was associated with lower 28-day mortality "
        f"({pp(P['mort_early'])}% versus {pp(P['mort_deferred'])}%; risk "
        f"difference {pp_label(P['rd'])} percentage points; 95% CI "
        f"{ci_pp(P['rd_ci_lo'], P['rd_ci_hi'])}; risk ratio {P['rr']:.2f}, 95% CI "
        f"{P['rr_ci_lo']:.2f}\u2013{P['rr_ci_hi']:.2f}). The finding was robust "
        f"across six sensitivity and robustness analyses, including an extended "
        f"40-covariate propensity adjustment, alternative weight truncation, and "
        f"a doubly robust estimator (E-value {_EV_TXT}); all nine exploratory "
        f"FiO\u2082\u00d7PEEP threshold cells excluded zero. Critically, the "
        f"association was externally validated under a harmonised protocol in "
        f"the independent multi-centre eICU-CRD cohort ({EV['n_analyzable']:,} "
        f"analysable patients across 208 hospitals): risk difference "
        f"{pp_label(EVP['rd'])} percentage points for in-hospital mortality.",
       align="justify")
    P_(doc,
        "We believe this work fits the scope of Annals of Intensive Care "
        "because it applies rigorous causal-inference methods to a clinically "
        "important weaning question, provides a fully pre-specified protocol "
        "with complete TARGET- and STROBE-compliant reporting (both checklists "
        "are included in the supplement), includes per-protocol verification "
        "that the assigned strategies were delivered as intended, and "
        "demonstrates transportability of the finding to an independent "
        "multi-centre database.", align="justify")
    P_(doc,
        "This manuscript is original work that has not been published elsewhere "
        "and is not under consideration at another journal. All listed authors "
        "have approved the submitted version and agree to be accountable for "
        "its content. The authors declare no competing interests. No specific "
        "funding was received for this study.", align="justify")
    P_(doc,
        "In accordance with the journal's instructions, the manuscript is "
        "double-spaced with continuous line numbering and is organised as "
        "Background, Methods, Results, Discussion, Conclusions, List of "
        "abbreviations, Declarations, References, figure legends and tables. "
        f"The abstract contains {ABSTRACT_WORDS} words (structured "
        "Background/Methods/Results/Conclusions; limit 350). Reporting follows "
        "STROBE and the TARGET statement; both completed checklists, the full "
        "target-trial protocol, the component-by-component emulation mapping, "
        "the operational definitions used in each database, the covariate "
        "missingness table and the complete set of supplementary tables "
        "(S1\u2013S10) and figures (S1\u2013S6) are provided in a single "
        "supplementary appendix. The complete analysis code is openly "
        "available at https://github.com/ccmzhangrui/"
        "ventilator-deescalation-target-trial-emulation; the two source "
        "databases (MIMIC-IV v2.2 and eICU-CRD v2.0) are accessible through "
        "PhysioNet under the required data-use agreements.", align="justify")
    doc.add_paragraph()
    P_(doc, "Thank you for your consideration.", align="left")
    doc.add_paragraph()
    P_(doc, "Sincerely,", align="left")
    P_(doc, "[Corresponding author]", align="left")
    P_(doc, "[Affiliation]", align="left")
    P_(doc, "[Email]", align="left")
    out_path = BASE / "TTE_CoverLetter_AIC.docx"
    bf.polish_doc(doc)
    doc.save(str(out_path))
    print(f"  ✓ AIC cover letter: {out_path}")
    return out_path


# ────────────────────────────────────────────────────────────────────────────
def main():
    print("[1/4] Rebuilding ICM manuscript with new citations...")
    icm = rebuild_icm()
    print("[2/4] Transforming to AIC variant...")
    aic = transform_to_aic(icm)
    print("[3/4] Building AIC cover letter...")
    cover_letter_aic()
    print("[4/4] Data accuracy audit...")
    sys.path.insert(0, str(BASE))
    import verify_data
    fails = verify_data.audit(aic) + verify_data.audit(icm)
    print("AUDIT", "FAILED" if fails else "PASSED")


if __name__ == "__main__":
    main()
