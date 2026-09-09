#!/usr/bin/env python3
"""Data accuracy audit: extract the AIC/ICM manuscript text and verify every
key number against the locked analysis JSONs (analysis_summary.json,
eicu_validation.json, robustness_results.json). Any mismatch -> FAIL list."""
import json, re, sys
from pathlib import Path
import docx

BASE = Path("/Users/zhangrui/Desktop/tte/TTE_Deescalation_v2")
RES = BASE / "results"

A = json.load(open(RES / "analysis_summary.json"))
P, F, S = A["primary"], A["flow"], A["strategy_counts"]
EV = json.load(open(RES / "eicu_validation.json"))
EVP, EVS = EV["primary_hospital_mortality"], EV["strategy_counts"]
RB = json.load(open(RES / "robustness_results.json"))


def docx_text(path):
    d = docx.Document(str(path))
    parts = [p.text for p in d.paragraphs]
    for t in d.tables:
        for r in t.rows:
            for c in r.cells:
                parts.append(c.text)
    return "\n".join(parts)


def pp(x):
    return f"{x*100:.1f}"


# (label, expected string derived from JSON, independent hard-locked constant)
CHECKS = [
    ("MIMIC total records",   f"{F['total_records']:,}", "51,992"),
    ("MIMIC eligible",        f"{F['eligible']:,}", "2,267"),
    ("analysable total",      f"{S['early'] + S['deferred']:,}", "1,902"),
    ("early n",               f"{S['early']}", "776"),
    ("deferred n",            f"{S['deferred']:,}", "1,126"),
    ("mort early %",          pp(P["mort_early"]), "9.2"),
    ("mort deferred %",       pp(P["mort_deferred"]), "13.7"),
    ("RD pp",                 f"{P['rd']*100:+.1f}", "-4.5"),
    ("RD CI lo",              f"{P['rd_ci_lo']*100:.1f}", "-7.3"),
    ("RD CI hi",              f"{P['rd_ci_hi']*100:.1f}", "-1.7"),
    ("RR",                    f"{P['rr']:.2f}", "0.67"),
    ("RR CI",                 f"{P['rr_ci_lo']:.2f}\u2013{P['rr_ci_hi']:.2f}", "0.50\u20130.86"),
    ("RMST diff",             f"{P['rmst_diff']:.2f}", "0.73"),
    ("VFD diff",              f"{P['vfd_diff']:.2f}", "2.12"),
    ("eICU analysable",       f"{EV['n_analyzable']:,}", "10,957"),
    ("eICU mort early %",     pp(EVP["mort_early"]), "9.2"),
    ("eICU mort deferred %",  pp(EVP["mort_deferred"]), "21.2"),
    ("eICU RD pp",            f"{EVP['rd']*100:+.1f}", "-12.0"),
    ("eICU RD CI",            f"{EVP['rd_ci_lo']*100:.1f} to {EVP['rd_ci_hi']*100:.1f}", "-13.3 to -10.7"),
    ("eICU RR",               f"{EVP['rr']:.2f}", "0.43"),
]


def audit(path):
    print(f"\n=== auditing {path.name} ===")
    txt = docx_text(path)
    fails = []
    # 1) JSON-derived values match hard-locked constants
    for label, derived, locked in CHECKS:
        if derived != locked:
            fails.append(f"LOCK MISMATCH {label}: json={derived!r} vs locked={locked!r}")
    # 2) key strings present in manuscript text
    for label, derived, _ in CHECKS:
        probe = derived.lstrip("+")
        if probe not in txt and derived not in txt:
            fails.append(f"MISSING IN TEXT {label}: {derived!r} not found")
    # 3) stale-value sentinels that must NOT appear
    for bad in ["13.9%", "319 eligible", "432 analysable", "+0.1 pp", "all-cause mortality"]:
        if bad in txt:
            fails.append(f"STALE/BANNED STRING present: {bad!r}")
    # 4) AIC structural checks (only for AIC file)
    if "AIC" in path.name:
        if "Take-home message" in txt:
            fails.append("Take-home box still present in AIC build")
        if "Research in context" in txt:
            fails.append("Research-in-context box still present in AIC build")
        for req in ["Ethics approval and consent to participate",
                    "Consent for publication", "Availability of data and materials",
                    "Competing interests", "Funding", "Authors' contributions",
                    "Acknowledgements", "List of abbreviations", "Conclusions",
                    "Background", "Keywords"]:
            if req not in txt:
                fails.append(f"AIC required section missing: {req!r}")
        for ds in ["10.13026/6mm1-ek67", "10.13026/C2WM1R"]:
            if ds not in txt:
                fails.append(f"dataset DOI missing: {ds}")
    # 5) new methodology citations (both builds)
    for cite in ["Gilding", "Trials terminated early"]:
        if cite not in txt:
            fails.append(f"methodology citation missing: {cite!r}")
    if fails:
        print("FAILURES:")
        for f_ in fails:
            print("  ✗", f_)
    else:
        print("ALL CHECKS PASSED ✓")
    return fails


if __name__ == "__main__":
    total = []
    for name in sys.argv[1:] or ["TTE_Deescalation_Manuscript_AIC.docx",
                                 "TTE_Deescalation_Manuscript_FINAL.docx"]:
        p = BASE / name
        if p.exists():
            total += audit(p)
    sys.exit(1 if total else 0)
