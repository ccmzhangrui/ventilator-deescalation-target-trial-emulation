#!/usr/bin/env python3
"""
eICU-CRD 2.0 extraction for TTE de-escalation external validation.

Stages (cached to eicu_cache/ as parquet):
  1. base      — patient.csv + apachePatientResult (first ICU stay, adults)
  2. resp      — respiratoryCharting vent settings (chunked, cohort stays only)
  3. treat     — treatment.csv ventilation mode events
  4. labs      — lab.csv key labs (chunked, cohort stays only)
  5. drugs     — infusionDrug vasopressor/sedation (cohort stays only)
  6. history   — pastHistory comorbidities

Usage: python eicu_extract.py [stage ...]   (default: all)
"""
import sys, re, json, time
from pathlib import Path
import numpy as np
import pandas as pd

EICU = Path("/Users/zhangrui/Documents/sofa2.0/eicu-collaborative-research-database-2.0")
CACHE = Path("/Users/zhangrui/Desktop/tte/TTE_Deescalation_v2/eicu_cache")
CACHE.mkdir(exist_ok=True)

VENT_LABELS = {
    "FiO2": "fio2", "PEEP": "peep", "Vent Rate": "vent_rate",
    "Pressure Support": "ps", "Pressure Control": "pc",
    "SaO2": "sao2", "Tidal Volume (set)": "tv_set",
    "Plateau Pressure": "pplat", "Total RR": "total_rr",
    "RR (patient)": "rr_patient", "Mean Airway Pressure": "map_airway",
    "Peak Insp. Pressure": "peakp",
}
# labels that prove invasive vent (exclude SaO2 / RR patient which exist on non-vent too)
VENT_PROOF = {"fio2", "peep", "vent_rate", "ps", "pc", "tv_set", "pplat",
              "total_rr", "map_airway", "peakp"}

LAB_WANT = {"lactate": "lactate", "paO2": "pao2", "pH": "ph",
            "bicarbonate": "bicarb", "creatinine": "creatinine",
            "platelets x 1000": "platelets", "WBC x 1000": "wbc",
            "total bilirubin": "tbil", "albumin": "albumin", "Hgb": "hgb",
            "-lymphs": "lymphs"}

DRUG_PATTERNS = {
    "norepi": r"norepinephrine|levophed|noradrenaline",
    "vasopressin": r"vasopressin",
    "epinephrine": r"epinephrine(?!.*local)",
    "phenylephrine": r"phenylephrine|neo-synephrine",
    "dopamine": r"dopamine",
    "propofol": r"propofol|diprivan",
    "fentanyl": r"fentanyl",
    "midazolam": r"midazolam|versed",
    "dexmedetomidine": r"dexmedetomidine|precedex",
}

HISTORY_MAP = {
    "copd": r"COPD|chronic obstructive|emphysema",
    "chf": r"congestive heart failure|CHF|heart failure",
    "diabetes": r"diabetes",
    "renal": r"renal failure|chronic kidney|CKD|dialysis",
    "malignancy": r"cancer|malignan|tumor|lymphoma|leukemia|metasta",
    "hypertension": r"hypertension",
    "cirrhosis": r"cirrhosis|liver failure|hepatic",
    "cad": r"coronary artery|myocardial infarction|CAD|angina",
}

MV_MODE_MAP = [
    # (regex on treatmentstring tail, mode)
    (r"assist controlled", "controlled"),
    (r"volume controlled", "controlled"),
    (r"pressure controlled", "controlled"),
    (r"synchronized intermittent", "controlled"),
    (r"permissive hypercapnea", "controlled"),
    (r"volume assured", "controlled"),
    (r"tidal volume", "controlled"),
    (r"pressure support", "assisted"),
]
NIV_PAT = r"non-invasive ventilation"
WEAN_PAT = r"weaning"


def log(msg):
    print(f"[{time.strftime('%H:%M:%S')}] {msg}", flush=True)


# ───────────────────────── Stage 1: base ─────────────────────────
def stage_base():
    log("stage base: patient.csv")
    p = pd.read_csv(EICU / "patient.csv", low_memory=False)
    p = p[p["unitvisitnumber"] == 1].copy()
    # age: "> 89" -> 90; numeric else
    p["age_num"] = pd.to_numeric(p["age"].replace({"> 89": "90", "": np.nan}),
                                 errors="coerce")
    p = p[p["age_num"] >= 18]
    p = p.rename(columns={"patientunitstayid": "stay_id"})
    keep = ["stay_id", "uniquepid", "gender", "age_num", "ethnicity",
            "hospitalid", "wardid", "apacheadmissiondx", "admissionheight",
            "admissionweight", "hospitaladmitsource", "unitadmitsource",
            "unittype", "unitstaytype", "unitdischargeoffset",
            "unitdischargestatus", "hospitaldischargeoffset",
            "hospitaldischargestatus", "hospitaladmitoffset"]
    p = p[keep]
    log(f"  first-stay adults: {len(p):,}")

    log("stage base: apachePatientResult (IV)")
    a = pd.read_csv(EICU / "apachePatientResult.csv")
    a = a[a["apacheversion"] == "IV"].rename(
        columns={"patientunitstayid": "stay_id"})
    a = a[["stay_id", "acutephysiologyscore", "apachescore",
           "predictedicumortality", "predictedhospitalmortality",
           "actualventdays", "unabridgedunitlos", "unabridgedhosplos"]]
    for c in ["predictedicumortality", "predictedhospitalmortality"]:
        a[c] = pd.to_numeric(a[c], errors="coerce")
    p = p.merge(a, on="stay_id", how="left")
    p.to_parquet(CACHE / "base.parquet")
    log(f"  base saved: {len(p):,} stays")
    return p


# ─────────────────────── Stage 2: respiratory ────────────────────
def stage_resp(base=None):
    if base is None:
        base = pd.read_parquet(CACHE / "base.parquet")
    stays = set(base["stay_id"])
    log("stage resp: respiratoryCharting chunked pass")
    usecols = ["patientunitstayid", "respchartoffset", "respchartvaluelabel",
               "respchartvalue"]
    out = []
    t0 = time.time()
    for i, ch in enumerate(pd.read_csv(EICU / "respiratoryCharting.csv",
                                       usecols=usecols, chunksize=2_000_000)):
        ch = ch[ch["respchartvaluelabel"].isin(VENT_LABELS.keys())]
        ch = ch[ch["patientunitstayid"].isin(stays)]
        if len(ch):
            ch["var"] = ch["respchartvaluelabel"].map(VENT_LABELS)
            ch["value"] = pd.to_numeric(
                ch["respchartvalue"].astype(str).str.extract(
                    r"(-?\d+\.?\d*)")[0], errors="coerce")
            ch = ch.dropna(subset=["value"])
            out.append(ch[["patientunitstayid", "respchartoffset",
                           "var", "value"]])
        if i % 3 == 0:
            log(f"  chunk {i}, kept {sum(len(x) for x in out):,} rows, "
                f"{time.time()-t0:.0f}s")
    resp = pd.concat(out, ignore_index=True).rename(
        columns={"patientunitstayid": "stay_id",
                 "respchartoffset": "offset_min"})
    resp.to_parquet(CACHE / "resp.parquet")
    log(f"  resp saved: {len(resp):,} rows, {resp['stay_id'].nunique():,} stays")
    return resp


# ─────────────────────── Stage 3: treatment ──────────────────────
def stage_treat(base=None):
    if base is None:
        base = pd.read_parquet(CACHE / "base.parquet")
    stays = set(base["stay_id"])
    log("stage treat: treatment.csv")
    t = pd.read_csv(EICU / "treatment.csv",
                    usecols=["patientunitstayid", "treatmentoffset",
                             "treatmentstring"])
    t = t[t["patientunitstayid"].isin(stays)].copy()
    ts = t["treatmentstring"].astype(str).str.lower()
    t["event"] = None
    t.loc[ts.str.contains(WEAN_PAT, na=False), "event"] = "weaning"
    t.loc[ts.str.contains(r"cpap/peep therapy", na=False), "event"] = "assisted"
    is_mv = ts.str.contains("mechanical ventilation", na=False)
    is_niv = ts.str.contains(NIV_PAT, na=False)
    t.loc[is_mv & ~is_niv, "event"] = "mv_generic"
    for pat, m in MV_MODE_MAP:
        sel = is_mv & ~is_niv & ts.str.contains(pat, na=False)
        t.loc[sel, "event"] = m
    t.loc[is_niv, "event"] = "niv"
    ev = (t.dropna(subset=["event"])
            .rename(columns={"patientunitstayid": "stay_id",
                             "treatmentoffset": "offset_min"})
            [["stay_id", "offset_min", "event"]])
    ev.to_parquet(CACHE / "treat.parquet")
    log(f"  treat saved: {len(ev):,} rows, {ev['stay_id'].nunique():,} stays; "
        f"events: {ev['event'].value_counts().to_dict()}")
    return ev


# ─────────────────────── Stage 4: labs ───────────────────────────
def stage_labs(base=None):
    if base is None:
        base = pd.read_parquet(CACHE / "base.parquet")
    stays = set(base["stay_id"])
    log("stage labs: lab.csv chunked pass")
    usecols = ["patientunitstayid", "labresultoffset", "labname", "labresult"]
    out = []
    t0 = time.time()
    for i, ch in enumerate(pd.read_csv(EICU / "lab.csv", usecols=usecols,
                                       chunksize=3_000_000)):
        ch = ch[ch["labname"].isin(LAB_WANT.keys())]
        ch = ch[ch["patientunitstayid"].isin(stays)]
        if len(ch):
            ch["var"] = ch["labname"].map(LAB_WANT)
            ch["value"] = pd.to_numeric(ch["labresult"], errors="coerce")
            ch = ch.dropna(subset=["value"])
            out.append(ch[["patientunitstayid", "labresultoffset",
                           "var", "value"]])
        if i % 5 == 0:
            log(f"  chunk {i}, kept {sum(len(x) for x in out):,}, "
                f"{time.time()-t0:.0f}s")
    labs = pd.concat(out, ignore_index=True).rename(
        columns={"patientunitstayid": "stay_id",
                 "labresultoffset": "offset_min"})
    labs.to_parquet(CACHE / "labs.parquet")
    log(f"  labs saved: {len(labs):,} rows")
    return labs


# ─────────────────────── Stage 5: drugs ──────────────────────────
def stage_drugs(base=None):
    if base is None:
        base = pd.read_parquet(CACHE / "base.parquet")
    stays = set(base["stay_id"])
    log("stage drugs: infusionDrug.csv")
    d = pd.read_csv(EICU / "infusionDrug.csv",
                    usecols=["patientunitstayid", "infusionoffset", "drugname"])
    d = d[d["patientunitstayid"].isin(stays)].copy()
    dn = d["drugname"].astype(str).str.lower()
    d["drug"] = None
    for k, pat in DRUG_PATTERNS.items():
        d.loc[d["drug"].isna() & dn.str.contains(pat, regex=True, na=False),
              "drug"] = k
    dr = (d.dropna(subset=["drug"])
            .rename(columns={"patientunitstayid": "stay_id",
                             "infusionoffset": "offset_min"})
            [["stay_id", "offset_min", "drug"]])
    dr.to_parquet(CACHE / "drugs.parquet")
    log(f"  drugs saved: {len(dr):,} rows; "
        f"{dr['drug'].value_counts().to_dict()}")
    return dr


# ─────────────────────── Stage 6: history ────────────────────────
def stage_history(base=None):
    if base is None:
        base = pd.read_parquet(CACHE / "base.parquet")
    stays = set(base["stay_id"])
    log("stage history: pastHistory.csv")
    h = pd.read_csv(EICU / "pastHistory.csv",
                    usecols=["patientunitstayid", "pasthistorypath"])
    h = h[h["patientunitstayid"].isin(stays)].copy()
    hp = h["pasthistorypath"].astype(str).str.lower()
    for k, pat in HISTORY_MAP.items():
        h[k] = hp.str.contains(pat, regex=True, na=False)
    hx = (h.rename(columns={"patientunitstayid": "stay_id"})
            .groupby("stay_id")[list(HISTORY_MAP.keys())].max()
            .reset_index())
    hx.to_parquet(CACHE / "history.parquet")
    log(f"  history saved: {len(hx):,} stays")
    return hx


if __name__ == "__main__":
    stages = sys.argv[1:] or ["base", "resp", "treat", "labs", "drugs",
                              "history"]
    base = None
    for s in stages:
        if s == "base":
            base = stage_base()
        elif s == "resp":
            stage_resp(base)
        elif s == "treat":
            stage_treat(base)
        elif s == "labs":
            stage_labs(base)
        elif s == "drugs":
            stage_drugs(base)
        elif s == "history":
            stage_history(base)
    log("ALL STAGES DONE")
