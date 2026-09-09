# Ventilator de-escalation after a 48-hour landmark: target trial emulation

Reproducible analysis code for the study

> **Early versus deferred de-escalation from controlled ventilation after a 48-hour
> landmark of physiological stability: a target trial emulation in MIMIC-IV with
> external validation in eICU-CRD**

## Design in one paragraph

We emulate a target trial comparing two ventilatory strategies over a **24-hour grace
window** that begins at a **48-hour landmark** after ICU admission (time zero):

| Protocol element | Specification |
|---|---|
| Eligibility | Adults (≥18 y) invasively ventilated >24 h, alive and still ventilated at 48 h, with a classifiable Day-2 mode, Day-2 FiO₂ ≤50% and PEEP ≤10 cmH₂O |
| Strategies | **Early** = step-down in support level by Day 3 (controlled → assisted/spontaneous, or successful extubation ≤72 h). **Deferred** = support level maintained or increased |
| Assignment | Clone–censor–weight (CCW); each patient cloned into both arms at time zero and censored when the observed data become inconsistent with the assigned strategy |
| Follow-up | Time zero = 48-hour landmark for eligibility, assignment and follow-up (no immortal time by design) |
| Outcome (primary) | 28-day in-hospital mortality (MIMIC-IV) / in-hospital mortality (eICU-CRD) |
| Estimand | Per-protocol effect, expressed as risk difference and risk ratio |
| Estimator | Stabilised inverse-probability-of-censoring weights (IPCW), truncated at the 1st–99th percentile; 21-covariate propensity model (40 in the extended analysis) |
| Inference | Patient-level bootstrap, 2,000 replicates (primary) / 800 (sensitivity); the propensity model is refit and weights re-truncated in each replicate |

Death is treated as an **outcome event**, never as a censoring event. Deaths occurring
inside the grace window form a separate, pre-specified exclusion category and are
reported in the flow diagrams.

## Repository layout

```
run_analysis.py          MIMIC-IV derivation pipeline (cohort -> weights -> effects -> CSV/JSON)
eicu_extract.py          eICU-CRD v2.0 cohort extraction (6 stages -> parquet cache)
eicu_analysis.py         External-validation pipeline (harmonised protocol)
compute_unweighted.py    Crude (unweighted) event counts + unweighted log-rank p
robustness_package.py    SA4 extended covariates, SA5 alt. truncation, AIPW, E-value, subgroups
ventilation_trajectory.py Per-protocol delivery verification (Figure S4 inputs)
matching_sensitivity.py  Propensity-score matching sensitivity analysis
make_flow_figures.py     Figures 1 and S5 (flow diagrams)
make_eicu_figures.py     Figure 4 (external validation)
make_design_schematic.py Figure S6 (CCW design schematic)
fix_figures.py           Figures 2, 3, S1, S2, S3
build_final.py           Builds manuscript / supplement / cover letter / slides (ICM style)
build_aic.py             Transforms the manuscript to Annals of Intensive Care format
verify_data.py           Audits every number in the built DOCX against results/*.json
results/                 Locked analysis outputs (JSON + CSV) — the single source of truth
figures/                 Publication-resolution PNGs as embedded in the manuscript
```

**Every number in the manuscript is interpolated from `results/` at build time.**
`verify_data.py` re-reads the built `.docx`, extracts all reported figures and asserts
that each matches the locked JSON/CSV. No result is typed by hand.

## Data access

The two databases are **not** redistributed here. Access requires credentialing and a
data-use agreement:

* MIMIC-IV v2.2 — <https://physionet.org/content/mimiciv/2.2/> (doi:10.13026/6mm1-ek67)
* eICU-CRD v2.0 — <https://physionet.org/content/eicu-crd/2.0/> (doi:10.13026/C2WM1R)

`run_analysis.py` expects the pre-extracted MIMIC-IV research sheet at the path set in
`EXCEL_PATH` (default `/Users/zhangrui/Documents/sofa2.0/副本minmic数据.xlsx`); edit this
constant to point at your own approved extract. `eicu_extract.py` expects the 25 released
eICU-CRD CSV tables in `EICU_DIR`.

Intermediate row-level artefacts (`results/*_analyzable.parquet`) are **excluded** from
this repository by design: they are derived from credentialed data and are regenerated
locally by the pipelines above. Only aggregate outputs are published.

## Environment

```bash
pip install -r requirements.txt
```

Python 3.13. Random seed is fixed (`rng_seed = 20260816`), so the bootstrap results are
deterministic.

## Reproducing the analysis

```bash
python run_analysis.py          # MIMIC-IV derivation -> results/*.json|csv
python compute_unweighted.py    # crude counts + unweighted log-rank p
python robustness_package.py    # SA4, SA5, AIPW, E-value, subgroups
python ventilation_trajectory.py
python eicu_extract.py          # eICU-CRD extract (requires local CSVs)
python eicu_analysis.py         # external validation
python make_flow_figures.py && python make_eicu_figures.py && \
       python fix_figures.py && python make_design_schematic.py
python build_aic.py             # manuscript + cover letter + number audit
```

`build_aic.py` prints `[PASS]`/`[FAIL]` for every audited number and raises on mismatch.

## Reporting

Reporting follows STROBE and the TARGET statement (21-item checklist reproduced in the
Supplement). The study was not registered; the full target-trial protocol is specified in
the Supplement and was fixed before the analytic dataset was finalised.

## Licence

Analysis code: MIT. Database content remains governed by the respective PhysioNet
data-use agreements.
