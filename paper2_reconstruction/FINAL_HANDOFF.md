# FINAL_HANDOFF — Paper II reconstruction study (revision 2026-10)

## State

- Branch: `devin/1790076395-paper2-reconstruction` (PR #1), repo
  `bougtoir/reproducibility-crisis-commercial-software`, dir
  `paper2_reconstruction/`.
- Status: **NOT SUBMISSION READY — SUBMISSION-READY EXCEPT D AND
  INVESTIGATOR SIGN-OFF** per `review/FINAL_ADVERSARIAL_DECISION.md`.
- Investigator sign-off package prepared at `investigator_signoff/` — 275
  proposed records (A 30 / B 90 / C 25 / E 130), decision fields empty.
- Final delivery artifacts: `manuscript/*_FINAL_PRE_D.*` and
  `dist/PaperII_FINAL_PRE_D_PACKAGE.zip` (rebuild:
  `make build && .venv/bin/python scripts/make_final_pred_package.py`).
- Frozen records unchanged: sealed blind outcome
  `350a0e7d…58ea5a`, primary mechanical adjudication `6fd42b89…65b28`,
  A/B/C/E import tree `f6b374b7…f78a1`. Zero slot reruns; zero reveal-code
  executions; no protocol/SAP/tolerance/sampling diffs.

## Canonical numbers (single source: results/manuscript_values.csv, 50 claims)

10,000 rows → 9,935 unique PMIDs (64 duplicated) → 7 EPJ fields → 461
route-eligible → 100 sampled → 30 slots → 10 attempted → frozen 0/10 vs
provisional adjusted 1/10; B 81/7/2; leakage 10→1/7/2; C 25/25; E 130/5/3;
refs 698/0 unresolved; pilot 7; gate-failed 7; external-reference barrier 28;
input-state 119/51/19/11/21/7.

## Rebuild

```
make setup && make fetch-vetted-wheels && make build && make check
```

`results/preparation_artifact_manifest.json` and
`dist/PaperII_preparation_NOT_SUBMISSION_READY.zip` regenerate.
`results/verification_adjusted/evidence_reference_resolution.csv` and
`verification_import_audit.json` require the private evidence mount; off-mount
they regenerate as `evidence_package_not_mounted` — restore committed bytes
instead of committing not-mounted output.

## Open items (all external to this package)

- Human Validation D execution, human adjudication of provisional A/B/C/E,
  investigator sign-off → then remove DRAFT banner, flip readiness, retarget
  journal.
- `make submission-check` stays non-zero until D completes.

## Taxonomy patch (documentation only)

The final numeric audit distinguishes A design/protocol constants, B
source-derived observed, C analysis-derived empirical, D external-source,
E implementation/formatting constants (dual-role N=100 / 3-slots documented).
No scientific result, sealed outcome, or manuscript estimate changed.


## Update — FINAL NON-HUMAN QC PATCH applied

Presentation-layer corrections completed (S-numbering, numeric citations,
OMML equations, asterisk removal, line-wrap joining). Classification
unchanged: SUBMISSION-READY EXCEPT D AND INVESTIGATOR SIGN-OFF. Remaining
tasks are exclusively human-dependent (sign-off, ethics determination, Human
Validation D, submission).
