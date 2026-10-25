# Revision changelog — manuscript package after input-state identifiability

Scope: targeted revision of the generated manuscript package (`src/paper2/documents.py`,
`src/paper2/input_state.py`, `src/paper2/verification_qc.py`). No empirical result,
denominator, sealed outcome, adjudication record or human-validation status changed.

## Unchanged frozen records (verified by `make check` / `review/VERIFICATION_ADJUSTED_QC.md`)

- sealed blind outcome SHA-256: `350a0e7dcedbb74ccad2472c642318c87cc7d272f1b6c88186a1419dc858ea5a`
- primary mechanical adjudication SHA-256: `6fd42b8964e055606cbd24ebddb44f42fa09147fa80bfa2d5bcaa161766c5b28`
- A/B/C/E imported review tree SHA-256: `f6b374b74c48c297e4c68dcdd780a4d2c1f162e8dcfc058111c524408961f78a1`
- reconstruction slots rerun: 0; original code executed: 0; Section D cases: 0 (pending)
- Paper III analyses: none

## Manuscript text (documents.py)

- Subtitle: "Prospective Paper II study" → completed-study wording with human validation pending.
- Abstract rewritten in completed-study tense: unique-PMID frame, ACCESS refinement, pilot and
  protocol freeze, 100/461 sampling, 10 attempted, 30 sealed slots, frozen 0/10 majority, separate
  provisional AI-assisted 1/10 that does not overwrite the sealed result, attrition as primary
  finding, D not performed.
- Methods heading: "prospectively frozen (protocol text)" → "procedures as executed under the
  prospectively frozen protocol". 2.1–2.6 converted from imperative/future to past tense.
  Structure (2.1, 2.1a, 2.2, 2.3, 2.4, 2.5, 2.6) preserved.
  - 2.1: amended deposited-data criterion, 461 route-eligible papers, retained attrition layers,
    delegated G1–G5 with investigator verification pending.
  - 2.2: pilot completed (7 papers × 3 slots, procedural review PD-01, rules unchanged, no pilot
    rate), PROTOCOL-FREEZE-2026-09-25, stratified sampling of 100 from 461, ten-item target freeze.
  - 2.3: "Use three clean, fixed slots…" → completed description with the frozen per-slot ceiling;
    "No qualified harness currently exists." removed and replaced by the calibrated statement that
    a functioning harness was implemented and used, listing controls that passed qualification
    probes, controls not formally qualified, and that prior model-training exposure cannot be
    certified absent.
  - 2.5: "The proposed weighted interval is not yet implemented" / "draft SAP" removed; analyses
    described as performed; no multivariable model; frozen SAP.
  - 2.6: sealing, mechanical adjudication, descriptive reveal and A/B/C/E import described as
    completed; Section D described as prespecified, supported by ten cases from frozen records,
    and not performed. "Prepare approximately 15 … cases" removed.
- Section 3 heading: "Observed preparation findings" → "Source corpus audit".
- Section 4 heading now reads "sealed delegated mechanical results (frozen)".
- Section 5 heading now reads "(A/B/C/E; provisional, separate from Section 4; D excluded)" with
  a lead paragraph: Section 4 frozen; Section 5 provisional pending investigator approval; does
  not overwrite Section 4; no reclassified non-executed paper was executed; reveal never changed
  a blind score.
- Discussion: added "Data/code availability requirements address only selected stages of the
  reproducibility chain; reconstructability also depends on explicit specification, stable input
  identity, executable environments and target-linked evaluation rules." and a statement that the
  framework is conceptual, not an empirically validated universal model. Existing points retained
  (attrition primary, non-executions ≠ failures, conditional success, resource ceilings, leakage
  false positives, agent failure ≠ human impossibility, descriptive reveal, input-state
  identifiability, no drift quantification).
- Statements: data/code availability statement updated (public repository now contains frozen
  protocols, sealed records and layers; inputs and full text retained internally, not
  redistributed); AI-assistance statement now describes the executed roles.
- References heading: "References used in this preparation draft" → "References".

## Tables

- Table 10 (`input_state.py`): code availability row kept in ACCESS, relabelled "ACCESS (artifact
  access)" with purpose "artifact reproducibility only, deliberately excluded from the Paper II
  blind intervention and used solely as a descriptive characteristic and post-seal reveal
  resource".
- Table 11: values re-derived from `paper_status.json` (100 papers; 28 external-reference
  barrier), 100 delegated specifications (119 reference entries in 51 papers; 19 unspecified in
  11 papers; 100/100 accession; 89/100 file list), 10 target freezes (21 inputs) and the retrieval
  ledger (21/21 SHA-256 and timestamps; 7/21 announced checksums, 7 matched). All values
  reconciled; version-identifier, schema/release and historical-version fields remain
  "not recorded"/"not assessed". Two item labels clarified ("recorded by the specifier from
  publication text"; "not specified in publication text"). No value was invented or rescored.

## References

- The 11 input-state references (FAIR, FORCE11, RDA dynamic-data citation, Rauber, Pröll & Rauber,
  Klump, Pasquier, SWHID, Sandve, Stodden, Zhao) remain in `data/verified_references.csv` with the
  DOIs cross-checked in `review/INPUT_STATE_CITATION_VERIFICATION.md`; QC re-verifies each DOI in
  manuscript and ledger. Terms remain provisional; no novelty claim.

## Figure 1

- Unchanged: ACCESS → RECONSTRUCT → EXECUTE → REPRODUCE → ROBUST; ACCESS sub-questions;
  Paper II highlight on RECONSTRUCT; ROBUST labelled future Paper III; Wet/Dry conceptual;
  drift note. Reviewed at journal width; no change required.

## QC (`verification_qc.py`)

New checks: no stale preparation wording; harness statement calibrated; Abstract completed-tense
with frozen/provisional separation and D pending; Section 4/5 labels; Table 10 code-availability
caveat; availability-chain sentence; no overclaim (guarantee, final adjudication, Paper III
analyses); Methods structure; changelog present with frozen hashes.

## Status

DRAFT — NOT SUBMISSION READY. Investigator verification/sign-off, human adjudication and Human
Validation D remain pending; `make submission-check` remains non-zero by design.

---

# Revision 2026-10 — hostile review / hard-code elimination pass

No empirical result, denominator, sealed outcome, adjudication record, protocol,
SAP, tolerance rule, sampling decision, or human-validation status changed.
No blind slot rerun; no reveal-code execution; no Paper III work.

## Code / provenance
- `documents.py`: all empirical literals in prose/captions/tables replaced by
  canonical `values` lookups; Section 7.1 software-redistribution claim softened
  per H5; `scripts/` added to shipped file list.
- `input_state.py`: `input_state_stats()` + `input_state_values()` — input-state
  counts now canonical frozen-descriptive rows (not empirical drift claims).
- `verification_layer.py`: 3 new canonical rows (pilot cohort, gate-failed
  cohort, external-reference barrier) marked out-of-denominator.
- `verification_qc.py`: abstract assertions derive expected substrings from
  `manuscript_values.csv` instead of duplicating literal results.
- `build.py`: emits the new canonical rows.
- `Makefile` + `scripts/fetch_vetted_wheels.sh`: hash-verified fetch of the
  pinned vetted wheels; fixes the one failing test on clean checkouts.

## New audit artifacts
- review/HARDCODE_AUDIT.md, results/NUMERIC_HARDCODE_FINDINGS.csv,
  results/NUMERIC_PROVENANCE.csv, results/MANUSCRIPT_NUMERIC_TRACEABILITY.csv
- review/HOSTILE_REVIEW.md (C1–C10), review/HOSTILE_REVIEW_ACTION_LOG.md,
  review/FINAL_ADVERSARIAL_DECISION.md (SUBMISSION-READY EXCEPT D)
- review/DATA_TEXT_FIGURE_CONSISTENCY.md, review/DOCUMENTATION_STALENESS_AUDIT.md,
  review/PACKAGE_SELF_CONSISTENCY.md, review/REFERENCE_AUDIT.md,
  results/CITATION_CLAIM_MAP.csv, review/TEST_REPORT.md, FINAL_HANDOFF.md
- README.md: dated "Current state" banner; review/EVIDENCE.md and
  review/REVIEWER_AUDIT.md: SUPERSEDED banners (provenance retained).

---

# Revision 2026-10 (pass 2) — FINAL pre-D finalization

No empirical result, sealed outcome, protocol, SAP, tolerance, sampling, or
adjudication record changed. No slot rerun; no reveal-code execution; no
human validation fabricated; no sign-off recorded.

## Added
- `investigator_signoff/`: A(30)/B(90)/C(25)/E(130) sign-off CSVs, per-layer
  summaries, DOCX attestation form, XLSX checklist, README — decision fields
  intentionally empty (generated by `scripts/make_investigator_signoff.py`).
- `scripts/make_final_pred_package.py`: emits
  `manuscript/Preparation_draft_EN_inline_FINAL_PRE_D.{docx,pdf}` (PDF via
  LibreOffice), `Supplement_FINAL_PRE_D.docx`, and
  `dist/PaperII_FINAL_PRE_D_PACKAGE.zip`.
- `docs/archive/README_pre_execution.md` (verbatim archive).
- review/: FINAL_HOSTILE_REVIEW.md (F1–F16), FINAL_HOSTILE_REVIEW_ACTION_LOG.md,
  DOCUMENTATION_FINALIZATION_AUDIT.md, PUBLIC_PACKAGE_CHECK.md,
  PRIVATE_EVIDENCE_CHECK.md, FINAL_TEST_REPORT.md.

## Changed
- README.md: rewritten to current canonical state only; stale pre-execution
  prose removed from root (archived). ACCESS decomposition kept for QC.
- FINAL_ADVERSARIAL_DECISION.md: classification
  **SUBMISSION-READY EXCEPT D AND INVESTIGATOR SIGN-OFF** with remaining-work table.
- documents.py zip list now ships `docs/` and `investigator_signoff/`.

## Verified
- Canonical recomputation: all headline numbers match manuscript_values.csv
  and source JSONs (10,000/9,935/64/7/461/100/10+30/90; frozen 0/10 vs
  adjusted 1/10; B 81/7/2; C 25/25; E 130/5/3; leakage 10→1/7/2; refs 698/0).
- `make check`: 146 passed. `make submission-check`: non-zero by design.
- Rendering QC: 10-page PDF, title/abstract/tables legible, DRAFT banner on
  every page.

---

# Patch 2026-10 — hard-code audit taxonomy correction (documentation only)

The numeric hard-code audit now distinguishes: A design/protocol constants;
B source-derived observed values; C analysis-derived empirical values;
D external-source values; E implementation/formatting constants. Dual-role
values (planned N=100 vs 100 sampled; 3 slots/paper vs 30 executed) documented
in both roles. Updated: HARDCODE_AUDIT.md, NUMERIC_HARDCODE_FINDINGS.csv
(taxonomy column), NUMERIC_PROVENANCE.csv, MANUSCRIPT_NUMERIC_TRACEABILITY.csv
(taxonomy+role columns), DATA_TEXT_FIGURE_CONSISTENCY.md.

**No scientific result changed; no sealed outcome changed; no manuscript
estimate changed; this patch corrected audit taxonomy only.**


## FINAL NON-HUMAN QC PATCH (this commit)

- Renumbered supplement/editable tables to S1–S11 in first-citation order;
  fixed the Methods-cites-Tables-10/11-before-Table-1 bug; Wilson planning
  table S1→S11 to keep the supplement sequence strict.
- Converted author–year citations to numeric [1]–[21] numbered by first
  appearance; reference list now numeric; proll_rauber_dynamic now cited [13].
- Converted all true equations in the supplement to native OMML (m:oMath):
  Ŝ = Σ_h N_h·ȳ_h, Ŝ/|E|, Ŝ/Σ_h N_h, π_h = n_h/N_h, weights, |E|/|U|
  cardinalities, k ≥ t / k + m < t, t = 1/2/3 thresholds, α = 0.05, α/(2H),
  eligibility interval, rounding-bin interval. Code/JSON/CLI/glob tokens left
  as literal text.
- Removed all formatting `*` from Word deliverables: **bold**/*italic* now
  render as real runs (including emphasis spanning code spans and wrapped
  source lines); hard-wrapped protocol lines joined into single paragraphs.
- Updated generator, verification_qc table checks, INPUT_STATE note and audit
  docs to the S-numbering. New audit artifacts:
  FIGURE_TABLE_ORDER_AUDIT.md + FIGURE_TABLE_RENUMBER_MAP.csv,
  REFERENCE_ORDER_AUDIT.md + REFERENCE_RENUMBER_MAP.csv +
  CITATION_SEQUENCE_CHECK.csv, OMML_EQUATION_AUDIT.md + OMML_EQUATION_MAP.csv,
  ASTERISK_AUDIT.md + ASTERISK_OCCURRENCES.csv, LINE_BREAK_AUDIT.md.
- No scientific result, denominator, estimate, interval, figure or table value
  changed; sealed outcomes untouched.

## FINAL NON-HUMAN COMPLETION PATCH (this commit)

- Added main-text citation of Supplementary Table S11 (Methods, precision
  planning); every supplementary object S1–S11 now cited from the main text.
- Fresh clean-clone reproducibility check from the pushed public branch:
  make setup → fetch-vetted-wheels → build → check, 146 tests passed.
- Added `make public-check` / `make private-evidence-check` targets to make
  the public/private boundary operational.
- New artifacts: PUBLIC_REPRODUCIBILITY_MATRIX.csv,
  FINAL_PUBLIC_END_TO_END_REPRODUCIBILITY_AUDIT.md.
- No scientific result, denominator, estimate, interval, figure or table value
  changed; sealed outcomes untouched.
