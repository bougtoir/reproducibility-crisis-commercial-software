# Data → text → figure consistency verification

Method: every empirical number rendered in the manuscript (Abstract, Methods,
Sections 4–7, captions) and every figure/table cell now originates from
`results/manuscript_values.csv`. Verification:

1. `manuscript_values.csv` (50 claims) — each row carries source_id,
   source_sha256, analysis entry point, derived_file, status, scope. See
   results/NUMERIC_PROVENANCE.csv.
2. Generator (`documents.py`) reads the `values` dict; no empirical literal
   remains in prose strings (HARDCODE_AUDIT.md).
3. `verification_qc.py` derives expected substrings from the canonical layer
   and asserts they appear in the built DOCX — a broken derivation fails the
   build, it cannot drift silently.
4. Figure builders (Figure 1 funnel; tables 1–8) read the same canonical
   rows; MANUSCRIPT_NUMERIC_TRACEABILITY.csv maps claim → manuscript location.
5. Cross-checked headline numbers against frozen sources: 10,000 rows →
   9,935 unique PMIDs (64 duplicated) → 7 fields → 461 eligible → 100 sampled
   → 30 slots → 10 attempted → frozen 0/10 vs provisional adjusted 1/10;
   B 81/7/2; leakage 10→1/7/2; C 25/25; E 130/5/3; refs 698/0 unresolved;
   input-state 119/51/19/11/28/21/7. All match `main_analysis.json`,
   `paper_status.json`, and `input_state_descriptives.csv`.
6. No figure/text/caption inconsistency remains; the one prior risk (literal
   copies) is structurally removed.

Status: PASS.

## Taxonomy patch (2026-10)

Re-verified after the audit-taxonomy correction: every manuscript number
remains generated from `manuscript_values.csv`. Classification now
distinguishes design/protocol constants (A), source-derived observed values
(B), analysis-derived empirical values (C), external-source values (D), and
implementation/formatting constants (E); N=100 and 3-slots documented in
both design and realized roles. No scientific result, sealed outcome, or
manuscript estimate changed. Status: PASS.


## FINAL NON-HUMAN QC PATCH addendum (figure/table + citation renumbering)

Object numbering changed: all supplement/editable tables are now numbered
S1–S11 (mapping in review/FIGURE_TABLE_RENUMBER_MAP.csv); main text retains
Table 1 and Figure 1. All in-text citations updated to the S-scheme; numeric
citations converted to [1]–[21] by first appearance (review/REFERENCE_*
artifacts). No numeric claim, denominator, estimate, interval, figure value or
conclusion changed — the renumbering is presentational only. DATA–TEXT–FIGURE
consistency re-verified after renumbering: all canonical values unchanged.
