# Numeric hard-code audit — corrected taxonomy (2026-10 patch)

Scope: `src/paper2/*.py`, `tests/*.py`, `review/*.md`, root docs, manuscript
generator, figure/table builders. Method: repository-wide numeric-literal
scan cross-checked against the canonical layer
(`results/manuscript_values.csv`, 50 claims) and the frozen protocol/SAP.

## Taxonomy (corrected)

### A — DESIGN / PROTOCOL CONSTANTS (fixed before outcome observation)

Legitimately hard-coded when identified as design constants. In this repo
they live in the frozen protocol/SAP and the code that enforces them:

- 3 blind slots per paper (`core.py` slot validation;
  `PROTOCOL-FREEZE-2026-09-25`)
- prospectively planned main sample size N = 100
- sampling seed (recorded in `results/frame_manifest.json`)
- per-slot tool-call / token / wall-clock / CPU-memory ceilings
- frozen alpha / tolerance thresholds; strict–majority–permissive
  success definitions

These are **inputs**, not results; they do not appear in
`manuscript_values.csv`.

### B — SOURCE-DERIVED OBSERVED VALUES (corpus audit)

Observed properties of the deposited corpus, **not** design constants.
Canonical rows carry `observed_source_audit` status and trace to the
source-corpus audit (11 claims): source_records 10,000; unique_pmids 9,935;
duplicated 64 (double/triple occurrence, within-field excess); EPJ fields 7;
unique PMID–field pairs; multi-field PMIDs; missing-DOI records.

### C — ANALYSIS-DERIVED EMPIRICAL VALUES (pipeline outputs)

Everything computed downstream: 461 route-eligible; 100 actually sampled;
10 attempted; 90 non-executed; 30 realized slots; non-execution-state counts;
L1–L5 run levels; stop reasons; failure codes; frozen 0/10; provisional
adjusted 1/10; 10/100 reachability; end-to-end yields; all confidence
intervals; B 81/7/2; leakage 1/7/2 of 10; C 25/25; E 130/5/3; policy-envelope
bounds; weighted ranges; Table 10/11 counts; input-state descriptives
(frozen-record descriptive, not a drift measurement); evidence refs 698/0;
D pending counter (0, `pending_external_not_performed`).

**Rule: must be generated from canonical outputs — never hand-copied.**
Enforced: `documents.py` renders all such numbers from `manuscript_values.csv`
lookups; `verification_qc.py` derives expected substrings from the same
layer, so a broken derivation fails the build.

### D — EXTERNAL-SOURCE VALUES

Target values extracted from source publications and external bibliographic
data (published coefficients, counts, thresholds being reconstructed;
verified-references metadata). These live in the adjudication/specification
records and `data/verified_references.csv`, are linked to publication
evidence, and are **not** study-derived quantities. No manuscript claim is
generated from them as if it were a study result.

### E — IMPLEMENTATION / FORMATTING CONSTANTS

Figure geometry, font sizes, margins, slide dimensions, schema versions —
non-scientific constants in `documents.py` (e.g. `Inches(0.7)`,
`figsize=(13.333, 7.5)`, fontsize values). Audited; all are layout only.

## Dual-role values (both roles documented)

- **N = 100**: planned sample size = A (design, protocol-frozen); "100
  papers were actually sampled" (`main_cohort_selected`) = C (realized).
- **3 slots**: slots-per-paper = A (design); "30 slots executed"
  (`main_blind_slots`) = C (realized).
- **10 / 90**: realized counts only — C (no design role).

## Re-scan results (post-taxonomy)

- No empirical result literal found in generator prose, tables, captions or
  figure labels — all render from canonical rows.
- Remaining hard-coded numbers are A (frozen protocol inputs, in
  protocol/build code, documented) or E (formatting). No C/B literal
  duplicates a canonical value.
- External-source targets are confined to evidence/specification records.
- Nothing was reclassified in `manuscript_values.csv`; the single source of
  truth is untouched — this patch corrected audit **taxonomy**, not results.

Machine-readable: `results/NUMERIC_HARDCODE_FINDINGS.csv` (taxonomy column),
`results/NUMERIC_PROVENANCE.csv`, `results/MANUSCRIPT_NUMERIC_TRACEABILITY.csv`.
