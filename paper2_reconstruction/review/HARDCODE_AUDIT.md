# Numeric hard-code audit (revision pass)

Scope: `src/paper2/*.py`, `tests/*.py`, `review/*.md`, manuscript generator
(`documents.py`), figure/table builders, root docs. Method: regex sweep for
empirical literals (counts, percentages, "N/M" fractions, the frozen corpus
numbers) cross-checked against `results/manuscript_values.csv` (the canonical
numeric layer, 50 claims).

## Categories (per revision prompt section D)

- **A — design constant**: fixed by protocol/SAP (30 slots, 10 attempts, 100
  sampled, 461 route-eligible, 7 EPJ fields, 10,000 source rows). These are
  protocol parameters, not results; they are stored once in
  `results/main_analysis.json` / frozen protocol and read by the generator.
- **B — external source**: numbers from cited external work (references).
  None are computed by this package; all live in the bibliography.
- **C — derived empirical**: numbers computed by the pipeline (A/B/C/E counts,
  leakage tallies, input-state descriptives, corpus dedup counts). **Rule:
  must appear only in `manuscript_values.csv`-producing code.**
- **D — suspicious**: literal in prose/rendering code that duplicates a
  derived value.
- **E — documentation example**: illustrative numbers in docs, not claims.

## Findings and resolutions

Before this pass, `documents.py` (manuscript generator) contained category-D
literals in prose strings — e.g. "100", "461", "Only 10", "30 sealed slots",
"81", "7", "2", "25/25", "130", "119", "51", "19", "11", "28", "21", "7"
rendered as fixed text in the Abstract, Methods 2.1/2.1a/2.2/2.6, Section 5
text, Discussion, and captions, plus literal denominators in table builders.
All were replaced by lookups into the canonical `values` dict built from
`manuscript_values.csv` (see `NUMERIC_HARDCODE_FINDINGS.csv`).

`verification_qc.py` asserted literal substrings ("100 papers were sampled
from 461", "proposes one", ...), i.e. the QC itself duplicated derived
numbers; it now derives expected substrings from `manuscript_values.csv`,
so the QC validates provenance rather than a second copy of the result.

`verification_layer.py` had no separate canonical rows for the pilot cohort,
the accessibility-gate failures, and the external-reference barrier count;
three rows were added so every number rendered in Methods 2.1a/Section 5
traces to a frozen source with scope `excluded_layer_outside_denominator`
or `pre_reconstruction_non_execution`.

`input_state.py` gained `input_state_values()` emitting frozen descriptive
rows (status `frozen_record_descriptive_not_rescored`, scope
`access_stage_descriptive_not_empirical_drift_measurement`) so the
INPUT-STATE numbers used in Section 5 and the identifiability note are
canonical, not restated.

## Remaining literals (audited, accepted)

- Category A constants appear inside protocol/screen code where they are
  *inputs* (slot count, sample size). Kept — they are the design, and the
  generator reads them from the frozen analysis outputs, not from code.
- Zero/one literals in tests asserting structural properties (e.g. "empty
  leak set", "no unresolved references") are test assertions on pipeline
  invariants, not copied empirical results.
- `paper_status.json` eligibility counts and `main_analysis.json`
  denominators are sources, not duplicates.

Machine-readable findings: `results/NUMERIC_HARDCODE_FINDINGS.csv`.
Provenance map: `results/NUMERIC_PROVENANCE.csv` (regenerated from
`manuscript_values.csv`). Render map: `results/MANUSCRIPT_NUMERIC_TRACEABILITY.csv`.
