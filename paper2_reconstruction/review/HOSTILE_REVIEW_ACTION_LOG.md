# Hostile review action log

| Finding | Action taken | Files |
|---|---|---|
| C1 frozen/provisional drift risk | Abstract + Section 4/5 numbers now render from canonical values only | src/paper2/documents.py, verification_qc.py |
| C2 denominator ambiguity | Named denominator layers + 3 new canonical rows (pilot 7, gate-failed 7, external-ref barrier 28) | src/paper2/verification_layer.py |
| C3 provisional-adjudication attack | Verified no protocol/SAP/tolerance diff; separation phrasing QC-enforced | (no change needed) |
| C4 input-state misread risk | input_state_values() frozen descriptive rows with descriptive-only scope | src/paper2/input_state.py, build.py |
| C5 software overclaim | Wording softened per prompt H5 | src/paper2/documents.py (Discussion 7.1) |
| C6 non-self-contained tests | fetch-vetted-wheels target + hash-verified fetch script + ledger | Makefile, scripts/fetch_vetted_wheels.sh |
| C7 private-evidence regeneration | Committed private-mount outputs restored; documented limitation | results/verification_adjusted/* |
| C8 stale docs | Superseded banners on README/EVIDENCE/REVIEWER_AUDIT | README.md, review/EVIDENCE.md, review/REVIEWER_AUDIT.md |
| C9 text-table traceability | MANUSCRIPT_NUMERIC_TRACEABILITY.csv; generator canonical-only | results/MANUSCRIPT_NUMERIC_TRACEABILITY.csv |
| C10 framing | No change (factual wording retained) | — |
