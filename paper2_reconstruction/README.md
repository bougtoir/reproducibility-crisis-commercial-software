# Paper II — Computational Reconstructability (pre-D final state)

Prospective, protocol-frozen reconstruction study of 100 EPJ papers sampled
from the commercial-software corpus. **Status: NOT SUBMISSION READY —
SUBMISSION-READY EXCEPT D AND INVESTIGATOR SIGN-OFF.**

## Current canonical state

- Source corpus: 10,000 records → 9,935 unique PMIDs (64 duplicated), 7 EPJ fields.
- Prospective pilot completed (7 papers × 3 slots); protocol + SAP frozen
  (`PROTOCOL-FREEZE-2026-09-25`).
- Main cohort: 100 papers sampled from 461 route-eligible.
- Blind reconstruction: 10 papers × 3 sealed slots executed; outcomes sealed.
- 90 papers non-executed with auditable barrier states
  (`results/paper_outcomes.csv`, `data/adjudication/main-study-20260925/paper_status.json`).
- Sealed delegated mechanical result (Section 4): 0/10 majority success.
- Provisional AI-assisted verification layer (Section 5): adjusted 1/10 —
  separate estimand, pending investigator sign-off.
- Human Validation D: **pending** (Section 6; not performed, not fabricated).
- Investigator sign-off: **pending** (`investigator_signoff/` package prepared,
  decision fields intentionally empty).
- Paper III analyses: excluded.

Canonical numbers live in `results/manuscript_values.csv`; status detail in
`WITHOUT_D_STATUS.md`; handoff in `FINAL_HANDOFF.md`.

The ACCESS stage is decomposed as: resource accessible? →
exact input state identifiable? →
historical input state retrievable? (see `INPUT_STATE_IDENTIFIABILITY_NOTE.md`).

## Public vs private

- Public: this repository — code, schemas, frozen outcome manifests, proposed
  review CSVs, derived results, generated manuscript artifacts, tests.
- Private/restricted: the evidence package (`/home/ubuntu/paper2_evidence/…`,
  not in this repo). Two artifacts — `results/verification_adjusted/
  evidence_reference_resolution.csv` and `verification_import_audit.json` —
  carry private-mount resolutions; committed bytes are authoritative and are
  not regenerated on public machines (see review/PRIVATE_EVIDENCE_CHECK.md).

## Checks

Public package:

```sh
make setup && make fetch-vetted-wheels && make build && make check
```

`fetch-vetted-wheels` downloads the three pinned wheels (sha256-verified
against `requirements.lock`) needed by the firewall/slot-package tests.
`make check` runs ruff, mypy and 146 pytest tests.

Full private-evidence check additionally requires the mounted evidence
package; differences are documented in `review/PRIVATE_EVIDENCE_CHECK.md`.

`make submission-check` **must currently fail** — submission readiness gates
on Human Validation D and investigator sign-off.

## Archive

Pre-execution documentation is preserved under `docs/archive/`
(`README_pre_execution.md`) and banner-marked `review/EVIDENCE.md`,
`review/REVIEWER_AUDIT.md`. See `review/DOCUMENTATION_FINALIZATION_AUDIT.md`.

## Layout

- `src/paper2/` — build/verification/generation pipeline (`make build`)
- `data/` — frame decisions, verified references, adjudication records, ledgers
- `results/` — canonical `manuscript_values.csv`, all derived tables
- `review/` — audits incl. FINAL_HOSTILE_REVIEW, HARDCODE_AUDIT, QC reports
- `investigator_signoff/` — sign-off request package (empty decision fields)
- `manuscript/` — generated DOCX/PDF/PPTX/SVG artifacts
- `dist/` — `PaperII_preparation_NOT_SUBMISSION_READY.zip`
