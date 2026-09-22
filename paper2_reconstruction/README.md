# Paper II — preparation package, not a completed empirical study

This branch preserves the EPJ source and creates a prospective reconstruction study
workspace. No reconstruction rates, human results, or protocol-freeze claims are
generated. `results/readiness.json` is authoritative for the current study status.

## Reproduce the preparation outputs

Prerequisites: Git, Python 3.10+, Python venv, and Make. From the repository root:

```sh
make -C paper2_reconstruction setup build check
```

`setup` installs the hash-locked dependencies in an isolated virtual environment.
The lock was resolved with uv 0.6.17 and an exclusion date of 2026-09-15.
`build` restores complete Git blobs from source commit
`2414f8d65cf2fe33c744c0a62cc05b6c77868367`, verifies hashes, creates immutable local
snapshots, and generates tables, a preparation DOCX, an editable figure PPTX, and
separate PNG/SVG figures. It does not download corpus papers or consult their code.
If the original commit is absent in a shallow checkout, fetch that commit first.

```sh
make -C paper2_reconstruction submission-check
```

This command **must currently fail**. Passing software tests does not make an
unperformed study submission-ready. The manuscript is a labelled preparation
draft, with prospective Methods and an observed source audit, not Paper II results.

## Data and interpretation

- The original deposited source files remain unchanged. Local raw snapshots under
  `data/raw/epj` are ignored by Git and can be recovered from the pinned public
  commit. `data/acquisition_ledger.csv` records exact provenance and checksums.
- `data/derived/source_record_bridge.csv` retains every source row and field.
  `paper_registry.csv` is a derived unique-PMID identity view, **not an approved
  replacement sampling frame**. No replacement papers have been sampled.
- EPJ code/data/software indicators are historical text detections. They are not
  access, eligibility, target, specification, or reconstruction assessments.
- `funnel_NOT_ASSESSED.csv` is an explicitly unassessed work queue. Empty
  experiment tables contain headers only; they are not zero-success experiments.
- `results/manuscript_values.csv` connects source-audit values to the source,
  checksum, analysis, and derived result. Planning intervals are analytic
  calculations under assumed proportions, not empirical estimates.
- Raw publisher evidence from the independent preparation audits is retained in
  the session evidence archives referenced in `review/EVIDENCE.md`. It is not
  redistributed publicly. The public reference ledger contains metadata and
  short analytical summaries only.

## Implementation boundary

Implemented: source snapshot/integrity checks, duplicate/conflict detection,
row/paper/membership separation, precision planning, deterministic stratified
sampling utility, missing-aware fixed-three-run aggregation, basic outcome/time
validation, reported-decimal comparison, document generation and unit tests.

Not implemented or validated: acquisition/classification of all paper texts,
an enforced solver information firewall, agent run controller, full target-type
scoring, pilot, protocol timestamp service, human assessments, empirical statistical
models, or final journal submission. The contract and protocol documents are
explicitly **DRAFT — NOT FROZEN**. A normal Devin child session is not evidence of
the specified isolation.

## Release

Work is directly on the public EPJ-linked repository, in a separate feature branch;
the original study is untouched. A private `wip` synchronization is not needed for
these changes. Do not merge other unrelated directories or upload full-text evidence.
Release only the reviewed Paper II code, lawful derived identifiers, protocols,
result provenance, and regenerated document artifacts. Resolve the source-frame
discrepancy and collect actual study evidence before tagging a submission release.
