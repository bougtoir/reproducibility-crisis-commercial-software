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
  `paper_registry.csv` is the derived unique-PMID identity view.
  The author approved this identity rule in `data/frame_decision.json`.
  `inference_frame.csv` contains one row per PMID with a deterministic disjoint
  sampling stratum; `results/frame_manifest.json` binds it to that decision and
  the unchanged deposit. No replacement papers have been sampled.
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
validation, reported-decimal comparison, deterministic document generation and
unit tests. Exact finite-population intervals and stratified policy identification
bounds are implemented and tested, including unresolved outcomes and gates.
The build refuses to overwrite modified paper-level funnel assessments
or populated experiment tables. Document metadata uses `SOURCE_DATE_EPOCH`
(default 1980-01-01 UTC), fixed SVG IDs and normalized Office/ZIP entries.

Infrastructure probes are separate from the preparation build:

```sh
cd paper2_reconstruction
PYTHONPATH=src .venv/bin/python -m paper2.isolation --output data/raw/isolation-new-run
PYTHONPATH=src .venv/bin/python -m paper2.model_api --output data/raw/api-new-run
PYTHONPATH=src .venv/bin/python -m paper2.controller --output data/raw/controller-new-run
PYTHONPATH=src .venv/bin/python -m paper2.scientific_image --output data/raw/science-new-run
PYTHONPATH=src .venv/bin/python -m paper2.acquisition --output data/raw/candidate-acquisition
PYTHONPATH=src .venv/bin/python -m paper2.corpus_acquisition --output data/raw/corpus-metadata
PYTHONPATH=src .venv/bin/python -m paper2.corpus_evidence \
  --metadata data/raw/corpus-metadata --output data/raw/corpus-articles
PYTHONPATH=src .venv/bin/python -m paper2.funnel_screen \
  --metadata data/raw/corpus-metadata --articles data/raw/corpus-articles \
  --output data/raw/pilot-screen --candidates-per-field 4
```

`paper2.funnel_screen` applies the candidate-review instrument to frame papers
in the deterministic pilot order, using only the retained identity-verified
metadata record and, when present, the retained open PMC JATS article of the same
paper (hash-checked against the corpus indexes). It records the article-text
route separately from the provisional G1/G2 proposals, downgrades abstract-only
negatives to uncertain, retains interrupted API calls as unknown usage without
retrying, and never writes to the funnel. `results/pilot_candidate_screening.json`
summarizes one such run; it is neither validated classification nor pilot selection.

Candidate evidence can be reviewed separately with `paper2.candidate_review`,
passing `--source` for the candidate-acquisition directory and `--output` for a
new private review directory. This instrument preserves final model responses
and checks citations against the retained publication bytes. Rejected evidence
remains unresolved; absent full text cannot establish a negative gate. Its
proposals are neither validated funnel labels nor selected targets. It does not
modify the funnel. An existing review directory resumes retained calls without
reissuing them; changed source text or prompts require a separately named run.
PDF extraction requires the system `pdftotext` utility.

The first command requires Linux Docker/cgroup v2 and the pinned Python image;
the second requires `DEEPSEEK_API_KEY` in the process environment and makes two
billable synthetic calls. Use a new output directory for every qualification.
The API client retains final content and token usage, never private reasoning or
the credential. Qualification uses one CPU and 2 GiB, not the proposed scientific
run envelope, and does not authorize a pilot or main experiment.

The controller permits only isolated Python and a structured final report, hashes
the effective context and tool registry, and seals a chained event journal.
Before removing a live worker it exports its work files through the isolated
process, preserving the archive and hashes. Absolute/traversing paths, links,
special files, duplicate members and oversized exports are rejected.
Its demonstration is synthetic; claimed reports are not adjudicated outcomes.
The scientific-image command builds a hash-locked general Python stack using
binary wheels and version-pinned ClustalW/IQ-TREE packages. It checks imports,
containment and alignment/tree construction on explicitly synthetic sequences,
retaining the outputs, executable hashes and installed system-package inventory.
The immutable local image identifier binds that environment; transitive apt
versions are recorded but are not fully locked for future rebuilds.
These checks do not establish paper-specific dependency closure or scientific validity.
The controller runs each API call in a separate trusted process with a wall deadline
limited by the remaining run budget. Expired calls stop without retry; a receipt
preserves unknown usage and does not assert server-side cancellation. Hidden
provider accounting and cancellation after a disconnected request remain unobservable.

`paper2.packages` checks the allowed/blocked CSV templates before preparing a
worker mount. Pass `--evidence-root`, `--allowed`, `--blocked`, `--output`,
`--paper-id` and `--policy-version`. Paths in the custody CSVs are relative to the
evidence root; `item_id` becomes the safe filename inside `output/input`.
It requires reviewed exact hashes, byte sizes, UTC dates, explicit rights and
complete allowed artifacts. Only the `input` directory may be mounted; the
blocked manifest and custody records remain outside it. Exact reviewed ZIP
members retain archive lineage; other content transformations are not qualified.
Unknown decisions, hash conflicts, symlinks, traversal and excessive sizes fail
closed. Mechanical checks do not establish the correctness or independence of
the declared custodian review and do not authorize empirical runs.

The candidate acquisition command ranks one unassessed candidate per assigned field,
preserves exact Europe PMC metadata and permitted full-text XML responses with
UTC/hash/rights receipts, and resumes from integrity-checked snapshots. HTTP errors,
missing PMC records and non-open metadata flags do not establish unavailable
inputs; lawful alternative full-text sources still require investigation.
These candidate acquisitions are neither final pilot selection nor G1–G5 labels.
Raw evidence remains Git-ignored and must be recovered through the session's
private evidence archive when moving to another machine.
After acquisition stops, run `paper2.corpus_evidence` with the same `--metadata`
and `--output` plus a new `--audit-report` path. The offline audit independently
recomputes article identities and receipt/index assertions. An interrupted index
requires explicit `--allow-partial`; its remaining frame rows stay unprocessed.
Complete index coverage means the configured PMC endpoints were assessed, not
that every paper's full text or research inputs are available.
Candidate reviews select numbered source segments; retained quotations are
resolved from those immutable segments without model rewriting. Invalid segment
IDs, unbound quotations, overlong summaries and incomplete responses remain
unresolved. These machine proposals do not validate G1–G5 or select pilot targets.

The corpus-metadata command queries the complete unique-PMID frame in bounded
batches. Invalid batch envelopes are retained and decomposed into smaller requests;
invalid single-record responses remain unknown. Every accepted record must have
the requested PMID and MED source. The index distinguishes identity verification,
not returned, invalid response and request failure. It stops after three consecutive
failed HTTP batches and preserves the partial index. Metadata alone is not a
full-text or input-access assessment.

`corpus_evidence` first verifies exact ordered frame coverage and every metadata
receipt/index assertion. It then retrieves standalone PMC article XML where the
metadata identifies an open endpoint, with two concurrent requests at most.
Article acceptance requires matching PMID/PMCID and a nonempty JATS body.
Failed, mismatched and unselected endpoints remain distinct; no status establishes
input-data availability or absence of lawful alternative article sources.
Article licences are retained for local review; no third-party bodies are published.

`paper2.timestamp --source FILE --output NEW_DIRECTORY` sends only an RFC 3161
SHA-256 digest request to FreeTSA, retains the payload/request/response and pinned
CA certificate, and verifies the signed response using OpenSSL. The external
timestamp proves existence of those bytes, not protocol adequacy, preregistration
or study completion. Synthetic reports may be timestamped without freezing any
scientific protocol.

Not implemented or validated: complete classification of all paper texts,
the complete solver information firewall, an authorized empirical run controller,
full target-type
scoring, pilot, prospective protocol freeze, human assessments, empirical statistical
models, or final journal submission. The contract and protocol documents are
explicitly **DRAFT — NOT FROZEN**. A normal Devin child session is not evidence of
the specified isolation.

## Release

Work is directly on the public EPJ-linked repository, in a separate feature branch;
the original study is untouched. A private `wip` synchronization is not needed for
these changes. Do not merge other unrelated directories or upload full-text evidence.
Release only the reviewed Paper II code, lawful derived identifiers, protocols,
result provenance, and regenerated document artifacts. The original distinct-paper
claim remains discrepant, but the author has approved the unique-PMID frame for
Paper II. Collect actual study evidence before tagging a submission release.
