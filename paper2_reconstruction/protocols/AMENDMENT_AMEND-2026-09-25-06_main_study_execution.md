# Protocol amendment AMEND-2026-09-25-06: main-study execution layer

**Status: written after the final protocol freeze `PROTOCOL-FREEZE-2026-09-25`
(SHA-256 `f3bdc79848672d06785081e339c8e6c26da26cae73b878d49687b933c2f4ef0f`)
and after the main cohort `MAIN-COHORT-2026-09-25` (SHA-256
`c93708937c2ccf7f67d2e9c79b30db549caf6112ef7e5f82dd93e2becb286b4f`) was
selected and timestamped, before any main-study paper was specified, before
any main-study input was retrieved and before any main-study solver call. It
records how the frozen protocol is executed for the cohort and one clerical
correction. No target, tolerance, conclusion or failure rule changes.**

## 1. Clerical correction to TARGET_SELECTION_RULES.md

The document hashed by the final freeze (SHA-256
`fc879735acc5698093d1476db7e4d87ac4174807f5cbea0eb2068e91162a99da`) still
carried the header line `**DRAFT — NOT FROZEN**` although AMEND-2026-09-25-05
froze it by hash together with every other protocol document. The header is
replaced by a frozen-status line that cites this amendment. No rule text
changes; the freeze record keeps the original hash, and the corrected hash is
recorded in `data/adjudication/main-study-20260925/` at the first main-study
freeze.

## 2. Delegated G1–G5 assessment and ten-field specification

For every cohort paper, in the order of `results/main_cohort_20260925.csv`, a
delegated model pass (the same provider/model as the solver, deterministic
settings, one request, journaled under `data/raw/main-study-20260925/
specifier-api/`) receives the segmented article text and the anonymous deposit
file listing retained at Stage B. It returns G1–G5 with verbatim quotes, the
exact target and its printed values, a blind description of the target, the
target type, location, required deposit files, leakage-suspect files, external
reference requirements, comparison metric, agreement criterion, specification
gaps and rejected alternatives.

The harness verifies programmatically that every quote occurs verbatim in the
cited segment or its immediate neighbours (segmentation is a harness artifact
that can split sentences), that every printed target value occurs in the cited
location
quotes, that the blind description contains no printed value, that the target
type has a complete comparison rule in TOLERANCE_RULES.md and that required
files exist in the listing. A paper is `specifiable` only if no gate is `no`,
the target value is located verbatim, at least one listed input is required and
no external reference resource beyond the served material is required.
Otherwise the paper is recorded with its machine-readable state
(`gate_failed_delegated`, `target_value_not_located_verbatim`,
`required_inputs_not_in_public_listing`,
`external_reference_resource_required_not_served`,
`specification_not_obtained`) and is **not** run. It stays in the cohort and
in the denominator defined by the SAP for post-selection ineligibility; it is
never replaced. Every delegated assessment is `investigator_verification:
pending` and `human_adjudication: pending`.

## 3. Input acquisition and leakage exclusion

Only the files the specification names are retrieved, anonymously, from the
public deposit route retained at Stage B, streamed to
`/home/ubuntu/paper2_evidence/main-inputs-20260925/<PMID>/` with an immutable
receipt (URL, identifier, UTC time, request conditions, bytes, SHA-256,
announced checksum, completeness). Frozen caps: 2 GiB per file, 8 GiB per
paper, retrieval stops when local free storage would fall below 12 GiB
(recorded as `local_storage_exhausted`, not as an access verdict). A file named
by the specifier as answer-bearing, or in which any printed target value occurs
as a whole numeric token, is excluded as
`target_or_data_leakage_*` and never served. Result-like header columns are
recorded for adjudication but do not by themselves exclude a file. A paper
whose every required input is leakage-excluded is recorded
`all_inputs_leak_target` and not run.

## 4. Per-paper prospective freeze

Before a paper's first slot its ten-field record (exact target, location with
quotes, required inputs with SHA-256 and bytes, allowed resources, blocked
original artifacts, comparison metric, agreement criterion, conclusion
criterion, final resource ceiling, stopping rule) is written to
`data/adjudication/main-study-20260925/freezes/<PMID>.json` and RFC 3161
timestamped. The solver instruction carries only the blind description.

## 5. Slots, order and quota stop

Exactly three sequential independent slots per specifiable paper, each in a
fresh isolated worker under the final ceiling of AMEND-2026-09-25-05
(1,500,000 provider tokens; other limits unchanged), on one machine. Papers are
processed strictly in cohort order. Execution ends when the investigator's
provider quota is exhausted (three consecutive slots ending in a controller
error before any tool call, or a provider stop during specification) or when
the cohort is complete. Papers not reached, or reached without all three slots
completed, are sealed as `not_run_resource_exhausted`; they remain in the
cohort as unresolved and are never replaced or counted as failures. Completed
slots are sealed before adjudication or any reveal.

## 6. What this amendment does not do

It does not change any frozen rule, does not re-run or re-score the pilot, does
not compute any success or failure rate, does not contact original authors and
does not perform any Paper III analysis.

## 7. Deposited code is withheld (DEV-2026-09-26-01)

On the first specifiable cohort paper the delegated specifier named deposited R
scripts in `required_files` and the harness served them. Deposited scripts,
notebooks and workflow files are original implementation artifacts under
INFORMATION_FIREWALL.md and are never served; they are excluded as
`original_implementation_artifact_withheld` by file name, the specifier is told
not to list them, and a paper whose every required input is code is recorded
`all_inputs_original_implementation` and not run. The two started slots are
void (retained under `voided/`, in no denominator); the paper is re-frozen and
run with three fresh slots. Details: `data/adjudication/main-study-20260925/
deviations/DEV-2026-09-26-01.json`.
