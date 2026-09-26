# Protocol amendment AMEND-2026-09-25-05: post-pilot final protocol freeze

**Status: written after the seven-paper prospective pilot was executed and
sealed (blind outcome SHA-256
`485967ff8a1caf85c1d9b1bc102ada617219e3636ea22b7dcbbc5337bb81f3db`) and
after the prespecified procedural review `PILOT-REVIEW-2026-09-25`, before any
pilot value was adjudicated, before any original implementation was revealed,
and before any main-study paper was selected. It freezes the reconstruction
protocol, target rules, tolerance rules, failure taxonomy, logging schema and
resource ceilings for the main study. Hashes of every frozen document are in
`data/adjudication/protocol_freeze_20260925.json`, RFC 3161 timestamped under
`data/raw/protocol-freeze-20260925/`.**

## Basis

The procedural review (data/adjudication/pilot_procedural_review_20260925.json)
examined only harness behaviour: stop reasons, ceilings, journal completeness,
report schema and value provenance. It computed no success or failure rate and
compared no observed value with a publication. Its findings and the rule each
receives are below. No target, tolerance or conclusion rule changes; nothing
below can raise or lower the observed success rate of any pilot slot, because
pilot slots are not re-run and remain outside the primary denominator.

## 1. PD-01 — resource ceilings made internally consistent

Sixteen of twenty-one pilot slots stopped at the cumulative provider-token
ceiling (400,000 tokens) after a median of 23.5 tool calls and at most 139 s of
a 14,400 s wall ceiling; no slot reached the 40-call ceiling and no slot ran a
heavy computation. The token ceiling therefore silently overrode the tool-call
and wall ceilings the protocol described as binding. The final ceiling sets the
cumulative provider-token ceiling to **1,500,000 tokens per slot**, the value
that makes the 40-call ceiling attainable at the 95th-percentile observed
context size (31,844 prompt tokens × 40 = 1,273,760, rounded up to the next
500,000). Wall (14,400 s), step (3,600 s), tool-call (40), completion (8,192),
memory (3 GiB), single CPU, no network and the 64 MiB artifact export cap are
unchanged. The ceiling is a harness consistency correction and is applied to
all main-study slots identically; it is never adjusted per paper.

## 2. PD-02 — failure codes are taxonomy identifiers

Reports carried free-text tokens (`NO_ANALYSIS_EXECUTED`,
`RUN_TERMINATED_LENGTH_LIMIT`, `TARGET_NOT_OBSERVED`) in `failure_codes`.
Final rule: `failure_codes` must be identifiers `F01`–`F20` or `F99` from
FAILURE_TAXONOMY.md. Any other token is retained verbatim in the sealed report,
mapped to `F99` for tabulation, and adjudicated as *unclassified*. No code is
added to the taxonomy after this freeze.

## 3. PD-03 — final report without any execution

One report was submitted before any worker execution and asserted a harness
termination that the journal does not show. Final rule: a final report with
zero journaled executions is classified `no_execution_report`. The slot counts
as attempted; L2 and L3 are `not_reached`; any `observed_value` is void; any
procedural statement in a report is checked against the journal and never
accepted from the report alone.

## 4. PD-04 — value provenance is required for L3 and L4

The served publication text displays the target value, so a value transcribed
from the article is indistinguishable, in the report alone, from a computed
one. Final rule: for L3 (`execution_successful`) and L4
(`numerical_target_reproduced`) the adjudicator must locate the observed value
(to the reported precision) in a journaled execution output produced by
journaled code that reads served inputs. If it cannot be located, L3 and L4 are
`not_established` and the slot is recorded with `F15` if code exists or
`no_execution_report` otherwise. The blind is a blind to the original
implementation and its outputs, not to the publication; the manuscript states
this and reports the provenance check explicitly.

## 5. Frozen documents

PROTOCOL.md, SAP.md, TOLERANCE_RULES.md, FAILURE_TAXONOMY.md,
TARGET_SELECTION_RULES.md, INFORMATION_FIREWALL.md, REVEAL_PROTOCOL.md, this
amendment and its predecessors, the logging schema (journal event kinds and the
report schema in `src/paper2/controller.py` and `src/paper2/isolation.py`) and
the final ceiling in `src/paper2/protocol_freeze.py`. Their SHA-256 values are
listed in the freeze record. After the freeze, any change to these documents
requires a new numbered amendment, and any change made after a main-study slot
has opened is reported as a deviation.

## 6. Main study sampling frame

The eligible reconstruction population is every unique-PMID paper for which
Stage A (deposit statement in lawfully retained text) and exhaustive Stage B
(anonymous listing resolved, within caps, not implementation-only) record a
verified open deposit route, minus (a) the seven `PROSPECTIVE_PILOT` papers,
(b) the seven `ACCESSIBILITY_GATE_FAILED` papers, (c) every paper carrying a
machine-readable exclusion class, and (d) every paper whose retained text was
inspected for any earlier candidate review. The SAP allocation (one place per
nonempty field, remaining places by proportional-allocation deficit with
lexicographic ties, within-field order by SHA-256 of `paper2-main-v1|paper_id`)
is executed once; N_h, n_h, π_h and weights are saved before any paper's G1–G5
or target freeze begins. Each selected paper then receives the ten-field
prospective freeze of AMEND-2026-09-25-04 §4 with the same G1–G5 delegated
assessment and pending investigator verification; papers failing G1–G5,
leakage review or deposited-input matching are recorded with their class and
are not replaced (the primary estimate is reported on attempted papers with the
attrition table; the SAP's finite-population bounds cover the unattempted).

## 7. Unchanged

Target selection hierarchy; tolerance rules; L1–L5 definitions and the
primary endpoint (≥ 2 valid successes among 3 fixed slots per paper); three
sequential single-VM slots; the information firewall; descriptive-only reveal;
no Paper III analysis; no original-author contact; pilot outside the primary
denominator.
