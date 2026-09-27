# Protocol amendment AMEND-2026-09-25-07: primary adjudication, descriptive reveal and prespecified analysis

**Status: written after the main blind seal (`blind_outcome.json`, SHA-256
`350a0e7dcedbb74ccad2472c642318c87cc7d272f1b6c88186a1419dc858ea5a`, RFC 3161
timestamped) and before any original corpus-paper implementation was read. It
records how the frozen rules (PROTOCOL-FREEZE-2026-09-25) are applied to the
sealed outcomes. No target, tolerance, conclusion, failure or estimand rule
changes. Human adjudication and investigator verification remain pending.**

## 1. Delegated primary adjudication (`src/paper2/main_adjudication.py`)

The sealed blind outcome is adjudicated mechanically from the frozen ten-field
specification, the retained served inputs, the journaled code events and the
journaled execution outputs of each slot. Original implementations are not
opened. Rules applied:

- PD-04 provenance: a target component counts as observed only when it appears,
  numerically identical, in the output of a successful execution of journaled
  code that reads the served inputs. Report-only values are `report_only` and
  set `human_review_flag` when they coexist with observed values.
- A slot with no journaled execution is `no_execution_report`; L2/L3 are
  `not_reached` and no taxonomy code is assigned unless a protocol stop code
  applies.
- Stop reasons map to F13 (`wall_limit`, `tool_limit`, `token_reserve_limit`)
  and F15 (`malformed_action_limit`). F20 is reserved for an executed run whose
  observed value does not reproduce the target. F17 marks contaminated slots as
  void, never as ordinary failure. Solver-reported reason text is kept verbatim
  and mapped to F01/F03/F20 only by the fixed patterns in the module; otherwise
  F99.
- Paper level: majority (t=2) primary; strict and permissive as sensitivities;
  incomplete triples are never reduced to 0/3–3/3.

The first mechanical record (SHA-256 `8fab19e0…4664a`) assigned codes to
no-execution slots, accepted approximate provenance matches and conflated L3
with L4. It was superseded before any human adjudication by the corrected
record (SHA-256 `6fd42b89…5b28`); both records and both timestamp receipts are
retained (`superseded/README.md`). Blind outcomes were not changed.

## 2. Descriptive reveal (`src/paper2/reveal_ledger.py`)

REVEAL_PROTOCOL.md was frozen by hash in AMEND-2026-09-25-05; its stale
`DRAFT — NOT FROZEN` header is replaced with a frozen-status line and no rule
text changes. Public original-code archives named in the retained article text
are acquired after the seal (receipts with URL, bytes, SHA-256 and UTC time in
`/home/ubuntu/paper2_evidence/main-reveal-20260925`). Paper-specific scripts
that were part of a data deposit and withheld from the solver
(`original_implementation_artifact_withheld` in the acquisition ledger) may be
described from the withheld copy. Every quoted snippet is verified verbatim
against the retained bytes; every source acquisition time is verified to
postdate the seal unless it is a withheld deposit file. Papers whose
publication names no code route receive a `missing` reveal assessment. Nothing
is executed; no blind score is revisited.

## 3. Prespecified analysis (`src/paper2/main_analysis.py`)

SAP.md is applied to the route-eligible population |E| = 461 with the seven
frozen strata. Non-attempted cohort papers are policy non-success under the
recorded delegated states (`primary_policy`), and additionally treated as
unresolved in a `verification_pending_envelope` because investigator
verification of the gate states is pending. The conditional rate among
attempted papers is secondary. Run-level L1–L5, stop reasons, slot states and
failure codes are tabulated by run and by paper. No model, perturbation,
multiverse or alternative-implementation analysis is performed.
