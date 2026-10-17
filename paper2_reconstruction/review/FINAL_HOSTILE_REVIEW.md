# Final hostile review (pre-D, from-scratch pass)

Reviewer composite: skeptical metascience referee + statistician +
reproducibility specialist + research-software engineer + desk-reject editor +
AI-claims skeptic. F1–F16; actions logged in FINAL_HOSTILE_REVIEW_ACTION_LOG.md.

## F1 — Representativeness
100 of 461 route-eligible of a fixed EPJ corpus: finite, source-bound. Manuscript
already scopes claims to "source-corpus-specific finite-sample estimate";
verified present. No universal claim permitted (kept).

## F2 — 10/100 attrition
Could be attacked as 90% waste. Reframed correctly in text as the *estimand*:
pre-reconstruction non-execution census. Verified phrasing "pre-reconstruction
non-execution" used; non-executed ≠ failure stated.

## F3 — Gate strictness (access/specification)
Referee may claim gates were unrealistically strict, inflating barriers.
Recorded states are auditable per-paper (paper_status.json + B review); gates
frozen ex ante in the protocol. No ex-post loosening made.

## F4 — Resource-ceiling dominance
Slots capped by wall-clock/compute ceiling; failures dominated by
wall_limit stop reasons could be "just budget". This is the prespecified
ceiling estimand — "under the prespecified resource ceiling" wording retained.

## F5 — Leakage detector false positives
10 flags → 1 confirmed/7 corrected/2 unresolved shows detector imperfection.
Reported verbatim; no claim of detector infallibility.

## F6 — Agent/model dependence
Single agent scaffold, single period. Stated as limitation; no
cross-model claim made.

## F7 — Slot independence
3 slots per paper share the paper — within-paper correlation means effective
denominator is closer to 10 than 30. Manuscript reports majority success at
paper level (not slot-level), which is the right unit; verified.

## F8 — Provider/tool limitations
Broker-side tool restrictions were controlled by the firewall harness
qualification (custodian controls). Documented; residual risk stated.

## F9 — Dual frozen/adjusted reporting
Risk: "adjusted" reads as upgrading results. Enforced separation (Section 4
sealed vs Section 5 provisional; D excluded; QC phrase-checked). Kept.

## F10 — No real human validation
The largest substantive gap — acknowledged as pending Section 6, not
fabricated. Package classified EXCEPT D AND INVESTIGATOR SIGN-OFF.

## F11 — Investigator sign-off
Provisional review is unsigned. Now partially mitigated: investigator_signoff/
package prepared with empty decision fields; still pending by design.

## F12 — Overlap with prior EPJ paper
Paper I covered the source corpus; Paper II adds the prospective
reconstruction funnel + input-state layer. Novelty statement grounds
combination (fixed corpus, full funnel, information boundary, withheld
implementation, prespecified budget, blind runs, sealed+verification layers,
descriptive reveal, planned human validation). No "first" claims.

## F13 — Novelty vs PaperBench/ReplicationBench/Paper2Code/CORE-Bench/
ScienceAgentBench
Those benchmarks use provided artifacts or author clarifications; this study
withholds original implementation and measures the *publication-only*
reconstruction path plus barrier census. Caveat text kept in related work.

## F14 — Input-state identifiability overclaim
Descriptive access-stage statistic only — canonical rows carry
`frozen_record_descriptive_not_rescored` / `..._not_empirical_drift_measurement`
scope; text states it did not measure drift frequency.

## F15 — Generalization beyond source frame
Limited to EPJ commercial-software corpus; stated.

## F16 — Remaining hard-coded prose risk
Rechecked (see updated HARDCODE_AUDIT.md): no new literal empirical values in
generator prose; design constants remain only where they are inputs.

## Rejected attacks (no change)
- "Rerun with bigger budget" — prohibited; ceiling is the estimand.
- "Reclassify non-executed as failures" — would misstate the estimand.
- "Report slot-level rates to look better" — wrong unit.
