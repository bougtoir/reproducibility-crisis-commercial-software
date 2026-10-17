# Hostile review (adversarial pre-submission pass, revision round)

Reviewer posture: assume a skeptical referee trying to desk-reject. Findings
C1–C10; each mapped to an action in HOSTILE_REVIEW_ACTION_LOG.md.

## C1 — "0/10 vs 1/10 is a non-result dressed as a finding"
The frozen delegated-mechanical outcome is 0 successes among 10 attempted;
the provisional adjusted estimate proposes 1/10 pending Human Validation D.
Risk: referee reads this as "the method failed". Mitigation already in text:
the paper's claim is the *measurement protocol and barriers census*, not a
reconstruction rate; the abstract states results are provisional and D is
pending. Strengthened: abstract now renders the frozen/adjusted split from
canonical values so the two numbers cannot drift.

## C2 — Denominator ambiguity
"10 attempted of 30 slots of 100 sampled of 461 eligible" invites
denominator-confusion complaints. Mitigation: Methods 2.1a/2.2 now name each
denominator layer explicitly, with pilot (7) and gate-failed (7) cohorts
labeled `excluded_layer_outside_denominator` in canonical rows.

## C3 — Provisional adjudication may look like outcome-shopping
Provisional A/B/C/E review could be attacked as post-hoc relabeling.
Mitigation: kept verbatim the separation "sealed delegated mechanical results
(frozen)" vs "provisional, separate from Section 4; D excluded"; QC check
enforces the phrasing. No outcome-informed change to success criteria was
made in this revision (prohibited and verified: no diff to
protocol/SAP/tolerance files).

## C4 — Input-state descriptives could be read as the drift measurement
119 entries / 51 papers / 19 unspecified (11 papers) could be mistaken for a
reconstructed-inputs result. Mitigation: canonical rows carry
`frozen_record_descriptive_not_rescored` status and
`access_stage_descriptive_not_empirical_drift_measurement` scope; Section 5
text states they are access-stage descriptives.

## C5 — Software-vs-data asymmetry overclaim
Original wording "versioned software can usually be redistributed or rebuilt"
is falsifiable (licensing, legacy-version access). **Fixed** per prompt H5:
now "may in some settings be preserved or rebuilt more readily than
third-party historical data states, although licensing and legacy-version
access can also prevent software redistribution."

## C6 — Package not self-contained
The slot-package test required wheels under `data/raw/vetted-wheels-20260924/`
(gitignored) with no fetch path — a reviewer reproducing tests on a clean
checkout fails. **Fixed**: `make fetch-vetted-wheels` downloads the three
pinned wheels and verifies each sha256 against requirements.lock; documented
in PACKAGE_SELF_CONSISTENCY.md and README.

## C7 — Private evidence absence looks like missing data
`evidence_reference_resolution.csv` / `verification_import_audit.json`
regenerate as `evidence_package_not_mounted` off the private evidence mount.
Referee risk: "results can't be regenerated". Documented: these artifacts
require the private evidence package; committed versions are the authoritative
private-mount outputs and are restored, not regenerated, on public machines.

## C8 — Stale documentation contradictions
Pre-main-study README/EVIDENCE/REVIEWER_AUDIT claims ("no main runs exist",
"no pilot", "no frozen outcomes", "no reveal") flatly contradict the frozen
results — a referee reading the repo top-down would call the record
inconsistent. **Fixed**: superseded banners + current-state pointers;
canonical README section added. See DOCUMENTATION_STALENESS_AUDIT.md.

## C9 — Traceability gap between text and tables
Numbers in prose had no enforced link to `manuscript_values.csv`; a future
edit could silently fork. **Fixed**: generator reads all empirical numbers
from canonical rows; MANUSCRIPT_NUMERIC_TRACEABILITY.csv maps each claim to
its manuscript location.

## C10 — "Only" framing
"Only 10 of 30 slots executed" can be spun as underpowered rather than
protocol-honest. Kept the factual phrasing (attempts are protocol-capped at
10); the barrier census — not power — is the estimand.

## Items rejected by the hostile reviewer (no change made)
- Re-executing B-reclassified papers / reveal code: prohibited, not done.
- Multiverse/alternate tolerances: prohibited, not done.
- Converting provisional review to human adjudication: prohibited, not done.
