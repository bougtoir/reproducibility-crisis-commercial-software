# Paper II protocol

**FROZEN 2026-09-25 by AMEND-2026-09-25-05 after the seven-paper prospective
pilot (`PILOT-FREEZE-2026-09-25`) was executed, sealed and procedurally
reviewed. The frozen text hash and RFC 3161 timestamp are in
`data/adjudication/protocol_freeze_20260925.json`. Investigator verification of
delegated G1–G5 gates remains pending; pilot and accessibility-layer papers are
excluded from the primary denominator. No main-study slot had opened at the
freeze. Later changes require a numbered amendment.**

## Question and scope

Can publication-level scientific description support independent reconstruction
of a computational research system? Define *publication-grounded independent
computational reconstruction* operationally, with a specified agent, information
boundary, resource budget, target and comparison rule. No claim of a new term or
first agentic reconstruction study is made.

The framework is ACCESS → RECONSTRUCT → EXECUTE → REPRODUCE → ROBUST. Paper II
measures reconstruction and its execution/numerical/conclusion consequences.
ROBUST is future Paper III work and is excluded. The Wet/Dry mapping is a conceptual
analogy, not a literal equivalence or an empirical result.

## Frame and eligibility

The source is the VOR-linked EPJ deposit, identified in the acquisition ledger.
Preserve all rows. The automatically generated source audit distinguishes source
records, unique PMIDs and overlapping field memberships. The requested distinct
paper count is not met by the inspected deposit. No new papers are substituted.
The author approved the unique-PMID inference frame while retaining all original
rows. The verbatim decision is in `data/frame_decision.json`; the generated
`results/frame_manifest.json` binds the decision, source and inference-frame hashes.
This resolves the Paper II identity rule, not the historical distinct-paper claim,
classification validity or prospective protocol freeze.

Paper identity is PMID; duplicate DOIs or titles trigger review, not automatic
merging. Records are traceable by commit, file and one-based data-row position.
Field labels describe recorded sampling provenance, not mutually exclusive
scientific disciplines. Neither papers nor historical sample selection can be
generalized to all PubMed or all science using unverified inclusion probabilities.

G1: yes/no/uncertain computationally testable, with evidence and reason.
Include a central claim produced by a specified computational transformation,
statistical estimation, simulation, optimization or predictive evaluation.
Descriptive quantitative analysis is eligible when reproducing its central
reported output requires applying a computational procedure to inputs. A purely
verbal claim or restatement of a supplied number is ineligible. Missing inputs
or inadequate reporting do not turn an otherwise computational paper into G1=no.

G2: yes/no/uncertain identifiable principal target under the frozen hierarchy.
G3: public / obtainable_without_payment / restricted_but_potentially_accessible /
unavailable / insufficiently_identified / not_applicable / unknown.
G4: required function available under the commercial-software policy, with
separate version and access reasons. G5: sufficient / insufficient / uncertain
specification for an attempt. G6–G9 concern created implementation, execution,
numerical agreement and preservation of the target-linked conclusion.

Store gates independently and derive sequential at-risk sets. Preserve unknown,
unassessed and inaccessible categories. No gate is inferred from an EPJ regex
flag. Ordinary free public registration may be allowed if available to any reader;
individual institutional permissions are restricted even when no fee is charged.
Transient failed downloads remain unknown after a documented limited retry policy,
not evidence that the underlying data do not exist.

Article-text obtainability is recorded separately from G1–G5 as a prospective
access status: `lawful_text_retained` (identity-verified open PMC JATS, a
licensed publisher/repository copy, or an author-supplied copy), `no_lawful_public_text_located` (only closed locations found after the documented route checks) or
`route_check_incomplete`. A paper whose allowed material cannot be lawfully
retained before the pilot cannot be a pilot case and is recorded as prospectively
unavailable to the pilot with the routes checked; this is not a G3 input verdict,
not a G1 negative and does not remove the paper from the main frame. Machine
screening of retained text yields provisional G1/G2 proposals only; human
adjudication against the retained bytes is required before any gate is recorded.
If the investigator delegates the first pass, the delegated gates are recorded as a
named primary-assessment record with a deviation entry, every gate citing segment
IDs of the hash-checked retained bytes, and remain provisional: the funnel stays
unassessed and no pilot selection is frozen until the investigator team verifies or overturns
each gate. Delegated or human adjudication may skip a candidate only for absent
lawfully retained text or for G1/G2 ineligibility, never for G3, G4, G5 or
expected reconstruction difficulty, and every skipped candidate is retained with
its reason in the deterministic chain.

**Amendment AMEND-2026-09-24-03** (`AMENDMENT_AMEND-2026-09-24-03_deposited_data_criterion.md`,
hash-frozen and RFC 3161 timestamped before any screening under it) restricts
the main reconstruction sample to papers whose deposited analysis data a third
party can obtain lawfully and immediately. Under the amendment, author-request,
application-required, institutional-licence, size-unannounced, above-cap,
listing-unavailable and output-only deposits are sampling-frame exclusions
recorded with their class, not G3 study outcomes. Papers that pass G1/G2 but
fail the amended access criterion form the `ACCESSIBILITY_GATE_FAILED` layer
(`data/adjudication/accessibility_gate_failed_*.json`): an audit sample of
pre-reproducibility access barriers for an auxiliary analysis or a separate
paper, from which no reconstruction success or failure rate may be computed
because no member reached a reconstruction attempt. The seven papers of the
provisional pilot `DEV-2026-09-24-02` are the first members of that layer and
are not in the commercial-software reproducibility sample.

## Prospective pilot and freeze sequence

After harness qualification, select seven pilot papers,
one from each nonempty disjoint field assignment, by the deterministic hash
procedure in SAP. These are feasibility cases excluded from the main sample.
Use main-frame replacements only if prospectively ineligible, recording every
reason; do not replace difficult reconstruction outcomes.

Pilot only the same information-access intervention intended for the main study.
Use it to finalize target comparability, tolerances, scoring, access procedures,
budgets, stopping rules and telemetry. No post-pilot numerical choices are final
until all named protocol documents, prompts, manifests, image and tool registry
are hashed and externally timestamped. A Git commit or this draft's date alone
does not establish prospective freeze.

Freeze all documents in this directory, resolve every `NOT_COMPLETED` gate,
record the model/version available to the investigator, preregister when an
appropriate public record is available, then sample the main study and seal
paper-specific targets before runs. Later changes require an explicit deviation
record, not a rewritten historical protocol.

## Measurement and governance

Three fixed independent slots per paper use identical prompts, packages and
permissions and equivalent model/resource settings. No shared mutable state,
parent debugging, sibling results or original implementation is permitted.
Outcomes L1–L5 remain separate. Numerical agreement is primary; conclusion
preservation is secondary and must not rescue a numerical mismatch.

Validated policy barriers remain in the all-eligible policy estimand but are not
invented failed runs. Unresolved eligibility, access, contamination or missing
runs remain unresolved. Agent failure cannot establish universal human
impossibility or publication inadequacy.

Human scientific interpretation and final author responsibility are required.
Author list, affiliations, contributions, funding, interests, ethics determination
for human performance data, and APC arrangements are not supplied or presumed.
