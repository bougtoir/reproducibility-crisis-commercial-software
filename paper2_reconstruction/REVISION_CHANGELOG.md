# Revision changelog — manuscript package after input-state identifiability

Scope: targeted revision of the generated manuscript package (`src/paper2/documents.py`,
`src/paper2/input_state.py`, `src/paper2/verification_qc.py`). No empirical result,
denominator, sealed outcome, adjudication record or human-validation status changed.

## Unchanged frozen records (verified by `make check` / `review/VERIFICATION_ADJUSTED_QC.md`)

- sealed blind outcome SHA-256: `350a0e7dcedbb74ccad2472c642318c87cc7d272f1b6c88186a1419dc858ea5a`
- primary mechanical adjudication SHA-256: `6fd42b8964e055606cbd24ebddb44f42fa09147fa80bfa2d5bcaa161766c5b28`
- A/B/C/E imported review tree SHA-256: `f6b374b74c48c297e4c68dcdd780a4d2c1f162e8dcfc058111c524408961f78a1`
- reconstruction slots rerun: 0; original code executed: 0; Section D cases: 0 (pending)
- Paper III analyses: none

## Manuscript text (documents.py)

- Subtitle: "Prospective Paper II study" → completed-study wording with human validation pending.
- Abstract rewritten in completed-study tense: unique-PMID frame, ACCESS refinement, pilot and
  protocol freeze, 100/461 sampling, 10 attempted, 30 sealed slots, frozen 0/10 majority, separate
  provisional AI-assisted 1/10 that does not overwrite the sealed result, attrition as primary
  finding, D not performed.
- Methods heading: "prospectively frozen (protocol text)" → "procedures as executed under the
  prospectively frozen protocol". 2.1–2.6 converted from imperative/future to past tense.
  Structure (2.1, 2.1a, 2.2, 2.3, 2.4, 2.5, 2.6) preserved.
  - 2.1: amended deposited-data criterion, 461 route-eligible papers, retained attrition layers,
    delegated G1–G5 with investigator verification pending.
  - 2.2: pilot completed (7 papers × 3 slots, procedural review PD-01, rules unchanged, no pilot
    rate), PROTOCOL-FREEZE-2026-09-25, stratified sampling of 100 from 461, ten-item target freeze.
  - 2.3: "Use three clean, fixed slots…" → completed description with the frozen per-slot ceiling;
    "No qualified harness currently exists." removed and replaced by the calibrated statement that
    a functioning harness was implemented and used, listing controls that passed qualification
    probes, controls not formally qualified, and that prior model-training exposure cannot be
    certified absent.
  - 2.5: "The proposed weighted interval is not yet implemented" / "draft SAP" removed; analyses
    described as performed; no multivariable model; frozen SAP.
  - 2.6: sealing, mechanical adjudication, descriptive reveal and A/B/C/E import described as
    completed; Section D described as prespecified, supported by ten cases from frozen records,
    and not performed. "Prepare approximately 15 … cases" removed.
- Section 3 heading: "Observed preparation findings" → "Source corpus audit".
- Section 4 heading now reads "sealed delegated mechanical results (frozen)".
- Section 5 heading now reads "(A/B/C/E; provisional, separate from Section 4; D excluded)" with
  a lead paragraph: Section 4 frozen; Section 5 provisional pending investigator approval; does
  not overwrite Section 4; no reclassified non-executed paper was executed; reveal never changed
  a blind score.
- Discussion: added "Data/code availability requirements address only selected stages of the
  reproducibility chain; reconstructability also depends on explicit specification, stable input
  identity, executable environments and target-linked evaluation rules." and a statement that the
  framework is conceptual, not an empirically validated universal model. Existing points retained
  (attrition primary, non-executions ≠ failures, conditional success, resource ceilings, leakage
  false positives, agent failure ≠ human impossibility, descriptive reveal, input-state
  identifiability, no drift quantification).
- Statements: data/code availability statement updated (public repository now contains frozen
  protocols, sealed records and layers; inputs and full text retained internally, not
  redistributed); AI-assistance statement now describes the executed roles.
- References heading: "References used in this preparation draft" → "References".

## Tables

- Table 10 (`input_state.py`): code availability row kept in ACCESS, relabelled "ACCESS (artifact
  access)" with purpose "artifact reproducibility only, deliberately excluded from the Paper II
  blind intervention and used solely as a descriptive characteristic and post-seal reveal
  resource".
- Table 11: values re-derived from `paper_status.json` (100 papers; 28 external-reference
  barrier), 100 delegated specifications (119 reference entries in 51 papers; 19 unspecified in
  11 papers; 100/100 accession; 89/100 file list), 10 target freezes (21 inputs) and the retrieval
  ledger (21/21 SHA-256 and timestamps; 7/21 announced checksums, 7 matched). All values
  reconciled; version-identifier, schema/release and historical-version fields remain
  "not recorded"/"not assessed". Two item labels clarified ("recorded by the specifier from
  publication text"; "not specified in publication text"). No value was invented or rescored.

## References

- The 11 input-state references (FAIR, FORCE11, RDA dynamic-data citation, Rauber, Pröll & Rauber,
  Klump, Pasquier, SWHID, Sandve, Stodden, Zhao) remain in `data/verified_references.csv` with the
  DOIs cross-checked in `review/INPUT_STATE_CITATION_VERIFICATION.md`; QC re-verifies each DOI in
  manuscript and ledger. Terms remain provisional; no novelty claim.

## Figure 1

- Unchanged: ACCESS → RECONSTRUCT → EXECUTE → REPRODUCE → ROBUST; ACCESS sub-questions;
  Paper II highlight on RECONSTRUCT; ROBUST labelled future Paper III; Wet/Dry conceptual;
  drift note. Reviewed at journal width; no change required.

## QC (`verification_qc.py`)

New checks: no stale preparation wording; harness statement calibrated; Abstract completed-tense
with frozen/provisional separation and D pending; Section 4/5 labels; Table 10 code-availability
caveat; availability-chain sentence; no overclaim (guarantee, final adjudication, Paper III
analyses); Methods structure; changelog present with frozen hashes.

## Status

DRAFT — NOT SUBMISSION READY. Investigator verification/sign-off, human adjudication and Human
Validation D remain pending; `make submission-check` remains non-zero by design.
