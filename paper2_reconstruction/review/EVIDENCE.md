# Evidence handoff and acquisition limits

The preparation workflow ran five read-only audits on separate VMs; all completed.
They did not run Paper II experiments. Raw evidence was handed back as attachments,
downloaded and verified in the parent session. This is an internal evidence archive
inventory; these URLs require organization access and are not a public-data licence.

| Audit | Preserved evidence archive |
|---|---|
| Corpus/VOR/identifiers | https://app.devin.ai/attachments/ac70ad61-14b3-4024-a6bc-ed1b59d82de0/paper2_corpus_audit-evidence.tar.gz |
| Terminology/literature | https://app.devin.ai/attachments/e962f30f-fe37-457c-88ec-df4657338d38/paper2_terminology_audit_20260922T113608Z.zip |
| Sampling/SAP | https://app.devin.ai/attachments/c5e9150c-19f6-4d5d-98e9-3a6bede7b8a7/paper2_sap_audit_20260922T1129.tar.gz |
| Information firewall | https://app.devin.ai/attachments/b50c3f4a-019c-4ca4-abfd-7c5856247505/paper2-firewall-audit-20260922T1128Z/paper2-firewall-audit-20260922T1128Z.zip |
| Journal requirements | https://app.devin.ai/attachments/36ad4b21-1668-452e-a5d2-5c11452c61d5/paper2_journal_audit_20260922.tar.gz |

Each archive includes source URLs, retrieval conditions and UTC dates, original
response bytes, checksums, rights and completeness assessment. Failed requests
are preserved as failures. Some publisher-origin requests returned client
challenges; extracted official content is distinguished from complete publisher
HTML. Wet antecedent full text was not accessible; only metadata/preview claims
are supported. A targeted literature search is not a systematic review.

Independent corpus audit located the VOR's repository link and matching principal
aggregates, but no corrected unique-paper corpus or historical validation
annotations in inspected releases/refs/Crossmark. Earlier Scientific Data draft
validation claims were not substantiated by deposited annotations and are not
imported. The EPJ VOR does not report those draft validation figures.

The preparation build requires only the pinned public source commit; it does not
need any internal archive. Public `data/verified_references.csv` preserves the
primary-source URLs, publication/version status, supported claims and caveats.
Policy and literary claims must be rechecked from those public primary sources
before the eventual submission. Do not redistribute the raw evidence wholesale.

## Corpus article acquisition (2026-09-22/23)

`results/corpus_article_acquisition_audit.json` is a copy of the retained
post-acquisition audit of the private `data/raw/corpus-articles-20260922`
directory. Acquisition was interrupted once at 7,680 indexed rows (the interrupted
index is retained as `article_index_interrupted_7680rows_20260922.csv`) and then
resumed from the immutable receipts without re-fetching or overwriting any body;
the final index covers the exact ordered 9,935-paper frame. Only the open Europe
PMC JATS endpoint was assessed: 3,966 identity-verified articles, 5,957 papers
with no PMC endpoint selected (other sources unassessed), 11 persistent HTTP 500
responses and 1 body-less JATS response. A second attempt for those 12 papers
(`data/raw/corpus-articles-retry-20260923`, first attempt retained) reproduced
the same server-side results. None of these states is a G1–G5 verdict, an
input-availability finding or a full-text-availability finding for the corpus.
The author confirmed on 2026-09-23 that no local full text exists for the three
closed-access pilot candidates (PMID 38007008, 41336033, 38950848); they remain
unresolved and no replacement candidate was drawn
(`data/raw/candidate-alternatives-20260922/author_disposition_20260922.json`).
Article bodies are third-party content and are not redistributed here.

## Pilot candidate screening (2026-09-23)

`results/pilot_candidate_screening.json` summarizes `paper2.funnel_screen` over
the first nine deterministic candidates per field (63 papers; private run
`data/raw/pilot-screen-20260923`, 472,329 provider-accounted tokens). Nine ranks
were needed because Environmental_Earth has no open PMC text before rank 9. For
the 41 candidates without open PMC text, OpenAlex best-OA locations were checked
and retained (`data/raw/pilot-oa-locations-20260923`): two licensed publisher PDFs
were retrieved (CC-BY, CC-BY-NC), three OA PDF locations returned HTTP 403 to the
scripted request (not access verdicts), ten report OA with a landing page only
(text not retained), and 26 are closed at every reported location. Provisional
G1/G2 proposals are machine output bound to retained source segments;
`results/pilot_adjudication_request.md` lists the first lawfully retained
candidate per field for human G1–G5 adjudication. No pilot case has been selected
and the funnel is unchanged.

## Delegated primary G1–G5 adjudication (2026-09-24)

On author instruction the first-pass G1–G5 adjudication was recorded by Devin
instead of the investigator, as deviation `DEV-2026-09-23-01`
(`data/adjudication/devin_primary_G1_G5_20260923.json`). Eleven candidates were
assessed against the hash-checked retained article bytes;
`paper2.primary_adjudication` re-validates every gate value, requires each cited
quote to occur verbatim inside the cited retained segment, and requires each
stratum to consume its deterministic order without gaps
(`results/pilot_primary_adjudication.json`, `results/pilot_primary_adjudication.csv`).
Three candidates were skipped for primary G1=no and one for unresolved G1/G2;
candidates without lawfully retained text remain unadjudicated rather than
negative, and their OA route evidence for ranks 10–13 is retained in
`data/raw/pilot-oa-locations-ranks10-13-20260923` (PMID 38252604 has no DOI in the
frame, so its route check is incomplete). G3/G4/G5 never caused replacement: the
provisional set retains restricted-input, commercial-software and
insufficient-specification cases as study outcomes. The screening order was
extended to 13 ranks per field only to reach the next lawfully retained
Social_Behavioral candidate (`results/pilot_candidate_screening.json`,
`extension_20260923_to_13_ranks`). These gates are provisional pending investigator
verification; the funnel remains `NOT_STARTED`, pilot selection is not frozen, and
no reconstruction has been attempted.

## Information-firewall control qualification (2026-09-24)

`paper2.firewall` separates custodian, broker, solver and reveal-reviewer roles
around one quarantined mixed archive holding the seven provisional pilot articles
(hash-checked retained open PMC JATS), a pinned free-software closure
(`packaging`, `python-dateutil`, `six`; hashes from `requirements.lock`, wheels
and ledger retained in `data/raw/vetted-wheels-20260924`, closure verified from
wheel `Requires-Dist` metadata) and two synthetic forbidden members standing in
for the original implementation and author notes. The broker defaults to deny,
serves only reviewed exact members to named actors, and writes a hash-chained
event for every request. Adversarial cases (blocked member, reveal member to a
solver, reveal member before freeze, unlisted item, unauthorised actor, tampered
custodian bytes in tests) were all denied and journalled. Inside the isolated
worker the solver saw only the served members, could not reach the host canary,
the archive, the network, the Docker socket or API credentials, imported the
closure offline from the served wheels, and a runaway step was stopped at the
wall limit. The sealed control outcome received an RFC 3161 receipt from the
pinned FreeTSA anchor before the reveal member was released to the reveal
reviewer only; a mutated outcome no longer matched the receipt
(`results/firewall_qualification.json`, all nine controls pass; the full run,
including served article bytes, is retained privately in
`/home/ubuntu/paper2_evidence/firewall-qual-20260924b`). Still unqualified:
model-side retrieval tool restriction for provider-hosted tools, three-slot
scheduling and budget comparability, and independent human adjudicator
governance. This is a control test on one candidate article, not a
reconstruction attempt; no pilot or main run is authorized.

## Lawful public input acquisition for the provisional pilot (2026-09-24)

`paper2.pilot_inputs` resolved the deposit route named in each provisional pilot
article and retained every listing response and every retrievable file under
`data/raw/pilot-inputs-20260924/` with URL, accession, UTC time, request
conditions, byte count, SHA-256, rights and completeness in a content-addressed
ledger. Four papers name no public route: ADNI requires a data-use application,
two state author request only, and one points to supplementary material only.
ENA `PRJEB51705` and NGDC `PRJCA012518` announce only sequencing archives far
above the 64 MiB retention cap. ProteomeXchange `PXD027610` is announced through
iProX and exposes no file sizes, so its completeness is recorded as
`service_announces_no_file_sizes` rather than assumed. Zenodo record
`10.5281/zenodo.6377228` yielded eleven retained tabular files; these are the
figure source tables of the target statistic itself, so they are held as raw
evidence for post-reveal validation and were withheld from every solver as
outcome leakage. No deposited study input was simultaneously public, below the
cap, machine-readable and free of outcome leakage. Route failures are recorded as
route observations, not as G3 input verdicts.

## Provisional pilot execution (2026-09-24)

Executed under deviation `DEV-2026-09-24-02`
(`data/adjudication/pilot_execution_DEV-2026-09-24-02.json`) on the principal
investigator's explicit authorisation while investigator verification of G1–G5 remains pending. Seven
provisional papers, one per stratum, three independent slots each. Every slot
received its own broker journal, its own served package (the hash-checked article
plus the pinned free-software closure) and a fresh isolated worker with no shared
mutable state and no parent or sibling debugging. Target and allowed manifests
were hashed and RFC 3161 timestamped before the first solver call, and the sealed
blind outcomes were timestamped before any reveal-class request; no reveal and no
adjudication were performed (`results/provisional_pilot_summary.json`; full run
evidence, which quotes restricted article text, is retained privately in
`/home/ubuntu/paper2_evidence/pilot-20260924b`).

A first harness iteration (`/home/ubuntu/paper2_evidence/pilot-20260924`) is
retained rather than discarded: twenty of its twenty-one slots stopped on harness
limits rather than on study evidence, so it measures the harness, not the papers.
The harness was then changed — full tool output retained but truncated in the
model context, an oversized print no longer terminating the worker, a truncated
or malformed action returned to the solver within a fixed budget, and a larger
run envelope — and all twenty-one slots were re-run. In that iteration twelve
slots returned a sealed report and nine exhausted the context budget. No slot
executed the target computation, because no deposited input was served; every
sealed report recorded a null observed value with explicit missing-input failure
codes, and none manufactured a value or an agreement. L1–L5 outcomes remain
unadjudicated, pilot runs are excluded from the main sample, and no success rate
is assessable.

## Second-pass gate verification, accessibility layer and amended criterion (2026-09-24)

On investigator instruction the provisional pilot papers were removed from the main
reconstruction sample before any adjudication. `paper2.second_pass` re-read all
37 retained quotes of the eleven primary G1–G5 assessments against the
hash-checked segments and confirmed every gate without correction
(`results/second_pass_G1_G5.json`); this is a delegated re-verification, and investigator
verification of G1–G5 remains `pending` in every record. Protocol amendment
`AMEND-2026-09-24-03` (`protocols/AMENDMENT_AMEND-2026-09-24-03_deposited_data_criterion.md`,
SHA-256 `e4c8f82d…74d23`) was hash-frozen and RFC 3161 timestamped
(`data/adjudication/amendment_AMEND-2026-09-24-03.json`, receipt directory
alongside) before any candidate was screened under it. It requires that a third
party can obtain the deposited analysis data lawfully and immediately through an
anonymous machine-readable listing, that every needed file has an announced size
within the 2 GiB / 8 GiB caps, and that the deposit is not itself the target
output. The seven provisional papers are recorded as the
`ACCESSIBILITY_GATE_FAILED` layer (`data/adjudication/accessibility_gate_failed_20260924.json`:
above-cap/output-only, size-unannounced, application-required,
listing-unavailable, author-request-only ×2, supplement-only-unverified), a
pre-reproducibility access-barrier audit sample for an auxiliary analysis or a
separate paper. Its validator rejects any outcome or rate field: no member
reached a reconstruction attempt, so no success or failure rate exists for them.

## Amended candidate selection (2026-09-24)

`paper2.deposit_screen` screened the ordered frame to 400 ranks per stratum
(2,800 candidates; Stage A `stage_a_statements_400.json`): 1,688 without lawfully
retained text, 409 naming no deposit, 248 author-request-only, 166
supplement-only, 61 code-only, 14 application-required and 214 naming an open
route. Stage B resolved open routes stratum by stratum in deterministic order
until one candidate passed the caps (430 candidates recorded with their
exclusion class; 7 route-eligible; `stage_b_routes.json`, every listing response
retained with URL, UTC time, bytes and SHA-256 under
`data/raw/deposit-screen-20260924/routes/`). Devin assessed each route-eligible
candidate's G1–G5 against the retained article segments and the custodian
inspected the structure of every deposit small enough to retrieve (private
receipts in `/home/ubuntu/paper2_evidence/deposit-inputs-20260924`; no analysis
run, no values compared, nothing served to any solver). GEO `GSE291941`
(PMID:41466177, Chemistry_Materials) bundles raw counts with log2FC, p-value and
FDR columns for the target comparison and was excluded as
`deposited_output_only`; the Chemistry_Materials chain was then resumed from
rank 42 with that exclusion carried as a prior class
(`stage_b_routes_continuation_chemistry.json`, hash-chained to Stage B), reaching
GEO `GSE186841` (PMID:35003117, two per-sample count tables, no result columns)
at rank 46. `paper2.amended_selection` re-validates every quote, rank, route
class and leakage class and writes `results/amended_candidate_selection.{csv,json}`:
one main-sample candidate in each of the seven strata (four unconditional; the
Biomedical_Basic target requires investigator pre-specification and the
Computational_Science candidate an input-completeness check before freeze).
Dryad file downloads for two candidates required a browser JavaScript challenge
but no account, payment or contact, and were retained anonymously; this route
condition is recorded. These dispositions are sampling-frame decisions from a
delegated record pending investigator verification; the funnel is not updated, no
target manifest is frozen, no reconstruction has started, and the summary
carries no rate.

## Prospective pilot: freeze, execution, procedural review and protocol freeze (2026-09-25)

Under AMEND-2026-09-25-04 the seven amended candidates became the
`PROSPECTIVE_PILOT` set (outside the primary denominator). `paper2.prospective_freeze`
wrote the ten-field specification for each (target as comparison criterion only,
target location with retained quotes, required inputs with per-file SHA-256 and
bytes, allowed resources, blocked original artifacts, metric, agreement rule,
conclusion rule, ceiling, stopping rule) together with delegated G1–G5 second-pass
states (`pilot_prospective_freeze_20260925.json`, SHA-256
`43cc2ec8…8b34f`, RFC 3161 receipt `pilot-freeze-20260925_receipt/`). All inputs
were lawful anonymous deposits or public references persisted under
`/home/ubuntu/paper2_evidence/pilot-inputs-20260925` (ENA PRJNA641521: 192 files,
PRJNA827817: 14 files; GEO GSE186841; Zenodo 17343533 filtered to raw `.txt`
only; Dryad dr7sqv9z5 and z8w9ghxfh; Figshare 15148851; SILVA 128, UNITE 7.2,
SARS-CoV-2 reference/annotation, ARTIC V4.1, constellations v0.1.3). Target-bearing
files were excluded as leakage and recorded. `paper2.pilot_run` executed exactly
three sequential blind slots per paper on one VM (21 slots, `paper2-science`
image, no network, non-root, capability-dropped, blind instruction carries the
`blind_target` not the reported value) and sealed the outcome before any reveal
(`pilot_execution_PILOT-FREEZE-2026-09-25/blind_outcome.json`, SHA-256
`485967ff…b1f3db`, timestamped). `paper2.pilot_review` then produced the
prespecified procedural review (`pilot_procedural_review_20260925.json`): 16/21
slots stopped at the 400,000 cumulative-token ceiling after a median of 23.5 tool
calls and ≤139 s (PD-01), free-text failure codes (PD-02), one report without any
execution (PD-03) and reported values without exported artifacts while the served
publication displays the value (PD-04). No value was adjudicated and no rate
computed. AMEND-2026-09-25-05 fixes the dispositions (token ceiling 1,500,000 for
consistency with the unchanged 40-call ceiling; taxonomy-only failure codes;
`no_execution_report` state; value-provenance requirement for L3/L4) and
`paper2.protocol_freeze` hashed PROTOCOL, SAP, TOLERANCE_RULES, FAILURE_TAXONOMY,
TARGET_SELECTION_RULES, INFORMATION_FIREWALL, REVEAL_PROTOCOL, the amendments and
the logging-schema modules (`protocol_freeze_20260925.json`, SHA-256
`f3bdc798…4ef0f`, timestamped).

## Eligible population and main-cohort selection (2026-09-25)

Stage A was run over the whole 9,935-paper frame (`stage_a_statements_full.json`;
819 papers name an open route) and Stage B exhaustively resolved every open route
(1,355 route resolutions, listings retained under
`/home/ubuntu/paper2_evidence/deposit-routes-full-20260925`; summary
`results/deposit_route_screen_full_20260925.{json,csv}`). Attrition to the eligible
reconstruction population (461): no lawful text 5,969; no deposit named 1,472;
author-request-only 827; supplement-only 569; code-only 270; above cap 143;
size unannounced 87; listing unavailable 68; application required 52; output-only
2; input mismatch 1; accessibility layer 7; prospective pilot 7. `paper2.main_selection`
drew the SAP cohort (N=100; N_h/n_h Biomedical 131/29, Computational 112/24,
Environmental 83/18, Clinical 51/11, Chemistry 35/8, Social 29/6, Physics 20/4;
within-field order SHA-256(`paper2-main-v1|paper_id`); π_h and weights saved) with
no outcome read (`main_cohort_20260925.json`, SHA-256 `c9370893…86b4f`,
timestamped). No main-study paper has been G1–G5 assessed, target-frozen,
input-acquired or reconstructed; that is the next stage.
