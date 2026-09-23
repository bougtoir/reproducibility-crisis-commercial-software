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
the first four deterministic candidates per field (28 papers; private run
`data/raw/pilot-screen-20260923`, 261,555 provider-accounted tokens). For the 17
candidates without open PMC text, OpenAlex best-OA locations were checked and
retained (`data/raw/pilot-oa-locations-20260923`): one CC-BY publisher PDF was
retrieved (PMID 39774361); one hybrid-OA publisher PDF returned HTTP 403 to the
scripted request (PMID 33944686; not an access verdict); two reported OA without a
PDF location; thirteen are closed at every reported location. Provisional G1/G2
proposals are machine output bound to retained source segments and await human
adjudication; no pilot case has been selected and the funnel is unchanged.
