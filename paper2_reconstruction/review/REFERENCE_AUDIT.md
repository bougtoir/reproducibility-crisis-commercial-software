# Reference audit (revision pass)

- Manuscript cites 21 references, each verified against
  `data/verified_references.csv` (28 verified records; 7 verified-but-uncited
  rows are retained for future Paper III and are not rendered).
- Every cited reference has a DOI or canonical URL, a recorded
  `publication_status`, a `supported_claim` string and an
  `applicability_caveat` stating where the external number may NOT be read
  into our claims (e.g. PaperBench's 21.0% is an average graded replication
  score, not a reconstruction proportion — the manuscript does not equate
  them).
- Claim↔reference mapping: `results/CITATION_CLAIM_MAP.csv` (generated from
  the verified ledger — no hand-copied claims).
- No fabricated references added this revision; publication statuses that
  could not be independently verified are labeled as such in the ledger
  (e.g. CORE-Bench "separate peer-reviewed venue not verified").
- Corpus-side evidence references: 698 total, 0 unresolved
  (`evidence_refs_total` / `evidence_refs_unresolved` canonical rows) —
  unchanged this revision.


## Addendum — numeric citation conversion (FINAL NON-HUMAN QC PATCH)

Author–year citations were converted to numeric [1]–[21] in first-appearance
order (EPJ numeric square-bracket default). proll_rauber_dynamic, previously
uncited, is now cited as [13] in the §7.1 cluster [11–14]. Reference list is
numeric; all 21 entries cited; all citations resolve. See
REFERENCE_ORDER_AUDIT.md, REFERENCE_RENUMBER_MAP.csv, CITATION_SEQUENCE_CHECK.csv.
Existence/status verdicts from the original audit are unchanged.
