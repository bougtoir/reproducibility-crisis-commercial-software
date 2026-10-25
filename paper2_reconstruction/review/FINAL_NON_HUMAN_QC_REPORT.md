# FINAL NON-HUMAN QC REPORT — formatting/numbering patch

| Gate | Result |
| --- | --- |
| HOSTILE REVIEW | PASS (F1–F16 closed previously; no new findings from this patch) |
| NUMERIC HARDCODE AUDIT | PASS (no canonical value touched; presentational renumber only) |
| PUBLIC CODE REPRODUCIBILITY | PASS (146 tests, make build green) |
| REFERENCE EXISTENCE | PASS (unchanged ledger verdicts) |
| ALL REFERENCES CITED | PASS (21/21 cited; proll_rauber_dynamic added at [13]) |
| ALL FIGURES/TABLES CITED | PASS (Table 1, Figure 1, S1–S10 cited; S11 referenced from supplement intro) |
| DATA–TEXT–FIGURE NUMERIC CONSISTENCY | PASS (re-verified post-renumber; zero value changes) |
| FIGURE/TABLE FIRST-CITATION NUMBERING | PASS (S-scheme map in FIGURE_TABLE_RENUMBER_MAP.csv) |
| REFERENCE FIRST-APPEARANCE NUMBERING | PASS (sequence 1–21 strict; CITATION_SEQUENCE_CHECK.csv) |
| WORD-NATIVE EQUATIONS (OMML) | PASS (29 OMML objects; 0 plain-text remnants) |
| MARKDOWN ASTERISK REMNANTS | PASS (0 formatting asterisks; only code/glob literals remain) |
| MANUAL CHARACTER-COUNT LINE WRAPPING | PASS (0; only intentional page breaks) |

Remaining human-dependent tasks (unchanged): investigator sign-off,
ethics determination, Human Validation D, submission.
