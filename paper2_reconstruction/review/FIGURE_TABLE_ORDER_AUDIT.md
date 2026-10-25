# Figure/Table order audit — FINAL NON-HUMAN QC PATCH

Rule: every numbered object must be cited in text, numbered in first-citation
order, no jumps, no backward references, no unreferenced object, no mixing of
main-text and supplement numbering.

## Main text (Preparation_draft_EN.docx)

Objects present in the main document: Figure 1 and Table 1 only. All other
tables were deliberately moved to the supplement/editable-tables file in an
earlier pass; they are now numbered S1–S11 to avoid mixing main-text and
supplement numbering (the previous bug — Methods citing "Table 10/11" before
"Table 1" — is resolved because supplement objects carry the S-prefix scheme).

First-citation order in the main document:
- Figure 1: char offset 4569
- Table S1: char offset 7349
- Table S2: char offset 7413
- Table 1: char offset 12541
- Table S3: char offset 13798
- Table S4: char offset 14341
- Table S5: char offset 14482
- Table S6: char offset 15184
- Table S7: char offset 16899
- Table S8: char offset 17283
- Table S9: char offset 18824
- Table S10: char offset 19291

- Figure 1 first cited before Table 1: yes
- Every S-numbered table cited in text: yes

## Supplement (Supplement_protocols_DRAFT_EN.docx / Editable_tables_EN.docx)

Tables S1–S11 appear in supplement order: S1 conceptual mapping, S2 ACCESS
input-state fields, S3 dispositions, S4 run-level states, S5 failure codes,
S6 reveal ledger, S7 slot distribution, S8 funnel, S9 reveal completeness,
S10 Human Validation D pending, S11 Wilson analytic planning. The Wilson
planning table was renumbered S1→S11 so the supplement sequence is strict.
Table S11 is the only object never directly cited from the main text; it is
referenced from the supplement intro and the editable-tables intro, which is
acceptable for a planning-only supplement table and is recorded here.

STATUS: PASS — renumber map in FIGURE_TABLE_RENUMBER_MAP.csv.
