# Line-break audit — FINAL NON-HUMAN QC PATCH

Requirement: no manual character-count line wrapping inside DOCX text.

- Hard-wrapped source lines in protocols/*.md are now joined into single
  paragraphs by the renderer (markdown() block-joining), so no manual
  text-wrapping breaks exist.
- w:br counts: {'Preparation_draft_EN.docx': {}, 'Supplement_protocols_DRAFT_EN.docx': {'page': 16}, 'Editable_tables_EN.docx': {}}. All 16 breaks in the supplement are type="page"
  (intentional page breaks between protocol documents), not text-wrapping.
- MANUAL_CHARACTER_COUNT_LINE_WRAPPING = 0.

STATUS: PASS.
