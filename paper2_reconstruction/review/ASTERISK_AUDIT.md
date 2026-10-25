# Asterisk audit — FINAL NON-HUMAN QC PATCH

Requirement: no Markdown emphasis markers may remain in Word deliverables.

- Total `*` occurrences in the three DOCX deliverables: 4.
- Classification: all remaining occurrences are inside backticked code spans —
  glob/file patterns (`*_*.json`, `leakage_*`, `paper2_evidence*`) or the
  verbatim quoted header `` `**DRAFT — NOT FROZEN**` ``. Classification:
  legitimate_code_or_filename. None are equation operators (the sole `*` math
  operator, in the rounding-bin expression, is now OMML multiplication) and
  none are formatting remnants.
- MARKDOWN_FORMATTING_ASTERISKS = 0 (all `**`/`x*` emphasis spans are now
  rendered as real bold/italic runs, including spans crossing code spans and
  hard-wrapped source lines).

STATUS: PASS — see ASTERISK_OCCURRENCES.csv.
