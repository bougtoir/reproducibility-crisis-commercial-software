# Investigator sign-off package (pre-D)

Purpose: convert the PROPOSED AI-assisted independent-review records into
investigator-adjudicated records. **No approval has occurred.** The
investigator_* columns are empty by design.

## Contents
- `A_30slots_signoff.csv` (30), `B_90papers_signoff.csv` (90),
  `C_G1G5_pilot_signoff.csv` (25), `E_reveal_signoff.csv` (130)
- Per-layer summaries `*_SIGNOFF_SUMMARY.md`
- `INVESTIGATOR_SIGNOFF_FORM.docx` — printable attestation
- `INVESTIGATOR_SIGNOFF_CHECKLIST.xlsx` — editable workbook

## Procedure
1. Review each row against its `evidence_ref` (private evidence package
   required for package-relative refs).
2. Enter approve / reject / modify / pending plus comment, initials,
   UTC datetime.
3. Sign the DOCX form. Return the completed CSVs; the verification layer
   is then updated to `investigator_signed` status.

This is not Human Validation D (separate, still pending).
