# WITHOUT_D_STATUS — complete without D versus D-dependent

Status: NOT SUBMISSION READY. Human validation (Section D) has not been performed.
The A/B/C/E rows are an AI-assisted independent evidence review, provisional, for investigator approval. They are not human-participant validation
and carry no investigator sign-off.

## Complete without D (regenerated from `results/`)

- Attempt reachability: 10/100 = 10.0% (Wilson 95% CI 5.5–17.4%).
- Conditional majority success among attempted papers — frozen mechanical: 0/10 = 0.0% (Wilson 95% CI 0.0–27.8%); verification-adjusted: 1/10 = 10.0% (Wilson 95% CI 1.8–40.4%).
- Observed end-to-end yield under the frozen workflow — frozen: 0/100 = 0.0% (Wilson 95% CI 0.0–3.7%); verification-adjusted: 1/100 = 1.0% (Wilson 95% CI 0.2–5.4%).
- 100-paper funnel (frozen / adjusted / unresolved envelope): `results/verification_adjusted/adjusted_B_funnel.csv`. Adjusted non-executed states: 83 verified barriers, 5 reclassified `specifiable_runnable` (not executed; attrition interpretation only), 2 unknown.
- Leakage detector audit: 1 confirmed, 7 corrected, 2 unresolved of 10 `all_inputs_leak_target` flags.
- Run-level L1–L5 (frozen and adjusted), failure taxonomy and stop reasons: `results/main_run_levels.csv`, `results/main_failure_codes.csv`, `results/verification_adjusted/verification_import_audit.json`.
- Descriptive reveal completeness (frozen and conservative adjusted labels): `results/verification_adjusted/adjusted_E_reveal_ledger.csv`.
- Sealed blind outcome unchanged: SHA-256 `350a0e7dcedbb74ccad2472c642318c87cc7d272f1b6c88186a1419dc858ea5a`.
- Frozen primary adjudication unchanged: SHA-256 `6fd42b8964e055606cbd24ebddb44f42fa09147fa80bfa2d5bcaa161766c5b28`.

## Not performed (by instruction)

No blind slot rerun; no execution of reclassified papers; no original-code execution;
no reveal-informed rescoring; no Paper III analyses; no investigator sign-off recorded.

## D-dependent (pending, external)

- Human-validation Results text, Table 9 and figure: empty/pending (`D_human_validation_cases_completed = 0`).
- Any claim about human reconstructability or agent-versus-human comparison.
- Conversion of the AI-assisted A/B/C/E proposals into signed investigator adjudication.
- Institutional ethics determination, validator recruitment and participation records.
- Submission readiness (`make submission-check` remains non-zero).
