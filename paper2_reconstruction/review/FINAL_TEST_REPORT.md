# Final test report (pre-D finalization)

Environment: clean-ish checkout, `.venv` from `make setup`, wheels via
`make fetch-vetted-wheels` (packaging-26.3, python-dateutil-2.9.0.post0,
six-1.17.0 — sha256 verified against requirements.lock).

```
ruff check src tests  -> clean
mypy                  -> clean
pytest -q             -> 146 passed, 14 warnings (matplotlib/pyparsing deprecation)
paper2.verification_qc -> all checks PASS
```

Suite covers: corpus audit, cohort construction, sealed-slot integrity,
firewall harness, pilot, verification layer import/QC, canonical value
emission, document generation, manifest bytes consistency, schema validation.
Nothing hidden; no expected failures beyond `make submission-check`
(intentionally non-zero: D + investigator sign-off pending).

Private-evidence-dependent artifacts: see PRIVATE_EVIDENCE_CHECK.md.
