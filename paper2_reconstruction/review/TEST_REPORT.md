# Test report (revision 2026-10 pass)

Environment: clean checkout + `make setup` + `make fetch-vetted-wheels`
(three pinned wheels, sha256-verified against requirements.lock).

```
ruff check src tests ... clean (auto-fix: import sort in input_state.py)
mypy ... clean
pytest -q ... 146 passed, 14 warnings (matplotlib/pyparsing deprecations only)
```

Previous state before this revision: 145 passed / 1 failed —
`test_slot_package_holds_only_publication_and_vetted_software` failed on clean
checkouts because the gitignored vetted-wheels closure had no fetch path.
Resolved by `make fetch-vetted-wheels`; no test was weakened or skipped.

Build QC: `paper2.verification_qc` all PASS (see
review/VERIFICATION_ADJUSTED_QC.md). `make submission-check` remains non-zero
by design: Human Validation D and investigator sign-off are pending.
