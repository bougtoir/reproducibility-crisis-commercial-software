# Public package check (pre-D)

Clean-checkout procedure and verified result:

```sh
make setup                  # hash-locked deps into .venv
make fetch-vetted-wheels    # 3 pinned wheels, sha256 vs requirements.lock
make build                  # build -> documents -> verification_qc (all PASS)
make check                  # ruff + mypy + pytest
```

Result: ruff clean; mypy clean; **146 tests passed** (14 warnings, all
dependency deprecations). `make submission-check` fails by design (D +
sign-off pending).

The public check exercises the full pipeline except private-evidence
resolution (see PRIVATE_EVIDENCE_CHECK.md). No test skipped or weakened.
