# Package self-consistency (revision pass)

## Check levels

- `PUBLIC_PACKAGE_CHECK` = `make setup && make fetch-vetted-wheels && make
  check` on a clean public checkout. Verified: ruff clean, mypy clean,
  **146 tests passed** (was 145/146 — the slot-package wheel test failed
  without a fetch path).
- `FULL_PRIVATE_EVIDENCE_CHECK` additionally requires the private evidence
  mount (`/home/ubuntu/paper2_evidence/`, not in this package). Two
  verification artifacts regenerate differently without it:
  `results/verification_adjusted/evidence_reference_resolution.csv` and
  `verification_import_audit.json` would contain
  `evidence_package_not_mounted`. The committed versions are the
  authoritative private-mount outputs and are restored, not regenerated, on
  public machines. This is an *omitted-private-evidence* difference, not a
  test failure and not a code defect.

## Fixes applied this revision

- `scripts/fetch_vetted_wheels.sh` + `make fetch-vetted-wheels`: downloads
  packaging-26.3, python-dateutil-2.9.0.post0, six-1.17.0 from PyPI into the
  gitignored `data/raw/vetted-wheels-20260924/` and verifies each sha256
  against `requirements.lock`; writes LEDGER.csv. The previously failing
  `test_slot_package_holds_only_publication_and_vetted_software` now passes.
- Multiprocessing/portability: tests use the library API, no spawn-dependent
  global state; the suite passes in the sandbox (no display/network beyond
  the pip fetch).
- Zip manifest: `results/preparation_artifact_manifest.json` regenerated;
  `scripts/` directory added to the shipped file list; root files
  INPUT_STATE_IDENTIFIABILITY_NOTE.md and REVISION_CHANGELOG.md already
  included.

## Not weakened

No test was removed, skipped, or relaxed. The QC layer remains a fail-closed
check suite (`make submission-check` still non-zero until D completes).
