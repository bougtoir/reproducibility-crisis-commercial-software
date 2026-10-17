# Private evidence check (pre-D)

The private evidence package (`/home/ubuntu/paper2_evidence/…`) is required
for two outputs whose evidence references are package-relative:

- `results/verification_adjusted/evidence_reference_resolution.csv`
- `results/verification_adjusted/verification_import_audit.json`

Behavior without the mount: refs resolve as `evidence_package_not_mounted`.
To prevent public rebuilds from clobbering authoritative outputs,
`verification_layer` now writes these files **only when the evidence package
is mounted or the file does not yet exist**. Committed bytes = the
authoritative private-mount outputs (698 evidence refs, 0 unresolved).

This is an expected external dependency, not a software failure: the same
code produces the committed bytes when the package is mounted.

Status on this machine: evidence package not mounted; committed authoritative
outputs retained; all other checks (incl. QC references to the audit JSON)
run against them and pass.
