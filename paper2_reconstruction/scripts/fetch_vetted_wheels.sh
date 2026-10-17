#!/usr/bin/env bash
# Download the pinned free-software closure used by the firewall harness tests
# into data/raw/vetted-wheels-20260924/ and verify every wheel against the
# sha256 hashes recorded in requirements.lock. Public PyPI artifacts only;
# data/raw/ is gitignored, so each checkout must run this once.
set -euo pipefail
cd "$(dirname "$0")/.."
PYTHON=${PYTHON:-.venv/bin/python}
DEST=data/raw/vetted-wheels-20260924
mkdir -p "$DEST"
"$PYTHON" -m pip download --no-deps -q -d "$DEST" \
    packaging==26.3 python-dateutil==2.9.0.post0 six==1.17.0
"$PYTHON" - <<'EOF'
import hashlib, re, sys
from pathlib import Path
lock = Path("requirements.lock").read_text()
allowed = set(re.findall(r"--hash=sha256:([0-9a-f]{64})", lock))
dest = Path("data/raw/vetted-wheels-20260924")
rows = []
for wheel in sorted(dest.glob("*.whl")):
    digest = hashlib.sha256(wheel.read_bytes()).hexdigest()
    if digest not in allowed:
        sys.exit(f"{wheel.name}: sha256 {digest} not pinned in requirements.lock")
    rows.append(f"{wheel.name},{digest}")
(dest / "LEDGER.csv").write_text("file,sha256\n" + "\n".join(rows) + "\n")
print(f"verified {len(rows)} pinned wheels in {dest}")
EOF
