import argparse
import json
import subprocess
from dataclasses import asdict
from datetime import datetime, timezone
from pathlib import Path

from paper2.core import sha256, snapshot
from paper2.isolation import DOCKER, PROBES, Worker

ROOT = Path(__file__).resolve().parents[2]
TAG = "paper2-science:qualification"
IMPORT_PROBE = """
import json
import Bio
import h5py
import matplotlib
import networkx
import numpy
import openpyxl
import pandas
import scipy
import sklearn
import statsmodels
import sympy
names = [Bio, h5py, matplotlib, networkx, numpy, openpyxl, pandas, scipy,
         sklearn, statsmodels, sympy]
print(json.dumps({module.__name__: module.__version__ for module in names},
                 sort_keys=True))
"""


def docker(*arguments: str) -> str:
    return subprocess.check_output(
        [DOCKER, *arguments], cwd=ROOT, text=True, timeout=300, env={"PATH": "/usr/bin:/bin"}
    ).strip()


def build(destination: Path) -> dict[str, object]:
    destination.mkdir(parents=True, exist_ok=False)
    docker("build", "--provenance=false", "-f", "Dockerfile.science", "-t", TAG, ".")
    image_id = docker("image", "inspect", "--format={{.Id}}", TAG)
    inspection = json.loads(docker("image", "inspect", image_id))
    snapshot(destination / "image-inspection.json", json.dumps(inspection, indent=2).encode())
    package = destination / "package"
    package.mkdir()
    snapshot(package / "allowed.txt", b"SYNTHETIC_ALLOWED")
    canary = destination / "host-canary.txt"
    snapshot(canary, b"SYNTHETIC_HOST_ONLY_NOT_MOUNTED")
    expected = {"allowed.txt": sha256(b"SYNTHETIC_ALLOWED")}
    worker = Worker(package, expected, image_id)
    try:
        imports = worker.execute(IMPORT_PROBE, seconds=30)
        probes = worker.execute(
            PROBES.replace(
                "HOST_CANARY_PATH",
                json.dumps(str(canary.resolve())),
            ),
            seconds=30,
        )
    finally:
        worker.close()
    snapshot(destination / "imports.json", json.dumps(asdict(imports), indent=2).encode())
    snapshot(destination / "probes.json", json.dumps(asdict(probes), indent=2).encode())
    lock = ROOT / "requirements-science.lock"
    report: dict[str, object] = {
        "scope": "synthetic_scientific_image_qualification_only",
        "status": (
            "scientific_image_smoke_pass"
            if imports.returncode == probes.returncode == 0
            and imports.stop_reason == probes.stop_reason == "completed"
            and all(json.loads(probes.output).values())
            else "scientific_image_smoke_failed"
        ),
        "image_id": image_id,
        "requirements_lock_sha256": sha256(lock.read_bytes()),
        "dockerfile_sha256": sha256((ROOT / "Dockerfile.science").read_bytes()),
        "imports": json.loads(imports.output),
        "execution_controls": json.loads(probes.output),
        "recorded_at_utc": datetime.now(timezone.utc).isoformat(),
        "pilot_authorized": False,
        "main_study_authorized": False,
        "limits": [
            "General Python libraries only; no paper-specific dependency closure",
            "No R, commercial software, GPU or external command-line scientific programs",
            "Synthetic import and containment checks are not scientific validation",
        ],
    }
    snapshot(destination / "report.json", (json.dumps(report, indent=2) + "\n").encode())
    return report


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    report = build(args.output)
    print(json.dumps(report, indent=2))
    if report["status"] != "scientific_image_smoke_pass":
        raise SystemExit(1)


if __name__ == "__main__":
    main()
