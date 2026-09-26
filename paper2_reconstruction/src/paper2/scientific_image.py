import argparse
import json
import subprocess
from dataclasses import asdict
from datetime import datetime, timezone
from pathlib import Path

from paper2.core import sha256, snapshot
from paper2.isolation import DOCKER, PROBES, Execution, Worker
from paper2.model_api import mapping

ROOT = Path(__file__).resolve().parents[2]
TAG = "paper2-science:qualification"
IMPORT_PROBE = """
import hashlib
import json
import subprocess
from pathlib import Path

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
libraries = {module.__name__: module.__version__ for module in names}
executables = {}
for name in ("clustalw", "iqtree2"):
    path = Path("/usr/bin") / name
    executables[name] = {
        "sha256": hashlib.sha256(path.read_bytes()).hexdigest(),
        "bytes": path.stat().st_size,
    }
packages = subprocess.check_output(
    ["dpkg-query", "-W", "-f=${binary:Package}\\t${Version}\\n"], text=True
)
print(json.dumps({"libraries": libraries, "executables": executables,
                  "installed_system_packages": packages}, sort_keys=True))
"""
PHYLOGENY_PROBE = """
import json
import subprocess
from pathlib import Path

from Bio import Phylo

alignment = subprocess.run(
    ["clustalw", "-INFILE=/input/synthetic.fasta", "-OUTFILE=/work/aligned.fasta",
     "-ALIGN", "-NEWTREE=/work/guide.dnd", "-OUTPUT=FASTA", "-TYPE=DNA", "-QUIET"],
    capture_output=True, text=True, timeout=20
)
Path("/work/clustalw.stdout").write_text(alignment.stdout)
Path("/work/clustalw.stderr").write_text(alignment.stderr)
alignment.check_returncode()
tree = subprocess.run(
    ["iqtree2", "-s", "/work/aligned.fasta", "-m", "JC", "-seed", "17",
     "-nt", "1", "-pre", "/work/synthetic"],
    capture_output=True, text=True, timeout=20
)
Path("/work/iqtree.stdout").write_text(tree.stdout)
Path("/work/iqtree.stderr").write_text(tree.stderr)
tree.check_returncode()
observed = Phylo.read("/work/synthetic.treefile", "newick")
print(json.dumps({
    "alignment_written": Path("/work/aligned.fasta").is_file(),
    "expected_taxa": sorted(tip.name for tip in observed.get_terminals())
                     == ["alpha", "beta", "delta", "gamma"],
    "finite_nonnegative_branches": all(
        node.branch_length is None or 0 <= node.branch_length < float("inf")
        for node in observed.find_clades()
    ),
}))
"""


def docker(*arguments: str) -> str:
    return subprocess.check_output(
        [DOCKER, *arguments], cwd=ROOT, text=True, timeout=300, env={"PATH": "/usr/bin:/bin"}
    ).strip()


def probe_output(execution: Execution) -> dict[str, object]:
    try:
        return mapping(json.loads(execution.output))
    except ValueError:
        return {"parse_error": "Probe did not return a JSON object"}


def build(destination: Path) -> dict[str, object]:
    destination.mkdir(parents=True, exist_ok=False)
    docker("build", "--provenance=false", "-f", "Dockerfile.science", "-t", TAG, ".")
    image_id = docker("image", "inspect", "--format={{.Id}}", TAG)
    inspection = json.loads(docker("image", "inspect", image_id))
    snapshot(destination / "image-inspection.json", json.dumps(inspection, indent=2).encode())
    package = destination / "package"
    package.mkdir()
    snapshot(package / "allowed.txt", b"SYNTHETIC_ALLOWED")
    sequences = {
        "alpha": "ACGT" * 25,
        "beta": "ACGT" * 24 + "ACGA",
        "gamma": "ACGA" * 25,
        "delta": "TCGA" * 25,
    }
    fasta = "".join(f">{name}\n{sequence}\n" for name, sequence in sequences.items()).encode()
    snapshot(package / "synthetic.fasta", fasta)
    canary = destination / "host-canary.txt"
    snapshot(canary, b"SYNTHETIC_HOST_ONLY_NOT_MOUNTED")
    expected = {"allowed.txt": sha256(b"SYNTHETIC_ALLOWED"), "synthetic.fasta": sha256(fasta)}
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
        phylogeny = worker.execute(PHYLOGENY_PROBE, seconds=50)
        artifacts = worker.export_artifacts(destination / "worker-artifacts")
    finally:
        worker.close()
    snapshot(destination / "imports.json", json.dumps(asdict(imports), indent=2).encode())
    snapshot(destination / "probes.json", json.dumps(asdict(probes), indent=2).encode())
    snapshot(destination / "phylogeny.json", json.dumps(asdict(phylogeny), indent=2).encode())
    lock = ROOT / "requirements-science.lock"
    controls = probe_output(probes)
    phylogeny_checks = probe_output(phylogeny)
    report: dict[str, object] = {
        "scope": "synthetic_scientific_image_qualification_only",
        "status": (
            "scientific_image_smoke_pass"
            if imports.returncode == probes.returncode == phylogeny.returncode == 0
            and imports.stop_reason == probes.stop_reason == phylogeny.stop_reason == "completed"
            and controls and all(value is True for value in controls.values())
            and phylogeny_checks and all(value is True for value in phylogeny_checks.values())
            else "scientific_image_smoke_failed"
        ),
        "image_id": image_id,
        "requirements_lock_sha256": sha256(lock.read_bytes()),
        "dockerfile_sha256": sha256((ROOT / "Dockerfile.science").read_bytes()),
        "imports": probe_output(imports),
        "execution_controls": controls,
        "synthetic_phylogeny": phylogeny_checks,
        "worker_artifacts": artifacts,
        "recorded_at_utc": datetime.now(timezone.utc).isoformat(),
        "pilot_authorized": False,
        "main_study_authorized": False,
        "limits": [
            "Listed libraries and phylogenetic tools only; dependency closure still unverified",
            "No R, commercial software or GPU",
            "System package inventory recorded; transitive apt versions are not fully locked",
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
