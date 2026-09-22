import argparse
import io
import json
import re
import stat
import sys
import zipfile
from datetime import datetime, timedelta
from pathlib import Path, PurePosixPath

from paper2.core import Row, read_csv, sha256, snapshot

ROLES = {
    "publication",
    "methodological_supplement",
    "public_input",
    "general_reference",
    "open_source_documentation",
}
MAX_SOURCE_BYTES = 128 * 1024 * 1024
MAX_SERVED_BYTES = 64 * 1024 * 1024
MAX_PACKAGE_BYTES = 256 * 1024 * 1024


def retained_file(root: Path, name: str, digest: str, size: str, limit: int) -> bytes:
    path = Path(name)
    if path.is_absolute() or ".." in path.parts or not path.parts:
        raise ValueError("Custody paths must be relative and remain within the evidence root")
    current = root
    for part in path.parts:
        current = current / part
        if current.is_symlink():
            raise ValueError("Custody paths must not contain symbolic links")
    if not current.is_file() or not size.isdigit():
        raise ValueError("Custody manifest requires an existing file and byte size")
    if not 0 < current.stat().st_size == int(size) <= limit:
        raise ValueError("Custody byte size differs or exceeds the package bound")
    body = current.read_bytes()
    if sha256(body) != digest:
        raise ValueError("Custody content differs from its reviewed SHA-256")
    return body


def zip_member(source: bytes, name: str) -> bytes:
    path = PurePosixPath(name)
    if path.is_absolute() or ".." in path.parts or not path.parts or "\\" in name:
        raise ValueError("Unsafe archive member name")
    with zipfile.ZipFile(io.BytesIO(source)) as archive:
        members = archive.infolist()
        names = [member.filename for member in members]
        if len(names) > 1000 or len(names) != len(set(names)):
            raise ValueError("Archive has excessive or duplicate members")
        member = archive.getinfo(name)
        kind = stat.S_IFMT(member.external_attr >> 16)
        if member.is_dir() or kind not in {0, stat.S_IFREG}:
            raise ValueError("Only regular reviewed archive members can be served")
        if not 0 < member.file_size <= MAX_SERVED_BYTES:
            raise ValueError("Archive member exceeds the serving bound")
        return archive.read(member)


def verify_review(row: Row, paper_id: str, policy_version: str) -> None:
    if (
        row["paper_id"] != paper_id
        or row["policy_version"] != policy_version
        or row["decision"] != "allow"
        or row["role"] not in ROLES
        or row["completeness"] != "complete"
    ):
        raise ValueError(
            "Serving requires a complete artifact explicitly allowed under this policy"
        )
    for key in ("rights", "reviewed_by", "decision_reason", "canonical_url", "identifier_version"):
        if row[key].strip().lower() in {"", "unknown", "unreviewed", "not_assessed"}:
            raise ValueError(f"Custody review is incomplete: {key}")
    reviewed = datetime.fromisoformat(row["reviewed_at_utc"].replace("Z", "+00:00"))
    retrieved = datetime.fromisoformat(row["retrieved_at_utc"].replace("Z", "+00:00"))
    if (
        reviewed.utcoffset() != timedelta(0)
        or retrieved.utcoffset() != timedelta(0)
        or reviewed < retrieved
    ):
        raise ValueError("Review and acquisition need UTC timestamps in chronological order")


def build_package(
    root: Path,
    allowed: Path,
    blocked: Path,
    output: Path,
    paper_id: str,
    policy_version: str,
) -> dict[str, object]:
    if output.exists():
        raise ValueError("A new custody package directory is required")
    root = root.resolve(strict=True)
    rows, denied = read_csv(allowed), read_csv(blocked)
    if not 1 <= len(rows) <= 1000 or not paper_id or not policy_version:
        raise ValueError("A paper, policy and nonempty reviewed manifest are required")
    blocked_hashes = {row["sha256"] for row in denied if row["sha256"]}
    blocked_locations = {
        (row["canonical_url"], row["archive_member_path"])
        for row in denied if row["canonical_url"]
    }
    prepared: dict[str, bytes] = {}
    total = 0
    transformations = []
    for row in rows:
        verify_review(row, paper_id, policy_version)
        name = row["item_id"]
        if not re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9_.-]{0,99}", name) or name in prepared:
            raise ValueError("Package item IDs must be unique safe filenames")
        if (
            row["served_sha256"] in blocked_hashes
            or (row["canonical_url"], row["archive_member_path"]) in blocked_locations
        ):
            raise ValueError("Blocked material conflicts with an allowed artifact")
        source = retained_file(
            root, row["source_local_path"], row["source_sha256"], row["source_bytes"],
            MAX_SOURCE_BYTES,
        )
        served = retained_file(
            root, row["served_local_path"], row["served_sha256"], row["served_bytes"],
            MAX_SERVED_BYTES,
        )
        if row["archive_member_path"]:
            if not row["transformation_record_id"] or zip_member(
                source, row["archive_member_path"]
            ) != served:
                raise ValueError("Served bytes do not match the reviewed archive member")
            transformations.append({
                "record_id": row["transformation_record_id"],
                "source_sha256": row["source_sha256"],
                "archive_member_path": row["archive_member_path"],
                "served_sha256": row["served_sha256"],
                "operation": "zipfile.ZipFile.read_exact_member",
                "python_version": sys.version,
            })
        elif source != served or row["transformation_record_id"]:
            raise ValueError("Unqualified content transformations cannot enter a blind package")
        total += len(served)
        if total > MAX_PACKAGE_BYTES:
            raise ValueError("Combined package exceeds the serving bound")
        prepared[name] = served
    output.mkdir(parents=True)
    expected = {}
    for name, body in prepared.items():
        snapshot(output / "input" / name, body)
        expected[name] = sha256(body)
    snapshot(output / "allowed.csv", allowed.read_bytes())
    snapshot(output / "blocked.csv", blocked.read_bytes())
    transformation_bytes = (json.dumps(transformations, indent=2) + "\n").encode()
    snapshot(output / "transformations.json", transformation_bytes)
    manifest = {
        "scope": "mechanical_custody_validation_not_semantic_or_human_adjudication",
        "paper_id": paper_id,
        "policy_version": policy_version,
        "allowed_manifest_sha256": sha256(allowed.read_bytes()),
        "blocked_manifest_sha256": sha256(blocked.read_bytes()),
        "builder_sha256": sha256(Path(__file__).read_bytes()),
        "transformations_sha256": sha256(transformation_bytes),
        "expected_input_sha256": expected,
        "served_bytes": total,
        "maximum_package_bytes": MAX_PACKAGE_BYTES,
        "mount_only": "input",
        "pilot_authorized": False,
        "main_study_authorized": False,
    }
    snapshot(output / "manifest.json", (json.dumps(manifest, indent=2) + "\n").encode())
    return manifest


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--evidence-root", type=Path, required=True)
    parser.add_argument("--allowed", type=Path, required=True)
    parser.add_argument("--blocked", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--paper-id", required=True)
    parser.add_argument("--policy-version", required=True)
    args = parser.parse_args()
    print(json.dumps(build_package(
        args.evidence_root, args.allowed, args.blocked, args.output,
        args.paper_id, args.policy_version,
    ), indent=2))


if __name__ == "__main__":
    main()
