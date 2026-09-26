import io
import json
import zipfile
from pathlib import Path

import pytest

from paper2.core import Row, sha256, snapshot, write_csv
from paper2.packages import build_package, zip_member


def fixture_row() -> Row:
    return {
        "item_id": "input.csv",
        "paper_id": "synthetic",
        "policy_version": "synthetic-only",
        "decision": "allow",
        "role": "public_input",
        "completeness": "complete",
        "rights": "synthetic fixture created for tests",
        "reviewed_by": "synthetic-test-custodian",
        "decision_reason": "synthetic data only",
        "canonical_url": "https://example.invalid/synthetic",
        "identifier_version": "synthetic-v1",
        "retrieved_at_utc": "2026-09-22T00:00:00Z",
        "reviewed_at_utc": "2026-09-22T00:00:01Z",
        "source_local_path": "source.csv",
        "source_sha256": sha256(b"value\n3\n"),
        "source_bytes": str(len(b"value\n3\n")),
        "served_local_path": "source.csv",
        "served_sha256": sha256(b"value\n3\n"),
        "served_bytes": str(len(b"value\n3\n")),
        "archive_member_path": "",
        "transformation_record_id": "",
    }


def prepare(tmp_path: Path, row: Row, denied: list[Row]) -> dict[str, object]:
    snapshot(tmp_path / "source.csv", b"value\n3\n")
    write_csv(tmp_path / "allowed.csv", [row], list(row))
    write_csv(
        tmp_path / "blocked.csv", denied, ["sha256", "canonical_url", "archive_member_path"]
    )
    return build_package(
        tmp_path, tmp_path / "allowed.csv", tmp_path / "blocked.csv", tmp_path / "package",
        "synthetic", "synthetic-only",
    )


def test_only_approved_input_bytes_enter_the_mount(tmp_path: Path) -> None:
    result = prepare(tmp_path, fixture_row(), [])
    assert list((tmp_path / "package/input").iterdir()) == [tmp_path / "package/input/input.csv"]
    assert result["expected_input_sha256"] == {"input.csv": sha256(b"value\n3\n")}
    assert result["pilot_authorized"] is False
    assert json.loads((tmp_path / "package/manifest.json").read_text()) == result


@pytest.mark.parametrize(
    ("field", "value"),
    [
        ("decision", "unknown"),
        ("reviewed_by", "unknown"),
        ("source_local_path", "../source.csv"),
        ("source_sha256", "incorrect"),
        ("item_id", "../escape"),
        ("policy_version", "wrong-policy"),
        ("completeness", "partial"),
        ("served_bytes", "0"),
    ],
)
def test_unreviewed_or_mismatched_input_is_denied(
    tmp_path: Path, field: str, value: str
) -> None:
    with pytest.raises(ValueError):
        prepare(tmp_path, {**fixture_row(), field: value}, [])
    assert not (tmp_path / "package").exists()


def test_blocked_content_and_symlinks_are_denied(tmp_path: Path) -> None:
    row = fixture_row()
    with pytest.raises(ValueError, match="Blocked material"):
        prepare(tmp_path, row, [
            {"sha256": row["served_sha256"], "canonical_url": "", "archive_member_path": ""}
        ])
    (tmp_path / "link.csv").symlink_to(tmp_path / "source.csv")
    with pytest.raises(ValueError, match="symbolic links"):
        prepare(tmp_path, {**row, "source_local_path": "link.csv"}, [])


def test_reviewed_archive_member_is_bound_to_original_bytes(tmp_path: Path) -> None:
    container = io.BytesIO()
    with zipfile.ZipFile(container, "w") as archive:
        archive.writestr("inputs/source.csv", b"value\n3\n")
        archive.writestr("unreviewed.txt", b"Must not be served")
    raw = container.getvalue()
    snapshot(tmp_path / "source.zip", raw)
    row = {
        **fixture_row(),
        "source_local_path": "source.zip",
        "source_sha256": sha256(raw),
        "source_bytes": str(len(raw)),
        "archive_member_path": "inputs/source.csv",
        "transformation_record_id": "synthetic-extraction",
    }
    result = prepare(tmp_path, row, [
        {"sha256": sha256(raw), "canonical_url": row["canonical_url"], "archive_member_path": ""}
    ])
    assert result["expected_input_sha256"] == {"input.csv": sha256(b"value\n3\n")}
    assert not (tmp_path / "package/input/unreviewed.txt").exists()
    with pytest.raises(ValueError):
        zip_member(raw, "../source.csv")
