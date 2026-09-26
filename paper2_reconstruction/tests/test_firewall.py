import io
import json
import tarfile
from pathlib import Path

import pytest

from paper2.controller import Journal
from paper2.core import sha256
from paper2.firewall import Broker, Custodian, Denied, Item, negative_cases, vetted_wheels

ARTICLE = b"<article>retained publication text</article>\n"
IMPLEMENTATION = b"SYNTHETIC_ORIGINAL_IMPLEMENTATION_CANARY\n"
NOTES = b"SYNTHETIC_AUTHOR_NOTES_CANARY\n"
MEMBERS = {
    "articles/1.xml": ARTICLE,
    "implementation/original_pipeline.py": IMPLEMENTATION,
    "implementation/original_notes.txt": NOTES,
}


def archive(destination: Path, members: dict[str, bytes]) -> Path:
    path = destination / "mixed.tar"
    with tarfile.open(path, mode="w") as handle:
        for name, payload in members.items():
            info = tarfile.TarInfo(name)
            info.size = len(payload)
            info.mtime = 0
            handle.addfile(info, io.BytesIO(payload))
    return path


def items() -> dict[str, Item]:
    return {
        "I000": Item(
            "I000", "1", "articles/1.xml", sha256(ARTICLE), len(ARTICLE),
            "publication_text", "allow", "reviewed retained publication text",
        ),
        "I001": Item(
            "I001", "", "implementation/original_pipeline.py", sha256(IMPLEMENTATION),
            len(IMPLEMENTATION), "reveal_original_implementation", "deny",
            "forbidden before the blind outcome freeze",
        ),
        "I002": Item(
            "I002", "", "implementation/original_notes.txt", sha256(NOTES), len(NOTES),
            "author_implementation_material", "deny", "never served",
        ),
    }


def broker(tmp_path: Path, members: dict[str, bytes] = MEMBERS) -> Broker:
    custodian = Custodian(archive(tmp_path, members), items())
    return Broker(custodian, Journal(tmp_path / "events"), tmp_path / "package")


def test_broker_serves_reviewed_member_and_denies_everything_else(tmp_path: Path) -> None:
    subject = broker(tmp_path)
    assert subject.request("solver_package_builder", "I000") == ARTICLE
    assert (tmp_path / "package" / "1.xml").read_bytes() == ARTICLE
    assert subject.served == {"1.xml": sha256(ARTICLE)}
    results = negative_cases(subject, items())
    assert all(value.startswith("denied") for value in results.values())
    assert sorted(results) == [
        "blocked_author_material_denied",
        "reveal_item_denied_before_freeze",
        "reveal_item_denied_to_solver",
        "unauthorised_actor_denied",
        "unlisted_item_denied",
    ]


def test_every_request_including_denials_is_journalled(tmp_path: Path) -> None:
    subject = broker(tmp_path)
    subject.request("solver_package_builder", "I000")
    with pytest.raises(Denied):
        subject.request("solver_package_builder", "I002")
    kinds = [
        json.loads(path.read_bytes())["kind"]
        for path in sorted((tmp_path / "events").glob("event-*.json"))
    ]
    assert kinds == ["access_allowed", "access_denied"]


def test_tampered_custodian_bytes_are_refused(tmp_path: Path) -> None:
    subject = broker(tmp_path, {**MEMBERS, "articles/1.xml": b"<article>swapped</article>\n"})
    with pytest.raises(Denied, match="reviewed identity"):
        subject.request("solver_package_builder", "I000")


def test_reveal_requires_a_freeze_receipt_covering_the_sealed_outcome(tmp_path: Path) -> None:
    subject = broker(tmp_path)
    outcome = tmp_path / "blind_outcome.json"
    outcome.write_bytes(b'{"sealed": true}\n')
    with pytest.raises(Denied, match="does not cover"):
        subject.freeze(
            outcome,
            {"source_sha256": "0" * 64, "response_sha256": "a", "trust_anchor_sha256": "b"},
        )
    with pytest.raises(Denied, match="before the blind outcome freeze"):
        subject.request("reveal_reviewer", "I001")
    subject.freeze(
        outcome,
        {
            "source_sha256": sha256(outcome.read_bytes()),
            "response_sha256": "a",
            "trust_anchor_sha256": "b",
        },
    )
    assert subject.request("reveal_reviewer", "I001") == IMPLEMENTATION
    assert not (tmp_path / "package" / "original_pipeline.py").exists()
    with pytest.raises(Denied):
        subject.request("solver_package_builder", "I001")


def test_dependency_closure_requires_every_reviewed_requirement(tmp_path: Path) -> None:
    with pytest.raises(ValueError, match="is missing"):
        vetted_wheels(tmp_path)
