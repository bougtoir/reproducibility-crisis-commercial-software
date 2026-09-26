import io
import json
import sys
import tarfile
from pathlib import Path

import pytest

from paper2.controller import Journal, Limits, decode_action
from paper2.core import sha256
from paper2.isolation import bounded_capture, regular_tar_members


def test_capture_enforces_live_output_and_time_limits() -> None:
    output, errors, status, reason = bounded_capture(
        [sys.executable, "-c", "print('observed')"], seconds=2, byte_limit=64
    )
    assert (output, errors, status, reason) == (b"observed\n", b"", 0, "completed")
    output, errors, _, reason = bounded_capture(
        [sys.executable, "-c", "print('x' * 10000)"], seconds=2, byte_limit=64
    )
    assert reason == "output_limit" and len(output) + len(errors) == 64
    _, _, _, reason = bounded_capture(
        [sys.executable, "-c", "import time; time.sleep(30)"], seconds=0.05, byte_limit=64
    )
    assert reason == "timeout"


def tar_member(name: str, *, kind: bytes = tarfile.REGTYPE) -> bytes:
    output = io.BytesIO()
    with tarfile.open(fileobj=output, mode="w") as archive:
        data = b"result"
        member = tarfile.TarInfo(name)
        member.type = kind
        member.size = len(data) if kind == tarfile.REGTYPE else 0
        archive.addfile(member, io.BytesIO(data))
    return output.getvalue()


def test_worker_artifact_tar_accepts_only_regular_relative_members() -> None:
    assert regular_tar_members(tar_member("./result.json")) == {"result.json": b"result"}
    for raw in [
        tar_member("../escape"),
        tar_member("/absolute"),
        tar_member("link", kind=tarfile.SYMTYPE),
    ]:
        with pytest.raises(ValueError):
            regular_tar_members(raw)


@pytest.mark.parametrize(
    "action",
    [
        {"action": "shell", "code": "true"},
        {"action": "fetch", "url": "https://example.org"},
        {"action": "python", "code": "print(1)", "host": True},
        {"action": "python", "code": ""},
        {"action": "python", "code": 123},
        {"action": "final", "report": {"private_reasoning": "must not be accepted"}},
    ],
)
def test_controller_rejects_unregistered_actions_and_fields(action: dict[str, object]) -> None:
    with pytest.raises(ValueError):
        decode_action(json.dumps(action))


def test_journal_hash_chain_and_no_overwrite(tmp_path: Path) -> None:
    journal = Journal(tmp_path)
    journal.record("first", {"observation": 1})
    first_hash = journal.head
    journal.record("second", {"observation": 2})
    entries = [path.read_bytes() for path in sorted(tmp_path.glob("event-*.json"))]
    assert sha256(entries[0]) == first_hash
    assert json.loads(entries[1])["previous_sha256"] == first_hash
    assert sha256(entries[1]) == journal.head
    with pytest.raises(ValueError):
        Journal(tmp_path).record("overwrite", None)
    assert (tmp_path / "event-0000.json").read_bytes() == entries[0]


def test_limits_reject_empty_or_unaccountable_contract() -> None:
    with pytest.raises(ValueError):
        Limits(seconds=0).validate()
    with pytest.raises(ValueError):
        Limits(tokens=1024).validate()
