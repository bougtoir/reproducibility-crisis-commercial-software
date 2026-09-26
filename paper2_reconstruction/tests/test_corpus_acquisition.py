import json
from pathlib import Path

import pytest

from paper2.core import sha256
from paper2.corpus_acquisition import fetch_group, metadata_query, validate_batch


def batch_body(records: list[dict[str, str]], hits: int) -> bytes:
    return json.dumps({"hitCount": hits, "resultList": {"result": records}}).encode()


def test_metadata_batch_checks_identity_completeness_and_uniqueness() -> None:
    record = {"id": "123", "source": "MED"}
    assert validate_batch(batch_body([record], 1), ["123", "456"]) == {"123": record}
    assert validate_batch(batch_body([], 0), ["123"]) == {}
    for body, identifiers in [
        (b'{"version": "6.9"}', ["123"]),
        (b'{"hitCount":0,"resultList":{}}', ["123"]),
        (batch_body([{}], 1), ["123"]),
        (batch_body([record], 2), ["123"]),
        (batch_body([record, record], 2), ["123"]),
        (batch_body([record], 1), ["456"]),
        (batch_body([{"id": "123", "source": "PMC"}], 1), ["123"]),
    ]:
        with pytest.raises(ValueError):
            validate_batch(body, identifiers)


def test_metadata_query_accepts_only_bounded_numeric_identifiers() -> None:
    assert "resultType=core" in metadata_query(["123", "456"])
    for identifiers in [[], ["bad OR MATCHALL"], ["1"] * 51]:
        with pytest.raises(ValueError):
            metadata_query(identifiers)


def test_invalid_batch_is_retained_and_decomposed_without_inventing_absence(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    requested = []

    def acquire_fixture(
        url: str, identifier: str, directory: Path
    ) -> tuple[Path, dict[str, object]]:
        requested.append(identifier)
        body = (
            batch_body([{"id": "123", "source": "MED"}], 1)
            if identifier == "PMID_BATCH:123"
            else b'{"version":"6.9"}'
        )
        path = tmp_path / sha256(url.encode())
        path.write_bytes(body)
        return path, {"http_status": 200, "sha256": sha256(body)}

    monkeypatch.setattr("paper2.corpus_acquisition.acquire", acquire_fixture)
    monkeypatch.setattr("paper2.corpus_acquisition.time.sleep", lambda seconds: None)
    rows = fetch_group(["123", "456"], tmp_path)
    assert requested == ["PMID_BATCH:123,456", "PMID_BATCH:123", "PMID_BATCH:456"]
    assert rows["123"]["status"] == "identity_verified"
    assert rows["456"]["status"] == "invalid_response"
    assert len(list(tmp_path.iterdir())) == 3
