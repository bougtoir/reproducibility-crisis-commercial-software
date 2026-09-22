import json
from pathlib import Path

import pytest

from paper2 import corpus_evidence
from paper2.core import sha256, snapshot, write_csv
from paper2.corpus_acquisition import metadata_query
from paper2.corpus_evidence import verify_article, verify_metadata


def test_article_identity_requires_both_identifiers_and_nonempty_body() -> None:
    raw = b"""<article><front><article-meta>
    <article-id pub-id-type="pmid">123</article-id>
    <article-id pub-id-type="pmc">456</article-id>
    </article-meta></front><body><p>Published methods</p></body></article>"""
    assert verify_article(raw, "123", "PMC456")["identity"] == "verified_pmid_and_pmcid"
    for altered in [
        raw.replace(b">123<", b">999<"),
        raw.replace(b">456<", b">999<"),
        raw.replace(b"<p>Published methods</p>", b""),
        b"<error>not available</error>",
    ]:
        with pytest.raises(ValueError):
            verify_article(altered, "123", "PMC456")


def test_metadata_audit_rejects_false_status_and_broken_receipts(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setattr(corpus_evidence, "ROOT", tmp_path)
    frame_path = tmp_path / "data/derived/inference_frame.csv"
    write_csv(frame_path, [{"paper_id": "PMID:123"}], ["paper_id"])
    source = tmp_path / "metadata"
    body = json.dumps(
        {
            "hitCount": 1,
            "resultList": {"result": [{"id": "123", "source": "MED", "isOpenAccess": "N"}]},
        }
    ).encode()
    receipt = {
        "identifier": "PMID_BATCH:123",
        "url": metadata_query(["123"]),
        "sha256": sha256(body),
        "bytes": len(body),
        "http_status": 200,
    }
    snapshot(source / "responses/example/body", body)
    snapshot(source / "responses/example/receipt.json", json.dumps(receipt).encode())
    row = {
        "paper_id": "PMID:123",
        "pmid": "123",
        "pmcid": "",
        "open_access_flag": "N",
        "frame_sha256": sha256(frame_path.read_bytes()),
        "source_snapshot": "responses/example/body",
        "source_snapshot_sha256": sha256(body),
        "status": "identity_verified",
    }
    write_csv(source / "metadata_index.csv", [row], list(row))
    assert verify_metadata(source) == [row]
    for changed in [
        {**row, "status": "not_returned"},
        {**row, "open_access_flag": "Y"},
        {**row, "source_snapshot_sha256": "not-the-hash"},
        {**row, "source_snapshot": "../outside"},
    ]:
        write_csv(source / "metadata_index.csv", [changed], list(changed))
        with pytest.raises(ValueError):
            verify_metadata(source)
