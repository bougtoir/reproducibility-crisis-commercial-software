import json
from pathlib import Path
from unittest.mock import patch

import pytest

from paper2.acquisition import API
from paper2.core import Row, sha256, snapshot, write_csv
from paper2.corpus_evidence import audit_articles, verify_article

BODY = b"""<article><front><article-meta>
<article-id pub-id-type="pmid">123</article-id>
<article-id pub-id-type="pmc">456</article-id>
</article-meta></front><body><p>Synthetic fixture</p></body></article>"""


def prepare(root: Path) -> tuple[Path, Path, list[Row], list[Row]]:
    metadata, source = root / "metadata", root / "articles"
    reference = {
        "paper_id": "PMID:123", "pmid": "123", "pmcid": "PMC456",
        "frame_sha256": "synthetic-frame", "source_snapshot_sha256": "synthetic-metadata",
        "status": "identity_verified", "open_access_flag": "Y",
    }
    closed = {**reference, "paper_id": "PMID:124", "pmid": "124", "pmcid": "",
              "open_access_flag": "N"}
    references = [reference, closed]
    write_csv(metadata / "metadata_index.csv", references, list(reference))
    snapshot(source / "metadata_index_snapshot.csv", (metadata / "metadata_index.csv").read_bytes())
    snapshot(source / "responses/example/body", BODY)
    receipt = {
        "sha256": sha256(BODY), "bytes": len(BODY), "identifier": "PMC456",
        "url": f"{API}/PMC456/fullTextXML", "http_status": 200,
    }
    snapshot(source / "responses/example/receipt.json", json.dumps(receipt).encode())
    snapshot(
        source / "responses/example/article_validation.json",
        json.dumps(verify_article(BODY, "123", "PMC456")).encode(),
    )
    row = {
        "paper_id": "PMID:123", "pmid": "123", "pmcid": "PMC456",
        "frame_sha256": "synthetic-frame", "metadata_sha256": "synthetic-metadata",
        "status": "identity_verified_article", "response_sha256": sha256(BODY),
        "response_path": "responses/example/body", "http_status": "200", "validation_error": "",
    }
    rows = [row, {
        **row, "paper_id": "PMID:124", "pmid": "124", "pmcid": "",
        "status": "pmc_endpoint_not_selected_other_sources_unassessed",
        "response_sha256": "", "response_path": "", "http_status": "",
    }]
    write_csv(source / "article_index.csv", rows, list(row))
    return metadata, source, references, rows


def test_article_audit_recomputes_identity_and_distinguishes_endpoint_scope(tmp_path: Path) -> None:
    metadata, source, references, rows = prepare(tmp_path)
    with patch("paper2.corpus_evidence.verify_metadata", return_value=references):
        report = audit_articles(metadata, source)
    assert report["indexed_papers"] == len(rows)
    assert report["coverage"] == "complete_frame_index"
    assert report["input_access"] == "not_assessed"
    assert report["article_status_counts"] == {
        "identity_verified_article": 1,
        "pmc_endpoint_not_selected_other_sources_unassessed": 1,
    }


@pytest.mark.parametrize(
    ("field", "value"),
    [
        ("response_sha256", "different"),
        ("response_path", "../escape"),
        ("status", "unverified_article_response"),
        ("metadata_sha256", "different"),
        ("http_status", "404"),
    ],
)
def test_article_audit_rejects_unsupported_index_assertions(
    tmp_path: Path, field: str, value: str
) -> None:
    metadata, source, references, rows = prepare(tmp_path)
    rows[0][field] = value
    write_csv(source / "article_index.csv", rows, list(rows[0]))
    with (
        patch("paper2.corpus_evidence.verify_metadata", return_value=references),
        pytest.raises(ValueError),
    ):
        audit_articles(metadata, source)


def test_incomplete_index_requires_explicit_partial_mode(tmp_path: Path) -> None:
    metadata, source, references, rows = prepare(tmp_path)
    write_csv(source / "article_index.csv", rows[:1], list(rows[0]))
    with patch("paper2.corpus_evidence.verify_metadata", return_value=references):
        with pytest.raises(ValueError, match="ordered frame"):
            audit_articles(metadata, source)
        report = audit_articles(metadata, source, allow_partial=True)
    assert report["coverage"] == "partial_frame_index"
    assert report["unprocessed_papers"] == 1


def test_modified_body_cannot_pass_with_unchanged_receipt(tmp_path: Path) -> None:
    metadata, source, references, _ = prepare(tmp_path)
    (source / "responses/example/body").write_bytes(BODY.replace(b">123<", b">999<"))
    with (
        patch("paper2.corpus_evidence.verify_metadata", return_value=references),
        pytest.raises(ValueError, match="receipt"),
    ):
        audit_articles(metadata, source)


def test_http_failure_remains_distinct_from_unavailable_input(tmp_path: Path) -> None:
    metadata, source, references, rows = prepare(tmp_path)
    receipt_path = source / "responses/example/receipt.json"
    receipt = json.loads(receipt_path.read_bytes())
    receipt["http_status"] = 404
    receipt_path.write_text(json.dumps(receipt))
    rows[0].update(http_status="404", status="request_failed_not_access_verdict")
    write_csv(source / "article_index.csv", rows, list(rows[0]))
    with patch("paper2.corpus_evidence.verify_metadata", return_value=references):
        report = audit_articles(metadata, source)
    assert report["input_access"] == "not_assessed"
    assert report["article_status_counts"] == {
        "request_failed_not_access_verdict": 1,
        "pmc_endpoint_not_selected_other_sources_unassessed": 1,
    }


@pytest.mark.parametrize("change", ["order", "snapshot", "unsupported_endpoint"])
def test_article_audit_rejects_wrong_frame_or_selection(tmp_path: Path, change: str) -> None:
    metadata, source, references, rows = prepare(tmp_path)
    if change == "order":
        rows.reverse()
    elif change == "snapshot":
        (source / "metadata_index_snapshot.csv").write_text("unrelated metadata")
    else:
        rows[1]["response_path"] = "responses/example/body"
    write_csv(source / "article_index.csv", rows, list(rows[0]))
    with (
        patch("paper2.corpus_evidence.verify_metadata", return_value=references),
        pytest.raises(ValueError),
    ):
        audit_articles(metadata, source)
