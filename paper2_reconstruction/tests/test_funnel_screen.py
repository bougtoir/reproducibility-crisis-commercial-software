import json
from pathlib import Path

import pytest

from paper2.core import Row, sha256, snapshot
from paper2.funnel_screen import metadata_source, retained_body, screen_paper
from paper2.model_api import MODEL

ABSTRACT = "Synthetic abstract " + " ".join(f"word{i}" for i in range(80))
ARTICLE = (
    b"<article><body><p>Synthetic article text reporting a fitted regression coefficient "
    b"of 0.42 for the primary computational endpoint using public inputs.</p></body></article>"
)


def prepare(root: Path) -> tuple[Row, Row, Row, Path, Path]:
    metadata, articles = root / "metadata", root / "articles"
    record = {"id": "123", "source": "MED", "title": "Synthetic title", "abstractText": ABSTRACT}
    body = json.dumps({"resultList": {"result": [record, {"id": "9", "source": "MED"}]}}).encode()
    snapshot(metadata / "responses/m/body", body)
    snapshot(articles / "responses/a/body", ARTICLE)
    paper = {"paper_id": "PMID:123", "sampling_stratum": "Physics_Engineering"}
    metadata_row = {
        "paper_id": "PMID:123", "pmid": "123", "status": "identity_verified",
        "source_snapshot": "responses/m/body", "source_snapshot_sha256": sha256(body),
    }
    article_row = {
        "paper_id": "PMID:123", "status": "identity_verified_article",
        "response_path": "responses/a/body", "response_sha256": sha256(ARTICLE),
    }
    return paper, metadata_row, article_row, metadata, articles


def store_completion(case: Path, content: dict[str, object]) -> None:
    completion = {
        "reported_model": MODEL, "finish_reason": "stop", "content": json.dumps(content),
        "prompt_tokens": 10, "completion_tokens": 5,
    }
    snapshot(case / "bounded_api/api/completion.json", json.dumps(completion).encode())


def review(article_hash: str, segment: str, metadata_hash: str) -> dict[str, object]:
    return {
        "computationally_testable": "yes",
        "principal_target_identifiable": "yes",
        "target_candidate": "fitted regression coefficient",
        "inputs_required": "public inputs",
        "access_observation": "unknown",
        "specification_observation": "unknown",
        "evidence": [
            {"source_sha256": article_hash, "segment_id": segment, "observation": "coefficient"},
            {"source_sha256": metadata_hash, "segment_id": "S0001", "observation": "abstract"},
        ],
    }


def test_screening_binds_review_to_retained_metadata_and_article(tmp_path: Path) -> None:
    paper, metadata_row, article_row, metadata, articles = prepare(tmp_path)
    case = tmp_path / "case"
    article_hash = sha256(ARTICLE)
    metadata_hash = metadata_source(metadata, metadata_row).source_sha256
    store_completion(case, review(article_hash, "S0001", metadata_hash))
    row = screen_paper(paper, metadata_row, article_row, metadata, articles, case)
    assert row["status"] == "provisional_machine_review_not_validated_or_selected"
    assert row["article_text_status"] == "identity_verified_open_pmc_jats"
    assert row["computationally_testable"] == "yes"
    assert row["provider_accounted_tokens"] == "15"
    sources = json.loads((case / "sources.json").read_text())
    assert [item["role"] for item in sources["sources"]] == [
        "title_and_abstract_only", "standalone_article_xml",
    ]


def test_abstract_only_negative_is_downgraded_to_uncertain(tmp_path: Path) -> None:
    paper, metadata_row, article_row, metadata, articles = prepare(tmp_path)
    article_row = {**article_row, "status": "pmc_endpoint_not_selected_other_sources_unassessed"}
    case = tmp_path / "case"
    metadata_hash = metadata_source(metadata, metadata_row).source_sha256
    negative = review(metadata_hash, "S0001", metadata_hash)
    negative["computationally_testable"] = "no"
    store_completion(case, negative)
    row = screen_paper(paper, metadata_row, article_row, metadata, articles, case)
    assert row["status"] == "abstract_only_negative_not_validated"
    assert row["computationally_testable"] == "uncertain"
    assert row["article_text_status"] == "no_open_pmc_endpoint_other_routes_unassessed"


def test_unverified_metadata_is_not_screened(tmp_path: Path) -> None:
    paper, metadata_row, article_row, metadata, articles = prepare(tmp_path)
    metadata_row = {**metadata_row, "status": "not_returned"}
    row = screen_paper(paper, metadata_row, article_row, metadata, articles, tmp_path / "c")
    assert row["status"] == "metadata_not_identity_verified_unscreened"
    assert row["provider_accounted_tokens"] == "unknown"
    assert not (tmp_path / "c").exists()


def test_tampered_or_escaping_retained_sources_are_rejected(tmp_path: Path) -> None:
    paper, metadata_row, article_row, metadata, articles = prepare(tmp_path)
    with pytest.raises(ValueError, match="indexed hash"):
        retained_body(articles, "responses/a/body", "0" * 64)
    with pytest.raises(ValueError, match="escapes"):
        retained_body(articles, "../metadata/responses/m/body", sha256(b""))
    with pytest.raises(ValueError, match="exactly one record"):
        metadata_source(metadata, {**metadata_row, "pmid": "8"})
