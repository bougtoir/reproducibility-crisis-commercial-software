import json
from pathlib import Path

import pytest

from paper2.core import sha256, snapshot
from paper2.primary_adjudication import render

ARTICLE = (
    b"<article><body><p>The fitted coefficient was 0.42 using public inputs."
    b"</p></body></article>"
)
SELECTED = "selected_provisional_pilot_case"
SEGMENT = "The fitted coefficient was 0.42 using public inputs."


def assessment(**overrides: object) -> dict[str, object]:
    row: dict[str, object] = {
        "paper_id": "PMID:123",
        "sampling_stratum": "Physics_Engineering",
        "candidate_rank": 1,
        "article_text_status": "identity_verified_open_pmc_jats",
        "G1_computationally_testable": "yes",
        "G1_reason": "central fitted coefficient",
        "G2_principal_target_identifiable": "yes",
        "G3_input_accessibility": "public",
        "G4_resource_accessibility": "available",
        "G5_specification_sufficient_for_attempt": "uncertain",
        "target_candidate": "fitted coefficient = 0.42",
        "alternatives_considered": "none",
        "pilot_decision": "provisional_pilot_case",
        "evidence": [{"segment_id": "S0001", "quote": "coefficient was 0.42", "gate": "G1,G2"}],
    }
    return {**row, **overrides}


def prepare(root: Path, assessments: list[dict[str, object]], chain: object) -> tuple[Path, Path]:
    screen = root / "screen"
    sources = {
        "sources": [
            {
                "role": "standalone_article_xml",
                "source_sha256": sha256(ARTICLE),
                "segments": [{"segment_id": "S0001", "text": SEGMENT}],
            }
        ]
    }
    snapshot(screen / "123" / "sources.json", json.dumps(sources).encode())
    record = {
        "record_id": "r",
        "scope": "devin_primary_assessment_pending_author_verification",
        "author_verification": "pending",
        "deviation": {"deviation_id": "DEV-1"},
        "assessments": assessments,
        "deterministic_candidate_chain": {"rule": "deterministic", "strata": chain},
    }
    path = root / "record.json"
    snapshot(path, (json.dumps(record) + "\n").encode())
    return path, screen


def test_validated_record_renders_pilot_cases(tmp_path: Path) -> None:
    chain = {
        "Physics_Engineering": [
            {"paper_id": "PMID:123", "candidate_rank": 1, "disposition": SELECTED}
        ]
    }
    path, screen = prepare(tmp_path, [assessment()], chain)
    results = tmp_path / "results"
    results.mkdir()
    summary = render(path, screen, results)
    assert summary["provisional_pilot_cases"]["Physics_Engineering"]["paper_id"] == "PMID:123"
    assert summary["formal_funnel_updated"] is False
    assert summary["record_sha256"] == sha256(path.read_bytes())
    rows = (results / "pilot_primary_adjudication.csv").read_text()
    assert "DEVIN_PRIMARY_PENDING_AUTHOR_VERIFICATION" in rows
    assert sha256(ARTICLE) in rows


def test_quote_absent_from_cited_segment_is_rejected(tmp_path: Path) -> None:
    chain = {
        "Physics_Engineering": [
            {"paper_id": "PMID:123", "candidate_rank": 1, "disposition": SELECTED}
        ]
    }
    evidence = [{"segment_id": "S0001", "quote": "coefficient was 0.99", "gate": "G1"}]
    path, screen = prepare(tmp_path, [assessment(evidence=evidence)], chain)
    (tmp_path / "results").mkdir()
    with pytest.raises(ValueError, match="verbatim"):
        render(path, screen, tmp_path / "results")


def test_g1_negative_cannot_be_a_pilot_case(tmp_path: Path) -> None:
    chain = {
        "Physics_Engineering": [
            {"paper_id": "PMID:123", "candidate_rank": 1, "disposition": SELECTED}
        ]
    }
    row = assessment(
        G1_computationally_testable="no", G2_principal_target_identifiable="not_applicable"
    )
    path, screen = prepare(tmp_path, [row], chain)
    (tmp_path / "results").mkdir()
    with pytest.raises(ValueError, match="G1=no cannot be a pilot case"):
        render(path, screen, tmp_path / "results")


def test_chain_must_consume_every_earlier_deterministic_rank(tmp_path: Path) -> None:
    chain = {
        "Physics_Engineering": [
            {"paper_id": "PMID:123", "candidate_rank": 2, "disposition": SELECTED}
        ]
    }
    path, screen = prepare(tmp_path, [assessment(candidate_rank=2)], chain)
    (tmp_path / "results").mkdir()
    with pytest.raises(ValueError, match="skips deterministic ranks"):
        render(path, screen, tmp_path / "results")
