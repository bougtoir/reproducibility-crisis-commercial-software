import json
from pathlib import Path

import pytest

from paper2.access_layer import AMENDMENT, LAYER, validate_layer
from paper2.deposit_screen import amended_record_classes, screen_routes, screened_papers

ROOT = Path(__file__).resolve().parents[1]


def member(**overrides: object) -> dict[str, object]:
    row: dict[str, object] = {
        "paper_id": "PMID:1",
        "layer": LAYER,
        "main_sample_membership": "excluded",
        "reconstruction_attempt_reached": False,
        "exclusion_classes": ["author_request_only"],
    }
    return {**row, **overrides}


def layer(**overrides: object) -> dict[str, object]:
    payload: dict[str, object] = {
        "amendment_id": AMENDMENT,
        "gate_verification": {"investigator_verification": "pending"},
        "rates": "prohibited: no member reached a reconstruction attempt",
        "members": [member()],
    }
    return {**payload, **overrides}


def test_layer_validates_and_counts_classes() -> None:
    summary = validate_layer(layer())
    assert summary["members"] == 1
    assert summary["exclusion_class_counts"] == {"author_request_only": 1}
    assert summary["investigator_verification"] == "pending"


def test_layer_rejects_outcome_fields() -> None:
    with pytest.raises(ValueError, match="outcome field"):
        validate_layer(layer(success_rate=0.0))


def test_layer_rejects_member_that_reached_an_attempt() -> None:
    with pytest.raises(ValueError, match="never reach"):
        validate_layer(layer(members=[member(reconstruction_attempt_reached=True)]))


def test_committed_layer_record_validates() -> None:
    path = ROOT / "data" / "adjudication" / "accessibility_gate_failed_20260924.json"
    summary = validate_layer(json.loads(path.read_text()))
    assert summary["members"] == 7
    assert summary["rates"].startswith("prohibited")


def test_route_continuation_skips_screened_papers_and_carries_inspection_class(
    tmp_path: Path,
) -> None:
    statements = [
        {
            "paper_id": "PMID:1",
            "sampling_stratum": "Chemistry_Materials",
            "candidate_rank": 1,
            "statement_class": "open_route_named_pending_listing",
            "hits": [],
        },
        {
            "paper_id": "PMID:2",
            "sampling_stratum": "Chemistry_Materials",
            "candidate_rank": 2,
            "statement_class": "open_route_named_pending_listing",
            "hits": [],
        },
        {
            "paper_id": "PMID:3",
            "sampling_stratum": "Chemistry_Materials",
            "candidate_rank": 3,
            "statement_class": "no_deposit_named",
            "hits": [],
        },
    ]
    record = tmp_path / "amended.json"
    record.write_text(
        json.dumps(
            {
                "assessments": [
                    {"paper_id": "PMID:2", "amended_disposition": "excluded_deposited_output_only"}
                ]
            }
        )
    )
    inspected = amended_record_classes(record)
    assert inspected == {"PMID:2": "deposited_output_only"}
    previous = tmp_path / "stage_b.json"
    previous.write_text(json.dumps({"rows": [{"paper_id": "PMID:1"}, {"paper_id": "PMID:2"}]}))
    screened = frozenset(screened_papers(previous) - set(inspected))
    rows = screen_routes(
        statements, tmp_path, 4, dict(inspected), ("Chemistry_Materials",), screened
    )
    assert [r["paper_id"] for r in rows] == ["PMID:2", "PMID:3"]
    assert rows[0]["exclusion_class"] == "deposited_output_only"
    assert rows[1]["exclusion_class"] == "no_deposit_named"
