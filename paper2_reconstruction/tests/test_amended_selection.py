import json
from pathlib import Path

import pytest

from paper2.amended_selection import AMENDMENT, render
from paper2.core import sha256, snapshot

ARTICLE = (
    b"<article><body><p>The fitted coefficient was 0.42 using public inputs.</p></body></article>"
)
SEGMENT = "The fitted coefficient was 0.42 using public inputs."


def assessment(**overrides: object) -> dict[str, object]:
    row: dict[str, object] = {
        "paper_id": "PMID:123",
        "sampling_stratum": "Physics_Engineering",
        "candidate_rank": "7",
        "article_text_status": "lawful_text_retained",
        "G1_computationally_testable": "yes",
        "G1_reason": "central fitted coefficient",
        "G2_principal_target_identifiable": "yes",
        "G3_input_accessibility": "public",
        "G4_resource_accessibility": "available",
        "G5_specification_sufficient_for_attempt": "uncertain",
        "target_candidate": "fitted coefficient = 0.42",
        "alternatives_considered": "none",
        "pilot_decision": "provisional_pilot_case",
        "deposit_role": "input_raw_table",
        "outcome_leakage": "none_in_deposited_files",
        "amended_disposition": "main_sample_candidate",
        "evidence": [{"segment_id": "S0001", "quote": "coefficient was 0.42", "gate": "G1,G2"}],
    }
    return {**row, **overrides}


def route(**overrides: object) -> dict[str, object]:
    row: dict[str, object] = {
        "paper_id": "PMID:123",
        "sampling_stratum": "Physics_Engineering",
        "candidate_rank": 7,
        "statement_class": "open_route_named_pending_listing",
        "exclusion_class": "",
        "routes": [
            {
                "kind": "figshare",
                "exclusion_class": "",
                "listed_bytes": 1000,
                "implementation_files_withheld": 1,
            }
        ],
    }
    return {**row, **overrides}


def prepare(
    root: Path,
    assessments: list[dict[str, object]],
    routes: list[dict[str, object]],
    continuation: list[dict[str, object]] | None = None,
) -> tuple[Path, Path, list[Path]]:
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
        "scope": "devin_delegated_amended_assessment",
        "author_verification": "pending",
        "amendment_id": AMENDMENT,
        "assessments": assessments,
    }
    record_path = root / "record.json"
    snapshot(record_path, (json.dumps(record) + "\n").encode())
    stage_b = root / "routes.json"
    snapshot(stage_b, json.dumps({"amendment_id": AMENDMENT, "rows": routes}).encode())
    paths = [stage_b]
    if continuation is not None:
        cont = root / "routes_cont.json"
        payload = {
            "amendment_id": AMENDMENT,
            "continues_from_sha256": sha256(stage_b.read_bytes()),
            "rows": continuation,
        }
        snapshot(cont, json.dumps(payload).encode())
        paths.append(cont)
    return record_path, screen, paths


def test_main_sample_candidate_is_rendered_without_rates(tmp_path: Path) -> None:
    record, screen, routes = prepare(tmp_path, [assessment()], [route()])
    results = tmp_path / "results"
    results.mkdir()
    summary = render(record, screen, routes, results)
    assert summary["main_sample_candidates"]["Physics_Engineering"]["paper_id"] == "PMID:123"
    assert summary["reconstruction_started"] is False
    assert summary["rates"].startswith("not applicable")
    assert summary["author_verification"] == "pending"
    csv_text = (results / "amended_candidate_selection.csv").read_text()
    assert "DEVIN_PRIMARY_PENDING_AUTHOR_VERIFICATION" in csv_text
    assert "1000" in csv_text


def test_disclosed_target_cannot_be_main_sample(tmp_path: Path) -> None:
    row = assessment(outcome_leakage="target_disclosed_by_deposited_file")
    record, screen, routes = prepare(tmp_path, [row], [route()])
    with pytest.raises(ValueError, match="disclosed target"):
        render(record, screen, routes, tmp_path)


def test_route_excluded_candidate_cannot_be_main_sample(tmp_path: Path) -> None:
    record, screen, routes = prepare(
        tmp_path, [assessment()], [route(exclusion_class="author_request_only")]
    )
    with pytest.raises(ValueError, match="disagrees"):
        render(record, screen, routes, tmp_path)


def test_g1_negative_cannot_be_main_sample(tmp_path: Path) -> None:
    record, screen, routes = prepare(
        tmp_path, [assessment(G1_computationally_testable="no")], [route()]
    )
    with pytest.raises(ValueError, match="G1=yes"):
        render(record, screen, routes, tmp_path)


def test_every_route_eligible_candidate_must_be_assessed(tmp_path: Path) -> None:
    extra = route(paper_id="PMID:456", sampling_stratum="Clinical_Medicine")
    record, screen, routes = prepare(tmp_path, [assessment()], [route(), extra])
    with pytest.raises(ValueError, match="PMID:456"):
        render(record, screen, routes, tmp_path)


def test_author_verification_must_stay_pending(tmp_path: Path) -> None:
    record, screen, routes = prepare(tmp_path, [assessment()], [route()])
    payload = json.loads(record.read_text())
    payload["author_verification"] = "verified"
    record.write_text(json.dumps(payload))
    with pytest.raises(ValueError, match="pending"):
        render(record, screen, routes, tmp_path)


def test_continuation_replaces_excluded_candidate_in_stratum(tmp_path: Path) -> None:
    excluded = assessment(
        outcome_leakage="target_disclosed_by_deposited_file",
        amended_disposition="excluded_deposited_output_only",
        pilot_decision="not_pilot_case_replaced",
    )
    record, screen, routes = prepare(
        tmp_path,
        [excluded],
        [route()],
        continuation=[route(exclusion_class="deposited_output_only", routes=[])],
    )
    results = tmp_path / "results"
    results.mkdir()
    summary = render(record, screen, routes, results)
    assert summary["main_sample_candidates"] == {}
    assert summary["strata_without_main_sample_candidate"] == ["Physics_Engineering"]
    assert [r["paper_id"] for r in summary["excluded_after_deposit_inspection"]] == ["PMID:123"]
    assert len(summary["route_records"]) == 2


def test_continuation_must_hash_chain_to_previous_record(tmp_path: Path) -> None:
    record, screen, routes = prepare(tmp_path, [assessment()], [route()], continuation=[])
    payload = json.loads(routes[1].read_text())
    payload["continues_from_sha256"] = "0" * 64
    routes[1].write_text(json.dumps(payload))
    with pytest.raises(ValueError, match="hash-chain"):
        render(record, screen, routes, tmp_path)
