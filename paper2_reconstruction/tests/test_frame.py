import json
from pathlib import Path

import pytest

from paper2.build import PINS, REVISION, ROOT
from paper2.core import assign_field, read_csv, sha256
from paper2.frame import confirm_frame


def test_confirmed_frame_preserves_exact_source_identity_and_decision() -> None:
    registry = read_csv(ROOT / "data/derived/paper_registry.csv")
    frame, manifest = confirm_frame(
        registry, ROOT / "data/frame_decision.json", REVISION, PINS["output/extracted_data.csv"]
    )
    generated = read_csv(ROOT / "data/derived/inference_frame.csv")
    assert generated == frame
    assert len(frame) == len(registry)
    assert {row["paper_id"] for row in frame} == {row["paper_id"] for row in registry}
    assert all(
        row["sampling_stratum"] == assign_field(row["paper_id"], row["epj_fields"].split(";"))
        for row in frame
    )
    assert manifest["source_record_count"] == len(
        read_csv(ROOT / "data/derived/source_record_bridge.csv")
    )
    assert manifest["protocol_freeze"] == "not_implied"
    assert confirm_frame(
        list(reversed(registry)),
        ROOT / "data/frame_decision.json",
        REVISION,
        PINS["output/extracted_data.csv"],
    ) == (frame, manifest)
    status = json.loads((ROOT / "results/readiness.json").read_text())
    assert status["paper_frame_decision"] == "approved_unique_pmid"
    assert status["distinct_10000_paper_precondition"] == "not_met"
    assert status["paper_frame_sha256"] == sha256(
        (ROOT / "data/derived/inference_frame.csv").read_bytes()
    )


def test_frame_rejects_a_decision_for_a_different_source(tmp_path: Path) -> None:
    decision = json.loads((ROOT / "data/frame_decision.json").read_text())
    decision["source_sha256"] = "0" * 64
    path = tmp_path / "decision.json"
    path.write_text(json.dumps(decision))
    with pytest.raises(ValueError, match="exact source"):
        confirm_frame([], path, REVISION, PINS["output/extracted_data.csv"])
