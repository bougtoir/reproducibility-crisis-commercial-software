import pytest

from paper2.candidate_review import Source, normalize, source_segments, validate_review


def test_segments_preserve_source_language_and_round_trip_whitespace() -> None:
    text = "Ｓｈａｎｎｏｎ 多样性指数 （Ｈ′）\n保留原文。 " * 40
    segments = source_segments(text)
    assert " ".join(segments.values()) == normalize(text)
    assert list(segments) == [f"S{index:04d}" for index in range(1, len(segments) + 1)]
    assert all(segment in normalize(text) for segment in segments.values())


def test_segment_evidence_is_resolved_from_source_and_unknown_ids_rejected() -> None:
    text = normalize("Ｓｈａｎｎｏｎ 多样性指数 （Ｈ′） 保留原文。 " * 40)
    source = Source("source-hash", "text-hash", "synthetic_only", text)
    evidence = [
        {
            "source_sha256": source.source_sha256,
            "segment_id": segment_id,
            "observation": "Synthetic unit-test observation only.",
        }
        for segment_id in ("S0001", "S0002")
    ]
    record: dict[str, object] = {
        "computationally_testable": "uncertain",
        "principal_target_identifiable": "uncertain",
        "target_candidate": "unknown",
        "inputs_required": "unknown",
        "access_observation": "unknown",
        "specification_observation": "unknown",
        "evidence": evidence,
    }
    resolved = validate_review(record, [source])
    assert resolved["evidence"] == [
        {**item, "quote": source_segments(text)[item["segment_id"]]} for item in evidence
    ]
    for invalid in ("S0000", "S9999", "../S0001"):
        with pytest.raises(ValueError, match="unregistered segment"):
            validate_review(
                {**record, "evidence": [{**item, "segment_id": invalid} for item in evidence]},
                [source],
            )
