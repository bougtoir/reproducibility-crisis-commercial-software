import pytest

from paper2.candidate_review import Source, validate_review


def test_candidate_review_rejects_unbound_or_invented_evidence() -> None:
    source = Source(
        "hash", "text-hash", "title_and_abstract_only", "A synthetic test source passage."
    )
    record: dict[str, object] = {
        "computationally_testable": "uncertain",
        "principal_target_identifiable": "uncertain",
        "target_candidate": "unknown",
        "inputs_required": "unknown",
        "access_observation": "unknown",
        "specification_observation": "unknown",
        "evidence": [
            {
                "source_sha256": "hash",
                "quote": "A synthetic test source passage.",
                "observation": "Unit-test fixture only.",
            }
        ],
    }
    assert validate_review(record, [source]) == record
    for evidence in [
        [],
        [{"source_sha256": "wrong", "quote": source.text, "observation": "test"}],
        [{"source_sha256": "hash", "quote": "Invented source passage.", "observation": "test"}],
    ]:
        with pytest.raises(ValueError):
            validate_review({**record, "evidence": evidence}, [source])
