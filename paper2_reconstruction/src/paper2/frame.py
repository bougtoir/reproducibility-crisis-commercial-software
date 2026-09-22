import json
from pathlib import Path

from paper2.core import Row, assign_field, sha256, utc_time


def confirm_frame(
    registry: list[Row], decision_path: Path, source_commit: str, source_sha256: str
) -> tuple[list[Row], dict[str, object]]:
    decision: dict[str, str] = json.loads(decision_path.read_text())
    required = {
        "decision": "approved_unique_pmid",
        "source_commit": source_commit,
        "source_sha256": source_sha256,
        "identity": "PMID",
        "original_rows": "preserve_all",
        "scope": "fixed_deposited_corpus_only",
        "protocol_freeze": "not_implied",
    }
    if any(decision.get(key) != value for key, value in required.items()):
        raise ValueError("Author decision does not approve this exact source and identity")
    utc_time(decision["recorded_at_utc"])
    if not registry or len({row["paper_id"] for row in registry}) != len(registry):
        raise ValueError("The inference frame must contain unique papers")
    frame = [
        {
            **row,
            "sampling_stratum": assign_field(row["paper_id"], row["epj_fields"].split(";")),
        }
        for row in sorted(registry, key=lambda row: row["paper_id"])
    ]
    manifest: dict[str, object] = {
        "status": decision["decision"],
        "author_decision_file": "data/frame_decision.json",
        "author_decision_sha256": sha256(decision_path.read_bytes()),
        "decision_recorded_at_utc": decision["recorded_at_utc"],
        "source_commit": source_commit,
        "source_sha256": source_sha256,
        "identity": "PMID",
        "paper_count": len(frame),
        "source_record_count": sum(int(row["record_multiplicity"]) for row in frame),
        "field_assignment": "minimum SHA256(paper2-field-v1|paper_id|recorded_field)",
        "scope": "fixed_deposited_corpus_only",
        "protocol_freeze": "not_implied",
    }
    return frame, manifest
