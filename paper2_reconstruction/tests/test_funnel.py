from pathlib import Path

import pytest

from paper2.build import ROOT, funnel_queue, protect_funnel
from paper2.core import read_csv, write_csv


def test_funnel_has_one_row_per_paper_and_keeps_source_memberships() -> None:
    frame = read_csv(ROOT / "data/derived/inference_frame.csv")
    queue = read_csv(ROOT / "data/derived/funnel_NOT_ASSESSED.csv")
    assert queue == funnel_queue(frame, paper_level=True)
    assert len({row["paper_id"] for row in queue}) == len(queue) == len(frame)
    assert sum(len(row["source_record_ids"].split(";")) for row in queue) == len(
        read_csv(ROOT / "data/derived/source_record_bridge.csv")
    )


@pytest.mark.parametrize("paper_level", [True, False])
def test_modified_current_or_legacy_funnel_is_never_overwritten(
    tmp_path: Path, paper_level: bool
) -> None:
    current = funnel_queue(
        read_csv(ROOT / "data/derived/inference_frame.csv")[:1], paper_level=True
    )
    legacy = funnel_queue(
        read_csv(ROOT / "data/derived/source_record_bridge.csv")[:1], paper_level=False
    )
    path = tmp_path / "queue.csv"
    rows = current if paper_level else legacy
    write_csv(path, rows, list(rows[0]))
    protect_funnel(path, current, legacy)
    changed = [dict(rows[0], G1_computationally_testable="yes", evidence_id="test-evidence")]
    write_csv(path, changed, list(changed[0]))
    original = path.read_bytes()
    with pytest.raises(ValueError, match="modified funnel"):
        protect_funnel(path, current, legacy)
    assert path.read_bytes() == original
