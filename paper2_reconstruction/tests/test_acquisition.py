import json
from pathlib import Path
from unittest.mock import patch

import pytest

from paper2.acquisition import acquire, candidate_order
from paper2.build import ROOT
from paper2.core import FIELDS, read_csv, sha256, snapshot


def test_candidate_ranking_is_input_order_invariant_and_in_frame() -> None:
    frame = read_csv(ROOT / "data/derived/inference_frame.csv")
    candidates = candidate_order(frame, 1)
    assert candidates == candidate_order(list(reversed(frame)), 1)
    assert len(candidates) == len(FIELDS)
    assert {row["sampling_stratum"] for row in candidates} == set(FIELDS)
    assert len({row["paper_id"] for row in candidates}) == len(candidates)
    assert {row["paper_id"] for row in candidates} <= {row["paper_id"] for row in frame}
    with pytest.raises(ValueError, match="positive"):
        candidate_order(frame, 0)


def test_cached_acquisition_is_verified_and_corruption_is_not_refetched(tmp_path: Path) -> None:
    url = "https://example.org/synthetic"
    target = tmp_path / sha256(url.encode())
    snapshot(target / "body", b"synthetic")
    snapshot(
        target / "receipt.json",
        json.dumps(
            {
                "url": url,
                "sha256": sha256(b"synthetic"),
                "http_status": 200,
            }
        ).encode(),
    )
    with patch("paper2.acquisition.urlopen", side_effect=AssertionError("No network expected")):
        path, receipt = acquire(url, "synthetic", tmp_path)
        assert path.read_bytes() == b"synthetic"
        assert receipt["http_status"] == 200
        path.write_bytes(b"corrupted synthetic evidence")
        with pytest.raises(ValueError, match="immutable receipt"):
            acquire(url, "synthetic", tmp_path)
        assert path.read_bytes() == b"corrupted synthetic evidence"
