import json

from paper2 import reveal_ledger as r
from paper2.core import sha256, utc_time


def test_recorded_ledger_postdates_seal_and_covers_every_attempted_paper() -> None:
    ledger = json.loads(r.LEDGER.read_bytes())
    sealed = json.loads(r.BLIND.read_bytes())
    assert ledger["blind_outcome_sha256"] == sha256(r.BLIND.read_bytes())
    assert ledger["primary_adjudication_sha256"] == sha256(r.ADJUDICATION.read_bytes())
    sealed_at = utc_time(sealed["sealed_at_utc"])
    assert utc_time(ledger["recorded_at_utc"]) > sealed_at
    attempted = {run["paper_id"] for run in sealed["runs"]}
    assert {p["paper_id"] for p in ledger["papers"]} == attempted
    assert ledger["original_code_executed"] is False
    assert ledger["blind_scores_revised_after_reveal"] is False
    for paper in ledger["papers"]:
        for source in paper["sources"]:
            if not source["acquired_before_seal_withheld_from_solver"]:
                assert utc_time(source["retrieved_at_utc"]) > sealed_at
        for snippet in paper["snippets"]:
            assert snippet["verified_verbatim"] is True
        if paper["reveal_assessment"] == "missing":
            assert paper["sources"] == [] and paper["dimensions"] == {}


def test_described_papers_cover_every_dimension_with_scaled_completeness() -> None:
    ledger = json.loads(r.LEDGER.read_bytes())
    observations = json.loads(r.OBSERVATIONS.read_bytes())
    scale = set(observations["completeness_scale"])
    for paper in ledger["papers"]:
        if paper["reveal_assessment"] == "described":
            assert set(paper["dimensions"]) == set(observations["dimensions"])
            assert {d["completeness"] for d in paper["dimensions"].values()} <= scale
