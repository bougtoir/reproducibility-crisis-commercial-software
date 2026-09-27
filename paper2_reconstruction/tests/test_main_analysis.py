import json

import pytest

from paper2 import main_analysis as m
from paper2.build import ROOT
from paper2.core import sha256


def test_analysis_matches_sealed_records_and_excludes_pilot_layers() -> None:
    analysis = m.analyse()
    sealed = json.loads(m.BLIND.read_bytes())
    adjudication = json.loads(m.ADJUDICATION.read_bytes())
    denominators = analysis["denominators"]
    assert denominators["sampled_main_cohort"] == len(sealed["papers"])
    assert denominators["runs"] == len(adjudication["runs"])
    assert sum(analysis["paper_dispositions"].values()) == denominators["sampled_main_cohort"]
    assert denominators["prospective_pilot_papers_excluded"] == 7
    assert analysis["inputs"]["blind_outcome_sha256"] == sha256(m.BLIND.read_bytes())
    assert "human_adjudication_pending" in analysis["status"]


def test_policy_estimates_are_bounds_not_point_claims() -> None:
    analysis = m.analyse()
    primary = analysis["policy_estimate_delegated_states_as_recorded"]
    envelope = analysis["policy_estimate_non_attempted_as_unresolved"]
    assert primary["eligible_population"] == 461
    assert primary["lower_sampling_policy_rate"] <= primary["upper_sampling_policy_rate"]
    assert envelope["upper_sampling_policy_rate"] >= primary["upper_sampling_policy_rate"]
    assert envelope["interval_label"] == "sampling_and_identification_envelope"


def test_unclassified_non_run_state_is_rejected() -> None:
    paper = {
        "paper_id": "PMID:0",
        "state": "not_run_gate_or_input",
        "eligibility": "specifiable",
        "acquisition": None,
    }
    with pytest.raises(ValueError, match="unclassified"):
        m.paper_disposition(paper, {})


def test_results_tables_are_regenerated_from_analysis() -> None:
    analysis = json.loads((ROOT / "results/main_analysis.json").read_text())
    assert analysis["paper_dispositions"] == m.analyse()["paper_dispositions"]
