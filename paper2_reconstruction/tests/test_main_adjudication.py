import pytest

from paper2 import main_adjudication as m
from paper2.core import sha256

SPEC = {"target_values": ["89%"]}
FREEZE = {"numerical_agreement_criterion": "rounds to 89%"}


def events(*executions: tuple[str, str, int]) -> list[dict[str, object]]:
    out: list[dict[str, object]] = []
    for code, output, rc in executions:
        out.append({"kind": "code", "value": code})
        out.append(
            {
                "kind": "execution",
                "value": {"code_sha256": sha256(code.encode()), "output": output, "returncode": rc},
            }
        )
    return out


def run(stop: str = "final_report", report: object = None) -> dict[str, object]:
    return {
        "paper_id": "PMID:1",
        "slot": 1,
        "stop_reason": stop,
        "claimed_report": report,
        "contamination_events": "none_recorded",
    }


def slot(
    monkeypatch: pytest.MonkeyPatch, journal: list[dict[str, object]], **kwargs: object
) -> dict[str, object]:
    monkeypatch.setattr(m, "journal", lambda _p, _s: journal)
    return m.adjudicate_slot(run(**kwargs), SPEC, FREEZE)


def test_no_execution_report_is_a_slot_state_not_a_failure_code(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    out = slot(monkeypatch, [], report={"observed_value": "89%", "failure_codes": ["reasons"]})
    assert out["slot_state"] == "no_execution_report"
    assert out["L2_implementation_created"] == "not_reached"
    assert out["L3_execution_successful"] == "not_reached"
    assert out["failure_codes"] == []
    assert out["run_success"] == "no"


def test_report_only_value_is_not_an_observation(monkeypatch: pytest.MonkeyPatch) -> None:
    journal = events(("open('/input/deposit/a.csv')\nprint(139)", "139\n", 0))
    out = slot(monkeypatch, journal, report={"observed_value": "89% (139/157)"})
    assert out["component_provenance"]["89"] is None
    assert out["L3_execution_successful"] == "yes"
    assert out["L4_numerical_target_reproduced"] == "no"
    assert out["failure_codes"] == ["F20"]
    assert out["human_review_flag"] is True


def test_served_input_output_establishes_target(monkeypatch: pytest.MonkeyPatch) -> None:
    journal = events(("open('/input/deposit/a.csv')\nprint(88.5)", "88.5\n", 0))
    out = slot(monkeypatch, journal, report={"observed_value": "88.5%"})
    assert out["L4_numerical_target_reproduced"] == "yes"
    assert out["run_success"] == "yes"
    assert out["failure_codes"] == []


def test_approximate_output_does_not_establish_component(monkeypatch: pytest.MonkeyPatch) -> None:
    journal = events(("open('/input/deposit/a.csv')", "88.54\n", 0))
    out = slot(monkeypatch, journal, report={"observed_value": "88.5"})
    assert out["L3_execution_successful"] == "not_established"
    assert out["failure_codes"] == ["F15"]


def test_code_not_reading_served_inputs_is_not_provenance(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    journal = events(("print(88.5)", "88.5\n", 0))
    out = slot(monkeypatch, journal, report={"observed_value": "88.5"})
    assert out["L2_implementation_created"] == "no"
    assert out["run_success"] == "no"


def test_resource_limit_stop_is_f13(monkeypatch: pytest.MonkeyPatch) -> None:
    journal = events(("open('/input/deposit/a.csv')", "", 0))
    out = slot(monkeypatch, journal, stop="tool_limit")
    assert out["failure_codes"] == ["F13"]
    assert out["clean"] is True


def test_verbatim_missing_input_maps_to_f01_without_f99(monkeypatch: pytest.MonkeyPatch) -> None:
    journal = events(("import os; os.listdir('/input/deposit')", "a.fastq\n", 0))
    report = {"observed_value": None, "failure_codes": ["MISSING_REQUIRED_INPUT: no counts"]}
    out = slot(monkeypatch, journal, report=report)
    assert out["failure_codes"] == ["F01"]


def test_firewall_contamination_voids_slot(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(m, "journal", lambda _p, _s: [])
    contaminated = {**run(), "contamination_events": "original_artifact_served"}
    out = m.adjudicate_slot(contaminated, SPEC, FREEZE)
    assert out["clean"] is False
    assert out["run_success"] != "yes"


def test_paper_majority_endpoint_requires_two_valid_successes() -> None:
    def s(state: str, i: int) -> dict[str, object]:
        return {"paper_id": "PMID:1", "slot": i, "run_success": state}

    one = m.adjudicate_paper([s("no", 1), s("no", 2), s("yes", 3)])
    assert one["majority"] == "failure" and one["permissive"] == "success"
    two = m.adjudicate_paper([s("yes", 1), s("no", 2), s("yes", 3)])
    assert two["majority"] == "success" and two["strict"] == "failure"
    partial = m.adjudicate_paper([s("yes", 1), s("not_assessable", 2), s("no", 3)])
    assert partial["majority"] == "unknown" and partial["complete_triple"] is False
