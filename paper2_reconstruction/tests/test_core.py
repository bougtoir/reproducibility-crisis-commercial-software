from decimal import Decimal
from pathlib import Path

import pytest

from paper2.core import (
    Candidate,
    assign_field,
    fixed_triple,
    gate_counts,
    paper_registry,
    rounded_agreement,
    run_success,
    snapshot,
    stratified_sample,
    validate_levels,
    validate_sequence,
    wilson,
)


def fixture_row(pmid: str, field: str) -> dict[str, str]:
    return {
        "pmid": pmid,
        "stratum": field,
        "doi": "",
        "pub_year": "2024",
        "code_available": "False",
        "data_available": "False",
        "has_pmc_fulltext": "False",
        "has_commercial_software": "False",
        "has_opensource_software": "False",
    }


def test_registry_preserves_duplicate_rows_and_overlapping_memberships() -> None:
    rows = [
        fixture_row("101", "Biomedical_Basic"),
        fixture_row("101", "Biomedical_Basic"),
        fixture_row("101", "Clinical_Medicine"),
        fixture_row("102", "Clinical_Medicine"),
    ]
    papers, bridge = paper_registry(rows)
    assert len(bridge) == 4
    assert len(papers) == 2
    assert papers[0]["record_multiplicity"] == "3"
    assert papers[0]["epj_fields"] == "Biomedical_Basic;Clinical_Medicine"
    assert len({row["source_record_id"] for row in bridge}) == 4
    assert papers[0]["epj_data_statement_detected"] == "False"


def test_conflicting_duplicate_fails_instead_of_silent_first_row() -> None:
    first = fixture_row("101", "Biomedical_Basic")
    second = fixture_row("101", "Clinical_Medicine") | {"doi": "different"}
    with pytest.raises(ValueError, match="Conflicting"):
        paper_registry([first, second])


def test_raw_snapshot_never_overwritten(tmp_path: Path) -> None:
    path = tmp_path / "raw"
    snapshot(path, b"original")
    snapshot(path, b"original")
    with pytest.raises(ValueError, match="Immutable"):
        snapshot(path, b"new")
    assert path.read_bytes() == b"original"


@pytest.mark.parametrize(
    ("states", "expected"),
    [
        (["yes", "yes", "yes"], "success"),
        (["yes", "yes", "not_assessable"], "success"),
        (["yes", "no", "not_assessable"], "unknown"),
        (["no", "no", "unknown"], "failure"),
        (["unknown", "unknown", "unknown"], "unknown"),
    ],
)
def test_fixed_slots_do_not_count_missing_as_failure(states, expected) -> None:
    assert fixed_triple(states) == expected


def test_thresholds_and_incomplete_triple() -> None:
    assert fixed_triple(["yes", "no", "unknown"], threshold=1) == "success"
    assert fixed_triple(["yes", "no", "unknown"], threshold=3) == "failure"
    with pytest.raises(ValueError):
        fixed_triple(["yes", "yes"])


def test_conclusion_agreement_can_survive_numerical_disagreement() -> None:
    validate_levels("yes", "yes", "no", "yes")
    assert run_success("yes", "yes", "no", "yes", True) == "no"
    assert run_success("yes", "yes", "yes", "yes", False) == "not_assessable"
    with pytest.raises(ValueError):
        validate_levels("no", "yes", "yes", "yes")
    with pytest.raises(ValueError):
        validate_levels("no", "no", "no", "no")
    assert run_success("no", "not_assessable", "not_assessable", "not_assessable", True) == "no"


def test_unobserved_funnel_is_not_a_zero_success_result() -> None:
    result = gate_counts(["not_assessed"] * 10)
    assert result["yes_among_known"] == "not_assessable"
    assert result["identification_lower"] == "0.0"
    assert result["identification_upper"] == "1.0"
    assert result["unresolved"] == "10"


def test_precision_uses_reported_decimal_places() -> None:
    assert rounded_agreement("1.234", "1.23", 2)
    assert not rounded_agreement("1.236", "1.23", 2)
    assert rounded_agreement("0.0004", "0.000", 3)
    assert not rounded_agreement("0.0006", "0.000", 3)
    with pytest.raises(ValueError):
        rounded_agreement("NaN", "1.23", 2)


def test_temporal_order_rejects_posthoc_target_and_early_reveal() -> None:
    times = [f"2026-09-22T12:0{i}:00Z" for i in range(5)]
    validate_sequence(*times)
    with pytest.raises(ValueError):
        validate_sequence(times[1], times[0], *times[2:])
    with pytest.raises(ValueError):
        validate_sequence(*times[:3], times[4], times[3])
    with pytest.raises(ValueError):
        validate_sequence(*(value.removesuffix("Z") for value in times))


def test_sampling_is_order_invariant_unique_weighted_and_stratified() -> None:
    candidates = [
        Candidate(str(i), "Biomedical_Basic" if i < 16 else "Clinical_Medicine") for i in range(20)
    ]
    selected = stratified_sample(candidates, 10, "prespecified-test-seed")
    assert selected == stratified_sample(list(reversed(candidates)), 10, "prespecified-test-seed")
    assert len({r["paper_id"] for r in selected}) == 10
    assert {r["sampling_stratum"] for r in selected} == {
        "Biomedical_Basic",
        "Clinical_Medicine",
    }
    assert sum(Decimal(r["weight"]) for r in selected) == 20
    with pytest.raises(ValueError):
        stratified_sample([*candidates, candidates[0]], 10, "seed")


def test_overlapping_memberships_get_stable_recorded_assignment() -> None:
    fields = ["Biomedical_Basic", "Clinical_Medicine", "Computational_Science"]
    assigned = assign_field("PMID:101", fields)
    assert assigned in fields
    assert assigned == assign_field("PMID:101", list(reversed(fields)) + [fields[0]])
    with pytest.raises(ValueError):
        assign_field("PMID:101", [])


def test_wilson_known_boundary_and_symmetry() -> None:
    lower, upper = wilson(50, 100)
    assert lower == pytest.approx(0.40383153)
    assert upper == pytest.approx(1 - lower)
    assert wilson(0, 100)[0] == 0
    assert wilson(100, 100)[1] == 1
    with pytest.raises(ValueError):
        wilson(0, 0)
