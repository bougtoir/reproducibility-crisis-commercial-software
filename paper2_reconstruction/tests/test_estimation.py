from math import comb, fsum

import pytest

from paper2.estimation import Stratum, stratified_policy_estimate, success_total_interval


def test_exact_small_population_interval_coverage_by_enumeration() -> None:
    alpha = 0.1
    for population in range(1, 13):
        for sample in range(1, population + 1):
            intervals = [
                success_total_interval(x, sample, population, alpha) for x in range(sample + 1)
            ]
            for successes in range(population + 1):
                missed = fsum(
                    comb(successes, x)
                    * comb(population - successes, sample - x)
                    / comb(population, sample)
                    for x in range(
                        max(0, sample - population + successes), min(sample, successes) + 1
                    )
                    if not intervals[x][0] <= successes <= intervals[x][1]
                )
                assert missed <= alpha + 1e-12


def test_boundary_and_census_strata_have_correct_support() -> None:
    assert success_total_interval(0, 10, 10) == (0, 0)
    assert success_total_interval(10, 10, 10) == (10, 10)
    assert success_total_interval(4, 10, 10) == (4, 4)
    low, high = success_total_interval(0, 10, 100)
    assert low == 0 and 0 < high < 100
    assert success_total_interval(10, 10, 100) == (100 - high, 100)


def test_unknown_slots_and_gates_widen_bounds_without_dropping_barriers() -> None:
    result = stratified_policy_estimate(
        [Stratum(10, 10, 3, 2), Stratum(20, 20, 8, 0)],
        eligible_population=50,
        unknown_gates=5,
        verified_barriers=15,
    )
    assert result.lower_success_total == result.lower_sampling_total == 11
    assert result.upper_success_total == result.upper_sampling_total == 18
    assert result.lower_policy_rate == 11 / 50
    assert result.upper_policy_rate == 18 / 50
    assert result.interval_label == "sampling_and_identification_envelope"


def test_unequal_stratum_weights_are_preserved() -> None:
    result = stratified_policy_estimate(
        [Stratum(100, 10, 5, 0), Stratum(900, 10, 1, 0)],
        1000,
        0,
        0,
    )
    assert result.lower_success_total == result.upper_success_total == 140
    assert result.lower_policy_rate == 0.14
    assert result.lower_sampling_policy_rate <= 0.14 <= result.upper_sampling_policy_rate
    assert result.interval_label == "simultaneous_hypergeometric_sampling_interval"


def test_barrier_only_and_unknown_only_domains_are_not_invented_runs() -> None:
    assert stratified_policy_estimate([], 10, 0, 10).upper_policy_rate == 0
    result = stratified_policy_estimate([], 10, 10, 0)
    assert result.lower_policy_rate == 0
    assert result.upper_policy_rate == 1


@pytest.mark.parametrize("counts", [(0, 0, 0, 0), (10, 0, 0, 0), (10, 5, 4, 2)])
def test_invalid_strata_are_rejected(counts: tuple[int, int, int, int]) -> None:
    with pytest.raises(ValueError, match="stratum"):
        Stratum(*counts).validate()


def test_denominator_mismatch_is_rejected() -> None:
    with pytest.raises(ValueError, match="partition"):
        stratified_policy_estimate([Stratum(10, 5, 3, 0)], 20, 5, 4)
