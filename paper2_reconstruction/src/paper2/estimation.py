from dataclasses import dataclass
from math import comb, fsum


def hypergeometric_tail(
    observed: int, sample: int, population: int, successes: int, *, upper: bool
) -> float:
    support_low = max(0, sample - population + successes)
    support_high = min(sample, successes)
    start = max(observed, support_low) if upper else support_low
    stop = support_high if upper else min(observed, support_high)
    denominator = comb(population, sample)
    return fsum(
        comb(successes, count) * comb(population - successes, sample - count) / denominator
        for count in range(start, stop + 1)
    )


def success_total_interval(
    observed: int, sample: int, population: int, alpha: float = 0.05
) -> tuple[int, int]:
    if not 0 <= observed <= sample <= population or sample < 1 or not 0 < alpha < 1:
        raise ValueError("Invalid hypergeometric interval inputs")
    low, high = observed, population - sample + observed
    left, right = low, high
    while left < right:
        middle = (left + right) // 2
        if hypergeometric_tail(observed, sample, population, middle, upper=True) >= alpha / 2:
            right = middle
        else:
            left = middle + 1
    lower = left
    left, right = low, high
    while left < right:
        middle = (left + right + 1) // 2
        if hypergeometric_tail(observed, sample, population, middle, upper=False) >= alpha / 2:
            left = middle
        else:
            right = middle - 1
    return lower, left


@dataclass(frozen=True)
class Stratum:
    population: int
    sample: int
    successes: int
    unresolved: int

    def validate(self) -> None:
        if (
            self.population < 1
            or not 1 <= self.sample <= self.population
            or self.successes < 0
            or self.unresolved < 0
            or self.successes + self.unresolved > self.sample
        ):
            raise ValueError("Invalid stratum counts")


@dataclass(frozen=True)
class Estimate:
    eligible_population: int
    attemptable_population: int
    unknown_gates: int
    verified_barriers: int
    lower_success_total: float
    upper_success_total: float
    lower_sampling_total: int
    upper_sampling_total: int
    lower_policy_rate: float
    upper_policy_rate: float
    lower_sampling_policy_rate: float
    upper_sampling_policy_rate: float
    interval_label: str


def stratified_policy_estimate(
    strata: list[Stratum],
    eligible_population: int,
    unknown_gates: int,
    verified_barriers: int,
    alpha: float = 0.05,
) -> Estimate:
    for stratum in strata:
        stratum.validate()
    attemptable = sum(stratum.population for stratum in strata)
    if (
        eligible_population < 1
        or unknown_gates < 0
        or verified_barriers < 0
        or attemptable + unknown_gates + verified_barriers != eligible_population
        or not 0 < alpha < 1
    ):
        raise ValueError("Eligible denominator must partition exactly into recorded gate domains")
    lower = fsum(s.population * s.successes / s.sample for s in strata)
    upper = fsum(s.population * (s.successes + s.unresolved) / s.sample for s in strata)
    lower_sampling, upper_sampling = 0, unknown_gates
    for stratum in strata:
        lower_sampling += success_total_interval(
            stratum.successes, stratum.sample, stratum.population, alpha / len(strata)
        )[0]
        upper_sampling += success_total_interval(
            stratum.successes + stratum.unresolved,
            stratum.sample,
            stratum.population,
            alpha / len(strata),
        )[1]
    upper += unknown_gates
    return Estimate(
        eligible_population,
        attemptable,
        unknown_gates,
        verified_barriers,
        lower,
        upper,
        lower_sampling,
        upper_sampling,
        lower / eligible_population,
        upper / eligible_population,
        lower_sampling / eligible_population,
        upper_sampling / eligible_population,
        (
            "sampling_and_identification_envelope"
            if unknown_gates or any(s.unresolved for s in strata)
            else "simultaneous_hypergeometric_sampling_interval"
        ),
    )
