"""Prespecified (SAP.md) descriptive analysis of the sealed, mechanically adjudicated main study.

Inputs are the frozen cohort record, the sealed blind outcome and the delegated primary
adjudication. Outputs are descriptive tables and the finite-population policy envelope.
No model fitting, perturbation or alternative-implementation analysis is performed.
Every estimate inherits the status of its inputs: delegated mechanical adjudication with
human adjudication and investigator verification pending.
"""

from __future__ import annotations

import json
from collections import Counter
from dataclasses import asdict
from pathlib import Path

from paper2.core import sha256, wilson, write_csv
from paper2.estimation import Estimate, Stratum, stratified_policy_estimate

ROOT = Path(__file__).resolve().parents[2]
STUDY_DIR = ROOT / "data" / "adjudication" / "main-study-20260925"
COHORT = ROOT / "data" / "adjudication" / "main_cohort_20260925.json"
BLIND = STUDY_DIR / "blind_outcome.json"
ADJUDICATION = STUDY_DIR / "primary_adjudication.json"
RESULTS = ROOT / "results"

LEVELS = (
    "L1_specification_extracted",
    "L2_implementation_created",
    "L3_execution_successful",
    "L4_numerical_target_reproduced",
    "L5_conclusion_preserved",
)
POLICY_NON_SUCCESS_STATES = {
    "external_reference_resource_required_not_served",
    "gate_failed_delegated",
    "required_inputs_not_in_public_listing",
    "specification_not_obtained",
    "target_value_not_located_verbatim",
}
ACQUISITION_NON_SUCCESS = {"all_inputs_leak_target", "no_input_retained"}
STATUS = (
    "delegated_mechanical_adjudication; human_adjudication_pending; "
    "investigator_verification_pending"
)


def mapping(value: object) -> dict[str, object]:
    if not isinstance(value, dict):
        raise ValueError("expected a JSON object")
    return {str(k): v for k, v in value.items()}


def listed(value: object) -> list[dict[str, object]]:
    if not isinstance(value, list):
        raise ValueError("expected a JSON list")
    return [mapping(v) for v in value]


def paper_disposition(paper: dict[str, object], endpoints: dict[str, dict[str, object]]) -> str:
    """Map a cohort paper to its policy-endpoint disposition under the frozen SAP wording."""
    pid = str(paper["paper_id"])
    if paper["state"] == "slots_completed":
        return f"attempted_{endpoints[pid]['primary_endpoint_majority']}"
    eligibility = str(paper["eligibility"])
    if eligibility in POLICY_NON_SUCCESS_STATES:
        return f"not_attempted_{eligibility}"
    if paper["acquisition"] in ACQUISITION_NON_SUCCESS:
        return f"not_attempted_{paper['acquisition']}"
    raise ValueError(f"{pid}: unclassified non-run state")


def analyse() -> dict[str, object]:
    cohort = mapping(json.loads(COHORT.read_bytes()))
    sealed = mapping(json.loads(BLIND.read_bytes()))
    adjudication = mapping(json.loads(ADJUDICATION.read_bytes()))
    if adjudication["blind_outcome_sha256"] != sha256(BLIND.read_bytes()):
        raise ValueError("adjudication does not reference the sealed blind outcome")
    papers = listed(sealed["papers"])
    runs = listed(adjudication["runs"])
    endpoints = {str(p["paper_id"]): p for p in listed(adjudication["papers"])}
    strata = listed(cohort["strata"])
    eligible = int(str(cohort["eligible_population"]))
    if len(papers) != int(str(cohort["selected_n"])):
        raise ValueError("sealed cohort size differs from the selection record")

    dispositions = {str(p["paper_id"]): paper_disposition(p, endpoints) for p in papers}
    disposition_counts = Counter(dispositions.values())

    def stratum(record: dict[str, object], unresolved_states: set[str]) -> Stratum:
        name = record["sampling_stratum"]
        members = [p for p in papers if p["sampling_stratum"] == name]
        successes = sum(dispositions[str(p["paper_id"])] == "attempted_success" for p in members)
        unresolved = sum(
            dispositions[str(p["paper_id"])] in unresolved_states
            or dispositions[str(p["paper_id"])] == "attempted_unknown"
            for p in members
        )
        return Stratum(int(str(record["N_h"])), len(members), successes, unresolved)

    not_attempted = {d for d in disposition_counts if d.startswith("not_attempted_")}
    primary = stratified_policy_estimate([stratum(s, set()) for s in strata], eligible, 0, 0)
    verification_pending = stratified_policy_estimate(
        [stratum(s, not_attempted) for s in strata], eligible, 0, 0
    )

    attempted = [p for p in papers if p["state"] == "slots_completed"]
    attempted_success = sum(
        endpoints[str(p["paper_id"])]["primary_endpoint_majority"] == "success" for p in attempted
    )
    conditional_lower, conditional_upper = wilson(attempted_success, len(attempted))

    level_counts = {level: Counter(str(r[level]) for r in runs) for level in LEVELS}
    stop_counts = Counter(str(r["stop_reason"]) for r in runs)
    slot_state_counts = Counter(str(r["slot_state"]) for r in runs)
    code_runs: Counter[str] = Counter()
    code_papers: dict[str, set[str]] = {}
    for run in runs:
        raw_codes = run["failure_codes"]
        for code in (str(c) for c in raw_codes) if isinstance(raw_codes, list) else ():
            code_runs[code] += 1
            code_papers.setdefault(code, set()).add(str(run["paper_id"]))
    sensitivity = {
        key: Counter(str(endpoints[str(p["paper_id"])][key]) for p in attempted)
        for key in ("strict", "majority", "permissive")
    }

    def estimate_row(label: str, e: Estimate) -> dict[str, object]:
        return {"analysis": label, **asdict(e)}

    return {
        "study": adjudication["study"],
        "sap": "protocols/SAP.md (frozen by AMEND-2026-09-25-05)",
        "status": STATUS,
        "inputs": {
            "cohort_selection": cohort["selection_id"],
            "blind_outcome_sha256": sha256(BLIND.read_bytes()),
            "primary_adjudication_sha256": sha256(ADJUDICATION.read_bytes()),
        },
        "denominators": {
            "route_eligible_population_E": eligible,
            "sampled_main_cohort": len(papers),
            "attempted_papers_complete_triples": len(attempted),
            "not_attempted_papers": len(papers) - len(attempted),
            "runs": len(runs),
            "prospective_pilot_papers_excluded": 7,
            "accessibility_gate_failed_papers_excluded": 7,
        },
        "paper_dispositions": dict(sorted(disposition_counts.items())),
        "primary_endpoint_majority_among_attempted": dict(
            Counter(
                str(endpoints[str(p["paper_id"])]["primary_endpoint_majority"]) for p in attempted
            )
        ),
        "sensitivity_thresholds_among_attempted": {k: dict(v) for k, v in sensitivity.items()},
        "conditional_rate_among_attempted": {
            "successes": attempted_success,
            "attempted": len(attempted),
            "wilson_95_lower": conditional_lower,
            "wilson_95_upper": conditional_upper,
            "label": "secondary; model-based unweighted binomial reference, not the headline",
        },
        "policy_estimate_delegated_states_as_recorded": estimate_row("primary_policy", primary),
        "policy_estimate_non_attempted_as_unresolved": estimate_row(
            "verification_pending_envelope", verification_pending
        ),
        "run_level": {
            "levels": {level: dict(counts) for level, counts in level_counts.items()},
            "stop_reasons": dict(stop_counts),
            "slot_states": dict(slot_state_counts),
            "failure_codes_runs": dict(sorted(code_runs.items())),
            "failure_codes_papers": {k: len(v) for k, v in sorted(code_papers.items())},
        },
        "not_performed_by_design": [
            "multiverse",
            "alternative_reasonable_implementations",
            "robustness_perturbation",
            "original_code_rerun",
            "random_intercept_model (fewer than 20 papers with successes)",
        ],
    }


def write_tables(analysis: dict[str, object]) -> None:
    dispositions = mapping(analysis["paper_dispositions"])
    write_csv(
        RESULTS / "main_paper_dispositions.csv",
        [{"disposition": k, "papers": v, "status": STATUS} for k, v in dispositions.items()],
        ["disposition", "papers", "status"],
    )
    run_level = mapping(analysis["run_level"])
    rows: list[dict[str, object]] = []
    for level, counts in mapping(run_level["levels"]).items():
        for state, n in sorted(mapping(counts).items()):
            rows.append({"level": level, "state": state, "runs": n})
    write_csv(RESULTS / "main_run_levels.csv", rows, ["level", "state", "runs"])
    code_runs = mapping(run_level["failure_codes_runs"])
    code_papers = mapping(run_level["failure_codes_papers"])
    write_csv(
        RESULTS / "main_failure_codes.csv",
        [{"code": c, "runs": code_runs[c], "papers": code_papers[c]} for c in code_runs],
        ["code", "runs", "papers"],
    )
    estimates = [
        mapping(analysis["policy_estimate_delegated_states_as_recorded"]),
        mapping(analysis["policy_estimate_non_attempted_as_unresolved"]),
    ]
    write_csv(RESULTS / "main_policy_estimates.csv", estimates, list(estimates[0]))


def main() -> None:
    analysis = analyse()
    (RESULTS / "main_analysis.json").write_text(json.dumps(analysis, indent=2) + "\n")
    write_tables(analysis)
    print(json.dumps({k: analysis[k] for k in ("denominators", "paper_dispositions")}, indent=2))


if __name__ == "__main__":
    main()
