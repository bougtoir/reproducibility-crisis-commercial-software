"""Primary (delegated, mechanical) adjudication of the sealed main-study blind outcomes.

Applies the frozen rules of AMEND-2026-09-25-05 (tolerance rules, failure taxonomy,
PD-04 provenance rule) to the sealed `blind_outcome.json` without reading any
original implementation. Every decision is derived from retained artifacts:
the frozen specification (reported target values), the frozen paper record
(agreement criterion), the sealed slot result (claimed report) and the journaled
code/execution events of each slot. Human adjudication remains pending and is
recorded as such; this module never alters a rule after seeing an outcome.
"""

from __future__ import annotations

import argparse
import json
import re
from datetime import datetime, timezone
from decimal import Decimal, InvalidOperation
from pathlib import Path

from paper2.core import (
    FAILURES,
    PaperState,
    State,
    fixed_triple,
    rounded_agreement,
    run_success,
    sha256,
    snapshot,
    write_csv,
)
from paper2.timestamp import stamp

ROOT = Path(__file__).resolve().parents[2]
STUDY_DIR = ROOT / "data" / "adjudication" / "main-study-20260925"
BLIND = STUDY_DIR / "blind_outcome.json"
SPECS = STUDY_DIR / "specifications"
FREEZES = STUDY_DIR / "freezes"
RUNS = Path("/home/ubuntu/paper2_evidence/main-run-20260925")
RESULTS = ROOT / "results"
ADJUDICATION = STUDY_DIR / "primary_adjudication.json"
AMENDMENT = "AMEND-2026-09-25-07"

NUMBER = re.compile(r"(?<![\w.])[-+]?\d[\d,]*(?:\.\d+)?(?![\w.])")
SERVED_MARKERS = ("/input/deposit", "deposit/", "'deposit'", '"deposit"')
VERBATIM_PATTERNS: tuple[tuple[str, str], ...] = (
    (r"missing[_ ]?(required[_ ])?input|input[_ ]absent|not served", "F01"),
    (r"not installed|tool[_ ]unavailable|missing[_ ]required[_ ]resource", "F03"),
    (r"not[_ ]reproduced", "F20"),
)
STOP_TO_CODE = {
    "tool_limit": "F13",
    "wall_limit": "F13",
    "token_reserve_limit": "F13",
    "malformed_action_limit": "F15",
}


def mapping(value: object) -> dict[str, object]:
    if not isinstance(value, dict):
        raise ValueError("expected a JSON object")
    return {str(k): v for k, v in value.items()}


def listed(value: object) -> list[object]:
    if not isinstance(value, list):
        raise ValueError("expected a JSON list")
    return value


def utc() -> str:
    return datetime.now(timezone.utc).isoformat()


def parse_number(text: str) -> Decimal | None:
    try:
        return Decimal(text.replace(",", "").rstrip("%"))
    except InvalidOperation:
        return None


def decimal_places(text: str) -> int:
    cleaned = text.replace(",", "").rstrip("%")
    return len(cleaned.split(".")[1]) if "." in cleaned else 0


def numbers_in(text: str) -> list[str]:
    return [m.group(0) for m in NUMBER.finditer(text)]


def observed_components(value: object) -> list[str]:
    """Flatten a claimed observed value into numeric strings as written by the solver."""
    if value is None or isinstance(value, bool):
        return []
    if isinstance(value, int | float):
        return [repr(value) if isinstance(value, float) else str(value)]
    if isinstance(value, str):
        return numbers_in(value)
    if isinstance(value, dict):
        out: list[str] = []
        for v in value.values():
            out.extend(observed_components(v))
        return out
    if isinstance(value, list):
        out = []
        for v in value:
            out.extend(observed_components(v))
        return out
    return []


def within_bin(observed: str, reference: str) -> bool:
    if parse_number(observed) is None or parse_number(reference) is None:
        return False
    return rounded_agreement(
        str(parse_number(observed)), str(parse_number(reference)), decimal_places(reference)
    )


def journal(paper_id: str, slot: int) -> list[dict[str, object]]:
    slot_dir = RUNS / "runs" / paper_id.replace(":", "_") / f"slot-{slot}" / "controller"
    return [mapping(json.loads(p.read_bytes())) for p in sorted(slot_dir.glob("event-*.json"))]


def executions(events: list[dict[str, object]]) -> list[tuple[str, dict[str, object]]]:
    """Pair each journaled code event with the execution record that follows it."""
    pairs: list[tuple[str, dict[str, object]]] = []
    pending: str | None = None
    for event in events:
        if event["kind"] == "code":
            value = event["value"]
            pending = value if isinstance(value, str) else str(mapping(value).get("code", ""))
        elif event["kind"] == "execution" and pending is not None:
            record = mapping(event["value"])
            if sha256(pending.encode()) == record.get("code_sha256"):
                pairs.append((pending, record))
            pending = None
    return pairs


def reads_served_inputs(code: str) -> bool:
    return any(marker in code for marker in SERVED_MARKERS)


def same_number(left: str, right: str) -> bool:
    a, b = parse_number(left), parse_number(right)
    return a is not None and b is not None and a == b


def provenance(component: str, pairs: list[tuple[str, dict[str, object]]]) -> str | None:
    """Return the sha256 of the first served-input execution whose output prints the value.

    The claimed component must appear verbatim (numerically identical) in the output of a
    successful execution that reads served inputs; approximate matches are not accepted.
    """
    for code, record in pairs:
        if not reads_served_inputs(code) or record.get("returncode") != 0:
            continue
        for token in numbers_in(str(record.get("output", ""))):
            if same_number(token, component):
                return str(record["code_sha256"])
    return None


def taxonomy_codes(reported: object) -> tuple[list[str], list[str], list[str]]:
    """Split solver-reported reasons into taxonomy codes, verbatim text and mapped codes."""
    codes, verbatim, mapped = [], [], []
    for item in listed(reported) if isinstance(reported, list) else []:
        text = str(item)
        match = re.match(r"^(F\d\d)\b", text)
        if match and match.group(1) in FAILURES:
            codes.append(match.group(1))
            continue
        verbatim.append(text)
        for pattern, code in VERBATIM_PATTERNS:
            if re.search(pattern, text, re.IGNORECASE):
                mapped.append(code)
                break
    return codes, verbatim, mapped


def adjudicate_slot(
    run: dict[str, object], spec: dict[str, object], freeze: dict[str, object]
) -> dict[str, object]:
    paper_id, slot = str(run["paper_id"]), int(str(run["slot"]))
    events = journal(paper_id, slot)
    pairs = executions(events)
    report = run.get("claimed_report")
    report_map = mapping(report) if isinstance(report, dict) else None
    contamination = str(run.get("contamination_events", "none_recorded"))
    clean = contamination == "none_recorded" and str(run["stop_reason"]) in {
        "final_report",
        *STOP_TO_CODE,
    }
    served_code = [c for c, _ in pairs if reads_served_inputs(c)]
    targets = [str(t) for t in listed(spec["target_values"])]
    observed = observed_components(None if report_map is None else report_map.get("observed_value"))

    established: dict[str, str | None] = {c: provenance(c, pairs) for c in observed}
    established_values = [c for c, digest in established.items() if digest is not None]
    matched = {t: next((c for c in established_values if within_bin(c, t)), None) for t in targets}
    unestablished = [c for c, d in established.items() if d is None]

    l1: State = "yes" if report_map is not None else "no"
    l2: State = "yes" if served_code else "no"
    if not pairs:
        l3_label = "no_execution_report"
    elif established_values:
        l3_label = "yes"
    else:
        l3_label = "not_established"
    l3: State = "yes" if l3_label == "yes" else "no"
    if l3 == "yes":
        l4: State = "yes" if all(v is not None for v in matched.values()) else "no"
    else:
        l4 = "not_assessable"
    l5: State = "unknown" if l4 == "yes" else "not_assessable"
    success = run_success(l2, l3, l4, l5, clean)

    reported_codes, verbatim, mapped = taxonomy_codes(
        None if report_map is None else report_map.get("failure_codes")
    )
    derived: list[str] = []
    stop_code = STOP_TO_CODE.get(str(run["stop_reason"]))
    if success != "yes":
        if l3 == "yes" and l4 == "no":
            derived.append("F20")
        elif pairs and observed and unestablished:
            derived.append("F15")
        if stop_code:
            derived.append(stop_code)
        if pairs:
            derived.extend(c for c in mapped if c != "F20" or l3 == "yes")
        if pairs and verbatim and not reported_codes and not derived:
            derived.append("F99")
    if not pairs and stop_code is None:
        codes: list[str] = []
    else:
        codes = list(dict.fromkeys(reported_codes + derived))
    return {
        "paper_id": paper_id,
        "slot": slot,
        "stop_reason": run["stop_reason"],
        "clean": clean,
        "contamination_events": contamination,
        "journaled_executions": len(pairs),
        "served_input_executions": len(served_code),
        "claimed_observed_value": None if report_map is None else report_map.get("observed_value"),
        "observed_components": observed,
        "component_provenance": established,
        "frozen_target_values": targets,
        "target_component_matches": matched,
        "agreement_criterion": freeze["numerical_agreement_criterion"],
        "L1_specification_extracted": l1,
        "L2_implementation_created": "not_reached" if not pairs else l2,
        "L3_execution_successful": "not_reached" if not pairs else l3_label,
        "slot_state": l3_label if not pairs else "executed",
        "L4_numerical_target_reproduced": l4 if l3 == "yes" else "not_established",
        "L5_conclusion_preserved": (
            "pending_human_adjudication" if l5 == "unknown" else "not_assessable"
        ),
        "scored_levels": {"L2": l2, "L3": l3, "L4": l4, "L5": l5},
        "report_only_components": unestablished,
        "human_review_flag": bool(unestablished and established_values),
        "run_success": success,
        "failure_codes": codes,
        "reported_failure_codes_verbatim": verbatim,
        "verbatim_mapped_codes": mapped,
        "adjudication": "delegated_primary_mechanical",
        "human_adjudication": "pending",
    }


def adjudicate_paper(slots: list[dict[str, object]]) -> dict[str, object]:
    states = [str(s["run_success"]) for s in sorted(slots, key=lambda s: int(str(s["slot"])))]
    typed: list[State] = []
    for state in states:
        if state == "yes":
            typed.append("yes")
        elif state == "no":
            typed.append("no")
        elif state == "unknown":
            typed.append("unknown")
        else:
            typed.append("not_assessable")
    k = typed.count("yes")
    m = sum(s in {"unknown", "not_assessable"} for s in typed)
    thresholds: dict[str, PaperState] = {
        "strict": fixed_triple(typed, 3),
        "majority": fixed_triple(typed, 2),
        "permissive": fixed_triple(typed, 1),
    }
    return {
        "paper_id": slots[0]["paper_id"],
        "slot_states": typed,
        "clean_successes": k,
        "missing_or_invalid_slots": m,
        "complete_triple": m == 0,
        "primary_endpoint_majority": thresholds["majority"],
        **thresholds,
    }


def adjudicate() -> dict[str, object]:
    sealed = mapping(json.loads(BLIND.read_bytes()))
    if sealed.get("adjudication") != "not_performed":
        raise ValueError("blind outcome is not in the sealed, unadjudicated state")
    runs = [mapping(r) for r in listed(sealed["runs"])]
    by_paper: dict[str, list[dict[str, object]]] = {}
    for run in runs:
        pmid = str(run["paper_id"]).removeprefix("PMID:")
        spec = mapping(mapping(json.loads((SPECS / f"{pmid}.json").read_bytes()))["specification"])
        freeze = mapping(json.loads((FREEZES / f"{pmid}.json").read_bytes()))
        by_paper.setdefault(str(run["paper_id"]), []).append(adjudicate_slot(run, spec, freeze))
    papers = [adjudicate_paper(slots) for _, slots in sorted(by_paper.items())]
    for paper in papers:
        if len(by_paper[str(paper["paper_id"])]) != 3:
            raise ValueError(f"{paper['paper_id']}: expected exactly three sealed slots")
    counts = {
        "papers_with_completed_triples": len(papers),
        "runs": sum(len(s) for s in by_paper.values()),
        "run_success": {
            state: sum(str(s["run_success"]) == state for ss in by_paper.values() for s in ss)
            for state in ("yes", "no", "unknown", "not_assessable")
        },
        "primary_endpoint_majority": {
            state: sum(str(p["majority"]) == state for p in papers)
            for state in ("success", "failure", "unknown")
        },
    }
    return {
        "scope": "main_study_primary_adjudication_of_sealed_blind_outcomes",
        "amendment": AMENDMENT,
        "study": sealed["study"],
        "protocol_amendment": sealed["protocol_amendment"],
        "execution_amendment": sealed["amendment"],
        "blind_outcome_sha256": sha256(BLIND.read_bytes()),
        "adjudicated_at_utc": utc(),
        "adjudicator": "delegated_primary_mechanical (frozen rules applied by code)",
        "human_adjudication": "pending",
        "investigator_verification": "pending",
        "original_implementation_consulted": False,
        "provenance_rule": (
            "L3/L4 observed values count only when located, to reported precision, in a "
            "journaled execution output of code that reads served inputs (PD-04)."
        ),
        "counts": counts,
        "papers": papers,
        "runs": [s for _, ss in sorted(by_paper.items()) for s in ss],
    }


def write_results(record: dict[str, object]) -> None:
    runs = [mapping(r) for r in listed(record["runs"])]
    papers = [mapping(p) for p in listed(record["papers"])]
    write_csv(
        RESULTS / "run_outcomes.csv",
        [
            {
                "paper_id": r["paper_id"],
                "run_id": f"{str(r['paper_id']).removeprefix('PMID:')}-slot{r['slot']}",
                "L1_specification": r["L1_specification_extracted"],
                "L2_implementation": r["L2_implementation_created"],
                "L3_execution": r["L3_execution_successful"],
                "L4_numerical": r["L4_numerical_target_reproduced"],
                "L5_conclusion": r["L5_conclusion_preserved"],
                "failure_code": ";".join(str(c) for c in listed(r["failure_codes"])),
                "clean": str(r["clean"]).lower(),
                "evidence_id": f"main-run-20260925/runs/{str(r['paper_id']).replace(':', '_')}"
                f"/slot-{r['slot']}",
            }
            for r in runs
        ],
        [
            "paper_id",
            "run_id",
            "L1_specification",
            "L2_implementation",
            "L3_execution",
            "L4_numerical",
            "L5_conclusion",
            "failure_code",
            "clean",
            "evidence_id",
        ],
    )
    write_csv(
        RESULTS / "paper_outcomes.csv",
        [
            {
                "paper_id": p["paper_id"],
                "clean_successes": p["clean_successes"],
                "missing_or_invalid_slots": p["missing_or_invalid_slots"],
                "strict": p["strict"],
                "majority": p["majority"],
                "permissive": p["permissive"],
                "complete_triple": str(p["complete_triple"]).lower(),
                "evidence_id": "main-study-20260925/primary_adjudication.json",
            }
            for p in papers
        ],
        [
            "paper_id",
            "clean_successes",
            "missing_or_invalid_slots",
            "strict",
            "majority",
            "permissive",
            "complete_triple",
            "evidence_id",
        ],
    )
    failures = []
    for r in runs:
        codes = [str(c) for c in listed(r["failure_codes"])]
        if r["run_success"] == "yes" or not codes:
            continue
        levels = mapping(r["scored_levels"])
        earliest = next((lv for lv in ("L2", "L3", "L4") if levels[lv] != "yes"), "none")
        failures.append(
            {
                "paper_id": r["paper_id"],
                "run_id": f"{str(r['paper_id']).removeprefix('PMID:')}-slot{r['slot']}",
                "primary_code": codes[0],
                "secondary_codes": ";".join(codes[1:]),
                "earliest_level": earliest,
                "observation": (
                    f"stop_reason={r['stop_reason']}; "
                    f"served_input_executions={r['served_input_executions']}; "
                    f"L3={r['L3_execution_successful']}"
                ),
                "attribution": "delegated_primary_mechanical; human adjudication pending",
                "evidence_id": f"main-run-20260925/runs/{str(r['paper_id']).replace(':', '_')}"
                f"/slot-{r['slot']}",
            }
        )
    write_csv(
        RESULTS / "failures.csv",
        failures,
        [
            "paper_id",
            "run_id",
            "primary_code",
            "secondary_codes",
            "earliest_level",
            "observation",
            "attribution",
            "evidence_id",
        ],
    )


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--dry-run", action="store_true")
    args = parser.parse_args()
    record = adjudicate()
    if args.dry_run:
        print(json.dumps(record["counts"], indent=2))
        return
    payload = (json.dumps(record, indent=2, sort_keys=True, default=str) + "\n").encode()
    snapshot(ADJUDICATION, payload)
    stamped = stamp(ADJUDICATION, RUNS / "adjudication-timestamp")
    write_results(record)
    print(
        json.dumps(
            {
                "status": "primary_adjudication_recorded_human_pending",
                "primary_adjudication_sha256": stamped["source_sha256"],
                "timestamp_status": stamped["status"],
                **mapping(record["counts"]),
            },
            indent=2,
        )
    )


if __name__ == "__main__":
    main()
