"""Prespecified procedural review of the sealed prospective pilot.

The review reads the sealed blind outcome and the per-slot controller journals
and derives procedural observations only: which ceiling stopped each slot, how
far each slot got, whether reports respected the logging schema, and whether
any report claims a value without an execution trail. It records defects and
the disposition each receives in the final protocol freeze. It computes no
success or failure rate and performs no adjudication of any observed value
against a publication; the pilot papers stay outside the primary denominator.
"""

from __future__ import annotations

import argparse
import json
import re
import statistics
from datetime import datetime, timezone
from pathlib import Path

from paper2.build import ROOT
from paper2.core import sha256, snapshot
from paper2.model_api import mapping
from paper2.prospective_freeze import CEILING, FREEZE_ID
from paper2.timestamp import stamp

REVIEW_ID = "PILOT-REVIEW-2026-09-25"
RECORD = ROOT / "data" / "adjudication" / "pilot_procedural_review_20260925.json"
STAMP_DIR = ROOT / "data" / "raw" / "pilot-review-20260925"
TAXONOMY = ROOT / "protocols" / "FAILURE_TAXONOMY.md"
CODE_PATTERN = re.compile(r"^F(0[1-9]|1[0-9]|20|99)$")


def taxonomy_codes() -> set[str]:
    codes = set(re.findall(r"^\| (F\d\d) \|", TAXONOMY.read_text(), flags=re.MULTILINE))
    if len(codes) < 20:
        raise ValueError("failure taxonomy could not be read")
    return codes


def slot_events(controller: Path) -> list[dict[str, object]]:
    events = []
    for path in sorted(controller.glob("event-*.json")):
        events.append(mapping(json.loads(path.read_text())))
    if not events:
        raise ValueError(f"no journal events under {controller}")
    return events


def slot_observation(run: dict[str, object], destination: Path) -> dict[str, object]:
    paper = str(run["paper_id"]).replace(":", "_")
    controller = destination / "runs" / paper / f"slot-{run['slot']}" / "controller"
    events = slot_events(controller)
    prompt_tokens = [
        int(str(mapping(e["value"])["prompt_tokens"]))
        for e in events
        if e["kind"] == "response" and "prompt_tokens" in mapping(e["value"])
    ]
    executions = [e for e in events if e["kind"] == "execution"]
    failed_executions = [
        e for e in executions if int(str(mapping(e["value"]).get("returncode", 0))) != 0
    ]
    report = run["claimed_report"]
    claimed = mapping(report) if isinstance(report, dict) else None
    codes: list[str] = []
    if claimed is not None:
        raw_codes = claimed.get("failure_codes", [])
        codes = [str(c) for c in raw_codes] if isinstance(raw_codes, list) else []
    artifacts = run["worker_artifacts"]
    observation = {
        "paper_id": run["paper_id"],
        "sampling_stratum": run["sampling_stratum"],
        "slot": run["slot"],
        "stop_reason": run["stop_reason"],
        "tool_calls": run["tool_calls"],
        "executions": len(executions),
        "failed_executions": len(failed_executions),
        "provider_accounted_tokens": run["provider_accounted_tokens"],
        "elapsed_seconds": run["elapsed_seconds"],
        "max_prompt_tokens": max(prompt_tokens, default=0),
        "report_submitted": claimed is not None,
        "observed_value_reported": claimed is not None
        and claimed.get("observed_value") is not None,
        "failure_codes": codes,
        "off_taxonomy_failure_codes": [c for c in codes if not CODE_PATTERN.match(c)],
        "worker_artifacts": len(artifacts) if isinstance(artifacts, list) else 0,
        "report_without_any_execution": claimed is not None and not executions,
        "value_without_exported_artifact": (
            claimed is not None
            and claimed.get("observed_value") is not None
            and not (isinstance(artifacts, list) and artifacts)
        ),
        "contamination_events": run["contamination_events"],
        "adjudication": run["adjudication"],
    }
    return observation


def quantile(values: list[float], q: float) -> float:
    if not values:
        return 0.0
    ordered = sorted(values)
    position = (len(ordered) - 1) * q
    low = int(position)
    high = min(low + 1, len(ordered) - 1)
    return ordered[low] + (ordered[high] - ordered[low]) * (position - low)


def review(destination: Path) -> dict[str, object]:
    outcome_path = destination / "blind_outcome.json"
    outcome = mapping(json.loads(outcome_path.read_text()))
    receipt = mapping(json.loads((destination / "outcome-timestamp" / "receipt.json").read_text()))
    if receipt["source_sha256"] != sha256(outcome_path.read_bytes()):
        raise ValueError("blind outcome does not match its timestamp receipt")
    if outcome["reveal"] != "not_performed" or outcome["adjudication"] != "not_performed":
        raise ValueError("procedural review requires an unrevealed, unadjudicated outcome")
    runs = outcome["runs"]
    if not isinstance(runs, list):
        raise ValueError("blind outcome has no runs")
    observations = [slot_observation(mapping(r), destination) for r in runs]
    taxonomy = taxonomy_codes()
    stop_counts: dict[str, int] = {}
    for o in observations:
        stop_counts[str(o["stop_reason"])] = stop_counts.get(str(o["stop_reason"]), 0) + 1
    wall = int(str(CEILING["wall_seconds_per_slot"]))
    tool_ceiling = int(str(CEILING["tool_calls_per_slot"]))
    token_ceiling = int(str(CEILING["provider_tokens_per_slot"]))
    elapsed = [float(str(o["elapsed_seconds"])) for o in observations]
    tool_calls = [int(str(o["tool_calls"])) for o in observations]
    prompts = [int(str(o["max_prompt_tokens"])) for o in observations]
    token_stopped = [o for o in observations if o["stop_reason"] == "token_reserve_limit"]
    off_taxonomy = sorted(
        {
            str(c)
            for o in observations
            for c in (
                o["off_taxonomy_failure_codes"]
                if isinstance(o["off_taxonomy_failure_codes"], list)
                else []
            )
        }
    )
    no_execution_reports = [o for o in observations if o["report_without_any_execution"]]
    unexported_values = [o for o in observations if o["value_without_exported_artifact"]]
    contamination = [o for o in observations if o["contamination_events"] != "none_recorded"]
    ceiling_analysis = {
        "frozen_ceiling": CEILING,
        "slots_stopped_by_token_reserve": len(token_stopped),
        "slots_total": len(observations),
        "max_elapsed_seconds": max(elapsed),
        "max_elapsed_fraction_of_wall_ceiling": max(elapsed) / wall,
        "max_tool_calls": max(tool_calls),
        "max_tool_calls_fraction_of_ceiling": max(tool_calls) / tool_ceiling,
        "median_tool_calls_at_token_stop": statistics.median(
            [int(str(o["tool_calls"])) for o in token_stopped]
        )
        if token_stopped
        else None,
        "p95_max_prompt_tokens": quantile([float(p) for p in prompts], 0.95),
        "cumulative_tokens_needed_for_tool_ceiling_at_p95_prompt": int(
            tool_ceiling * quantile([float(p) for p in prompts], 0.95)
        ),
        "binding_ceiling": "provider_tokens_per_slot"
        if len(token_stopped) > len(observations) / 2
        else "none_dominant",
        "token_ceiling_reached_before_any_heavy_computation": all(
            float(str(o["elapsed_seconds"])) < 600 for o in token_stopped
        ),
    }
    defects: list[dict[str, object]] = []
    if ceiling_analysis["binding_ceiling"] == "provider_tokens_per_slot":
        needed = int(
            str(ceiling_analysis["cumulative_tokens_needed_for_tool_ceiling_at_p95_prompt"])
        )
        defects.append(
            {
                "defect_id": "PD-01",
                "class": "resource_ceiling_internal_inconsistency",
                "observation": (
                    f"{len(token_stopped)} of {len(observations)} slots stopped at the cumulative "
                    f"provider-token ceiling ({token_ceiling}) after a median of "
                    f"{ceiling_analysis['median_tool_calls_at_token_stop']} tool calls and at most "
                    f"{max(elapsed):.0f} s of a {wall} s wall ceiling; the tool-call ceiling "
                    f"({tool_ceiling}) was never reachable because cumulative prompt tokens at the "
                    f"p95 context size require about {needed} tokens."
                ),
                "disposition": (
                    "final freeze sets the cumulative provider-token ceiling so that the "
                    "tool-call ceiling is attainable at the p95 observed context size; the wall, "
                    "step and tool-call ceilings are unchanged. This restores consistency between "
                    "ceilings and is not a tuning of target, tolerance or conclusion rules."
                ),
            }
        )
    if off_taxonomy:
        defects.append(
            {
                "defect_id": "PD-02",
                "class": "logging_schema_violation",
                "observation": f"free-text failure codes outside the taxonomy: {off_taxonomy}",
                "disposition": (
                    "final freeze requires failure_codes to be taxonomy identifiers; a report "
                    "carrying any other token is recorded as F99 with the raw token retained "
                    "verbatim, and adjudication treats it as unclassified, never as a new code."
                ),
            }
        )
    if no_execution_reports:
        defects.append(
            {
                "defect_id": "PD-03",
                "class": "report_without_execution",
                "observation": (
                    f"{len(no_execution_reports)} final report(s) submitted before any worker "
                    "execution while asserting that the harness had terminated the run; the "
                    "journal shows no such termination."
                ),
                "disposition": (
                    "final freeze classifies a final report with zero executions as "
                    "no_execution_report: the slot counts as attempted, L2 and L3 are recorded "
                    "as not_reached, any observed_value is void, and any procedural claim in the "
                    "report is checked against the journal, never accepted from the report."
                ),
            }
        )
    defects.append(
        {
            "defect_id": "PD-04",
            "class": "value_provenance_requirement",
            "observation": (
                f"{len(unexported_values)} report(s) with a non-null observed_value exported no "
                "worker artifact; the served publication text itself displays the reported "
                "value, so a transcribed value cannot be distinguished from a computed one by "
                "the report alone."
            ),
            "disposition": (
                "final freeze makes value provenance an adjudication requirement for L3 and L4: "
                "the observed value must appear in a journaled execution output produced by "
                "journaled code that reads served inputs; otherwise L3/L4 are not_established. "
                "Blinding remains blinding to the original implementation, not to the "
                "publication, and the manuscript states this explicitly."
            ),
        }
    )
    if contamination:
        defects.append(
            {
                "defect_id": "PD-05",
                "class": "firewall_contamination",
                "observation": f"{len(contamination)} slot(s) recorded contamination events",
                "disposition": "investigated before any main-study slot",
            }
        )
    return {
        "review_id": REVIEW_ID,
        "scope": "procedural_review_only_no_outcome_adjudication_no_rates",
        "freeze_id": FREEZE_ID,
        "reviewed_at_utc": datetime.now(timezone.utc).isoformat(),
        "blind_outcome_sha256": receipt["source_sha256"],
        "blind_outcome_timestamp_response_sha256": receipt["response_sha256"],
        "taxonomy_codes": sorted(taxonomy),
        "stop_reason_counts": dict(sorted(stop_counts.items())),
        "slots_with_report": sum(1 for o in observations if o["report_submitted"]),
        "slots_with_reported_value": sum(1 for o in observations if o["observed_value_reported"]),
        "ceiling_analysis": ceiling_analysis,
        "defects": defects,
        "rules_not_changed": [
            "target selection rules",
            "tolerance rules",
            "conclusion criteria",
            "wall, step and tool-call ceilings",
            "failure taxonomy codes",
            "three fixed slots per paper",
        ],
        "success_rate": "not_computed_pilot_is_procedural_only",
        "pilot_papers_in_primary_denominator": False,
        "investigator_verification": "pending",
        "human_adjudication": "pending",
        "observations": observations,
    }


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--destination", type=Path, required=True)
    parser.add_argument("--stamp", action="store_true")
    args = parser.parse_args()
    if RECORD.exists():
        payload = RECORD.read_bytes()
        record = mapping(json.loads(payload))
        rebuilt = review(args.destination)
        for key in ("observations", "defects", "ceiling_analysis", "stop_reason_counts"):
            if rebuilt[key] != record[key]:
                raise ValueError(
                    f"review record no longer reproduces from the sealed outcome: {key}"
                )
    else:
        record = review(args.destination)
        payload = (json.dumps(record, indent=2, ensure_ascii=False) + "\n").encode()
        snapshot(RECORD, payload)
    print(RECORD, sha256(payload))
    print(json.dumps(record["stop_reason_counts"]))
    for defect in record["defects"] if isinstance(record["defects"], list) else []:
        entry = mapping(defect)
        print(entry["defect_id"], entry["class"])
    if args.stamp:
        print(json.dumps(stamp(RECORD, STAMP_DIR), indent=2)[:400])


if __name__ == "__main__":
    main()
