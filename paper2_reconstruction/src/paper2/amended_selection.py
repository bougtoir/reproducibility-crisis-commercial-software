"""Render the amended (AMEND-2026-09-24-03) candidate selection from Devin's delegated record.

The record holds G1-G5 assessments of every Stage B route-eligible candidate plus a
custodian deposit inspection. This module re-validates each quote against the
retained article text, joins the Stage B route row, applies the main-sample rule
and writes a results table. Nothing here is a reconstruction outcome: dispositions
are sampling-frame decisions and remain pending author verification.
"""

from __future__ import annotations

import argparse
import json
from pathlib import Path

from paper2.build import ROOT
from paper2.core import sha256, write_csv
from paper2.model_api import mapping
from paper2.primary_adjudication import validate_assessment

AMENDMENT = "AMEND-2026-09-24-03"
MAIN = "main_sample_candidate"
CONDITIONAL_PREFIX = "main_sample_candidate_conditional_on_"
EXCLUDED_PREFIX = "excluded_"
COLUMNS = (
    "paper_id",
    "sampling_stratum",
    "candidate_rank",
    "G1_computationally_testable",
    "G2_principal_target_identifiable",
    "G3_input_accessibility",
    "G4_resource_accessibility",
    "G5_specification_sufficient_for_attempt",
    "route_kinds",
    "route_exclusion_class",
    "candidate_data_bytes",
    "implementation_files_withheld",
    "deposit_role",
    "outcome_leakage",
    "amended_disposition",
    "article_source_sha256",
    "evidence_segment_ids",
    "assessment_status",
)


def route_rows(routes: list[Path]) -> tuple[dict[str, dict[str, object]], set[str]]:
    """Join a Stage B record with its continuations.

    Later records override earlier rows (a continuation re-records a candidate whose
    deposit inspection excluded it); the returned set holds every paper that was
    route-eligible in any record, i.e. every paper that required a G1-G5 assessment.
    """
    by_paper: dict[str, dict[str, object]] = {}
    ever_eligible: set[str] = set()
    previous = b""
    for path in routes:
        data = path.read_bytes()
        payload = mapping(json.loads(data))
        if payload["amendment_id"] != AMENDMENT:
            raise ValueError(f"{path.name}: route record belongs to a different amendment")
        if previous and payload.get("continues_from_sha256") != sha256(previous):
            raise ValueError(
                f"{path.name}: continuation does not hash-chain to the previous record"
            )
        rows = payload["rows"]
        if not isinstance(rows, list):
            raise ValueError(f"{path.name}: route record has no rows")
        for r in rows:
            by_paper[str(mapping(r)["paper_id"])] = mapping(r)
            if mapping(r)["exclusion_class"] == "":
                ever_eligible.add(str(mapping(r)["paper_id"]))
        previous = data
    return by_paper, ever_eligible


def disposition_is_consistent(row: dict[str, object], route: dict[str, object]) -> None:
    paper_id = str(row["paper_id"])
    disposition = str(row["amended_disposition"])
    g1, g2 = str(row["G1_computationally_testable"]), str(row["G2_principal_target_identifiable"])
    leakage = str(row["outcome_leakage"])
    excluded = disposition.startswith(EXCLUDED_PREFIX)
    if route["exclusion_class"] not in ("", disposition.removeprefix(EXCLUDED_PREFIX)):
        raise ValueError(f"{paper_id}: route exclusion class disagrees with the disposition")
    if route["exclusion_class"] != "" and not excluded:
        raise ValueError(f"{paper_id}: route-excluded candidate cannot be a main-sample candidate")
    if leakage.startswith("target_disclosed") and not excluded:
        raise ValueError(f"{paper_id}: disclosed target requires an excluded disposition")
    if (g1 != "yes" or g2 == "no") and not excluded:
        raise ValueError(f"{paper_id}: main-sample candidate requires G1=yes and G2!=no")
    if excluded and row["pilot_decision"] != "not_pilot_case_replaced":
        raise ValueError(f"{paper_id}: excluded candidate must carry not_pilot_case_replaced")
    if not excluded and disposition != MAIN and not disposition.startswith(CONDITIONAL_PREFIX):
        raise ValueError(f"{paper_id}: unknown disposition {disposition!r}")


def render(record_path: Path, screen: Path, routes: list[Path], results: Path) -> dict[str, object]:
    record = mapping(json.loads(record_path.read_bytes()))
    if record["amendment_id"] != AMENDMENT or record["author_verification"] != "pending":
        raise ValueError(
            "record must reference the frozen amendment and keep author verification pending"
        )
    assessments = record["assessments"]
    if not isinstance(assessments, list):
        raise ValueError("assessments must be a list")
    by_paper, eligible = route_rows(routes)
    rows: list[dict[str, object]] = []
    for item in assessments:
        row = mapping(item)
        paper_id = str(row["paper_id"])
        route = by_paper[paper_id]
        if int(str(row["candidate_rank"])) != int(str(route["candidate_rank"])):
            raise ValueError(f"{paper_id}: candidate rank differs from the route record")
        disposition_is_consistent(row, route)
        validated = validate_assessment(screen, row)
        route_list = route["routes"]
        if not isinstance(route_list, list):
            raise ValueError(f"{paper_id}: route list missing")
        rows.append(
            {
                **validated,
                "route_kinds": ";".join(sorted({str(mapping(r)["kind"]) for r in route_list})),
                "route_exclusion_class": route["exclusion_class"],
                "candidate_data_bytes": sum(
                    int(str(mapping(r)["listed_bytes"]))
                    for r in route_list
                    if mapping(r)["exclusion_class"] == ""
                ),
                "implementation_files_withheld": sum(
                    int(str(mapping(r).get("implementation_files_withheld", 0))) for r in route_list
                ),
                "deposit_role": row["deposit_role"],
                "outcome_leakage": row["outcome_leakage"],
                "amended_disposition": row["amended_disposition"],
            }
        )
    assessed = {str(r["paper_id"]) for r in rows}
    if assessed != eligible:
        raise ValueError(
            f"assessed set differs from route-eligible set: {sorted(assessed ^ eligible)}"
        )
    strata = [
        str(r["sampling_stratum"])
        for r in rows
        if not str(r["amended_disposition"]).startswith(EXCLUDED_PREFIX)
    ]
    if len(strata) != len(set(strata)):
        raise ValueError("more than one main-sample candidate in a stratum")
    write_csv(
        results / "amended_candidate_selection.csv",
        [{column: r[column] for column in COLUMNS} for r in rows],
        COLUMNS,
    )
    summary: dict[str, object] = {
        "scope": str(record["scope"]),
        "amendment_id": AMENDMENT,
        "record_sha256": sha256(record_path.read_bytes()),
        "route_records": [{"path": p.name, "sha256": sha256(p.read_bytes())} for p in routes],
        "author_verification": "pending",
        "route_eligible_assessed": len(rows),
        "main_sample_candidates": {
            str(r["sampling_stratum"]): {
                "paper_id": r["paper_id"],
                "candidate_rank": r["candidate_rank"],
                "disposition": r["amended_disposition"],
                "target_candidate": r["target_candidate"],
            }
            for r in sorted(rows, key=lambda r: str(r["sampling_stratum"]))
            if not str(r["amended_disposition"]).startswith(EXCLUDED_PREFIX)
        },
        "excluded_after_deposit_inspection": [
            {
                "paper_id": r["paper_id"],
                "sampling_stratum": r["sampling_stratum"],
                "disposition": r["amended_disposition"],
                "outcome_leakage": r["outcome_leakage"],
            }
            for r in rows
            if str(r["amended_disposition"]).startswith(EXCLUDED_PREFIX)
        ],
        "strata_without_main_sample_candidate": sorted(
            {str(r["sampling_stratum"]) for r in rows} - set(strata)
        ),
        "reconstruction_started": False,
        "rates": "not applicable: no candidate has reached a reconstruction attempt",
    }
    (results / "amended_candidate_selection.json").write_text(json.dumps(summary, indent=2) + "\n")
    return summary


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--record",
        type=Path,
        default=ROOT / "data" / "adjudication" / "devin_amended_G1_G5_20260924.json",
    )
    parser.add_argument(
        "--screen", type=Path, default=ROOT / "data" / "raw" / "deposit-screen-20260924" / "screen"
    )
    parser.add_argument(
        "--routes",
        type=Path,
        nargs="+",
        default=[
            ROOT / "data" / "raw" / "deposit-screen-20260924" / "stage_b_routes.json",
            ROOT
            / "data"
            / "raw"
            / "deposit-screen-20260924"
            / "stage_b_routes_continuation_chemistry.json",
        ],
    )
    parser.add_argument("--results", type=Path, default=ROOT / "results")
    args = parser.parse_args()
    summary = render(args.record, args.screen, args.routes, args.results)
    print(json.dumps(summary, indent=2))


if __name__ == "__main__":
    main()
