"""Main-cohort selection under the frozen SAP (AMEND-2026-09-25-05 §6).

Builds the eligible reconstruction population from the exhaustive Stage A/B
deposit-route record, removes the prospective pilot, the accessibility layer and
every machine-readable exclusion class, then draws the field-stratified cohort by
the SAP's deterministic rule: one place per nonempty field, remaining places by
largest proportional-allocation deficit with lexicographic ties, within-field
order by SHA-256 of ``paper2-main-v1|paper_id``. Nothing about any outcome is
read here; the record is snapshotted and timestamped before any selected paper's
target freeze.
"""

from __future__ import annotations

import argparse
import json
from datetime import datetime, timezone
from pathlib import Path

from paper2.build import ROOT
from paper2.core import FIELDS, sha256, snapshot, wilson
from paper2.model_api import mapping
from paper2.timestamp import stamp

SELECTION_ID = "MAIN-COHORT-2026-09-25"
AMENDMENT = "AMEND-2026-09-25-05"
RANK_DOMAIN = "paper2-main-v1"
RECORD = ROOT / "data" / "adjudication" / "main_cohort_20260925.json"
CSV = ROOT / "results" / "main_cohort_20260925.csv"
STAMP_DIR = ROOT / "data" / "raw" / "main-cohort-20260925"
PILOT = ROOT / "data" / "adjudication" / "pilot_prospective_freeze_20260925.json"
ACCESS_LAYER = ROOT / "data" / "adjudication" / "accessibility_gate_failed_20260924.json"
PROTOCOL_FREEZE = ROOT / "data" / "adjudication" / "protocol_freeze_20260925.json"
TARGET_N = 100


def rank_key(paper_id: str) -> str:
    return sha256(f"{RANK_DOMAIN}|{paper_id}".encode())


def excluded_papers() -> dict[str, str]:
    out: dict[str, str] = {}
    pilot = mapping(json.loads(PILOT.read_bytes()))
    papers = pilot["papers"]
    if not isinstance(papers, list):
        raise ValueError("pilot freeze has no papers")
    for entry in papers:
        out[str(mapping(entry)["paper_id"])] = "prospective_pilot"
    layer = mapping(json.loads(ACCESS_LAYER.read_bytes()))
    members = layer["members"]
    if not isinstance(members, list):
        raise ValueError("accessibility layer has no members")
    for entry in members:
        out[str(mapping(entry)["paper_id"])] = "accessibility_gate_failed"
    return out


def eligible(routes: Path) -> tuple[list[dict[str, object]], dict[str, int], str]:
    payload = mapping(json.loads(routes.read_bytes()))
    if payload["stage"] != "B_routes_exhaustive":
        raise ValueError("main selection requires the exhaustive Stage B route record")
    rows = payload["rows"]
    if not isinstance(rows, list):
        raise ValueError("route record has no rows")
    removed = excluded_papers()
    frame: list[dict[str, object]] = []
    attrition: dict[str, int] = {}
    for raw in rows:
        row = mapping(raw)
        paper_id = str(row["paper_id"])
        reason = removed.get(paper_id)
        if reason is None and row["exclusion_class"] != "":
            reason = str(row["exclusion_class"])
        if reason is not None:
            attrition[reason] = attrition.get(reason, 0) + 1
            continue
        frame.append(
            {
                "paper_id": paper_id,
                "sampling_stratum": str(row["sampling_stratum"]),
                "candidate_rank": row["candidate_rank"],
                "statement_class": row["statement_class"],
                "route_status": row.get("route_status", ""),
                "rank_key": rank_key(paper_id),
            }
        )
    frame.sort(key=lambda r: (str(r["sampling_stratum"]), str(r["rank_key"])))
    return frame, dict(sorted(attrition.items())), sha256(routes.read_bytes())


def allocate(sizes: dict[str, int], target: int) -> dict[str, int]:
    """One place per nonempty field, then largest proportional-allocation deficit."""
    total = sum(sizes.values())
    planned = min(target, total)
    allocation = {field: (1 if sizes[field] else 0) for field in sizes}
    while sum(allocation.values()) < planned:
        best: tuple[float, str] | None = None
        for field in sorted(sizes):
            if allocation[field] >= sizes[field]:
                continue
            deficit = planned * sizes[field] / total - allocation[field]
            if best is None or deficit > best[0] + 1e-12:
                best = (deficit, field)
        if best is None:
            break
        allocation[best[1]] += 1
    return allocation


def select(routes: Path) -> dict[str, object]:
    freeze = mapping(json.loads(PROTOCOL_FREEZE.read_bytes()))
    frame, attrition, routes_sha = eligible(routes)
    assigned = {row["paper_id"]: row for row in frame}
    if len(assigned) != len(frame):
        raise ValueError("eligible frame contains duplicate paper ids")
    sizes = {field: sum(1 for r in frame if r["sampling_stratum"] == field) for field in FIELDS}
    allocation = allocate(sizes, TARGET_N)
    cohort: list[dict[str, object]] = []
    strata: list[dict[str, object]] = []
    for field in FIELDS:
        ordered = [r for r in frame if r["sampling_stratum"] == field]
        chosen = ordered[: allocation[field]]
        n_h, big_n_h = len(chosen), len(ordered)
        strata.append(
            {
                "sampling_stratum": field,
                "N_h": big_n_h,
                "n_h": n_h,
                "pi_h": (n_h / big_n_h) if big_n_h else 0.0,
                "weight": (big_n_h / n_h) if n_h else None,
                "ordered_frame_sha256": sha256(
                    "\n".join(str(r["paper_id"]) for r in ordered).encode()
                ),
            }
        )
        for position, row in enumerate(chosen, start=1):
            cohort.append({**row, "within_field_position": position})
    n = len(cohort)
    precision = [
        {
            "assumed_p": p,
            "wilson_lower": wilson(round(p * n), n)[0],
            "wilson_upper": wilson(round(p * n), n)[1],
        }
        for p in (0.2, 0.3, 0.5, 0.7, 0.8)
    ]
    return {
        "selection_id": SELECTION_ID,
        "amendment": AMENDMENT,
        "protocol_freeze_sha256": sha256(PROTOCOL_FREEZE.read_bytes()),
        "protocol_freeze_id": freeze["freeze_id"],
        "route_record": str(routes.relative_to(ROOT)),
        "route_record_sha256": routes_sha,
        "rank_domain": RANK_DOMAIN,
        "target_n": TARGET_N,
        "selected_n": n,
        "eligible_population": len(frame),
        "attrition": attrition,
        "strata": strata,
        "planned_precision_reference": precision,
        "precision_note": (
            "analytic unweighted binomial reference at the realised n, not a study result "
            "and not a sampling interval for the stratified design"
        ),
        "outcomes_inspected": "none",
        "pilot_in_primary_denominator": False,
        "investigator_verification": "pending",
        "cohort": cohort,
    }


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--routes",
        type=Path,
        default=ROOT / "data/raw/deposit-screen-20260924/stage_b_routes_full.json",
    )
    parser.add_argument("--stamp", action="store_true")
    args = parser.parse_args()
    if RECORD.exists():
        payload = RECORD.read_bytes()
        record = mapping(json.loads(payload))
        rebuilt = select(args.routes)
        for key in ("cohort", "strata", "attrition", "eligible_population"):
            if rebuilt[key] != record[key]:
                raise ValueError(f"main cohort no longer reproduces from its sources: {key}")
    else:
        record = {**select(args.routes), "selected_at_utc": datetime.now(timezone.utc).isoformat()}
        payload = (json.dumps(record, indent=2, ensure_ascii=False) + "\n").encode()
        snapshot(RECORD, payload)
        cohort = record["cohort"]
        if not isinstance(cohort, list):
            raise ValueError("cohort missing")
        lines = ["paper_id,sampling_stratum,within_field_position,rank_key"]
        lines += [
            f"{mapping(r)['paper_id']},{mapping(r)['sampling_stratum']},"
            f"{mapping(r)['within_field_position']},{mapping(r)['rank_key']}"
            for r in cohort
        ]
        snapshot(CSV, ("\n".join(lines) + "\n").encode())
    print(RECORD, sha256(payload))
    print(
        json.dumps(
            {
                "eligible_population": record["eligible_population"],
                "selected_n": record["selected_n"],
                "strata": [
                    {k: mapping(s)[k] for k in ("sampling_stratum", "N_h", "n_h")}
                    for s in (record["strata"] if isinstance(record["strata"], list) else [])
                ],
                "attrition": record["attrition"],
            },
            indent=2,
        )
    )
    if args.stamp:
        receipt = stamp(RECORD, STAMP_DIR)
        print(receipt["source_sha256"], receipt["status"])


if __name__ == "__main__":
    main()
