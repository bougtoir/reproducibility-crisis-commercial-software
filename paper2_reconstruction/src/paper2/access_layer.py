"""Freeze the deposited-data amendment and build the accessibility layer.

The accessibility layer (`ACCESSIBILITY_GATE_FAILED`) retains papers that
passed the computational gates but whose study inputs no third party could
obtain lawfully and immediately. It is an audit sample of pre-reproducibility
access barriers, not a reconstruction sample: no member reached a
reconstruction attempt, so the record may not carry any success or failure
rate, and validation rejects one.
"""

from __future__ import annotations

import argparse
import json
from datetime import datetime, timezone
from pathlib import Path

from paper2.build import ROOT
from paper2.core import sha256, verification_status
from paper2.model_api import mapping
from paper2.timestamp import stamp

AMENDMENT = "AMEND-2026-09-24-03"
LAYER = "ACCESSIBILITY_GATE_FAILED"
CLASSES = {
    "no_lawful_text_retained",
    "no_deposit_named",
    "author_request_only",
    "application_required",
    "institutional_licence",
    "listing_unavailable",
    "size_unannounced",
    "above_cap",
    "deposited_output_only",
    "code_only_deposit",
    "supplement_only_unverified",
    "deposited_input_mismatch",
}
FORBIDDEN_KEY_TERMS = (
    "rate",
    "proportion",
    "success",
    "failure_count",
    "L1",
    "L2",
    "L3",
    "L4",
    "L5",
)

# Deposit status recorded by paper2.pilot_inputs -> amendment exclusion class.
STATUS_CLASS = {
    "not_retrievable_by_public_route_application_required": "application_required",
    "not_retrievable_by_public_route_on_request_only": "author_request_only",
    "not_retrievable_by_public_route_supplement_only": "supplement_only_unverified",
}
KIND_CLASS = {
    "ena_filereport": "above_cap",
    "proteomexchange": "size_unannounced",
    "ngdc_bioproject": "listing_unavailable",
    "zenodo": "deposited_output_only",
}


INSTRUCTION_BASIS = {
    AMENDMENT: (
        "Author instruction 2026-09-24: restrict the main analysis to papers with "
        "deposited analysis data that a third party can obtain lawfully and "
        "immediately; retain excluded papers with reasons; hold the seven "
        "provisional pilot papers in a separate accessibility layer without any "
        "reconstruction success or failure rate."
    ),
    "AMEND-2026-09-25-04": (
        "Investigator instruction 2026-09-25: the seven AMEND-2026-09-24-03 candidates "
        "are a prospective pilot outside the primary denominator; investigator "
        "verification replaces author verification; ten fields are frozen per paper "
        "before reconstruction; deposited_input_mismatch is an exclusion class; no "
        "Paper III analysis."
    ),
}
ESTIMAND = {
    AMENDMENT: (
        "reconstruction success among computational papers with immediately "
        "obtainable deposited analysis data"
    ),
    "AMEND-2026-09-25-04": (
        "unchanged from AMEND-2026-09-24-03; estimated only on the post-freeze main "
        "cohort, never on the prospective pilot"
    ),
}


def freeze_amendment(
    document: Path, destination: Path, record: Path, amendment_id: str = AMENDMENT
) -> dict[str, object]:
    if amendment_id not in INSTRUCTION_BASIS or amendment_id not in document.name:
        raise ValueError(f"unknown amendment or mismatched document for {amendment_id}")
    receipt = stamp(document, destination)
    payload = {
        "amendment_id": amendment_id,
        "scope": "protocol_amendment_hash_frozen_before_any_screening_under_it",
        "instruction_basis": INSTRUCTION_BASIS[amendment_id],
        "document": str(document.relative_to(ROOT)),
        "document_sha256": receipt["source_sha256"],
        "document_bytes": receipt["source_bytes"],
        "timestamp_response_sha256": receipt["response_sha256"],
        "timestamp_status": receipt["status"],
        "timestamp_receipt_dir": str(destination),
        "frozen_at_utc": receipt["retrieved_at_utc"],
        "applies_to": "every candidate screened after frozen_at_utc",
        "main_sample_estimand": ESTIMAND[amendment_id],
    }
    record.write_text(json.dumps(payload, indent=2, ensure_ascii=False) + "\n")
    return payload


def deposit_classes(ledger: dict[str, object], paper_id: str) -> list[dict[str, object]]:
    deposits = ledger["deposits"]
    if not isinstance(deposits, list):
        raise ValueError("ledger has no deposit list")
    rows = []
    for item in deposits:
        deposit = mapping(item)
        if deposit["paper_id"] != paper_id:
            continue
        status = str(deposit["status"])
        kind = str(deposit["kind"])
        if status in STATUS_CLASS:
            klass = STATUS_CLASS[status]
        elif status.startswith("listing_retained") and kind in KIND_CLASS:
            klass = KIND_CLASS[kind]
        elif status == "public_files_retained_below_cap" and kind in KIND_CLASS:
            klass = KIND_CLASS[kind]
        else:
            raise ValueError(f"{paper_id}: unmapped deposit status {status!r} ({kind})")
        listing = deposit.get("listing")
        rows.append(
            {
                "accession": deposit["accession"],
                "kind": kind,
                "pilot_input_status": status,
                "route_note": deposit["note"],
                "exclusion_class": klass,
                "listing_sha256": mapping(listing)["sha256"] if isinstance(listing, dict) else "",
                "listed_files": deposit.get("listed_files", 0),
                "listed_bytes": deposit.get("listed_bytes", 0),
            }
        )
    if not rows:
        raise ValueError(f"{paper_id}: no deposit route recorded")
    return rows


def build_layer(
    primary: Path, ledger_path: Path, amendment: Path, summary: Path, record: Path
) -> dict[str, object]:
    gates = mapping(json.loads(primary.read_bytes()))
    ledger = mapping(json.loads(ledger_path.read_bytes()))
    amend = mapping(json.loads(amendment.read_bytes()))
    pilot = mapping(json.loads(summary.read_bytes()))
    assessments = gates["assessments"]
    if not isinstance(assessments, list):
        raise ValueError("primary record has no assessments")
    members = []
    for item in assessments:
        row = mapping(item)
        if row["pilot_decision"] != "provisional_pilot_case":
            continue
        paper_id = str(row["paper_id"])
        routes = deposit_classes(ledger, paper_id)
        members.append(
            {
                "paper_id": paper_id,
                "sampling_stratum": row["sampling_stratum"],
                "candidate_rank": row["candidate_rank"],
                "layer": LAYER,
                "main_sample_membership": "excluded",
                "commercial_software_reproducibility_sample": "not_a_member",
                "primary_gates": {
                    key: row[key]
                    for key in (
                        "G1_computationally_testable",
                        "G2_principal_target_identifiable",
                        "G3_input_accessibility",
                        "G4_resource_accessibility",
                        "G5_specification_sufficient_for_attempt",
                    )
                },
                "target_candidate": row["target_candidate"],
                "evidence_segment_ids": [str(mapping(e)["segment_id"]) for e in row["evidence"]]
                if isinstance(row["evidence"], list)
                else [],
                "access_routes": routes,
                "exclusion_classes": sorted({str(r["exclusion_class"]) for r in routes}),
                "reconstruction_attempt_reached": False,
                "provisional_pilot_slots": pilot["slots_per_paper"][paper_id]
                if isinstance(pilot["slots_per_paper"], dict)
                else None,
                "note": (
                    "pilot slots ran without any deposited input; their sealed reports are "
                    "harness observations of input absence, not reconstruction outcomes"
                ),
            }
        )
    if len(members) != 7:
        raise ValueError(f"expected the seven provisional pilot papers, found {len(members)}")
    payload = {
        "record_id": "accessibility_gate_failed_20260924",
        "layer": LAYER,
        "scope": (
            "pre_reproducibility_access_barrier_audit_sample; excluded from the main "
            "reconstruction sample; retained for an auxiliary access-availability analysis "
            "or a separate paper"
        ),
        "amendment_id": amend["amendment_id"],
        "amendment_document_sha256": amend["document_sha256"],
        "primary_record_sha256": sha256(primary.read_bytes()),
        "input_ledger_sha256": sha256(ledger_path.read_bytes()),
        "pilot_summary_sha256": sha256(summary.read_bytes()),
        "pilot_deviation": pilot["deviation"],
        "gate_verification": {
            "investigator_verification": "pending",
            "devin_second_pass": (
                "each retained gate re-read against the hash-checked retained segments on "
                "2026-09-24 and left unchanged; primary G3 values are superseded for "
                "sampling purposes by the amendment exclusion classes"
            ),
        },
        "rates": "prohibited: no member reached a reconstruction attempt",
        "class_table_note": (
            "supplement_only_unverified extends the frozen amendment table: the only "
            "named input location is article supplementary material whose contents were "
            "not verified as the consumed analysis data (criterion 1 unmet, not a "
            "no_deposit_named verdict)"
        ),
        "recorded_utc": datetime.now(timezone.utc).isoformat(),
        "members": sorted(members, key=lambda m: str(m["sampling_stratum"])),
    }
    validate_layer(payload)
    record.write_text(json.dumps(payload, indent=2, ensure_ascii=False) + "\n")
    return payload


def _keys(value: object) -> list[str]:
    found: list[str] = []
    if isinstance(value, dict):
        for key, inner in value.items():
            found.append(str(key))
            found.extend(_keys(inner))
    elif isinstance(value, list):
        for inner in value:
            found.extend(_keys(inner))
    return found


def validate_layer(payload: dict[str, object]) -> dict[str, object]:
    members = payload["members"]
    if not isinstance(members, list) or not members:
        raise ValueError("layer needs members")
    for key in _keys(payload):
        if key == "rates":
            continue
        lowered = key.lower()
        if any(term.lower() in lowered for term in FORBIDDEN_KEY_TERMS):
            raise ValueError(f"layer record may not carry outcome field {key!r}")
    counts: dict[str, int] = {}
    for item in members:
        member = mapping(item)
        if member["layer"] != LAYER or member["main_sample_membership"] != "excluded":
            raise ValueError(f"{member['paper_id']}: not marked as excluded layer member")
        if member["reconstruction_attempt_reached"] is not False:
            raise ValueError(f"{member['paper_id']}: layer members never reach an attempt")
        classes = member["exclusion_classes"]
        if not isinstance(classes, list) or not classes:
            raise ValueError(f"{member['paper_id']}: exclusion class required")
        for klass in classes:
            if str(klass) not in CLASSES:
                raise ValueError(f"{member['paper_id']}: unknown class {klass!r}")
            counts[str(klass)] = counts.get(str(klass), 0) + 1
    return {
        "layer": LAYER,
        "amendment_id": payload["amendment_id"],
        "members": len(members),
        "exclusion_class_counts": dict(sorted(counts.items())),
        "rates": payload["rates"],
        "investigator_verification": verification_status(mapping(payload["gate_verification"])),
    }


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    sub = parser.add_subparsers(dest="command", required=True)
    freeze = sub.add_parser("freeze")
    freeze.add_argument(
        "--document",
        type=Path,
        default=ROOT / "protocols" / f"AMENDMENT_{AMENDMENT}_deposited_data_criterion.md",
    )
    freeze.add_argument("--destination", type=Path, required=True)
    freeze.add_argument(
        "--record",
        type=Path,
        default=ROOT / "data" / "adjudication" / f"amendment_{AMENDMENT}.json",
    )
    freeze.add_argument("--amendment-id", default=AMENDMENT)
    layer = sub.add_parser("layer")
    layer.add_argument(
        "--primary",
        type=Path,
        default=ROOT / "data" / "adjudication" / "devin_primary_G1_G5_20260923.json",
    )
    layer.add_argument("--ledger", type=Path, required=True)
    layer.add_argument(
        "--amendment",
        type=Path,
        default=ROOT / "data" / "adjudication" / f"amendment_{AMENDMENT}.json",
    )
    layer.add_argument(
        "--summary", type=Path, default=ROOT / "results" / "provisional_pilot_summary.json"
    )
    layer.add_argument(
        "--record",
        type=Path,
        default=ROOT / "data" / "adjudication" / "accessibility_gate_failed_20260924.json",
    )
    args = parser.parse_args()
    if args.command == "freeze":
        print(
            json.dumps(
                freeze_amendment(
                    args.document, args.destination, args.record, args.amendment_id
                ),
                indent=2,
            )
        )
    else:
        payload = build_layer(args.primary, args.ledger, args.amendment, args.summary, args.record)
        print(json.dumps(validate_layer(payload), indent=2))


if __name__ == "__main__":
    main()
