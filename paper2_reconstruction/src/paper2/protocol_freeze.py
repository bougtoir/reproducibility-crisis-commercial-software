"""Final post-pilot protocol freeze (AMEND-2026-09-25-05).

Hashes the frozen protocol documents, the logging schema modules and the
final resource ceiling into one record, snapshots it immutably and timestamps
it. The main-study runner reads its ceiling from this module, so a ceiling
change is visible as a change to the frozen hash.
"""

from __future__ import annotations

import argparse
import json
from datetime import datetime, timezone

from paper2.build import ROOT
from paper2.core import sha256, snapshot
from paper2.model_api import mapping
from paper2.pilot_review import RECORD as REVIEW_RECORD
from paper2.prospective_freeze import CEILING as PILOT_CEILING
from paper2.timestamp import stamp

AMENDMENT = "AMEND-2026-09-25-05"
FREEZE_ID = "PROTOCOL-FREEZE-2026-09-25"
RECORD = ROOT / "data" / "adjudication" / "protocol_freeze_20260925.json"
STAMP_DIR = ROOT / "data" / "raw" / "protocol-freeze-20260925"
FINAL_CEILING: dict[str, object] = {**PILOT_CEILING, "provider_tokens_per_slot": 1_500_000}
FROZEN_DOCUMENTS = (
    "protocols/PROTOCOL.md",
    "protocols/SAP.md",
    "protocols/TOLERANCE_RULES.md",
    "protocols/FAILURE_TAXONOMY.md",
    "protocols/TARGET_SELECTION_RULES.md",
    "protocols/INFORMATION_FIREWALL.md",
    "protocols/REVEAL_PROTOCOL.md",
    "protocols/AMENDMENT_AMEND-2026-09-24-03_deposited_data_criterion.md",
    "protocols/AMENDMENT_AMEND-2026-09-25-04_pilot_role_and_prospective_freeze.md",
    "protocols/AMENDMENT_AMEND-2026-09-25-05_final_protocol_freeze.md",
)
LOGGING_SCHEMA_MODULES = (
    "src/paper2/controller.py",
    "src/paper2/isolation.py",
    "src/paper2/protocol_freeze.py",
)
SLOT_STATES = ("completed_report", "no_execution_report", "ceiling_stop", "harness_error")


def build() -> dict[str, object]:
    review = mapping(json.loads(REVIEW_RECORD.read_bytes()))
    return {
        "freeze_id": FREEZE_ID,
        "amendment": AMENDMENT,
        "basis_review_id": review["review_id"],
        "basis_review_sha256": sha256(REVIEW_RECORD.read_bytes()),
        "basis_blind_outcome_sha256": review["blind_outcome_sha256"],
        "documents": {
            path: {
                "sha256": sha256((ROOT / path).read_bytes()),
                "bytes": (ROOT / path).stat().st_size,
            }
            for path in FROZEN_DOCUMENTS
        },
        "logging_schema_modules": {
            path: sha256((ROOT / path).read_bytes()) for path in LOGGING_SCHEMA_MODULES
        },
        "final_ceiling": FINAL_CEILING,
        "pilot_ceiling": PILOT_CEILING,
        "ceiling_change": {
            "provider_tokens_per_slot": {
                "from": PILOT_CEILING["provider_tokens_per_slot"],
                "to": FINAL_CEILING["provider_tokens_per_slot"],
                "reason": "PD-01 internal consistency with the unchanged tool-call ceiling",
            }
        },
        "slot_states": SLOT_STATES,
        "failure_code_rule": "taxonomy identifiers only; other tokens retained verbatim as F99",
        "value_provenance_rule": "L3/L4 require the value in journaled execution output",
        "primary_endpoint": "at_least_two_valid_successes_among_three_fixed_slots",
        "pilot_in_primary_denominator": False,
        "paper_iii_analyses": "not_performed",
        "original_author_contact": "none",
        "investigator_verification": "pending",
    }


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--stamp", action="store_true")
    args = parser.parse_args()
    if RECORD.exists():
        payload = RECORD.read_bytes()
        record = mapping(json.loads(payload))
        rebuilt = build()
        for key in ("documents", "final_ceiling", "basis_review_sha256"):
            if rebuilt[key] != record[key]:
                raise ValueError(f"frozen protocol no longer matches its record: {key}")
    else:
        record = {**build(), "frozen_at_utc": datetime.now(timezone.utc).isoformat()}
        payload = (json.dumps(record, indent=2, ensure_ascii=False) + "\n").encode()
        snapshot(RECORD, payload)
    print(RECORD, sha256(payload))
    if args.stamp:
        receipt = stamp(RECORD, STAMP_DIR)
        print(receipt["source_sha256"], receipt.get("verification"))


if __name__ == "__main__":
    main()
