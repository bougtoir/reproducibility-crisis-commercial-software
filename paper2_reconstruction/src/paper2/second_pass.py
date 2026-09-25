"""Render the delegated second-pass verification of the primary G1-G5 record.

The second pass re-reads every recorded gate against the same hash-checked
retained article segments, re-checks every primary evidence quote verbatim and
cites the retained data-availability statement behind each paper's access class
under amendment AMEND-2026-09-24-03. It is a delegated re-verification, not
investigator verification: `investigator_verification` stays pending until the
investigator team records gate decisions personally.
"""

from __future__ import annotations

import argparse
import json
from datetime import datetime, timezone
from pathlib import Path

from paper2.access_layer import AMENDMENT, CLASSES
from paper2.build import ROOT
from paper2.core import sha256, write_csv
from paper2.model_api import mapping
from paper2.primary_adjudication import GATES, article_segments

VERDICTS = {"confirmed", "corrected"}
GATE_KEYS = (
    "G1_computationally_testable",
    "G2_principal_target_identifiable",
    "G3_input_accessibility",
    "G4_resource_accessibility",
    "G5_specification_sufficient_for_attempt",
)
COLUMNS = (
    "paper_id",
    "sampling_stratum",
    "candidate_rank",
    "article_source_sha256",
    "primary_gates",
    "second_pass_verdict",
    "quotes_checked",
    "availability_segment_id",
    "access_class_under_amendment",
    "investigator_verification",
)


def verify(screen: Path, row: dict[str, object], note: dict[str, object]) -> dict[str, object]:
    paper_id = str(row["paper_id"])
    source_sha, segments = article_segments(screen, paper_id)
    evidence = row["evidence"]
    if not isinstance(evidence, list) or not evidence:
        raise ValueError(f"{paper_id}: primary record cites no evidence")
    checked = 0
    for item in evidence:
        quote = mapping(item)
        segment_id = str(quote["segment_id"])
        if segment_id not in segments:
            raise ValueError(f"{paper_id}: unknown segment {segment_id}")
        if str(quote["quote"]) not in segments[segment_id]:
            raise ValueError(f"{paper_id}: quote absent from {segment_id}")
        checked += 1
    availability = str(note["availability_segment_id"])
    if availability not in segments:
        raise ValueError(f"{paper_id}: unknown availability segment {availability}")
    if str(note["availability_quote"]) not in segments[availability]:
        raise ValueError(f"{paper_id}: availability quote absent from {availability}")
    verdict = str(note["verdict"])
    if verdict not in VERDICTS:
        raise ValueError(f"{paper_id}: verdict {verdict!r} not allowed")
    access_class = str(note["access_class"])
    if access_class not in CLASSES:
        raise ValueError(f"{paper_id}: access class {access_class!r} not in the frozen table")
    gates = {key: row[key] for key in GATE_KEYS}
    revised = note.get("revised_gates", {})
    if verdict == "confirmed" and revised:
        raise ValueError(f"{paper_id}: a confirmed gate set may not carry revisions")
    if verdict == "corrected" and not revised:
        raise ValueError(f"{paper_id}: a corrected gate set must state the revised gates")
    return {
        "paper_id": paper_id,
        "sampling_stratum": row["sampling_stratum"],
        "candidate_rank": row["candidate_rank"],
        "article_source_sha256": source_sha,
        "article_segments": len(segments),
        "primary_gates": gates,
        "second_pass_verdict": verdict,
        "revised_gates": revised,
        "quotes_rechecked": checked,
        "availability_segment_id": availability,
        "availability_quote": note["availability_quote"],
        "access_class_under_amendment": access_class,
        "note": note["note"],
        "investigator_verification": "pending",
    }


def render(
    primary: Path, notes: Path, screen: Path, record: Path, results: Path
) -> dict[str, object]:
    gates = mapping(json.loads(primary.read_bytes()))
    review = mapping(json.loads(notes.read_bytes()))
    assessments = gates["assessments"]
    entries = review["notes"]
    if not isinstance(assessments, list) or not isinstance(entries, list):
        raise ValueError("primary record and notes must both hold lists")
    by_paper = {str(mapping(n)["paper_id"]): mapping(n) for n in entries}
    if len(by_paper) != len(entries):
        raise ValueError("duplicate second-pass note")
    rows = []
    for item in assessments:
        row = mapping(item)
        paper_id = str(row["paper_id"])
        if paper_id not in by_paper:
            raise ValueError(f"{paper_id}: no second-pass note; every gate must be re-read")
        rows.append(verify(screen, row, by_paper[paper_id]))
    extra = set(by_paper) - {str(r["paper_id"]) for r in rows}
    if extra:
        raise ValueError(f"notes for papers absent from the primary record: {sorted(extra)}")
    payload = {
        "record_id": "devin_second_pass_G1_G5_20260924",
        "scope": (
            "delegated_second_pass_reverification_pending_investigator_verification_"
            "not_formal_human_adjudication"
        ),
        "reviewer": review["reviewer"],
        "gates": list(GATES),
        "amendment_id": AMENDMENT,
        "primary_record": str(primary.relative_to(ROOT)),
        "primary_record_sha256": sha256(primary.read_bytes()),
        "notes_sha256": sha256(notes.read_bytes()),
        "screen_directory": str(screen),
        "evidence_rule": (
            "Every primary evidence quote was re-checked verbatim inside the cited segment "
            "of the hash-checked retained article text, and each access class cites the "
            "retained data-availability statement verbatim."
        ),
        "investigator_verification": "pending",
        "investigator_verification_note": (
            "This record does not substitute for the investigator team's verification; gates stay "
            "provisional and the funnel stays unassessed until the investigator team "
            "records decisions."
        ),
        "recorded_utc": datetime.now(timezone.utc).isoformat(),
        "papers": len(rows),
        "verdict_counts": {
            verdict: sum(1 for row in rows if row["second_pass_verdict"] == verdict)
            for verdict in sorted(VERDICTS)
        },
        "quotes_rechecked": sum(int(str(row["quotes_rechecked"])) for row in rows),
        "assessments": rows,
    }
    record.write_text(json.dumps(payload, indent=2, ensure_ascii=False) + "\n")
    write_csv(
        results / "second_pass_G1_G5.csv",
        [
            {
                "paper_id": row["paper_id"],
                "sampling_stratum": row["sampling_stratum"],
                "candidate_rank": row["candidate_rank"],
                "article_source_sha256": row["article_source_sha256"],
                "primary_gates": "; ".join(
                    f"{key.split('_')[0]}={mapping(row['primary_gates'])[key]}" for key in GATE_KEYS
                ),
                "second_pass_verdict": row["second_pass_verdict"],
                "quotes_checked": row["quotes_rechecked"],
                "availability_segment_id": row["availability_segment_id"],
                "access_class_under_amendment": row["access_class_under_amendment"],
                "investigator_verification": row["investigator_verification"],
            }
            for row in rows
        ],
        COLUMNS,
    )
    summary = {key: payload[key] for key in ("record_id", "papers", "verdict_counts")}
    summary["quotes_rechecked"] = payload["quotes_rechecked"]
    summary["investigator_verification"] = payload["investigator_verification"]
    (results / "second_pass_G1_G5.json").write_text(
        json.dumps(summary, indent=2, ensure_ascii=False) + "\n"
    )
    return payload


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--primary",
        type=Path,
        default=ROOT / "data" / "adjudication" / "devin_primary_G1_G5_20260923.json",
    )
    parser.add_argument(
        "--notes",
        type=Path,
        default=ROOT / "data" / "adjudication" / "second_pass_notes_20260924.json",
    )
    parser.add_argument("--screen", type=Path, required=True)
    parser.add_argument(
        "--record",
        type=Path,
        default=ROOT / "data" / "adjudication" / "devin_second_pass_G1_G5_20260924.json",
    )
    parser.add_argument("--results", type=Path, default=ROOT / "results")
    args = parser.parse_args()
    payload = render(args.primary, args.notes, args.screen, args.record, args.results)
    print(
        json.dumps(
            {
                "record_id": payload["record_id"],
                "papers": payload["papers"],
                "verdict_counts": payload["verdict_counts"],
                "quotes_rechecked": payload["quotes_rechecked"],
                "investigator_verification": payload["investigator_verification"],
            },
            indent=2,
        )
    )


if __name__ == "__main__":
    main()
