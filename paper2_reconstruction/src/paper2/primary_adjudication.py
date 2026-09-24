"""Validate and render the delegated primary G1–G5 adjudication record.

The record is authored by the delegated adjudicator and is never a formal
funnel update. Each evidence quote must occur verbatim inside the cited
segment of the hash-checked retained article text, so the judgements remain
auditable against the same bytes a later human verifier will see.
"""

from __future__ import annotations

import argparse
import json
from pathlib import Path

from paper2.build import ROOT
from paper2.core import sha256, write_csv
from paper2.model_api import mapping

GATES = ("G1", "G2", "G3", "G4", "G5")
ALLOWED = {
    "G1_computationally_testable": {"yes", "no", "uncertain"},
    "G2_principal_target_identifiable": {"yes", "no", "uncertain", "not_applicable"},
    "G3_input_accessibility": {
        "public",
        "obtainable_without_payment",
        "restricted_but_potentially_accessible",
        "unavailable",
        "insufficiently_identified",
        "not_applicable",
        "unknown",
    },
    "G4_resource_accessibility": {"available", "unavailable", "uncertain", "not_applicable"},
    "G5_specification_sufficient_for_attempt": {
        "sufficient",
        "insufficient",
        "uncertain",
        "not_applicable",
    },
    "pilot_decision": {"provisional_pilot_case", "not_pilot_case_replaced"},
}
COLUMNS = (
    "paper_id",
    "sampling_stratum",
    "candidate_rank",
    "article_text_status",
    "article_source_sha256",
    "G1_computationally_testable",
    "G1_reason",
    "G2_principal_target_identifiable",
    "G3_input_accessibility",
    "G4_resource_accessibility",
    "G5_specification_sufficient_for_attempt",
    "target_candidate",
    "pilot_decision",
    "evidence_segment_ids",
    "assessment_status",
)


def article_segments(screen: Path, paper_id: str) -> tuple[str, dict[str, str]]:
    pmid = paper_id.removeprefix("PMID:")
    payload = mapping(json.loads((screen / pmid / "sources.json").read_text()))
    sources = payload["sources"]
    if not isinstance(sources, list):
        raise ValueError(f"{paper_id}: sources.json has no source list")
    articles = [mapping(s) for s in sources if mapping(s)["role"] == "standalone_article_xml"]
    if len(articles) != 1:
        raise ValueError(f"{paper_id}: expected exactly one retained article source")
    article = articles[0]
    segments = article["segments"]
    if not isinstance(segments, list):
        raise ValueError(f"{paper_id}: article has no segments")
    return str(article["source_sha256"]), {
        str(mapping(s)["segment_id"]): str(mapping(s)["text"]) for s in segments
    }


def validate_assessment(screen: Path, row: dict[str, object]) -> dict[str, object]:
    paper_id = str(row["paper_id"])
    for key, allowed in ALLOWED.items():
        if str(row[key]) not in allowed:
            raise ValueError(f"{paper_id}: {key}={row[key]!r} not in {sorted(allowed)}")
    g1, g2 = row["G1_computationally_testable"], row["G2_principal_target_identifiable"]
    if g1 == "no" and row["pilot_decision"] != "not_pilot_case_replaced":
        raise ValueError(f"{paper_id}: G1=no cannot be a pilot case")
    if row["pilot_decision"] == "provisional_pilot_case" and (g1 != "yes" or g2 == "no"):
        raise ValueError(f"{paper_id}: pilot case requires G1=yes and G2!=no")
    digest, segments = article_segments(screen, paper_id)
    evidence = row["evidence"]
    if not isinstance(evidence, list) or not evidence:
        raise ValueError(f"{paper_id}: evidence list required")
    ids: list[str] = []
    for item in evidence:
        entry = mapping(item)
        segment_id = str(entry["segment_id"])
        quote = str(entry["quote"])
        if segment_id not in segments:
            raise ValueError(f"{paper_id}: unknown segment {segment_id}")
        if quote not in segments[segment_id]:
            raise ValueError(f"{paper_id}: quote not found verbatim in {segment_id}: {quote!r}")
        ids.append(segment_id)
    return {
        **{column: row.get(column, "") for column in COLUMNS if column in row},
        "article_source_sha256": digest,
        "evidence_segment_ids": ";".join(ids),
        "assessment_status": "DEVIN_PRIMARY_PENDING_AUTHOR_VERIFICATION",
    }


def validate_chain(record: dict[str, object], pilot: list[dict[str, object]]) -> dict[str, object]:
    """Check that each stratum consumes its deterministic order without gaps."""
    chain = mapping(mapping(record["deterministic_candidate_chain"])["strata"])
    selected = {str(r["sampling_stratum"]): int(str(r["candidate_rank"])) for r in pilot}
    counts: dict[str, int] = {}
    for stratum, entries in chain.items():
        if not isinstance(entries, list):
            raise ValueError(f"{stratum}: chain entries must be a list")
        ranks = [int(str(mapping(e)["candidate_rank"])) for e in entries]
        if ranks != list(range(1, len(ranks) + 1)):
            raise ValueError(f"{stratum}: chain skips deterministic ranks {ranks}")
        if selected.get(stratum) != ranks[-1]:
            raise ValueError(f"{stratum}: chain must end at the selected candidate")
        for entry in entries[:-1]:
            disposition = str(mapping(entry)["disposition"])
            if disposition.startswith("selected"):
                raise ValueError(f"{stratum}: an earlier candidate is also marked selected")
            counts[disposition] = counts.get(disposition, 0) + 1
    missing = sorted(set(selected) - set(chain))
    if missing:
        raise ValueError(f"strata without a recorded chain: {missing}")
    return {
        "strata": len(chain),
        "candidates_consumed": sum(
            len(entries) for entries in chain.values() if isinstance(entries, list)
        ),
        "skip_reason_counts": dict(sorted(counts.items())),
    }


def render(record_path: Path, screen: Path, results: Path) -> dict[str, object]:
    record = mapping(json.loads(record_path.read_bytes()))
    assessments = record["assessments"]
    if not isinstance(assessments, list):
        raise ValueError("assessments must be a list")
    rows = [validate_assessment(screen, mapping(a)) for a in assessments]
    pilot = [r for r in rows if r["pilot_decision"] == "provisional_pilot_case"]
    strata = [str(r["sampling_stratum"]) for r in pilot]
    if len(strata) != len(set(strata)):
        raise ValueError("more than one provisional pilot case in a stratum")
    replaced = [r for r in rows if r["pilot_decision"] == "not_pilot_case_replaced"]
    for r in replaced:
        later = [
            p
            for p in pilot
            if p["sampling_stratum"] == r["sampling_stratum"]
            and int(str(p["candidate_rank"])) > int(str(r["candidate_rank"]))
        ]
        if not later:
            raise ValueError(f"{r['paper_id']}: replaced without a later pilot case in stratum")
    chain = validate_chain(record, pilot)
    write_csv(results / "pilot_primary_adjudication.csv", rows, COLUMNS)
    summary: dict[str, object] = {
        "scope": str(record["scope"]),
        "record_sha256": sha256(record_path.read_bytes()),
        "deviation_id": mapping(record["deviation"])["deviation_id"],
        "author_verification": str(record["author_verification"]),
        "papers_assessed": len(rows),
        "provisional_pilot_cases": {
            str(r["sampling_stratum"]): {
                "paper_id": r["paper_id"],
                "candidate_rank": r["candidate_rank"],
                "target_candidate": r["target_candidate"],
                "G3_input_accessibility": r["G3_input_accessibility"],
            }
            for r in sorted(pilot, key=lambda r: str(r["sampling_stratum"]))
        },
        "replaced_candidates": [
            {
                "paper_id": r["paper_id"],
                "sampling_stratum": r["sampling_stratum"],
                "candidate_rank": r["candidate_rank"],
                "G1_computationally_testable": r["G1_computationally_testable"],
            }
            for r in replaced
        ],
        "gate_counts": {
            key: {
                value: sum(1 for r in rows if r[key] == value)
                for value in sorted({str(r[key]) for r in rows})
            }
            for key in (
                "G1_computationally_testable",
                "G2_principal_target_identifiable",
                "G3_input_accessibility",
                "G4_resource_accessibility",
                "G5_specification_sufficient_for_attempt",
            )
        },
        "deterministic_candidate_chain": chain,
        "formal_funnel_updated": False,
        "pilot_selection_frozen": False,
    }
    (results / "pilot_primary_adjudication.json").write_text(json.dumps(summary, indent=2) + "\n")
    return summary


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--record",
        type=Path,
        default=ROOT / "data" / "adjudication" / "devin_primary_G1_G5_20260923.json",
    )
    parser.add_argument(
        "--screen", type=Path, default=ROOT / "data" / "raw" / "pilot-screen-20260923"
    )
    parser.add_argument("--results", type=Path, default=ROOT / "results")
    args = parser.parse_args()
    summary = render(args.record, args.screen, args.results)
    print(json.dumps(summary["provisional_pilot_cases"], indent=2, ensure_ascii=False))


if __name__ == "__main__":
    main()
