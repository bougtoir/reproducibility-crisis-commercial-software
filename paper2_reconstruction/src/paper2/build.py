import argparse
import csv
import json
import subprocess
from collections import Counter
from datetime import datetime, timezone
from pathlib import Path

from paper2.core import FIELDS, Row, paper_registry, read_csv, sha256, snapshot, wilson, write_csv

ROOT = Path(__file__).resolve().parents[2]
REPO = ROOT.parent
REVISION = "2414f8d65cf2fe33c744c0a62cc05b6c77868367"
SOURCE_DIR = "reproducibility-crisis-commercial-software"
PINS = {
    "output/extracted_data.csv": "08068a5662968c8ffd1b70c7d61b3a6a4ed0cc42dda02f40ca0d113be3b750d0",
    "output/sampled_pmids.csv": "b920521604ad567207e6ecd9319876f7a8cb3f8f5fecd8a77b3a07116c8275f2",
    "output/summary_stats.json": "53717bd581f8913ae2192dc3717e48eeb2732e12b3712d4ce94e73b98b751c1b",
}
SOURCE_FILES = (
    *PINS,
    "output/sampling_stats.csv",
    "output/sampling.log",
    "pubmed_sampler.py",
    "run_extraction.py",
    "README.md",
)
EMPTY_TABLES = {
    "sampling_manifest.csv": [
        "paper_id",
        "sampling_stratum",
        "stratum_population",
        "stratum_sample",
        "inclusion_probability",
        "weight",
        "seed",
        "selected_at_utc",
        "protocol_sha256",
    ],
    "target_manifest.csv": [
        "paper_id",
        "target_id",
        "source_id",
        "location",
        "reported_value",
        "metric",
        "rationale",
        "tolerance_rule",
        "reported_precision",
        "claim",
        "selected_at_utc",
    ],
    "agent_run_manifest.csv": [
        "paper_id",
        "run_id",
        "slot",
        "model",
        "model_version",
        "prompt_sha256",
        "package_sha256",
        "environment_digest",
        "start_utc",
        "stop_utc",
        "firewall_status",
        "context_sha256",
        "tool_registry_sha256",
        "budget_id",
        "access_log_sha256",
    ],
    "run_outcomes.csv": [
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
    "paper_outcomes.csv": [
        "paper_id",
        "clean_successes",
        "missing_or_invalid_slots",
        "strict",
        "majority",
        "permissive",
        "complete_triple",
        "evidence_id",
    ],
    "failures.csv": [
        "paper_id",
        "run_id",
        "primary_code",
        "secondary_codes",
        "earliest_level",
        "observation",
        "attribution",
        "attribution_certainty",
        "evidence_id",
    ],
    "burden.csv": [
        "run_id",
        "wall_seconds",
        "input_tokens",
        "output_tokens",
        "attempts",
        "revisions",
        "dependency_count",
        "assumption_count",
        "ambiguity_count",
        "cost_usd",
        "evidence_id",
    ],
    "reveal.csv": [
        "paper_id",
        "item",
        "description_status",
        "publication_evidence",
        "independent_choice",
        "original_choice",
        "reveal_utc",
        "blind_freeze_id",
    ],
    "blind_freeze_manifest.csv": [
        "artifact",
        "bytes",
        "sha256",
        "frozen_at_utc",
        "independent_timestamp_receipt",
    ],
    "human_results.csv": [
        "case_id",
        "validator_id",
        "consent_record",
        "package_sha256",
        "start_utc",
        "stop_utc",
        "implementation",
        "execution",
        "numerical",
        "conclusion",
        "evidence_id",
    ],
}


def git_bytes(relative: str) -> bytes:
    return subprocess.check_output(
        ["git", "show", f"{REVISION}:{SOURCE_DIR}/{relative}"],
        cwd=REPO,
    )


def acquire_pinned_source() -> list[Row]:
    ledger_path = ROOT / "data/acquisition_ledger.csv"
    previous = read_csv(ledger_path) if ledger_path.exists() else []
    by_path = {row["path"]: row for row in previous}
    rows = []
    for relative in SOURCE_FILES:
        data = git_bytes(relative)
        digest = sha256(data)
        if relative in PINS and digest != PINS[relative]:
            raise ValueError(f"Pinned source hash mismatch: {relative}")
        target = Path("data/raw/epj") / relative
        snapshot(ROOT / target, data)
        record = {
            "source_id": relative,
            "version": REVISION,
            "url": f"https://github.com/bougtoir/{SOURCE_DIR}/blob/"
            f"{REVISION}/{SOURCE_DIR}/{relative}",
            "retrieved_at_utc": datetime.now(timezone.utc).isoformat(),
            "request_conditions": "git show pinned commit; full blob; no resampling",
            "path": str(target),
            "bytes": str(len(data)),
            "sha256": digest,
            "rights": "Existing author deposit; no new third-party redistribution licence inferred",
            "completeness": "complete_deposited_file; original_API_responses_not_deposited",
        }
        if str(target) in by_path:
            old = by_path[str(target)]
            if any(old[key] != record[key] for key in ("version", "bytes", "sha256")):
                raise ValueError("Existing acquisition ledger conflicts with pinned source")
            record["retrieved_at_utc"] = old["retrieved_at_utc"]
        rows.append(record)
    write_csv(ledger_path, rows, list(rows[0]))
    return rows


def prepare_experiment_tables(derived: Path) -> None:
    for filename, columns in EMPTY_TABLES.items():
        path = derived / filename
        if not path.exists():
            continue
        with path.open(newline="") as handle:
            header = next(csv.reader(handle), [])
        if header != columns:
            raise ValueError(f"Unexpected experimental table schema: {path}")
        if read_csv(path):
            raise ValueError(
                f"Preparation build must not overwrite observed experiment data: {path}"
            )
    for filename, columns in EMPTY_TABLES.items():
        path = derived / filename
        write_csv(path, [], columns)


def build() -> dict[str, object]:
    prepare_experiment_tables(ROOT / "data/derived")
    ledger = acquire_pinned_source()
    rows = read_csv(ROOT / "data/raw/epj/output/extracted_data.csv")
    sampled = read_csv(ROOT / "data/raw/epj/output/sampled_pmids.csv")
    if [(r["pmid"], r["stratum"]) for r in rows] != [(r["pmid"], r["stratum"]) for r in sampled]:
        raise ValueError("Ordered PMID/stratum source sequences disagree")
    registry, bridge = paper_registry(rows)
    for record in bridge:
        record["source_commit"] = REVISION
        record["source_file"] = f"{SOURCE_DIR}/output/extracted_data.csv"
    derived, results = ROOT / "data/derived", ROOT / "results"
    derived.mkdir(parents=True, exist_ok=True)
    results.mkdir(parents=True, exist_ok=True)
    write_csv(derived / "paper_registry.csv", registry, list(registry[0]))
    write_csv(derived / "source_record_bridge.csv", bridge, list(bridge[0]))
    duplicate_rows = [r for r in registry if int(r["record_multiplicity"]) > 1]
    write_csv(derived / "duplicate_pmids.csv", duplicate_rows, list(registry[0]))
    funnel = [
        {
            "source_record_id": row["source_record_id"],
            "paper_id": row["paper_id"],
            "epj_field": row["epj_field"],
            "G1_computationally_testable": "uncertain",
            "G1_reason": "not_assessed",
            "G2_principal_target_identifiable": "uncertain",
            "G3_input_accessibility": "unknown",
            "G4_resource_accessibility": "unknown",
            "G5_specification_sufficient_for_attempt": "uncertain",
            "G6_implementation_created": "not_assessed",
            "G7_execution_successful": "not_assessed",
            "G8_numerical_target_reproduced": "not_assessed",
            "G9_corresponding_conclusion_preserved": "not_assessed",
            "assessment_status": "NOT_STARTED",
            "evidence_id": "not_assessable",
        }
        for row in bridge
    ]
    write_csv(derived / "funnel_NOT_ASSESSED.csv", funnel, list(funnel[0]))
    characteristics = []
    for field in FIELDS:
        records = [r for r in rows if r["stratum"] == field]
        characteristics.append(
            {
                "field": field,
                "source_rows": str(len(records)),
                "unique_pmids_within_field": str(len({r["pmid"] for r in records})),
                "commercial_mentions": str(
                    sum(r["has_commercial_software"] == "True" for r in records)
                ),
                "open_source_mentions": str(
                    sum(r["has_opensource_software"] == "True" for r in records)
                ),
                "code_statement_detected": str(sum(r["code_available"] == "True" for r in records)),
                "data_statement_detected": str(sum(r["data_available"] == "True" for r in records)),
            }
        )
    write_csv(results / "corpus_characteristics.csv", characteristics, list(characteristics[0]))
    counts = Counter(r["pmid"] for r in rows)
    pairs = {(r["pmid"], r["stratum"]) for r in rows}
    membership_count = Counter(pmid for pmid, _ in pairs)
    metrics = {
        "source_records": len(rows),
        "unique_pmids": len(registry),
        "excess_records": len(rows) - len(registry),
        "duplicated_pmids": len(duplicate_rows),
        "double_occurrence_pmids": sum(value == 2 for value in counts.values()),
        "triple_occurrence_pmids": sum(value == 3 for value in counts.values()),
        "unique_pmid_field_pairs": len(pairs),
        "within_field_excess_records": len(rows) - len(pairs),
        "multi_field_pmids": sum(value > 1 for value in membership_count.values()),
        "epj_fields": len(characteristics),
        "missing_doi_records": sum(not r["doi"] for r in rows),
    }
    (results / "corpus_audit.json").write_text(json.dumps(metrics, indent=2) + "\n")
    precision = []
    planned_n = 100
    for numerator in (20, 30, 50, 70, 80):
        lower, upper = wilson(numerator, planned_n)
        precision.append(
            {
                "planned_n": str(planned_n),
                "assumed_p": str(numerator / planned_n),
                "wilson_lower": f"{lower:.9f}",
                "wilson_upper": f"{upper:.9f}",
                "total_width": f"{upper - lower:.9f}",
                "half_width": f"{(upper - lower) / 2:.9f}",
                "interpretation": "analytic_planning_not_empirical; unweighted_binomial_reference",
            }
        )
    write_csv(results / "precision_planning.csv", precision, list(precision[0]))
    values = [
        {
            "claim_id": key,
            "value": str(value),
            "unit": "count",
            "status": "observed_source_audit",
            "source_id": "output/extracted_data.csv",
            "source_sha256": PINS["output/extracted_data.csv"],
            "analysis": "src/paper2/build.py:build",
            "derived_file": "results/corpus_audit.json",
            "claim_scope": "deposited_corpus_not_reconstruction_outcome",
        }
        for key, value in metrics.items()
    ]
    write_csv(results / "manuscript_values.csv", values, list(values[0]))
    status: dict[str, object] = {
        "study_status": "PREPARATION_ONLY_NOT_SUBMISSION_READY",
        "corpus_identity": "VOR-linked deposited corpus located and commit pinned",
        "distinct_10000_paper_precondition": "not_met",
        "paper_frame_decision": "awaiting_author_confirmation",
        "historical_extraction_validation": "not_verified_no_annotations_recovered",
        "historical_raw_API_snapshots": "not_recovered",
        "pilot": "NOT_COMPLETED",
        "protocol_freeze": "NOT_COMPLETED",
        "funnel_assessment": "NOT_STARTED",
        "intensive_sample": "NOT_SELECTED",
        "reconstruction_experiments": "NOT_COMPLETED",
        "blind_freeze": "NOT_COMPLETED",
        "descriptive_reveal": "NOT_COMPLETED",
        "human_validation": "NOT_COMPLETED",
        "firewall_harness": "proposed_not_implemented_or_validated",
        "primary_success_rate": "not_assessable",
        "sources_verified": len(ledger),
    }
    (results / "readiness.json").write_text(json.dumps(status, indent=2) + "\n")
    return status


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--require-submission-ready", action="store_true")
    args = parser.parse_args()
    status = build()
    print(json.dumps(status, indent=2))
    if args.require_submission_ready:
        raise SystemExit(
            "Submission blocked: empirical study and required validations are incomplete."
        )


if __name__ == "__main__":
    main()
