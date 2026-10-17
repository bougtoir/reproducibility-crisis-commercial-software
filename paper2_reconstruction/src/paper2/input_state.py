"""Input-state identifiability: ACCESS-stage refinement, conceptual table and descriptives.

Nothing here rescoring any gate, slot or adjudication record. The descriptive summary
reads only frozen public records (delegated specifications, frozen target packages and
the retrieval ledger of the served inputs) and reports fields as recorded or not recorded.
"""

import argparse
import json
from collections import Counter
from pathlib import Path

from paper2.core import Row, read_csv, write_csv

ROOT = Path(__file__).resolve().parents[2]
STUDY_DIR = ROOT / "data" / "adjudication" / "main-study-20260925"
LEDGER = STUDY_DIR / "input_state_ledger.csv"
RESULTS = ROOT / "results"

ACCESS_QUESTIONS = (
    "resource accessible?",
    "exact input state identifiable?",
    "historical input state retrievable?",
)
DEFINITIONS = {
    "data availability": "whether a named resource can be accessed at all",
    "input-state identifiability": (
        "the extent to which the exact state of the input data used in an analysis can be "
        "uniquely identified from the publication and associated records"
    ),
    "historical-state retrievability": (
        "the extent to which that exact previously used data state can still be obtained "
        "by an independent third party"
    ),
}
KEY_STATEMENT = "Public availability is not equivalent to reproducible input identity."
MINIMUM_REPORTING = (
    "repository or accession",
    "release/version",
    "retrieval date/time",
    "exact query or API parameters",
    "file names/list",
    "schema/version",
    "cryptographic checksum",
    "transformation log",
    "license/redistribution status",
)
# (requirement, stage, purpose); ACCESS rows first, then the retained framework rows.
STAGE_TABLE: tuple[tuple[str, str, str], ...] = (
    ("repository/accession", "ACCESS", "identify data source"),
    ("release/version", "ACCESS", "identify exact dataset state"),
    ("retrieval date/time", "ACCESS", "anchor dynamic resource state in time"),
    ("exact query/API parameters", "ACCESS", "reconstruct query-generated input"),
    ("file list/schema", "ACCESS", "fix input composition/structure"),
    ("SHA-256/content hash", "ACCESS / REPRODUCE", "verify byte-level identity"),
    (
        "transformation log",
        "RECONSTRUCT / REPRODUCE",
        "trace derivation from raw to analytical input",
    ),
    (
        "license/redistribution status",
        "ACCESS",
        "explain why exact snapshot can or cannot be redistributed",
    ),
    (
        "code availability",
        "ACCESS (artifact access)",
        "locate the original implementation; artifact reproducibility only, deliberately "
        "excluded from the Paper II blind intervention and used solely as a descriptive "
        "characteristic and post-seal reveal resource",
    ),
    ("software/version reporting", "ACCESS / EXECUTE", "identify the executable environment"),
    ("container/environment", "EXECUTE", "reinstate the execution environment"),
    ("detailed Methods", "RECONSTRUCT", "specify the procedure well enough to implement"),
    ("parameters/defaults", "RECONSTRUCT", "fix analytical choices that change results"),
    ("random seed", "EXECUTE / REPRODUCE", "control stochastic variation"),
    ("resource requirements", "EXECUTE", "state compute, memory and time needed"),
    ("independent reconstruction", "RECONSTRUCT → REPRODUCE", "test the description itself"),
    ("robustness testing", "ROBUST", "vary reasonable choices (Paper III; excluded here)"),
)


def stage_table() -> list[Row]:
    return [
        {"requirement": requirement, "stage": stage, "purpose": purpose}
        for requirement, stage, purpose in STAGE_TABLE
    ]


def load(path: Path) -> dict[str, object]:
    value = json.loads(path.read_text())
    if not isinstance(value, dict):
        raise ValueError(f"{path}: expected a JSON object")
    return value


def dicts(value: object) -> list[dict[str, object]]:
    if not isinstance(value, list):
        return []
    return [item for item in value if isinstance(item, dict)]


def unspecified(version: str) -> bool:
    text = version.strip().lower()
    return text in {"", "none", "unknown", "n/a", "not reported"} or any(
        marker in text for marker in ("not specified", "unspecified", "not stated")
    )


def extract_ledger(private_root: Path) -> list[Row]:
    """Build the public retrieval ledger of the served inputs from private receipts."""
    receipts: dict[str, dict[str, object]] = {}
    for receipt_path in private_root.glob("*/*/*/receipt.json"):
        record = load(receipt_path)
        receipts[str(record["sha256"])] = record
    rows: list[Row] = []
    for freeze in sorted((STUDY_DIR / "freezes").glob("*.json")):
        package = load(freeze)
        for item in dicts(package["required_inputs"]):
            receipt = receipts[str(item["sha256"])]
            announced = str(receipt.get("announced_md5", ""))
            rows.append(
                {
                    "paper_id": str(package["paper_id"]),
                    "identifier": str(item["identifier"]),
                    "repository_kind": str(item["identifier"]).split(":")[0],
                    "file_name": str(item["name"]),
                    "url": str(receipt["url"]),
                    "retrieved_at_utc": str(receipt["retrieved_at_utc"]),
                    "repository_version_identifier": "not recorded (live service snapshot)",
                    "query_or_api_parameters": str(receipt["request_conditions"]),
                    "bytes": str(item["bytes"]),
                    "sha256": str(item["sha256"]),
                    "announced_checksum": "md5 announced" if announced else "none announced",
                    "announced_checksum_matches": (
                        "yes" if announced and announced == str(receipt.get("md5", "")) else ""
                    ),
                    "schema_or_release_identifier": "not recorded",
                    "license_redistribution_status": str(receipt["rights"]),
                    "historical_version_check": "not assessed",
                }
            )
    return rows


def input_state_stats() -> dict[str, int]:
    """Structured ACCESS-stage counts behind input_state_descriptives.csv.

    Single computation for both the descriptive table and the manuscript values
    layer: every number the prose cites must come from this dict.
    """
    status = load(STUDY_DIR / "paper_status.json")
    kinds: Counter[str] = Counter()
    accession_papers = listing_papers = 0
    reference_papers = reference_entries = version_stated = version_unspecified = 0
    papers_with_unspecified = 0
    for paper_id in status:
        record = load(STUDY_DIR / "specifications" / f"{paper_id.removeprefix('PMID:')}.json")
        listing = dicts(record.get("listing"))
        accession_papers += any(entry.get("accession") for entry in listing)
        listing_papers += any(entry.get("files") for entry in listing)
        for entry in listing:
            kinds[str(entry.get("kind"))] += 1
        specification = record.get("specification")
        if not isinstance(specification, dict):
            continue
        references = dicts(specification.get("required_public_references"))
        if not references:
            continue
        reference_papers += 1
        flags = [unspecified(str(reference.get("version", ""))) for reference in references]
        reference_entries += len(flags)
        version_unspecified += sum(flags)
        version_stated += len(flags) - sum(flags)
        papers_with_unspecified += any(flags)
    ledger = read_csv(LEDGER)
    external = sum(
        1
        for value in status.values()
        if isinstance(value, dict)
        and value.get("eligibility") == "external_reference_resource_required_not_served"
    )
    return {
        "status_papers": len(status),
        "accession_papers": accession_papers,
        "listing_papers": listing_papers,
        "repository_kinds": sum(kinds.values()),
        "reference_papers": reference_papers,
        "reference_entries": reference_entries,
        "version_stated_entries": version_stated,
        "version_unspecified_entries": version_unspecified,
        "papers_with_unspecified_version": papers_with_unspecified,
        "external_reference_barrier_papers": external,
        "served_files": len(ledger),
        "served_files_timestamped": sum(bool(r["retrieved_at_utc"]) for r in ledger),
        "served_files_announced_checksum": sum(
            r["announced_checksum"] == "md5 announced" for r in ledger
        ),
        "served_files_checksum_matched": sum(
            r["announced_checksum_matches"] == "yes" for r in ledger
        ),
    }


def descriptives() -> list[Row]:
    """Descriptive ACCESS-stage metadata as recorded in the frozen main-study records."""
    stats = input_state_stats()
    status_papers = stats["status_papers"]
    ledger_rows = stats["served_files"]
    kinds: Counter[str] = Counter()
    status = load(STUDY_DIR / "paper_status.json")
    for paper_id in status:
        record = load(STUDY_DIR / "specifications" / f"{paper_id.removeprefix('PMID:')}.json")
        for entry in dicts(record.get("listing")):
            kinds[str(entry.get("kind"))] += 1
    analysis = load(RESULTS / "main_analysis.json")
    denominators = analysis.get("denominators")
    if not isinstance(denominators, dict):
        raise ValueError("main_analysis.json denominators missing")
    attempted = int(str(denominators["attempted_papers_complete_triples"]))
    source = f"delegated specification listing ({status_papers} papers)"
    frozen = (
        "frozen target packages and retrieval ledger "
        f"({attempted} attempted papers)"
    )
    rows = [
        (
            "repository/accession recorded",
            f"{stats['accession_papers']}/{status_papers} papers",
            source,
        ),
        (
            "repository kinds in listings",
            "; ".join(f"{kind} {count}" for kind, count in sorted(kinds.items())),
            source,
        ),
        (
            "file list recorded in listing",
            f"{stats['listing_papers']}/{status_papers} papers",
            source,
        ),
        (
            "external public reference resources named by specifier",
            f"{stats['reference_entries']} entries in {stats['reference_papers']} papers",
            source,
        ),
        (
            "reference version/release recorded by the specifier from publication text",
            f"{stats['version_stated_entries']}/{stats['reference_entries']} entries",
            source,
        ),
        (
            "reference version/release not specified in publication text",
            f"{stats['version_unspecified_entries']}/{stats['reference_entries']} entries "
            f"({stats['papers_with_unspecified_version']} papers with at least one)",
            source,
        ),
        (
            "not run because an external public reference resource was required",
            f"{stats['external_reference_barrier_papers']}/{status_papers} papers "
            "(frozen barrier state, unchanged)",
            "paper_status.json",
        ),
        (
            "served inputs with SHA-256 recorded",
            f"{ledger_rows}/{ledger_rows} files",
            frozen,
        ),
        (
            "served inputs with retrieval timestamp recorded",
            f"{stats['served_files_timestamped']}/{ledger_rows} files",
            frozen,
        ),
        (
            "served inputs with repository-announced checksum",
            f"{stats['served_files_announced_checksum']}/{ledger_rows} "
            f"files; {stats['served_files_checksum_matched']} matched",
            frozen,
        ),
        (
            "repository version/release identifier of served inputs",
            "not recorded (live service snapshots)",
            frozen,
        ),
        ("query/API parameters", "anonymous GET of listed files; no query-generated input", frozen),
        ("schema/release identifier", "not recorded", frozen),
        ("historical version check", "not assessed", frozen),
    ]
    return [{"item": item, "value": value, "source": src} for item, value, src in rows]


def input_state_values() -> list[Row]:
    """manuscript_values.csv rows for the ACCESS-stage descriptives (frozen records)."""
    stats = input_state_stats()
    entries = (
        ("input_state_reference_entries", stats["reference_entries"]),
        ("input_state_reference_papers", stats["reference_papers"]),
        (
            "input_state_version_unspecified_entries",
            stats["version_unspecified_entries"],
        ),
        (
            "input_state_papers_with_unspecified_version",
            stats["papers_with_unspecified_version"],
        ),
        (
            "input_state_external_reference_barrier_papers",
            stats["external_reference_barrier_papers"],
        ),
        ("input_state_served_files", stats["served_files"]),
        (
            "input_state_served_files_checksum_matched",
            stats["served_files_checksum_matched"],
        ),
    )
    return [
        {
            "claim_id": key,
            "value": str(value),
            "unit": "count",
            "status": "frozen_record_descriptive_not_rescored",
            "source_id": "data/adjudication/main-study-20260925/",
            "source_sha256": "",
            "analysis": "src/paper2/input_state.py:input_state_stats",
            "derived_file": "results/input_state_descriptives.csv",
            "claim_scope": "access_stage_descriptive_not_empirical_drift_measurement",
        }
        for key, value in entries
    ]


def save(path: Path, rows: list[Row]) -> None:
    write_csv(path, rows, list(rows[0]))


def write_outputs() -> None:
    save(RESULTS / "requirement_stage_table.csv", stage_table())
    save(RESULTS / "input_state_descriptives.csv", descriptives())


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--extract-ledger", type=Path, help="private main-inputs directory")
    arguments = parser.parse_args()
    if arguments.extract_ledger is not None:
        save(LEDGER, extract_ledger(arguments.extract_ledger))
    write_outputs()


if __name__ == "__main__":
    main()
