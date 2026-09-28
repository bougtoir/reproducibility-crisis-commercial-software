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
    ("code availability", "ACCESS", "locate the original implementation (not used in Paper II)"),
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


def descriptives() -> list[Row]:
    """Descriptive ACCESS-stage metadata as recorded in the frozen main-study records."""
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
    source = "delegated specification listing (100 papers)"
    frozen = "frozen target packages and retrieval ledger (10 attempted papers)"
    rows = [
        ("repository/accession recorded", f"{accession_papers}/{len(status)} papers", source),
        (
            "repository kinds in listings",
            "; ".join(f"{kind} {count}" for kind, count in sorted(kinds.items())),
            source,
        ),
        ("file list recorded in listing", f"{listing_papers}/{len(status)} papers", source),
        (
            "external public reference resources named by specifier",
            f"{reference_entries} entries in {reference_papers} papers",
            source,
        ),
        (
            "reference version/release stated in publication text",
            f"{version_stated}/{reference_entries} entries",
            source,
        ),
        (
            "reference version/release not specified",
            f"{version_unspecified}/{reference_entries} entries "
            f"({papers_with_unspecified} papers with at least one)",
            source,
        ),
        (
            "not run because an external public reference resource was required",
            f"{external}/{len(status)} papers (frozen barrier state, unchanged)",
            "paper_status.json",
        ),
        ("served inputs with SHA-256 recorded", f"{len(ledger)}/{len(ledger)} files", frozen),
        (
            "served inputs with retrieval timestamp recorded",
            f"{sum(bool(r['retrieved_at_utc']) for r in ledger)}/{len(ledger)} files",
            frozen,
        ),
        (
            "served inputs with repository-announced checksum",
            f"{sum(r['announced_checksum'] == 'md5 announced' for r in ledger)}/{len(ledger)} "
            f"files; {sum(r['announced_checksum_matches'] == 'yes' for r in ledger)} matched",
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
