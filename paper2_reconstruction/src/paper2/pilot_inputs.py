"""Lawful public-input acquisition for the provisional pilot papers.

Each deposit named in the retained article text is resolved through its public
service API, the listing response is retained verbatim, and only analysis-ready
files below a fixed byte cap are downloaded. Deposits that require an approved
application, that are available on request only, or that exceed the cap are
recorded as such: they are access observations, not G3 verdicts, and no
substitute input is invented. Nothing here reveals an original implementation,
so acquisition stays outside the solver package until the broker serves it.
"""

from __future__ import annotations

import argparse
import json
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from xml.etree import ElementTree

from paper2.acquisition import acquire, acquire_stream
from paper2.build import ROOT
from paper2.core import sha256, snapshot
from paper2.model_api import mapping

FILE_CAP = 64 * 1024 * 1024
TABULAR = (".csv", ".tsv", ".txt", ".xlsx", ".geno", ".snp", ".ind", ".zip", ".gz", ".json")


@dataclass(frozen=True)
class Deposit:
    paper_id: str
    accession: str
    kind: str
    listing_url: str
    note: str


DEPOSITS = (
    Deposit(
        "PMID:36493775",
        "10.5281/zenodo.6377228",
        "zenodo",
        "https://zenodo.org/api/records/6377228",
        "genotype dataset and raw figure tables named in the article data statement",
    ),
    Deposit(
        "PMID:36493775",
        "PRJEB51705",
        "ena_filereport",
        "https://www.ebi.ac.uk/ena/portal/api/filereport?accession=PRJEB51705"
        "&result=read_run&fields=run_accession,submitted_bytes,fastq_bytes&format=tsv",
        "raw BAM/read deposit; volume expected to exceed the per-file cap and run budget",
    ),
    Deposit(
        "PMID:36358920",
        "PXD027610",
        "proteomexchange",
        "https://proteomecentral.proteomexchange.org/cgi/GetDataset?ID=PXD027610"
        "&outputMode=XML",
        "DIA mass-spectrometry deposit named in the article",
    ),
    Deposit(
        "PMID:36409083",
        "PRJCA012518",
        "ngdc_bioproject",
        "https://ngdc.cncb.ac.cn/bioproject/browse/PRJCA012518",
        "genome deposit named in the article; service exposes no documented file API",
    ),
    Deposit(
        "PMID:39730532",
        "ADNI",
        "application_required",
        "",
        "ADNI requires an individually approved data-use application; not obtainable"
        " by free registration open to any reader",
    ),
    Deposit(
        "PMID:34031721",
        "author_request",
        "on_request_only",
        "",
        "video/tracking data and code available on reasonable request; no deposit",
    ),
    Deposit(
        "PMID:33584682",
        "article_supplementary_material",
        "supplement_only",
        "",
        "data statement points to the article and Supplementary Material only",
    ),
    Deposit(
        "PMID:39472468",
        "author_request",
        "on_request_only",
        "",
        "fluorescence training/test datasets available from the corresponding author"
        " on reasonable request only",
    ),
)


def zenodo_files(body: bytes) -> list[dict[str, object]]:
    record = mapping(json.loads(body))
    entries = record.get("files", [])
    files: list[dict[str, object]] = []
    if not isinstance(entries, list):
        return files
    for entry in entries:
        item = mapping(entry)
        links = mapping(item.get("links", {}))
        files.append(
            {
                "name": str(item.get("key", "")),
                "bytes": item.get("size"),
                "checksum": item.get("checksum"),
                "url": str(links.get("self", "")),
            }
        )
    return files


def proteomexchange_files(body: bytes) -> list[dict[str, object]]:
    """List the announced deposit files; ProteomeXchange announces no file sizes."""
    root = ElementTree.fromstring(body.decode())
    files: list[dict[str, object]] = []
    for item in root.iter("DatasetFile"):
        locations = [
            str(parameter.get("value", ""))
            for parameter in item.iter("cvParam")
            if str(parameter.get("name", "")) == "URI"
        ]
        files.append(
            {
                "name": str(item.get("name", "")),
                "bytes": None,
                "checksum": None,
                "url": locations[0] if locations else "",
            }
        )
    return files


def pride_files(body: bytes) -> list[dict[str, object]]:
    payload = json.loads(body)
    entries = payload if isinstance(payload, list) else mapping(payload).get("_embedded", [])
    files: list[dict[str, object]] = []
    if not isinstance(entries, list):
        return files
    for entry in entries:
        item = mapping(entry)
        locations = item.get("publicFileLocations", [])
        url = ""
        if isinstance(locations, list):
            for location in locations:
                candidate = mapping(location)
                if str(candidate.get("name", "")).startswith("FTP"):
                    url = str(candidate.get("value", ""))
        files.append(
            {
                "name": str(item.get("fileName", "")),
                "bytes": item.get("fileSizeBytes"),
                "checksum": item.get("checksum"),
                "url": url,
            }
        )
    return files


def serveable(entry: dict[str, object]) -> bool:
    size = entry.get("bytes")
    name = str(entry.get("name", "")).lower()
    return (
        isinstance(size, int)
        and 0 < size <= FILE_CAP
        and name.endswith(TABULAR)
        and str(entry.get("url", "")).startswith("http")
    )


def acquire_deposit(deposit: Deposit, destination: Path) -> dict[str, object]:
    record: dict[str, object] = {
        "paper_id": deposit.paper_id,
        "accession": deposit.accession,
        "kind": deposit.kind,
        "note": deposit.note,
        "file_cap_bytes": FILE_CAP,
    }
    if not deposit.listing_url:
        record["status"] = f"not_retrievable_by_public_route_{deposit.kind}"
        record["retained_files"] = []
        return record
    directory = destination / deposit.accession.replace("/", "_")
    listing_path, receipt = acquire(
        deposit.listing_url,
        deposit.accession,
        directory / "listing",
        request_conditions="GET public deposit listing; no implementation material requested",
    )
    record["listing"] = receipt
    if receipt["http_status"] != 200:
        record["status"] = "listing_request_failed_not_an_access_verdict"
        record["retained_files"] = []
        return record
    body = listing_path.read_bytes()
    try:
        if deposit.kind == "zenodo":
            files = zenodo_files(body)
        elif deposit.kind == "pride_files":
            files = pride_files(body)
        elif deposit.kind == "proteomexchange":
            files = proteomexchange_files(body)
        else:
            files = []
    except (ValueError, KeyError, TypeError):
        record["status"] = "listing_not_machine_readable"
        record["retained_files"] = []
        return record
    record["listed_files"] = len(files)
    record["listed_bytes"] = sum(int(f["bytes"]) for f in files if isinstance(f["bytes"], int))
    retained: list[dict[str, object]] = []
    excluded: list[dict[str, object]] = []
    for entry in files:
        if not serveable(entry):
            excluded.append(
                {
                    "name": entry["name"],
                    "bytes": entry["bytes"],
                    "reason": "exceeds_per_file_cap_or_not_analysis_ready_or_no_public_url",
                }
            )
            continue
        path, file_receipt = acquire(
            str(entry["url"]),
            f"{deposit.accession}:{entry['name']}",
            directory / "files",
            request_conditions="GET deposited analysis input below the frozen per-file cap",
        )
        retained.append(
            {
                "name": entry["name"],
                "declared_bytes": entry["bytes"],
                "declared_checksum": entry["checksum"],
                "retained_bytes": file_receipt["bytes"],
                "retained_sha256": file_receipt["sha256"],
                "http_status": file_receipt["http_status"],
                "path": str(path.resolve()),
            }
        )
    record["retained_files"] = retained
    record["excluded_files"] = excluded
    record["status"] = (
        "public_files_retained_below_cap"
        if retained
        else "listing_retained_no_file_below_cap_and_analysis_ready"
    )
    return record


def acquire_all(destination: Path) -> dict[str, object]:
    destination.mkdir(parents=True, exist_ok=True)
    records = [acquire_deposit(deposit, destination) for deposit in DEPOSITS]
    ledger: dict[str, object] = {
        "scope": (
            "lawful_public_input_acquisition_for_provisional_pilot_papers; "
            "access observations only, not G3 verdicts"
        ),
        "retrieved_at_utc": datetime.now(timezone.utc).isoformat(),
        "file_cap_bytes": FILE_CAP,
        "deposits": records,
        "papers_with_retained_public_input": sorted(
            {
                str(record["paper_id"])
                for record in records
                if record["status"] == "public_files_retained_below_cap"
            }
        ),
        "papers_without_retained_public_input": sorted(
            {str(record["paper_id"]) for record in records}
            - {
                str(record["paper_id"])
                for record in records
                if record["status"] == "public_files_retained_below_cap"
            }
        ),
    }
    content = (json.dumps(ledger, indent=2) + "\n").encode()
    digest = sha256(content)
    snapshot(destination / f"ledger-{digest[:16]}.json", content)
    ledger["ledger_sha256"] = digest
    return ledger


def fetch_ena_runs(report: Path, destination: Path, accession: str) -> dict[str, object]:
    """Stream every FASTQ listed in a retained ENA file report to persistent storage.

    The report is the retained anonymous ENA listing; each file is fetched over
    HTTPS from the listed path, verified against the listed md5 and byte count, and
    receipted. Nothing is unpacked or read.
    """
    lines = report.read_text().splitlines()
    header = lines[0].split("\t")
    rows = [dict(zip(header, line.split("\t"), strict=True)) for line in lines[1:]]
    files: list[dict[str, object]] = []
    for row in rows:
        paths = row["fastq_ftp"].split(";")
        md5s = row["fastq_md5"].split(";")
        sizes = row["fastq_bytes"].split(";")
        for path, digest, size in zip(paths, md5s, sizes, strict=True):
            _, receipt = acquire_stream(
                f"https://{path}",
                f"ena:{accession} {row['run_accession']} {path.rsplit('/', 1)[-1]}",
                destination / row["run_accession"],
                request_conditions=(
                    "GET anonymous ENA FASTQ mirror over HTTPS; path, md5 and bytes taken "
                    "from the retained file report; no login, payment or author contact"
                ),
                expected_md5=digest,
                expected_bytes=int(size),
            )
            files.append(
                {
                    "run_accession": row["run_accession"],
                    "file": path.rsplit("/", 1)[-1],
                    "bytes": receipt["bytes"],
                    "sha256": receipt["sha256"],
                    "completeness": receipt["completeness"],
                }
            )
            print(row["run_accession"], receipt["bytes"], receipt["completeness"], flush=True)
    ledger = {
        "accession": accession,
        "report_sha256": sha256(report.read_bytes()),
        "retrieved_at_utc": datetime.now(timezone.utc).isoformat(),
        "files": files,
        "complete_files": sum(1 for f in files if f["completeness"] == "complete_response"),
        "total_bytes": sum(int(str(f["bytes"])) for f in files),
    }
    snapshot(
        destination / f"{accession}_fetch_ledger.json",
        (json.dumps(ledger, indent=2) + "\n").encode(),
    )
    return ledger


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--output", type=Path, default=ROOT / "data" / "raw" / "pilot-inputs-20260924"
    )
    parser.add_argument("--ena-report", type=Path)
    parser.add_argument("--accession")
    args = parser.parse_args()
    if args.ena_report is not None:
        if args.accession is None:
            raise SystemExit("--accession is required with --ena-report")
        ena = fetch_ena_runs(args.ena_report, args.output, args.accession)
        print(json.dumps({k: ena[k] for k in ("accession", "complete_files", "total_bytes")}))
        return
    ledger = acquire_all(args.output)
    deposits = ledger["deposits"]
    summary = {
        "papers_with_retained_public_input": ledger["papers_with_retained_public_input"],
        "papers_without_retained_public_input": ledger["papers_without_retained_public_input"],
        "deposits": [
            {key: record[key] for key in ("paper_id", "accession", "status")}
            for record in (deposits if isinstance(deposits, list) else [])
        ],
    }
    print(json.dumps(summary, indent=2))


if __name__ == "__main__":
    main()
