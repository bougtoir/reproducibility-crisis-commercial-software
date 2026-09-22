import argparse
import json
import time
import xml.etree.ElementTree as ET
from collections import Counter
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

from paper2.acquisition import API, acquire
from paper2.build import ROOT
from paper2.core import Row, read_csv, sha256, snapshot, write_csv
from paper2.corpus_acquisition import metadata_query, validate_batch
from paper2.model_api import mapping


def verify_metadata(source: Path) -> list[Row]:
    frame_path = ROOT / "data/derived/inference_frame.csv"
    frame = read_csv(frame_path)
    frame_hash = sha256(frame_path.read_bytes())
    index = read_csv(source / "metadata_index.csv")
    if [row["paper_id"] for row in index] != [row["paper_id"] for row in frame]:
        raise ValueError("Metadata acquisition index does not cover the exact ordered frame")
    verified: dict[str, tuple[str, dict[str, dict[str, object]], str]] = {}
    requested: dict[str, set[str]] = {}
    for row in index:
        if row["frame_sha256"] != frame_hash:
            raise ValueError("Metadata index refers to another inference frame")
        if row["paper_id"] != f"PMID:{row['pmid']}":
            raise ValueError("Metadata index has inconsistent PMID identity")
        relative = Path(row["source_snapshot"])
        if relative.is_absolute() or ".." in relative.parts:
            raise ValueError("Metadata snapshot path must stay within its acquisition directory")
        body_path = source / relative
        if str(relative) not in verified:
            body = body_path.read_bytes()
            receipt = mapping(json.loads(body_path.with_name("receipt.json").read_bytes()))
            digest = sha256(body)
            if receipt["sha256"] != digest or receipt["bytes"] != len(body):
                raise ValueError("Metadata response does not match its acquisition receipt")
            identifier = receipt["identifier"]
            if not isinstance(identifier, str) or not identifier.startswith("PMID_BATCH:"):
                raise ValueError("Unexpected metadata acquisition identifier")
            identifiers = identifier.removeprefix("PMID_BATCH:").split(",")
            if receipt["url"] != metadata_query(identifiers):
                raise ValueError("Metadata receipt identifier disagrees with the requested URL")
            requested[str(relative)] = set(identifiers)
            records = {}
            response_status = "request_failed"
            if receipt["http_status"] == 200:
                try:
                    records = validate_batch(body, identifiers)
                except ValueError:
                    response_status = "invalid_response"
                else:
                    response_status = "not_returned"
            verified[str(relative)] = digest, records, response_status
        digest, records, response_status = verified[str(relative)]
        if row["pmid"] not in requested[str(relative)]:
            raise ValueError("Metadata response was not requested for this PMID")
        if digest != row["source_snapshot_sha256"]:
            raise ValueError("Metadata index references another response hash")
        record = records.get(row["pmid"])
        expected_status = "identity_verified" if record is not None else response_status
        if row["status"] != expected_status:
            raise ValueError("Metadata index identity verdict conflicts with the retained response")
        if record is not None and (
            row["pmcid"] != str(record.get("pmcid", ""))
            or row["open_access_flag"] != str(record.get("isOpenAccess", "unknown"))
        ):
            raise ValueError("Metadata index attributes conflict with the retained response")
    return index


def verify_article(body: bytes, pmid: str, pmcid: str) -> dict[str, object]:
    root = ET.fromstring(body)
    if root.tag != "article":
        raise ValueError("Full-text response is not a standalone JATS article")
    identities = {
        element.get("pub-id-type"): "".join(element.itertext()).strip()
        for element in root.findall("./front/article-meta/article-id")
    }
    pmc_identity = identities.get("pmc") or identities.get("pmcid") or ""
    if identities.get("pmid") != pmid or pmc_identity.removeprefix("PMC") != pmcid[3:]:
        raise ValueError("Full-text response identity does not match the requested paper")
    body_element = root.find("body")
    if body_element is None or not "".join(body_element.itertext()).strip():
        raise ValueError("Full-text response has no article body")
    licenses = [
        " ".join(" ".join(element.itertext()).split())
        for element in root.findall("./front/article-meta/permissions/license")
    ]
    return {
        "identity": "verified_pmid_and_pmcid",
        "completeness": "nonempty_jats_body; supplement_and_input_coverage_not_assessed",
        "article_license_text": licenses,
        "reuse_decision": "local_evidence_only; redistribution_not_authorized_by_this_check",
    }


def acquire_article(row: Row, destination: Path) -> Row:
    result = {
        "paper_id": row["paper_id"],
        "pmid": row["pmid"],
        "pmcid": row["pmcid"],
        "frame_sha256": row["frame_sha256"],
        "metadata_sha256": row["source_snapshot_sha256"],
        "status": "pmc_endpoint_not_selected_other_sources_unassessed",
        "response_sha256": "",
        "response_path": "",
        "http_status": "",
        "validation_error": "",
    }
    pmcid = row["pmcid"]
    if (
        row["status"] != "identity_verified"
        or row["open_access_flag"] != "Y"
        or not pmcid.startswith("PMC")
        or not pmcid[3:].isdigit()
    ):
        return result
    body, receipt = acquire(
        f"{API}/{pmcid}/fullTextXML",
        pmcid,
        destination / "responses",
        request_conditions="GET standalone JATS article; no supplement, repository or code fetched",
    )
    result.update(
        response_sha256=str(receipt["sha256"]),
        response_path=str(body.relative_to(destination)),
        http_status=str(receipt["http_status"]),
        status="request_failed_not_access_verdict",
    )
    if receipt["http_status"] == 200:
        try:
            validation = verify_article(body.read_bytes(), row["pmid"], pmcid)
        except (ValueError, ET.ParseError) as error:
            result.update(
                status="unverified_article_response",
                validation_error=f"{type(error).__name__}: {error}",
            )
        else:
            snapshot(
                body.with_name("article_validation.json"),
                (json.dumps(validation, indent=2) + "\n").encode(),
            )
            result["status"] = "identity_verified_article"
    return result


def acquire_corpus_articles(source: Path, destination: Path) -> None:
    index = verify_metadata(source)
    snapshot(
        destination / "metadata_index_snapshot.csv",
        (source / "metadata_index.csv").read_bytes(),
    )
    rows = []
    failures = 0
    with ThreadPoolExecutor(max_workers=2) as executor:
        for start in range(0, len(index), 2):
            batch = list(
                executor.map(
                    lambda row: acquire_article(row, destination),
                    index[start : start + 2],
                )
            )
            rows.extend(batch)
            write_csv(destination / "article_index.csv", rows, list(rows[0]))
            for row in batch:
                if row["http_status"]:
                    failures = (
                        failures + 1 if row["http_status"] in {"0", "429", "502", "503"} else 0
                    )
            if failures >= 3:
                raise RuntimeError(
                    "Three failed article requests; partial index and receipts retained"
                )
            if start % 100 == 0:
                print(f"Article-source assessment: {len(rows)}/{len(index)}", flush=True)
            if any(row["http_status"] for row in batch):
                time.sleep(0.35)
    report = {
        "scope": "acquisition_only_not_G1_G5_classification",
        "frame_sha256": index[0]["frame_sha256"],
        "metadata_index_sha256": sha256((source / "metadata_index.csv").read_bytes()),
        "article_index_sha256": sha256((destination / "article_index.csv").read_bytes()),
        "papers_in_frame": len(index),
        "metadata_status_counts": dict(Counter(row["status"] for row in index)),
        "article_status_counts": dict(Counter(row["status"] for row in rows)),
        "input_access": "not_assessed",
        "pilot_completed": False,
    }
    snapshot(destination / "acquisition_audit.json", (json.dumps(report, indent=2) + "\n").encode())


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--metadata", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    acquire_corpus_articles(args.metadata, args.output)


if __name__ == "__main__":
    main()
