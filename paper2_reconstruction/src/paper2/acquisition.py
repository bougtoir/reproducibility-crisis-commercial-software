import argparse
import json
import time
from datetime import datetime, timezone
from pathlib import Path
from urllib.error import HTTPError, URLError
from urllib.parse import urlencode
from urllib.request import Request, urlopen

from paper2.build import ROOT
from paper2.core import FIELDS, Row, read_csv, sha256, snapshot, write_csv
from paper2.model_api import mapping

API = "https://www.ebi.ac.uk/europepmc/webservices/rest"


def acquire(
    url: str,
    identifier: str,
    directory: Path,
    *,
    request_conditions: str = "GET; exact identifier; no supplements or implementation fetched",
) -> tuple[Path, dict[str, object]]:
    target = directory / sha256(url.encode())
    receipt_path, body_path = target / "receipt.json", target / "body"
    if receipt_path.exists():
        receipt = mapping(json.loads(receipt_path.read_text()))
        if receipt["url"] != url or receipt["sha256"] != sha256(body_path.read_bytes()):
            raise ValueError("Stored acquisition differs from immutable receipt")
        return body_path, receipt
    request = Request(url, headers={"User-Agent": "PaperIIResearch/0.1 (public metadata audit)"})
    status, error, body, content_type, final_url = 0, "", b"", "", url
    try:
        with urlopen(request, timeout=90) as response:
            status = response.status
            body = response.read()
            content_type = response.headers.get("Content-Type", "")
            final_url = response.url
    except HTTPError as exception:
        status, error, body = exception.code, str(exception), exception.read()
    except (URLError, TimeoutError) as exception:
        error = type(exception).__name__
    snapshot(body_path, body)
    receipt = {
        "url": url,
        "final_url": final_url,
        "identifier": identifier,
        "version": "live_service_snapshot; see response for record version",
        "retrieved_at_utc": datetime.now(timezone.utc).isoformat(),
        "request_conditions": request_conditions,
        "http_status": status,
        "error": error,
        "content_type": content_type,
        "path": str(body_path.resolve()),
        "bytes": len(body),
        "sha256": sha256(body),
        "rights": "local evidence only; record/article-specific reuse rights require review",
        "completeness": "complete_response"
        if status == 200
        else "failed_request_not_access_verdict",
    }
    snapshot(receipt_path, (json.dumps(receipt, indent=2) + "\n").encode())
    return body_path, receipt


def candidate_order(frame: list[Row], per_field: int) -> list[Row]:
    if per_field < 1:
        raise ValueError("Candidate count must be positive")
    candidates = []
    for field in FIELDS:
        ordered = sorted(
            (row for row in frame if row["sampling_stratum"] == field),
            key=lambda row: (
                sha256(f"paper2-pilot-v1|{row['paper_id']}".encode()),
                row["paper_id"],
            ),
        )
        for rank, row in enumerate(ordered[:per_field], 1):
            candidates.append(
                {
                    "paper_id": row["paper_id"],
                    "pmid": row["pmid"],
                    "sampling_stratum": field,
                    "candidate_rank": str(rank),
                    "rank_sha256": sha256(f"paper2-pilot-v1|{row['paper_id']}".encode()),
                    "status": "acquisition_candidate_not_selected_or_assessed",
                }
            )
    return candidates


def acquire_candidates(destination: Path, per_field: int) -> list[Row]:
    frame_path = ROOT / "data/derived/inference_frame.csv"
    candidates = candidate_order(read_csv(frame_path), per_field)
    records = []
    for candidate in candidates:
        pmid = candidate["pmid"]
        parameters = urlencode(
            {
                "query": f"EXT_ID:{pmid} AND SRC:MED",
                "format": "json",
                "resultType": "core",
                "pageSize": "1",
            }
        )
        body_path, receipt = acquire(
            f"{API}/search?{parameters}", f"PMID:{pmid}", destination / "responses"
        )
        record = {
            **candidate,
            "metadata_sha256": str(receipt["sha256"]),
            "metadata_status": str(receipt["http_status"]),
            "pmcid": "",
            "open_access_flag": "unknown",
            "fulltext_status": "not_requested",
            "fulltext_sha256": "",
            "frame_sha256": sha256(frame_path.read_bytes()),
        }
        if receipt["http_status"] == 200:
            data = mapping(json.loads(body_path.read_bytes()))
            result_list = mapping(data["resultList"])["result"]
            if data["hitCount"] == 1 and isinstance(result_list, list) and len(result_list) == 1:
                article = mapping(result_list[0])
                if article["id"] != pmid or article["source"] != "MED":
                    raise ValueError("Europe PMC response identity mismatch")
                record["pmcid"] = str(article.get("pmcid", ""))
                record["open_access_flag"] = str(article.get("isOpenAccess", "unknown"))
                pmcid = record["pmcid"]
                if (
                    record["open_access_flag"] == "Y"
                    and pmcid.startswith("PMC")
                    and pmcid[3:].isdigit()
                ):
                    _, fulltext = acquire(
                        f"{API}/{pmcid}/fullTextXML", pmcid, destination / "responses"
                    )
                    record["fulltext_status"] = str(fulltext["http_status"])
                    record["fulltext_sha256"] = str(fulltext["sha256"])
            else:
                record["metadata_status"] = "unresolved_identity_or_result_count"
        records.append(record)
        write_csv(destination / "acquisition_index.csv", records, list(records[0]))
        print(
            f"{candidate['paper_id']} metadata={record['metadata_status']} "
            f"fulltext={record['fulltext_status']}",
            flush=True,
        )
        time.sleep(0.35)
    return records


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--candidates-per-field", type=int, default=1)
    args = parser.parse_args()
    acquire_candidates(args.output, args.candidates_per_field)


if __name__ == "__main__":
    main()
