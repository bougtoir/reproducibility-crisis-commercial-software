import argparse
import json
import time
from pathlib import Path
from urllib.parse import urlencode

from paper2.acquisition import API, acquire
from paper2.build import ROOT
from paper2.core import Row, read_csv, sha256, write_csv
from paper2.model_api import mapping


def metadata_query(identifiers: list[str]) -> str:
    if not identifiers or len(identifiers) > 50 or any(not pmid.isdigit() for pmid in identifiers):
        raise ValueError("A metadata batch requires 1–50 numeric PMIDs")
    query = "(" + " OR ".join(f"EXT_ID:{pmid}" for pmid in identifiers) + ") AND SRC:MED"
    return f"{API}/search?" + urlencode(
        {
            "query": query,
            "format": "json",
            "resultType": "core",
            "pageSize": "100",
        }
    )


def validate_batch(body: bytes, identifiers: list[str]) -> dict[str, dict[str, object]]:
    data = mapping(json.loads(body))
    if "resultList" not in data or "hitCount" not in data:
        raise ValueError("Metadata service returned no result envelope")
    results = mapping(data["resultList"]).get("result")
    if (
        not isinstance(results, list)
        or type(data["hitCount"]) is not int
        or data["hitCount"] != len(results)
    ):
        raise ValueError("Metadata batch is truncated or malformed; no completeness claim allowed")
    records: dict[str, dict[str, object]] = {}
    for value in results:
        record = mapping(value)
        pmid = record.get("id")
        if record.get("source") != "MED" or not isinstance(pmid, str) or pmid not in identifiers:
            raise ValueError("Metadata record is outside the requested source/identity set")
        if pmid in records:
            raise ValueError("Duplicate PMID returned by metadata service")
        records[pmid] = record
    return records


def fetch_group(pmids: list[str], destination: Path) -> dict[str, Row]:
    body, receipt = acquire(
        metadata_query(pmids), "PMID_BATCH:" + ",".join(pmids), destination / "responses"
    )
    time.sleep(0.35)
    status = "request_failed"
    records: dict[str, dict[str, object]] = {}
    if receipt["http_status"] == 200:
        try:
            records = validate_batch(body.read_bytes(), pmids)
            status = "not_returned"
        except ValueError:
            if len(pmids) > 1:
                middle = len(pmids) // 2
                return fetch_group(pmids[:middle], destination) | fetch_group(
                    pmids[middle:], destination
                )
            status = "invalid_response"
    rows = {}
    for pmid in pmids:
        record = records.get(pmid)
        rows[pmid] = {
            "source_snapshot_sha256": str(receipt["sha256"]),
            "source_snapshot": str(body.relative_to(destination)),
            "status": "identity_verified" if record is not None else status,
            "pmcid": str(record.get("pmcid", "")) if record else "",
            "open_access_flag": str(record.get("isOpenAccess", "unknown")) if record else "unknown",
        }
    return rows


def acquire_metadata(destination: Path) -> list[Row]:
    frame_path = ROOT / "data/derived/inference_frame.csv"
    frame = read_csv(frame_path)
    frame_hash = sha256(frame_path.read_bytes())
    rows: list[Row] = []
    failure_streak = 0
    for start in range(0, len(frame), 10):
        batch = frame[start : start + 10]
        records = fetch_group([row["pmid"] for row in batch], destination)
        failure_streak = (
            failure_streak + 1
            if all(row["status"] == "request_failed" for row in records.values())
            else 0
        )
        rows.extend(
            {
                "paper_id": row["paper_id"],
                "pmid": row["pmid"],
                "frame_sha256": frame_hash,
                **records[row["pmid"]],
            }
            for row in batch
        )
        write_csv(destination / "metadata_index.csv", rows, list(rows[0]))
        print(f"Metadata snapshots: {len(rows)}/{len(frame)}", flush=True)
        if failure_streak >= 3:
            raise RuntimeError("Three consecutive failed batches; partial acquisition retained")
    return rows


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    acquire_metadata(args.output)


if __name__ == "__main__":
    main()
