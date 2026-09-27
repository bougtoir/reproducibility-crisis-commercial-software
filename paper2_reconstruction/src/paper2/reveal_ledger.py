"""Descriptive original-code reveal ledger for the sealed main study.

Reads hand-authored descriptive observations, verifies that every cited source is a
retained public acquisition (receipt SHA-256 equals the retained bytes), that every
quoted snippet occurs verbatim in the retained archive, and that every acquisition
postdates the blind seal. Original code is never executed. Nothing here reads or
modifies the blind outcome or the primary adjudication beyond checking hashes/times.
"""

from __future__ import annotations

import argparse
import io
import json
import tarfile
from datetime import datetime, timezone
from pathlib import Path

from paper2.core import sha256, sha256_file, snapshot, utc_time, write_csv
from paper2.timestamp import stamp

ROOT = Path(__file__).resolve().parents[2]
STUDY_DIR = ROOT / "data" / "adjudication" / "main-study-20260925"
BLIND = STUDY_DIR / "blind_outcome.json"
ADJUDICATION = STUDY_DIR / "primary_adjudication.json"
OBSERVATIONS = STUDY_DIR / "reveal" / "observations.json"
LEDGER = STUDY_DIR / "reveal" / "reveal_ledger.json"
EVIDENCE = Path("/home/ubuntu/paper2_evidence/main-reveal-20260925")
INPUTS = Path("/home/ubuntu/paper2_evidence/main-inputs-20260925")
WITHHELD = "original_implementation_artifact_withheld"
RESULTS = ROOT / "results"


def mapping(value: object) -> dict[str, object]:
    if not isinstance(value, dict):
        raise ValueError("expected a JSON object")
    return {str(k): v for k, v in value.items()}


def listed(value: object) -> list[object]:
    if not isinstance(value, list):
        raise ValueError("expected a JSON list")
    return value


def receipt_path(reference: str) -> Path:
    path = Path(reference)
    return path if path.is_absolute() else EVIDENCE / path


def withheld_from_solver(paper_id: str, name: str) -> bool:
    """True when the input-acquisition ledger records the deposit file as withheld original code."""
    path = INPUTS / paper_id.removeprefix("PMID:") / "acquisition.json"
    ledger = mapping(json.loads(path.read_bytes()))
    return any(
        mapping(e)["name"] == name and mapping(e)["reason"] == WITHHELD
        for e in listed(ledger["excluded"])
    )


def verify_source(
    paper_id: str, source: dict[str, object], sealed_at: datetime
) -> dict[str, object]:
    receipt_file = receipt_path(str(source["receipt"]))
    receipt = mapping(json.loads(receipt_file.read_bytes()))
    body = Path(str(receipt["path"]))
    digest = sha256_file(body)
    if digest != receipt["sha256"] or digest != source["sha256"]:
        raise ValueError(f"{source['identifier']}: retained bytes do not match receipt")
    if receipt.get("http_status") != 200 or receipt.get("completeness") != "complete_response":
        raise ValueError(f"{source['identifier']}: acquisition incomplete")
    retrieved = utc_time(str(receipt["retrieved_at_utc"]))
    before_seal = retrieved <= sealed_at
    if before_seal and source["kind"] != "withheld_deposit_file":
        raise ValueError(f"{source['identifier']}: acquired before the blind seal")
    if source["kind"] == "withheld_deposit_file":
        name = str(source["identifier"]).rsplit(":", 1)[-1]
        if not withheld_from_solver(paper_id, name):
            raise ValueError(f"{source['identifier']}: not recorded as withheld from the solver")
    return {
        "acquired_before_seal_withheld_from_solver": before_seal,
        "identifier": source["identifier"],
        "kind": source["kind"],
        "url": receipt["url"],
        "retrieved_at_utc": receipt["retrieved_at_utc"],
        "bytes": receipt["bytes"],
        "sha256": digest,
        "receipt": str(receipt_file),
        "body": str(body),
    }


def member_text(body: Path, kind: str, member: str) -> str:
    if kind != "github_tarball":
        return body.read_text(errors="replace")
    with tarfile.open(fileobj=io.BytesIO(body.read_bytes()), mode="r:gz") as archive:
        for info in archive.getmembers():
            if info.isfile() and info.name.split("/", 1)[-1] == member:
                extracted = archive.extractfile(info)
                if extracted is None:
                    break
                return extracted.read().decode(errors="replace")
    raise ValueError(f"{member}: not found in {body}")


def verify_snippets(
    snippets: list[dict[str, object]], sources: dict[str, dict[str, object]]
) -> list[dict[str, object]]:
    out: list[dict[str, object]] = []
    for snippet in snippets:
        source = sources[str(snippet["source"])]
        text = member_text(Path(str(source["body"])), str(source["kind"]), str(snippet["member"]))
        quote = str(snippet["quote"])
        if quote not in text:
            raise ValueError(f"{snippet['source']}:{snippet['member']}: quote not verbatim")
        out.append({**snippet, "verified_verbatim": True, "source_sha256": source["sha256"]})
    return out


def build() -> dict[str, object]:
    sealed = mapping(json.loads(BLIND.read_bytes()))
    sealed_at = utc_time(str(sealed["sealed_at_utc"]))
    adjudication = mapping(json.loads(ADJUDICATION.read_bytes()))
    observations = mapping(json.loads(OBSERVATIONS.read_bytes()))
    attempted = {str(mapping(r)["paper_id"]) for r in listed(sealed["runs"])}
    papers: list[dict[str, object]] = []
    scale = {str(s) for s in listed(observations["completeness_scale"])}
    dimensions = [str(d) for d in listed(observations["dimensions"])]
    for raw in listed(observations["papers"]):
        paper = mapping(raw)
        pid = str(paper["paper_id"])
        if pid not in attempted:
            raise ValueError(f"{pid}: not an attempted main-study paper")
        sources = {
            str(mapping(s)["identifier"]): verify_source(pid, mapping(s), sealed_at)
            for s in listed(paper["sources"])
        }
        snippets = verify_snippets([mapping(s) for s in listed(paper["snippets"])], sources)
        described = mapping(paper["dimensions"])
        if paper["reveal_assessment"] == "described":
            missing = [d for d in dimensions if d not in described]
            if missing or not sources:
                raise ValueError(f"{pid}: described reveal lacks {missing or 'sources'}")
            for name, entry in described.items():
                if mapping(entry)["completeness"] not in scale:
                    raise ValueError(f"{pid}:{name}: completeness outside scale")
        elif sources or snippets or described:
            raise ValueError(f"{pid}: missing reveal must not carry descriptions")
        papers.append(
            {
                "paper_id": pid,
                "original_code_status": paper["original_code_status"],
                "reveal_assessment": paper["reveal_assessment"],
                "sources": list(sources.values()),
                "snippets": snippets,
                "dimensions": described,
            }
        )
    if {p["paper_id"] for p in papers} != attempted:
        raise ValueError("reveal ledger must cover every attempted paper exactly once")
    return {
        "scope": observations["scope"],
        "study": observations["study"],
        "amendment": observations["amendment"],
        "reveal_reviewer": observations["reveal_reviewer"],
        "code_availability_source": observations["code_availability_source"],
        "blind_outcome_sha256": sha256(BLIND.read_bytes()),
        "blind_sealed_at_utc": sealed["sealed_at_utc"],
        "primary_adjudication_sha256": sha256(ADJUDICATION.read_bytes()),
        "primary_adjudication_at_utc": adjudication["adjudicated_at_utc"],
        "recorded_at_utc": datetime.now(timezone.utc).isoformat(),
        "original_code_executed": False,
        "blind_scores_revised_after_reveal": False,
        "human_validators_exposed_to_reveal": False,
        "counts": {
            "attempted_papers": len(papers),
            "described": sum(p["reveal_assessment"] == "described" for p in papers),
            "missing": sum(p["reveal_assessment"] == "missing" for p in papers),
        },
        "papers": papers,
    }


def write_results(ledger: dict[str, object]) -> None:
    rows: list[dict[str, object]] = []
    for raw in listed(ledger["papers"]):
        paper = mapping(raw)
        described = mapping(paper["dimensions"])
        if not described:
            rows.append(
                {
                    "paper_id": paper["paper_id"],
                    "original_code_status": paper["original_code_status"],
                    "reveal_assessment": paper["reveal_assessment"],
                    "dimension": "",
                    "completeness": "not_available",
                    "description": "",
                }
            )
        for name, entry in described.items():
            item = mapping(entry)
            rows.append(
                {
                    "paper_id": paper["paper_id"],
                    "original_code_status": paper["original_code_status"],
                    "reveal_assessment": paper["reveal_assessment"],
                    "dimension": name,
                    "completeness": item["completeness"],
                    "description": item["value"],
                }
            )
    write_csv(
        RESULTS / "reveal_ledger.csv",
        rows,
        [
            "paper_id",
            "original_code_status",
            "reveal_assessment",
            "dimension",
            "completeness",
            "description",
        ],
    )


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--dry-run", action="store_true")
    args = parser.parse_args()
    ledger = build()
    if args.dry_run:
        print(json.dumps(ledger["counts"], indent=2))
        return
    snapshot(LEDGER, (json.dumps(ledger, indent=2, sort_keys=True) + "\n").encode())
    stamped = stamp(LEDGER, Path("/home/ubuntu/paper2_evidence/main-run-20260925/reveal-timestamp"))
    write_results(ledger)
    print(
        json.dumps(
            {
                "status": "descriptive_reveal_recorded",
                "reveal_ledger_sha256": stamped["source_sha256"],
                "timestamp_status": stamped["status"],
                **mapping(ledger["counts"]),
            },
            indent=2,
        )
    )


if __name__ == "__main__":
    main()
