"""Provisional G1/G2 screening of frame papers from retained corpus evidence.

Sources are the identity-verified Europe PMC metadata record (title and
abstract) and, when retained, the identity-verified open PMC JATS article of
the same paper. Every observation is bound to immutable retained bytes. The
output is provisional machine review for prospective pilot screening; it is not
a validated G1-G5 assessment, an input-access verdict or a selected target.
"""

import argparse
import json
from pathlib import Path
from xml.etree import ElementTree

from paper2.acquisition import candidate_order
from paper2.build import ROOT
from paper2.candidate_review import SYSTEM, Source, normalize, source_segments, validate_review
from paper2.core import Row, read_csv, sha256, snapshot, write_csv
from paper2.model_api import MODEL, bounded_complete, mapping

ARTICLE_TEXT_STATUS = {
    "identity_verified_article": "identity_verified_open_pmc_jats",
    "pmc_endpoint_not_selected_other_sources_unassessed": (
        "no_open_pmc_endpoint_other_routes_unassessed"
    ),
    "request_failed_not_access_verdict": "open_pmc_request_failed_not_access_verdict",
    "unverified_article_response": "open_pmc_response_unverified",
}


def retained_body(root: Path, relative: str, digest: str) -> bytes:
    path = (root / relative).resolve()
    if root.resolve() not in path.parents:
        raise ValueError("Retained source path escapes the acquisition directory")
    body = path.read_bytes()
    if sha256(body) != digest:
        raise ValueError("Retained source differs from the indexed hash")
    return body


def metadata_source(metadata: Path, row: Row) -> Source:
    body = retained_body(metadata, row["source_snapshot"], row["source_snapshot_sha256"])
    results = mapping(json.loads(body)["resultList"])["result"]
    if not isinstance(results, list):
        raise ValueError("Metadata response has no result list")
    records = [
        mapping(item)
        for item in results
        if mapping(item)["id"] == row["pmid"] and mapping(item)["source"] == "MED"
    ]
    if len(records) != 1:
        raise ValueError("Metadata snapshot does not contain exactly one record for the paper")
    record = records[0]
    text = normalize(str(record.get("title", "")) + "\n" + str(record.get("abstractText", "")))
    return Source(sha256(body), sha256(text.encode()), "title_and_abstract_only", text)


def article_source(articles: Path, row: Row) -> Source:
    body = retained_body(articles, row["response_path"], row["response_sha256"])
    root = ElementTree.fromstring(body)
    if root.tag != "article":
        raise ValueError("Retained article is not a JATS article")
    text = normalize("\n".join(root.itertext()))
    if len(text) > 500_000:
        raise ValueError("Article exceeds review context bound; not truncated")
    return Source(sha256(body), sha256(text.encode()), "standalone_article_xml", text)


def screen_paper(
    paper: Row,
    metadata_row: Row,
    article_row: Row,
    metadata: Path,
    articles: Path,
    case: Path,
) -> Row:
    if metadata_row["status"] != "identity_verified":
        return {
            **paper,
            "status": "metadata_not_identity_verified_unscreened",
            "article_text_status": ARTICLE_TEXT_STATUS[article_row["status"]],
            "computationally_testable": "uncertain",
            "principal_target_identifiable": "uncertain",
            "sources_sha256": "",
            "review_sha256": "",
            "provider_accounted_tokens": "unknown",
        }
    sources = [metadata_source(metadata, metadata_row)]
    if article_row["status"] == "identity_verified_article":
        sources.append(article_source(articles, article_row))
    context = {
        "paper_id": paper["paper_id"],
        "scope": "provisional_pilot_screening_not_validated_funnel_or_outcome",
        "system_prompt_sha256": sha256(SYSTEM.encode()),
        "sources": [
            {
                "source_sha256": item.source_sha256,
                "text_sha256": item.text_sha256,
                "role": item.role,
                "segments": [
                    {"segment_id": segment_id, "text": text}
                    for segment_id, text in source_segments(item.text).items()
                ],
            }
            for item in sources
        ],
    }
    payload = json.dumps(context, ensure_ascii=False)
    snapshot(case / "sources.json", payload.encode())
    completion_path = case / "bounded_api/api/completion.json"
    if not (case / "bounded_api").exists():
        try:
            bounded_complete(
                [{"role": "system", "content": SYSTEM}, {"role": "user", "content": payload}],
                case / "bounded_api",
                4096,
            )
        except RuntimeError:
            pass
    if not completion_path.exists():
        return {
            **paper,
            "status": "api_call_incomplete_retained_not_retried",
            "article_text_status": ARTICLE_TEXT_STATUS[article_row["status"]],
            "computationally_testable": "uncertain",
            "principal_target_identifiable": "uncertain",
            "sources_sha256": sha256(payload.encode()),
            "review_sha256": "",
            "provider_accounted_tokens": "unknown",
        }
    response = mapping(json.loads(completion_path.read_text()))
    if response["reported_model"] != MODEL or response["finish_reason"] != "stop":
        raise ValueError("Incomplete or wrong-model screening response")
    try:
        reviewed = validate_review(json.loads(str(response["content"])), sources)
        snapshot(case / "review.json", (json.dumps(reviewed, indent=2) + "\n").encode())
        status = "provisional_machine_review_not_validated_or_selected"
    except ValueError as error:
        reviewed = {}
        status = "rejected_source_evidence"
        snapshot(case / "rejected.json", (json.dumps({"error": str(error)}) + "\n").encode())
    has_article = len(sources) == 2
    g1 = str(reviewed.get("computationally_testable", "uncertain"))
    g2 = str(reviewed.get("principal_target_identifiable", "uncertain"))
    if not has_article and (g1 == "no" or g2 == "no"):
        g1 = "uncertain" if g1 == "no" else g1
        g2 = "uncertain" if g2 == "no" else g2
        status = "abstract_only_negative_not_validated"
    return {
        **paper,
        "status": status,
        "article_text_status": ARTICLE_TEXT_STATUS[article_row["status"]],
        "computationally_testable": g1,
        "principal_target_identifiable": g2,
        "sources_sha256": sha256(payload.encode()),
        "review_sha256": sha256(completion_path.read_bytes()),
        "provider_accounted_tokens": str(
            int(str(response["prompt_tokens"])) + int(str(response["completion_tokens"]))
        ),
    }


def screen_pilot_candidates(
    metadata: Path, articles: Path, destination: Path, per_field: int
) -> list[Row]:
    frame_path = ROOT / "data/derived/inference_frame.csv"
    frame = read_csv(frame_path)
    frame_sha256 = sha256(frame_path.read_bytes())
    metadata_rows = {row["paper_id"]: row for row in read_csv(metadata / "metadata_index.csv")}
    article_rows = {row["paper_id"]: row for row in read_csv(articles / "article_index.csv")}
    if (articles / "metadata_index_snapshot.csv").read_bytes() != (
        metadata / "metadata_index.csv"
    ).read_bytes():
        raise ValueError("Article acquisition uses another metadata snapshot")
    destination.mkdir(parents=True, exist_ok=True)
    rows = []
    for candidate in candidate_order(frame, per_field):
        metadata_row = metadata_rows[candidate["paper_id"]]
        article_row = article_rows[candidate["paper_id"]]
        if frame_sha256 not in (metadata_row["frame_sha256"], article_row["frame_sha256"]) or (
            metadata_row["frame_sha256"] != article_row["frame_sha256"]
        ):
            raise ValueError("Retained evidence is bound to another frame")
        paper = {
            "paper_id": candidate["paper_id"],
            "sampling_stratum": candidate["sampling_stratum"],
            "candidate_rank": candidate["candidate_rank"],
            "rank_sha256": candidate["rank_sha256"],
            "frame_sha256": frame_sha256,
            "metadata_status": metadata_row["status"],
        }
        row = screen_paper(
            paper, metadata_row, article_row, metadata, articles, destination / candidate["pmid"]
        )
        rows.append(row)
        write_csv(destination / "screen_index.csv", rows, list(rows[0]))
        print(
            f"{row['paper_id']} {row['sampling_stratum']} rank={row['candidate_rank']} "
            f"text={row['article_text_status']} G1={row['computationally_testable']} "
            f"G2={row['principal_target_identifiable']} {row['status']}",
            flush=True,
        )
    return rows


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--metadata", type=Path, required=True)
    parser.add_argument("--articles", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--candidates-per-field", type=int, default=1)
    args = parser.parse_args()
    screen_pilot_candidates(args.metadata, args.articles, args.output, args.candidates_per_field)


if __name__ == "__main__":
    main()
