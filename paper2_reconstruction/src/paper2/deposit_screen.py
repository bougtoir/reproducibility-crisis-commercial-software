"""Screen candidates under amendment AMEND-2026-09-24-03 (deposited-data criterion).

Stage A (`statements`) walks every stratum in the unchanged deterministic
candidate order over the retained, identity-verified article text and records,
for every paper reached, the data-deposit statements it detects: repository
accessions, controlled-access consortia, request-only phrases, supplement-only
phrases and code-only repositories. Detection is regex evidence with segment
IDs and verbatim quotes, never an access verdict.

Stage B (`routes`) resolves, in that order and only for papers whose retained
text names an anonymously listable deposit, the public listing endpoint of each
named deposit, retains the listing response verbatim and applies the amendment
caps to the announced sizes. Every paper reached before the first
route-verified candidate of a stratum is retained with its exclusion class.
"""

from __future__ import annotations

import argparse
import json
import re
from datetime import datetime, timezone
from pathlib import Path
from xml.etree import ElementTree

from paper2.access_layer import AMENDMENT, CLASSES
from paper2.acquisition import acquire, candidate_order
from paper2.build import ROOT
from paper2.candidate_review import normalize, source_segments
from paper2.core import FIELDS, Row, read_csv, sha256, write_csv
from paper2.funnel_screen import retained_body
from paper2.model_api import mapping
from paper2.pilot_inputs import zenodo_files

FILE_CAP = 2 * 1024**3
TOTAL_CAP = 8 * 1024**3

# (kind, pattern, route). `open` routes have anonymous machine-readable listings
# that stage B knows how to resolve; `controlled` routes need an application;
# `code` routes are blocked by the firewall; `phrase` routes carry no deposit.
PATTERNS: tuple[tuple[str, str, str], ...] = (
    ("zenodo", r"10\.5281/zenodo\.(\d+)|zenodo\.org/records?/(\d+)", "open"),
    ("figshare", r"10\.6084/m9\.figshare\.(\d+)", "open"),
    ("dryad", r"10\.5061/dryad\.([0-9a-z]+)", "open"),
    ("osf", r"osf\.io/([a-z0-9]{5})\b", "open"),
    ("mendeley", r"10\.17632/([0-9a-z]+\.\d+)", "open"),
    ("dataverse", r"10\.7910/DVN/([0-9A-Z]+)", "open"),
    ("geo", r"\b(GSE\d{4,})\b", "open"),
    ("bioproject", r"\b(PRJ[NED][AB]\d+)\b", "open"),
    ("sra_study", r"\b([SED]RP\d{5,})\b", "open"),
    ("proteomexchange", r"\b(PXD\d{6})\b", "open"),
    ("metabolights", r"\b(MTBLS\d+)\b", "open"),
    ("arrayexpress", r"\b(E-[A-Z]{4}-\d+)\b", "open"),
    ("dbgap", r"\b(phs\d{6})\b", "controlled"),
    ("ega", r"\b(EGA[SD]\d{11})\b", "controlled"),
    ("adni", r"\b(ADNI)\b", "controlled"),
    ("ukbiobank", r"(UK Biobank)", "controlled"),
    ("github", r"github\.com/([\w.-]+/[\w.-]+)", "code"),
    (
        "request_only",
        r"((?:available|obtained|provided|accessible)[^.]{0,80}?(?:up)?on (?:reasonable )?request"
        r"|from the corresponding author[^.]{0,40}?(?:up)?on (?:reasonable )?request"
        r"|available (?:up)?on request)",
        "phrase",
    ),
    (
        "supplement_only",
        r"((?:included|contained|presented) in (?:this|the) (?:published )?article"
        r"(?:\s*/\s*| and (?:its )?| or (?:its )?)?(?:supplementary|supplemental)?"
        r"(?: (?:material|information|files))?)",
        "phrase",
    ),
    ("no_data", r"(No data are associated with this article|Not applicable)", "phrase"),
)
# Deposited files that are implementation material stay behind the firewall and
# never count towards eligibility (they are original code, not analysis data).
IMPLEMENTATION_FILE = re.compile(
    r"(^[\w.-]+/[\w.-]+-v?\d+(\.\d+)+[\w.-]*\.(zip|tar\.gz)$"  # GitHub release archive
    r"|(^|[_\-/ ])(scripts?|code|src|software|notebooks?|pipeline)([_\-/ .]|$)"
    r"|\.(py|r|rmd|ipynb|m|jl|sh|cpp|java|js)$)",
    re.IGNORECASE,
)
ROUTE_CLASS = {
    "controlled": "application_required",
    "code": "code_only_deposit",
    "request_only": "author_request_only",
    "supplement_only": "supplement_only_unverified",
    "no_data": "no_deposit_named",
}


def article_text(articles: Path, index: dict[str, Row], paper_id: str) -> str | None:
    row = index.get(paper_id)
    if row is None or row["status"] != "identity_verified_article":
        return None
    body = retained_body(articles, row["response_path"], row["response_sha256"])
    root = ElementTree.fromstring(body)
    return normalize("\n".join(root.itertext()))


def detect(text: str) -> list[dict[str, object]]:
    segments = source_segments(text)
    hits: list[dict[str, object]] = []
    seen: set[tuple[str, str]] = set()
    for segment_id, segment in segments.items():
        for kind, pattern, route in PATTERNS:
            for match in re.finditer(pattern, segment):
                accession = next((g for g in match.groups() if g), match.group(0))
                key = (kind, accession)
                if key in seen:
                    continue
                seen.add(key)
                hits.append(
                    {
                        "kind": kind,
                        "route": route,
                        "accession": accession,
                        "segment_id": segment_id,
                        "quote": match.group(0)[:200],
                    }
                )
    return hits


def statement_class(hits: list[dict[str, object]]) -> str:
    routes = {str(h["route"]) for h in hits}
    kinds = {str(h["kind"]) for h in hits}
    if "open" in routes:
        return "open_route_named_pending_listing"
    if "controlled" in routes:
        return "application_required"
    if "request_only" in kinds:
        return "author_request_only"
    if "code" in routes:
        return "code_only_deposit"
    if "supplement_only" in kinds:
        return "supplement_only_unverified"
    return "no_deposit_named"


def screen_statements(
    frame: list[Row], articles: Path, index: dict[str, Row], per_field: int
) -> list[dict[str, object]]:
    rows: list[dict[str, object]] = []
    for candidate in candidate_order(frame, per_field):
        paper_id = candidate["paper_id"]
        text = article_text(articles, index, paper_id)
        if text is None:
            rows.append(
                {
                    **candidate,
                    "article_text_status": "no_lawful_text_retained",
                    "statement_class": "no_lawful_text_retained",
                    "hits": [],
                }
            )
            continue
        hits = detect(text)
        rows.append(
            {
                **candidate,
                "article_text_status": "lawful_text_retained",
                "article_text_sha256": sha256(text.encode()),
                "statement_class": statement_class(hits),
                "hits": hits,
            }
        )
    return rows


def listing_url(kind: str, accession: str) -> str:
    if kind == "zenodo":
        return f"https://zenodo.org/api/records/{accession}"
    if kind == "figshare":
        return f"https://api.figshare.com/v2/articles/{accession}"
    if kind == "dryad":
        doi = f"doi:10.5061/dryad.{accession}".replace(":", "%3A").replace("/", "%2F")
        return f"https://datadryad.org/api/v2/datasets/{doi}"
    if kind == "osf":
        return f"https://api.osf.io/v2/guids/{accession}/"
    if kind == "mendeley":
        return f"https://data.mendeley.com/public-api/datasets/{accession.split('.')[0]}/files"
    if kind == "dataverse":
        return (
            "https://dataverse.harvard.edu/api/datasets/:persistentId/versions/:latest/files"
            f"?persistentId=doi:10.7910/DVN/{accession}"
        )
    if kind == "geo":
        stem = accession[:-3] + "nnn"
        return f"https://ftp.ncbi.nlm.nih.gov/geo/series/{stem}/{accession}/suppl/"
    if kind in {"bioproject", "sra_study"}:
        return (
            f"https://www.ebi.ac.uk/ena/portal/api/filereport?accession={accession}"
            "&result=read_run&fields=run_accession,fastq_ftp,fastq_bytes,submitted_bytes"
            "&format=tsv"
        )
    if kind == "proteomexchange":
        return (
            f"https://proteomecentral.proteomexchange.org/cgi/GetDataset?ID={accession}"
            "&outputMode=XML"
        )
    if kind == "metabolights":
        return (
            f"https://www.ebi.ac.uk/metabolights/ws/studies/{accession}/files?include_raw_data=true"
        )
    if kind == "arrayexpress":
        return f"https://www.ebi.ac.uk/biostudies/api/v1/studies/{accession}/info"
    raise ValueError(f"no listing route for {kind}")


def parse_listing(kind: str, body: bytes) -> list[dict[str, object]]:
    """Return [{name, bytes}] with bytes None when the service announces no size."""
    if kind == "zenodo":
        return [{"name": f["name"], "bytes": f["bytes"]} for f in zenodo_files(body)]
    if kind == "figshare":
        payload = mapping(json.loads(body))
        files = payload["files"]
        if not isinstance(files, list):
            raise ValueError("figshare record has no file list")
        return [{"name": mapping(f)["name"], "bytes": mapping(f)["size"]} for f in files]
    if kind == "dryad":
        payload = mapping(json.loads(body))
        embedded = mapping(payload.get("_embedded", {}))
        versions = embedded.get("stash:versions", [])
        dryad_files: list[dict[str, object]] = []
        if isinstance(versions, list):
            for version in versions:
                listed = mapping(mapping(version).get("_embedded", {})).get("stash:files", [])
                if not isinstance(listed, list):
                    continue
                for f in listed:
                    entry = mapping(f)
                    dryad_files.append({"name": entry["path"], "bytes": entry.get("size")})
        return dryad_files
    if kind == "osf":
        payload = mapping(json.loads(body))
        data = payload["data"]
        if not isinstance(data, list):
            raise ValueError("osf listing has no data list")
        return [
            {
                "name": mapping(mapping(f)["attributes"])["name"],
                "bytes": mapping(mapping(f)["attributes"]).get("size"),
            }
            for f in data
        ]
    if kind == "dataverse":
        payload = mapping(json.loads(body))
        data = payload["data"]
        if not isinstance(data, list):
            raise ValueError("dataverse listing has no data list")
        return [
            {
                "name": mapping(mapping(f)["dataFile"])["filename"],
                "bytes": mapping(mapping(f)["dataFile"]).get("filesize"),
            }
            for f in data
        ]
    if kind == "geo":
        text = body.decode("utf-8", "replace")
        geo_files: list[dict[str, object]] = []
        for match in re.finditer(
            r'href="([^"?/]+)"[^\n]*?(\d{4}-\d\d-\d\d \d\d:\d\d)\s+(\S+)', text
        ):
            size = match.group(3)
            geo_files.append({"name": match.group(1), "bytes": _ftp_size(size)})
        return geo_files
    if kind in {"bioproject", "sra_study"}:
        lines = body.decode("utf-8", "replace").strip().splitlines()
        if not lines:
            return []
        header = lines[0].split("\t")
        runs: list[dict[str, object]] = []
        for line in lines[1:]:
            cells = dict(zip(header, line.split("\t"), strict=False))
            for column in ("fastq_bytes", "submitted_bytes"):
                for size in cells.get(column, "").split(";"):
                    if size:
                        runs.append(
                            {"name": f"{cells['run_accession']}:{column}", "bytes": int(size)}
                        )
        return runs
    if kind == "proteomexchange":
        root = ElementTree.fromstring(body)
        return [
            {"name": node.attrib.get("name", node.attrib.get("id", "")), "bytes": None}
            for node in root.iter("DatasetFile")
        ]
    if kind == "metabolights":
        payload = mapping(json.loads(body))
        study = payload.get("study", [])
        if not isinstance(study, list):
            raise ValueError("metabolights listing has no file list")
        return [{"name": mapping(f)["file"], "bytes": None} for f in study]
    if kind == "arrayexpress":
        payload = mapping(json.loads(body))
        return [{"name": "biostudies_info", "bytes": payload.get("ftpLink") and None}]
    raise ValueError(f"no parser for {kind}")


def _ftp_size(text: str) -> int | None:
    units = {"K": 1024, "M": 1024**2, "G": 1024**3}
    if text.isdigit():
        return int(text)
    if text[-1] in units and text[:-1].replace(".", "", 1).isdigit():
        return int(float(text[:-1]) * units[text[-1]])
    return None


def _get(url: str, label: str, directory: Path) -> tuple[bytes | None, dict[str, object]]:
    path, receipt = acquire(
        url,
        label,
        directory,
        request_conditions="GET anonymous public deposit listing; no implementation material",
    )
    return (path.read_bytes() if receipt["http_status"] == 200 else None), receipt


def follow_listing(
    kind: str, accession: str, body: bytes, directory: Path, record: dict[str, object]
) -> list[dict[str, object]] | str:
    """Resolve the second request some services need before files are listed.

    Returns the file list, or a route status string when the identifier turns
    out not to be a data deposit or the follow-up request failed.
    """
    label = f"{kind}:{accession}"
    if kind == "dryad":
        links = mapping(mapping(json.loads(body))["_links"])
        href = str(mapping(links["stash:version"])["href"])
        files_body, receipt = _get(f"https://datadryad.org{href}/files", label, directory)
        record["follow_up"] = receipt
        if files_body is None:
            return f"version_files_http_{receipt['http_status']}"
        embedded = mapping(mapping(json.loads(files_body))["_embedded"])
        listed = embedded["stash:files"]
        if not isinstance(listed, list):
            raise ValueError("dryad version has no file list")
        return [{"name": mapping(f)["path"], "bytes": mapping(f).get("size")} for f in listed]
    if kind == "osf":
        data = mapping(mapping(json.loads(body))["data"])
        guid_type = str(data["type"])
        if guid_type == "preprints":
            return "identifier_is_a_preprint_not_a_data_deposit"
        if guid_type not in {"nodes", "registrations"}:
            return f"identifier_is_an_osf_{guid_type}_not_a_data_deposit"
        providers_body, receipt = _get(
            f"https://api.osf.io/v2/{guid_type}/{accession}/files/", label, directory
        )
        record["follow_up"] = receipt
        if providers_body is None:
            return f"providers_http_{receipt['http_status']}"
        providers = mapping(json.loads(providers_body))["data"]
        if not isinstance(providers, list):
            raise ValueError("osf node has no provider list")
        files: list[dict[str, object]] = []
        for provider in providers:
            name = str(mapping(mapping(provider)["attributes"])["provider"])
            listing_body, listing_receipt = _get(
                f"https://api.osf.io/v2/{guid_type}/{accession}/files/{name}/", label, directory
            )
            record[f"provider_{name}"] = listing_receipt
            if listing_body is None:
                continue
            entries = mapping(json.loads(listing_body))["data"]
            if not isinstance(entries, list):
                continue
            for entry in entries:
                attributes = mapping(mapping(entry)["attributes"])
                suffix = "/" if attributes["kind"] == "folder" else ""
                files.append(
                    {
                        "name": f"{name}:{attributes['name']}{suffix}",
                        "bytes": attributes.get("size"),
                    }
                )
        return files
    return parse_listing(kind, body)


def resolve_route(hit: dict[str, object], destination: Path) -> dict[str, object]:
    kind, accession = str(hit["kind"]), str(hit["accession"])
    url = listing_url(kind, accession)
    directory = destination / kind / accession.replace("/", "_")
    body, receipt = _get(url, f"{kind}:{accession}", directory)
    record: dict[str, object] = {
        "kind": kind,
        "accession": accession,
        "listing_url": url,
        "listing": receipt,
    }
    if body is None:
        record["exclusion_class"] = "listing_unavailable"
        record["route_status"] = f"listing_http_{receipt['http_status']}"
        return record
    try:
        resolved = follow_listing(kind, accession, body, directory, record)
    except (ValueError, KeyError, TypeError, ElementTree.ParseError):
        record["exclusion_class"] = "listing_unavailable"
        record["route_status"] = "listing_not_machine_readable"
        return record
    if isinstance(resolved, str):
        record["exclusion_class"] = (
            "no_deposit_named" if "not_a_data_deposit" in resolved else "listing_unavailable"
        )
        record["route_status"] = resolved
        return record
    files = resolved
    all_files = files
    for entry in files:
        entry["role"] = (
            "implementation_withheld"
            if IMPLEMENTATION_FILE.search(str(entry["name"]))
            else "candidate_analysis_data"
        )
    record["listed_files"] = len(files)
    record["implementation_files_withheld"] = sum(
        1 for f in files if f["role"] == "implementation_withheld"
    )
    files = [f for f in files if f["role"] == "candidate_analysis_data"]
    record["candidate_data_files"] = len(files)
    sized = [f for f in files if isinstance(f["bytes"], int)]
    record["files_with_size"] = len(sized)
    listed_bytes = sum(int(str(f["bytes"])) for f in sized)
    largest = max((int(str(f["bytes"])) for f in sized), default=0)
    record["listed_bytes"] = listed_bytes
    record["largest_file_bytes"] = largest
    record["files"] = all_files[:500]
    if not files and record["implementation_files_withheld"]:
        record["exclusion_class"] = "code_only_deposit"
        record["route_status"] = "every_listed_file_is_implementation_material"
    elif not files:
        record["exclusion_class"] = "listing_unavailable"
        record["route_status"] = "listing_empty"
    elif len(sized) != len(files):
        record["exclusion_class"] = "size_unannounced"
        record["route_status"] = "some_or_all_files_without_announced_size"
    elif largest > FILE_CAP or listed_bytes > TOTAL_CAP:
        record["exclusion_class"] = "above_cap"
        record["route_status"] = "announced_sizes_exceed_amendment_caps"
    else:
        record["exclusion_class"] = ""
        record["route_status"] = "anonymous_listing_with_sizes_within_caps"
    return record


def prior_classes(second_pass: Path) -> dict[str, str]:
    """Access classes already recorded for papers assessed before the amendment."""
    payload = mapping(json.loads(second_pass.read_bytes()))
    rows = payload["assessments"]
    if not isinstance(rows, list):
        raise ValueError("second-pass record has no assessments")
    return {
        str(mapping(r)["paper_id"]): str(mapping(r)["access_class_under_amendment"]) for r in rows
    }


def amended_record_classes(record: Path) -> dict[str, str]:
    """Exclusion classes decided after deposit inspection of route-eligible candidates."""
    payload = mapping(json.loads(record.read_bytes()))
    rows = payload["assessments"]
    if not isinstance(rows, list):
        raise ValueError("amended assessment record has no assessments")
    classes = {}
    for r in rows:
        disposition = str(mapping(r)["amended_disposition"])
        if disposition.startswith("excluded_"):
            classes[str(mapping(r)["paper_id"])] = disposition.removeprefix("excluded_")
    return classes


def screened_papers(route_record: Path) -> set[str]:
    rows = mapping(json.loads(route_record.read_bytes()))["rows"]
    if not isinstance(rows, list):
        raise ValueError("route record has no rows")
    return {str(mapping(r)["paper_id"]) for r in rows}


def screen_routes(
    statements: list[dict[str, object]],
    destination: Path,
    per_field_limit: int,
    prior: dict[str, str],
    fields: tuple[str, ...] = FIELDS,
    already_screened: frozenset[str] = frozenset(),
    exhaustive: bool = False,
) -> list[dict[str, object]]:
    """Resolve open routes stratum by stratum until one candidate passes the caps.

    ``already_screened`` papers were consumed by an earlier Stage B record and are
    skipped without re-resolution, so a continuation run resumes the deterministic
    chain where the earlier record stopped. With ``exhaustive`` every paper of every
    stratum is resolved, which enumerates the route-eligible population.
    """
    rows = []
    for field in fields:
        passed = 0
        for row in statements:
            if row["sampling_stratum"] != field or (passed and not exhaustive):
                continue
            if str(row["paper_id"]) in already_screened:
                continue
            hits = row["hits"]
            if not isinstance(hits, list):
                continue
            open_hits = [mapping(h) for h in hits if mapping(h)["route"] == "open"]
            entry = {
                "paper_id": row["paper_id"],
                "sampling_stratum": field,
                "candidate_rank": row["candidate_rank"],
                "statement_class": row["statement_class"],
            }
            paper_id = str(row["paper_id"])
            if paper_id in prior:
                entry["exclusion_class"] = prior[paper_id]
                entry["route_status"] = "class_recorded_in_prior_G1_G5_or_deposit_inspection_record"
                entry["routes"] = []
                rows.append(entry)
                continue
            if not open_hits:
                entry["exclusion_class"] = (
                    row["statement_class"]
                    if row["statement_class"] in CLASSES
                    else "no_deposit_named"
                )
                entry["routes"] = []
                rows.append(entry)
                continue
            routes = [resolve_route(h, destination) for h in open_hits[:per_field_limit]]
            entry["routes"] = routes
            eligible = [r for r in routes if not r["exclusion_class"]]
            if eligible:
                entry["exclusion_class"] = ""
                entry["route_status"] = "deposit_route_verified_pending_G1_G2_and_leakage_review"
                passed += 1
            else:
                entry["exclusion_class"] = str(routes[0]["exclusion_class"])
            rows.append(entry)
    return rows


def retain_sources(
    route_rows: list[dict[str, object]], articles: Path, index: dict[str, Row], destination: Path
) -> list[str]:
    """Write hash-checked segmented article text for every route-eligible candidate.

    The layout matches ``primary_adjudication.article_segments`` so that G1-G5
    assessments of the amended candidates cite the same kind of retained evidence.
    """
    retained = []
    for row in route_rows:
        if row["exclusion_class"] != "":
            continue
        paper_id = str(row["paper_id"])
        article = index[paper_id]
        body = retained_body(articles, article["response_path"], article["response_sha256"])
        text = normalize("\n".join(ElementTree.fromstring(body).itertext()))
        payload = {
            "paper_id": paper_id,
            "scope": "amended_candidate_article_text_for_G1_G5_assessment",
            "amendment_id": AMENDMENT,
            "sources": [
                {
                    "source_sha256": sha256(body),
                    "text_sha256": sha256(text.encode()),
                    "role": "standalone_article_xml",
                    "retained_response_path": article["response_path"],
                    "segments": [
                        {"segment_id": k, "text": v} for k, v in source_segments(text).items()
                    ],
                }
            ],
        }
        target = destination / paper_id.removeprefix("PMID:") / "sources.json"
        target.parent.mkdir(parents=True, exist_ok=True)
        if not target.exists():
            target.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n")
        retained.append(paper_id)
    return retained


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    sub = parser.add_subparsers(dest="command", required=True)
    statements = sub.add_parser("statements")
    statements.add_argument("--per-field", type=int, default=40)
    statements.add_argument(
        "--articles", type=Path, default=ROOT / "data/raw/corpus-articles-20260922"
    )
    statements.add_argument("--output", type=Path, required=True)
    routes = sub.add_parser("routes")
    routes.add_argument("--statements", type=Path, required=True)
    routes.add_argument("--destination", type=Path, required=True)
    routes.add_argument("--output", type=Path, required=True)
    routes.add_argument("--routes-per-paper", type=int, default=4)
    routes.add_argument(
        "--second-pass",
        type=Path,
        default=ROOT / "data" / "adjudication" / "devin_second_pass_G1_G5_20260924.json",
    )
    routes.add_argument("--fields", nargs="*", default=list(FIELDS))
    routes.add_argument("--continue-from", type=Path, default=None)
    routes.add_argument("--amended-record", type=Path, default=None)
    routes.add_argument("--exhaustive", action="store_true")
    retain = sub.add_parser("retain")
    retain.add_argument("--routes", type=Path, required=True)
    retain.add_argument("--articles", type=Path, default=ROOT / "data/raw/corpus-articles-20260922")
    retain.add_argument("--destination", type=Path, required=True)
    args = parser.parse_args()
    if args.command == "retain":
        index = {r["paper_id"]: r for r in read_csv(args.articles / "article_index.csv")}
        route_rows = mapping(json.loads(args.routes.read_bytes()))["rows"]
        if not isinstance(route_rows, list):
            raise ValueError("routes file has no rows")
        print(
            retain_sources([mapping(r) for r in route_rows], args.articles, index, args.destination)
        )
        return
    if args.command == "statements":
        frame = read_csv(ROOT / "data/derived/inference_frame.csv")
        index = {r["paper_id"]: r for r in read_csv(args.articles / "article_index.csv")}
        rows = screen_statements(frame, args.articles, index, args.per_field)
        payload = {
            "amendment_id": AMENDMENT,
            "stage": "A_statements",
            "per_field": args.per_field,
            "recorded_utc": datetime.now(timezone.utc).isoformat(),
            "rows": rows,
        }
        args.output.write_text(json.dumps(payload, indent=2, ensure_ascii=False) + "\n")
        counts: dict[str, int] = {}
        for row in rows:
            key = f"{row['sampling_stratum']}:{row['statement_class']}"
            counts[key] = counts.get(key, 0) + 1
        print(json.dumps(dict(sorted(counts.items())), indent=2))
    else:
        payload = mapping(json.loads(args.statements.read_bytes()))
        stage_a = payload["rows"]
        if not isinstance(stage_a, list):
            raise ValueError("statements file has no rows")
        prior = prior_classes(args.second_pass)
        inspected = (
            amended_record_classes(args.amended_record) if args.amended_record is not None else {}
        )
        prior.update(inspected)
        screened = (
            frozenset(screened_papers(args.continue_from) - set(inspected))
            if args.continue_from is not None
            else frozenset()
        )
        rows = screen_routes(
            [mapping(r) for r in stage_a],
            args.destination,
            args.routes_per_paper,
            prior,
            tuple(args.fields),
            screened,
            args.exhaustive,
        )
        stage = "B_routes" if args.continue_from is None else "B_routes_continuation"
        out = {
            "amendment_id": AMENDMENT,
            "stage": "B_routes_exhaustive" if args.exhaustive else stage,
            "continues_from_sha256": (
                sha256(args.continue_from.read_bytes()) if args.continue_from is not None else ""
            ),
            "fields": list(args.fields),
            "statements_sha256": sha256(args.statements.read_bytes()),
            "file_cap_bytes": FILE_CAP,
            "total_cap_bytes": TOTAL_CAP,
            "recorded_utc": datetime.now(timezone.utc).isoformat(),
            "rows": rows,
        }
        args.output.write_text(json.dumps(out, indent=2, ensure_ascii=False) + "\n")
        write_csv(
            args.output.with_suffix(".csv"),
            [
                {
                    "paper_id": r["paper_id"],
                    "sampling_stratum": r["sampling_stratum"],
                    "candidate_rank": r["candidate_rank"],
                    "statement_class": r["statement_class"],
                    "exclusion_class": r["exclusion_class"],
                    "routes": "; ".join(
                        f"{mapping(x)['kind']}:{mapping(x)['accession']}={mapping(x)['route_status']}"
                        for x in r["routes"]
                    )
                    if isinstance(r["routes"], list)
                    else "",
                }
                for r in rows
            ],
            (
                "paper_id",
                "sampling_stratum",
                "candidate_rank",
                "statement_class",
                "exclusion_class",
                "routes",
            ),
        )
        for r in rows:
            print(
                r["sampling_stratum"],
                r["candidate_rank"],
                r["paper_id"],
                r["exclusion_class"] or "ELIGIBLE_ROUTE",
            )


if __name__ == "__main__":
    main()
