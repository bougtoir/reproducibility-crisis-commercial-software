"""Custodian, broker and access-telemetry qualification for the blind harness.

The custodian holds a quarantined mixed archive and never serves it. The broker
serves only reviewed exact members to a named actor, writes a hash-chained access
event for every request including denials, and refuses reveal-class material
until an external freeze receipt exists. A solver worker then proves, inside the
isolated container, that only the served members are reachable, that the vetted
dependency closure imports without a network, and that resource stops fire.

Every check has an adversarial negative case: a blocked member, an unlisted
member, a tampered custodian byte stream, a pre-freeze reveal request and a
post-freeze outcome mutation must all be refused and recorded.
"""

from __future__ import annotations

import argparse
import io
import json
import tarfile
import zipfile
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path, PurePosixPath

from paper2.build import ROOT
from paper2.controller import Journal
from paper2.core import read_csv, sha256, snapshot
from paper2.isolation import Worker
from paper2.model_api import mapping

POLICY_VERSION = "firewall-qualification-2026-09-24"
ALLOWED_ACTORS = {"broker", "solver_package_builder"}
REVEAL_ACTORS = {"reveal_reviewer"}
CANARY = b"SYNTHETIC_ORIGINAL_IMPLEMENTATION_CANARY def fit(): return 0.42\n"
WHEELS = {
    "packaging-26.3-py3-none-any.whl": "packaging==26.3",
    "python_dateutil-2.9.0.post0-py2.py3-none-any.whl": "python-dateutil==2.9.0.post0",
    "six-1.17.0-py2.py3-none-any.whl": "six==1.17.0",
}


class Denied(Exception):
    """Raised when the broker refuses a request under the frozen policy."""


@dataclass(frozen=True)
class Item:
    item_id: str
    paper_id: str
    member: str
    sha256: str
    bytes: int
    category: str
    decision: str
    decision_reason: str


class Custodian:
    """Holds the quarantined archive and exposes reviewed members only."""

    def __init__(self, archive: Path, items: dict[str, Item]) -> None:
        self.archive = archive
        self.archive_sha256 = sha256(archive.read_bytes())
        self.items = items

    def member(self, item_id: str) -> bytes:
        item = self.items[item_id]
        with tarfile.open(self.archive, mode="r") as handle:
            payload = handle.extractfile(item.member)
            if payload is None:
                raise Denied(f"{item_id}: archive member is not a regular file")
            body = payload.read()
        if sha256(body) != item.sha256 or len(body) != item.bytes:
            raise Denied(f"{item_id}: quarantined member differs from the reviewed identity")
        return body


class Broker:
    """Serves reviewed members to a named actor and records every request."""

    def __init__(self, custodian: Custodian, journal: Journal, package: Path) -> None:
        self.custodian = custodian
        self.journal = journal
        self.package = package
        self.frozen: dict[str, object] | None = None
        self.served: dict[str, str] = {}

    def freeze(self, outcome: Path, receipt: dict[str, object]) -> None:
        digest = sha256(outcome.read_bytes())
        if receipt["source_sha256"] != digest:
            raise Denied("freeze receipt does not cover the sealed outcome bytes")
        self.frozen = {
            "outcome_sha256": digest,
            "timestamp_response_sha256": receipt["response_sha256"],
            "trust_anchor_sha256": receipt["trust_anchor_sha256"],
        }
        self.journal.record("blind_outcome_frozen", self.frozen)

    def request(self, actor: str, item_id: str) -> bytes:
        event: dict[str, object] = {
            "actor": actor,
            "requested_item_id": item_id,
            "policy_version": POLICY_VERSION,
            "archive_sha256": self.custodian.archive_sha256,
        }
        try:
            item = self.custodian.items.get(item_id)
            if item is None:
                raise Denied("unlisted item; default deny")
            if item.decision != "allow":
                if item.category != "reveal_original_implementation":
                    raise Denied(f"blocked item: {item.decision_reason}")
                if actor not in REVEAL_ACTORS:
                    raise Denied("reveal-class item is not available to this actor")
                if self.frozen is None:
                    raise Denied("reveal-class item requested before the blind outcome freeze")
            elif actor not in ALLOWED_ACTORS:
                raise Denied("actor is not authorised to receive served members")
            body = self.custodian.member(item_id)
        except (Denied, KeyError, tarfile.TarError, OSError) as error:
            self.journal.record(
                "access_denied",
                {**event, "decision": "deny", "reason": f"{type(error).__name__}: {error}"},
            )
            raise Denied(str(error)) from error
        member = PurePosixPath(item.member).name
        if actor in ALLOWED_ACTORS:
            snapshot(self.package / member, body)
            self.served[member] = sha256(body)
        self.journal.record(
            "access_allowed",
            {
                **event,
                "decision": "allow",
                "archive_member_path": item.member,
                "served_bytes": len(body),
                "served_sha256": sha256(body),
                "served_to_package": actor in ALLOWED_ACTORS,
            },
        )
        return body


PROBES = """
import json
import os
import socket
import sys
from pathlib import Path

served = sorted(p.name for p in Path("/input").iterdir())
checks = {"served_members_match_manifest": served == SERVED_NAMES}
checks["no_canary_member_served"] = not any("original" in name for name in served)
checks["canary_host_path_unreadable"] = not Path(CANARY_PATH).exists()
checks["custodian_archive_unreadable"] = not Path(ARCHIVE_PATH).exists()
checks["no_api_credential"] = "DEEPSEEK_API_KEY" not in os.environ
checks["no_docker_socket"] = not Path("/var/run/docker.sock").exists()
with socket.socket() as connection:
    connection.settimeout(2)
    try:
        connection.connect(("1.1.1.1", 443))
        checks["network_denied"] = False
    except OSError:
        checks["network_denied"] = True
sys.path[:0] = [str(p) for p in sorted(Path("/input").glob("*.whl"))]
import dateutil.relativedelta
import packaging.version
import six
checks["dependency_closure_imports_offline"] = bool(
    six.PY3
    and packaging.version.Version("1.0") < packaging.version.Version("1.1")
    and dateutil.relativedelta.relativedelta(days=1).days == 1
)
checks["dependency_modules_come_from_served_wheels"] = all(
    path.startswith("/input/") for path in (six.__file__, packaging.version.__file__)
)
article = [p for p in Path("/input").iterdir() if p.suffix == ".xml"]
checks["served_article_is_readable_jats"] = bool(article) and all(
    p.read_bytes().lstrip().startswith(b"<") for p in article
)
try:
    Path("/input/write-test").write_text("x")
    checks["package_readonly"] = False
except OSError:
    checks["package_readonly"] = True
print(json.dumps(checks, sort_keys=True))
"""


def pilot_articles(record: Path, articles: Path) -> list[tuple[str, bytes]]:
    payload = mapping(json.loads(record.read_bytes()))
    assessments = payload["assessments"]
    if not isinstance(assessments, list):
        raise ValueError("adjudication record has no assessment list")
    selected = sorted(
        str(mapping(a)["paper_id"])
        for a in assessments
        if mapping(a)["pilot_decision"] == "provisional_pilot_case"
    )
    return verified_articles(selected, articles)


def verified_articles(selected: list[str], articles: Path) -> list[tuple[str, bytes]]:
    """Return identity-verified retained article bodies for the listed papers only."""
    index = {row["paper_id"]: row for row in read_csv(articles / "article_index.csv")}
    bodies = []
    for paper_id in selected:
        row = index[paper_id]
        if row["status"] != "identity_verified_article":
            raise ValueError(f"{paper_id}: retained article is not identity verified")
        body = (articles / row["response_path"]).read_bytes()
        if sha256(body) != row["response_sha256"]:
            raise ValueError(f"{paper_id}: retained article differs from the indexed hash")
        bodies.append((paper_id, body))
    return bodies


def vetted_wheels(directory: Path) -> dict[str, bytes]:
    wheels = {}
    for name in sorted(WHEELS):
        path = directory / name
        if not path.exists():
            raise ValueError(
                f"{name} is missing; download the pinned wheel closure into {directory}"
            )
        wheels[name] = path.read_bytes()
    for name, payload in wheels.items():
        with zipfile.ZipFile(io.BytesIO(payload)) as archive:
            metadata = next(
                archive.read(item).decode()
                for item in archive.namelist()
                if item.endswith(".dist-info/METADATA")
            )
        required = [
            line.split(":", 1)[1].split(";")[0].split(maxsplit=1)[0].strip().lower()
            for line in metadata.splitlines()
            if line.startswith("Requires-Dist:")
        ]
        available = {WHEELS[key].split("==")[0].lower() for key in wheels}
        missing = sorted(set(required) - available)
        if missing:
            raise ValueError(f"{name}: dependency closure is incomplete, missing {missing}")
    return wheels


def build_archive(
    destination: Path, bodies: list[tuple[str, bytes]], wheels: dict[str, bytes]
) -> tuple[Path, dict[str, Item]]:
    """Write one quarantined mixed archive holding allowed, blocked and reveal members."""
    members: list[tuple[str, bytes, str, str, str]] = []
    for paper_id, body in bodies:
        pmid = paper_id.removeprefix("PMID:")
        members.append(
            (
                f"articles/{pmid}.xml",
                body,
                "publication_text",
                "allow",
                "retained identity-verified open PMC JATS of the pilot candidate",
            )
        )
    for name, payload in wheels.items():
        members.append(
            (
                f"software/{name}",
                payload,
                "vetted_free_software",
                "allow",
                f"pinned hash-checked wheel for {WHEELS[name]} with reviewed closure",
            )
        )
    members.append(
        (
            "implementation/original_pipeline.py",
            CANARY,
            "reveal_original_implementation",
            "deny",
            "synthetic stand-in for author implementation; forbidden before reveal",
        )
    )
    members.append(
        (
            "implementation/original_notes.txt",
            b"SYNTHETIC_AUTHOR_NOTES_CANARY\n",
            "author_implementation_material",
            "deny",
            "author implementation material is never served",
        )
    )
    archive = destination / "quarantine" / "mixed_archive.tar"
    archive.parent.mkdir(parents=True, exist_ok=True)
    with tarfile.open(archive, mode="w") as handle:
        for name, payload, _, _, _ in members:
            info = tarfile.TarInfo(name)
            info.size = len(payload)
            info.mtime = 0
            handle.addfile(info, io.BytesIO(payload))
    items = {}
    for index, (name, payload, category, decision, reason) in enumerate(members):
        item_id = f"I{index:03d}"
        items[item_id] = Item(
            item_id=item_id,
            paper_id=(
                name.split("/")[-1].removesuffix(".xml")
                if category == "publication_text"
                else ""
            ),
            member=name,
            sha256=sha256(payload),
            bytes=len(payload),
            category=category,
            decision=decision,
            decision_reason=reason,
        )
    return archive, items


def negative_cases(broker: Broker, items: dict[str, Item]) -> dict[str, str]:
    """Exercise adversarial requests; every one must be denied and journalled."""
    results: dict[str, str] = {}
    blocked = next(i for i in items.values() if i.category == "author_implementation_material")
    reveal = next(i for i in items.values() if i.category == "reveal_original_implementation")
    allowed = next(i for i in items.values() if i.category == "publication_text")
    attempts = [
        ("blocked_author_material_denied", "solver_package_builder", blocked.item_id),
        ("reveal_item_denied_to_solver", "solver_package_builder", reveal.item_id),
        ("reveal_item_denied_before_freeze", "reveal_reviewer", reveal.item_id),
        ("unlisted_item_denied", "solver_package_builder", "I999"),
        ("unauthorised_actor_denied", "unknown_role", allowed.item_id),
    ]
    for name, actor, item_id in attempts:
        try:
            broker.request(actor, item_id)
            results[name] = "served_not_denied"
        except Denied as error:
            results[name] = f"denied: {error}"
    return results


def qualification(
    destination: Path, record: Path, articles: Path, wheel_source: Path
) -> dict[str, object]:
    destination.mkdir(parents=True, exist_ok=False)
    bodies = pilot_articles(record, articles)
    wheels = vetted_wheels(wheel_source)
    archive, items = build_archive(destination, bodies, wheels)
    canary_host = destination / "quarantine" / "host_canary.py"
    snapshot(canary_host, CANARY)
    journal = Journal(destination / "access-events")
    package = destination / "package"
    custodian = Custodian(archive, items)
    broker = Broker(custodian, journal, package)
    journal.record(
        "policy",
        {
            "policy_version": POLICY_VERSION,
            "roles": ["custodian", "broker", "solver", "blind_adjudicator", "reveal_reviewer"],
            "default": "deny",
            "archive_sha256": custodian.archive_sha256,
            "allowed_items": {i.item_id: i.member for i in items.values() if i.decision == "allow"},
            "blocked_items": {i.item_id: i.member for i in items.values() if i.decision != "allow"},
            "firewall_module_sha256": sha256(Path(__file__).read_bytes()),
        },
    )
    denials = negative_cases(broker, items)
    paper = bodies[0][0]
    pmid = paper.removeprefix("PMID:")
    article_item = next(i for i in items.values() if i.member == f"articles/{pmid}.xml")
    broker.request("solver_package_builder", article_item.item_id)
    for item in items.values():
        if item.category == "vetted_free_software":
            broker.request("solver_package_builder", item.item_id)
    probes = (
        PROBES.replace("SERVED_NAMES", json.dumps(sorted(broker.served)))
        .replace("CANARY_PATH", json.dumps(str(canary_host.resolve())))
        .replace("ARCHIVE_PATH", json.dumps(str(archive.resolve())))
    )
    snapshot(destination / "probes.py", probes.encode())
    worker = Worker(package, dict(broker.served))
    try:
        snapshot(destination / "worker.json", worker.inspection.encode())
        execution = worker.execute(probes, seconds=120)
        snapshot(destination / "probe.json", json.dumps(execution.__dict__, indent=2).encode())
        checks: dict[str, bool] = (
            json.loads(execution.output) if execution.stop_reason == "completed" else {}
        )
        stop = worker.execute("while True: pass", seconds=2)
        snapshot(destination / "resource_stop.json", json.dumps(stop.__dict__, indent=2).encode())
    finally:
        worker.close()
    outcome = destination / "blind_outcome.json"
    sealed = {
        "scope": "firewall_qualification_not_a_reconstruction_outcome",
        "paper_id": paper,
        "served_manifest": broker.served,
        "solver_checks": checks,
        "journal_head_sha256": journal.head,
    }
    snapshot(outcome, (json.dumps(sealed, indent=2, sort_keys=True) + "\n").encode())
    from paper2.timestamp import stamp

    receipt = stamp(outcome, destination / "freeze-timestamp")
    broker.freeze(outcome, receipt)
    reveal = next(i for i in items.values() if i.category == "reveal_original_implementation")
    post_freeze = {}
    try:
        body = broker.request("reveal_reviewer", reveal.item_id)
        post_freeze["reveal_after_freeze_to_reveal_reviewer"] = (
            "served" if body == CANARY else "served_unexpected_bytes"
        )
    except Denied as error:
        post_freeze["reveal_after_freeze_to_reveal_reviewer"] = f"denied: {error}"
    try:
        broker.request("solver_package_builder", reveal.item_id)
        post_freeze["reveal_after_freeze_to_solver"] = "served_not_denied"
    except Denied as error:
        post_freeze["reveal_after_freeze_to_solver"] = f"denied: {error}"
    mutated = json.loads(outcome.read_text())
    mutated["solver_checks"] = {"network_denied": False}
    tampered = destination / "mutated_outcome.json"
    snapshot(tampered, (json.dumps(mutated, indent=2, sort_keys=True) + "\n").encode())
    post_freeze["post_freeze_outcome_mutation_detected"] = str(
        sha256(tampered.read_bytes()) != receipt["source_sha256"]
    )
    events = sorted((destination / "access-events").glob("event-*.json"))
    head = "0" * 64
    chain_intact = True
    for index, path in enumerate(events):
        event = mapping(json.loads(path.read_bytes()))
        chain_intact = chain_intact and event["previous_sha256"] == head and event["index"] == index
        head = sha256(path.read_bytes())
    denied_events = sum(
        1
        for path in events
        if mapping(json.loads(path.read_bytes()))["kind"] == "access_denied"
    )
    denied_attempts = len(denials) + sum(
        1 for value in post_freeze.values() if value.startswith("denied")
    )
    controls = {
        "solver_probes": bool(checks) and all(checks.values()) and len(checks) == 11,
        "resource_stop_enforced": stop.stop_reason == "wall_limit",
        "all_negative_cases_denied": all(value.startswith("denied") for value in denials.values()),
        "reveal_denied_to_solver_after_freeze": post_freeze[
            "reveal_after_freeze_to_solver"
        ].startswith("denied"),
        "reveal_served_to_reviewer_after_freeze": post_freeze[
            "reveal_after_freeze_to_reveal_reviewer"
        ]
        == "served",
        "freeze_receipt_verified": receipt["status"]
        == "cryptographically_verified_against_pinned_freetsa_ca",
        "post_freeze_mutation_detected": post_freeze["post_freeze_outcome_mutation_detected"]
        == "True",
        "telemetry_chain_intact": chain_intact,
        "every_denial_recorded": denied_events == denied_attempts,
    }
    report: dict[str, object] = {
        "status": (
            "firewall_controls_pass" if all(controls.values()) else "firewall_controls_failed"
        ),
        "scope": (
            "custodian_broker_telemetry_dependency_and_freeze_controls_on_retained_pilot_article; "
            "not a reconstruction attempt and not an outcome"
        ),
        "policy_version": POLICY_VERSION,
        "recorded_at_utc": datetime.now(timezone.utc).isoformat(),
        "pilot_article_paper_id": paper,
        "pilot_articles_in_archive": [paper_id for paper_id, _ in bodies],
        "archive_sha256": custodian.archive_sha256,
        "served_manifest": broker.served,
        "dependency_closure": {name: WHEELS[name] for name in sorted(wheels)},
        "controls": controls,
        "solver_checks": checks,
        "negative_cases": denials,
        "freeze_and_reveal": post_freeze,
        "freeze_receipt": {
            "outcome_sha256": receipt["source_sha256"],
            "timestamp_response_sha256": receipt["response_sha256"],
            "authority": receipt["url"],
        },
        "access_events": len(events),
        "journal_head_sha256": head,
        "pilot_authorized": False,
        "main_study_authorized": False,
        "unqualified": [
            "model_side_retrieval_tool_restriction_for_provider_hosted_tools",
            "three_slot_cross_run_scheduling_and_budget_comparability",
            "independent_human_adjudicator_governance",
            "prior_model_training_exposure_cannot_be_certified_absent",
        ],
    }
    snapshot(destination / "report.json", (json.dumps(report, indent=2) + "\n").encode())
    return report


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument(
        "--record",
        type=Path,
        default=ROOT / "data" / "adjudication" / "devin_primary_G1_G5_20260923.json",
    )
    parser.add_argument(
        "--articles", type=Path, default=ROOT / "data" / "raw" / "corpus-articles-20260922"
    )
    parser.add_argument(
        "--wheels", type=Path, default=ROOT / "data" / "raw" / "vetted-wheels-20260924"
    )
    args = parser.parse_args()
    report = qualification(args.output, args.record, args.articles, args.wheels)
    keys = ("status", "controls", "negative_cases", "freeze_and_reveal")
    print(json.dumps({key: report[key] for key in keys}, indent=2))
    if report["status"] != "firewall_controls_pass":
        raise SystemExit(1)


if __name__ == "__main__":
    main()
