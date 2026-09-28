"""Verification-adjusted sensitivity layer (A/B/C/E review import, D excluded).

Imports the AI-assisted independent evidence review of the sealed main study into a
separate layer. The sealed blind outcome, the delegated primary adjudication and the
frozen main analysis are read but never modified. Every proposed row is validated
against the frozen repository before any adjusted quantity is computed; adjusted
quantities are written only under ``results/verification_adjusted``.

The imported rows are AI-assisted evidence review proposals for investigator
approval. They are not human-participant validation (Section D) and they do not
populate investigator sign-off fields.
"""

from __future__ import annotations

import json
import re
import shutil
import sys
from collections import Counter
from dataclasses import asdict
from pathlib import Path

from paper2.core import Row, State, fixed_triple, read_csv, sha256, wilson, write_csv
from paper2.estimation import Estimate, Stratum, stratified_policy_estimate
from paper2.main_analysis import (
    ACQUISITION_NON_SUCCESS,
    LEVELS,
    POLICY_NON_SUCCESS_STATES,
    listed,
    mapping,
    paper_disposition,
)

ROOT = Path(__file__).resolve().parents[2]
ADJ = ROOT / "data" / "adjudication"
STUDY_DIR = ADJ / "main-study-20260925"
COHORT = ADJ / "main_cohort_20260925.json"
BLIND = STUDY_DIR / "blind_outcome.json"
ADJUDICATION = STUDY_DIR / "primary_adjudication.json"
REVEAL_LEDGER = STUDY_DIR / "reveal" / "reveal_ledger.json"
IMPORT_DIR = STUDY_DIR / "verification_adjusted"
RESULTS = ROOT / "results" / "verification_adjusted"
FROZEN_ANALYSIS = ROOT / "results" / "main_analysis.json"

LAYER_ID = "verification_adjusted_sensitivity_layer"
REVIEW_ROLE = "AI-assisted independent evidence review, provisional, for investigator approval"
NOT_HUMAN_VALIDATION = (
    "Not human-participant validation (Section D); D remains pending and external."
)

REVIEW_FILES = {
    "A": "PaperII_A_30slots_independent_review_PROPOSED.csv",
    "B": "PaperII_B_90papers_investigator_verification_PROPOSED.csv",
    "C": "PaperII_C_G1G5_pilot_candidate_verification_PROPOSED.csv",
    "E": "PaperII_E_reveal_review_PROPOSED.csv",
}

LEVEL_COLUMNS = {
    "L1": LEVELS[0],
    "L2": LEVELS[1],
    "L3": LEVELS[2],
    "L4": LEVELS[3],
    "L5": LEVELS[4],
}
LEVEL_VALUES = {"yes", "no", "not_established", "not_reached", "not_assessable", "unknown"}
NON_EXECUTED_STATES = POLICY_NON_SUCCESS_STATES | ACQUISITION_NON_SUCCESS
CORRECTED_B_STATES = NON_EXECUTED_STATES | {"specifiable_runnable"}
COMPLETENESS_VALUES = {"explicit", "partial", "ambiguous", "not_reported", "not_available"}
SEGMENT_ID = re.compile(r"^S\d{4}$")
DESCRIPTOR = re.compile(r"^(?:[a-z_]+=|PMID/slot|zenodo:|github:|geo:|osf:|figshare:)")

C_LAYER_SOURCES = {
    "primary_plus_second_pass": ADJ / "devin_primary_G1_G5_20260923.json",
    "access_layer": ADJ / "accessibility_gate_failed_20260924.json",
    "pilot": ADJ / "pilot_prospective_freeze_20260925.json",
}


def evidence_root() -> Path | None:
    """Optional local evidence package used to resolve package-relative evidence_ref tokens."""
    for candidate in (
        Path("/home/ubuntu/paper2_evidence/PaperII_independent_adjudication_LIGHT"),
        Path("/home/ubuntu/paper2_evidence/PaperII_independent_adjudication_FULL"),
    ):
        if candidate.is_dir():
            return candidate
    return None


def codes(value: object) -> list[str]:
    if not isinstance(value, list):
        raise ValueError("expected a JSON list of failure codes")
    return [str(v) for v in value]


def frozen_state(paper: dict[str, object]) -> str:
    if paper["state"] == "slots_completed":
        return "slots_completed"
    eligibility = str(paper["eligibility"])
    if eligibility in POLICY_NON_SUCCESS_STATES:
        return eligibility
    return str(paper["acquisition"])


def c_paper_ids(layer: str) -> set[str]:
    data = mapping(json.loads(C_LAYER_SOURCES[layer].read_bytes()))
    if layer == "primary_plus_second_pass":
        return {str(a["paper_id"]) for a in listed(data["assessments"])}
    if layer == "pilot":
        return {str(p["paper_id"]) for p in listed(data["papers"])}
    return set(re.findall(r"PMID:\d+", C_LAYER_SOURCES[layer].read_text()))


class Importer:
    def __init__(self, source: Path) -> None:
        self.source = source
        self.errors: list[str] = []
        self.cohort = mapping(json.loads(COHORT.read_bytes()))
        self.sealed = mapping(json.loads(BLIND.read_bytes()))
        self.adjudication = mapping(json.loads(ADJUDICATION.read_bytes()))
        self.ledger = mapping(json.loads(REVEAL_LEDGER.read_bytes()))
        if self.adjudication["blind_outcome_sha256"] != sha256(BLIND.read_bytes()):
            raise ValueError("adjudication does not reference the sealed blind outcome")
        self.papers = listed(self.sealed["papers"])
        self.runs = {
            (str(r["paper_id"]), int(str(r["slot"]))): r for r in listed(self.adjudication["runs"])
        }
        self.endpoints = {str(p["paper_id"]): p for p in listed(self.adjudication["papers"])}
        self.segments = set(
            re.findall(r"S\d{4}", (ADJ / "second_pass_notes_20260924.json").read_text())
        )
        self.evidence = evidence_root()
        self.reference_rows: list[Row] = []
        self.hashes: dict[str, str] = {}

    def fail(self, message: str) -> None:
        self.errors.append(message)

    # ----------------------------------------------------------------- import
    def copy_reviews(self) -> dict[str, list[Row]]:
        IMPORT_DIR.mkdir(parents=True, exist_ok=True)
        reviews: dict[str, list[Row]] = {}
        for section, name in REVIEW_FILES.items():
            src = self.source / name
            dst = IMPORT_DIR / name
            data = src.read_bytes()
            if dst.exists() and dst.read_bytes() != data:
                raise ValueError(f"{dst} exists with different content; refusing to overwrite")
            if not dst.exists():
                shutil.copyfile(src, dst)
            self.hashes[name] = sha256(data)
            reviews[section] = read_csv(dst)
        return reviews

    # -------------------------------------------------------- evidence_ref
    def resolve(self, section: str, row: Row, token: str) -> str:
        token = token.strip().strip("/")
        if not token:
            return "empty"
        paper = row["paper_id"].removeprefix("PMID:")
        repo_candidates = [
            ROOT / token,
            STUDY_DIR / token,
            STUDY_DIR / "reveal" / token,
            STUDY_DIR / "specifications" / Path(token).name,
        ]
        if any(c.is_file() for c in repo_candidates):
            return "repository_file"
        if SEGMENT_ID.match(token):
            return "segment_id" if token in self.segments else "unresolved"
        if self.evidence is not None:
            slot = row.get("slot", "")
            candidates = [self.evidence / token]
            if section == "A" and slot:
                candidates.append(
                    self.evidence
                    / "05_MAIN_RUN_EVIDENCE_10x3"
                    / f"PMID_{paper}"
                    / f"slot-{slot}"
                    / token
                )
            if section == "B":
                state = row["recorded_state"]
                candidates.append(
                    self.evidence / "08_NONEXECUTED_90_EVIDENCE" / state / paper / token
                )
            if section == "C":
                candidates.append(
                    self.evidence / "09_G1G5_PILOT_AND_CANDIDATE_HISTORY" / "records" / token
                )
                candidates.append(ADJ / Path(token).name)
            if section == "E":
                candidates.append(self.evidence / "11_REVEAL_EVIDENCE" / "records" / token)
            if any(c.exists() for c in candidates):
                return "evidence_package_file"
        if DESCRIPTOR.match(token):
            return "descriptor"
        if self.evidence is None:
            return "evidence_package_not_mounted"
        return "unresolved"

    def check_refs(self, section: str, rows: list[Row], key: str) -> None:
        for row in rows:
            for token in row["evidence_ref"].split(";"):
                if not token.strip():
                    continue
                kind = self.resolve(section, row, token)
                self.reference_rows.append(
                    {
                        "section": section,
                        "paper_id": row["paper_id"],
                        "key": row.get(key, ""),
                        "evidence_ref": token.strip(),
                        "resolution": kind,
                    }
                )
                if kind == "unresolved":
                    self.fail(f"{section} {row['paper_id']} unresolved evidence_ref {token!r}")

    # ------------------------------------------------------------ section A
    def validate_a(self, rows: list[Row]) -> None:
        keys = [(r["paper_id"], int(r["slot"])) for r in rows]
        if sorted(keys) != sorted(self.runs):
            self.fail("A: paper_id/slot set differs from the frozen 30 sealed slots")
        for row in rows:
            key = (row["paper_id"], int(row["slot"]))
            frozen = self.runs.get(key)
            if frozen is None:
                continue
            levels = {c: row[c] for c in LEVEL_COLUMNS}
            if any(v not in LEVEL_VALUES for v in levels.values()):
                self.fail(f"A {key}: invalid level value {levels}")
            if row["run_success"] not in {"yes", "no", "unknown", "not_assessable"}:
                self.fail(f"A {key}: invalid run_success {row['run_success']!r}")
            necessary = (levels["L1"], levels["L2"], levels["L3"], levels["L4"])
            expected = (
                "yes"
                if all(v == "yes" for v in necessary)
                else "no"
                if any(v in {"no", "not_reached", "not_established"} for v in necessary)
                else "unknown"
            )
            if row["run_success"] != expected:
                self.fail(
                    f"A {key}: run_success {row['run_success']} inconsistent with levels {levels}"
                )
            if row["agree_with_mechanical"] == "yes":
                for short, full in LEVEL_COLUMNS.items():
                    if short == "L5":
                        continue
                    if str(frozen[full]) != row[short]:
                        self.fail(
                            f"A {key}: agree=yes but {short} differs from frozen {frozen[full]}"
                        )
                if str(frozen["run_success"]) != row["run_success"]:
                    self.fail(f"A {key}: agree=yes but run_success differs from frozen")
            if row["adjudicator_id"] != "AI_independent_evidence_review_FOR_INVESTIGATOR_APPROVAL":
                self.fail(f"A {key}: unexpected adjudicator_id {row['adjudicator_id']!r}")
        self.check_refs("A", rows, "slot")

    # ------------------------------------------------------------ section B
    def validate_b(self, rows: list[Row]) -> None:
        non_executed = {
            str(p["paper_id"]): frozen_state(p)
            for p in self.papers
            if p["state"] != "slots_completed"
        }
        ids = [r["paper_id"] for r in rows]
        if sorted(ids) != sorted(non_executed):
            self.fail("B: paper_id set differs from the frozen 90 non-executed papers")
        for row in rows:
            frozen = non_executed.get(row["paper_id"])
            if frozen is None:
                continue
            if row["recorded_state"] != frozen:
                self.fail(
                    f"B {row['paper_id']}: recorded_state {row['recorded_state']} "
                    f"!= frozen {frozen}"
                )
            if row["verification"] not in {"confirmed", "state_incorrect", "unknown"}:
                self.fail(f"B {row['paper_id']}: invalid verification {row['verification']!r}")
            if row["verification"] == "state_incorrect":
                if (
                    row["corrected_state"] not in CORRECTED_B_STATES
                    or row["corrected_state"] == frozen
                ):
                    self.fail(
                        f"B {row['paper_id']}: invalid corrected_state {row['corrected_state']!r}"
                    )
            elif row["corrected_state"]:
                self.fail(f"B {row['paper_id']}: corrected_state given without state_incorrect")
        self.check_refs("B", rows, "recorded_state")

    # ------------------------------------------------------------ section C
    def validate_c(self, rows: list[Row]) -> None:
        for layer in C_LAYER_SOURCES:
            expected = c_paper_ids(layer)
            got = {r["paper_id"] for r in rows if r["layer"] == layer}
            if got != expected:
                self.fail(
                    f"C {layer}: paper set differs from the frozen record "
                    f"({sorted(got ^ expected)})"
                )
        for row in rows:
            if row["layer"] not in C_LAYER_SOURCES:
                self.fail(f"C {row['paper_id']}: unknown layer {row['layer']!r}")
            if row["verification"] != "confirmed" or row["corrected_state"]:
                self.fail(f"C {row['paper_id']}: post-outcome C changes are not permitted")
        self.check_refs("C", rows, "layer")

    # ------------------------------------------------------------ section E
    def validate_e(self, rows: list[Row]) -> None:
        ledger = {str(p["paper_id"]): p for p in listed(self.ledger["papers"])}
        pairs = Counter((r["paper_id"], r["dimension"]) for r in rows)
        if any(n > 1 for n in pairs.values()):
            self.fail("E: duplicated paper/dimension rows")
        if {r["paper_id"] for r in rows} != set(ledger):
            self.fail("E: paper set differs from the reveal ledger")
        for row in rows:
            paper = ledger.get(row["paper_id"])
            if paper is None:
                continue
            dims = mapping(paper["dimensions"])
            if dims:
                if row["dimension"] not in dims:
                    self.fail(f"E {row['paper_id']}: unknown dimension {row['dimension']}")
                    continue
                recorded = str(mapping(dims[row["dimension"]])["completeness"])
            else:
                recorded = "not_available"
            if row["recorded_completeness"] != recorded:
                self.fail(
                    f"E {row['paper_id']}/{row['dimension']}: recorded "
                    f"{row['recorded_completeness']} != ledger {recorded}"
                )
            if row["agree"] == "yes" and row["corrected_completeness"] not in {"", recorded}:
                self.fail(f"E {row['paper_id']}/{row['dimension']}: agree=yes with a correction")
            if row["agree"] == "no" and row["corrected_completeness"] not in COMPLETENESS_VALUES:
                self.fail(f"E {row['paper_id']}/{row['dimension']}: invalid corrected completeness")
            if row["agree"] == "no" and not dims:
                self.fail(f"E {row['paper_id']}: cannot correct a paper without a code route")
        dims_expected = sum(max(len(mapping(p["dimensions"])), 13) for p in ledger.values())
        if len(rows) != dims_expected:
            self.fail(f"E: {len(rows)} rows, expected {dims_expected}")
        self.check_refs("E", rows, "dimension")

    # ------------------------------------------------------------ analysis
    def adjusted_a(self, rows: list[Row]) -> tuple[list[Row], dict[str, object]]:
        slot_rows: list[Row] = []
        by_paper: dict[str, dict[int, Row]] = {}
        for row in sorted(rows, key=lambda r: (r["paper_id"], int(r["slot"]))):
            key = (row["paper_id"], int(row["slot"]))
            frozen = self.runs[key]
            out: Row = {
                "paper_id": row["paper_id"],
                "slot": row["slot"],
                "frozen_run_success": str(frozen["run_success"]),
                "adjusted_run_success": row["run_success"],
                "changed": "yes" if row["agree_with_mechanical"] == "no" else "no",
                "adjusted_failure_codes": row["failure_codes"],
                "frozen_failure_codes": ";".join(str(c) for c in codes(frozen["failure_codes"])),
                "reviewer_id": row["adjudicator_id"],
                "reviewed_at_utc": row["adjudicated_at_utc"],
                "investigator_signoff": "",
                "layer": LAYER_ID,
            }
            for short, full in LEVEL_COLUMNS.items():
                out[f"frozen_{short}"] = str(frozen[full])
                out[f"adjusted_{short}"] = row[short]
            slot_rows.append(out)
            by_paper.setdefault(row["paper_id"], {})[int(row["slot"])] = row

        def slot_state(value: str) -> State:
            if value in {"yes", "no", "unknown", "not_assessable"}:
                return value  # type: ignore[return-value]
            raise ValueError(value)

        paper_rows: list[Row] = []
        adjusted_endpoints: dict[str, str] = {}
        for pid, slots in sorted(by_paper.items()):
            states = [slot_state(slots[i]["run_success"]) for i in (1, 2, 3)]
            frozen_paper = self.endpoints[pid]
            frozen_k = int(str(frozen_paper["clean_successes"]))
            k = states.count("yes")
            adjusted_endpoints[pid] = fixed_triple(states)
            paper_rows.append(
                {
                    "paper_id": pid,
                    "frozen_successful_slots": f"{frozen_k}/3",
                    "adjusted_successful_slots": f"{k}/3",
                    "frozen_majority": str(frozen_paper["primary_endpoint_majority"]),
                    "adjusted_majority": adjusted_endpoints[pid],
                    "adjusted_strict": fixed_triple(states, 3),
                    "adjusted_permissive": fixed_triple(states, 1),
                    "changed": "yes" if frozen_k != k else "no",
                }
            )
        frozen_dist = Counter(f"{self.endpoints[p]['clean_successes']}/3" for p in by_paper)
        adjusted_dist = Counter(r["adjusted_successful_slots"] for r in paper_rows)
        distribution = [
            {
                "successful_slots": f"{k}/3",
                "frozen_papers": frozen_dist.get(f"{k}/3", 0),
                "adjusted_papers": adjusted_dist.get(f"{k}/3", 0),
            }
            for k in range(4)
        ]
        levels = {
            level: {
                "frozen": dict(Counter(r[f"frozen_{short}"] for r in slot_rows)),
                "adjusted": dict(Counter(r[f"adjusted_{short}"] for r in slot_rows)),
            }
            for short, level in LEVEL_COLUMNS.items()
        }
        codes_frozen: Counter[str] = Counter()
        codes_adjusted: Counter[str] = Counter()
        for r in slot_rows:
            codes_frozen.update(c for c in r["frozen_failure_codes"].split(";") if c)
            codes_adjusted.update(c for c in r["adjusted_failure_codes"].split(";") if c)
        summary = {
            "slots_reviewed": len(slot_rows),
            "slots_mechanical_supported": sum(r["changed"] == "no" for r in slot_rows),
            "slots_proposed_changed": sum(r["changed"] == "yes" for r in slot_rows),
            "papers": paper_rows,
            "distribution": distribution,
            "adjusted_endpoints": adjusted_endpoints,
            "levels": levels,
            "failure_codes_runs": {
                "frozen": dict(sorted(codes_frozen.items())),
                "adjusted": dict(sorted(codes_adjusted.items())),
            },
        }
        return slot_rows, summary

    def adjusted_b(self, rows: list[Row]) -> tuple[list[Row], dict[str, object]]:
        by_id = {r["paper_id"]: r for r in rows}
        funnel_rows: list[Row] = []
        adjusted_state: dict[str, str] = {}
        for paper in self.papers:
            pid = str(paper["paper_id"])
            frozen = frozen_state(paper)
            review = by_id.get(pid)
            if review is None:
                adjusted, envelope, verification = frozen, frozen, "attempted_not_in_B_scope"
            elif review["verification"] == "confirmed":
                adjusted, envelope, verification = frozen, frozen, "confirmed"
            elif review["verification"] == "state_incorrect":
                adjusted, envelope, verification = (
                    review["corrected_state"],
                    "unresolved",
                    "corrected",
                )
            else:
                adjusted, envelope, verification = "unresolved", "unresolved", "unknown"
            adjusted_state[pid] = adjusted
            funnel_rows.append(
                {
                    "paper_id": pid,
                    "sampling_stratum": str(paper["sampling_stratum"]),
                    "frozen_state": frozen,
                    "verification": verification,
                    "adjusted_state": adjusted,
                    "envelope_state": envelope,
                    "reviewer_id": review["reviewer_id"] if review else "",
                    "reviewed_at_utc": review["verified_at_utc"] if review else "",
                    "investigator_signoff": "",
                }
            )
        leak = [r for r in rows if r["recorded_state"] == "all_inputs_leak_target"]
        detector = {
            "classifications": len(leak),
            "confirmed": sum(r["verification"] == "confirmed" for r in leak),
            "corrected": sum(r["verification"] == "state_incorrect" for r in leak),
            "unresolved": sum(r["verification"] == "unknown" for r in leak),
            "corrected_to": dict(
                Counter(
                    r["corrected_state"] for r in leak if r["verification"] == "state_incorrect"
                )
            ),
            "methodological_finding": (
                "whole-token matches inside compressed or binary deposited content and raw "
                "input matrices produced false-positive target-leakage flags"
            ),
        }
        summary = {
            "papers_reviewed": len(rows),
            "confirmed": sum(r["verification"] == "confirmed" for r in rows),
            "state_incorrect": sum(r["verification"] == "state_incorrect" for r in rows),
            "unknown": sum(r["verification"] == "unknown" for r in rows),
            "corrections_by_frozen_state": dict(
                Counter(r["recorded_state"] for r in rows if r["verification"] != "confirmed")
            ),
            "funnel": {
                "frozen": dict(sorted(Counter(r["frozen_state"] for r in funnel_rows).items())),
                "adjusted": dict(sorted(Counter(r["adjusted_state"] for r in funnel_rows).items())),
                "envelope": dict(sorted(Counter(r["envelope_state"] for r in funnel_rows).items())),
            },
            "leakage_detector_audit": detector,
            "reruns_performed": 0,
            "note": (
                "corrections are attrition interpretation only; no reclassified paper was executed"
            ),
        }
        return funnel_rows, summary

    def adjusted_e(self, rows: list[Row]) -> tuple[list[Row], dict[str, object]]:
        out: list[Row] = []
        for row in sorted(rows, key=lambda r: (r["paper_id"], r["dimension"])):
            corrected = (
                row["corrected_completeness"]
                if row["agree"] == "no"
                else row["recorded_completeness"]
            )
            out.append(
                {
                    "paper_id": row["paper_id"],
                    "dimension": row["dimension"],
                    "frozen_completeness": row["recorded_completeness"],
                    "adjusted_completeness": corrected,
                    "changed": "yes" if row["agree"] == "no" else "no",
                    "reviewer_id": row["reviewer_id"],
                    "reviewed_at_utc": row["reviewed_at_utc"],
                    "investigator_signoff": "",
                }
            )
        described = [r for r in out if r["frozen_completeness"] != "not_available"]
        summary = {
            "papers": len({r["paper_id"] for r in out}),
            "papers_with_code_route": len({r["paper_id"] for r in described}),
            "papers_without_code_route": len({r["paper_id"] for r in out})
            - len({r["paper_id"] for r in described}),
            "dimension_rows": len(out),
            "proposed_corrections": sum(r["changed"] == "yes" for r in out),
            "completeness_described_papers": {
                "frozen": dict(
                    sorted(Counter(r["frozen_completeness"] for r in described).items())
                ),
                "adjusted": dict(
                    sorted(Counter(r["adjusted_completeness"] for r in described).items())
                ),
            },
            "original_code_executed": False,
            "blind_scores_revised_from_reveal": False,
        }
        return out, summary

    def estimands(
        self, a: dict[str, object], b: dict[str, object], funnel: list[Row]
    ) -> dict[str, object]:
        n_cohort = len(self.papers)
        attempted = [p for p in self.papers if p["state"] == "slots_completed"]
        endpoints_adj = mapping(a["adjusted_endpoints"])
        frozen_success = sum(
            self.endpoints[str(p["paper_id"])]["primary_endpoint_majority"] == "success"
            for p in attempted
        )
        adjusted_success = sum(endpoints_adj[str(p["paper_id"])] == "success" for p in attempted)

        def ci(k: int, n: int) -> dict[str, object]:
            lo, hi = wilson(k, n)
            return {
                "numerator": k,
                "denominator": n,
                "proportion": k / n,
                "wilson_95_lower": lo,
                "wilson_95_upper": hi,
            }

        strata = listed(self.cohort["strata"])
        eligible = int(str(self.cohort["eligible_population"]))
        state_by_id = {r["paper_id"]: r for r in funnel}

        def stratum(record: dict[str, object], column: str) -> Stratum:
            name = record["sampling_stratum"]
            members = [p for p in self.papers if p["sampling_stratum"] == name]
            successes = unresolved = 0
            for p in members:
                pid = str(p["paper_id"])
                if p["state"] == "slots_completed":
                    if endpoints_adj[pid] == "success":
                        successes += 1
                    elif endpoints_adj[pid] == "unknown":
                        unresolved += 1
                elif state_by_id[pid][column] in {"unresolved", "specifiable_runnable"}:
                    unresolved += 1
            return Stratum(int(str(record["N_h"])), len(members), successes, unresolved)

        def row(label: str, e: Estimate) -> dict[str, object]:
            return {"analysis": label, **asdict(e)}

        adjusted_policy = stratified_policy_estimate(
            [stratum(s, "adjusted_state") for s in strata], eligible, 0, 0
        )
        envelope_policy = stratified_policy_estimate(
            [stratum(s, "envelope_state") for s in strata], eligible, 0, 0
        )
        return {
            "attempt_reachability": {
                **ci(len(attempted), n_cohort),
                "frame": "identical under frozen and adjusted layers",
            },
            "conditional_success_among_attempted": {
                "frozen_mechanical": ci(frozen_success, len(attempted)),
                "verification_adjusted": ci(adjusted_success, len(attempted)),
                "label": "secondary; conditional on attempt; unweighted binomial reference",
            },
            "observed_end_to_end_yield_100": {
                "frozen_mechanical": ci(frozen_success, n_cohort),
                "verification_adjusted": ci(adjusted_success, n_cohort),
                "label": (
                    "observed yield under the frozen workflow, "
                    "not an intrinsic reproducibility probability"
                ),
            },
            "policy_estimate_eligible_population": {
                "verification_adjusted_states": row(
                    "verification_adjusted_policy", adjusted_policy
                ),
                "verification_adjusted_unresolved_envelope": row(
                    "verification_adjusted_envelope", envelope_policy
                ),
            },
            "non_executed_papers_are_reconstruction_failures": False,
        }

    # ---------------------------------------------------------------- driver
    def run(self) -> dict[str, object]:
        reviews = self.copy_reviews()
        self.validate_a(reviews["A"])
        self.validate_b(reviews["B"])
        self.validate_c(reviews["C"])
        self.validate_e(reviews["E"])
        for section, rows in reviews.items():
            for row in rows:
                if any("participant" in k for k in row) or "human_validation" in row:
                    self.fail(f"{section}: Section D participant field detected")
        if self.errors:
            raise ValueError("verification import failed:\n" + "\n".join(self.errors))

        RESULTS.mkdir(parents=True, exist_ok=True)
        slot_rows, a_summary = self.adjusted_a(reviews["A"])
        funnel_rows, b_summary = self.adjusted_b(reviews["B"])
        e_rows, e_summary = self.adjusted_e(reviews["E"])
        c_summary = {
            "records_reviewed": len(reviews["C"]),
            "confirmed": sum(r["verification"] == "confirmed" for r in reviews["C"]),
            "by_layer": dict(sorted(Counter(r["layer"] for r in reviews["C"]).items())),
            "post_outcome_corrections": 0,
            "pilot_in_main_denominator": False,
            "accessibility_gate_failed_role": "auxiliary access-barrier layer",
        }
        estimands = self.estimands(a_summary, b_summary, funnel_rows)
        frozen_analysis = mapping(json.loads(FROZEN_ANALYSIS.read_bytes()))
        frozen_dispositions = {
            str(p["paper_id"]): paper_disposition(p, self.endpoints) for p in self.papers
        }
        if dict(Counter(frozen_dispositions.values())) != mapping(
            frozen_analysis["paper_dispositions"]
        ):
            raise ValueError("frozen dispositions no longer match results/main_analysis.json")

        write_csv(RESULTS / "adjusted_A_slots.csv", slot_rows, list(slot_rows[0]))
        papers = listed(a_summary["papers"])
        write_csv(RESULTS / "adjusted_A_papers.csv", papers, list(papers[0]))
        dist = listed(a_summary["distribution"])
        write_csv(RESULTS / "adjusted_A_distribution.csv", dist, list(dist[0]))
        write_csv(RESULTS / "adjusted_B_funnel_papers.csv", funnel_rows, list(funnel_rows[0]))
        funnel = mapping(b_summary["funnel"])
        states = sorted(
            set().union(*(mapping(funnel[k]) for k in ("frozen", "adjusted", "envelope")))
        )
        funnel_table = [
            {
                "state": s,
                "frozen": mapping(funnel["frozen"]).get(s, 0),
                "verification_adjusted": mapping(funnel["adjusted"]).get(s, 0),
                "unresolved_envelope": mapping(funnel["envelope"]).get(s, 0),
            }
            for s in states
        ]
        write_csv(RESULTS / "adjusted_B_funnel.csv", funnel_table, list(funnel_table[0]))
        write_csv(RESULTS / "adjusted_E_reveal_ledger.csv", e_rows, list(e_rows[0]))
        write_csv(
            RESULTS / "evidence_reference_resolution.csv",
            self.reference_rows,
            list(self.reference_rows[0]),
        )
        policy = mapping(estimands["policy_estimate_eligible_population"])
        policy_rows = [mapping(v) for v in policy.values()]
        write_csv(RESULTS / "adjusted_policy_estimates.csv", policy_rows, list(policy_rows[0]))

        audit = {
            "layer": LAYER_ID,
            "review_role": REVIEW_ROLE,
            "human_participant_validation": NOT_HUMAN_VALIDATION,
            "investigator_signoff": (
                "not recorded; fields left empty pending explicit investigator approval"
            ),
            "source_files_sha256": self.hashes,
            "frozen_inputs_sha256": {
                "blind_outcome": sha256(BLIND.read_bytes()),
                "primary_adjudication": sha256(ADJUDICATION.read_bytes()),
                "reveal_ledger": sha256(REVEAL_LEDGER.read_bytes()),
                "main_cohort": sha256(COHORT.read_bytes()),
            },
            "frozen_inputs_modified": False,
            "rows": {k: len(v) for k, v in reviews.items()},
            "evidence_references": {
                "total": len(self.reference_rows),
                "by_resolution": dict(
                    sorted(Counter(r["resolution"] for r in self.reference_rows).items())
                ),
                "unresolved": sum(r["resolution"] == "unresolved" for r in self.reference_rows),
                "evidence_package_root": str(self.evidence) if self.evidence else None,
                "note": (
                    "package-relative references were resolved against the local private evidence "
                    "package when mounted; otherwise they are recorded as "
                    "evidence_package_not_mounted"
                ),
            },
            "validation_errors": self.errors,
            "A": a_summary,
            "B": b_summary,
            "C": c_summary,
            "E": e_summary,
            "estimands": estimands,
            "prohibited_actions_not_performed": [
                "blind slot rerun",
                "original code execution",
                "sealed outcome modification",
                "human validation D",
                "reveal-informed rescoring of blind runs",
            ],
        }
        (RESULTS / "verification_import_audit.json").write_text(
            json.dumps(audit, indent=2, default=str) + "\n"
        )
        return audit


FROZEN_STATUS = "frozen_mechanical_delegated_adjudication"
ADJUSTED_STATUS = "verification_adjusted_ai_review_provisional_not_investigator_signed"


def verification_values(audit: dict[str, object], analysis: dict[str, object]) -> list[Row]:
    """Machine-readable manuscript numerators/denominators, frozen and adjusted side by side."""
    est = mapping(audit["estimands"])
    a = mapping(audit["A"])
    b = mapping(audit["B"])
    e = mapping(audit["E"])
    c = mapping(audit["C"])
    detector = mapping(b["leakage_detector_audit"])
    funnel = mapping(b["funnel"])
    reach = mapping(est["attempt_reachability"])
    cond = mapping(est["conditional_success_among_attempted"])
    yield_ = mapping(est["observed_end_to_end_yield_100"])
    denominators = mapping(analysis["denominators"])
    entries: list[tuple[str, object, str, str]] = [
        (
            "main_eligible_population",
            denominators["route_eligible_population_E"],
            FROZEN_STATUS,
            "denominator",
        ),
        ("main_cohort_selected", denominators["sampled_main_cohort"], FROZEN_STATUS, "denominator"),
        (
            "main_attempted_papers",
            reach["numerator"],
            FROZEN_STATUS,
            "attempt_reachability_numerator",
        ),
        (
            "main_non_executed_papers",
            denominators["not_attempted_papers"],
            FROZEN_STATUS,
            "pre_reconstruction_non_execution",
        ),
        ("main_blind_slots", denominators["runs"], FROZEN_STATUS, "runs"),
        (
            "frozen_majority_successes_among_attempted",
            mapping(cond["frozen_mechanical"])["numerator"],
            FROZEN_STATUS,
            "conditional_numerator",
        ),
        (
            "adjusted_majority_successes_among_attempted",
            mapping(cond["verification_adjusted"])["numerator"],
            ADJUSTED_STATUS,
            "conditional_numerator",
        ),
        (
            "frozen_end_to_end_successes_100",
            mapping(yield_["frozen_mechanical"])["numerator"],
            FROZEN_STATUS,
            "observed_yield_numerator",
        ),
        (
            "adjusted_end_to_end_successes_100",
            mapping(yield_["verification_adjusted"])["numerator"],
            ADJUSTED_STATUS,
            "observed_yield_numerator",
        ),
        ("A_slots_reviewed", a["slots_reviewed"], ADJUSTED_STATUS, "review_scope"),
        (
            "A_slots_mechanical_supported",
            a["slots_mechanical_supported"],
            ADJUSTED_STATUS,
            "review_result",
        ),
        ("A_slots_proposed_changed", a["slots_proposed_changed"], ADJUSTED_STATUS, "review_result"),
        ("B_papers_reviewed", b["papers_reviewed"], ADJUSTED_STATUS, "review_scope"),
        ("B_confirmed", b["confirmed"], ADJUSTED_STATUS, "review_result"),
        ("B_state_incorrect", b["state_incorrect"], ADJUSTED_STATUS, "review_result"),
        ("B_unknown", b["unknown"], ADJUSTED_STATUS, "review_result"),
        (
            "B_adjusted_specifiable_runnable_not_executed",
            mapping(funnel["adjusted"]).get("specifiable_runnable", 0),
            ADJUSTED_STATUS,
            "attrition_interpretation_only",
        ),
        ("leakage_flags_total", detector["classifications"], FROZEN_STATUS, "detector_audit"),
        ("leakage_flags_confirmed", detector["confirmed"], ADJUSTED_STATUS, "detector_audit"),
        ("leakage_flags_corrected", detector["corrected"], ADJUSTED_STATUS, "detector_audit"),
        ("leakage_flags_unresolved", detector["unresolved"], ADJUSTED_STATUS, "detector_audit"),
        ("C_records_reviewed", c["records_reviewed"], ADJUSTED_STATUS, "historical_pilot_layer"),
        ("C_confirmed", c["confirmed"], ADJUSTED_STATUS, "historical_pilot_layer"),
        ("E_dimension_rows", e["dimension_rows"], ADJUSTED_STATUS, "descriptive_reveal"),
        (
            "E_proposed_corrections",
            e["proposed_corrections"],
            ADJUSTED_STATUS,
            "descriptive_reveal",
        ),
        (
            "E_papers_with_code_route",
            e["papers_with_code_route"],
            FROZEN_STATUS,
            "descriptive_reveal",
        ),
        (
            "evidence_refs_total",
            mapping(audit["evidence_references"])["total"],
            ADJUSTED_STATUS,
            "import_audit",
        ),
        (
            "evidence_refs_unresolved",
            mapping(audit["evidence_references"])["unresolved"],
            ADJUSTED_STATUS,
            "import_audit",
        ),
        (
            "D_human_validation_cases_completed",
            0,
            "pending_external_not_performed",
            "section_D_pending",
        ),
    ]
    return [
        {
            "claim_id": key,
            "value": str(value),
            "unit": "count",
            "status": status,
            "source_id": "data/adjudication/main-study-20260925/",
            "source_sha256": str(mapping(audit["frozen_inputs_sha256"])["blind_outcome"]),
            "analysis": "src/paper2/verification_layer.py:Importer.run",
            "derived_file": "results/verification_adjusted/verification_import_audit.json",
            "claim_scope": scope,
        }
        for key, value, status, scope in entries
    ]


def pct(item: dict[str, object]) -> str:
    return (
        f"{item['numerator']}/{item['denominator']} = {100 * float(str(item['proportion'])):.1f}% "
        f"(Wilson 95% CI {100 * float(str(item['wilson_95_lower'])):.1f}–"
        f"{100 * float(str(item['wilson_95_upper'])):.1f}%)"
    )


def without_d_status(audit: dict[str, object], analysis: dict[str, object]) -> None:
    est = mapping(audit["estimands"])
    cond = mapping(est["conditional_success_among_attempted"])
    yield_ = mapping(est["observed_end_to_end_yield_100"])
    b = mapping(audit["B"])
    adjusted = mapping(mapping(b["funnel"])["adjusted"])
    detector = mapping(b["leakage_detector_audit"])
    frozen_sha = mapping(audit["frozen_inputs_sha256"])
    barriers = sum(
        int(str(v))
        for k, v in adjusted.items()
        if k not in {"slots_completed", "specifiable_runnable", "unresolved"}
    )
    lines = [
        "# WITHOUT_D_STATUS — complete without D versus D-dependent",
        "",
        "Status: NOT SUBMISSION READY. Human validation (Section D) has not been performed.",
        f"The A/B/C/E rows are an {REVIEW_ROLE}. They are not human-participant validation",
        "and carry no investigator sign-off.",
        "",
        "## Complete without D (regenerated from `results/`)",
        "",
        f"- Attempt reachability: {pct(mapping(est['attempt_reachability']))}.",
        "- Conditional majority success among attempted papers — frozen mechanical: "
        f"{pct(mapping(cond['frozen_mechanical']))}; verification-adjusted: "
        f"{pct(mapping(cond['verification_adjusted']))}.",
        "- Observed end-to-end yield under the frozen workflow — frozen: "
        f"{pct(mapping(yield_['frozen_mechanical']))}; verification-adjusted: "
        f"{pct(mapping(yield_['verification_adjusted']))}.",
        "- 100-paper funnel (frozen / adjusted / unresolved envelope): "
        "`results/verification_adjusted/adjusted_B_funnel.csv`. Adjusted non-executed states: "
        f"{barriers} verified barriers, {adjusted.get('specifiable_runnable', 0)} reclassified "
        f"`specifiable_runnable` (not executed; attrition interpretation only), "
        f"{adjusted.get('unresolved', 0)} unknown.",
        f"- Leakage detector audit: {detector['confirmed']} confirmed, {detector['corrected']} "
        f"corrected, {detector['unresolved']} unresolved of {detector['classifications']} "
        "`all_inputs_leak_target` flags.",
        "- Run-level L1–L5 (frozen and adjusted), failure taxonomy and stop reasons: "
        "`results/main_run_levels.csv`, `results/main_failure_codes.csv`, "
        "`results/verification_adjusted/verification_import_audit.json`.",
        "- Descriptive reveal completeness (frozen and conservative adjusted labels): "
        "`results/verification_adjusted/adjusted_E_reveal_ledger.csv`.",
        f"- Sealed blind outcome unchanged: SHA-256 `{frozen_sha['blind_outcome']}`.",
        f"- Frozen primary adjudication unchanged: SHA-256 `{frozen_sha['primary_adjudication']}`.",
        "",
        "## Not performed (by instruction)",
        "",
        "No blind slot rerun; no execution of reclassified papers; no original-code execution;",
        "no reveal-informed rescoring; no Paper III analyses; no investigator sign-off recorded.",
        "",
        "## D-dependent (pending, external)",
        "",
        "- Human-validation Results text, Table 9 and figure: empty/pending "
        "(`D_human_validation_cases_completed = 0`).",
        "- Any claim about human reconstructability or agent-versus-human comparison.",
        "- Conversion of the AI-assisted A/B/C/E proposals into signed investigator adjudication.",
        "- Institutional ethics determination, validator recruitment and participation records.",
        "- Submission readiness (`make submission-check` remains non-zero).",
        "",
    ]
    (ROOT / "WITHOUT_D_STATUS.md").write_text("\n".join(lines))


def main(argv: list[str] | None = None) -> None:
    args = argv if argv is not None else sys.argv[1:]
    source = Path(args[0]) if args else IMPORT_DIR
    audit = Importer(source).run()
    est = mapping(audit["estimands"])
    cond = mapping(mapping(est["conditional_success_among_attempted"])["verification_adjusted"])
    print(
        f"imported {LAYER_ID}: adjusted conditional {cond['numerator']}/{cond['denominator']}, "
        f"unresolved refs {mapping(audit['evidence_references'])['unresolved']}"
    )


if __name__ == "__main__":
    main()
