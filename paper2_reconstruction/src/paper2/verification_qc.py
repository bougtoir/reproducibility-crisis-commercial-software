"""QC for the verification-adjusted layer: frozen preservation, arithmetic, consistency."""

import hashlib
import json
import re
from pathlib import Path

from docx import Document

from paper2.core import read_csv, sha256
from paper2.documents import REFERENCES, STAGES
from paper2.input_state import ACCESS_QUESTIONS, KEY_STATEMENT, STAGE_TABLE
from paper2.main_analysis import mapping
from paper2.verification_layer import (
    ADJUDICATION,
    BLIND,
    IMPORT_DIR,
    RESULTS,
    REVEAL_LEDGER,
    ROOT,
)

FROZEN_BLIND_SHA = "350a0e7dcedbb74ccad2472c642318c87cc7d272f1b6c88186a1419dc858ea5a"
FROZEN_ADJUDICATION_SHA = "6fd42b8964e055606cbd24ebddb44f42fa09147fa80bfa2d5bcaa161766c5b28"
FROZEN_ABCE_SHA = "f6b374b74c48c297e4c68dcdd780a4d2c1f162e8dcfc058111c524408961f78a"
D_TERMS = ("validator_id", "consent_record", "human_results")
STAGE_SEQUENCE = "ACCESS → RECONSTRUCT → EXECUTE → REPRODUCE → ROBUST"
UNCALIBRATED = (
    "public datasets are not reproducible",
    "all public databases change",
    "quantified dataset drift",
    "measured dataset drift",
)
RERUN_TERMS = ("reran", "re-ran", "rerun slot", "executed the original code")


def tree_sha(directory: Path) -> str:
    digest = hashlib.sha256()
    for path in sorted(directory.rglob("*")):
        if path.is_file():
            digest.update(path.relative_to(directory).as_posix().encode())
            digest.update(bytes.fromhex(sha256(path.read_bytes())))
    return digest.hexdigest()


def docx_text(path: Path) -> str:
    document = Document(str(path))
    parts = [p.text for p in document.paragraphs]
    for table in document.tables:
        for row in table.rows:
            parts.extend(cell.text for cell in row.cells)
    return "\n".join(parts)


def run() -> list[tuple[str, bool, str]]:
    audit = json.loads((RESULTS / "verification_import_audit.json").read_text())
    frozen = json.loads((ROOT / "results/main_analysis.json").read_text())
    values = {r["claim_id"]: r for r in read_csv(ROOT / "results/manuscript_values.csv")}
    est = mapping(audit["estimands"])
    cond = mapping(est["conditional_success_among_attempted"])
    b = mapping(audit["B"])
    funnel = mapping(b["funnel"])
    detector = mapping(b["leakage_detector_audit"])
    manuscript = docx_text(ROOT / "manuscript/Preparation_draft_EN.docx")
    tables_doc = docx_text(ROOT / "manuscript/Editable_tables_EN.docx")
    supplement = docx_text(ROOT / "manuscript/Supplement_protocols_DRAFT_EN.docx")
    slots = read_csv(RESULTS / "adjusted_A_slots.csv")
    papers = read_csv(RESULTS / "adjusted_A_papers.csv")
    dist = read_csv(RESULTS / "adjusted_A_distribution.csv")
    adjusted_success = sum(p["adjusted_majority"] == "success" for p in papers)
    frozen_success = sum(p["frozen_majority"] == "success" for p in papers)
    checks = [
        (
            "blind outcome byte-identical to seal",
            sha256(BLIND.read_bytes()) == FROZEN_BLIND_SHA,
            "",
        ),
        (
            "primary adjudication unchanged",
            sha256(ADJUDICATION.read_bytes()) == FROZEN_ADJUDICATION_SHA,
            "",
        ),
        (
            "reveal ledger references sealed blind outcome",
            json.loads(REVEAL_LEDGER.read_text())["blind_outcome_sha256"] == FROZEN_BLIND_SHA,
            "",
        ),
        (
            "frozen main_analysis conditional successes unchanged (0)",
            frozen["conditional_rate_among_attempted"]["successes"] == 0,
            "",
        ),
        (
            "A rows 30 / B 90 / C 25 / E 130",
            audit["rows"] == {"A": 30, "B": 90, "C": 25, "E": 130},
            "",
        ),
        (
            "no unresolved evidence references",
            mapping(audit["evidence_references"])["unresolved"] == 0,
            "",
        ),
        ("no validation errors", audit["validation_errors"] == [], ""),
        (
            "no reruns / no original code executed",
            b["reruns_performed"] == 0 and not mapping(audit["E"])["original_code_executed"],
            "",
        ),
        (
            "no investigator sign-off populated",
            all(r["investigator_signoff"] == "" for r in slots),
            "",
        ),
        (
            "adjusted conditional numerator equals paper table majority count",
            mapping(cond["verification_adjusted"])["numerator"] == adjusted_success
            and mapping(cond["frozen_mechanical"])["numerator"] == frozen_success,
            f"adjusted {adjusted_success}, frozen {frozen_success}",
        ),
        (
            "distribution sums to 10 papers (frozen and adjusted)",
            sum(int(r["frozen_papers"]) for r in dist) == 10
            and sum(int(r["adjusted_papers"]) for r in dist) == 10,
            "",
        ),
        (
            "funnel columns each sum to 100",
            all(
                sum(int(str(v)) for v in mapping(funnel[k]).values()) == 100
                for k in ("frozen", "adjusted", "envelope")
            ),
            "",
        ),
        (
            "B verification counts sum to 90",
            sum(int(str(b[k])) for k in ("confirmed", "state_incorrect", "unknown")) == 90,
            "",
        ),
        (
            "leakage audit partitions flags",
            sum(int(str(detector[k])) for k in ("confirmed", "corrected", "unresolved"))
            == int(str(detector["classifications"])),
            "",
        ),
        (
            "manuscript_values frozen/adjusted numerators match audit",
            values["frozen_majority_successes_among_attempted"]["value"]
            == str(mapping(cond["frozen_mechanical"])["numerator"])
            and values["adjusted_majority_successes_among_attempted"]["value"]
            == str(mapping(cond["verification_adjusted"])["numerator"]),
            "",
        ),
        (
            "manuscript narrative cites frozen and adjusted conditional values",
            f"{mapping(cond['frozen_mechanical'])['numerator']}/10 (" in manuscript
            and f"{mapping(cond['verification_adjusted'])['numerator']}/10 (" in manuscript,
            "",
        ),
        ("manuscript cites Tables 6-9", all(f"Table {n}" in manuscript for n in (6, 7, 8, 9)), ""),
        (
            "Tables 6-9 present in editable tables and supplement",
            all(f"Table {n}." in tables_doc and f"Table {n}." in supplement for n in (6, 7, 8, 9)),
            "",
        ),
        (
            "manuscript labels review as AI-assisted and provisional, D pending",
            "AI-assisted" in manuscript
            and "provisional" in manuscript
            and "Section D" in manuscript
            and "pending" in manuscript,
            "",
        ),
        (
            "manuscript does not claim submission readiness",
            "NOT SUBMISSION READY" in manuscript
            and not re.search(
                r"\bsubmission[- ]ready\b(?!\W*NOT)", manuscript.replace("NOT SUBMISSION READY", "")
            ),
            "",
        ),
        ("D case count is zero", values["D_human_validation_cases_completed"]["value"] == "0", ""),
        (
            "no D result terms in adjusted outputs",
            not any(t in (RESULTS / "verification_import_audit.json").read_text() for t in D_TERMS),
            "",
        ),
    ]
    checks.extend(input_state_checks(manuscript, tables_doc, supplement))
    return checks


def input_state_checks(
    manuscript: str, tables_doc: str, supplement: str
) -> list[tuple[str, bool, str]]:
    svg = (ROOT / "manuscript/figure1_framework.svg").read_text()
    readme = (ROOT / "README.md").read_text()
    note = (ROOT / "INPUT_STATE_IDENTIFIABILITY_NOTE.md").read_text()
    stage_rows = read_csv(ROOT / "results/requirement_stage_table.csv")
    descriptives = read_csv(ROOT / "results/input_state_descriptives.csv")
    ledger = {r["reference_id"]: r for r in read_csv(ROOT / "data/verified_references.csv")}
    verification = (ROOT / "review/INPUT_STATE_CITATION_VERIFICATION.md").read_text()
    new_ids = [
        "fair_principles",
        "force11_data_citation",
        "rda_dynamic_data_citation",
        "rauber_dynamic_subsets",
        "proll_rauber_dynamic",
        "klump_versioning",
        "pasquier_provenance",
        "swhid_content_hash",
        "sandve_ten_rules",
        "stodden_enhancing",
        "zhao_annotation_versions",
    ]
    lower = manuscript.lower()
    return [
        ("top-level stage order unchanged", " → ".join(STAGES) == STAGE_SEQUENCE, STAGE_SEQUENCE),
        (
            "ACCESS subcomponents in Figure 1, README and note",
            all(q in svg and q in readme and q in note for q in ACCESS_QUESTIONS),
            "; ".join(ACCESS_QUESTIONS),
        ),
        (
            "requirement-to-stage table complete",
            [(r["requirement"], r["stage"], r["purpose"]) for r in stage_rows] == list(STAGE_TABLE)
            and all(row[0] in tables_doc and row[0] in supplement for row in STAGE_TABLE),
            f"{len(STAGE_TABLE)} rows",
        ),
        (
            "input-state descriptives carry no invented fields",
            all(
                any(w in r["value"] for w in ("not recorded", "not assessed"))
                for r in descriptives
                if r["item"] in ("schema/release identifier", "historical version check")
            )
            and "not recorded (live service snapshots)" in tables_doc,
            "",
        ),
        (
            "calibrated wording present and uncalibrated wording absent",
            KEY_STATEMENT in manuscript
            and "not by itself establish" in manuscript
            and "did not quantify dataset drift" in manuscript
            and not any(term in lower for term in UNCALIBRATED),
            "",
        ),
        (
            "no novelty claim for provisional terms",
            "no novelty is claimed" in manuscript and "novel" not in note.lower(),
            "",
        ),
        (
            "new citations verified in ledger and verification table",
            all(
                i in REFERENCES
                and i in ledger
                and ledger[i]["DOI"] in manuscript
                and ledger[i]["DOI"] in verification
                for i in new_ids
            ),
            f"{len(new_ids)} references",
        ),
        (
            "A/B/C/E adjudication records unchanged",
            tree_sha(IMPORT_DIR) == FROZEN_ABCE_SHA,
            "",
        ),
        (
            "no rerun or original-code execution introduced",
            not any(t in lower for t in RERUN_TERMS)
            and "no original code was executed" in manuscript,
            "",
        ),
    ]


def write_report() -> None:
    checks = run()
    lines = [
        "# Verification-adjusted layer QC report",
        "",
        "Generated by `src/paper2/verification_qc.py` from current `results/` and `manuscript/`.",
        "",
        "| Check | Result | Note |",
        "|---|---|---|",
    ]
    for name, ok, note in checks:
        lines.append(f"| {name} | {'PASS' if ok else 'FAIL'} | {note} |")
    passed = sum(ok for _, ok, _ in checks)
    lines += ["", f"{passed}/{len(checks)} checks passed.", ""]
    (ROOT / "review/VERIFICATION_ADJUSTED_QC.md").write_text("\n".join(lines))
    if passed != len(checks):
        raise SystemExit("verification QC failed")


if __name__ == "__main__":
    write_report()
