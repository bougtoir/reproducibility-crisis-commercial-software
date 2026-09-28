import csv
import json
from pathlib import Path

import pytest

from paper2 import verification_layer as vl
from paper2.core import sha256

ROOT = Path(__file__).resolve().parents[1]
AUDIT = ROOT / "results/verification_adjusted/verification_import_audit.json"


def audit() -> dict[str, object]:
    return json.loads(AUDIT.read_text())


def test_import_validates_against_frozen_records_and_preserves_seals() -> None:
    result = vl.Importer(vl.IMPORT_DIR).run()
    assert result["validation_errors"] == []
    assert result["rows"] == {"A": 30, "B": 90, "C": 25, "E": 130}
    frozen = result["frozen_inputs_sha256"]
    assert isinstance(frozen, dict)
    assert frozen["blind_outcome"] == sha256(vl.BLIND.read_bytes())
    assert frozen["primary_adjudication"] == sha256(vl.ADJUDICATION.read_bytes())
    refs = result["evidence_references"]
    assert isinstance(refs, dict) and refs["unresolved"] == 0
    hashes = result["source_files_sha256"]
    assert isinstance(hashes, dict)
    for name, digest in hashes.items():
        assert sha256((vl.IMPORT_DIR / name).read_bytes()) == digest


def test_adjusted_layer_is_separate_and_keeps_estimands_apart() -> None:
    a = audit()
    est = a["estimands"]
    assert isinstance(est, dict)
    reach = est["attempt_reachability"]
    cond = est["conditional_success_among_attempted"]
    yield_ = est["observed_end_to_end_yield_100"]
    assert reach["numerator"] == 10 and reach["denominator"] == 100
    assert cond["frozen_mechanical"]["numerator"] == 0
    assert cond["verification_adjusted"]["numerator"] == 1
    assert cond["verification_adjusted"]["denominator"] == 10
    assert yield_["verification_adjusted"]["denominator"] == 100
    assert round(cond["verification_adjusted"]["wilson_95_lower"], 3) == 0.018
    assert round(cond["verification_adjusted"]["wilson_95_upper"], 3) == 0.404
    frozen_analysis = json.loads((ROOT / "results/main_analysis.json").read_text())
    assert frozen_analysis["conditional_rate_among_attempted"]["successes"] == 0


def test_b_corrections_are_interpretation_only_and_unsigned() -> None:
    a = audit()
    b = a["B"]
    assert isinstance(b, dict)
    assert b["reruns_performed"] == 0
    assert b["confirmed"] + b["state_incorrect"] + b["unknown"] == 90
    funnel = b["funnel"]
    for key in ("frozen", "adjusted", "envelope"):
        assert sum(funnel[key].values()) == 100
    detector = b["leakage_detector_audit"]
    assert (
        detector["confirmed"] + detector["corrected"] + detector["unresolved"]
        == detector["classifications"]
    )
    with (ROOT / "results/verification_adjusted/adjusted_A_slots.csv").open() as handle:
        rows = list(csv.DictReader(handle))
    assert len(rows) == 30 and all(r["investigator_signoff"] == "" for r in rows)


def test_c_and_e_are_descriptive_and_d_absent() -> None:
    a = audit()
    assert a["C"]["post_outcome_corrections"] == 0
    assert a["E"]["original_code_executed"] is False
    assert a["E"]["blind_scores_revised_from_reveal"] is False
    text = AUDIT.read_text()
    assert "participant" not in text.lower() or "Not human-participant validation" in text
    with (ROOT / "results/manuscript_values.csv").open() as handle:
        values = {r["claim_id"]: r for r in csv.DictReader(handle)}
    assert values["D_human_validation_cases_completed"]["value"] == "0"
    assert values["adjusted_majority_successes_among_attempted"]["status"].startswith(
        "verification_adjusted"
    )
    assert values["frozen_majority_successes_among_attempted"]["value"] == "0"


def test_tampered_review_is_rejected(tmp_path: Path) -> None:
    for name in vl.REVIEW_FILES.values():
        (tmp_path / name).write_bytes((vl.IMPORT_DIR / name).read_bytes())
    path = tmp_path / vl.REVIEW_FILES["A"]
    with path.open(newline="") as handle:
        rows = list(csv.DictReader(handle))
    rows[0]["slot"] = "9"
    with path.open("w", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=list(rows[0]))
        writer.writeheader()
        writer.writerows(rows)
    importer = vl.Importer(tmp_path)
    importer.copy_reviews = lambda: {  # type: ignore[method-assign]
        s: vl.read_csv(tmp_path / n) for s, n in vl.REVIEW_FILES.items()
    }
    with pytest.raises(ValueError, match="frozen 30 sealed slots"):
        importer.run()
