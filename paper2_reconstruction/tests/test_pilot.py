import json
from pathlib import Path

import pytest

from paper2.firewall import Custodian, build_archive, vetted_wheels
from paper2.pilot import DEVIATION, INSTRUCTION, input_status, pilot_papers, serve_slot

ARTICLE = b"<article>retained publication text</article>\n"
RECORD = {
    "assessments": [
        {
            "paper_id": "PMID:2",
            "sampling_stratum": "Clinical_Medicine",
            "target_candidate": "second target",
            "pilot_decision": "provisional_pilot_case",
        },
        {
            "paper_id": "PMID:1",
            "sampling_stratum": "Biomedical_Basic",
            "target_candidate": "first target",
            "pilot_decision": "provisional_pilot_case",
        },
        {
            "paper_id": "PMID:3",
            "sampling_stratum": "Biomedical_Basic",
            "target_candidate": "not selected",
            "pilot_decision": "not_selected_lower_rank",
        },
    ]
}
LEDGER = {
    "deposits": [
        {"paper_id": "PMID:1", "accession": "D1", "status": "listing_retained_no_file_below_cap"},
        {"paper_id": "PMID:1", "accession": "D2", "status": "not_retrievable_by_public_route"},
    ]
}


def write(path: Path, payload: dict[str, object]) -> Path:
    path.write_text(json.dumps(payload))
    return path


def test_pilot_papers_takes_only_provisional_cases_in_deterministic_order(tmp_path: Path) -> None:
    papers = pilot_papers(write(tmp_path / "record.json", RECORD))
    assert [paper["paper_id"] for paper in papers] == ["PMID:1", "PMID:2"]
    assert papers[0]["target_candidate"] == "first target"


def test_pilot_papers_refuses_a_record_without_a_provisional_case(tmp_path: Path) -> None:
    empty = {"assessments": [dict(RECORD["assessments"][2])]}  # type: ignore[index]
    with pytest.raises(ValueError):
        pilot_papers(write(tmp_path / "empty.json", empty))


def test_input_status_reports_every_deposit_state_for_the_paper(tmp_path: Path) -> None:
    ledger = write(tmp_path / "ledger.json", LEDGER)
    assert input_status(ledger, "PMID:1") == (
        "D1: listing_retained_no_file_below_cap; D2: not_retrievable_by_public_route"
    )
    assert input_status(ledger, "PMID:9") == "no public deposit route recorded"


def test_slot_package_holds_only_publication_and_vetted_software(tmp_path: Path) -> None:
    wheels = vetted_wheels(
        Path(__file__).resolve().parents[1] / "data" / "raw" / "vetted-wheels-20260924"
    )
    archive, items = build_archive(tmp_path, [("PMID:1", ARTICLE)], wheels)
    custodian = Custodian(archive, items)
    served, journal, article = serve_slot(custodian, tmp_path / "slot", "PMID:1", "solver_slot_1")
    assert article == "1.xml"
    assert sorted(served) == sorted(["1.xml", *wheels])
    package = sorted(path.name for path in (tmp_path / "slot" / "package").iterdir())
    assert package == sorted(served)
    assert not any("original" in name for name in served)
    events = [
        json.loads(path.read_bytes())
        for path in sorted((tmp_path / "slot" / "access-events").glob("event-*.json"))
    ]
    assert events[0]["value"]["deviation"] == DEVIATION
    assert all(event["kind"] != "access_denied" for event in events)
    assert len(journal.head) == 64 and len(events) == 1 + len(served)


def test_instruction_states_absent_inputs_and_forbids_fabrication() -> None:
    text = INSTRUCTION.format(
        slot=1, target="T", served="- /input/1.xml", article="1.xml", input_status="D1: absent"
    )
    assert "Deposited study inputs were NOT served for this run: D1: absent" in text
    assert "Do not fabricate" in text
    assert "no access to the original implementation" in text
