import json
from pathlib import Path

import pytest

from paper2.core import sha256, snapshot
from paper2.pilot_inputs import (
    FILE_CAP,
    Deposit,
    acquire_deposit,
    proteomexchange_files,
    serveable,
    zenodo_files,
)

PROTEOMEXCHANGE = b"""<?xml version="1.0" encoding="UTF-8"?>
<ProteomeXchangeDataset id="PXD027610">
  <DatasetFileList>
    <DatasetFile id="FILE_1" name="table.csv">
      <cvParam name="URI" value="https://example.org/table.csv"/>
    </DatasetFile>
    <DatasetFile id="FILE_2" name="run.raw"/>
  </DatasetFileList>
</ProteomeXchangeDataset>
"""


def test_zenodo_listing_keeps_announced_size_and_checksum() -> None:
    body = json.dumps(
        {"files": [{"key": "TableZ4.xlsx", "size": 12489, "checksum": "md5:x", "links": {"self": "https://zenodo.org/f"}}]}
    ).encode()
    assert zenodo_files(body) == [
        {
            "name": "TableZ4.xlsx",
            "bytes": 12489,
            "checksum": "md5:x",
            "url": "https://zenodo.org/f",
        }
    ]


def test_proteomexchange_listing_records_absent_file_sizes() -> None:
    files = proteomexchange_files(PROTEOMEXCHANGE)
    assert [file["name"] for file in files] == ["table.csv", "run.raw"]
    assert all(file["bytes"] is None for file in files)
    assert files[0]["url"] == "https://example.org/table.csv"
    assert not any(serveable(file) for file in files)


@pytest.mark.parametrize(
    "entry",
    [
        {"name": "a.csv", "bytes": FILE_CAP + 1, "url": "https://example.org/a.csv"},
        {"name": "a.bam", "bytes": 10, "url": "https://example.org/a.bam"},
        {"name": "a.csv", "bytes": None, "url": "https://example.org/a.csv"},
        {"name": "a.csv", "bytes": 10, "url": "ftp://example.org/a.csv"},
        {"name": "a.csv", "bytes": 0, "url": "https://example.org/a.csv"},
    ],
)
def test_oversized_unsized_binary_or_non_http_entries_are_not_served(
    entry: dict[str, object],
) -> None:
    assert not serveable(entry)
    assert serveable({"name": "a.csv", "bytes": 10, "url": "https://example.org/a.csv"})


def test_restricted_route_is_recorded_without_claiming_unavailability(tmp_path: Path) -> None:
    deposit = Deposit(
        paper_id="PMID:39730532",
        accession="ADNI",
        kind="application_required",
        listing_url="",
        note="data use application required",
    )
    record = acquire_deposit(deposit, tmp_path)
    assert record["status"] == "not_retrievable_by_public_route_application_required"
    assert record["retained_files"] == []
    assert record["note"] == "data use application required"


def test_retained_listing_is_immutable(tmp_path: Path) -> None:
    body = b"{}\n"
    path = tmp_path / "listing" / "response"
    snapshot(path, body)
    assert sha256(path.read_bytes()) == sha256(body)
    with pytest.raises(ValueError):
        snapshot(path, b"{\"files\": []}\n")
    assert path.read_bytes() == body
