import json
import shutil
import zipfile
from pathlib import Path

import pytest

from paper2 import documents
from paper2.build import ROOT
from paper2.core import sha256


def test_document_build_is_byte_reproducible_and_archives_only_known_files(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    for directory in ("protocols", "results", "data", "schemas", "review", "src", "tests"):
        shutil.copytree(
            ROOT / directory,
            tmp_path / directory,
            ignore=shutil.ignore_patterns("raw", "__pycache__"),
        )
    for name in ("requirements.lock", "pyproject.toml", "README.md", "Makefile"):
        shutil.copyfile(ROOT / name, tmp_path / name)
    monkeypatch.setattr(documents, "ROOT", tmp_path)
    documents.build_documents()
    manifest = tmp_path / "results/preparation_artifact_manifest.json"
    package = tmp_path / "dist/PaperII_preparation_NOT_SUBMISSION_READY.zip"
    before_manifest, before_package = manifest.read_bytes(), sha256(package.read_bytes())
    (tmp_path / "manuscript/private-unlisted.txt").write_text("synthetic exclusion canary")
    documents.build_documents()
    assert manifest.read_bytes() == before_manifest
    assert sha256(package.read_bytes()) == before_package
    assert "manuscript/private-unlisted.txt" not in json.loads(manifest.read_text())
    with zipfile.ZipFile(package) as archive:
        assert "manuscript/private-unlisted.txt" not in archive.namelist()
        assert "results/preparation_artifact_manifest.json" in archive.namelist()
    with zipfile.ZipFile(tmp_path / "manuscript/Preparation_draft_EN.docx") as archive:
        assert {member.date_time[:3] for member in archive.infolist()} == {
            (documents.BUILD_TIME.year, documents.BUILD_TIME.month, documents.BUILD_TIME.day)
        }
