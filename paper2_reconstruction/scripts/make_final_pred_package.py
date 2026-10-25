#!/usr/bin/env python3
"""Assemble the FINAL_PRE_D delivery artifacts.

Creates near-final named copies of the manuscript artifacts, renders the
manuscript PDF via LibreOffice, and builds
dist/PaperII_FINAL_PRE_D_PACKAGE.zip from the manifest file list plus the
FINAL_PRE_D artifacts. Run AFTER `make build`.
"""
import json
import shutil
import subprocess
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
MAN = ROOT / "manuscript"
DIST = ROOT / "dist"

final_docx = MAN / "Preparation_draft_EN_inline_FINAL_PRE_D.docx"
final_supp = MAN / "Supplement_FINAL_PRE_D.docx"
shutil.copyfile(MAN / "Preparation_draft_EN.docx", final_docx)
shutil.copyfile(MAN / "Supplement_protocols_DRAFT_EN.docx", final_supp)

subprocess.run(
    ["soffice", "--headless", "--convert-to", "pdf",
     "--outdir", str(MAN), str(final_docx)],
    check=True, capture_output=True, timeout=300,
)
final_pdf = final_docx.with_suffix(".pdf")
assert final_pdf.exists(), "PDF conversion failed"

subprocess.run(
    ["soffice", "--headless", "--convert-to", "pdf",
     "--outdir", str(MAN), str(final_supp)],
    check=True, capture_output=True, timeout=300,
)
final_supp_pdf = final_supp.with_suffix(".pdf")
assert final_supp_pdf.exists(), "supplement PDF conversion failed"

manifest = json.loads(
    (ROOT / "results/preparation_artifact_manifest.json").read_text()
)
files = [ROOT / rel for rel in manifest] + [
    final_docx, final_pdf, final_supp, final_supp_pdf,
    ROOT / "results/preparation_artifact_manifest.json",
]
out = DIST / "PaperII_FINAL_PRE_D_PACKAGE.zip"
with zipfile.ZipFile(out, "w", compression=zipfile.ZIP_DEFLATED) as z:
    for p in sorted(set(files)):
        if p.is_file():
            z.writestr(str(p.relative_to(ROOT)), p.read_bytes())
print(f"{out.name}: {len(files)} entries, {out.stat().st_size} bytes")
