import json
import os
import zipfile
from datetime import datetime, timezone
from io import BytesIO
from pathlib import Path

import matplotlib
from docx import Document
from docx.document import Document as WordDocument
from docx.shared import Inches, Pt, RGBColor
from matplotlib import pyplot as plt
from matplotlib.patches import FancyBboxPatch
from pptx import Presentation
from pptx.dml.color import RGBColor as SlideColor
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches as SlideInches
from pptx.util import Pt as SlidePt

from paper2.build import ROOT
from paper2.core import Row, read_csv, sha256
from paper2.input_state import ACCESS_QUESTIONS, KEY_STATEMENT, MINIMUM_REPORTING
from paper2.main_analysis import listed, mapping

STATUS = (
    "DRAFT — NOT SUBMISSION READY: main study sealed; delegated mechanical adjudication; "
    "A/B/C/E AI-assisted review imported as a provisional sensitivity layer; investigator "
    "sign-off, human adjudication and human validation (Section D) pending"
)
BUILD_TIME = datetime.fromtimestamp(
    int(os.environ.get("SOURCE_DATE_EPOCH", "315532800")), timezone.utc
)
ARTIFACT_NAMES = (
    "Preparation_draft_EN.docx",
    "Supplement_protocols_DRAFT_EN.docx",
    "Editable_tables_EN.docx",
    "Figures_editable_EN.pptx",
    "figure1_framework.svg",
    "figure1_framework.pdf",
    "figure1_framework.png",
)
TITLE = "From scientific description to independent computational reconstruction"
SUBTITLE = "Prospective Paper II study linked to the EPJ commercial-software corpus"
CAPTION = (
    "Figure 1. Conceptual Wet/Dry framework. ACCESS is refined internally into three "
    "questions (resource accessible; exact input state identifiable; historical input state "
    "retrievable) because public resources can drift over time. Paper II focuses on "
    "RECONSTRUCT and observes EXECUTE and REPRODUCE. ROBUST belongs to future Paper III and "
    "is excluded. The two pathways are an analogy, not a literal equivalence or an empirical "
    "result."
)
ACCESS_NOTE = "Public resources can drift over time; availability ≠ identity of the input state."
STAGES = ["ACCESS", "RECONSTRUCT", "EXECUTE", "REPRODUCE", "ROBUST"]
STAGE_COLORS = ["#E2E8F0", "#0F766E", "#CCFBF1", "#CCFBF1", "#FEF3C7"]
WET = [
    "Methods",
    "Materials",
    "Environmental /\nprocedural specification",
    "Reconstruct\nexperiment",
    "Execute",
    "Measurement /\nconclusion",
]
DRY = [
    "Methods",
    "Data / resources",
    "Computational\nspecification",
    "Independent\nimplementation",
    "Execute",
    "Numerical result /\nconclusion",
]
REFERENCES = (
    "acm_terms",
    "corebench",
    "fair_principles",
    "force11_data_citation",
    "goodman_crossref",
    "klump_versioning",
    "nasem_metadata",
    "paper2code",
    "paperbench_arxiv",
    "parent_crossref",
    "pasquier_provenance",
    "proll_rauber_dynamic",
    "rauber_dynamic_subsets",
    "rda_dynamic_data_citation",
    "replicationbench",
    "sandve_ten_rules",
    "scienceagentbench_arxiv",
    "stodden_enhancing",
    "swhid_content_hash",
    "wet_crossref",
    "zhao_annotation_versions",
)


def style_document(document: WordDocument) -> None:
    section = document.sections[0]
    section.top_margin = Inches(0.7)
    section.bottom_margin = Inches(0.7)
    section.left_margin = Inches(0.75)
    section.right_margin = Inches(0.75)
    normal = document.styles["Normal"]
    normal.font.name = "Calibri"
    normal.font.size = Pt(10)
    normal.paragraph_format.space_after = Pt(6)
    document.core_properties.author = "Paper II preparation; authorship not finalized"
    document.core_properties.subject = STATUS
    document.core_properties.created = BUILD_TIME
    document.core_properties.modified = BUILD_TIME
    footer = section.footer.paragraphs[0]
    footer.text = STATUS
    footer.runs[0].font.size = Pt(8)
    footer.runs[0].font.color.rgb = RGBColor.from_string("9A3412")


def heading(document: WordDocument, title: str, level: int = 1) -> None:
    document.add_heading(title, level=level)


def paragraph(document: WordDocument, text: str) -> None:
    document.add_paragraph(text)


def caption(document: WordDocument, text: str) -> None:
    line = document.add_paragraph(text)
    line.paragraph_format.space_before = Pt(14)
    line.paragraph_format.space_after = Pt(8)
    line.paragraph_format.keep_with_next = text.startswith("Table")
    line.runs[0].italic = True


def table(document: WordDocument, rows: list[Row], columns: list[tuple[str, str]]) -> None:
    result = document.add_table(rows=1, cols=len(columns))
    result.style = "Table Grid"
    for cell, (_, label) in zip(result.rows[0].cells, columns, strict=False):
        cell.text = label
    for row in rows:
        for cell, (key, _) in zip(result.add_row().cells, columns, strict=False):
            cell.text = row[key].replace("_", " ")
    for row_index, table_row in enumerate(result.rows):
        for cell in table_row.cells:
            for line in cell.paragraphs:
                line.paragraph_format.keep_with_next = row_index < len(result.rows) - 1
                for run in line.runs:
                    run.font.size = Pt(9)


def markdown(document: WordDocument, text: str) -> None:
    lines = text.splitlines()
    index = 0
    while index < len(lines):
        line = lines[index]
        if line.startswith("|"):
            cells = []
            while index < len(lines) and lines[index].startswith("|"):
                cells.append([cell.strip() for cell in lines[index].strip("|").split("|")])
                index += 1
            columns = [(str(column), name) for column, name in enumerate(cells[0])]
            rows = [{str(column): cell for column, cell in enumerate(row)} for row in cells[2:]]
            table(document, rows, columns)
            continue
        if line.startswith("# "):
            heading(document, line[2:], 1)
        elif line.startswith("## "):
            heading(document, line[3:], 2)
        elif line.startswith("- "):
            document.add_paragraph(line[2:], style="List Bullet")
        elif line.strip():
            paragraph(document, line.strip("*"))
        index += 1


def framework(output: Path) -> None:
    matplotlib.rcParams["svg.hashsalt"] = "paper2-framework-v1"
    matplotlib.rcParams["svg.fonttype"] = "none"
    matplotlib.rcParams["pdf.fonttype"] = 42
    figure, axis = plt.subplots(figsize=(13.333, 7.5))
    axis.set_xlim(0, 13.333)
    axis.set_ylim(0, 7.5)
    axis.axis("off")
    axis.text(
        0.2,
        7.1,
        "From description to a working scientific system",
        fontsize=21,
        weight="bold",
        color="#0F172A",
    )
    axis.text(0.2, 6.65, "Conceptual framework • Paper II preparation", fontsize=13)
    for index, (name, color) in enumerate(zip(STAGES, STAGE_COLORS, strict=False)):
        x = 0.2 + index * 2.6
        axis.add_patch(
            FancyBboxPatch(
                (x, 5.25),
                2.35,
                0.9,
                boxstyle="round,pad=0.03",
                facecolor=color,
                edgecolor="#64748B",
            )
        )
        axis.text(
            x + 1.175,
            5.7,
            name,
            ha="center",
            va="center",
            weight="bold",
            fontsize=14,
            color="white" if index == 1 else "#0F172A",
        )
        if index < len(STAGES) - 1:
            axis.text(x + 2.47, 5.7, "→", ha="center", va="center", fontsize=19)
    axis.add_patch(
        FancyBboxPatch(
            (0.2, 4.28),
            2.35,
            0.82,
            boxstyle="round,pad=0.02",
            facecolor="#F8FAFC",
            edgecolor="#94A3B8",
            linestyle="--",
        )
    )
    for index, question in enumerate(ACCESS_QUESTIONS):
        axis.text(0.32, 4.95 - index * 0.26, f"├ {question}", fontsize=8.6, color="#0F172A")
    axis.text(3.0, 4.88, "Paper II focus and downstream observations", fontsize=12, color="#0F766E")
    axis.text(10.6, 4.88, "Paper III: excluded", fontsize=12, color="#92400E")
    axis.text(3.0, 4.45, ACCESS_NOTE, fontsize=9.5, color="#475569", style="italic")
    for name, items, y, fill in (("WET", WET, 3.0, "#EFF6FF"), ("DRY", DRY, 1.45, "#F0FDFA")):
        axis.text(0.2, y + 1.0, name, fontsize=14, weight="bold", color="#0F172A")
        for index, label in enumerate(items):
            x = 0.2 + index * 2.17
            axis.add_patch(
                FancyBboxPatch(
                    (x, y),
                    1.95,
                    0.88,
                    boxstyle="round,pad=0.02",
                    facecolor=fill,
                    edgecolor="#CBD5E1",
                )
            )
            axis.text(x + 0.975, y + 0.44, label, fontsize=10.5, ha="center", va="center")
            if index < len(items) - 1:
                axis.text(x + 2.06, y + 0.44, "→", ha="center", va="center", fontsize=16)
    axis.text(
        0.2,
        0.85,
        "Independent implementation ≠ access to the original code",
        fontsize=13,
        color="#0F172A",
    )
    axis.text(
        0.2,
        0.45,
        "Conceptual analogy only. No robustness or multiverse analysis.",
        fontsize=12,
        color="#475569",
    )
    figure.tight_layout()
    for suffix in ("svg", "pdf", "png"):
        metadata: dict[str, object] = (
            {"CreationDate": BUILD_TIME, "ModDate": BUILD_TIME}
            if suffix == "pdf"
            else {"Date": BUILD_TIME.isoformat()}
        )
        figure.savefig(
            output / f"figure1_framework.{suffix}",
            dpi=300,
            facecolor="white",
            metadata=metadata,
        )
    plt.close(figure)
    svg = output / "figure1_framework.svg"
    svg.write_text("\n".join(line.rstrip() for line in svg.read_text().splitlines()) + "\n")


def slides(output: Path, characteristics: list[Row]) -> None:
    presentation = Presentation()
    presentation.core_properties.created = BUILD_TIME
    presentation.core_properties.modified = BUILD_TIME
    presentation.slide_width = SlideInches(13.333)
    presentation.slide_height = SlideInches(7.5)
    slide = presentation.slides.add_slide(presentation.slide_layouts[6])

    def textbox(x: float, y: float, width: float, height: float, text: str, size: int) -> None:
        shape = slide.shapes.add_textbox(
            SlideInches(x),
            SlideInches(y),
            SlideInches(width),
            SlideInches(height),
        )
        frame = shape.text_frame
        frame.margin_left = 0
        frame.margin_right = 0
        frame.word_wrap = True
        frame.text = text
        for line in frame.paragraphs:
            line.font.size = SlidePt(size)
            line.font.color.rgb = SlideColor.from_string("0F172A")

    textbox(0.3, 0.2, 12.5, 0.5, "Figure 1. From description to a working scientific system", 24)
    for index, (name, color) in enumerate(zip(STAGES, STAGE_COLORS, strict=False)):
        shape = slide.shapes.add_shape(
            MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE,
            SlideInches(0.3 + index * 2.6),
            SlideInches(1.05),
            SlideInches(2.25),
            SlideInches(0.72),
        )
        shape.fill.solid()
        shape.fill.fore_color.rgb = SlideColor.from_string(color[1:])
        shape.line.color.rgb = SlideColor.from_string("64748B")
        shape.text = name
        line = shape.text_frame.paragraphs[0]
        line.alignment = PP_ALIGN.CENTER
        line.font.size = SlidePt(17)
        line.font.bold = True
        line.font.color.rgb = SlideColor.from_string("FFFFFF" if index == 1 else "0F172A")
        if index < 4:
            textbox(2.58 + index * 2.6, 1.13, 0.3, 0.5, "→", 20)
    textbox(0.3, 1.82, 2.4, 0.9, "\n".join(f"├ {q}" for q in ACCESS_QUESTIONS), 9)
    textbox(3.0, 1.82, 7.5, 0.4, "Paper II focus and downstream observations", 15)
    textbox(3.0, 2.22, 7.5, 0.4, ACCESS_NOTE, 11)
    textbox(10.7, 1.82, 2.3, 0.5, "Paper III: excluded", 14)
    for name, items, y in (("WET", WET, 2.85), ("DRY", DRY, 4.5)):
        textbox(0.3, y - 0.5, 2.0, 0.4, name, 18)
        for index, label in enumerate(items):
            shape = slide.shapes.add_shape(
                MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE,
                SlideInches(0.3 + index * 2.16),
                SlideInches(y),
                SlideInches(1.91),
                SlideInches(0.95),
            )
            shape.fill.solid()
            shape.fill.fore_color.rgb = SlideColor.from_string("F0FDFA")
            shape.line.color.rgb = SlideColor.from_string("CBD5E1")
            shape.text_frame.word_wrap = True
            shape.text = label
            for line in shape.text_frame.paragraphs:
                line.font.size = SlidePt(13)
                line.font.color.rgb = SlideColor.from_string("0F172A")
                line.alignment = PP_ALIGN.CENTER
            if index < 5:
                textbox(2.22 + index * 2.16, y + 0.28, 0.22, 0.4, "→", 16)
    textbox(0.3, 6.12, 12.5, 0.9, CAPTION, 13)
    slide = presentation.slides.add_slide(presentation.slide_layouts[6])
    textbox(0.3, 0.2, 12.5, 0.6, "Table 1. Deposited source rows and PMID identities", 24)
    grid = slide.shapes.add_table(
        len(characteristics) + 1,
        3,
        SlideInches(0.5),
        SlideInches(1.2),
        SlideInches(12.0),
        SlideInches(4.8),
    ).table
    for col, label in enumerate(("Recorded EPJ field", "Source rows", "Unique PMIDs in field")):
        grid.cell(0, col).text = label
    for row_index, row in enumerate(characteristics, 1):
        for col, key in enumerate(("field", "source_rows", "unique_pmids_within_field")):
            grid.cell(row_index, col).text = row[key].replace("_", " ")
    for table_row in grid.rows:
        for cell in table_row.cells:
            for line in cell.text_frame.paragraphs:
                line.font.size = SlidePt(16)
    textbox(
        0.5,
        6.35,
        12.0,
        0.7,
        "Observed source audit only. Field memberships overlap; do not sum within-field "
        "unique counts as a unique-paper total.",
        15,
    )
    presentation.save(str(output / "Figures_editable_EN.pptx"))


def reference_text(row: Row) -> str:
    authors = row["authors"]
    if row["reference_id"] == "parent_crossref":
        authors = "Onishi T; Ikenoue T [name order follows supplied author specification]"
    elif row["reference_id"] == "wet_crossref":
        authors = "Onishi T [name order follows supplied author specification]"
    elif row["reference_id"] == "nasem_metadata":
        authors = "National Academies of Sciences, Engineering, and Medicine"
    return f"{authors} ({row['year']}). {row['title']}. {row['publication_status']}. {row['URL']}"


def zip_member(archive: zipfile.ZipFile, name: str, content: bytes) -> None:
    stamp = BUILD_TIME.timetuple()
    member = zipfile.ZipInfo(name, (stamp.tm_year, stamp.tm_mon, stamp.tm_mday, 0, 0, 0))
    member.compress_type = zipfile.ZIP_DEFLATED
    member.external_attr = 0o100644 << 16
    archive.writestr(member, content)


def normalize_office_archive(path: Path) -> None:
    output = BytesIO()
    with zipfile.ZipFile(path) as original, zipfile.ZipFile(output, "w") as normalized:
        for name in sorted(original.namelist()):
            zip_member(normalized, name, original.read(name))
    path.write_bytes(output.getvalue())


def main_study_paragraphs(analysis: dict[str, object]) -> list[str]:
    denominators = analysis["denominators"]
    conditional = analysis["conditional_rate_among_attempted"]
    primary = analysis["policy_estimate_delegated_states_as_recorded"]
    pending = analysis["policy_estimate_non_attempted_as_unresolved"]
    levels = analysis["run_level"]
    if not (
        isinstance(denominators, dict)
        and isinstance(conditional, dict)
        and isinstance(primary, dict)
        and isinstance(pending, dict)
        and isinstance(levels, dict)
    ):
        raise ValueError("main_analysis.json has an unexpected shape")
    code_runs = levels["failure_codes_runs"]
    stop = levels["stop_reasons"]
    sensitivity = analysis["sensitivity_thresholds_among_attempted"]
    if not (
        isinstance(code_runs, dict) and isinstance(stop, dict) and isinstance(sensitivity, dict)
    ):
        raise ValueError("main_analysis.json run_level has an unexpected shape")
    thresholds = "; ".join(
        f"{name}: " + ", ".join(f"{k} {v}" for k, v in sorted(counts.items()))
        for name, counts in sensitivity.items()
    )
    ledger = json.loads(
        (ROOT / "data/adjudication/main-study-20260925/reveal/reveal_ledger.json").read_text()
    )
    reveal_counts = ledger["counts"]
    return [
        "Status: every number in this section is a delegated mechanical adjudication of "
        "sealed blind outcomes under frozen rules (AMEND-2026-09-25-07). Human adjudication "
        "and investigator verification are pending, and the seven prospective pilot papers "
        "and seven ACCESSIBILITY_GATE_FAILED papers are outside every denominator below.",
        f"Exhaustive route screening of the unique-PMID frame left "
        f"{denominators['route_eligible_population_E']} route-eligible papers, from which "
        f"{denominators['sampled_main_cohort']} were sampled in seven frozen strata. "
        f"{denominators['attempted_papers_complete_triples']} sampled papers completed three "
        f"blind slots ({denominators['runs']} runs); {denominators['not_attempted_papers']} "
        "were not run and retain machine-readable states (Table 2): external public reference "
        "resource required, delegated G1–G5 negative, required inputs absent from the public "
        "listing, specification rejected by the quote/leakage validator, target value not "
        "located verbatim, every candidate input leaking the target, or no input retained.",
        f"Among attempted papers the majority endpoint was met by {conditional['successes']} of "
        f"{conditional['attempted']} (secondary conditional rate; Wilson reference "
        f"{conditional['wilson_95_lower']:.2f}–{conditional['wilson_95_upper']:.2f}). "
        f"Prespecified thresholds — {thresholds}. "
        "Run-level L1–L5 frequencies are in Table 3. Stop reasons: "
        + ", ".join(f"{k} {v}" for k, v in stop.items())
        + ". Failure codes by run (Table 4): "
        + ", ".join(f"{k} {v}" for k, v in code_runs.items())
        + ". F13 denotes the frozen resource ceiling, not demonstrated publication inadequacy.",
        "Under the frozen SAP the headline policy estimand retains non-attempted papers as "
        "policy non-success. Taking the delegated states as recorded, the simultaneous "
        f"hypergeometric upper bound on the eligible-population success rate is "
        f"{primary['upper_sampling_policy_rate']:.3f} (weighted point estimate "
        f"{primary['upper_policy_rate']:.3f}). Because gate verification is "
        "pending, an identification envelope that treats every non-attempted paper as "
        f"unresolved reaches {pending['upper_sampling_policy_rate']:.3f}; it is a bound, not an "
        "estimate. No random-intercept model, perturbation, multiverse or alternative-"
        "implementation analysis was performed.",
        "Descriptive original-code reveal (Table 5) was recorded only after the seal: "
        f"{reveal_counts['described']} attempted papers named public code routes or shipped "
        "scripts inside the data deposit (withheld from the solver) and are described; "
        f"{reveal_counts['missing']} named none and receive a missing reveal assessment. "
        "Nothing was executed and no blind score was revisited.",
    ]


def pct(item: dict[str, object]) -> str:
    return (
        f"{item['numerator']}/{item['denominator']} ({100 * float(str(item['proportion'])):.1f}%; "
        f"Wilson 95% CI {100 * float(str(item['wilson_95_lower'])):.1f}–"
        f"{100 * float(str(item['wilson_95_upper'])):.1f}%)"
    )


def verification_paragraphs(audit: dict[str, object]) -> list[str]:
    est = mapping(audit["estimands"])
    a = mapping(audit["A"])
    b = mapping(audit["B"])
    c = mapping(audit["C"])
    e = mapping(audit["E"])
    refs = mapping(audit["evidence_references"])
    detector = mapping(b["leakage_detector_audit"])
    funnel = mapping(b["funnel"])
    adjusted = mapping(funnel["adjusted"])
    cond = mapping(est["conditional_success_among_attempted"])
    yield_ = mapping(est["observed_end_to_end_yield_100"])
    policy = mapping(est["policy_estimate_eligible_population"])
    adj_policy = mapping(policy["verification_adjusted_states"])
    env_policy = mapping(policy["verification_adjusted_unresolved_envelope"])
    changed = [p for p in listed(a["papers"]) if p["changed"] == "yes"]
    changed_text = "; ".join(
        f"{p['paper_id']} {p['frozen_successful_slots']} → {p['adjusted_successful_slots']}"
        for p in changed
    )
    barriers = sum(
        int(str(v))
        for k, v in adjusted.items()
        if k not in {"slots_completed", "specifiable_runnable", "unresolved"}
    )
    completeness = mapping(e["completeness_described_papers"])
    return [
        "Status: this section reports a separate verification-adjusted sensitivity layer built "
        "from an AI-assisted independent evidence review of sections A (30 sealed slots), B (90 "
        "non-executed papers), C (historical G1–G5, pilot and access-barrier layers) and E "
        "(descriptive reveal). The review is provisional and awaits investigator approval; it "
        "is not human-participant validation, no investigator sign-off field was populated, and "
        f"the sealed blind outcome is byte-identical (SHA-256 "
        f"{str(mapping(audit['frozen_inputs_sha256'])['blind_outcome'])[:16]}…). Every reviewed "
        f"paper/slot identifier validated against the frozen records and {refs['total']} cited "
        f"evidence references resolved ({refs['unresolved']} unresolved).",
        f"Three estimands are kept separate. Attempt reachability among the 100 selected papers "
        f"was {pct(mapping(est['attempt_reachability']))} under both layers. Conditional majority "
        f"success among attempted papers was {pct(mapping(cond['frozen_mechanical']))} under the "
        f"frozen mechanical adjudication and {pct(mapping(cond['verification_adjusted']))} under "
        f"the verification-adjusted layer, in which {a['slots_proposed_changed']} of "
        f"{a['slots_reviewed']} slot classifications were proposed for change and "
        f"{a['slots_mechanical_supported']} were supported ({changed_text}; Table 6). The observed "
        "end-to-end yield under the frozen workflow was "
        f"{pct(mapping(yield_['frozen_mechanical']))} "
        f"frozen and {pct(mapping(yield_['verification_adjusted']))} adjusted. None of these is an "
        "intrinsic probability that a publication is scientifically reproducible.",
        f"Verification of the {b['papers_reviewed']} non-executed papers confirmed "
        f"{b['confirmed']} recorded states, proposed {b['state_incorrect']} state corrections and "
        f"left {b['unknown']} unresolved (Table 7). Non-confirmed states by frozen category: "
        + ", ".join(f"{k} {v}" for k, v in mapping(b["corrections_by_frozen_state"]).items())
        + f". Of {detector['classifications']} all_inputs_leak_target flags, "
        f"{detector['confirmed']} confirmed, {detector['corrected']} corrected and "
        f"{detector['unresolved']} unresolved. Whole-token matches inside compressed or binary "
        "deposited content and raw input matrices produced false-positive target-leakage flags; "
        "this is reported as a methodological finding. Under the adjusted layer the cohort "
        f"comprises 10 attempted papers, {barriers} verified pre-reconstruction barriers, "
        f"{adjusted.get('specifiable_runnable', 0)} papers reclassified as specifiable and "
        f"runnable but never executed, and {adjusted.get('unresolved', 0)} unknown. "
        "Reclassified papers were "
        "not run; the correction is attrition interpretation only, and non-executed papers are "
        "not reconstruction failures.",
        "For the eligible-population policy estimand, treating reclassified and unknown papers as "
        f"unresolved gives a simultaneous sampling-and-identification envelope of "
        f"{adj_policy['lower_sampling_policy_rate']:.3f}–"
        f"{adj_policy['upper_sampling_policy_rate']:.3f} (weighted point range "
        f"{adj_policy['lower_policy_rate']:.3f}–{adj_policy['upper_policy_rate']:.3f}); "
        "treating every non-confirmed leakage classification as unresolved widens the upper "
        f"bound to {env_policy['upper_sampling_policy_rate']:.3f}. These are bounds, not "
        "estimates.",
        f"Section C reviewed {c['records_reviewed']} historical G1–G5, pilot and "
        f"ACCESSIBILITY_GATE_FAILED records and supported {c['confirmed']}; no post-outcome "
        "change was made, the pilot stays outside the main denominator and the access-barrier "
        f"layer remains auxiliary. Section E reviewed {e['dimension_rows']} reveal dimension rows "
        f"for {e['papers']} attempted papers ({e['papers_with_code_route']} with a code route) and "
        f"proposed {e['proposed_corrections']} conservative completeness corrections, chiefly "
        "where "
        "deposited downstream scripts start from pre-computed inputs and cannot establish upstream "
        "normalisation, filtering or exclusion completeness (Table 8; explicit "
        f"{mapping(completeness['frozen']).get('explicit', 0)} → "
        f"{mapping(completeness['adjusted']).get('explicit', 0)} dimensions). The reveal remains "
        "descriptive; no original code was executed and no blind score was revisited.",
    ]


def verification_tables(document: WordDocument, results: Path) -> None:
    caption(
        document,
        "Table 6. Paper-level successful-slot distribution among 10 attempted papers: frozen "
        "mechanical versus verification-adjusted (AI-assisted review, provisional).",
    )
    table(
        document,
        read_csv(results / "adjusted_A_distribution.csv"),
        [
            ("successful_slots", "Successful slots"),
            ("frozen_papers", "Frozen papers"),
            ("adjusted_papers", "Adjusted papers"),
        ],
    )
    caption(
        document,
        "Table 7. 100-paper funnel by state: frozen record, verification-adjusted state and "
        "unresolved envelope (no reclassified paper was executed).",
    )
    table(
        document,
        read_csv(results / "adjusted_B_funnel.csv"),
        [
            ("state", "State"),
            ("frozen", "Frozen"),
            ("verification_adjusted", "Adjusted"),
            ("unresolved_envelope", "Envelope"),
        ],
    )
    caption(
        document,
        "Table 8. Descriptive reveal completeness corrections proposed by the E review "
        "(papers with a code route only; descriptive, not executed).",
    )
    table(
        document,
        [r for r in read_csv(results / "adjusted_E_reveal_ledger.csv") if r["changed"] == "yes"],
        [
            ("paper_id", "Paper"),
            ("dimension", "Dimension"),
            ("frozen_completeness", "Frozen"),
            ("adjusted_completeness", "Adjusted"),
        ],
    )
    caption(
        document,
        "Table 9. Human validation (Section D): PENDING — no participant, case or outcome "
        "data exist.",
    )
    table(
        document,
        [{"item": "Cases completed", "value": "0 (pending external human validation)"}],
        [("item", "Item"), ("value", "Value")],
    )
    caption(
        document,
        "Table 10. Conceptual mapping of reproducibility requirements to framework stages. "
        "ACCESS is refined into resource accessibility, input-state identifiability and "
        "historical-state retrievability; the mapping is conceptual, not an empirical result.",
    )
    table(
        document,
        read_csv(results.parent / "requirement_stage_table.csv"),
        [("requirement", "Requirement"), ("stage", "Stage"), ("purpose", "Purpose")],
    )
    caption(
        document,
        "Table 11. ACCESS-stage input-state fields as recorded in the frozen main-study records "
        "(descriptive; fields not recorded are reported as such; no dataset drift was measured).",
    )
    table(
        document,
        read_csv(results.parent / "input_state_descriptives.csv"),
        [("item", "Item"), ("value", "Value as recorded"), ("source", "Source")],
    )


def main_study_tables(
    document: WordDocument,
    dispositions: list[Row],
    run_levels: list[Row],
    failure_codes: list[Row],
    reveal: list[Row],
) -> None:
    caption(
        document, "Table 2. Main-cohort paper dispositions (delegated; human adjudication pending)."
    )
    table(document, dispositions, [("disposition", "Disposition"), ("papers", "Papers")])
    caption(document, "Table 3. Run-level L1–L5 states across 30 sealed slots.")
    table(document, run_levels, [("level", "Level"), ("state", "State"), ("runs", "Runs")])
    caption(document, "Table 4. Failure taxonomy codes by run and by paper.")
    table(document, failure_codes, [("code", "Code"), ("runs", "Runs"), ("papers", "Papers")])
    caption(document, "Table 5. Descriptive original-code reveal ledger (post-seal, not executed).")
    table(
        document,
        reveal,
        [
            ("paper_id", "Paper"),
            ("reveal_assessment", "Assessment"),
            ("dimension", "Dimension"),
            ("completeness", "Completeness"),
            ("description", "Description"),
        ],
    )


def build_documents() -> None:
    output = ROOT / "manuscript"
    output.mkdir(parents=True, exist_ok=True)
    values = {
        r["claim_id"]: int(r["value"]) for r in read_csv(ROOT / "results/manuscript_values.csv")
    }
    characteristics = read_csv(ROOT / "results/corpus_characteristics.csv")
    precision = read_csv(ROOT / "results/precision_planning.csv")
    references = read_csv(ROOT / "data/verified_references.csv")
    analysis = json.loads((ROOT / "results/main_analysis.json").read_text())
    dispositions = read_csv(ROOT / "results/main_paper_dispositions.csv")
    run_levels = read_csv(ROOT / "results/main_run_levels.csv")
    failure_codes = read_csv(ROOT / "results/main_failure_codes.csv")
    reveal = read_csv(ROOT / "results/reveal_ledger.csv")
    verification_dir = ROOT / "results/verification_adjusted"
    audit = json.loads((verification_dir / "verification_import_audit.json").read_text())
    framework(output)
    slides(output, characteristics)
    document = Document()
    style_document(document)
    heading(document, TITLE, 0)
    paragraph(document, SUBTITLE)
    paragraph(document, STATUS)
    paragraph(document, "Authors, affiliations and corresponding author: NOT FINALIZED.")
    paragraph(document, "Working journal: EPJ Research Infrastructures.")
    heading(document, "Abstract (draft)")
    paragraph(
        document,
        "Access to research artifacts does not establish whether the published scientific "
        "description supports an independent implementation. This prospective study will "
        "estimate publication-grounded computational reconstructability under a specified "
        "agent, information boundary and resource budget, without consulting the original "
        "implementation or purchasing study-specific commercial software. The source audit "
        f"identified {values['source_records']:,} deposited records representing "
        f"{values['unique_pmids']:,} distinct PMIDs. The author approved the unique-PMID "
        "frame while preserving every deposited row. Existing variables remain historical "
        "text detections rather than validated access assessments. The proposed design "
        "separates eligibility, input and resource access, specification, implementation, "
        "execution, numerical agreement and target-linked conclusion preservation. Three "
        "fixed independent run slots per sampled paper support majority, strict and "
        "permissive outcome definitions while retaining missing or contaminated slots as "
        "unresolved. Targets, tolerances and resource ceilings were frozen prospectively. "
        "After a seven-paper prospective pilot and protocol freeze, 100 papers were sampled "
        "from 461 route-eligible papers. Only 10 reached blinded independent reconstruction "
        "under the prespecified access, input and specification rules; 90 stopped before "
        "reconstruction with machine-readable barrier states. Delegated mechanical adjudication "
        "of the 30 sealed slots found no paper meeting the majority criterion; a separate "
        "AI-assisted verification review, provisional and not investigator-signed, proposes one. "
        "Verification of the 90 non-executed papers supported most recorded barriers and "
        "identified a small number of pre-execution misclassifications from an overly "
        "conservative target-leakage detector. Large pre-reconstruction attrition is the primary "
        "empirical finding; the reconstruction result is conditional on attempt, bounded by the "
        "resource ceiling, and does not show human impossibility. Human validation has not been "
        "performed and remains a limitation until real validators complete it.",
    )
    paragraph(
        document,
        "Keywords: computational reconstruction; reproducibility; scientific specification; "
        "research infrastructure; software accessibility; research agents",
    )
    heading(document, "1 Introduction")
    paragraph(
        document,
        "Computational reproducibility terminology varies across communities (Goodman et al., "
        "2016; National Academies of Sciences, Engineering, and Medicine, 2019). ACM's "
        "post-2020 terminology distinguishes execution of authors' artifacts from results "
        "obtained using independently developed artifacts (ACM Publications Board, 2020). "
        "We therefore use publication-grounded independent computational reconstruction "
        "as an operational description rather than proposing a universal definition.",
    )
    paragraph(
        document,
        "The EPJ antecedent examines commercial software and version accessibility "
        "(Onishi and Ikenoue, 2026). Those conditions motivate, but do not answer, whether "
        "the scientific description contains enough operational detail to implement the "
        "reported method independently. The Wet antecedent concerns environmental "
        "confounding in clonal systems (Onishi, 2026). Its extension here is a conceptual "
        "analogy: nominal materials or data may be available while operational specification "
        "remains incomplete. The Wet full text was not accessible in the preparation audit; "
        "no detailed empirical claim is imported from it.",
    )
    paragraph(
        document,
        "Independent paper-to-code reconstruction already has substantial precedent. "
        "PaperBench restricts original implementations but supplies author clarifications "
        "(Starace et al., 2025); ReplicationBench uses curated astrophysics tasks and supplied "
        "execution information (Ye et al., 2025). Paper2Code studies repository generation "
        "(Seo et al., 2026), while ScienceAgentBench evaluates curated scientific programming "
        "tasks (Chen et al., 2025). CORE-Bench instead supplies original code and data "
        "(Siegel et al., 2024; revised 2026). Their endpoints and information packages differ "
        "and must not be pooled as a publication-level reconstructability rate.",
    )
    paragraph(
        document,
        "The proposed contribution is a finite-corpus estimation study with an explicit "
        "publication-only information boundary and zero study-specific software expenditure. "
        "It makes no first-study or terminological novelty claim. Figure 1 shows the "
        "proposed framework and the exclusion of robustness experiments.",
    )
    document.add_picture(str(output / "figure1_framework.png"), width=Inches(6.65))
    document.paragraphs[-1].paragraph_format.keep_with_next = True
    caption(document, CAPTION)
    heading(
        document, "2 Methods — prospectively frozen (protocol text; see supplement for amendments)"
    )
    for title, text in (
        (
            "2.1 Source frame and eligibility",
            "Preserve deposited source records and derive paper identity separately. G1 assesses "
            "whether a central claim depends on a computational procedure with a potentially "
            "objective output. G2 identifies the principal target; G3/G4 assess lawful input and "
            "zero-purchase resource access; G5 describes specification sufficiency. Neither access "
            "nor specification failures exclude otherwise eligible papers from the all-eligible "
            "policy estimand. Historical code/data flags cannot replace these assessments.",
        ),
        (
            "2.1a Input accessibility versus input-state identifiability",
            "Input accessibility was distinguished from input-state identifiability. Within ACCESS "
            "three questions were kept separate: whether a named resource is accessible; whether "
            "the exact state of the input data used in the original analysis can be uniquely "
            "identified from the publication and associated records (input-state "
            "identifiability); and whether that exact previously used state can still be "
            "obtained by an independent third party (historical-state retrievability). For "
            "externally maintained public datasets that authors could not necessarily "
            "redistribute or freeze, we recorded, where available, the repository or accession, "
            "release or version, retrieval date or timestamp, query parameters, file identity, "
            "schema or release identifier, and cryptographic checksum. Public availability alone "
            "was not treated as evidence that the exact historical analytical input remained "
            "identifiable or retrievable. This refinement is descriptive: the frozen G1–G5 gates "
            "and barrier states were not rescored, and no dataset drift was measured. Table 10 "
            "maps these reporting requirements to framework stages; Table 11 summarizes the "
            "fields actually recorded in the frozen main-study records.",
        ),
        (
            "2.2 Sampling and target selection",
            "Using the approved frame, a diverse pilot will finalize operational rules. The "
            "main sample is approximately 100 papers, stratified using a deterministic disjoint "
            "assignment of the seven overlapping EPJ fields. Save within-field sampling "
            "probabilities and weights. Before each run, select one central target by the frozen "
            "hierarchy and encode its associated claim. No outcome-informed target replacement "
            "is permitted. Planning precision in the supplement is not empirical evidence.",
        ),
        (
            "2.3 Independent reconstruction intervention",
            "Use three clean, fixed slots per paper with the same target, package and permissions "
            "and equivalent model and resource settings. Permitted materials are the publication, "
            "methodological supplement, vetted inputs and general documentation. Original "
            "implementations and other runs are prohibited. The broker must restrict model-side "
            "tools and execution networking. Separate VMs alone are insufficient. Prior model "
            "training exposure cannot be certified absent. No qualified harness currently exists.",
        ),
        (
            "2.4 Outcomes and comparison",
            "L1 specification, L2 implementation, L3 execution, L4 numerical reproduction and L5 "
            "conclusion preservation remain separate. A run succeeds only when an independent "
            "implementation executes and meets the frozen target-specific numerical rule. Majority "
            "paper success requires two clean successes among three fixed slots. Missing or "
            "contaminated slots remain unresolved, with bounds rather than reduced denominators. "
            "Exact counts, displayed-precision scalar comparisons and target-specific stochastic "
            "rules avoid arbitrary universal tolerances. Matching numbers do not by themselves "
            "prove implementation fidelity.",
        ),
        (
            "2.5 Statistical analysis and failure attribution",
            "Use design-weighted finite-corpus estimates and separately reported uncertainty and "
            "unknown-status bounds. Fixed-frame census counts do not have paper-sampling error. "
            "Report complete-triple consistency, hierarchical outcomes, failures and burden. "
            "Do not fit unjustified multivariable models or interpret code-sharing associations "
            "causally. Distinguish observed reconstruction failure from demonstrated publication "
            "inadequacy. The proposed weighted interval is not yet implemented; details and "
            "prespecified sensitivities are in the draft SAP.",
        ),
        (
            "2.6 Freeze, reveal and human validation",
            "Seal targets before runs and all outcomes before any original-code reveal. Reveal is "
            "descriptive only; no perturbation or revised score is allowed. Prepare approximately "
            "15 outcome-balanced human cases after agent freeze, with validators masked to agent "
            "and reveal results. Actual human work, institutional determination and participation "
            "records are indispensable. No human outcome is simulated. No robustness or multiverse "
            "analysis is part of Paper II.",
        ),
    ):
        heading(document, title, 2)
        paragraph(document, text)
    heading(document, "3 Observed preparation findings")
    paragraph(
        document,
        f"The two pinned source CSVs contain {values['source_records']:,} rows and "
        f"{values['unique_pmids']:,} unique PMIDs, with "
        f"{values['excess_records']:,} excess records over a unique-paper count. "
        f"There are {values['duplicated_pmids']:,} repeated identifiers: "
        f"{values['double_occurrence_pmids']:,} occur twice and "
        f"{values['triple_occurrence_pmids']:,} occurs three times. "
        f"The deposit has {values['unique_pmid_field_pairs']:,} distinct PMID-field "
        f"memberships, including {values['within_field_excess_records']:,} within-field "
        f"excess records and {values['multi_field_pmids']:,} PMIDs spanning fields. "
        "Non-field extracted values agree across repeated records. Table 1 presents "
        "source-row and within-field identity counts; it is not an eligibility or "
        "reconstruction-results table.",
    )
    caption(document, "Table 1. Observed deposited corpus structure.")
    table(
        document,
        characteristics,
        [
            ("field", "Recorded EPJ field"),
            ("source_rows", "Source rows"),
            ("unique_pmids_within_field", "Unique PMIDs within field"),
        ],
    )
    paragraph(
        document,
        "The original sampling code caps annual candidate retrieval and its nominal "
        "random-offset expression is always zero. Historical candidate/API snapshots and "
        "validation annotations were not recovered. Consequently original population-wide "
        "selection probabilities and extraction accuracy are not established by this "
        "audit. Main-study outcomes below are delegated mechanical adjudications that await "
        "human adjudication; no human-validated reconstruction rate exists.",
    )
    heading(document, "4 Main-study execution and delegated mechanical results")
    for text in main_study_paragraphs(analysis):
        paragraph(document, text)
    heading(document, "5 Verification-adjusted sensitivity layer (A/B/C/E; D excluded)")
    for text in verification_paragraphs(audit):
        paragraph(document, text)
    paragraph(
        document,
        "Tables 6–11 are supplied separately in the editable tables file and the supplement.",
    )
    heading(document, "6 Human validation (Section D) — pending")
    paragraph(
        document,
        "No human validator has been recruited, no institutional ethics determination has been "
        "obtained and no human case has been scored. The Results text, table and figure for "
        "Section D are intentionally empty (Table 9). The AI-assisted review in Section 5 is not "
        "a substitute for human validation.",
    )
    heading(document, "7 Discussion boundary and submission hold")
    paragraph(
        document,
        "The primary empirical message is the large pre-reconstruction attrition: 90 of 100 "
        "sampled papers stopped at access, input or specification barriers before any blinded "
        "reconstruction under the frozen rules, and verification supported most of these "
        "barriers. The 10-paper reconstruction result is conditional on attempt, obtained under "
        "frozen resource ceilings and a fixed agent, and is reported side by side as a frozen "
        "mechanical value and a provisional verification-adjusted value. Non-executed papers "
        "are pre-reconstruction non-executions, not reconstruction failures. Verification "
        "identified a small number of pre-execution classification problems, chiefly leakage "
        "false positives from whole-token matching in compressed or binary content; the "
        "detector's conservatism is a methodological limitation and finding. Agent failure does "
        "not prove human impossibility. Original-code availability may correlate with reporting "
        "practice but cannot be interpreted causally here.",
    )
    heading(document, "7.1 Dynamic public datasets and input-state identifiability", 2)
    paragraph(
        document,
        f"{KEY_STATEMENT} Externally maintained public datasets, registries, genomic and "
        "administrative databases and APIs may be updated, corrected, reannotated or "
        "restructured, and query or API responses may change over time, while retaining the "
        "same repository, accession, endpoint or dataset name (Klump et al., 2021; Rauber et "
        "al., 2016; Rauber et al., 2021). The same repository, accession or URL therefore does "
        "not necessarily denote the same version, and the same version label does not by "
        "itself establish the same bytes. This does not imply that all same-accession "
        "resources change; it means that public availability does not by itself establish "
        "that the exact historical analytical input remains identifiable or retrievable. Three "
        "abilities are distinct: the ability to redistribute an input, which authors may lack "
        "for third-party data under licence or legal constraints; the ability to identify the "
        "exact state that was used; and the ability to retrieve that historical state later. "
        "A content hash verifies byte-level identity when the bytes are available (Di Cosmo et "
        "al., 2020) but does not recover historical bytes that are no longer archived. Where "
        "direct archival redistribution is not permitted, reproducibility must rely on "
        "sufficient state identification and provenance: repository or accession, release or "
        "version, retrieval timestamp, query parameters, file identity, checksum and "
        "transformation provenance (Wilkinson et al., 2016; Data Citation Synthesis Group, "
        "2014; Pasquier et al., 2017; Sandve et al., 2013; Stodden et al., 2016).",
    )
    paragraph(
        document,
        "Software-version drift, the concern of the EPJ antecedent (Onishi and Ikenoue, 2026), "
        "and dataset-version drift are analogous but distinct: both can separate a nominally "
        "available resource from the state actually used, but versioned software can usually "
        "be redistributed or rebuilt, whereas third-party data states may be neither "
        "redistributable nor recoverable. Paper II did not quantify dataset drift and makes no "
        "empirical claim about its frequency. The frozen records show only that repository or "
        "accession identity was recorded for every sampled paper, that the delegated specifier "
        "named external public reference resources for about half of the specified papers "
        "with a minority of those entries lacking a version in the publication text, that 28 "
        "papers stopped at the frozen barrier requiring an unserved external public reference "
        "resource (a state that the frozen protocol defined and this refinement leaves "
        "unchanged), and that the served inputs were logged with retrieval timestamps and "
        "SHA-256 digests but without repository version identifiers (Table 11). Reference "
        "resource identity can matter for results (Zhao and Zhang, 2015). Input-state "
        "identifiability is therefore presented as a conceptual implication, a reporting "
        "requirement and a limitation relevant to reconstructability, not as a measured "
        "outcome. A focused bibliographic search found no established use of the terms "
        "input-state identifiability or historical-state retrievability; they are used "
        "provisionally and no novelty is claimed. Data-state reproducibility was not adopted "
        "because no established compatible usage was found.",
    )
    paragraph(
        document,
        "Proposed minimum reporting set for third-party or public input data (a recommendation, "
        "not an empirically validated standard), where available: "
        + "; ".join(f"({i}) {item}" for i, item in enumerate(MINIMUM_REPORTING, 1))
        + ". Table 10 assigns each element to the framework stage it serves.",
    )
    paragraph(
        document,
        "Submission is on hold: investigator verification and sign-off of the A/B/C/E "
        "proposals, human adjudication of the 30 sealed slots and human validation (Section D) "
        "are incomplete. This document must not be submitted as a completed empirical study.",
    )
    heading(document, "Statements and declarations — pending author completion")
    paragraph(
        document,
        "Funding, competing interests, author contributions, affiliations and correspondence: "
        "not provided. Human-validation ethics/consent determination: not completed. "
        "Data/code: the public EPJ-linked repository contains the source and preparation "
        "pipeline; no Paper II reconstruction dataset exists. Third-party full text is "
        "retained internally under its applicable rights and is not redistributed.",
    )
    paragraph(
        document,
        "AI assistance: Devin performed source, literature and design audits and generated "
        "preparation code and draft materials. These preparation sessions are not blind "
        "reconstruction replicates. Human investigators must validate the design, factual "
        "claims, interpretation and final manuscript, and retain scientific responsibility.",
    )
    heading(document, "References used in this preparation draft")
    paragraph(
        document,
        "Antecedent author names follow the supplied author specification. Publisher/"
        "Crossref metadata reverse the first author's name components; verify the final "
        "bibliographic form before submission. Preprints and version-specific claims "
        "are identified in the reference ledger.",
    )
    chosen = [row for row in references if row["reference_id"] in REFERENCES]
    for row in sorted(chosen, key=reference_text):
        paragraph(document, reference_text(row))
    document.save(str(output / "Preparation_draft_EN.docx"))

    supplement = Document()
    style_document(supplement)
    heading(supplement, "Paper II — draft protocols and preparation supplement", 0)
    paragraph(supplement, STATUS)
    paragraph(
        supplement,
        "Table S1 gives analytic precision planning at assumed proportions. These values "
        "are not observed outcomes and do not replace design-based intervals.",
    )
    caption(supplement, "Table S1. Wilson planning intervals (unweighted binomial reference).")
    table(
        supplement,
        precision,
        [
            ("planned_n", "Planned n"),
            ("assumed_p", "Assumed p"),
            ("wilson_lower", "95% lower"),
            ("wilson_upper", "95% upper"),
            ("half_width", "Half-width"),
        ],
    )
    main_study_tables(supplement, dispositions, run_levels, failure_codes, reveal)
    verification_tables(supplement, verification_dir)
    for path in sorted((ROOT / "protocols").glob("*.md")):
        supplement.add_page_break()
        markdown(supplement, path.read_text())
    supplement.save(str(output / "Supplement_protocols_DRAFT_EN.docx"))

    tables = Document()
    style_document(tables)
    heading(tables, "Editable preparation tables", 0)
    paragraph(tables, "Table 1 is the observed source audit; Table S1 is analytic planning only.")
    caption(tables, "Table 1. Observed deposited corpus structure.")
    table(
        tables,
        characteristics,
        [
            ("field", "Recorded EPJ field"),
            ("source_rows", "Source rows"),
            ("unique_pmids_within_field", "Unique PMIDs within field"),
        ],
    )
    main_study_tables(tables, dispositions, run_levels, failure_codes, reveal)
    verification_tables(tables, verification_dir)
    caption(tables, "Table S1. Analytic Wilson precision planning.")
    table(
        tables,
        precision,
        [
            ("planned_n", "Planned n"),
            ("assumed_p", "Assumed p"),
            ("wilson_lower", "95% lower"),
            ("wilson_upper", "95% upper"),
            ("half_width", "Half-width"),
        ],
    )
    tables.save(str(output / "Editable_tables_EN.docx"))
    for name in ARTIFACT_NAMES:
        if name.endswith((".docx", ".pptx")):
            normalize_office_archive(output / name)
    audit = (
        "# Repository audit — observed source only\n\n"
        f"Source records: {values['source_records']:,}; unique PMIDs: "
        f"{values['unique_pmids']:,}; excess records: {values['excess_records']:,}.\n\n"
        "The VOR links this deposited corpus. No corrected corpus or historical validation "
        "annotations were recovered in the audited public sources. Sampling candidates and "
        "original API snapshots are unavailable; historical flags are text-detection proxies. "
        "No replacement sample, blind reconstruction or human validation has occurred.\n\n"
        "The author approved the unique-PMID inference frame while preserving all source rows; "
        "data/frame_decision.json and results/frame_manifest.json record its provenance.\n\n"
        "See results/manuscript_values.csv for machine-readable numerical provenance, "
        "data/acquisition_ledger.csv for exact source hashes, review/EVIDENCE.md for retained "
        "audit evidence, and results/readiness.json for incomplete study gates.\n"
    )
    (ROOT / "review/REPOSITORY_AUDIT.md").write_text(audit)
    files = [
        *(output / name for name in ARTIFACT_NAMES),
        *sorted((ROOT / "protocols").glob("*.md")),
        *sorted((ROOT / "results").glob("*.csv")),
        *sorted((ROOT / "results").glob("*.json")),
        *sorted((ROOT / "results/verification_adjusted").glob("*")),
        *sorted((ROOT / "data/adjudication/main-study-20260925/verification_adjusted").glob("*")),
        ROOT / "WITHOUT_D_STATUS.md",
        *sorted((ROOT / "src/paper2").glob("*.py")),
        *sorted((ROOT / "data/derived").glob("*.csv")),
        *sorted((ROOT / "data").glob("*.csv")),
        *sorted((ROOT / "data").glob("*.json")),
        *sorted((ROOT / "schemas").glob("*")),
        ROOT / "requirements.lock",
        ROOT / "requirements-science.in",
        ROOT / "requirements-science.lock",
        ROOT / "Dockerfile.science",
        ROOT / ".dockerignore",
        ROOT / "pyproject.toml",
        ROOT / "README.md",
        ROOT / "Makefile",
        *sorted((ROOT / "review").glob("*.md")),
        *sorted((ROOT / "tests").glob("*.py")),
    ]
    manifest = {
        str(path.relative_to(ROOT)): {
            "bytes": path.stat().st_size,
            "sha256": sha256(path.read_bytes()),
        }
        for path in files
        if path.is_file()
        and path.name not in {"preparation_artifact_manifest.json", "VERIFICATION_ADJUSTED_QC.md"}
    }
    (ROOT / "results/preparation_artifact_manifest.json").write_text(
        json.dumps(manifest, indent=2) + "\n",
    )
    dist = ROOT / "dist"
    dist.mkdir(exist_ok=True)
    with zipfile.ZipFile(
        dist / "PaperII_preparation_NOT_SUBMISSION_READY.zip", "w", compression=zipfile.ZIP_DEFLATED
    ) as archive:
        for path in sorted({*files, ROOT / "results/preparation_artifact_manifest.json"}):
            if path.is_file():
                zip_member(archive, str(path.relative_to(ROOT)), path.read_bytes())


if __name__ == "__main__":
    build_documents()
