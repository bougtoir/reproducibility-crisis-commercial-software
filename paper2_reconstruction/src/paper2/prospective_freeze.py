"""Prospective ten-field freeze and delegated investigator verification for the pilot.

For each of the seven prospective-pilot papers this module fixes, before any
reconstruction slot starts: the exact target, its location in the publication,
the required inputs with byte-level identity, the allowed resources, the blocked
original implementation artifacts, the comparison metric, the numerical agreement
criterion, the conclusion criterion, the resource ceiling and the stopping rule.

Every served input is an actual file held in the persistent evidence area; the
record stores its size and SHA-256 so the served package can be checked against
this freeze. Gate reviews are a delegated second pass by the research assistant
(Devin) over the retained article text; ``investigator_verification`` stays
``pending`` until a human adjudicator records a decision. Original study authors
are not contacted and their implementation artifacts are never served.
"""

from __future__ import annotations

import argparse
import json
from dataclasses import asdict, dataclass
from datetime import datetime, timezone
from pathlib import Path

from paper2.build import ROOT
from paper2.core import sha256, sha256_file, snapshot
from paper2.model_api import mapping
from paper2.primary_adjudication import article_segments
from paper2.timestamp import stamp

FREEZE_ID = "PILOT-FREEZE-2026-09-25"
AMENDMENT = "AMEND-2026-09-25-04"
EVIDENCE = Path("/home/ubuntu/paper2_evidence/pilot-inputs-20260925")
LEGACY = Path("/home/ubuntu/paper2_evidence/deposit-inputs-20260924")
SCREEN = ROOT / "data" / "raw" / "deposit-screen-20260924" / "screen"
ADJUDICATION = ROOT / "data" / "adjudication" / "devin_amended_G1_G5_20260924.json"
RECORD = ROOT / "data" / "adjudication" / "pilot_prospective_freeze_20260925.json"
STAMP_DIR = ROOT / "data" / "raw" / "pilot-freeze-20260925"
IMAGE = "paper2-science:pilot-2026-09-25"
GATES = ("G1", "G2", "G3", "G4", "G5")
BANNED_WORDING = "author verification"

CEILING: dict[str, object] = {
    "slots_per_paper": 3,
    "wall_seconds_per_slot": 14400,
    "step_seconds": 3600,
    "provider_tokens_per_slot": 400000,
    "tool_calls_per_slot": 40,
    "max_completion_tokens": 8192,
    "worker_cpus": 1,
    "worker_memory": "3g",
    "worker_work_storage": "disk-backed /work, exported artifacts capped at 64 MiB",
    "network": "none",
    "image": IMAGE,
}
STOPPING_RULE = (
    "final report submitted by the slot",
    "wall, step, token, tool-call or output ceiling reached (recorded as the stop reason)",
    "required input absent from the served package: report null observation with F01/F18",
    "specification gap that cannot be closed from the served text: report null with F04-F08",
    "controller or worker error: recorded as harness stop, never as an outcome",
    "no rerun, no extension and no rule change after the slot starts",
)
CONCLUSION_RULE = (
    "Conclusion agreement is scored only when the numerical criterion is met or when the "
    "publication states a directional claim that the observed value can be compared with; "
    "otherwise the conclusion component is not_assessable."
)


@dataclass(frozen=True)
class Served:
    served_as: str
    source: Path
    identifier: str
    role: str


@dataclass(frozen=True)
class Blocked:
    artifact: str
    reason: str


@dataclass(frozen=True)
class GateReview:
    gate: str
    delegated_assessment: str
    second_pass: str
    rationale: str


@dataclass(frozen=True)
class Quote:
    segment_id: str
    quote: str


@dataclass(frozen=True)
class Specification:
    paper_id: str
    stratum: str
    exact_target: str
    blind_target: str
    target_type: str
    target_location: str
    location_evidence: tuple[Quote, ...]
    required_inputs: tuple[Served, ...]
    input_notes: tuple[str, ...]
    allowed_resources: tuple[str, ...]
    blocked_original_artifacts: tuple[Blocked, ...]
    comparison_metric: str
    numerical_agreement_criterion: str
    conclusion_criterion: str
    specification_gaps: tuple[str, ...]
    gate_reviews: tuple[GateReview, ...]


def receipt_body(directory: Path, sha256: str | None = None) -> Path:
    bodies = sorted(directory.glob("*/body"))
    if sha256 is not None:
        bodies = [
            body
            for body in bodies
            if mapping(json.loads((body.parent / "receipt.json").read_text()))["sha256"] == sha256
        ]
    if len(bodies) != 1:
        raise ValueError(f"expected exactly one acquired body under {directory}")
    return bodies[0]


def ena_runs(accession: str, platform: str | None = None) -> tuple[Served, ...]:
    ledger_path = EVIDENCE / "files" / f"ena_{accession}" / f"{accession}_fetch_ledger.json"
    ledger = mapping(json.loads(ledger_path.read_text()))
    files = ledger["files"]
    if not isinstance(files, list):
        raise ValueError(f"{accession}: fetch ledger has no file list")
    alias_lines = (EVIDENCE / "ena" / f"{accession}_alias.tsv").read_text().splitlines()
    header = alias_lines[0].split("\t")
    rows = [dict(zip(header, line.split("\t"), strict=True)) for line in alias_lines[1:]]
    by_run = {row["run_accession"]: row for row in rows}
    served: list[Served] = []
    for entry in files:
        item = mapping(entry)
        if item["completeness"] != "complete_response":
            raise ValueError(f"{accession} {item['run_accession']}: download incomplete")
        run = str(item["run_accession"])
        row = by_run[run]
        if platform is not None and row["instrument_platform"] != platform:
            continue
        name = str(item["file"])
        body = receipt_body(EVIDENCE / "files" / f"ena_{accession}" / run, str(item["sha256"]))
        served.append(
            Served(
                f"reads/{row['instrument_platform'].lower()}/{name}",
                body,
                f"ena:{accession} {run} library {row['library_name']} {name}",
                "deposited_input",
            )
        )
    if not served:
        raise ValueError(f"{accession}: no complete runs for platform {platform}")
    return tuple(served)


def reference(subdir: str, served_as: str, identifier: str, url_fragment: str) -> Served:
    for receipt in sorted((EVIDENCE / "references" / subdir).glob("*/receipt.json")):
        payload = mapping(json.loads(receipt.read_text()))
        if url_fragment in str(payload["url"]):
            if payload["completeness"] != "complete_response":
                raise ValueError(f"{identifier}: reference download incomplete")
            return Served(served_as, receipt.with_name("body"), identifier, "public_reference")
    raise ValueError(f"{identifier}: no receipt matching {url_fragment}")


def legacy(body: Path, served_as: str, identifier: str) -> Served:
    return Served(served_as, body, identifier, "deposited_input")


def dryad_browser(deposit: str, name: str, identifier: str) -> Served:
    return Served(
        f"deposit/{name}",
        EVIDENCE / "files" / f"dryad_{deposit}_browser" / name,
        identifier,
        "deposited_input",
    )


def review(gate: str, delegated: str, second_pass: str, rationale: str) -> GateReview:
    return GateReview(gate, delegated, second_pass, rationale)


def specifications() -> tuple[Specification, ...]:
    return (
        Specification(
            paper_id="PMID:33016314",
            stratum="Biomedical_Basic",
            exact_target=(
                "Number of unique OTUs in the filtered data set: 2763 bacterial (16S rRNA, "
                "Illumina MiSeq) and 357 fungal (ITS, PacBio Sequel II)"
            ),
            blind_target=(
                "The number of unique bacterial (16S rRNA) OTUs and the number of unique "
                "fungal (ITS) OTUs in the filtered data set, as reported in the Results"
            ),
            target_type="count",
            target_location="Results, first paragraph of the sequencing-data description",
            location_evidence=(
                Quote(
                    "S0067",
                    "which were clustered into 2763 unique bacterial OTUs and 357 unique "
                    "fungal OTUs",
                ),
                Quote(
                    "S0057",
                    "threshold of 8) at 97% sequence similarity with unoise3 as part of "
                    "usearch v11.0.667. Singletons were removed (abundance threshold 2)",
                ),
            ),
            required_inputs=(
                *ena_runs("PRJNA641521"),
                Served(
                    "metadata/PRJNA641521_run_library_alias.tsv",
                    EVIDENCE / "ena" / "PRJNA641521_alias.tsv",
                    "ena:PRJNA641521 run/library alias report (library CH1-CH96 per run)",
                    "public_metadata",
                ),
                Served(
                    "metadata/PRJNA641521_sample_metadata.tsv",
                    EVIDENCE / "ena" / "PRJNA641521_meta.tsv",
                    "ena:PRJNA641521 sample/instrument metadata report",
                    "public_metadata",
                ),
                reference(
                    "silva_128",
                    "references/SILVA_128_SSURef_Nr99_tax_silva.fasta.gz",
                    "SILVA release 128 SSURef_Nr99_tax_silva "
                    "(md5 9f69e7583fca9b7900837279a3ad20ab)",
                    "SILVA_128_SSURef_Nr99_tax_silva.fasta.gz",
                ),
                reference(
                    "unite_7.2",
                    "references/unite_v7.2_sh_general_release_01.12.2017.zip",
                    "UNITE general FASTA release 7.2 (doi:10.15156/BIO/587475) main file",
                    "3a13b7f6-bedd-4cb2-8d29-752b5addeaae",
                ),
                reference(
                    "unite_7.2",
                    "references/unite_v7.2_sh_general_release_s_01.12.2017.zip",
                    "UNITE general FASTA release 7.2 (doi:10.15156/BIO/587475) singletons file",
                    "354807ca-e58e-49b7-9e65-8ba35158a320",
                ),
            ),
            input_notes=(
                "192 ENA runs: 96 Illumina MiSeq (bacteria) and 96 PacBio Sequel II (fungi), "
                "one FASTQ per run; ENA reports library_layout SINGLE for the MiSeq runs "
                "although the article describes merging of paired reads.",
                "Reads are deposited per library (CH1-CH96), so the proprietary SMRT Link "
                "demultiplexing step named in the article precedes the served inputs.",
            ),
            allowed_resources=(
                "vsearch 2.22.1 (free reimplementation of the usearch/unoise3 algorithms)",
                "cutadapt, fastp, seqtk, Python scientific stack in the qualified image",
                "SILVA 128 and UNITE 7.2 reference files served in the package",
            ),
            blocked_original_artifacts=(
                Blocked("usearch v11.0.667 binary", "proprietary freeware; not redistributable"),
                Blocked(
                    "PacBio SMRT Link 6.0.0.47841", "commercial software; step precedes deposit"
                ),
                Blocked("any author script or OTU table", "author_implementation_material"),
            ),
            comparison_metric="exact integer comparison of the two OTU counts",
            numerical_agreement_criterion=(
                "both counts equal (2763 and 357); each component recorded separately"
            ),
            conclusion_criterion=CONCLUSION_RULE,
            specification_gaps=(
                "read type of the PacBio deposit (subreads vs CCS) not stated",
                "MiSeq deposit layout (single) versus described paired-read merging",
                "unoise3 is specified through usearch defaults not printed in the text",
            ),
            gate_reviews=(
                review("G1", "yes", "confirmed", "OTU table is a specified computation on reads"),
                review("G2", "yes", "confirmed", "two integer counts stated in S0067"),
                review(
                    "G3",
                    "public",
                    "confirmed",
                    "all 192 FASTQ files downloaded anonymously and checksum-verified",
                ),
                review(
                    "G4",
                    "available",
                    "confirmed_with_note",
                    "usearch binary is proprietary freeware; vsearch is the free equivalent; "
                    "SMRT Link demultiplexing precedes the deposited per-library reads",
                ),
                review(
                    "G5",
                    "uncertain",
                    "confirmed",
                    "layout and read-type ambiguities recorded as specification gaps",
                ),
            ),
        ),
        Specification(
            paper_id="PMID:35003117",
            stratum="Chemistry_Materials",
            exact_target=(
                "Number of differentially expressed genes between 3D- and 2D-cultured "
                "RAW264.7 macrophages: 6762 total, 5949 down-regulated, 813 up-regulated"
            ),
            blind_target=(
                "The number of differentially expressed genes between 3D- and 2D-cultured "
                "macrophages (total, down-regulated and up-regulated) under the published "
                "thresholds"
            ),
            target_type="count",
            target_location="Results, transcriptome paragraph",
            location_evidence=(
                Quote(
                    "S0050",
                    "among which 6762 were differentially expressed (fold change > 2, p = "
                    "0.05, p adj = 0.05) in both males and females. A total of 5949 genes "
                    "were downregulated and 813 were upregulated.",
                ),
                Quote(
                    "S0024", "DEGs were screened using Degseq, with |log2 fold change| ≥ 1 and q"
                ),
            ),
            required_inputs=(
                legacy(
                    LEGACY / "PMID_35003117" / "x" / "GSM5662247_2D_count.txt.gz",
                    "deposit/GSM5662247_2D_count.txt.gz",
                    "geo:GSE186841 GSM5662247 2D gene counts",
                ),
                legacy(
                    LEGACY / "PMID_35003117" / "x" / "GSM5662248_3D_count.txt.gz",
                    "deposit/GSM5662248_3D_count.txt.gz",
                    "geo:GSE186841 GSM5662248 3D gene counts",
                ),
            ),
            input_notes=(
                "The two count tables are the members of GSE186841_RAW.tar (sha256 "
                "7584208757ee6522657de24b54b5c3f6ee3064a2af5c5ab4225fe74a960e4649).",
            ),
            allowed_resources=(
                "R 4.2.2 with DEGseq 1.52.0, edgeR and DESeq2 in the qualified image",
                "Python scientific stack in the qualified image",
            ),
            blocked_original_artifacts=(
                Blocked("author DEG lists or scripts", "author_implementation_material"),
            ),
            comparison_metric="exact integer comparison of total, down and up counts",
            numerical_agreement_criterion="all three counts equal; components recorded separately",
            conclusion_criterion=CONCLUSION_RULE,
            specification_gaps=(
                "DEGseq method variant and normalisation not stated",
                "'in both males and females' with one count column per condition",
                "methods mix RNA-seq with an Affymetrix array description",
            ),
            gate_reviews=(
                review("G1", "yes", "confirmed", "DEG screen on deposited counts"),
                review("G2", "yes", "confirmed", "three counts stated in S0050"),
                review("G3", "public", "confirmed", "GEO archive retrieved anonymously"),
                review(
                    "G4", "available", "confirmed", "DEGseq 1.52.0 installed from Bioconductor 3.16"
                ),
                review("G5", "uncertain", "confirmed", "gaps recorded; attempt possible"),
            ),
        ),
        Specification(
            paper_id="PMID:41444829",
            stratum="Clinical_Medicine",
            exact_target=(
                "Held-out test-set accuracy (98%) and ROC AUC (0.98) of the gradient-boosting "
                "tumour-detection model"
            ),
            blind_target=(
                "Held-out test-set accuracy and ROC AUC of the published gradient-boosting "
                "tumour-detection model"
            ),
            target_type="prediction_performance",
            target_location="Results, model performance paragraph (Fig. 1e)",
            location_evidence=(
                Quote(
                    "S0036",
                    "an accuracy of 98%, an 88% sensitivity at 95% specificity, and an area "
                    "under the curve (AUC) of the receiver operating characteristic (ROC) "
                    "curve of 0.98",
                ),
                Quote(
                    "S0174",
                    "We found that 58 of the 126 features had significant differences between "
                    "means ( P < 0.10)",
                ),
            ),
            required_inputs=(
                Served(
                    "deposit/Experiments_raw_txt_only.zip",
                    EVIDENCE / "served_packages" / "Experiments_raw_txt_only.zip",
                    "zenodo:17343533 Experiments.zip raw per-well .txt spectra only "
                    "(custodian-filtered; member manifest retained)",
                    "custodian_filtered_deposit",
                ),
                Served(
                    "deposit/Experiments_raw_txt_only.manifest.json",
                    EVIDENCE / "served_packages" / "Experiments_raw_txt_only.manifest.json",
                    "member manifest of the filtered archive with source archive sha256",
                    "custodian_filtered_deposit",
                ),
                Served(
                    "deposit/Labels.zip",
                    receipt_body(EVIDENCE / "files" / "zenodo_labels_zip"),
                    "zenodo:17343533 Labels.zip",
                    "deposited_input",
                ),
                Served(
                    "deposit/Corrections.zip",
                    receipt_body(EVIDENCE / "files" / "zenodo_corrections_zip"),
                    "zenodo:17343533 Corrections.zip",
                    "deposited_input",
                ),
            ),
            input_notes=(
                "Experiments.zip (sha256 "
                "7f72f86d11a38c9b06c39ca3b9b3fbceba267373c3300fe71b6f3eec6efe9a1d, "
                "md5 verified) contains raw per-well .txt spectra plus author-processed .mat, "
                "'with Fits' .csv and .png files; only the .txt members are served byte-identical.",
            ),
            allowed_resources=(
                "scikit-learn, hyperopt-equivalent Bayesian optimisation or any free optimiser, "
                "Python scientific stack in the qualified image",
            ),
            blocked_original_artifacts=(
                Blocked("Scripts.zip", "author_implementation_material"),
                Blocked(
                    ".mat, 'with Fits' .csv and .png members of Experiments.zip",
                    "author-processed intermediates derived by the withheld implementation",
                ),
            ),
            comparison_metric=(
                "accuracy and ROC AUC on the held-out test set as defined by the publication"
            ),
            numerical_agreement_criterion=(
                "accuracy rounds to 98% and AUC rounds to 0.98 under nearest rounding; the "
                "split identity must match the publication or the component is not_assessable"
            ),
            conclusion_criterion=CONCLUSION_RULE,
            specification_gaps=(
                "train/test split membership and random seed sit in withheld scripts",
                "hyperparameter search space in Supplementary Table 3 not retained",
                "spectral preprocessing from raw .txt to the 126 QWN features not fully printed",
            ),
            gate_reviews=(
                review("G1", "yes", "confirmed", "ML performance on deposited spectra"),
                review("G2", "yes", "confirmed", "accuracy and AUC stated in S0036"),
                review(
                    "G3",
                    "public",
                    "confirmed",
                    "all three data archives downloaded; Experiments.zip md5 verified",
                ),
                review("G4", "available", "confirmed", "scikit-learn stack is free"),
                review(
                    "G5",
                    "uncertain",
                    "confirmed",
                    "split identity likely unrecoverable; prediction rule requires identical split",
                ),
            ),
        ),
        Specification(
            paper_id="PMID:36286480",
            stratum="Computational_Science",
            exact_target=(
                "Coefficient of determination r^2 = 0.817 between estimated and expected lineage "
                "frequencies across the seven SARS-CoV-2 mixture standards"
            ),
            blind_target=(
                "The coefficient of determination between estimated and expected lineage "
                "frequencies across the seven SARS-CoV-2 mixture samples"
            ),
            target_type="deterministic_scalar",
            target_location="Results, SARS-CoV-2 mixture validation paragraph (Fig. S1A)",
            location_evidence=(
                Quote(
                    "S0027",
                    "MixviR’s estimates of the frequencies of each lineage were also broadly "
                    "consistent with expectations (see Fig. S1A in the supplemental material; "
                    "r 2 = 0.817)",
                ),
                Quote(
                    "S0054",
                    "The presence of lineage l in sample i is inferred by comparing the ratio "
                    "n li /N l to an adjustable threshold value (default = 0.5)",
                ),
            ),
            required_inputs=(
                *ena_runs("PRJNA827817"),
                Served(
                    "metadata/PRJNA827817_sample_metadata.tsv",
                    EVIDENCE / "ena" / "PRJNA827817_meta.tsv",
                    "ena:PRJNA827817 sample metadata with expected lineage ratios",
                    "public_metadata",
                ),
                reference(
                    "sarscov2_reference",
                    "references/GCF_009858895.2_ASM985889v3_genomic.fna.gz",
                    "NCBI GCF_009858895.2 SARS-CoV-2 Wuhan-Hu-1 genome",
                    "genomic.fna.gz",
                ),
                reference(
                    "sarscov2_reference",
                    "references/GCF_009858895.2_ASM985889v3_genomic.gff.gz",
                    "NCBI GCF_009858895.2 annotation",
                    "genomic.gff.gz",
                ),
                reference(
                    "constellations_v0.1.3",
                    "references/constellations-v0.1.3.tar.gz",
                    "cov-lineages/constellations v0.1.3 (2022-02-08) lineage-defining sites",
                    "v0.1.3",
                ),
                reference(
                    "artic_v4_1",
                    "references/artic_nCoV-2019_V4.1_primer.bed",
                    "ARTIC nCoV-2019 V4.1 primer scheme",
                    "V4.1",
                ),
            ),
            input_notes=(
                "Expected lineage ratios are recorded in the public ENA sample descriptions "
                "(design inputs, not deposit leakage).",
                "The article's mutation list is an outbreak.info snapshot of 2022-02-17 kept in "
                "the author repository; the served substitute is the independent "
                "cov-lineages/constellations release closest before that date. Version "
                "mismatch is a frozen, recorded input substitution (F18 risk).",
                "The amplicon primer scheme is not stated in the retained text; the ARTIC V4.1 "
                "bed is provided as a public reference whose use is a recorded assumption.",
            ),
            allowed_resources=(
                "minimap2, samtools, bcftools, ivar, fastp, cutadapt in the qualified image",
                "Python and R scientific stacks in the qualified image",
            ),
            blocked_original_artifacts=(
                Blocked("MixviR R package and GitHub repository", "author_implementation_material"),
                Blocked("DRAGEN COVID Lineage v3.5.10", "commercial cloud application"),
                Blocked("author mutation_files snapshot", "author_implementation_material"),
            ),
            comparison_metric="r^2 across all lineage-frequency pairs of the seven mixtures",
            numerical_agreement_criterion="observed r^2 rounds to 0.817 at three decimals",
            conclusion_criterion=CONCLUSION_RULE,
            specification_gaps=(
                "variant-calling parameters other than the DRAGEN thresholds (=1) not stated",
                "definition of the pairs entering r^2 (absent lineages as zero or excluded)",
                "mutation list version differs from the author snapshot by construction",
            ),
            gate_reviews=(
                review("G1", "yes", "confirmed", "validation statistic from known mixtures"),
                review("G2", "yes", "confirmed", "r^2 = 0.817 stated in S0027"),
                review(
                    "G3",
                    "unknown",
                    "revised:public_with_substituted_reference",
                    "all seven paired FASTQ sets downloaded and checksum-verified; the "
                    "lineage-mutation list is available only as a substitute public release",
                ),
                review(
                    "G4",
                    "available",
                    "confirmed",
                    "article names BCFtools/GATK as accepted VCF generators",
                ),
                review("G5", "uncertain", "confirmed", "estimator rule stated in S0054-S0055"),
            ),
        ),
        Specification(
            paper_id="PMID:35388942",
            stratum="Environmental_Earth",
            exact_target=(
                "Mean interval between active mound-rising events: 5.9 years before warming "
                "(1850-1961) and 3.0 years after warming (1961-2010)"
            ),
            blind_target=(
                "The mean interval in years between active mound-rising (tree-leaning) "
                "events before warming and after warming, as defined in the publication"
            ),
            target_type="deterministic_scalar",
            target_location="Results, section on annual activities of tree leaning (Figure 6c)",
            location_evidence=(
                Quote(
                    "S0049",
                    "The interval between active mound‐rising events was shorter after warming "
                    "(3.0 years) than the interval before warming (5.9 years; Figure 6c )",
                ),
                Quote(
                    "S0033",
                    "Peaks of annual intensity of tree leaning >2 were identified as active "
                    "tree leaning events",
                ),
            ),
            required_inputs=(
                Served(
                    "deposit/Data_S1-3.xlsx",
                    EVIDENCE / "files" / "dryad_dr7sqv9z5_Data_S1-3_browser" / "body",
                    "dryad:dr7sqv9z5 Data_S1-3.xlsx (file 1389355, version 168699)",
                    "deposited_input",
                ),
            ),
            input_notes=(
                "Data S2 holds ring widths by sample and direction; Data S3 wood radius; Data "
                "S1 is a carbon-stock summary unrelated to the target.",
            ),
            allowed_resources=("Python scientific stack (pandas, numpy, scipy, openpyxl)",),
            blocked_original_artifacts=(
                Blocked("SigmaPlot 14.0 project files", "commercial software; not deposited"),
                Blocked("author event lists", "author_implementation_material"),
            ),
            comparison_metric="two mean intervals in years",
            numerical_agreement_criterion=(
                "each interval rounds to the reported value at one decimal"
            ),
            conclusion_criterion=CONCLUSION_RULE,
            specification_gaps=(
                "event definition across trees (per tree vs pooled) and interval averaging rule",
                "text cites Data S1 for the leaning ratio while the deposit's S1 is carbon stocks",
            ),
            gate_reviews=(
                review("G1", "yes", "confirmed", "indices and event counts from ring widths"),
                review("G2", "yes", "confirmed", "5.9 and 3.0 years stated in S0049"),
                review(
                    "G3", "public", "confirmed", "Dryad file downloaded; bytes match storageSize"
                ),
                review("G4", "available", "confirmed", "free stack suffices"),
                review(
                    "G5",
                    "sufficient",
                    "revised:uncertain",
                    "interval aggregation rule is not printed; recorded as a gap",
                ),
            ),
        ),
        Specification(
            paper_id="PMID:35767539",
            stratum="Physics_Engineering",
            exact_target=(
                "Least-squares estimates of the logistic paper-count function on WoS-Stat: "
                "mu_n = 33.263, sigma_n = 14.743, kappa_n = 17242.068"
            ),
            blind_target=(
                "The least-squares parameter estimates (mu_n, sigma_n, kappa_n) of the "
                "published function f_n fitted to the annual paper counts of WoS-Stat"
            ),
            target_type="deterministic_scalar",
            target_location="Modeling of WoS-Stat network, estimation paragraph (Fig 5a)",
            location_evidence=(
                Quote(
                    "S0052",
                    "For f n , We adopt the least squares method to estimate parameters and "
                    "obtain estimates μ ^ n = 33 . 263 , σ ^ n = 14 . 743 , κ ^ n =",
                ),
                Quote("S0053", "17242 . 068 , and η ^ n = 328 . 047 ."),
            ),
            required_inputs=(
                dryad_browser("z8w9ghxfh", "wos-stat_nodes.csv", "dryad:z8w9ghxfh nodes"),
                dryad_browser("z8w9ghxfh", "wos-stat_edges.csv", "dryad:z8w9ghxfh edges"),
                dryad_browser("z8w9ghxfh", "README.md", "dryad:z8w9ghxfh README"),
            ),
            input_notes=("Publication years 1981-2016 map to t = 1..36 per the article.",),
            allowed_resources=("Python scientific stack (numpy, scipy, pandas, networkx)",),
            blocked_original_artifacts=(
                Blocked("author fitting scripts", "author_implementation_material"),
            ),
            comparison_metric="three parameter estimates",
            numerical_agreement_criterion=(
                "each estimate rounds to the reported value at three decimals; eta_n recorded "
                "as a secondary component"
            ),
            conclusion_criterion=CONCLUSION_RULE,
            specification_gaps=(
                "least-squares initialisation and optimiser unspecified (F10 risk)",
                "eta_n estimator definition not printed",
            ),
            gate_reviews=(
                review("G1", "yes", "confirmed", "parameter fit on deposited network"),
                review("G2", "yes", "confirmed", "estimates stated in S0052-S0053"),
                review("G3", "public", "confirmed", "three Dryad files downloaded and hashed"),
                review("G4", "available", "confirmed", "free stack suffices"),
                review("G5", "uncertain", "confirmed", "functional form and method stated"),
            ),
        ),
        Specification(
            paper_id="PMID:35136578",
            stratum="Social_Behavioral",
            exact_target=(
                "Odds ratio for previous exposure to colic (OR = 32.250) and body condition "
                "score (OR = 0.022) with the Table 2 significance pattern"
            ),
            blind_target=(
                "The odds ratios for previous exposure to colic and for body condition "
                "score as reported in Table 2"
            ),
            target_type="deterministic_scalar",
            target_location="Results, clinical findings and risk factors (Table 2)",
            location_evidence=(
                Quote(
                    "S0037",
                    "previous exposure to colic (p < 0.000; OR = 32.250), body condition score "
                    "(p < 0.000; OR = 0.022)",
                ),
                Quote(
                    "S0031",
                    "The final analysis was done using logistic regression analysis with "
                    "selected variables depending on p-values.",
                ),
            ),
            required_inputs=(
                Served(
                    "deposit/Risk_factor_in_horse_with_colic.xlsx",
                    receipt_body(EVIDENCE / "files" / "figshare_colic_xlsx"),
                    "figshare:15148851 file 29101587",
                    "deposited_input",
                ),
            ),
            input_notes=(
                "Sheets: Risk factor (per-horse coded factors), Clinical sign, Hematology.",
            ),
            allowed_resources=("Python scientific stack (statsmodels, scipy, pandas, openpyxl)",),
            blocked_original_artifacts=(
                Blocked("SPSS v.25 outputs or syntax", "commercial software; not deposited"),
            ),
            comparison_metric="two odds ratios",
            numerical_agreement_criterion="each OR rounds to the reported value at three decimals",
            conclusion_criterion=CONCLUSION_RULE,
            specification_gaps=(
                "univariable versus multivariable origin of Table 2 ORs",
                "variable selection rule 'depending on p-values' unspecified",
            ),
            gate_reviews=(
                review("G1", "yes", "confirmed", "regression on deposited records"),
                review("G2", "yes", "confirmed", "ORs stated in S0037"),
                review("G3", "public", "confirmed", "Figshare file downloaded anonymously"),
                review(
                    "G4", "available", "confirmed", "SPSS analyses reproducible in free software"
                ),
                review("G5", "uncertain", "confirmed", "selection rule gap recorded"),
            ),
        ),
    )


def served_record(item: Served) -> dict[str, object]:
    if not item.source.is_file():
        raise ValueError(f"served input missing on persistent storage: {item.source}")
    return {
        "served_as": item.served_as,
        "identifier": item.identifier,
        "role": item.role,
        "source_path": str(item.source.resolve()),
        "bytes": item.source.stat().st_size,
        "sha256": sha256_file(item.source),
    }


def verify_quotes(paper_id: str, quotes: tuple[Quote, ...]) -> str:
    digest, segments = article_segments(SCREEN, paper_id)
    for quote in quotes:
        if quote.quote not in segments.get(quote.segment_id, ""):
            raise ValueError(
                f"{paper_id}: quote not verbatim in {quote.segment_id}: {quote.quote!r}"
            )
    return digest


def delegated_assessments() -> dict[str, dict[str, object]]:
    record = mapping(json.loads(ADJUDICATION.read_text()))
    assessments = record["assessments"]
    if not isinstance(assessments, list):
        raise ValueError("adjudication record has no assessments")
    return {
        str(mapping(a)["paper_id"]): mapping(a)
        for a in assessments
        if mapping(a)["pilot_decision"] == "provisional_pilot_case"
    }


def check_gate_reviews(spec: Specification, delegated: dict[str, object]) -> None:
    keys = {
        "G1": "G1_computationally_testable",
        "G2": "G2_principal_target_identifiable",
        "G3": "G3_input_accessibility",
        "G4": "G4_resource_accessibility",
        "G5": "G5_specification_sufficient_for_attempt",
    }
    if tuple(r.gate for r in spec.gate_reviews) != GATES:
        raise ValueError(f"{spec.paper_id}: gate reviews must cover G1-G5 in order")
    for item in spec.gate_reviews:
        if str(delegated[keys[item.gate]]) != item.delegated_assessment:
            raise ValueError(
                f"{spec.paper_id} {item.gate}: delegated value {delegated[keys[item.gate]]!r} "
                f"differs from review {item.delegated_assessment!r}"
            )


def build() -> dict[str, object]:
    delegated = delegated_assessments()
    specs = specifications()
    if {s.paper_id for s in specs} != set(delegated):
        raise ValueError("freeze must cover exactly the provisional pilot cases")
    papers = []
    for spec in specs:
        check_gate_reviews(spec, delegated[spec.paper_id])
        digest = verify_quotes(spec.paper_id, spec.location_evidence)
        served = [served_record(item) for item in spec.required_inputs]
        names = [s["served_as"] for s in served]
        if len(set(names)) != len(names):
            raise ValueError(f"{spec.paper_id}: duplicate served path")
        papers.append(
            {
                "paper_id": spec.paper_id,
                "sampling_stratum": spec.stratum,
                "role": "prospective_pilot_pipeline_validation_only",
                "article_source_sha256": digest,
                "exact_target": spec.exact_target,
                "blind_target": spec.blind_target,
                "target_type": spec.target_type,
                "target_location": spec.target_location,
                "location_evidence": [asdict(q) for q in spec.location_evidence],
                "required_inputs": served,
                "required_inputs_count": len(served),
                "required_inputs_bytes": sum(int(str(s["bytes"])) for s in served),
                "input_notes": list(spec.input_notes),
                "allowed_resources": list(spec.allowed_resources),
                "blocked_original_artifacts": [asdict(b) for b in spec.blocked_original_artifacts],
                "comparison_metric": spec.comparison_metric,
                "numerical_agreement_criterion": spec.numerical_agreement_criterion,
                "conclusion_criterion": spec.conclusion_criterion,
                "resource_ceiling": CEILING,
                "stopping_rule": list(STOPPING_RULE),
                "specification_gaps": list(spec.specification_gaps),
                "gate_reviews": [asdict(r) for r in spec.gate_reviews],
                "delegated_assessment_sha256": sha256(
                    json.dumps(
                        delegated[spec.paper_id], sort_keys=True, ensure_ascii=False
                    ).encode()
                ),
            }
        )
    record: dict[str, object] = {
        "freeze_id": FREEZE_ID,
        "amendment": AMENDMENT,
        "frozen_at_utc": datetime.now(timezone.utc).isoformat(),
        "verifier": "Devin delegated second pass over retained article text and acquired inputs",
        "investigator_verification": "pending",
        "human_adjudication": "pending",
        "original_author_contact": "none",
        "denominator_role": "excluded_from_primary_reconstruction_rate",
        "paper_iii_analyses": "not_performed",
        "papers": papers,
    }
    text = json.dumps(record, indent=2, ensure_ascii=False)
    if BANNED_WORDING in text.lower():
        raise ValueError("banned wording in freeze record")
    return record


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--stamp", action="store_true", help="RFC 3161 timestamp the record")
    arguments = parser.parse_args()
    if RECORD.exists():
        payload = RECORD.read_bytes()
        record = mapping(json.loads(payload))
        rebuilt = build()
        for key in ("papers", "freeze_id", "amendment", "denominator_role"):
            if rebuilt[key] != record[key]:
                raise ValueError(f"frozen record no longer reproduces from sources: {key}")
    else:
        record = build()
        payload = (json.dumps(record, indent=2, ensure_ascii=False) + "\n").encode()
        snapshot(RECORD, payload)
    print(RECORD, sha256(payload))
    papers = record["papers"]
    assert isinstance(papers, list)
    for paper in papers:
        entry = mapping(paper)
        print(entry["paper_id"], entry["required_inputs_count"], entry["required_inputs_bytes"])
    if arguments.stamp:
        receipt = stamp(RECORD, STAMP_DIR)
        print(json.dumps(receipt, indent=2)[:600])


if __name__ == "__main__":
    main()
