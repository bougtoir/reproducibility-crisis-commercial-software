#!/usr/bin/env python3
"""
PubMed Stratified Random Sampling Pipeline for Reproducibility Crisis Study.

This script:
1. Queries PubMed E-utilities for papers published 2020+ in English
2. Performs stratified random sampling by research field (MeSH-based)
3. Extracts metadata (PMID, title, journal, date, DOI, etc.)
4. Attempts to retrieve full text from PMC for software extraction
5. Extracts software names, versions, and related metadata from Methods sections
"""

import os
import re
import json
import time
import random
import logging
import xml.etree.ElementTree as ET
from collections import defaultdict
from datetime import datetime
from pathlib import Path

import requests
import pandas as pd
from tqdm import tqdm

# ── Configuration ──────────────────────────────────────────────────────────
API_KEY = os.environ.get("NCBI_API_KEY", "")
BASE_URL = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/"
RATE_LIMIT = 0.11 if API_KEY else 0.34  # seconds between requests
TARGET_SAMPLE = 10000  # full study size
OUTPUT_DIR = Path("/home/ubuntu/reproducibility_study/output")
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(OUTPUT_DIR / "sampling.log"),
        logging.StreamHandler(),
    ],
)
logger = logging.getLogger(__name__)

# ── MeSH-based strata definitions ─────────────────────────────────────────
# We define 7 strata based on broad MeSH categories
STRATA = {
    "Biomedical_Basic": {
        "query_addition": '("Molecular Biology"[MeSH] OR "Genetics"[MeSH] OR "Biochemistry"[MeSH] OR "Cell Biology"[MeSH] OR "Microbiology"[MeSH])',
        "target_n": 1429,
    },
    "Clinical_Medicine": {
        "query_addition": '("Therapeutics"[MeSH] OR "Diagnosis"[MeSH] OR "Clinical Trials as Topic"[MeSH] OR "Surgical Procedures, Operative"[MeSH])',
        "target_n": 1429,
    },
    "Chemistry_Materials": {
        "query_addition": '("Chemistry"[MeSH] OR "Materials Science"[MeSH] OR "Nanotechnology"[MeSH] OR "Polymers"[MeSH])',
        "target_n": 1429,
    },
    "Physics_Engineering": {
        "query_addition": '("Physics"[MeSH] OR "Biomedical Engineering"[MeSH] OR "Signal Processing, Computer-Assisted"[MeSH] OR "Imaging, Three-Dimensional"[MeSH])',
        "target_n": 1429,
    },
    "Social_Behavioral": {
        "query_addition": '("Psychology"[MeSH] OR "Behavioral Sciences"[MeSH] OR "Sociology"[MeSH] OR "Public Health"[MeSH] OR "Epidemiology"[MeSH])',
        "target_n": 1429,
    },
    "Computational_Science": {
        "query_addition": '("Computational Biology"[MeSH] OR "Artificial Intelligence"[MeSH] OR "Machine Learning"[MeSH] OR "Bioinformatics"[MeSH])',
        "target_n": 1429,
    },
    "Environmental_Earth": {
        "query_addition": '("Environmental Sciences"[MeSH] OR "Ecology"[MeSH] OR "Climate"[MeSH] OR "Conservation of Natural Resources"[MeSH])',
        "target_n": 1426,
    },
}

# ── Software detection patterns ───────────────────────────────────────────
SOFTWARE_PATTERNS = [
    # Commercial statistical software
    (r'\bMATLAB\b', 'MATLAB', 'commercial'),
    (r'\bSPSS\b', 'SPSS', 'commercial'),
    (r'\bSAS\s+(software|Institute|version|v\d|9\.|OnDemand)', 'SAS', 'commercial'),
    (r'\bStata\b(?!\s*-)', 'Stata', 'commercial'),
    (r'\bGraphPad\s*Prism\b', 'GraphPad Prism', 'commercial'),
    (r'\bPrism\s+\d', 'GraphPad Prism', 'commercial'),
    (r'\bMathematica\b', 'Mathematica', 'commercial'),
    (r'\bOriginPro\b', 'OriginPro', 'commercial'),
    (r'\bOrigin\s*Lab\b', 'OriginLab', 'commercial'),
    (r'\bJMP\b(?:\s+Pro|\s+\d|\s+software)', 'JMP', 'commercial'),
    (r'\bMinitab\b', 'Minitab', 'commercial'),
    (r'\bEViews\b', 'EViews', 'commercial'),
    (r'\bGenStat\b', 'GenStat', 'commercial'),
    (r'\bLabVIEW\b', 'LabVIEW', 'commercial'),
    (r'\bFlowJo\b', 'FlowJo', 'commercial'),
    (r'\bTreeStar\b', 'FlowJo', 'commercial'),
    (r'\bGaussian\s*(09|16|0[0-9])\b', 'Gaussian', 'commercial'),
    (r'\bVASP\b', 'VASP', 'commercial'),
    (r'\bCOMSOL\b', 'COMSOL', 'commercial'),
    (r'\bANSYS\b', 'ANSYS', 'commercial'),
    (r'\bAbaqus\b', 'Abaqus', 'commercial'),
    (r'\bMplus\b', 'Mplus', 'commercial'),
    (r'\bLISREL\b', 'LISREL', 'commercial'),
    (r'\bAMOS\b(?:\s+\d|\s+software|\s+version)', 'AMOS', 'commercial'),
    (r'\bHLM\s*\d', 'HLM', 'commercial'),
    (r'\bEndNote\b', 'EndNote', 'commercial'),
    (r'\bAdobe\s+(Photoshop|Illustrator|InDesign|Acrobat|Premiere)', 'Adobe \\1', 'commercial'),
    (r'\bPhotoshop\b', 'Adobe Photoshop', 'commercial'),
    (r'\bMicrosoft\s+Excel\b', 'Microsoft Excel', 'commercial'),
    (r'\bExcel\s+(version|v\d|20[01]\d|software|spreadsheet)', 'Microsoft Excel', 'commercial'),
    (r'\bMicrosoft\s+Office\b', 'Microsoft Office', 'commercial'),
    (r'\bSchrodinger\b|\bSchrödinger\b', 'Schrödinger', 'commercial'),
    (r'\bMOE\b(?:\s+\d|\s+software|\s+\()', 'MOE', 'commercial'),
    (r'\bGeneious\b', 'Geneious', 'commercial'),
    (r'\bCLC\s+(Genomics|Workbench|bio)', 'CLC Workbench', 'commercial'),
    (r'\bSnapGene\b', 'SnapGene', 'commercial'),
    (r'\bIngenuity\b', 'Ingenuity Pathway Analysis', 'commercial'),
    (r'\bIPA\b(?:\s+software|\s+\(Qiagen|\s+analysis)', 'Ingenuity Pathway Analysis', 'commercial'),
    (r'\bPartek\b', 'Partek', 'commercial'),
    (r'\bBioRender\b', 'BioRender', 'commercial'),
    (r'\bIMARIS\b|\bImaris\b', 'Imaris', 'commercial'),
    (r'\bMetaMorph\b', 'MetaMorph', 'commercial'),
    (r'\bZEN\s+(software|Blue|Black|\d)', 'Zeiss ZEN', 'commercial'),
    (r'\bNIS[\-\s]Elements\b', 'NIS-Elements', 'commercial'),
    (r'\bcellSens\b', 'cellSens', 'commercial'),
    (r'\bSoftMax\b', 'SoftMax Pro', 'commercial'),
    (r'\bGen5\b(?:\s+software|\s+\d)', 'Gen5', 'commercial'),
    (r'\bTrimmomatic\b', 'Trimmomatic', 'open_source'),

    # Open-source / free software
    (r'\bR\s*\(\s*(?:version\s*)?\d', 'R', 'open_source'),
    (r'\bR\s+software\b', 'R', 'open_source'),
    (r'\bR\s+(?:v\d|version|Core Team|Foundation|package)', 'R', 'open_source'),
    (r'\bPython\b(?:\s+\d|\s+v\d|\s+version|\s+software|\s+script|\s+package|\s+library)', 'Python', 'open_source'),
    (r'\bImageJ\b', 'ImageJ', 'open_source'),
    (r'\bFIJI\b|\bFiji\b', 'Fiji/ImageJ', 'open_source'),
    (r'\bCellProfiler\b', 'CellProfiler', 'open_source'),
    (r'\bBLAST\b(?:n|p|x|\+|\s)', 'BLAST', 'open_source'),
    (r'\bBowtie\s*2?\b', 'Bowtie', 'open_source'),
    (r'\bBWA\b(?:\s|-|/)', 'BWA', 'open_source'),
    (r'\bSamtools\b|\bsamtools\b', 'Samtools', 'open_source'),
    (r'\bGATK\b', 'GATK', 'open_source'),
    (r'\bSTAR\b(?:\s+aligner|\s+v\d|\s+\d)', 'STAR', 'open_source'),
    (r'\bHISAT2?\b|\bhisat2?\b', 'HISAT2', 'open_source'),
    (r'\bSalmon\b(?:\s+v\d|\s+version|\s+software|\s+\()', 'Salmon', 'open_source'),
    (r'\bkallisto\b', 'kallisto', 'open_source'),
    (r'\bDESeq2?\b', 'DESeq2', 'open_source'),
    (r'\bedgeR\b', 'edgeR', 'open_source'),
    (r'\bCufflinks\b', 'Cufflinks', 'open_source'),
    (r'\bStringTie\b', 'StringTie', 'open_source'),
    (r'\bBedtools\b|\bbedtools\b', 'Bedtools', 'open_source'),
    (r'\bPyMOL\b', 'PyMOL', 'open_source'),
    (r'\bVMD\b(?:\s+\d|\s+software|\s+\()', 'VMD', 'open_source'),
    (r'\bGROMACS\b|\bgromacs\b', 'GROMACS', 'open_source'),
    (r'\bNAMD\b', 'NAMD', 'open_source'),
    (r'\bOpenBabel\b|Open\s+Babel', 'Open Babel', 'open_source'),
    (r'\bRDKit\b', 'RDKit', 'open_source'),
    (r'\bAutoDock\b', 'AutoDock', 'open_source'),
    (r'\bQuantum\s+ESPRESSO\b', 'Quantum ESPRESSO', 'open_source'),
    (r'\bGNUplot\b|\bgnuplot\b', 'gnuplot', 'open_source'),
    (r'\bggplot2\b', 'ggplot2', 'open_source'),
    (r'\bmatplotlib\b', 'matplotlib', 'open_source'),
    (r'\bscikit[\-\s]learn\b', 'scikit-learn', 'open_source'),
    (r'\bTensorFlow\b|\btensorflow\b', 'TensorFlow', 'open_source'),
    (r'\bPyTorch\b|\bpytorch\b', 'PyTorch', 'open_source'),
    (r'\bKeras\b', 'Keras', 'open_source'),
    (r'\bJASP\b(?:\s+software|\s+\d|\s+version|\s+\()', 'JASP', 'open_source'),
    (r'\bjamovi\b', 'jamovi', 'open_source'),
    (r'\bPSPP\b', 'PSPP', 'open_source'),
    (r'\bGNU\s+Octave\b', 'GNU Octave', 'open_source'),
    (r'\bSageMath\b', 'SageMath', 'open_source'),
    (r'\bMaxima\b(?:\s+software|\s+\d)', 'Maxima', 'open_source'),
    (r'\bCytoscape\b', 'Cytoscape', 'open_source'),
    (r'\bGephi\b', 'Gephi', 'open_source'),
    (r'\bRevMan\b', 'RevMan', 'open_source'),
    (r'\bOpenCV\b', 'OpenCV', 'open_source'),
    (r'\bSCIkit\-image\b|\bskimage\b', 'scikit-image', 'open_source'),
    (r'\bNextflow\b', 'Nextflow', 'open_source'),
    (r'\bSnakemake\b', 'Snakemake', 'open_source'),
    (r'\bDocker\b', 'Docker', 'open_source'),
    (r'\bSingularity\b(?:\s+container|\s+image)', 'Singularity', 'open_source'),
    (r'\bConda\b|\bconda\b', 'Conda', 'open_source'),
    (r'\bBioconductor\b', 'Bioconductor', 'open_source'),
    (r'\bSEURAT\b|\bSeurat\b', 'Seurat', 'open_source'),
    (r'\bScanpy\b|\bscanpy\b', 'Scanpy', 'open_source'),
]

# Version pattern
VERSION_PATTERN = re.compile(
    r'(?:version|ver\.?|v\.?)\s*(\d+(?:\.\d+){0,3}(?:\s*[a-zA-Z]\d*)?)'
    r'|'
    r'(\d+\.\d+(?:\.\d+){0,2}(?:\s*[a-zA-Z]\d*)?)',
    re.IGNORECASE,
)

# Code/data availability patterns
CODE_AVAIL_PATTERNS = [
    r'(?:code|script|software)\s+(?:is\s+)?(?:available|accessible|deposited|archived)',
    r'github\.com/[^\s\)]+',
    r'gitlab\.com/[^\s\)]+',
    r'bitbucket\.org/[^\s\)]+',
    r'zenodo\.org/(?:record|doi)/[^\s\)]+',
    r'figshare\.com/[^\s\)]+',
    r'(?:source\s+)?code\s+(?:is\s+)?(?:freely\s+)?available',
    r'open[\s-]?source',
]

DATA_AVAIL_PATTERNS = [
    r'(?:data|dataset)\s+(?:is\s+)?(?:available|accessible|deposited|archived)',
    r'data\s+availability\s+statement',
    r'(?:data|dataset)\s+(?:can\s+be\s+)?(?:downloaded|accessed|obtained)',
    r'GEO\s+(?:accession|series)',
    r'(?:SRA|ENA|DDBJ)\s+(?:accession|database)',
    r'ArrayExpress',
    r'Dryad',
    r'NCBI\s+(?:GEO|SRA)',
]

REPRODUCIBILITY_PATTERNS = [
    r'reproduc(?:ib|e)',
    r'replic(?:ab|at)',
    r'availability\s+statement',
    r'(?:code|data)\s+sharing',
    r'FAIR\s+(?:data|principles)',
]


# ── API Helper Functions ──────────────────────────────────────────────────

def api_request(endpoint, params, max_retries=5):
    """Make a request to PubMed E-utilities with rate limiting and retry."""
    if API_KEY:
        params["api_key"] = API_KEY
    url = BASE_URL + endpoint
    for attempt in range(max_retries):
        try:
            time.sleep(RATE_LIMIT)
            resp = requests.get(url, params=params, timeout=30)
            resp.raise_for_status()
            return resp
        except requests.exceptions.RequestException as e:
            wait = 2 ** attempt
            logger.warning(f"Request failed (attempt {attempt+1}/{max_retries}): {e}. Retrying in {wait}s...")
            time.sleep(wait)
    raise RuntimeError(f"Failed after {max_retries} attempts: {url}")


def esearch(query, retmax=0, retstart=0, use_history=False):
    """Search PubMed and return count + PMIDs or web environment (XML-based)."""
    params = {
        "db": "pubmed",
        "term": query,
        "retmax": retmax,
        "retstart": retstart,
    }
    if use_history:
        params["usehistory"] = "y"
    resp = api_request("esearch.fcgi", params)
    root = ET.fromstring(resp.content)
    count_elem = root.find("Count")
    count = int(count_elem.text) if count_elem is not None else 0
    id_list = [id_elem.text for id_elem in root.findall(".//IdList/Id")]
    result = {"esearchresult": {"count": str(count), "idlist": id_list}}
    if use_history:
        qk = root.find("QueryKey")
        we = root.find("WebEnv")
        if qk is not None:
            result["esearchresult"]["querykey"] = qk.text
        if we is not None:
            result["esearchresult"]["webenv"] = we.text
    return result


def efetch_xml(pmids):
    """Fetch full records for a list of PMIDs in XML format."""
    params = {
        "db": "pubmed",
        "id": ",".join(str(p) for p in pmids),
        "rettype": "xml",
        "retmode": "xml",
    }
    resp = api_request("efetch.fcgi", params)
    return ET.fromstring(resp.content)


def pmc_fetch_fulltext(pmc_id):
    """Fetch full text XML from PMC."""
    params = {
        "db": "pmc",
        "id": pmc_id,
        "rettype": "xml",
        "retmode": "xml",
    }
    try:
        resp = api_request("efetch.fcgi", params)
        return ET.fromstring(resp.content)
    except Exception as e:
        logger.debug(f"PMC fetch failed for {pmc_id}: {e}")
        return None


# ── Sampling Functions ────────────────────────────────────────────────────

def get_stratum_count(stratum_query_addition):
    """Get the total count of papers for a stratum."""
    base_query = (
        '("2020/01/01"[PDAT] : "3000"[PDAT]) '
        'AND eng[la] '
        'AND "journal article"[pt] '
        'NOT review[pt] '
        'NOT "case reports"[pt] '
        'AND hasabstract '
    )
    full_query = base_query + "AND " + stratum_query_addition
    result = esearch(full_query, retmax=0)
    count = int(result["esearchresult"]["count"])
    return count, full_query


def sample_pmids_from_stratum(full_query, total_count, target_n):
    """Randomly sample PMIDs from a stratum using year-split sampling.
    
    PubMed limits retstart to 9999, so for large result sets we split
    by publication year (2020-2025) and sample proportionally from each year.
    """
    if total_count <= target_n:
        result = esearch(full_query, retmax=total_count)
        return result["esearchresult"]["idlist"]

    # Split by year for better coverage and to work within PubMed's retstart limit
    years = ["2020", "2021", "2022", "2023", "2024", "2025"]
    all_candidates = []
    
    for year in years:
        year_query = full_query.replace(
            '("2020/01/01"[PDAT] : "3000"[PDAT])',
            f'("{year}/01/01"[PDAT] : "{year}/12/31"[PDAT])'
        )
        # First get count for this year
        count_result = esearch(year_query, retmax=0)
        year_count = int(count_result["esearchresult"]["count"])
        
        if year_count == 0:
            continue
            
        # Fetch up to 9999 PMIDs (PubMed limit)
        fetch_size = min(9999, year_count)
        # Use a random offset within the safe range
        max_start = max(0, min(year_count, 9999) - fetch_size)
        start = random.randint(0, max_start) if max_start > 0 else 0
        
        result = esearch(year_query, retmax=fetch_size, retstart=start)
        year_pmids = result["esearchresult"]["idlist"]
        all_candidates.extend(year_pmids)
        logger.info(f"    Year {year}: {year_count:,} total, fetched {len(year_pmids)}")
    
    if len(all_candidates) <= target_n:
        return all_candidates
    
    return random.sample(all_candidates, target_n)


def run_stratified_sampling():
    """Run stratified random sampling across all strata."""
    logger.info("=" * 60)
    logger.info("Starting stratified random sampling")
    logger.info(f"Target total sample: {TARGET_SAMPLE}")
    logger.info("=" * 60)

    all_pmids = {}
    stratum_stats = {}

    for stratum_name, config in STRATA.items():
        logger.info(f"\n--- Stratum: {stratum_name} ---")
        count, full_query = get_stratum_count(config["query_addition"])
        logger.info(f"  Total available: {count:,}")
        logger.info(f"  Target sample: {config['target_n']}")

        pmids = sample_pmids_from_stratum(full_query, count, config["target_n"])
        logger.info(f"  Sampled: {len(pmids)}")

        all_pmids[stratum_name] = pmids
        stratum_stats[stratum_name] = {
            "total_available": count,
            "target_n": config["target_n"],
            "sampled_n": len(pmids),
        }

    total_sampled = sum(len(v) for v in all_pmids.values())
    logger.info(f"\nTotal PMIDs sampled: {total_sampled}")

    # Save sampling stats
    stats_df = pd.DataFrame(stratum_stats).T
    stats_df.to_csv(OUTPUT_DIR / "sampling_stats.csv")

    return all_pmids, stratum_stats


# ── Data Extraction Functions ─────────────────────────────────────────────

def extract_text_recursive(element):
    """Recursively extract all text from an XML element."""
    if element is None:
        return ""
    text = element.text or ""
    for child in element:
        text += extract_text_recursive(child)
        if child.tail:
            text += child.tail
    return text


def parse_pubmed_article(article_elem):
    """Parse a PubMed article XML element into a dict."""
    record = {}

    # PMID
    pmid_elem = article_elem.find(".//PMID")
    record["pmid"] = pmid_elem.text if pmid_elem is not None else ""

    # Title
    title_elem = article_elem.find(".//ArticleTitle")
    record["title"] = extract_text_recursive(title_elem) if title_elem is not None else ""

    # Journal
    journal_elem = article_elem.find(".//Journal/Title")
    record["journal"] = journal_elem.text if journal_elem is not None else ""

    # Journal abbreviation
    journal_abbr = article_elem.find(".//Journal/ISOAbbreviation")
    record["journal_abbr"] = journal_abbr.text if journal_abbr is not None else ""

    # ISSN
    issn_elem = article_elem.find(".//Journal/ISSN")
    record["issn"] = issn_elem.text if issn_elem is not None else ""

    # Publication date
    pub_date = article_elem.find(".//PubDate")
    if pub_date is not None:
        year = pub_date.find("Year")
        month = pub_date.find("Month")
        day = pub_date.find("Day")
        record["pub_year"] = year.text if year is not None else ""
        record["pub_month"] = month.text if month is not None else ""
        record["pub_day"] = day.text if day is not None else ""
    else:
        record["pub_year"] = record["pub_month"] = record["pub_day"] = ""

    # DOI
    doi_elem = article_elem.find('.//ArticleId[@IdType="doi"]')
    if doi_elem is None:
        doi_elem = article_elem.find('.//ELocationID[@EIdType="doi"]')
    record["doi"] = doi_elem.text if doi_elem is not None else ""

    # PMC ID
    pmc_elem = article_elem.find('.//ArticleId[@IdType="pmc"]')
    record["pmc_id"] = pmc_elem.text if pmc_elem is not None else ""

    # Abstract
    abstract_parts = article_elem.findall(".//Abstract/AbstractText")
    abstract_texts = []
    for part in abstract_parts:
        label = part.get("Label", "")
        text = extract_text_recursive(part)
        if label:
            abstract_texts.append(f"{label}: {text}")
        else:
            abstract_texts.append(text)
    record["abstract"] = " ".join(abstract_texts)

    # MeSH terms
    mesh_terms = []
    for mesh in article_elem.findall(".//MeshHeading/DescriptorName"):
        mesh_terms.append(mesh.text)
    record["mesh_terms"] = "; ".join(mesh_terms)

    # Affiliations
    affiliations = []
    for aff in article_elem.findall(".//Affiliation"):
        if aff.text:
            affiliations.append(aff.text)
    record["affiliations"] = "; ".join(affiliations[:3])  # Keep first 3

    # Country from affiliation (simple extraction)
    record["country"] = ""
    if affiliations:
        # Try to get country from last part of first affiliation
        parts = affiliations[0].split(",")
        if parts:
            record["country"] = parts[-1].strip().rstrip(".")

    # Grant/funding info
    grants = []
    for grant in article_elem.findall(".//Grant"):
        agency = grant.find("Agency")
        if agency is not None and agency.text:
            grants.append(agency.text)
    record["funding_agencies"] = "; ".join(set(grants))

    # Publication types
    pub_types = []
    for pt in article_elem.findall(".//PublicationType"):
        pub_types.append(pt.text)
    record["publication_types"] = "; ".join(pub_types)

    return record


def detect_software(text):
    """Detect software mentions in text. Returns list of (name, license_type) tuples."""
    found = []
    for pattern, name, license_type in SOFTWARE_PATTERNS:
        if re.search(pattern, text):
            # Check for version near the software mention
            found.append((name, license_type))
    return list(set(found))


def extract_version_for_software(text, software_name):
    """Extract version number mentioned near a software name."""
    # Search in a window around the software mention
    # Escape special regex characters in the software name
    escaped_name = re.escape(software_name)
    # Look for version within 100 chars of software name
    pattern = re.compile(
        escaped_name + r'.{0,80}?(?:version|ver\.?|v\.?)\s*(\d+(?:\.\d+){0,3}(?:\s*[a-zA-Z]\d*)?)',
        re.IGNORECASE | re.DOTALL,
    )
    match = pattern.search(text)
    if match:
        return match.group(1).strip()

    # Also try: SoftwareName X.Y.Z pattern
    pattern2 = re.compile(
        escaped_name + r'\s+(\d+\.\d+(?:\.\d+){0,2})',
        re.IGNORECASE,
    )
    match2 = pattern2.search(text)
    if match2:
        return match2.group(1).strip()

    return ""


def check_code_availability(text):
    """Check if code availability is mentioned."""
    for pattern in CODE_AVAIL_PATTERNS:
        if re.search(pattern, text, re.IGNORECASE):
            return True
    return False


def check_data_availability(text):
    """Check if data availability is mentioned."""
    for pattern in DATA_AVAIL_PATTERNS:
        if re.search(pattern, text, re.IGNORECASE):
            return True
    return False


def check_reproducibility_statement(text):
    """Check if reproducibility/replicability is mentioned."""
    for pattern in REPRODUCIBILITY_PATTERNS:
        if re.search(pattern, text, re.IGNORECASE):
            return True
    return False


def extract_github_urls(text):
    """Extract GitHub/GitLab URLs from text."""
    urls = re.findall(r'(?:https?://)?(?:github|gitlab)\.com/[\w\-\.]+/[\w\-\.]+', text)
    return "; ".join(urls) if urls else ""


# ── PMC Full Text Processing ─────────────────────────────────────────────

def extract_methods_from_pmc(pmc_xml):
    """Extract methods section text from PMC full-text XML."""
    if pmc_xml is None:
        return ""

    methods_text = ""
    # Look for sections with Methods/Materials-related titles
    for sec in pmc_xml.iter("sec"):
        title = sec.find("title")
        if title is not None and title.text:
            title_lower = title.text.lower()
            if any(kw in title_lower for kw in
                   ["method", "material", "experimental", "procedure",
                    "software", "statistical analys", "data analys",
                    "computational", "supplement"]):
                methods_text += " " + extract_text_recursive(sec)

    # If no methods section found, try body text
    if not methods_text:
        body = pmc_xml.find(".//body")
        if body is not None:
            methods_text = extract_text_recursive(body)

    return methods_text


# ── Main Processing Pipeline ─────────────────────────────────────────────

def process_papers(all_pmids, stratum_stats):
    """Process all sampled papers: fetch metadata and extract software info."""
    logger.info("\n" + "=" * 60)
    logger.info("Starting data extraction pipeline")
    logger.info("=" * 60)

    all_records = []
    batch_size = 50  # PubMed efetch batch size

    for stratum_name, pmids in all_pmids.items():
        logger.info(f"\nProcessing stratum: {stratum_name} ({len(pmids)} papers)")

        for i in tqdm(range(0, len(pmids), batch_size), desc=stratum_name):
            batch = pmids[i:i + batch_size]
            try:
                xml_root = efetch_xml(batch)
            except Exception as e:
                logger.error(f"efetch failed for batch starting at {i}: {e}")
                continue

            for article in xml_root.findall(".//PubmedArticle"):
                record = parse_pubmed_article(article)
                record["stratum"] = stratum_name

                # Combine abstract for software detection
                search_text = record["abstract"]
                software_detected = detect_software(search_text)

                # Try PMC full text if available
                pmc_text = ""
                if record["pmc_id"]:
                    pmc_xml = pmc_fetch_fulltext(record["pmc_id"])
                    if pmc_xml is not None:
                        pmc_text = extract_methods_from_pmc(pmc_xml)
                        record["has_pmc_fulltext"] = True
                        # Also detect software from full text
                        software_from_fulltext = detect_software(pmc_text)
                        software_detected = list(set(software_detected + software_from_fulltext))
                    else:
                        record["has_pmc_fulltext"] = False
                else:
                    record["has_pmc_fulltext"] = False

                # Combined text for analysis
                combined_text = search_text + " " + pmc_text

                # Software information
                software_names = [s[0] for s in software_detected]
                software_types = {s[0]: s[1] for s in software_detected}

                record["software_mentioned"] = "; ".join(software_names)
                record["software_count"] = len(software_names)
                record["has_commercial_software"] = any(
                    t == "commercial" for t in software_types.values()
                )
                record["has_opensource_software"] = any(
                    t == "open_source" for t in software_types.values()
                )
                record["commercial_software_list"] = "; ".join(
                    [n for n, t in software_types.items() if t == "commercial"]
                )
                record["opensource_software_list"] = "; ".join(
                    [n for n, t in software_types.items() if t == "open_source"]
                )
                record["commercial_software_count"] = sum(
                    1 for t in software_types.values() if t == "commercial"
                )
                record["opensource_software_count"] = sum(
                    1 for t in software_types.values() if t == "open_source"
                )

                # Version extraction for each software
                versions = {}
                for sw_name in software_names:
                    ver = extract_version_for_software(combined_text, sw_name)
                    if ver:
                        versions[sw_name] = ver
                record["software_versions"] = json.dumps(versions) if versions else ""
                record["version_mentioned_count"] = len(versions)
                record["version_mention_rate"] = (
                    len(versions) / len(software_names) if software_names else 0
                )

                # Code/data availability
                record["code_available"] = check_code_availability(combined_text)
                record["data_available"] = check_data_availability(combined_text)
                record["reproducibility_mentioned"] = check_reproducibility_statement(combined_text)
                record["repository_urls"] = extract_github_urls(combined_text)

                all_records.append(record)

        logger.info(f"  Completed {stratum_name}: {len([r for r in all_records if r['stratum'] == stratum_name])} records")

    return all_records


# ── Impact Factor Mapping ─────────────────────────────────────────────────

def load_sjr_data():
    """Load SJR data if available, otherwise return empty mapping."""
    sjr_path = OUTPUT_DIR / "sjr_data.csv"
    if sjr_path.exists():
        df = pd.read_csv(sjr_path)
        return dict(zip(df["Issn"].str.strip(), df["SJR"]))
    return {}


# ── Software Availability Database ────────────────────────────────────────

SOFTWARE_AVAILABILITY = {
    "MATLAB": {"current_version": "R2025a", "legacy_available": True, "free_access": "MATLAB Online Basic (20h/month)", "license": "commercial"},
    "SPSS": {"current_version": "31", "legacy_available": False, "free_access": "14-day trial only", "license": "commercial"},
    "SAS": {"current_version": "Viya", "legacy_available": False, "free_access": "SAS OnDemand for Academics (current only)", "license": "commercial"},
    "Stata": {"current_version": "19", "legacy_available": False, "free_access": "None", "license": "commercial"},
    "GraphPad Prism": {"current_version": "10", "legacy_available": False, "free_access": "30-day trial only", "license": "commercial"},
    "Mathematica": {"current_version": "14.2", "legacy_available": False, "free_access": "Wolfram Engine (CLI only, non-production)", "license": "commercial"},
    "OriginPro": {"current_version": "2025", "legacy_available": False, "free_access": "Trial only", "license": "commercial"},
    "JMP": {"current_version": "18", "legacy_available": False, "free_access": "30-day trial", "license": "commercial"},
    "FlowJo": {"current_version": "10.10", "legacy_available": False, "free_access": "Trial only", "license": "commercial"},
    "Gaussian": {"current_version": "16 Rev C.02", "legacy_available": True, "free_access": "None (maintenance program)", "license": "commercial"},
    "VASP": {"current_version": "6.4", "legacy_available": True, "free_access": "None (group license)", "license": "commercial"},
    "COMSOL": {"current_version": "6.4", "legacy_available": False, "free_access": "Trial only", "license": "commercial"},
    "ANSYS": {"current_version": "2025 R1", "legacy_available": False, "free_access": "Student version (limited)", "license": "commercial"},
    "Abaqus": {"current_version": "2025", "legacy_available": False, "free_access": "Student edition (limited)", "license": "commercial"},
    "LabVIEW": {"current_version": "2024 Q3", "legacy_available": False, "free_access": "Community Edition", "license": "commercial"},
    "Mplus": {"current_version": "8.11", "legacy_available": False, "free_access": "Demo (limited)", "license": "commercial"},
    "Adobe Photoshop": {"current_version": "26.x", "legacy_available": False, "free_access": "7-day trial", "license": "commercial"},
    "Microsoft Excel": {"current_version": "365", "legacy_available": False, "free_access": "Excel Online (limited)", "license": "commercial"},
    "EndNote": {"current_version": "21", "legacy_available": False, "free_access": "EndNote Basic (limited web)", "license": "commercial"},
    "Schrödinger": {"current_version": "2025-1", "legacy_available": False, "free_access": "Academic license available", "license": "commercial"},
    "Geneious": {"current_version": "2025.0", "legacy_available": False, "free_access": "Trial only", "license": "commercial"},
    "CLC Workbench": {"current_version": "25.0", "legacy_available": False, "free_access": "Trial only", "license": "commercial"},
    "Ingenuity Pathway Analysis": {"current_version": "current", "legacy_available": False, "free_access": "None", "license": "commercial"},
    "Imaris": {"current_version": "10.2", "legacy_available": False, "free_access": "Viewer only (free)", "license": "commercial"},
    "BioRender": {"current_version": "current", "legacy_available": False, "free_access": "Free tier (limited exports)", "license": "commercial"},
    "Minitab": {"current_version": "22", "legacy_available": False, "free_access": "30-day trial", "license": "commercial"},
    "Zeiss ZEN": {"current_version": "3.9", "legacy_available": False, "free_access": "ZEN lite (free, limited)", "license": "commercial"},
    "NIS-Elements": {"current_version": "6.0", "legacy_available": False, "free_access": "Viewer only", "license": "commercial"},
    "Partek": {"current_version": "current", "legacy_available": False, "free_access": "Trial only", "license": "commercial"},
    "SnapGene": {"current_version": "8.0", "legacy_available": False, "free_access": "Viewer only (free)", "license": "commercial"},
    "MetaMorph": {"current_version": "current", "legacy_available": False, "free_access": "None", "license": "commercial"},
    "EViews": {"current_version": "14", "legacy_available": False, "free_access": "Student version (limited)", "license": "commercial"},
    "LISREL": {"current_version": "12", "legacy_available": False, "free_access": "Student version (limited)", "license": "commercial"},
    "HLM": {"current_version": "8", "legacy_available": False, "free_access": "Student version", "license": "commercial"},
    "MOE": {"current_version": "2024", "legacy_available": False, "free_access": "Academic license", "license": "commercial"},
    "AMOS": {"current_version": "29", "legacy_available": False, "free_access": "None", "license": "commercial"},
    "GenStat": {"current_version": "24", "legacy_available": False, "free_access": "None", "license": "commercial"},
    # Open source - always available
    "R": {"current_version": "4.5.0", "legacy_available": True, "free_access": "Fully free", "license": "open_source"},
    "Python": {"current_version": "3.13", "legacy_available": True, "free_access": "Fully free", "license": "open_source"},
    "ImageJ": {"current_version": "1.54", "legacy_available": True, "free_access": "Fully free", "license": "open_source"},
    "Fiji/ImageJ": {"current_version": "2.15", "legacy_available": True, "free_access": "Fully free", "license": "open_source"},
}


def check_version_availability(software_name, version_str):
    """Check if a specific version of software is available."""
    info = SOFTWARE_AVAILABILITY.get(software_name, {})
    if not info:
        return "unknown"
    if info.get("license") == "open_source":
        return "available"
    if info.get("legacy_available"):
        return "legacy_available"
    if version_str:
        current = info.get("current_version", "")
        if current and version_str in current:
            return "current"
        return "likely_unavailable"
    return "unknown"


# ── Commercial Software Cost Database ─────────────────────────────────────

SOFTWARE_COSTS_USD = {
    "MATLAB": 2350,  # Standard individual license
    "SPSS": 1188,    # Annual subscription
    "SAS": 8500,     # Annual (estimated, varies greatly)
    "Stata": 595,    # Stata/SE annual
    "GraphPad Prism": 252,  # Annual individual
    "Mathematica": 385,   # Annual
    "OriginPro": 1100,   # Single user
    "JMP": 1785,     # Annual
    "FlowJo": 500,   # Annual academic
    "Gaussian": 3000, # Academic group
    "VASP": 4500,    # Academic group
    "COMSOL": 5000,  # Annual (estimated)
    "ANSYS": 5000,   # Annual (estimated)
    "Abaqus": 5000,  # Annual (estimated)
    "LabVIEW": 800,  # Annual
    "Mplus": 695,    # Single user
    "Adobe Photoshop": 264,  # Annual
    "Microsoft Excel": 100,  # Annual (365)
    "EndNote": 274,   # One-time
    "Schrödinger": 0,  # Academic free
    "Geneious": 575,  # Annual academic
    "CLC Workbench": 2500,  # Annual (estimated)
    "Ingenuity Pathway Analysis": 5000,  # Annual (estimated)
    "Imaris": 3000,   # Annual (estimated)
    "BioRender": 396,  # Annual academic
    "Minitab": 1610,  # Annual
    "Zeiss ZEN": 0,   # Bundled with hardware
    "NIS-Elements": 0, # Bundled with hardware
    "Partek": 3000,   # Annual (estimated)
    "SnapGene": 185,  # Annual academic
    "MetaMorph": 3000, # Estimated
    "EViews": 595,    # Annual academic
    "LISREL": 495,    # Single user
    "HLM": 450,      # Annual
    "MOE": 0,        # Academic free
    "AMOS": 1188,    # Bundled with SPSS
    "GenStat": 500,  # Annual academic
}


def estimate_replication_cost(software_list_str):
    """Estimate the total cost to replicate based on commercial software used."""
    if not software_list_str:
        return 0
    softwares = [s.strip() for s in software_list_str.split(";")]
    total = 0
    for sw in softwares:
        cost = SOFTWARE_COSTS_USD.get(sw, 0)
        total += cost
    return total


# ── Main Execution ────────────────────────────────────────────────────────

def main():
    logger.info("=" * 60)
    logger.info("REPRODUCIBILITY CRISIS EMPIRICAL STUDY")
    logger.info("PubMed Stratified Random Sampling Pipeline")
    logger.info(f"Date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    logger.info(f"Target sample: {TARGET_SAMPLE}")
    logger.info(f"API Key: {'Set' if API_KEY else 'Not set (rate limited)'}")
    logger.info("=" * 60)

    # Step 1: Stratified sampling
    all_pmids, stratum_stats = run_stratified_sampling()

    # Save PMIDs
    pmid_data = []
    for stratum, pmids in all_pmids.items():
        for pmid in pmids:
            pmid_data.append({"stratum": stratum, "pmid": pmid})
    pd.DataFrame(pmid_data).to_csv(OUTPUT_DIR / "sampled_pmids.csv", index=False)
    logger.info(f"Saved {len(pmid_data)} PMIDs to sampled_pmids.csv")

    # Step 2: Process papers
    all_records = process_papers(all_pmids, stratum_stats)

    # Step 3: Add availability and cost information
    for record in all_records:
        if record["software_versions"]:
            versions = json.loads(record["software_versions"])
            availability = {}
            for sw, ver in versions.items():
                availability[sw] = check_version_availability(sw, ver)
            record["version_availability"] = json.dumps(availability)
        else:
            record["version_availability"] = ""

        record["estimated_replication_cost_usd"] = estimate_replication_cost(
            record.get("commercial_software_list", "")
        )

    # Step 4: Save to CSV
    df = pd.DataFrame(all_records)
    df.to_csv(OUTPUT_DIR / "extracted_data.csv", index=False)
    logger.info(f"\nSaved {len(df)} records to extracted_data.csv")

    # Step 5: Generate summary statistics
    generate_summary(df, stratum_stats)

    logger.info("\n" + "=" * 60)
    logger.info("Pipeline complete!")
    logger.info("=" * 60)

    return df


def generate_summary(df, stratum_stats):
    """Generate summary statistics and save them."""
    summary = {
        "total_papers": len(df),
        "papers_with_software": int((df["software_count"] > 0).sum()),
        "papers_with_commercial_sw": int(df["has_commercial_software"].sum()),
        "papers_with_opensource_sw": int(df["has_opensource_software"].sum()),
        "papers_with_version": int((df["version_mentioned_count"] > 0).sum()),
        "papers_with_code_available": int(df["code_available"].sum()),
        "papers_with_data_available": int(df["data_available"].sum()),
        "papers_with_pmc_fulltext": int(df["has_pmc_fulltext"].sum()),
        "papers_with_reproducibility_mention": int(df["reproducibility_mentioned"].sum()),
        "mean_software_per_paper": float(df["software_count"].mean()),
        "mean_commercial_per_paper": float(df["commercial_software_count"].mean()),
        "mean_version_mention_rate": float(df["version_mention_rate"].mean()),
        "mean_replication_cost_usd": float(df["estimated_replication_cost_usd"].mean()),
        "median_replication_cost_usd": float(df["estimated_replication_cost_usd"].median()),
    }

    # Per-stratum summary
    stratum_summary = {}
    for stratum in df["stratum"].unique():
        sdf = df[df["stratum"] == stratum]
        stratum_summary[stratum] = {
            "n": len(sdf),
            "sw_detection_rate": float((sdf["software_count"] > 0).mean()),
            "commercial_rate": float(sdf["has_commercial_software"].mean()),
            "version_mention_rate": float(sdf["version_mention_rate"].mean()),
            "code_available_rate": float(sdf["code_available"].mean()),
            "mean_replication_cost": float(sdf["estimated_replication_cost_usd"].mean()),
        }

    with open(OUTPUT_DIR / "summary_stats.json", "w") as f:
        json.dump({"overall": summary, "by_stratum": stratum_summary}, f, indent=2)

    logger.info("\n--- SUMMARY ---")
    for k, v in summary.items():
        if isinstance(v, float):
            logger.info(f"  {k}: {v:.3f}")
        else:
            logger.info(f"  {k}: {v}")


if __name__ == "__main__":
    random.seed(42)  # For reproducibility
    df = main()
