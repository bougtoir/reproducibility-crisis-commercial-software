# The Hidden Cost of Reproducibility: Commercial Software Dependency in Published Research

Empirical study of commercial software dependency in 10,000 published papers and the version accessibility gap.

## Overview

This project investigates a critical yet underexplored barrier to research reproducibility: dependency on commercial software whose specific versions are often inaccessible for replication. We conducted a cross-sectional empirical study of **10,000 papers** published between 2020 and 2025, sampled from PubMed using stratified random sampling across seven research fields.

### Key Findings (N=10,000)

| Metric | Value |
|--------|-------|
| Commercial software dependency rate | 18.5% (1,853 papers) |
| Software detection rate | 29.2% (2,922 papers) |
| Version reporting rate | 20.2% |
| Mean replication cost (commercial SW papers) | $340 |
| Commercial SW versions "likely unavailable" | 97.0% |
| Code availability statement rate | 5.3% |
| Data availability statement rate | 3.7% |

## Project Structure

```
reproducibility-crisis-commercial-software/
├── README.md
├── pubmed_sampler.py          # PubMed sampling & software extraction pipeline
├── run_extraction.py          # Extraction runner script
├── generate_figures.py        # Generate 8 publication-quality color figures
├── create_report_docx.py      # Generate report DOCX (English & Japanese)
├── create_paper_docx.py       # Generate Scientific Data paper DOCX (English & Japanese)
├── create_figures_pptx.py     # Generate PPTX with one figure per slide
└── output/
    ├── extracted_data.csv             # Complete dataset (10,000 papers, 35 variables)
    ├── sampled_pmids.csv              # Sampled PubMed IDs with stratum assignments
    ├── sampling_stats.csv             # Stratum-level sampling statistics
    ├── summary_stats.json             # Summary statistics (overall & by stratum)
    ├── sampling.log                   # Sampling process log
    ├── report_english.docx            # Research report (English)
    ├── report_japanese.docx           # Research report (Japanese)
    ├── paper_scientific_data_english.docx   # Scientific Data article (English)
    ├── paper_scientific_data_japanese.docx  # Scientific Data article (Japanese)
    ├── figures_english.pptx           # Figures PowerPoint (English)
    ├── figures_japanese.pptx          # Figures PowerPoint (Japanese)
    └── figures/
        ├── fig1_software_rates_by_field.png
        ├── fig2_top_commercial_software.png
        ├── fig3_version_and_availability.png
        ├── fig4_replication_costs.png
        ├── fig5_software_landscape.png
        ├── fig6_version_availability.png
        ├── fig7_software_heatmap.png
        └── fig8_pmc_impact.png
```

## Methodology

### Sampling Strategy
- **Source:** PubMed E-utilities API
- **Period:** January 2020 -- March 2026
- **Method:** Stratified random sampling across 7 research fields
- **Year-split strategy** to work around PubMed's retstart limit (max offset 9,999)

### Seven Research Strata
1. Biomedical Basic (Molecular Biology, Genetics, Biochemistry, etc.)
2. Clinical Medicine (Therapeutics, Diagnosis, Clinical Trials, etc.)
3. Chemistry & Materials (Chemistry, Materials Science, Nanotechnology, etc.)
4. Physics & Engineering (Physics, Biomedical Engineering, etc.)
5. Social & Behavioral (Psychology, Epidemiology, Public Health, etc.)
6. Computational Science (AI, Machine Learning, Bioinformatics, etc.)
7. Environmental & Earth (Environmental Sciences, Ecology, Climate, etc.)

### Software Detection
- 95+ regular expression patterns covering 45 commercial and 50 open-source tools
- Detection on both abstracts and PMC full-text Methods sections
- Version number extraction within 80 characters of software mention
- Version availability assessment based on vendor policy database

### Vendor Policy Survey
Systematic survey of legacy version access policies for 30 major commercial software tools, finding that **no vendor** explicitly offers "reproducibility licenses" for verification purposes.

## Requirements

```
Python >= 3.10
pandas
requests
tqdm
matplotlib
seaborn
python-docx
python-pptx
Pillow
lxml
```

## Usage

```bash
# 1. Run PubMed sampling
python pubmed_sampler.py

# 2. Run data extraction (if sampling and extraction are separate steps)
python run_extraction.py

# 3. Generate figures
python generate_figures.py

# 4. Generate report DOCX (English & Japanese)
python create_report_docx.py

# 5. Generate Scientific Data paper DOCX (English & Japanese)
python create_paper_docx.py

# 6. Generate figures PPTX (English & Japanese)
python create_figures_pptx.py
```

## Target Journal

**Scientific Data** (Nature Portfolio) -- Article format

## License

This project is provided for academic and research purposes.

## Citation

If you use this dataset or code, please cite:

> [Author names]. The Hidden Cost of Reproducibility: Commercial Software Dependency in Published Research and the Version Accessibility Gap. *Scientific Data* (in preparation).
