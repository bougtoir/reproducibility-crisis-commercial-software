# Pilot candidate adjudication request (human G1–G5, not machine output)

Scope: the lowest-ranked deterministic candidate per field whose article text is lawfully retained. Machine fields below are provisional proposals bound to retained source segments; the author must adjudicate G1 (computational claim), G2 (principal target), G3 (input access), G4 (specification sufficiency) and G5 (feasibility under the compute envelope) against the retained article before any pilot case is recorded. Leave any gate `unknown` rather than guessing. Do not consult the original code or repository.

## Biomedical_Basic — PMID:41032336 (rank 1)

- DOI: 10.1093/femsle/fnaf103
- Article text route: open_pmc_jats
- Machine proposal G1=yes G2=yes
- Proposed target: Maximum-likelihood cpn60 phylogeny (GTR + gamma + invariant sites, 500 bootstrap replicates) used to assign 48 orange-yellow environmental colonies to Pantoea species and detect a candidate new species.
- Inputs required (proposal): cpn60 (and some 16S rRNA) amplicon sequences from environmental isolates and reference strains, deposited in GenBank (PV699648-PV699692, PV688339) and supplied as FASTA supplementary material.
- Access observation: Sequence data are stated as deposited in GenBank under accession numbers PV699648-PV699692 and PV688339, and also available in supplementary material in FASTA format.
- Specification observation: The phylogeny is described with model, gamma categories, and bootstrap replicates, but no code, alignment files, or tree files are reported as shared; only sequences are stated as available.

| Gate | Author decision (yes/no/unknown) | Note |
|---|---|---|
| G1 | | |
| G2 | | |
| G3 | | |
| G4 | | |
| G5 | | |

## Chemistry_Materials — PMID:36358920 (rank 2)

- DOI: 10.3390/biom12111569
- Article text route: open_pmc_jats
- Machine proposal G1=yes G2=yes
- Proposed target: Differential urinary proteins and protein chemical modification types distinguishing high-fat-diet ApoE-/- mice from controls across time points, derived from label-free DIA/DDA proteomic quantification and open/restricted modification searches.
- Inputs required (proposal): Urine LC-MS/MS raw files (DDA and DIA) from experimental and control mice at seven time points; DDA search results (pdResult) and 10 DDA raw files for spectrum library; mouse UniProt database; iRT peptide sequences; sample group/time-point labels.
- Access observation: The article states mass spectrometry proteomics data were deposited to ProteomeXchange via iProX with dataset identifier PXD027610; this is a data-availability statement, not confirmation that files are retrievable or usable.
- Specification observation: Reporting includes screening thresholds (FC >=1.5 or <=0.67, p<0.05), KNN imputation, CV<0.3, FDR<1%, and random-combination permutation controls, but the abstract-only source omits these; full-text methods are needed to reproduce the pipeline.

| Gate | Author decision (yes/no/unknown) | Note |
|---|---|---|
| G1 | | |
| G2 | | |
| G3 | | |
| G4 | | |
| G5 | | |

## Clinical_Medicine — PMID:39730532 (rank 1)

- DOI: 10.1038/s41598-024-81563-z
- Article text route: open_pmc_jats
- Machine proposal G1=yes G2=yes
- Proposed target: Classification of Alzheimer's disease diagnostic groups (CN vs AD and sMCI vs pMCI) from structural MRI using a hybrid ML+DL model
- Inputs required (proposal): ADNI1 structural T1-weighted sMRI scans (256x256x168, 1mm spacing), preprocessed via CAT/SPM12 (denoising, bias correction, skull stripping, MNI152 registration, Brainnetome atlas parcellation into 246 regions); gray matter volumes and CNN-extracted features
- Access observation: The article states the dataset is publicly available via the ADNI access-data link (https://adni.loni.usc.edu/data-samples/access-data/), and the article is open access under CC BY-NC-ND.
- Specification observation: The paper reports architecture and hyperparameter tuning procedures and performance metrics, but does not report code availability, random seeds, or full reproducibility details; data availability is stated only as a public ADNI link.

| Gate | Author decision (yes/no/unknown) | Note |
|---|---|---|
| G1 | | |
| G2 | | |
| G3 | | |
| G4 | | |
| G5 | | |

## Computational_Science — PMID:36409083 (rank 3)

- DOI: 10.1128/msystems.00831-22
- Article text route: open_pmc_jats
- Machine proposal G1=yes G2=yes
- Proposed target: Genomic traits (lineage-specific genes such as paiB, cfim, hysA-VSaβ, hlb, and Opp system arrangement) associated with virulence and fitness phenotypes that may account for the success of epidemic S. aureus clones ST59 and ST398.
- Inputs required (proposal): Whole-genome sequencing data (Illumina and Nanopore) from 142 S. aureus isolates; core gene alignment; phenotypic measurements (proteolysis, hemolysis, qPCR expression, neutrophil assays, adhesion/invasion); metadata on ST, spa type, SCCmec, resistance genes.
- Access observation: Raw Illumina short-read data from 157 newly generated S. aureus isolates are stated as available on NGDC under BioProject PRJCA012518; metadata in Data Set S1.
- Specification observation: The abstract and article describe comparative genomic and phenotypic analyses but do not report a single pre-specified computational model, algorithm, or statistical pipeline for the central claim; methods list tools (SPAdes, Roary, IQ-TREE, etc.) but no unified reproducible computational workflow is specified.

| Gate | Author decision (yes/no/unknown) | Note |
|---|---|---|
| G1 | | |
| G2 | | |
| G3 | | |
| G4 | | |
| G5 | | |

## Environmental_Earth — PMID:39472468 (rank 9)

- DOI: 10.1038/s41598-024-76149-8
- Article text route: open_pmc_jats
- Machine proposal G1=yes G2=yes
- Proposed target: Detection of rare non-native (BTS) alleles in pooled DNA samples via a custom Random Forest classifier applied to Fluidigm SNP-type allele fluorescence intensities, with accuracy/sensitivity/specificity evaluated against exon-capture sequence data.
- Inputs required (proposal): Fluidigm SNP-type assay endpoint fluorescence intensities (raw and control-corrected X- and Y-allele intensities) per SNP per DNA-pool; known-genotype training pools at defined RARatios; exon-capture sequence data for validation; R packages randomForest, lme4, BayesianFirstAid.
- Access observation: The article is open access (CC BY) and states that additional datasets generated and analyzed are available from the corresponding author on reasonable request; a full description of the genotype assay is in the supplemental materials.
- Specification observation: Data availability is stated only as available from the corresponding author on reasonable request; no repository accession or code-sharing statement is reported in the supplied text.

| Gate | Author decision (yes/no/unknown) | Note |
|---|---|---|
| G1 | | |
| G2 | | |
| G3 | | |
| G4 | | |
| G5 | | |

## Physics_Engineering — PMID:34031721 (rank 2)

- DOI: 10.1007/s00259-021-05387-z
- Article text route: open_pmc_jats
- Machine proposal G1=yes G2=yes
- Proposed target: Video-based marker tracking and computer-assisted movement analysis to quantify surgical dexterity and decision-making during robotic radioguided surgery.
- Inputs required (proposal): Endoscopic video with visible tracking markers, marker geometry, camera calibration parameters, and recorded task/trial metadata.
- Access observation: The article states code and data are available on reasonable request; the article is open access under CC BY.
- Specification observation: The article reports preprocessing steps and movement features but does not provide code, exact marker geometry, or full parameter settings.

| Gate | Author decision (yes/no/unknown) | Note |
|---|---|---|
| G1 | | |
| G2 | | |
| G3 | | |
| G4 | | |
| G5 | | |

## Social_Behavioral — PMID:33574863 (rank 4)

- DOI: 10.33073/pjm-2020-037
- Article text route: open_pmc_jats
- Machine proposal G1=yes G2=yes
- Proposed target: Alkaline-lignin degradation rate and Lip/Mnp enzyme activity dynamics of four bacterial strains, plus GC-MS-based inference of DBP aerobic metabolic pathway from degradation products
- Inputs required (proposal): OD600 measurements of alkaline-lignin liquid cultures, colony diameters, Lip/Mnp absorbance readings (310 nm, 465 nm) with extinction coefficients and reaction volumes/times, GC-MS peak tables (RT, compound presence/absence across strains and control)
- Access observation: The article is open access under CC BY-NC-ND 4.0 and includes supplementary material links, but no data or code repository is stated.
- Specification observation: Enzyme activity formulas and statistical methods (one-way ANOVA, Duncan) are reported, but raw OD/absorbance values, replicate-level data, and GC-MS peak intensities are not provided; degradation rate calculation from OD is not fully specified.

| Gate | Author decision (yes/no/unknown) | Note |
|---|---|---|
| G1 | | |
| G2 | | |
| G3 | | |
| G4 | | |
| G5 | | |

## Prospectively unavailable earlier candidates

Earlier-ranked candidates in each field lack lawfully retained text after the open PMC and OpenAlex route checks (see `pilot_candidate_screening.json`). They stay in the frame with eligibility unknown; they are not G1 negatives and not input-unavailable verdicts.
