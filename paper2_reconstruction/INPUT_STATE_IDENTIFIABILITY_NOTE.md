# Input-state identifiability (conceptual note)

Status: conceptual refinement of the ACCESS stage of the Paper II framework. It changes no gate outcome, no barrier state, no sealed blind outcome, no A/B/C/E record and no Section D status. Paper II did not measure dataset drift; nothing here is an empirical result.

## Definitions (provisional)

- **Data availability**: whether a named resource (repository, accession, endpoint, dataset name) can be accessed at all.
- **Input-state identifiability**: the extent to which the exact state of the input data used in an analysis can be uniquely identified from the publication and associated records.
- **Historical-state retrievability**: the extent to which that exact previously used data state can still be obtained by an independent third party.

A focused bibliographic audit (Crossref/DataCite metadata, 2026-09-25; see `review/INPUT_STATE_CITATION_VERIFICATION.md`) found the underlying ideas well established in the FAIR, data-citation, dynamic-data and provenance literature, but did not find established use of these exact phrases. They are therefore used provisionally and no priority or originality is claimed. "Data-state reproducibility" was not adopted because no established compatible usage was found.

Key statement: **Public availability is not equivalent to reproducible input identity.**

Three things that are easy to conflate:

```text
same repository / accession / URL
≠ same version
≠ same bytes
```

## Relation to ACCESS

No new top-level stage is added. ACCESS is refined internally:

```text
ACCESS
  ├─ resource accessible?
  ├─ exact input state identifiable?
  └─ historical input state retrievable?
RECONSTRUCT → EXECUTE → REPRODUCE → ROBUST
```

A paper can pass the first question while failing the second (the publication names GEO series X and "the human genome" but no release), or pass the first two while failing the third (release named, but the repository no longer serves that release). Externally maintained resources may be updated, corrected, reannotated, restructured, or answer a query differently over time while keeping the same identifier. This does not mean every same-accession resource changes; it means that availability alone does not establish that the exact historical analytical input is identifiable or retrievable.

Three abilities are distinct and should not be substituted for each other:

| Ability | Who holds it | What it does not imply |
|---|---|---|
| redistribute the input | often *not* the authors for third-party data (licence/legal limits) | inability to redistribute ≠ inability to identify |
| identify the exact state | authors, by reporting | identification ≠ retrievability |
| retrieve the historical state | depends on the repository keeping versions | a hash verifies bytes that exist; it cannot recover bytes that are gone |

## Relation to the EPJ software-version work

The EPJ Research Infrastructures antecedent (Onishi and Ikenoue, 2026) concerns commercial software dependency and the version accessibility gap: the software named in a paper may no longer be obtainable in the version used. Dataset-version drift is analogous but distinct. Both separate a nominally available resource from the state actually used. Versioned software can, however, often be redistributed, rebuilt or emulated, whereas an exact third-party data state may be neither redistributable by the authors nor recoverable from the source. Paper II inherits the software concern and adds the data-state concern as a conceptual implication and a limitation on reconstructability; it does not quantify either drift.

## What Paper II actually recorded

The frozen main-study records (Table 11 in the editable tables; `results/input_state_descriptives.csv`; `data/adjudication/main-study-20260925/input_state_ledger.csv`) contain, as recorded:

- repository or accession for every sampled paper (delegated specification listings);
- a file list for most papers;
- external public reference resources named by the delegated specifier for about half of the specified papers, with a minority of entries lacking a version in the publication text;
- the frozen barrier state "external public reference resource required, not served" for 28 papers — a protocol-defined state that this note leaves unchanged;
- for the 21 served inputs of the 10 attempted papers: URL, UTC retrieval timestamp, byte count, SHA-256, whether the repository announced a checksum and whether it matched; **no** repository version/release identifier (live-service snapshots), **no** schema/release identifier, and **no** historical-version check.

Fields that were not recorded are reported as "not recorded" or "not assessed". None were reconstructed after the fact.

## Reporting recommendation (proposed, not validated)

For third-party or public input data, report where available:

1. repository or accession;
2. release/version;
3. retrieval date/time;
4. exact query or API parameters;
5. file names/list;
6. schema/version;
7. cryptographic checksum;
8. transformation log (raw → analytical input);
9. licence/redistribution status (why an exact snapshot can or cannot be redistributed).

Table 10 (`results/requirement_stage_table.csv`) assigns each element to the framework stage it serves and retains the existing rows for code availability, software/version reporting, container/environment, detailed Methods, parameters/defaults, random seed, resource requirements, independent reconstruction and robustness testing.

## Claim calibration

Prefer: "Public availability does not by itself establish that the exact historical analytical input remains identifiable or retrievable."

Do not write: "Public datasets are not reproducible"; do not imply that all same-accession resources change, that a hash alone solves versioning when the historical bytes are unavailable, that authors can always archive third-party data, or that Paper II quantified dataset drift.

## Sources

FAIR principles (Wilkinson et al., 2016); Joint Declaration of Data Citation Principles (2014); RDA dynamic-data citation recommendations (Rauber et al., 2016; Pröll and Rauber, 2013; Rauber et al., 2021); dataset versioning framework (Klump et al., 2021); provenance (Pasquier et al., 2017); content-based identifiers (Di Cosmo et al., 2020); reproducibility guidance (Sandve et al., 2013; Stodden et al., 2016); reference-resource dependence of results (Zhao and Zhang, 2015). Verification records: `data/verified_references.csv`, `review/INPUT_STATE_CITATION_VERIFICATION.md`.
