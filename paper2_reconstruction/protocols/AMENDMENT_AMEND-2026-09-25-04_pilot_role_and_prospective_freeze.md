# Protocol amendment AMEND-2026-09-25-04: pilot role, investigator verification and prospective target freeze

**Status: investigator-instructed amendment, hash-frozen before any candidate
selected under AMEND-2026-09-24-03 is verified, frozen or reconstructed.
Applies to every pilot and main-study record written after the freeze.**

## Basis

On 2026-09-25 the principal investigator instructed that (1) the seven
candidates selected under AMEND-2026-09-24-03 are a prospective pilot /
pipeline-validation set and not the main analysis sample; (2) G1–G5
verification and a prospective target specification must be complete for all
seven before any blind reconstruction begins; (3) the verifier is the
research team, so records must say "investigator verification" or "human
adjudication", not "author verification", and original authors are not
contacted for the primary reconstruction protocol; (4) accessibility and
screening failures stay in the full-corpus attrition layer with
machine-readable reasons; (5)–(6) the deposited-data and leakage rules of
AMEND-2026-09-24-03 remain in force; (7) ten fields are frozen per pilot paper
before reconstruction; (9) pilot results are used only to find procedural
defects, never to raise the observed success rate; (12) no Paper III analysis
is performed.

## 1. Terminology

`investigator_verification` replaces `author_verification` as the key and
phrase for the research team's review of delegated (Devin) G1–G5 gates.
Records written before this amendment keep their original key; readers accept
both. "Author" is reserved for the original study authors and for the
investigators' own authorship of this study. The primary reconstruction
protocol never contacts original authors; author clarification is not an
allowed resource.

## 2. Role of the seven AMEND-2026-09-24-03 candidates

The candidates selected under AMEND-2026-09-24-03 form the
`PROSPECTIVE_PILOT` set: `main_sample_membership =
excluded_from_primary_denominator`. Their outcomes validate the pipeline
(target rules, tolerance rules, failure taxonomy, logging schema, resource
ceilings) and are reported descriptively. They enter no primary
reconstruction-rate numerator or denominator. The main cohort is drawn after
the post-pilot protocol freeze from the full eligible population by the SAP's
precision calculation and stratified random sampling (target N ≈ 100), with
pilot papers removed from the sampling frame.

## 3. Additional exclusion class

| class | meaning |
| --- | --- |
| `deposited_input_mismatch` | the named deposit resolves anonymously but its listed samples/files are not the datasets the article reports analysing (identifier sets disjoint or incomplete); the target inputs are therefore not deposited under the named route |

The class is recorded with the article's own identifiers and the deposit's
listed identifiers as evidence. It is a sampling-frame restriction like every
AMEND-2026-09-24-03 class, not a reconstruction outcome, and the deterministic
per-stratum chain resumes at the next rank.

## 4. Prospective freeze fields (per pilot and per main-study paper)

Before any slot is opened, the target manifest records, and the freeze hashes:

1. `exact_target` — the numerical value(s) exactly as published, with units
   and reporting precision;
2. `target_location` — section, table, figure or sentence, with the retained
   segment ID and verbatim quote;
3. `required_inputs` — repository, accession, file names, versions or
   run/sample identifiers, bytes and SHA-256 where retained;
4. `allowed_resources` — free software, public references and the package
   contents served to every slot;
5. `blocked_original_artifacts` — original code, repositories, notebooks,
   containers, processed result files and author clarification, all withheld;
6. `comparison_metric` — the TOLERANCE_RULES scorer applied;
7. `numerical_agreement_criterion` — the prospective numeric rule with
   reporting precision;
8. `conclusion_criterion` — the minimal target-linked claim and how agreement
   or disagreement with it is decided;
9. `resource_ceiling` — wall-clock, model-call, token and byte ceilings per
   slot;
10. `stopping_rule` — what ends a slot (report, ceiling, unresolved input,
    specification gap) and how each end state is labelled.

Known specification gaps (for example an undisclosed train/test split) are
written into the manifest before the slot opens; a gap is a recorded
`specification_insufficient` state, never a reason to substitute an easier
target after the run.

## 5. Investigator verification record

For every pilot paper the investigator team records, gate by gate, `confirm`
or `overturn` of the delegated value with a reason, after re-reading the cited
retained segments and re-checking every quote verbatim. Overturns are kept
alongside the delegated value. The record is validated against the retained
text and the amended record before the target freeze may reference it.

## 6. Paper III boundary

No multiverse analysis, no deliberate alternative reasonable implementation,
no robustness perturbation and no test of whether an undocumented
implementation choice changes a conclusion is performed in Paper II. The
original-code reveal after the blind freeze is descriptive only.

## Sequence

1. Freeze this amendment (SHA-256 + RFC 3161 receipt).
2. Complete deposit inspection, investigator G1–G5 verification and the ten
   freeze fields for all seven pilot candidates; resume a stratum's chain
   where a candidate is excluded.
3. Freeze the pilot target manifest (SHA-256 + RFC 3161), then run three
   independent blind slots per paper.
4. Pilot review; freeze the final protocol and SAP; select the main cohort.
