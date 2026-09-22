# Failure taxonomy

**DRAFT — NOT FROZEN**

| Code | Category |
|---|---|
| F01 | input_unavailable |
| F02 | commercial_dependency |
| F03 | software_or_version_unavailable |
| F04 | algorithm_underspecified |
| F05 | parameter_missing |
| F06 | preprocessing_unspecified |
| F07 | data_transformation_ambiguous |
| F08 | software_default_unspecified |
| F09 | randomization_or_seed_unspecified |
| F10 | numerical_implementation_ambiguous |
| F11 | hidden_manual_step_suspected |
| F12 | dependency_or_environment_failure |
| F13 | computational_resource_limit |
| F14 | publication_internal_inconsistency |
| F15 | agent_reasoning_or_coding_failure |
| F16 | target_result_insufficiently_reported |
| F17 | firewall_contamination |
| F18 | input_identity_or_version_ambiguous |
| F19 | access_or_download_failure |
| F20 | runs_but_target_not_reproduced |
| F99 | other |

Record primary and optional secondary codes, earliest failed level, source
location or execution evidence, observed event, hypothesized attribution and
certainty separately. Allow uncertain, multiple and not_classifiable. For
simultaneous unresolved causes do not force an arbitrary primary code.

A documented failed download is F19, not proof of F01. An incorrect agent program
is not automatically F04. A difference from author code after reveal is descriptive
evidence, not proof that the publication made reconstruction impossible. F11
remains suspected unless evidence establishes a hidden step. F17 invalidates the
slot for clean inference; it is not an ordinary failed reconstruction.
