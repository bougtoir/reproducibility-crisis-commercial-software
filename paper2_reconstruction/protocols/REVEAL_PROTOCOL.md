# Blind freeze and descriptive reveal

**DRAFT — NOT FROZEN; original corpus-paper implementations have not been revealed.**

Close all main blind slots, adjudicate and freeze outcomes/codes/comparisons,
then hash code, inputs, outputs, public action logs, access events and manifests.
Record byte counts, UTC freeze, independent timestamp receipt and artifact IDs
in `blind_freeze_manifest.csv`. Seal human blind packages before reveal.

The required ordering is target selected < run start < run stop <= blind freeze
< original-code reveal. No missing timestamp may be invented or backdated.
A checksum alone proves content identity, not when the content existed.

A separate reveal reviewer may then inspect original code where lawfully
available. Compare publication, independent implementation and original choices
for normalization, filters, missing-data handling, initialization, seeds,
stopping, defaults, preprocessing order, exclusions, manual transformations,
constants, boundaries, data cleaning and versions. Use explicit/partial/
ambiguous/not_reported for description completeness. Preserve source evidence.

Reveal is descriptive only. Do not rerun with original choices, substitute
original code, perturb any discrepancy or revise the frozen blind score.
Unavailable original code yields a missing reveal assessment, not a negative
description-quality finding. Human validators must never see reveal material.
