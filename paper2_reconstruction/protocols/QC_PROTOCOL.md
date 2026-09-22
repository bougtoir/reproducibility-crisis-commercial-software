# Quality-control protocol

**DRAFT — NOT FROZEN**

First perform a reviewer-perspective critique of question/novelty, statistical
design, figures/tables, reproducibility and claim strength. Rank each issue by
scientific severity, expected benefit and feasibility. Mechanical formatting
cannot overrule a fatal scientific gap.

Required final audits:

| Audit | Evidence required |
|---|---|
| Fabrication | Every empirical row from real source or observed run; empty human/agent tables not scored. |
| Numerical | Manuscript/table/figure values derived from canonical results and raw source hashes. |
| Denominator | Unique paper versus source row; at-risk versus full frame; missing and contaminated slots; sample weights. |
| Figures/tables | Sequential citations, no orphan elements, English labels, editable outputs and journal-appropriate placement. |
| Methods/code | Every claimed method implemented and exercised; no unperformed freeze, interval or validation claims. |
| Endpoint | Target and tolerance sealed before run; no post-hoc easiest target. |
| Firewall | Qualified broker and control-plane restriction, access telemetry, canary failures, contamination adjudication. |
| Independence | Separate contexts and mutable storage; no parent/sibling debugging. |
| Commercial policy | Software invoices/entitlements or explicit zero-purchase evidence; compute costs distinct. |
| Failure attribution | Observations and causal explanations distinguished; uncertain categories preserved. |
| Human blinding | Genuine participation, ethics/consent determination, masked packages and scores. |
| Reveal | Frozen outcomes precede original-code access; no perturbation or revised score. |
| References/novelty | Primary source, version, DOI, status and supported claim; no unsupported “first”. |
| Discussion/abstract | No inference from agent failure to human impossibility; no unavailable results. |
| Supplement | All protocol versions, deviations, schemas, rights/access paths and reproducible code. |
| Rebuild | Clean public checkout and stated inputs regenerate every submitted result without private files. |

Automated checks currently cover immutable sources, duplicate conflicts,
identity/membership preservation, unknown-aware slot logic, implementation/
execution/numerical consistency, timestamp ordering, scalar rounding,
sampling invariance and reference interval calculations. These are software
checks, not validation of source labels or the unimplemented experimental harness.

Before submission, repeat targeted literature and live journal-guideline checks,
generate the final empirical manuscript from the frozen study data, review all
author declarations, render every document and manually inspect it, and archive
a versioned lawful release. `submission-check` must remain blocked until those
requirements are implemented and evidenced; never weaken it to make a build green.
