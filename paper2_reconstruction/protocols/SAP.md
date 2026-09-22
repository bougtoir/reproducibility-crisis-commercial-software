# Statistical analysis plan

**DRAFT — NOT FROZEN. No main-study outcomes have been inspected.**

## Estimands and denominators

The population is the author-approved unique-PMID frame derived from the fixed
deposit, not all publications. Its exact count, ordered CSV hash and source/decision
provenance are in `results/frame_manifest.json`. Preserve source rows and all recorded
field memberships in the separate bridge; no replacement publications are added.
Let E denote G1=yes papers and U denote G1=uncertain papers. The proposed primary
estimand is end-to-end policy demonstration among E, with all verified access
barriers retained. For an assessable attempted paper the endpoint is at least
two valid successes among three fixed slots. A verified indispensable gate
barrier is policy non-success, with `not_attempted_gate_barrier`; it does not
produce three failed runs. Unresolved gates yield an unknown endpoint.
The conditional rate among attempted papers is secondary, with its own denominator.
G5 insufficiency alone must not select difficult but attemptable papers out.

For an assessed gate, report yes, no and unknown/unassessed counts over both the
original paper set and the actual conditional risk set. Supply identification
bounds [yes/N, (yes+unresolved)/N] and the resolved-case rate separately. Do not
multiply marginal gate rates to construct a funnel. Full-frame census proportions
have no paper-sampling uncertainty; any binomial interval must be labelled a
model-based reference, not a census sampling confidence interval. Classification
uncertainty and validation sampling uncertainty require separate reporting.

## Sampling

Use disjoint assignment before looking at outcomes: among each paper's recorded
field labels choose the label with minimum SHA256 of
`paper2-field-v1|paper_id|field`. Retain every original membership in the bridge.
After removing pilot cases and deciding eligibility, allocate about 100 main
papers across nonempty assigned fields. Set one place per nonempty field, then
assign each remaining place to the field with largest proportional allocation
deficit, respecting its capacity; resolve ties lexicographically as implemented.
Within field rank by SHA256 of `paper2-main-v1|paper_id`.
Selection is reproducible and has equal within-field probability under the
prespecified pseudorandom ranking model. Save seed, ordered frame hash, N_h, n_h,
pi_h=n_h/N_h and weight N_h/n_h. No pilot or main selection has occurred.

This field-stratified design is the conservative default. Do not add commercial,
code or data cross-strata unless the pilot shows adequate occupancy and the
amended allocation is frozen before outcomes. Software categories are
commercial_only/open_only/both/neither_detected/unknown; a negative regex match
does not establish absence. Known policy barriers can be tallied by census;
sample the attemptable remainder without removing it from the headline denominator.

## Precision and estimates

The build calculates Wilson intervals for planned n=100 at assumed p of 0.20,
0.30, 0.50, 0.70 and 0.80. These are analytic unweighted binomial reference
calculations, not study results or guarantees for the stratified design.

For the attemptable stratum populations, estimate the success total as
S_hat = sum_h N_h*ybar_h. The all-eligible policy estimate is S_hat/|E|,
retaining verified indispensable barriers as policy non-success in |E|.
The conditional attemptable-domain estimate is S_hat/sum_h N_h and must
not be labelled the headline estimate. Eligible papers with unresolved gates
remain in |E| and add zero/their count to lower/upper success-total bounds.
For finite-population uncertainty with small strata or boundary successes, invert
the hypergeometric distribution for the possible number of successful papers
K_h in each stratum. Use two tail probabilities alpha/(2H) for H sampled strata
and alpha=0.05; aggregate simultaneous lower and upper K_h bounds, then divide by
|E| for the headline or sum_h N_h for the conditional domain. Add unresolved-gate
counts to the upper headline bound. Label this as a sampling-and-identification
envelope when unknown gates remain. This conservative interval avoids zero estimated
variance in all-success/all-failure strata. Census strata have exact totals.
The estimator and interval are implemented in `src/paper2/estimation.py`.
Exhaustive enumeration checks finite-population coverage for all sample sizes
and success totals in populations up to 12; boundary, census, unequal-weight,
unknown-outcome and denominator checks are also included. These software tests
are not empirical study estimates; protocol freeze remains pending.

If a paper outcome is unresolved, calculate endpoint lower/upper estimates by
assigning unresolved outcomes 0/1 solely for identification bounds, explicitly
labelled. Obtain sampling bounds for each endpoint; do not present the resulting
envelope as an ordinary missing-at-random confidence interval. In E plus U,
worst-case eligibility bounds are [L_E/(|E|+|U|),
(U_E+|U|)/(|E|+|U|)], where L_E and U_E are eligible-set lower/upper success totals.
For sampled outcomes, carry sampling bounds into these limits.

## Run and paper summaries

Each valid run requires implementation, execution and target agreement.
For fixed slots, k = clean successes and m = missing/invalid slots.
Threshold t is achieved if k>=t, failed if k+m<t, otherwise unknown.
Use t=2 primary, t=3 strict and t=1 permissive.
Show 0/3–3/3 only for fully assessable clean triples, and display the incomplete
triple count alongside. Never reduce a contaminated triple to a new denominator.

Report each L1–L5 separately, including applicable conditional and unconditional
denominators. Preserve distinct point/interval/p-value/conclusion comparisons.
Failure frequency uses papers and runs as separate units; multiple codes may
sum above the denominator. Burden uses median, quartiles, range and missingness,
with durations of budget-censored runs identified rather than treated as fast
completed successes. Report software purchase, agent/API and compute costs separately.

## Models and sensitivities

Do not treat three runs as three independent papers. Descriptive estimates are
primary. A paper-random-intercept logistic model is optional only if at least
20 distinct papers with successes and 20 with failures support every candidate
predictor degree of freedom, clustering is identifiable, and no separation or
unstable fit is present. Freeze any reduced predictor set before outcome access;
never screen predictors by observed p values. Given the planned size and seven
fields, an adjusted multi-predictor model is likely unjustified and may be omitted.
A code-sharing association is observational, not causal; original code stays hidden.

Prespecified sensitivity descriptions: strict/majority/permissive; clean versus
all-assigned accounting without reclassifying contaminated runs as valid;
uncertain-target exclusion; lower/upper unknown bounds; field, software category
and code-detection subgroups. Report sparsity; do not fish for interactions.
No alternative models/defaults/implementations, perturbations, multiverse,
robustness scores or Paper III experiments are permitted.
