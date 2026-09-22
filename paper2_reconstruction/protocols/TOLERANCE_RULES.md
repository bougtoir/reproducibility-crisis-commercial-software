# Target-specific agreement

**DRAFT — NOT FROZEN; generic scalar comparison is implemented, other scorers pending.**

No universal percentage tolerance. Store both target and observed values,
units, absolute error and meaningful relative error; relative error near zero
is not informative. Review reporting precision rather than demanding hidden digits.

| Target | Prospective rule |
|---|---|
| Counts | Exact, unless publication explicitly defines stochastic/approximate counts. |
| Deterministic scalar/effect | Agreement within the displayed rounding bin. Default nearest rounding uses [v−0.5*10^-d, v+0.5*10^-d); if truncation or another convention is explicit, use it. Declare ambiguous boundary cases prospectively. |
| Interval | Point estimate, lower bound and upper bound separately; all required components meet their individual precision rules. Inferential interpretation remains separate. |
| P value | Preserve exact reported precision or reported inequality. Significance category alone does not meet an exact-value target. |
| Proportion | Numerator and denominator when reported; otherwise displayed precision and units. |
| Prediction/performance | Identical prespecified metric, evaluation population and split. Unknown split identity is not cured by a similar score on different data. |
| Table | Prospectively identified central cells; all must meet their cell rules. Retain element errors. |
| Figure | Prefer underlying values. If extraction is unavoidable, freeze coordinates, extraction uncertainty and comparison before the run; no default pixel similarity. |
| Stochastic estimand | Compare the reported estimand/distributional summary under a prespecified criterion. Pilot fixes simulation count and error control within the resource ceiling. No optional rerunning until agreement. |
| Optimization | Objective value and constraints first. Exact decision variables only for a unique solution; equivalent optima are not failures. |

At pilot, reject an unscoreable target and proceed down the hierarchy before
main runs; log the reason. No stochastic, figure or non-scalar target may be
launched with an unresolved comparison rule. Numerical agreement does not itself
prove faithful algorithm implementation. Review fidelity using the published
method without introducing alternative analyses.
