# OMML equation audit — FINAL NON-HUMAN QC PATCH

Requirement: all true mathematical expressions in Word deliverables are native
Word equations (m:oMath / OMML), not plain-text pseudo-equations.

## Result

- OMML equation objects in the built supplement: 29 (main manuscript
  and editable-tables contain no true equations — verified by text scan).
- Plain-text remnants of equation patterns (S_hat, sum_h, pi_h, >=, <=, ^,
  alpha, |E|, |U|, t=, _h subscripts): 0 — verified by regex scan of every
  paragraph and table cell in all three DOCX deliverables.
- Converted constructs: accents (Ŝ, ȳ), subscripts (N_h, n_h, K_h, π_h, L_E,
  U_E), fractions (π_h = n_h/N_h, Ŝ/|E|, α/(2H)), n-ary summation (Σ_h),
  absolute-value delimiters (|E|, |U|, |E|+|U|), bracketed intervals
  ([L_E/(|E|+|U|), …] and [v−0.5×10^−d, v+0.5×10^−d)), inequalities (k ≥ t,
  k + m < t, t = 1/2/3, ≤ in the freeze-ordering constraint), Greek letters.

## Not converted (by design)

- Code identifiers, file paths, JSON keys, globs and CLI tokens in backticks
  (e.g. `FILE_CAP`, `accessibility_gate_failed_*.json`, `NO_ANALYSIS_EXECUTED`,
  `G1 = yes`, `G2 != no` gate-state comparisons). These are literal code/state
  strings, not mathematical expressions; marking them up as equations would
  corrupt their meaning.

STATUS: PASS. Map in OMML_EQUATION_MAP.csv.
