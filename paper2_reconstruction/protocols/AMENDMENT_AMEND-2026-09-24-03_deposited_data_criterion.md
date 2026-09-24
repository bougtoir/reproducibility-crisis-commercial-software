# Protocol amendment AMEND-2026-09-24-03: deposited-data eligibility criterion

**Status: author-instructed amendment, hash-frozen before any candidate is
screened under it. Applies to every candidate selected after the freeze.**

## Basis

The provisional pilot (`DEV-2026-09-24-02`) reached no reconstruction attempt:
none of the seven provisional papers had study inputs that a third party could
obtain lawfully and immediately. On 2026-09-24 the author instructed that the
main analysis of publication-grounded reconstruction be restricted to papers
whose deposited analysis data are lawfully and immediately obtainable by a
third party, that every excluded paper be retained with its exclusion reason,
and that the seven provisional papers be moved to a separate accessibility
layer without any reconstruction success or failure rate.

## Amended eligibility (replaces the G3-as-outcome rule for the main sample)

A paper is eligible for the main reconstruction sample only if, in addition to
lawfully retained article text, primary G1 = yes and G2 != no, all of the
following hold at screening time and are recorded with retained evidence:

1. **Deposited analysis data are named in the article.** The article or its
   open supplement names a repository accession, persistent identifier or
   supplementary data file that holds the data consumed by the target
   computation (not only the outputs of that computation).
2. **Third-party lawful access is immediate.** The listing endpoint of the
   deposit answers an anonymous request with a machine-readable file list.
   No account approval, data-access committee, data-use agreement, author
   contact, institutional licence or payment may stand between a third party
   and the bytes. Free registration that is granted automatically is
   acceptable only if the file list itself is anonymous.
3. **The deposit is practically retrievable.** Every file needed for the
   target computation has an announced size, each file is at most
   `FILE_CAP = 2 GiB`, and the needed files total at most `TOTAL_CAP = 8 GiB`.
   Deposits that announce no sizes, or whose needed files exceed either cap,
   are excluded as practically unretrievable.
4. **The deposit is not outcome leakage.** Files that directly contain the
   target value (figure source tables, result tables) do not qualify as
   analysis data; they are withheld from every solver and held for
   post-reveal validation.

## Exclusion classes (retained, never silently dropped)

| class | meaning |
| --- | --- |
| `no_lawful_text_retained` | article text not lawfully retained; unadjudicated, not negative |
| `no_deposit_named` | no accession, persistent identifier or data file detected in the retained text |
| `author_request_only` | data stated as available on request or from the corresponding author |
| `application_required` | controlled-access repository or consortium (e.g. ADNI, dbGaP, EGA, UK Biobank) |
| `institutional_licence` | access depends on an institutional contract or licence |
| `listing_unavailable` | named route returned no anonymous machine-readable file list |
| `size_unannounced` | route lists files without sizes |
| `above_cap` | needed files exceed `FILE_CAP` or `TOTAL_CAP` |
| `deposited_output_only` | the only deposited files are the target outputs (outcome leakage) |
| `code_only_deposit` | the only named deposit is source code (blocked by the firewall) |
| `primary_G1_no` / `primary_G2_no` | fails the unchanged computational gates |

Exclusion under this amendment is a **sampling frame restriction**, not a G3
study outcome. The main-sample estimand therefore becomes reconstruction
success among computational papers with immediately obtainable deposited
analysis data, and every manuscript statement of the estimand must say so.

## Accessibility layer

Papers excluded for access reasons after G1 = yes and G2 != no are retained in
`data/adjudication/accessibility_gate_failed_*.json` with class, route
evidence and the deposit statement segment IDs. The seven provisional pilot
papers are the first members of this layer (`ACCESSIBILITY_GATE_FAILED`). This
layer is an audit sample of pre-reproducibility access barriers for an
auxiliary analysis or a separate paper. **No reconstruction success or failure
rate may be computed from it**, because no member reached a reconstruction
attempt; the validation code rejects any rate field in the layer record.

## Sequence

1. Freeze this amendment (SHA-256 + RFC 3161 receipt) before any screening.
2. Screen candidates in the unchanged deterministic order per stratum,
   retaining every listing response and every exclusion reason.
3. Record primary G1–G5 for the first deposit-eligible candidate per stratum
   as a delegated assessment pending author verification.
4. Only then freeze pilot selection under the amended criterion.
