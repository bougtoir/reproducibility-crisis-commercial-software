# Documentation staleness audit (revision pass)

Searched root docs and `review/*.md` for claims contradicted by the frozen
main-study records ("no main runs", "no pilot", "no frozen outcomes",
"no reconstruction rates", "no candidate attempted", "funnel NOT_STARTED").

| File | Stale claims found | Action |
|---|---|---|
| README.md | Title block + body state the pre-execution posture ("no reconstruction rates … are generated") | Added dated "Current state" banner at top pointing to canonical status; body retained as provenance of the staged protocol |
| review/EVIDENCE.md | Multiple "no pilot / not frozen / no reconstruction attempted" sections written pre-execution | SUPERSEDED banner inserted after title |
| review/REVIEWER_AUDIT.md | Pre-main-study review ("HOLD" verdict, no frozen outcomes) | SUPERSEDED banner; successor = review/HOSTILE_REVIEW.md |
| WITHOUT_D_STATUS.md | Current — regenerated facts match canonical values | Minor: unchanged, still authoritative |
| INPUT_STATE_IDENTIFIABILITY_NOTE.md | Current | unchanged |

Policy applied: superseded documents are banner-marked and cross-linked, not
deleted (they are provenance for the staged protocol history). Canonical
status order: `results/manuscript_values.csv` > `WITHOUT_D_STATUS.md` >
`results/readiness.json` > prose docs.
