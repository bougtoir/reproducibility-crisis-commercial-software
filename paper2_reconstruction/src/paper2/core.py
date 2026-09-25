import csv
import hashlib
import math
from collections import Counter, defaultdict
from collections.abc import Mapping, Sequence
from dataclasses import dataclass
from datetime import datetime
from decimal import Decimal
from pathlib import Path
from typing import Literal

State = Literal["yes", "no", "unknown", "not_assessable"]
PaperState = Literal["success", "failure", "unknown"]
Row = dict[str, str]
STATES = {"yes", "no", "unknown", "not_assessable"}
FAILURES = {f"F{number:02d}" for number in range(1, 21)} | {"F99"}
FIELDS = (
    "Biomedical_Basic",
    "Clinical_Medicine",
    "Chemistry_Materials",
    "Physics_Engineering",
    "Social_Behavioral",
    "Computational_Science",
    "Environmental_Earth",
)


def sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def verification_status(record: Mapping[str, object]) -> str:
    """Read the human-verification state of a gate record.

    Records written before amendment AMEND-2026-09-25-04 (terminology) spell the key
    `author_verification`; they are hash-bound evidence and are read, not rewritten.
    """
    for key in ("investigator_verification", "author_verification"):
        if key in record:
            return str(record[key])
    raise ValueError("record carries no investigator verification state")


def read_csv(path: Path) -> list[Row]:
    with path.open(newline="", encoding="utf-8-sig") as handle:
        return list(csv.DictReader(handle))


def write_csv(path: Path, rows: Sequence[Mapping[str, object]], columns: Sequence[str]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(
            handle, fieldnames=columns, extrasaction="raise", lineterminator="\n"
        )
        writer.writeheader()
        writer.writerows(rows)


def snapshot(path: Path, data: bytes) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    if path.exists():
        if path.read_bytes() != data:
            raise ValueError(f"Immutable snapshot differs: {path}")
    else:
        with path.open("xb") as handle:
            handle.write(data)


def paper_registry(rows: Sequence[Row]) -> tuple[list[Row], list[Row]]:
    groups: dict[str, list[tuple[int, Row]]] = defaultdict(list)
    bridge = []
    for index, row in enumerate(rows, 1):
        pmid = row["pmid"]
        if not pmid.isdigit() or row["stratum"] not in FIELDS:
            raise ValueError("Invalid PMID or EPJ field")
        groups[pmid].append((index, row))
        bridge.append(
            {
                "source_record_id": f"epj-{index:05d}",
                "paper_id": f"PMID:{pmid}",
                "pmid": pmid,
                "epj_field": row["stratum"],
            }
        )
    registry = []
    for pmid, records in sorted(groups.items()):
        first = records[0][1]
        compared = {key: value for key, value in first.items() if key != "stratum"}
        for _, row in records[1:]:
            if {key: value for key, value in row.items() if key != "stratum"} != compared:
                raise ValueError(f"Conflicting repeated PMID: {pmid}")
        labels = sorted({row["stratum"] for _, row in records})
        registry.append(
            {
                "paper_id": f"PMID:{pmid}",
                "pmid": pmid,
                "doi": first["doi"],
                "epj_fields": ";".join(labels),
                "source_record_ids": ";".join(f"epj-{index:05d}" for index, _ in records),
                "record_multiplicity": str(len(records)),
                "pub_year": first["pub_year"],
                "epj_code_statement_detected": first["code_available"],
                "epj_data_statement_detected": first["data_available"],
                "epj_pmc_retrieval_flag": first["has_pmc_fulltext"],
                "epj_commercial_detected": first["has_commercial_software"],
                "epj_open_source_detected": first["has_opensource_software"],
            }
        )
    return registry, bridge


def wilson(successes: int, n: int, z: float = 1.959963984540054) -> tuple[float, float]:
    if n <= 0 or not 0 <= successes <= n:
        raise ValueError("Wilson interval requires 0 <= successes <= n and n > 0")
    p = successes / n
    denominator = 1 + z * z / n
    center = (p + z * z / (2 * n)) / denominator
    width = z * math.sqrt(p * (1 - p) / n + z * z / (4 * n * n)) / denominator
    lower = 0.0 if successes == 0 else max(0.0, center - width)
    upper = 1.0 if successes == n else min(1.0, center + width)
    return lower, upper


def fixed_triple(states: Sequence[State], threshold: int = 2) -> PaperState:
    if len(states) != 3 or threshold not in {1, 2, 3} or any(s not in STATES for s in states):
        raise ValueError("Exactly three valid slots and threshold 1, 2 or 3 are required")
    k = states.count("yes")
    m = sum(state in {"unknown", "not_assessable"} for state in states)
    if k >= threshold:
        return "success"
    if k + m < threshold:
        return "failure"
    return "unknown"


def validate_levels(
    implementation: State, execution: State, numerical: State, conclusion: State
) -> None:
    levels = (implementation, execution, numerical, conclusion)
    if any(level not in STATES for level in levels):
        raise ValueError("Invalid level")
    if execution == "yes" and implementation != "yes":
        raise ValueError("Execution requires an implementation")
    if numerical in {"yes", "no"} and execution != "yes":
        raise ValueError("Numerical comparison requires execution")
    if conclusion in {"yes", "no"} and execution != "yes":
        raise ValueError("Conclusion comparison requires execution")


def run_success(
    implementation: State,
    execution: State,
    numerical: State,
    conclusion: State,
    clean: bool,
) -> State:
    validate_levels(implementation, execution, numerical, conclusion)
    if not clean:
        return "not_assessable"
    necessary = (implementation, execution, numerical)
    if "no" in necessary:
        return "no"
    if all(value == "yes" for value in necessary):
        return "yes"
    return "unknown"


def rounded_agreement(observed: str, reported: str, decimal_places: int) -> bool:
    if decimal_places < 0:
        raise ValueError("Decimal places must be nonnegative")
    value, reference = Decimal(observed), Decimal(reported)
    if not value.is_finite() or not reference.is_finite():
        raise ValueError("Nonfinite target")
    half_unit = Decimal(10) ** -decimal_places / 2
    return reference - half_unit <= value < reference + half_unit


def utc_time(text: str) -> datetime:
    time = datetime.fromisoformat(text.replace("Z", "+00:00"))
    offset = time.utcoffset()
    if offset is None or offset.total_seconds() != 0:
        raise ValueError("Timestamp must include UTC timezone")
    return time


def validate_sequence(target: str, start: str, end: str, freeze: str, reveal: str) -> None:
    times = [utc_time(value) for value in (target, start, end, freeze, reveal)]
    if not (times[0] < times[1] < times[2] <= times[3] < times[4]):
        raise ValueError("Require target < start < end <= blind freeze < reveal")


@dataclass(frozen=True)
class Candidate:
    paper_id: str
    field: str


def assign_field(paper_id: str, fields: Sequence[str]) -> str:
    if not fields or set(fields) - set(FIELDS):
        raise ValueError("A nonempty valid recorded membership is required")
    return min(
        set(fields),
        key=lambda field: (sha256(f"paper2-field-v1|{paper_id}|{field}".encode()), field),
    )


def stratified_sample(candidates: Sequence[Candidate], total: int, seed: str) -> list[Row]:
    if len({item.paper_id for item in candidates}) != len(candidates):
        raise ValueError("Sampling requires a unique paper frame")
    if not candidates or total < 1 or total > len(candidates):
        raise ValueError("Invalid sample size")
    grouped: dict[str, list[Candidate]] = defaultdict(list)
    for item in candidates:
        if item.field not in FIELDS:
            raise ValueError("Unknown field")
        grouped[item.field].append(item)
    if total < len(grouped):
        raise ValueError("Sample too small to represent every nonempty stratum")
    quotas = dict.fromkeys(grouped, 1)
    for _ in range(total - len(grouped)):
        available = [field for field in grouped if quotas[field] < len(grouped[field])]
        chosen = max(
            available,
            key=lambda field: (
                total * len(grouped[field]) / len(candidates) - quotas[field],
                field,
            ),
        )
        quotas[chosen] += 1
    output = []
    for field in sorted(grouped):
        members = sorted(
            grouped[field],
            key=lambda item: (sha256(f"{seed}|{item.paper_id}".encode()), item.paper_id),
        )
        count = quotas[field]
        for item in members[:count]:
            output.append(
                {
                    "paper_id": item.paper_id,
                    "sampling_stratum": field,
                    "stratum_population": str(len(members)),
                    "stratum_sample": str(count),
                    "inclusion_probability": str(count / len(members)),
                    "weight": str(len(members) / count),
                    "seed": seed,
                }
            )
    return output


def gate_counts(states: Sequence[str]) -> Row:
    allowed = {"yes", "no", "uncertain", "unknown", "not_assessed"}
    if not states or set(states) - allowed:
        raise ValueError("Nonempty gate data with valid states required")
    counts = Counter(states)
    total = len(states)
    unresolved = sum(counts[key] for key in ("uncertain", "unknown", "not_assessed"))
    known = counts["yes"] + counts["no"]
    return {
        "n": str(total),
        "yes": str(counts["yes"]),
        "no": str(counts["no"]),
        "unresolved": str(unresolved),
        "yes_among_all": str(counts["yes"] / total),
        "yes_among_known": str(counts["yes"] / known) if known else "not_assessable",
        "identification_lower": str(counts["yes"] / total),
        "identification_upper": str((counts["yes"] + unresolved) / total),
    }
