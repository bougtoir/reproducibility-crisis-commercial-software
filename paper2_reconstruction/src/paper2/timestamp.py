import argparse
import http.client
import json
import subprocess
from datetime import datetime, timezone
from pathlib import Path

from paper2.acquisition import acquire
from paper2.core import sha256, snapshot

CA_SHA256 = "2151b61137ffa86bf664691ba67e7da0b19f98c758e3d228d5d8ebf27e044438"


def stamp(source: Path, destination: Path) -> dict[str, object]:
    destination.mkdir(parents=True, exist_ok=False)
    payload = source.read_bytes()
    retained_source = destination / "timestamped_payload"
    snapshot(retained_source, payload)
    ca_path, ca_receipt = acquire(
        "https://freetsa.org/files/cacert.pem",
        "FreeTSA_CA",
        destination / "authority",
        request_conditions="GET timestamp authority trust anchor; compare to published SHA-256",
    )
    if ca_receipt["http_status"] != 200 or ca_receipt["sha256"] != CA_SHA256:
        raise ValueError("Timestamp authority certificate does not match the reviewed trust anchor")
    query = subprocess.run(
        ["openssl", "ts", "-query", "-data", str(retained_source), "-sha256", "-cert"],
        capture_output=True,
        check=True,
        timeout=30,
    ).stdout
    snapshot(destination / "request.tsq", query)
    connection = http.client.HTTPSConnection("freetsa.org", timeout=60)
    try:
        connection.request(
            "POST",
            "/tsr",
            body=query,
            headers={"Content-Type": "application/timestamp-query"},
        )
        response = connection.getresponse()
        body = response.read()
        status = response.status
    finally:
        connection.close()
    response_path = destination / "response.tsr"
    snapshot(response_path, body)
    receipt = {
        "source_sha256": sha256(payload),
        "source_bytes": len(payload),
        "source_name": source.name,
        "url": "https://freetsa.org/tsr",
        "retrieved_at_utc": datetime.now(timezone.utc).isoformat(),
        "request_conditions": "POST RFC3161 SHA-256 digest and nonce; source bytes not transmitted",
        "request_sha256": sha256(query),
        "response_sha256": sha256(body),
        "response_bytes": len(body),
        "http_status": status,
        "trust_anchor_sha256": CA_SHA256,
    }
    snapshot(destination / "receipt.json", (json.dumps(receipt, indent=2) + "\n").encode())
    if status != 200:
        raise ValueError("Timestamp service did not return HTTP 200; no retry performed")
    verification = subprocess.run(
        [
            "openssl",
            "ts",
            "-verify",
            "-data",
            str(retained_source),
            "-in",
            str(response_path),
            "-CAfile",
            str(ca_path),
        ],
        capture_output=True,
        timeout=30,
    )
    snapshot(destination / "verification.stdout", verification.stdout)
    snapshot(destination / "verification.stderr", verification.stderr)
    verification.check_returncode()
    details = subprocess.run(
        ["openssl", "ts", "-reply", "-in", str(response_path), "-text"],
        capture_output=True,
        check=True,
        timeout=30,
    ).stdout
    snapshot(destination / "response_details.txt", details)
    result = {
        **receipt,
        "status": "cryptographically_verified_against_pinned_freetsa_ca",
        "scope": "file_existence_timestamp_only_not_study_readiness_or_preregistration",
        "source_path": str(source.resolve()),
        "receipt_path": str((destination / "receipt.json").resolve()),
    }
    snapshot(destination / "report.json", (json.dumps(result, indent=2) + "\n").encode())
    return result


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--source", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    print(json.dumps(stamp(args.source, args.output), indent=2))


if __name__ == "__main__":
    main()
