import argparse
import http.client
import json
import os
import time
from dataclasses import asdict, dataclass
from datetime import datetime, timezone
from pathlib import Path

from paper2.core import sha256, snapshot

MODEL = "deepseek-flash"


def mapping(value: object) -> dict[str, object]:
    if not isinstance(value, dict) or any(not isinstance(key, str) for key in value):
        raise ValueError("Expected a JSON object with string keys")
    return dict(value)


@dataclass(frozen=True)
class Completion:
    response_id: str
    reported_model: str
    model_version: str
    content: str
    finish_reason: str
    prompt_tokens: int
    completion_tokens: int
    elapsed_seconds: float


def complete(
    messages: list[dict[str, str]], archive: Path, max_tokens: int, timeout_seconds: int = 120
) -> Completion:
    archive.mkdir(parents=True, exist_ok=False)
    request = {
        "model": MODEL,
        "messages": messages,
        "thinking": {"type": "disabled"},
        "response_format": {"type": "json_object"},
        "temperature": 0,
        "max_tokens": max_tokens,
        "stream": False,
    }
    payload = json.dumps(request, ensure_ascii=False).encode()
    snapshot(archive / "request.json", payload)
    start = time.monotonic()
    if timeout_seconds < 1 or timeout_seconds > 120:
        raise ValueError("API timeout must be between one and 120 seconds")
    connection = http.client.HTTPSConnection("api.deepseek.com", timeout=timeout_seconds)
    try:
        connection.request(
            "POST",
            "/chat/completions",
            body=payload,
            headers={
                "Authorization": f"Bearer {os.environ['DEEPSEEK_API_KEY']}",
                "Content-Type": "application/json",
            },
        )
        response = connection.getresponse()
        body = response.read(8 * 1024 * 1024 + 1)
        receipt = {
            "url": "https://api.deepseek.com/chat/completions",
            "http_status": response.status,
            "recorded_at_utc": datetime.now(timezone.utc).isoformat(),
            "request_sha256": sha256(payload),
            "response_storage": "final_content_and_usage_only; private_reasoning_not_retained",
        }
        snapshot(archive / "receipt.json", json.dumps(receipt, indent=2).encode())
        if response.status != 200:
            raise RuntimeError(f"Model API returned HTTP {response.status}; request not retried")
        if len(body) > 8 * 1024 * 1024:
            raise ValueError("Model response exceeded storage limit")
        data = mapping(json.loads(body))
    finally:
        connection.close()
    choices = data["choices"]
    if not isinstance(choices, list) or len(choices) != 1:
        raise ValueError("Expected exactly one completion")
    choice = mapping(choices[0])
    message = mapping(choice["message"])
    usage = mapping(data["usage"])
    content, finish = message["content"], choice["finish_reason"]
    model, response_id = data["model"], data["id"]
    prompt_tokens, completion_tokens = usage["prompt_tokens"], usage["completion_tokens"]
    if not all(isinstance(value, str) for value in (content, finish, model, response_id)):
        raise ValueError("Missing provider identity, finish reason or final content")
    if not isinstance(prompt_tokens, int) or not isinstance(completion_tokens, int):
        raise ValueError("Provider did not return accounted token usage")
    result = Completion(
        str(response_id),
        str(model),
        "not_reported_by_provider",
        str(content),
        str(finish),
        prompt_tokens,
        completion_tokens,
        time.monotonic() - start,
    )
    snapshot(archive / "completion.json", json.dumps(asdict(result), indent=2).encode())
    return result


def qualify(destination: Path) -> dict[str, object]:
    destination.mkdir(parents=True, exist_ok=False)
    messages = [
        {
            "role": "system",
            "content": "Synthetic qualification only. Return JSON. No tools are available.",
        },
        {
            "role": "user",
            "content": "Return a JSON object with answer equal to 17 + 29, and scope='synthetic'.",
        },
    ]
    results = [complete(messages, destination / f"request-{index}", 128) for index in range(2)]
    passed = all(
        mapping(json.loads(result.content)) == {"answer": 46, "scope": "synthetic"}
        and result.finish_reason == "stop"
        and result.reported_model == MODEL
        for result in results
    )
    report: dict[str, object] = {
        "status": "API_smoke_pass" if passed else "API_smoke_failed",
        "scope": "synthetic_software_qualification_only",
        "model": MODEL,
        "controller_sha256": sha256(Path(__file__).read_bytes()),
        "registered_provider_tools": [],
        "private_reasoning_stored": False,
        "pilot_authorized": False,
        "main_study_authorized": False,
        "limits": [
            "No empirical solver controller or adjudication has been qualified",
            "Provider hidden context and training exposure are not observable",
            "Provider model version is not reported; requested model is an alias",
        ],
    }
    snapshot(destination / "report.json", (json.dumps(report, indent=2) + "\n").encode())
    return report


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    report = qualify(args.output)
    print(json.dumps(report, indent=2))
    if report["status"] != "API_smoke_pass":
        raise SystemExit(1)


if __name__ == "__main__":
    main()
