import json
import time
from dataclasses import asdict
from pathlib import Path
from unittest.mock import patch

import pytest

from paper2.controller import Limits, run
from paper2.core import sha256, snapshot
from paper2.model_api import Completion, bounded_complete


def slow_api(
    messages: list[dict[str, str]], archive: Path, max_tokens: int, timeout_seconds: int
) -> None:
    snapshot(archive / "started.json", b'{"synthetic_request": true}')
    time.sleep(30)


def completed_api(
    messages: list[dict[str, str]], archive: Path, max_tokens: int, timeout_seconds: int
) -> None:
    result = Completion("synthetic", "synthetic", "unknown", '{"value": 7}', "stop", 11, 5, 0.1)
    snapshot(archive / "completion.json", json.dumps(asdict(result)).encode())


def test_live_api_process_deadline_retains_unknown_usage_without_retry(tmp_path: Path) -> None:
    archive = tmp_path / "bounded"
    started = time.monotonic()
    with (
        patch("paper2.model_api.complete", slow_api),
        pytest.raises(RuntimeError, match="not retried"),
    ):
        bounded_complete([{"role": "user", "content": "synthetic"}], archive, 128, 1)
    assert time.monotonic() - started < 5
    boundary = json.loads((archive / "boundary.json").read_text())
    assert boundary["status"] == "deadline_exceeded"
    assert boundary["retries"] == 0
    assert boundary["provider_inflight_cancellation"] == "not_confirmed"
    assert (archive / "api/started.json").is_file()
    assert not (archive / "api/completion.json").exists()


def test_completed_api_process_retains_final_content_and_usage(tmp_path: Path) -> None:
    archive = tmp_path / "bounded"
    with patch("paper2.model_api.complete", completed_api):
        result = bounded_complete([{"role": "user", "content": "synthetic"}], archive, 128, 10)
    assert result.content == '{"value": 7}'
    assert (result.prompt_tokens, result.completion_tokens) == (11, 5)
    assert json.loads((archive / "boundary.json").read_text())["exitcode"] == 0


def test_interrupted_request_is_not_reported_as_zero_token_usage(tmp_path: Path) -> None:
    package = tmp_path / "package"
    snapshot(package / "input.txt", b"synthetic")
    with (
        patch("paper2.controller.Worker", autospec=True) as worker,
        patch("paper2.controller.bounded_complete", side_effect=RuntimeError("deadline")),
    ):
        worker.return_value.inspection = "[]"
        worker.return_value.export_artifacts.return_value = []
        result = run(
            package, {"input.txt": sha256(b"synthetic")}, "synthetic", tmp_path / "run", Limits()
        )
    assert result["requests"] == 1
    assert result["api_requests_with_usage"] == 0
    assert result["provider_accounted_tokens"] == "unknown"
    assert result["known_provider_accounted_tokens"] == 0
    assert result["usage_completeness"] == "partial"
