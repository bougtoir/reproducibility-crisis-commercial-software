import json
from pathlib import Path
from unittest.mock import patch

import pytest

from paper2.core import sha256
from paper2.model_api import complete


@pytest.mark.parametrize("status", [200, 302, 503])
def test_api_retains_only_final_content_and_never_retries_or_redirects(
    tmp_path: Path, status: int
) -> None:
    class Response:
        def __init__(self) -> None:
            self.status = status

        def read(self, limit: int) -> bytes:
            return json.dumps(
                {
                    "id": "synthetic-response",
                    "model": "synthetic-model",
                    "choices": [
                        {
                            "finish_reason": "stop",
                            "message": {
                                "content": '{"answer": 46}',
                                "reasoning_content": "SYNTHETIC_PRIVATE_REASONING",
                            },
                        }
                    ],
                    "usage": {"prompt_tokens": 12, "completion_tokens": 6},
                }
            ).encode()

    class Connection:
        calls = 0
        closed = False

        def __init__(self, host: str, timeout: int) -> None:
            assert host == "api.deepseek.com"
            assert timeout == 120

        def request(self, method: str, path: str, body: bytes, headers: dict[str, str]) -> None:
            Connection.calls += 1
            assert method == "POST"
            assert path == "/chat/completions"
            assert headers["Authorization"] == "Bearer SYNTHETIC_SECRET"
            request = json.loads(body)
            assert request["thinking"] == {"type": "disabled"}
            assert "tools" not in request

        def getresponse(self) -> Response:
            return Response()

        def close(self) -> None:
            Connection.closed = True

    archive = tmp_path / "call"
    with (
        patch("paper2.model_api.http.client.HTTPSConnection", Connection),
        patch.dict("os.environ", {"DEEPSEEK_API_KEY": "SYNTHETIC_SECRET"}),
    ):
        if status == 200:
            result = complete([{"role": "user", "content": "Synthetic JSON test"}], archive, 128)
            assert result.content == '{"answer": 46}'
            assert result.model_version == "not_reported_by_provider"
            assert result.prompt_tokens == 12
        else:
            with pytest.raises(RuntimeError, match=f"HTTP {status}"):
                complete([{"role": "user", "content": "Synthetic JSON test"}], archive, 128)
    assert Connection.calls == 1
    assert Connection.closed
    receipt = json.loads((archive / "receipt.json").read_text())
    assert receipt["request_sha256"] == sha256((archive / "request.json").read_bytes())
    for path in archive.iterdir():
        assert "SYNTHETIC_PRIVATE_REASONING" not in path.read_text()
        assert "SYNTHETIC_SECRET" not in path.read_text()
