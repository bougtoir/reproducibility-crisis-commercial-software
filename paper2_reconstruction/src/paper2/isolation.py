import argparse
import io
import json
import os
import selectors
import subprocess
import tarfile
import time
import uuid
from dataclasses import asdict, dataclass
from datetime import datetime, timezone
from pathlib import Path, PurePosixPath

from paper2.core import sha256, sha256_file, snapshot

IMAGE = "python@sha256:65a93d69fa75478d554f4ad27c85c1e69fa184956261b4301ebaf6dbb0a3543d"
DOCKER = "/usr/bin/docker"
MAX_ARTIFACT_BYTES = 64 * 1024 * 1024


@dataclass(frozen=True)
class Execution:
    code_sha256: str
    returncode: int
    output: str
    stop_reason: str
    elapsed_seconds: float


def docker(*arguments: str) -> str:
    return subprocess.check_output(
        [DOCKER, *arguments], text=True, timeout=30, env={"PATH": "/usr/bin:/bin"}
    ).strip()


def regular_tar_members(raw: bytes) -> dict[str, bytes]:
    if len(raw) > MAX_ARTIFACT_BYTES:
        raise ValueError("Worker artifact archive exceeded the fixed export limit")
    files: dict[str, bytes] = {}
    with tarfile.open(fileobj=io.BytesIO(raw), mode="r|") as archive:
        for index, member in enumerate(archive):
            if index >= 1000:
                raise ValueError("Worker artifact archive has excessive members")
            parts = tuple(part for part in PurePosixPath(member.name).parts if part != ".")
            if not parts:
                continue
            unsafe_path = (
                ".." in parts or PurePosixPath(member.name).is_absolute() or len(parts) > 12
            )
            if member.isdir() and not unsafe_path:
                continue
            if not member.isfile() or member.size > MAX_ARTIFACT_BYTES or unsafe_path:
                raise ValueError("Worker artifacts contain an unsafe or non-regular member")
            name = str(PurePosixPath(*parts))
            if name in files or len(files) >= 1000:
                raise ValueError("Worker artifact archive has duplicate or excessive members")
            payload = archive.extractfile(member)
            if payload is None:
                raise ValueError("Worker artifact could not be read")
            files[name] = payload.read()
    return files


def bounded_capture(
    arguments: list[str], seconds: float, byte_limit: int
) -> tuple[bytes, bytes, int | None, str]:
    if seconds <= 0 or byte_limit <= 0:
        raise ValueError("Capture limits must be positive")
    process = subprocess.Popen(
        arguments, stdout=subprocess.PIPE, stderr=subprocess.PIPE, env={"PATH": "/usr/bin:/bin"}
    )
    streams: dict[str, bytearray] = {"stdout": bytearray(), "stderr": bytearray()}
    total = 0
    reason = "completed"
    started = time.monotonic()
    try:
        with selectors.DefaultSelector() as selector:
            assert process.stdout is not None and process.stderr is not None
            selector.register(process.stdout, selectors.EVENT_READ, "stdout")
            selector.register(process.stderr, selectors.EVENT_READ, "stderr")
            while selector.get_map():
                remaining = seconds - (time.monotonic() - started)
                if remaining <= 0:
                    reason = "timeout"
                    break
                for key, _ in selector.select(min(0.1, remaining)):
                    chunk = os.read(key.fd, 65536)
                    if not chunk:
                        selector.unregister(key.fileobj)
                        continue
                    retained = chunk[: byte_limit - total]
                    streams[key.data].extend(retained)
                    total += len(retained)
                    if len(retained) != len(chunk):
                        reason = "output_limit"
                        break
                if reason != "completed":
                    break
            if reason == "completed":
                try:
                    process.wait(timeout=max(0.001, seconds - (time.monotonic() - started)))
                except subprocess.TimeoutExpired:
                    reason = "timeout"
    finally:
        if process.poll() is None:
            process.kill()
        process.wait()
        if process.stdout is not None:
            process.stdout.close()
        if process.stderr is not None:
            process.stderr.close()
    return bytes(streams["stdout"]), bytes(streams["stderr"]), process.returncode, reason


class Worker:
    def __init__(
        self,
        package: Path,
        expected: dict[str, str],
        image: str = IMAGE,
        *,
        memory: str = "2g",
        work: Path | None = None,
    ) -> None:
        """Start an isolated worker over an exact package.

        ``work`` selects a disk-backed writable /work directory instead of the default
        512 MiB tmpfs, for runs whose intermediate files exceed memory-backed storage.
        """
        actual: dict[str, str] = {}
        for path in package.rglob("*"):
            if path.is_symlink():
                raise ValueError("Input packages must not contain symlinks")
            if path.is_file():
                actual[str(path.relative_to(package))] = sha256_file(path)
        if not expected or actual != expected or package.is_symlink():
            raise ValueError("Input package differs from exact approved manifest")
        self.name = f"paper2-qualification-{uuid.uuid4().hex}"
        self.closed = False
        if work is None:
            work_mount = ("--tmpfs", "/work:rw,nosuid,nodev,size=512m,mode=1777")
        else:
            work.mkdir(parents=True, exist_ok=True)
            work.chmod(0o1777)
            work_mount = ("--mount", f"type=bind,source={work.resolve()},target=/work")
        docker(
            "run",
            "--detach",
            "--name",
            self.name,
            "--network",
            "none",
            "--read-only",
            "--user",
            "65534:65534",
            "--cap-drop",
            "ALL",
            "--security-opt",
            "no-new-privileges",
            "--pids-limit",
            "64",
            "--memory",
            memory,
            "--memory-swap",
            memory,
            "--cpus",
            "1",
            "--workdir",
            "/work",
            *work_mount,
            "--tmpfs",
            "/tmp:rw,nosuid,nodev,size=64m,mode=1777",
            "--mount",
            f"type=bind,source={package.resolve()},target=/input,readonly",
            image,
            "sleep",
            "infinity",
        )
        self.inspection = docker("inspect", self.name)

    def execute(self, code: str, seconds: float = 30, max_output: int = 65536) -> Execution:
        if self.closed:
            raise RuntimeError("A sealed worker cannot execute additional code")
        start = time.monotonic()
        output = bytearray()
        reason = "completed"
        with (
            subprocess.Popen(
                [DOCKER, "exec", self.name, "python", "-I", "-u", "-c", code],
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                env={"PATH": "/usr/bin:/bin"},
            ) as process,
            selectors.DefaultSelector() as selector,
        ):
            assert process.stdout is not None
            selector.register(process.stdout, selectors.EVENT_READ)
            while selector.get_map():
                remaining = seconds - (time.monotonic() - start)
                if remaining <= 0:
                    reason = "wall_limit"
                    break
                for key, _ in selector.select(min(remaining, 0.2)):
                    chunk = os.read(key.fd, min(8192, max_output + 1 - len(output)))
                    if not chunk:
                        selector.unregister(key.fileobj)
                        continue
                    output.extend(chunk)
                if len(output) > max_output:
                    reason = "output_limit"
                    break
            if reason != "completed":
                process.kill()
            if reason == "wall_limit":
                docker("kill", self.name)
                self.closed = True
            returncode = process.wait(timeout=10)
        return Execution(
            sha256(code.encode()),
            returncode,
            bytes(output[:max_output]).decode(errors="replace"),
            reason,
            time.monotonic() - start,
        )

    def export_artifacts(self, destination: Path) -> list[dict[str, object]]:
        exporter = (
            "import sys,tarfile;"
            "a=tarfile.open(fileobj=sys.stdout.buffer,mode='w|');"
            "a.add('/work',arcname='.',recursive=True,"
            "filter=lambda i: None if i.name.split('/')[1:2]==['scratch'] else i);a.close()"
        )
        output, errors, status, reason = bounded_capture(
            [DOCKER, "exec", self.name, "python", "-c", exporter],
            seconds=30,
            byte_limit=MAX_ARTIFACT_BYTES,
        )
        snapshot(destination / "artifacts.tar", output)
        snapshot(destination / "export.stderr", errors)
        if reason != "completed" or status != 0:
            raise ValueError(f"Worker artifact export stopped: {reason}; status={status}")
        records = []
        for name, payload in regular_tar_members(output).items():
            snapshot(destination / "files" / name, payload)
            records.append({"path": name, "bytes": len(payload), "sha256": sha256(payload)})
        return records

    def close(self) -> None:
        docker("rm", "--force", self.name)
        self.closed = True


PROBES = """
import json
import os
import socket
from pathlib import Path

checks = {"nonroot": os.getuid() == 65534}
checks["input_readable"] = Path("/input/allowed.txt").read_text() == "SYNTHETIC_ALLOWED"
checks["no_host_canary"] = not Path(HOST_CANARY_PATH).exists()
checks["no_docker_socket"] = not Path("/var/run/docker.sock").exists()
checks["no_api_credential"] = "DEEPSEEK_API_KEY" not in os.environ
checks["no_sibling_state"] = not Path("/work/sibling.txt").exists()
for name, path in (
    ("input_readonly", "/input/write-test"),
    ("root_readonly", "/root-write-test"),
):
    try:
        Path(path).write_text("synthetic")
        checks[name] = False
    except OSError:
        checks[name] = True
with socket.socket() as connection:
    connection.settimeout(2)
    try:
        connection.connect(("1.1.1.1", 443))
        checks["external_network_denied"] = False
    except OSError:
        checks["external_network_denied"] = True
status = Path("/proc/self/status").read_text()
checks["capabilities_empty"] = "CapEff:\\t0000000000000000" in status
checks["no_new_privileges"] = "NoNewPrivs:\\t1" in status
checks["seccomp_filter"] = "Seccomp:\\t2" in status
checks["memory_limit"] = Path("/sys/fs/cgroup/memory.max").read_text().strip() == "2147483648"
checks["pids_limit"] = Path("/sys/fs/cgroup/pids.max").read_text().strip() == "64"
Path("/work/sibling.txt").write_text("SYNTHETIC_SIBLING")
print(json.dumps(checks, sort_keys=True))
"""


def qualification(destination: Path) -> dict[str, object]:
    destination.mkdir(parents=True, exist_ok=False)
    package = destination / "package"
    package.mkdir()
    snapshot(package / "allowed.txt", b"SYNTHETIC_ALLOWED")
    snapshot(destination / "host-canary.txt", b"SYNTHETIC_FORBIDDEN")
    probes = PROBES.replace(
        "HOST_CANARY_PATH", json.dumps(str((destination / "host-canary.txt").resolve()))
    )
    snapshot(destination / "probes.py", probes.encode())
    expected = {"allowed.txt": sha256(b"SYNTHETIC_ALLOWED")}
    executions = []
    all_passed = True
    for index in range(2):
        worker = Worker(package, expected)
        try:
            snapshot(destination / f"worker-{index}.json", worker.inspection.encode())
            result = worker.execute(probes)
            snapshot(
                destination / f"probe-{index}.json", json.dumps(asdict(result), indent=2).encode()
            )
            checks: dict[str, bool] = json.loads(result.output)
            passed = (
                result.returncode == 0
                and result.stop_reason == "completed"
                and len(checks) == 14
                and all(checks.values())
            )
            all_passed = all_passed and passed
            executions.append(checks)
            limit = worker.execute("while True: pass", seconds=1)
            all_passed = all_passed and limit.stop_reason == "wall_limit"
            snapshot(
                destination / f"timeout-{index}.json", json.dumps(asdict(limit), indent=2).encode()
            )
        finally:
            worker.close()
    report: dict[str, object] = {
        "status": "execution_controls_pass" if all_passed else "execution_controls_failed",
        "scope": "synthetic_software_qualification_only",
        "main_study_authorized": False,
        "pilot_authorized": False,
        "image": IMAGE,
        "harness_sha256": sha256(Path(__file__).read_bytes()),
        "recorded_at_utc": datetime.now(timezone.utc).isoformat(),
        "checks": executions,
        "unqualified": [
            "model_API_controller_and_effective_context",
            "package_custodian_review_and_broker",
            "dependency_closure_and_scientific_packages",
            "telemetry_completeness_and_budget_comparability",
            "independent_adjudication_and_external_freeze",
        ],
    }
    snapshot(destination / "report.json", (json.dumps(report, indent=2) + "\n").encode())
    return report


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    report = qualification(args.output)
    print(json.dumps(report, indent=2))
    if report["status"] != "execution_controls_pass":
        raise SystemExit(1)


if __name__ == "__main__":
    main()
