import hashlib
import json
import os
import socket
import subprocess
import sys
import time
import webbrowser
from pathlib import Path
from typing import Optional


APP_HOST = os.environ.get("ANSWER_SHEET_HOST", os.environ.get("HOST", "127.0.0.1"))
_port_str = os.environ.get("ANSWER_SHEET_PORT", os.environ.get("PORT", "8000"))
try:
    APP_PORT = int(_port_str)
except ValueError:
    APP_PORT = 8000
APP_URL = f"http://{APP_HOST}:{APP_PORT}"

REPO_DIR = Path(__file__).resolve().parent
VENV_DIR = REPO_DIR / ".venv"
REQ_PATH = REPO_DIR / "requirements.txt"
MARKER_PATH = REPO_DIR / ".answer_sheet_studio_install.json"
OUTPUTS_DIR = REPO_DIR / "outputs"
OUTPUTS_DIR.mkdir(exist_ok=True)
LAUNCHER_LOG = OUTPUTS_DIR / "launcher.log"
SERVER_LOG = OUTPUTS_DIR / "server.log"


def _log(line: str) -> None:
    ts = time.strftime("%Y-%m-%d %H:%M:%S")
    try:
        with open(LAUNCHER_LOG, "a", encoding="utf-8") as f:
            f.write(f"[{ts}] {line}\n")
    except Exception:
        pass


def is_port_open(host: str, port: int, timeout: float = 0.25) -> bool:
    try:
        with socket.create_connection((host, port), timeout=timeout):
            return True
    except OSError:
        return False


def _sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def _read_marker() -> Optional[dict]:
    if not MARKER_PATH.exists():
        return None
    try:
        return json.loads(MARKER_PATH.read_text(encoding="utf-8"))
    except Exception:
        return None


def _write_marker(data: dict) -> None:
    try:
        MARKER_PATH.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        pass


def _venv_python() -> Path:
    if os.name == "nt":
        return VENV_DIR / "Scripts" / "python.exe"
    py3 = VENV_DIR / "bin" / "python3"
    return py3 if py3.exists() else (VENV_DIR / "bin" / "python")


def _run_logged(args: list[str]) -> int:
    _log(f"$ {' '.join(args)}")
    with open(LAUNCHER_LOG, "a", encoding="utf-8") as log:
        p = subprocess.run(
            args,
            cwd=str(REPO_DIR),
            stdout=log,
            stderr=subprocess.STDOUT,
            text=True,
            errors="replace",
        )
        return int(p.returncode)


def ensure_venv() -> Path:
    py = _venv_python()
    if py.exists():
        return py
    _log("Creating .venv ...")
    rc = _run_logged([sys.executable, "-m", "venv", str(VENV_DIR)])
    if rc != 0 or not py.exists():
        raise SystemExit(f"Failed to create venv. See {LAUNCHER_LOG}")
    return py


def ensure_requirements(py: Path) -> None:
    if not REQ_PATH.exists():
        return
    req_sha = _sha256_file(REQ_PATH)
    marker = _read_marker() or {}
    if marker.get("requirements_sha256") == req_sha:
        return

    _log("Installing requirements ...")
    _run_logged([str(py), "-m", "pip", "install", "--upgrade", "pip"])
    rc = _run_logged([str(py), "-m", "pip", "install", "--only-binary=:all:", "-r", str(REQ_PATH)])
    if rc != 0:
        rc = _run_logged([str(py), "-m", "pip", "install", "-r", str(REQ_PATH)])
    if rc != 0:
        raise SystemExit(f"Failed to install requirements. See {LAUNCHER_LOG}")

    _write_marker({"requirements_sha256": req_sha, "installed_at": int(time.time())})


def start_server(py: Path) -> None:
    if is_port_open(APP_HOST, APP_PORT):
        _log(f"Server already running at {APP_URL}")
        webbrowser.open(APP_URL)
        return

    _log("Starting server ...")
    with open(SERVER_LOG, "a", encoding="utf-8") as server_log:
        kwargs = {
            "cwd": str(REPO_DIR),
            "stdout": server_log,
            "stderr": subprocess.STDOUT,
            "text": True,
            "errors": "replace",
        }
        if os.name == "nt":
            creationflags = 0
            creationflags |= getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0)
            creationflags |= getattr(subprocess, "DETACHED_PROCESS", 0)
            kwargs["creationflags"] = creationflags
        else:
            kwargs["start_new_session"] = True

        proc = subprocess.Popen([str(py), str(REPO_DIR / "run_app.py")], **kwargs)

        for _ in range(60):  # ~30 seconds
            if is_port_open(APP_HOST, APP_PORT):
                _log(f"Server running at {APP_URL}")
                webbrowser.open(APP_URL)
                return
            if proc.poll() is not None:
                raise SystemExit(f"Server exited (code {proc.returncode}). See {SERVER_LOG}")
            time.sleep(0.5)

    raise SystemExit(f"Server not reachable at {APP_URL}. See {SERVER_LOG}")


def main() -> None:
    if sys.version_info < (3, 10):
        raise SystemExit("Answer Sheet Studio requires Python 3.10+.")

    try:
        py = ensure_venv()
        ensure_requirements(py)
        start_server(py)
    except SystemExit:
        raise
    except Exception as exc:
        _log(f"ERROR: {exc!r}")
        raise


if __name__ == "__main__":
    main()
