import hashlib
import json
import os
import socket
import subprocess
import sys
import time
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from threading import Thread
from typing import Optional


APP_HOST = os.environ.get("ANSWER_SHEET_HOST", os.environ.get("HOST", "127.0.0.1"))
_port_str = os.environ.get("ANSWER_SHEET_PORT", os.environ.get("PORT", "8000"))
try:
    APP_PORT = int(_port_str)
except ValueError:
    APP_PORT = 8000
APP_URL = f"http://{APP_HOST}:{APP_PORT}"
OPEN_BROWSER = os.environ.get("ANSWER_SHEET_OPEN_BROWSER", "1").strip().lower() not in {"0", "false", "no"}
CLI_PROGRESS = os.environ.get("ANSWER_SHEET_PROGRESS", "").strip().lower() in {"1", "true", "yes", "cli", "terminal"}
WIZARD_MODE = os.environ.get("ANSWER_SHEET_WIZARD", "").strip().lower() in {"1", "true", "yes", "wizard"}

REPO_DIR = Path(__file__).resolve().parent
REQ_PATH = REPO_DIR / "requirements.txt"
OUTPUTS_DIR = REPO_DIR / "outputs"
OUTPUTS_DIR.mkdir(exist_ok=True)
LAUNCHER_LOG = OUTPUTS_DIR / "launcher.log"
SERVER_LOG = OUTPUTS_DIR / "server.log"
PROGRESS_URL_PATH = OUTPUTS_DIR / "progress_url.txt"

_PROGRESS_STATE: dict = {
    "phase": "starting",
    "message": "Starting…",
    "app_url": APP_URL,
    "app_up": False,
}


def _log(line: str) -> None:
    ts = time.strftime("%Y-%m-%d %H:%M:%S")
    if CLI_PROGRESS:
        try:
            print(f"[{ts}] {line}", flush=True)
        except Exception:
            pass
    try:
        with open(LAUNCHER_LOG, "a", encoding="utf-8") as f:
            f.write(f"[{ts}] {line}\n")
    except Exception:
        pass


def _set_progress(phase: str, message: str) -> None:
    _PROGRESS_STATE["phase"] = phase
    _PROGRESS_STATE["message"] = message
    _PROGRESS_STATE["ts"] = int(time.time())
    if CLI_PROGRESS:
        try:
            print(f"[{phase}] {message}", flush=True)
        except Exception:
            pass


def _tail_text(path: Path, max_lines: int = 80, max_bytes: int = 64 * 1024) -> str:
    try:
        if not path.exists():
            return ""
        size = path.stat().st_size
        with open(path, "rb") as f:
            f.seek(max(0, size - max_bytes))
            data = f.read()
        text = data.decode("utf-8", errors="replace")
        lines = text.splitlines()
        if len(lines) > max_lines:
            lines = lines[-max_lines:]
        return "\n".join(lines)
    except Exception:
        return ""


def _wizard_pause(message: str) -> None:
    if not WIZARD_MODE:
        return
    try:
        input(f"\n{message}\nPress Enter to continue...")
    except Exception:
        return


class _ProgressHandler(BaseHTTPRequestHandler):
    def log_message(self, fmt: str, *args) -> None:
        return

    def _send(self, status: int, content_type: str, body: bytes) -> None:
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self) -> None:  # noqa: N802
        if self.path.startswith("/status"):
            payload = dict(_PROGRESS_STATE)
            payload["log_tail"] = _tail_text(LAUNCHER_LOG)
            body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
            self._send(200, "application/json; charset=utf-8", body)
            return

        if self.path not in {"/", "/index.html"}:
            self._send(404, "text/plain; charset=utf-8", b"Not found")
            return

        html = f"""<!doctype html>
<html lang="en">
  <head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>Answer Sheet Studio - Installing</title>
    <style>
      :root {{ color-scheme: light dark; }}
      body {{ font-family: ui-sans-serif, system-ui, -apple-system, Segoe UI, Roboto, Helvetica, Arial; margin: 24px; }}
      .card {{ max-width: 900px; margin: 0 auto; padding: 20px; border-radius: 12px; border: 1px solid rgba(127,127,127,.25); }}
      .title {{ font-size: 20px; font-weight: 650; margin: 0 0 8px; }}
      .muted {{ opacity: .8; margin: 0 0 12px; }}
      .row {{ display: flex; gap: 12px; align-items: center; flex-wrap: wrap; }}
      .pill {{ font-size: 12px; padding: 4px 10px; border-radius: 999px; border: 1px solid rgba(127,127,127,.35); }}
      pre {{ white-space: pre-wrap; word-break: break-word; margin: 12px 0 0; padding: 12px; border-radius: 10px; border: 1px solid rgba(127,127,127,.25); max-height: 50vh; overflow: auto; }}
      progress {{ width: 100%; height: 16px; }}
    </style>
  </head>
  <body>
    <div class="card">
      <p class="title">Answer Sheet Studio</p>
      <p class="muted">Installing dependencies / starting server… This may take a few minutes on first run.</p>
      <div class="row">
        <span class="pill" id="phase">starting</span>
        <span class="pill" id="msg">Starting…</span>
      </div>
      <div style="margin-top:12px">
        <progress id="bar" max="100" value="10"></progress>
      </div>
      <pre id="log" aria-label="launcher log"></pre>
    </div>
    <script>
      (() => {{
        const phaseEl = document.getElementById("phase");
        const msgEl = document.getElementById("msg");
        const logEl = document.getElementById("log");
        const bar = document.getElementById("bar");
        const weights = {{
          starting: 5,
          venv: 15,
          pip_upgrade: 30,
          pip_install: 55,
          server: 80,
          done: 100,
          error: 100
        }};
        const poll = async () => {{
          try {{
            const res = await fetch("/status", {{ cache: "no-store" }});
            const j = await res.json();
            phaseEl.textContent = j.phase || "starting";
            msgEl.textContent = j.message || "";
            bar.value = weights[j.phase] ?? 10;
            if (j.log_tail) {{
              logEl.textContent = j.log_tail;
              logEl.scrollTop = logEl.scrollHeight;
            }}
            if (j.app_up && j.app_url) {{
              window.location.href = j.app_url;
            }}
          }} catch (e) {{
            // ignore transient errors
          }}
        }};
        poll();
        setInterval(poll, 1000);
      }})();
    </script>
  </body>
</html>
"""
        self._send(200, "text/html; charset=utf-8", html.encode("utf-8"))


class _ProgressServer:
    def __init__(self) -> None:
        self._httpd: Optional[ThreadingHTTPServer] = None
        self.url: Optional[str] = None

    def start(self) -> None:
        httpd = ThreadingHTTPServer(("127.0.0.1", 0), _ProgressHandler)
        self._httpd = httpd
        host, port = httpd.server_address[:2]
        self.url = f"http://{host}:{port}/"
        try:
            PROGRESS_URL_PATH.write_text(self.url, encoding="utf-8")
        except Exception:
            pass
        Thread(target=httpd.serve_forever, daemon=True).start()

    def stop(self) -> None:
        if not self._httpd:
            return
        try:
            self._httpd.shutdown()
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


def _venv_root_dir() -> Path:
    explicit = os.environ.get("ANSWER_SHEET_VENV_DIR", "").strip()
    if explicit:
        return Path(explicit).expanduser()

    # Windows: reuse a stable venv across re-downloads/unzips of the repo.
    if os.name == "nt":
        base = os.environ.get("LOCALAPPDATA") or str(Path.home())
        return Path(base) / "AnswerSheetStudio" / "venvs"

    return REPO_DIR / ".venv"


def _venv_dir() -> Path:
    root = _venv_root_dir()
    if root == REPO_DIR / ".venv":
        return root
    try:
        req_sha = _sha256_file(REQ_PATH) if REQ_PATH.exists() else "default"
    except Exception:
        req_sha = "default"
    return root / req_sha[:12]


def _marker_path() -> Path:
    venv_dir = _venv_dir()
    if venv_dir == REPO_DIR / ".venv":
        return REPO_DIR / ".answer_sheet_studio_install.json"
    return venv_dir / ".answer_sheet_studio_install.json"


def _read_marker() -> Optional[dict]:
    marker_path = _marker_path()
    if not marker_path.exists():
        return None
    try:
        return json.loads(marker_path.read_text(encoding="utf-8"))
    except Exception:
        return None


def _write_marker(data: dict) -> None:
    try:
        marker_path = _marker_path()
        marker_path.parent.mkdir(parents=True, exist_ok=True)
        marker_path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        pass


def _venv_python() -> Path:
    venv_dir = _venv_dir()
    if os.name == "nt":
        pyw = venv_dir / "Scripts" / "pythonw.exe"
        return pyw if pyw.exists() else (venv_dir / "Scripts" / "python.exe")
    py3 = venv_dir / "bin" / "python3"
    return py3 if py3.exists() else (venv_dir / "bin" / "python")


def _run_logged(args: list[str]) -> int:
    _log(f"$ {' '.join(args)}")
    with open(LAUNCHER_LOG, "a", encoding="utf-8") as log:
        kwargs = {}
        if os.name == "nt":
            kwargs["creationflags"] = getattr(subprocess, "CREATE_NO_WINDOW", 0)
        p = subprocess.run(
            args,
            cwd=str(REPO_DIR),
            stdout=log,
            stderr=subprocess.STDOUT,
            text=True,
            errors="replace",
            **kwargs,
        )
        return int(p.returncode)


def ensure_venv() -> Path:
    py = _venv_python()
    if py.exists():
        _set_progress("venv", "Virtual environment already exists.")
        _wizard_pause("Virtual environment is ready.")
        return py
    _set_progress("venv", "Creating virtual environment…")
    _log("Creating .venv ...")
    venv_dir = _venv_dir()
    try:
        venv_dir.parent.mkdir(parents=True, exist_ok=True)
    except Exception:
        pass
    rc = _run_logged([sys.executable, "-m", "venv", str(venv_dir)])
    if rc != 0 or not py.exists():
        _set_progress("error", "Failed to create venv (see launcher.log).")
        raise SystemExit(f"Failed to create venv. See {LAUNCHER_LOG}")
    _wizard_pause("Virtual environment created.")
    return py


def ensure_requirements(py: Path) -> None:
    if not REQ_PATH.exists():
        return
    req_sha = _sha256_file(REQ_PATH)
    marker = _read_marker() or {}
    if marker.get("requirements_sha256") == req_sha:
        _set_progress("pip_install", "Python packages already installed.")
        _wizard_pause("Python packages are already installed.")
        return

    _set_progress("pip_upgrade", "Upgrading pip…")
    _log("Installing requirements ...")
    _run_logged([str(py), "-m", "pip", "install", "--upgrade", "pip"])
    _set_progress("pip_install", "Installing Python packages…")
    rc = _run_logged([str(py), "-m", "pip", "install", "--only-binary=:all:", "-r", str(REQ_PATH)])
    if rc != 0:
        rc = _run_logged([str(py), "-m", "pip", "install", "-r", str(REQ_PATH)])
    if rc != 0:
        _set_progress("error", "Failed to install requirements (see launcher.log).")
        raise SystemExit(f"Failed to install requirements. See {LAUNCHER_LOG}")

    _write_marker({"requirements_sha256": req_sha, "installed_at": int(time.time())})
    _wizard_pause("Python packages installed.")


def start_server(py: Path) -> None:
    if is_port_open(APP_HOST, APP_PORT):
        _log(f"Server already running at {APP_URL}")
        _PROGRESS_STATE["app_up"] = True
        _set_progress("done", "Server already running.")
        _wizard_pause(f"Server is already running at {APP_URL}")
        if OPEN_BROWSER:
            webbrowser.open(APP_URL)
        return

    _set_progress("server", "Starting server…")
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
            creationflags |= getattr(subprocess, "CREATE_NO_WINDOW", 0)
            kwargs["creationflags"] = creationflags
        else:
            kwargs["start_new_session"] = True

        proc = subprocess.Popen([str(py), str(REPO_DIR / "run_app.py")], **kwargs)

        for _ in range(60):  # ~30 seconds
            if is_port_open(APP_HOST, APP_PORT):
                _log(f"Server running at {APP_URL}")
                _PROGRESS_STATE["app_up"] = True
                _set_progress("done", "Server is ready.")
                _wizard_pause(f"Server is ready at {APP_URL}")
                if OPEN_BROWSER:
                    webbrowser.open(APP_URL)
                return
            if proc.poll() is not None:
                _set_progress("error", f"Server exited (code {proc.returncode}). See server.log.")
                raise SystemExit(f"Server exited (code {proc.returncode}). See {SERVER_LOG}")
            time.sleep(0.5)

    _set_progress("error", "Server not reachable (see server.log).")
    raise SystemExit(f"Server not reachable at {APP_URL}. See {SERVER_LOG}")


def main() -> None:
    progress = _ProgressServer()
    try:
        if not CLI_PROGRESS:
            progress.start()
        _set_progress("starting", "Preparing installation…")
        if (not CLI_PROGRESS) and OPEN_BROWSER and progress.url:
            webbrowser.open(progress.url)
    except Exception:
        pass

    try:
        if sys.version_info < (3, 10):
            _log("ERROR: Python 3.10+ is required to run the app.")
            _set_progress("error", "Python 3.10+ is required. Please install Python 3.11 (recommended).")
            if not CLI_PROGRESS:
                time.sleep(8)
            raise SystemExit("Answer Sheet Studio requires Python 3.10+.")
        py = ensure_venv()
        ensure_requirements(py)
        start_server(py)
    except SystemExit:
        raise
    except Exception as exc:
        _log(f"ERROR: {exc!r}")
        _set_progress("error", f"Unexpected error: {exc!r}")
        raise
    finally:
        try:
            if not CLI_PROGRESS:
                progress.stop()
        except Exception:
            pass


if __name__ == "__main__":
    main()
