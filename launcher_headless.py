import hashlib
import json
import locale
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

SUPPORTED_LANGS = ("zh-Hant", "en")
DEFAULT_LANG = "zh-Hant"


def _normalize_lang(value: str) -> Optional[str]:
    if not value:
        return None
    s = value.strip().lower().replace("_", "-")
    if not s:
        return None
    if s.startswith("zh"):
        return "zh-Hant"
    if s.startswith("en"):
        return "en"
    return None


def _detect_launcher_lang() -> str:
    env_lang = _normalize_lang(os.environ.get("ANSWER_SHEET_LANG", ""))
    if env_lang:
        return env_lang
    for key in ("LC_ALL", "LC_CTYPE", "LANG", "LANGUAGE"):
        value = os.environ.get(key, "")
        lang = _normalize_lang(value)
        if lang:
            return lang
    try:
        loc = locale.getdefaultlocale()[0] if locale.getdefaultlocale() else ""
    except Exception:
        loc = ""
    lang = _normalize_lang(loc or "")
    return lang or DEFAULT_LANG


LAUNCHER_LANG = _detect_launcher_lang()

LAUNCHER_I18N = {
    "zh-Hant": {
        "html_lang": "zh-Hant",
        "title": "Answer Sheet Studio - 安裝中",
        "brand": "Answer Sheet Studio",
        "subtitle": "正在安裝相依套件 / 啟動伺服器…首次執行可能需要幾分鐘。",
        "open_app": "開啟 Answer Sheet Studio",
        "close_tab": "可以關閉此分頁。",
        "not_found": "找不到頁面",
        "phase_map": {
            "starting": "啟動中",
            "venv": "建立環境",
            "pip_upgrade": "更新工具",
            "pip_install": "安裝套件",
            "server": "啟動伺服器",
            "done": "完成",
            "error": "錯誤",
        },
        "msg_map": {
            "Starting…": "正在啟動...",
            "Preparing installation…": "準備安裝環境...",
            "Virtual environment already exists.": "虛擬環境已就緒",
            "Virtual environment is ready.": "虛擬環境準備完成",
            "Recreating virtual environment (old version).": "正在更新虛擬環境版本...",
            "Creating virtual environment…": "正在建立虛擬環境...",
            "Failed to create venv (see launcher.log).": "建立環境失敗 (請看 log)",
            "Virtual environment created.": "虛擬環境建立完成",
            "Python packages already installed.": "Python 套件已安裝",
            "Python packages are already installed.": "套件檢查完成",
            "Upgrading pip…": "正在更新安裝工具 (pip)...",
            "Installing requirements ...": "正在安裝相依套件...",
            "Installing Python packages…": "正在下載並安裝 Python 套件 (OpenCV, FastAPI)... 這可能需要幾分鐘",
            "Failed to install requirements (see launcher.log).": "套件安裝失敗 (請看 log)",
            "Python packages installed.": "套件安裝完成",
            "Server already running.": "伺服器已在執行中",
            "Starting server…": "正在啟動伺服器...",
            "Server is ready.": "伺服器已就緒！",
            "Server not reachable (see server.log).": "伺服器連線失敗",
            "Unexpected error": "發生未預期的錯誤",
            "Python 3.10+ is required": "需要 Python 3.10 以上版本",
        },
    },
    "en": {
        "html_lang": "en",
        "title": "Answer Sheet Studio - Installing",
        "brand": "Answer Sheet Studio",
        "subtitle": "Installing dependencies / starting server… This may take a few minutes on first run.",
        "open_app": "Open Answer Sheet Studio",
        "close_tab": "You can close this tab.",
        "not_found": "Not found",
        "phase_map": {
            "starting": "Starting",
            "venv": "Preparing venv",
            "pip_upgrade": "Upgrading pip",
            "pip_install": "Installing packages",
            "server": "Starting server",
            "done": "Done",
            "error": "Error",
        },
        "msg_map": {},
    },
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
        if os.name == "nt" and "�" in text:
            try:
                mbcs_text = data.decode("mbcs", errors="replace")
                if mbcs_text.count("�") < text.count("�"):
                    text = mbcs_text
            except Exception:
                pass
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
        i18n = LAUNCHER_I18N.get(LAUNCHER_LANG, LAUNCHER_I18N[DEFAULT_LANG])
        if self.path.startswith("/status"):
            payload = dict(_PROGRESS_STATE)
            payload["log_tail"] = _tail_text(LAUNCHER_LOG)
            body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
            self._send(200, "application/json; charset=utf-8", body)
            return

        if self.path not in {"/", "/index.html"}:
            self._send(404, "text/plain; charset=utf-8", i18n["not_found"].encode("utf-8"))
            return

        msg_map = json.dumps(i18n["msg_map"], ensure_ascii=False)
        phase_map = json.dumps(i18n["phase_map"], ensure_ascii=False)
        html = f"""<!doctype html>
<html lang="{i18n['html_lang']}">
  <head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>{i18n['title']}</title>
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
      <p class="title">{i18n['brand']}</p>
      <p class="muted">{i18n['subtitle']}</p>
      <div class="row">
        <span class="pill" id="phase">{i18n['phase_map'].get('starting', 'starting')}</span>
        <span class="pill" id="msg">{i18n['msg_map'].get('Starting…', 'Starting...')}</span>
      </div>
      <div style="margin-top:12px">
        <progress id="bar" max="100" value="10"></progress>
      </div>
      <div id="done" style="display:none; margin-top:12px">
        <a id="open" href="{APP_URL}" target="_blank" rel="noopener">{i18n['open_app']}</a>
        <span style="opacity:.7">{i18n['close_tab']}</span>
      </div>
      <pre id="log" aria-label="launcher log" style="display:none"></pre>
    </div>
    <script>
      (() => {{
        const phaseEl = document.getElementById("phase");
        const msgEl = document.getElementById("msg");
        const logEl = document.getElementById("log");
        const bar = document.getElementById("bar");
        const doneEl = document.getElementById("done");
        const openEl = document.getElementById("open");
        const phaseMap = {phase_map};
        const weights = {{
          starting: 5,
          venv: 15,
          pip_upgrade: 30,
          pip_install: 55,
          server: 80,
          done: 100,
          error: 100
        }};
        const msgMap = {msg_map};
        let timer = null;
        const poll = async () => {{
          try {{
            const res = await fetch("/status", {{ cache: "no-store" }});
            const j = await res.json();
            phaseEl.textContent = phaseMap[j.phase] || j.phase || "starting";
            const rawMsg = j.message || "";
            let displayMsg = rawMsg;
            // Simple substring matching for improved UX
            for (const [key, val] of Object.entries(msgMap)) {{
               if (rawMsg.includes(key)) {{
                   displayMsg = val;
                   break;
               }}
            }}
            msgEl.textContent = displayMsg;
            bar.value = weights[j.phase] ?? 10;
            if (j.log_tail) {{
              logEl.textContent = j.log_tail;
              logEl.scrollTop = logEl.scrollHeight;
            }}
            if (j.app_up && j.app_url) {{
              // Navigate away so this page stops polling before the launcher exits.
              window.location.replace(j.app_url);
              return;
            }}
            if (j.app_url) {{
              openEl.href = j.app_url;
              // Auto-redirect on success
              if (j.phase === "done") {{
                  setTimeout(() => {{
                      window.location.replace(j.app_url);
                  }}, 1500);
              }}
            }}
            if (j.phase === "done" || j.phase === "error") {{
              doneEl.style.display = "block";
              if (j.phase === "error") {{
                  logEl.style.display = "block";
              }}
              if (timer) {{
                clearInterval(timer);
                timer = null;
              }}
            }}
          }} catch (e) {{
            // ignore transient errors
          }}
        }};
        poll();
        timer = setInterval(poll, 1000);
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


def _pick_free_port() -> int:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind(("127.0.0.1", 0))
        return int(s.getsockname()[1])


def _set_app_port(port: int) -> None:
    global APP_PORT, APP_URL
    APP_PORT = int(port)
    APP_URL = f"http://{APP_HOST}:{APP_PORT}"
    _PROGRESS_STATE["app_url"] = APP_URL
    os.environ["ANSWER_SHEET_PORT"] = str(APP_PORT)
    os.environ["PORT"] = str(APP_PORT)



def _health_ok(timeout: float = 1.0) -> bool:
    health_url = f"{APP_URL}/health"
    try:
        import urllib.request
        req = urllib.request.Request(health_url, method="GET")
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            status = getattr(resp, "status", None) or resp.getcode()
            if status != 200:
                return False
            body = resp.read(256).decode("utf-8", errors="ignore").strip().lower()
            if not body:
                return False
            if body.startswith("ok"):
                return True
            if body.startswith("{"):
                try:
                    data = json.loads(body)
                except Exception:
                    return False
                if isinstance(data, dict):
                    status_val = str(data.get("status", "")).strip().lower()
                    if status_val == "ok":
                        return True
                    ok_val = data.get("ok")
                    if ok_val is True:
                        return True
            return False
    except Exception:
        return False


def _wait_for_server(timeout_total: float, per_request_timeout: float = 1.0) -> bool:
    deadline = time.monotonic() + timeout_total
    while time.monotonic() < deadline:
        if _health_ok(timeout=per_request_timeout):
            return True
        time.sleep(0.25)
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

    # macOS: reuse a stable venv across re-downloads/unzips of the repo.
    if sys.platform == "darwin":
        return Path.home() / "Library" / "Application Support" / "AnswerSheetStudio" / "venvs"

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
        p = subprocess.run(
            args,
            cwd=str(REPO_DIR),
            stdout=log,
            stderr=subprocess.STDOUT,
            text=True,
            errors="replace",
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0) if os.name == "nt" else 0,
        )
        return int(p.returncode)


import shutil

def _python_meets_version(py: Path) -> bool:
    probe = "import sys; raise SystemExit(0 if sys.version_info >= (3, 10) else 1)"
    try:
        kwargs = {}
        if os.name == "nt":
            kwargs["creationflags"] = getattr(subprocess, "CREATE_NO_WINDOW", 0)
        rc = subprocess.run([str(py), "-c", probe], **kwargs).returncode
        return rc == 0
    except Exception:
        return False


def _resolve_best_python() -> Optional[Path]:
    """
    Attempts to find the best available Python executable for creating the venv.
    
    On Windows:
      1. Tries to use the `py` launcher to find Python 3.11, 3.12, 3.13, or 3.10.
    
    On macOS/Linux:
      1. Tries to find `python3.11`, `python3.12`, `python3.13`, or `python3.10` in PATH.
      
    Returns the absolute path to the executable if found, otherwise None.
    """
    preferred_versions = ["3.11", "3.12", "3.13", "3.10"]
    
    if os.name == "nt":
        # Windows: Use `py` launcher if available
        if not shutil.which("py"):
            return None
            
        for ver in preferred_versions:
            try:
                # Ask `py` for the executable path of a specific version
                # -3.11, -3.12, etc.
                cmd = ["py", f"-{ver}", "-c", "import sys; print(sys.executable)"]
                # We use subprocess directly here (not _run_logged) to just query quietly
                # Prevent window popup with creationflags if possible, though 'py' usually handles it.
                kwargs = {}
                if os.name == "nt":
                    kwargs["creationflags"] = getattr(subprocess, "CREATE_NO_WINDOW", 0)
                
                proc = subprocess.run(
                    cmd, 
                    capture_output=True, 
                    text=True, 
                    check=True,
                    **kwargs
                )
                path_str = proc.stdout.strip()
                if path_str:
                    p = Path(path_str)
                    if p.exists():
                        _log(f"Resolved Python {ver} via py launcher: {p}")
                        return p
            except Exception:
                continue
                
    else:
        # macOS / Linux: Check for specific binary names in PATH
        for ver in preferred_versions:
            bin_name = f"python{ver}"
            path_str = shutil.which(bin_name)
            if path_str:
                p = Path(path_str)
                if p.exists():
                     _log(f"Resolved {bin_name} via PATH: {p}")
                     return p

    return None


def ensure_venv() -> Path:
    py = _venv_python()
    if py.exists():
        if os.name != "nt" and not _python_meets_version(py):
            _set_progress("venv", "Recreating virtual environment (old version).")
            _log("Removing incompatible .venv ...")
            try:
                shutil.rmtree(_venv_dir(), ignore_errors=True)
            except Exception:
                pass
        else:
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

    # Determine which python executable to use for creating the venv
    creator_python = _resolve_best_python()
    if not creator_python:
        # Fallback: use the python running this script
        creator_python = Path(sys.executable)
        _log(f"Using current python for venv creation: {creator_python}")

    rc = _run_logged([str(creator_python), "-m", "venv", str(venv_dir)])
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
    if os.name != "nt":
        if _wait_for_server(5.0):
            _log(f"Server already running at {APP_URL}")
            _PROGRESS_STATE["app_up"] = True
            _set_progress("done", "Server already running.")
            _wizard_pause(f"Server is already running at {APP_URL}")
            return
        if is_port_open(APP_HOST, APP_PORT):
            old_port = APP_PORT
            new_port = _pick_free_port()
            _set_app_port(new_port)
            _log(f"Port {old_port} is in use. Switching to {new_port}.")

    if is_port_open(APP_HOST, APP_PORT):
        _log(f"Port {APP_PORT} is BUSY. Attempting to determine if it is our server...")
        # If it's busy, we might want to kill it if it was a previous instance of THIS app.
        # On Windows, we can use `netstat -ano | findstr :8000` to find the PID.
        if os.name == "nt":
            try:
                cmd = f'netstat -ano | findstr :{APP_PORT}'
                res = subprocess.run(cmd, shell=True, capture_output=True, text=True)
                lines = res.stdout.strip().splitlines()
                for line in lines:
                    if "LISTENING" in line:
                        parts = line.strip().split()
                        pid = parts[-1]
                        _log(f"Found process {pid} listening on port {APP_PORT}. Killing it for a fresh start.")
                        subprocess.run(f"taskkill /F /PID {pid} /T", shell=True, capture_output=True, creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
                time.sleep(1) # Wait for release
            except Exception as e:
                _log(f"Failed to kill existing process: {e}")

    if os.name == "nt" and is_port_open(APP_HOST, APP_PORT):
        _log(f"Server already running at {APP_URL}")
        _PROGRESS_STATE["app_up"] = True
        _set_progress("done", "Server already running.")
        _wizard_pause(f"Server is already running at {APP_URL}")
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

        for _ in range(80):  # ~40 seconds
            # Use /health endpoint to verify the app is really ready
            if _health_ok(timeout=1.0):
                _log(f"Server is healthy at {APP_URL}")
                _PROGRESS_STATE["app_up"] = True
                _set_progress("done", "Server is ready.")
                _wizard_pause(f"Server is ready at {APP_URL}")
                return

            if proc.poll() is not None:
                log_tail = _tail_text(SERVER_LOG)
                if "Address already in use" in log_tail or "EADDRINUSE" in log_tail:
                    _set_progress(
                        "error",
                        f"Port {APP_PORT} is already in use. Set ANSWER_SHEET_PORT and retry.",
                    )
                    raise SystemExit(f"Port {APP_PORT} already in use. See {SERVER_LOG}")
                _set_progress("error", f"Server exited (code {proc.returncode}). See server.log.")
                raise SystemExit(f"Server exited (code {proc.returncode}). See {SERVER_LOG}")
            time.sleep(0.5)

    _set_progress("error", "Server not reachable (see server.log).")
    raise SystemExit(f"Server not reachable at {APP_URL}. See {SERVER_LOG}")


def main() -> None:
    progress = _ProgressServer()
    progress_started = False
    try:
        if not CLI_PROGRESS:
            progress.start()
            progress_started = True
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
                if progress_started:
                    # Give the progress page time to load and navigate away.
                    # When OPEN_BROWSER=0 (e.g. start_windows.vbs opens the progress URL itself),
                    # keep the progress server alive longer to avoid a race where the launcher
                    # exits before the browser actually loads the page.
                    try:
                        time.sleep(15 if not OPEN_BROWSER else 3)
                    except Exception:
                        pass
                progress.stop()
        except Exception:
            pass


if __name__ == "__main__":
    main()
