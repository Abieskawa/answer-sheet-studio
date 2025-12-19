import errno
import hashlib
import json
import os
import sys
import shutil
import subprocess
import threading
import queue
import webbrowser
import zipfile
import socket
import time
from pathlib import Path
from typing import List, Optional, Tuple
import tkinter as tk
from tkinter import messagebox

APP_HOST = "127.0.0.1"
APP_PORT = int(os.environ.get("ANSWER_SHEET_PORT", "8000")) if "ANSWER_SHEET_PORT" in os.environ else 8000
APP_URL = f"http://{APP_HOST}:{APP_PORT}"
DEFAULT_ZIP_NAME = "answer-sheet-studio.zip"

SKIP_DIR_NAMES = {".git", ".venv", "__pycache__", "_incoming", "node_modules", "outputs", ".DS_Store", "Thumbs.db"}

SUPPORTED_PYTHON_MIN = (3, 10)
SUPPORTED_PYTHON_MAX_EXCLUSIVE = (3, 14)  # 3.14+ may lack wheels for numpy/opencv on some platforms
PREFERRED_PYTHON_MINORS = (10, 11, 12, 13)
INSTALL_MARKER_NAME = ".answer_sheet_studio_install.json"

WINDOW_BG = "#f5f5f7"
CARD_BG = "#ffffff"
CARD_BORDER = "#e2e8f0"
HEADING_FG = "#0f172a"
SUBTEXT_FG = "#475569"
CONSOLE_BG = "#0f172a"
CONSOLE_FG = "#e5e7eb"
STATUS_STYLES = {
    "info": {"bg": "#dbeafe", "fg": "#1e3a8a"},
    "success": {"bg": "#dcfce7", "fg": "#166534"},
    "warning": {"bg": "#fef9c3", "fg": "#854d0e"},
    "error": {"bg": "#fee2e2", "fg": "#b91c1c"},
    "pending": {"bg": "#ede9fe", "fg": "#5b21b6"},
}

def is_port_open(host: str, port: int, timeout=0.25) -> bool:
    try:
        with socket.create_connection((host, port), timeout=timeout):
            return True
    except OSError:
        return False

def _format_py_version(v: Tuple[int, int]) -> str:
    return f"{v[0]}.{v[1]}"

def _is_supported_python(v: Tuple[int, int]) -> bool:
    return SUPPORTED_PYTHON_MIN <= v < SUPPORTED_PYTHON_MAX_EXCLUSIVE

def _read_venv_version(venv_dir: Path) -> Optional[Tuple[int, int]]:
    cfg = venv_dir / "pyvenv.cfg"
    if not cfg.exists():
        return None
    try:
        raw = cfg.read_text(encoding="utf-8", errors="ignore")
    except Exception:
        return None
    for line in raw.splitlines():
        if line.lower().startswith("version"):
            parts = line.split("=", 1)
            if len(parts) != 2:
                continue
            v_str = parts[1].strip()
            segs = v_str.split(".")
            if len(segs) < 2:
                continue
            try:
                return int(segs[0]), int(segs[1])
            except ValueError:
                return None
    return None

def _sha256_file(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()

def _load_install_marker(path: Path) -> Optional[dict]:
    if not path.exists():
        return None
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return None

def _write_install_marker(path: Path, data: dict) -> bool:
    try:
        path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        return True
    except Exception:
        return False

def _probe_python_cmd(cmd_prefix: List[str]) -> Optional[Tuple[Tuple[int, int], str]]:
    probe_code = (
        "import sys; "
        "print('PYVER=%d.%d' % (sys.version_info[0], sys.version_info[1])); "
        "print('PYEXE=' + sys.executable)"
    )
    try:
        out = subprocess.check_output(
            cmd_prefix + ["-c", probe_code],
            stderr=subprocess.STDOUT,
            text=True,
        )
    except Exception:
        return None

    ver_line = None
    exe_line = None
    for line in out.splitlines():
        if line.startswith("PYVER="):
            ver_line = line
        elif line.startswith("PYEXE="):
            exe_line = line
    if not ver_line or not exe_line:
        return None

    v_str = ver_line.split("=", 1)[1].strip()
    segs = v_str.split(".")
    if len(segs) < 2:
        return None
    try:
        v = (int(segs[0]), int(segs[1]))
    except ValueError:
        return None
    exe = exe_line.split("=", 1)[1].strip()
    return v, exe

def _select_venv_builder_cmd() -> Optional[Tuple[List[str], Tuple[int, int], str]]:
    """
    Return (cmd_prefix, (major, minor), executable_path) for creating a venv.
    Preference:
      1) current interpreter if supported
      2) Windows: `py -3.10`/`py -3.11`/... if available
      3) macOS/Linux: `python3.10`/`python3.11`/... if available
      4) `python3`/`python` if supported
    """
    cur = (sys.version_info[0], sys.version_info[1])
    if _is_supported_python(cur):
        return [sys.executable], cur, sys.executable

    if os.name == "nt" and shutil.which("py"):
        for minor in PREFERRED_PYTHON_MINORS:
            cmd = ["py", f"-3.{minor}"]
            probe = _probe_python_cmd(cmd)
            if probe is None:
                continue
            v, exe = probe
            if _is_supported_python(v):
                return cmd, v, exe

    for minor in PREFERRED_PYTHON_MINORS:
        exe = shutil.which(f"python3.{minor}")
        if not exe:
            continue
        probe = _probe_python_cmd([exe])
        if probe is None:
            continue
        v, exe_path = probe
        if _is_supported_python(v):
            return [exe], v, exe_path

    for exe_name in ("python3", "python"):
        exe = shutil.which(exe_name)
        if not exe:
            continue
        probe = _probe_python_cmd([exe])
        if probe is None:
            continue
        v, exe_path = probe
        if _is_supported_python(v):
            return [exe], v, exe_path

    return None

def check_port_available(host: str, port: int):
    """Return (True, None) if port can be bound, otherwise (False, reason)."""
    sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
    try:
        sock.bind((host, port))
    except PermissionError as exc:
        reason = (
            f"Permission denied when binding to {host}:{port}.\n\n"
            "macOS blocked Python from accepting incoming connections. "
            "Allow Python in System Settings → Network → Firewall (or temporarily disable the firewall) "
            "or choose another port by setting ANSWER_SHEET_PORT before launching."
        )
        return False, reason
    except OSError as exc:
        if exc.errno == errno.EADDRINUSE:
            return False, f"Port {port} is already in use. Close the other process or set ANSWER_SHEET_PORT to another value."
        return False, f"Failed to reserve port {port}: {exc}"
    finally:
        sock.close()
    return True, None

def find_project_files(repo_dir: Path, max_depth: int = 4):
    """
    Try to find requirements.txt and run_app.py in repo root or within a few subfolders.
    Returns (requirements_path, run_app_path) or (None, None).
    Uses a bounded BFS (depth-limited) to avoid scanning huge folders.
    """
    req = repo_dir / "requirements.txt"
    run = repo_dir / "run_app.py"
    if req.exists() and run.exists():
        return req, run

    # BFS over directories to a fixed depth
    queue_dirs = [(repo_dir, 0)]
    req_candidates = []
    run_candidates = []

    while queue_dirs:
        d, depth = queue_dirs.pop(0)
        if depth > max_depth:
            continue
        try:
            for entry in d.iterdir():
                if entry.is_dir():
                    if entry.name in SKIP_DIR_NAMES:
                        continue
                    queue_dirs.append((entry, depth + 1))
                else:
                    if entry.name == "requirements.txt":
                        req_candidates.append(entry)
                    elif entry.name == "run_app.py":
                        run_candidates.append(entry)
        except (PermissionError, FileNotFoundError):
            continue

    # Prefer closest to repo root
    req_candidates.sort(key=lambda p: len(p.relative_to(repo_dir).parts))
    run_candidates.sort(key=lambda p: len(p.relative_to(repo_dir).parts))

    if req_candidates and run_candidates:
        return req_candidates[0], run_candidates[0]
    return None, None

def merge_zip_into_repo(repo_dir: Path, zip_path: Path, log_fn):
    """
    Extract zip into _incoming, then merge its top-level files into repo_dir.
    If zip contains a single top folder, merge from inside that folder.
    """
    incoming = repo_dir / "_incoming"
    if incoming.exists():
        shutil.rmtree(incoming, ignore_errors=True)
    incoming.mkdir(parents=True, exist_ok=True)

    log_fn(f"Extracting: {zip_path}")
    with zipfile.ZipFile(zip_path, "r") as z:
        z.extractall(incoming)

    # If the zip has a single top folder, use it
    children = [p for p in incoming.iterdir() if p.name != ".DS_Store"]
    root = children[0] if (len(children) == 1 and children[0].is_dir()) else incoming

    log_fn(f"Merging files from: {root}")
    for item in root.iterdir():
        if item.name in SKIP_DIR_NAMES:
            continue
        dest = repo_dir / item.name
        if item.is_dir():
            shutil.copytree(item, dest, dirs_exist_ok=True)
        else:
            shutil.copy2(item, dest)

class LauncherApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Answer Sheet Studio Launcher")
        self.geometry("1020x680")

        self.repo_dir = Path(__file__).resolve().parent
        self.proc = None
        self.log_q = queue.Queue()

        # surface exceptions (avoid "nothing happens")
        self.report_callback_exception = self._report_tk_exception

        self._build_ui()
        self.after(120, self._drain_log_queue)

    def _report_tk_exception(self, exc, val, tb):
        import traceback
        msg = "".join(traceback.format_exception(exc, val, tb))
        self._append_log("ERROR:\n" + msg)
        messagebox.showerror("Launcher error", msg)

    def _build_ui(self):
        self.configure(bg=WINDOW_BG)

        root = tk.Frame(self, padx=20, pady=20, bg=WINDOW_BG)
        root.pack(fill="both", expand=True)

        header = tk.Frame(root, bg=WINDOW_BG)
        header.pack(fill="x", pady=(0, 12))
        tk.Label(
            header,
            text="Answer Sheet Studio Launcher",
            font=("Helvetica", 20, "bold"),
            fg=HEADING_FG,
            bg=WINDOW_BG,
            anchor="w",
        ).pack(fill="x")
        tk.Label(
            header,
            text="Bootstrap the environment, start the server, and open the web UI without touching Terminal.",
            font=("Helvetica", 12),
            fg=SUBTEXT_FG,
            bg=WINDOW_BG,
            anchor="w",
            wraplength=960,
            justify="left",
        ).pack(fill="x", pady=(4, 0))

        info_card = tk.Frame(
            root,
            bg=CARD_BG,
            highlightbackground=CARD_BORDER,
            highlightcolor=CARD_BORDER,
            highlightthickness=1,
            bd=0,
            padx=18,
            pady=18,
        )
        info_card.pack(fill="x", pady=(0, 18))

        tk.Label(
            info_card,
            text=f"Repository: {self.repo_dir}",
            font=("Helvetica", 12, "bold"),
            fg=HEADING_FG,
            bg=CARD_BG,
            anchor="w",
        ).pack(fill="x")

        tk.Label(
            info_card,
            text="Tip: 'Install & Run' creates .venv, installs requirements, and launches run_app.py. Re-run it after pulling updates.",
            font=("Helvetica", 11),
            fg=SUBTEXT_FG,
            bg=CARD_BG,
            anchor="w",
            wraplength=960,
            justify="left",
        ).pack(fill="x", pady=(6, 0))

        btns = tk.Frame(info_card, bg=CARD_BG)
        btns.pack(fill="x", pady=(12, 0))

        def add_btn(text, cmd):
            tk.Button(
                btns,
                text=text,
                command=cmd,
                padx=18,
                pady=8,
            ).pack(side="left", padx=(0, 10))

        add_btn("Update (git pull)", self.on_git_pull)
        add_btn("Install & Run", self.on_install_and_run)
        add_btn("Open in Browser", lambda: webbrowser.open(APP_URL))
        add_btn("Stop Server", self.on_stop)

        status_frame = tk.Frame(root, bg=WINDOW_BG)
        status_frame.pack(fill="x", pady=(0, 14))
        self.status_var = tk.StringVar(value="")
        self.status_label = tk.Label(
            status_frame,
            textvariable=self.status_var,
            anchor="w",
            font=("Helvetica", 13, "bold"),
            padx=16,
            pady=10,
            bg=STATUS_STYLES["info"]["bg"],
            fg=STATUS_STYLES["info"]["fg"],
        )
        self.status_label.pack(fill="x")
        self._update_status("Ready.", "info")

        log_card = tk.Frame(
            root,
            bg=CARD_BG,
            highlightbackground=CARD_BORDER,
            highlightcolor=CARD_BORDER,
            highlightthickness=1,
            bd=0,
            padx=0,
            pady=12,
        )
        log_card.pack(fill="both", expand=True)

        log_header = tk.Frame(log_card, bg=CARD_BG)
        log_header.pack(fill="x", padx=20, pady=(0, 10))
        tk.Label(
            log_header,
            text="Command log",
            font=("Helvetica", 14, "bold"),
            fg=HEADING_FG,
            bg=CARD_BG,
            anchor="w",
        ).pack(fill="x")
        tk.Label(
            log_header,
            text="This captures every command, pip output, and server log so you can see what happened.",
            font=("Helvetica", 11),
            fg=SUBTEXT_FG,
            bg=CARD_BG,
            anchor="w",
            wraplength=960,
            justify="left",
        ).pack(fill="x")

        console = tk.Frame(log_card, bg=CARD_BG)
        console.pack(fill="both", expand=True, padx=14, pady=(0, 4))

        self.log = tk.Text(
            console,
            wrap="word",
            bg=CONSOLE_BG,
            fg=CONSOLE_FG,
            insertbackground="#10b981",
            font=("Menlo", 12),
            borderwidth=0,
            highlightthickness=0,
            padx=16,
            pady=14,
        )
        self.log.pack(side="left", fill="both", expand=True)

        sb = tk.Scrollbar(console, command=self.log.yview)
        sb.pack(side="right", fill="y")
        self.log.configure(yscrollcommand=sb.set)
        self.log.tag_configure("cmd", foreground="#93c5fd")
        self.log.tag_configure("error", foreground="#fda4af")
        self.log.tag_configure("notice", foreground="#86efac")
        self.log.tag_configure("muted", foreground="#94a3b8")

        intro = (
            "Launcher ready.\n"
            "1) Click 'Install & Run' to set up dependencies and start the FastAPI server.\n"
            "2) Click 'Open in Browser' any time after the server reports that it is running.\n"
            "3) Use 'Stop Server' to terminate the background process.\n\n"
        )
        self.log.insert("end", intro, "muted")
        self.log.configure(state="disabled")

    def _update_status(self, text: str, tone: str = "info"):
        style = STATUS_STYLES.get(tone, STATUS_STYLES["info"])
        self.status_var.set(text)
        self.status_label.configure(bg=style["bg"], fg=style["fg"])

    def _append_log(self, s: str):
        self.log.configure(state="normal")
        tag = None
        if s.startswith("$ "):
            tag = "cmd"
        elif s.upper().startswith("ERROR"):
            tag = "error"
        elif s.startswith("Clicked"):
            tag = "notice"
        text = s if s.endswith("\n") else s + "\n"
        self.log.insert("end", text, tag)
        self.log.see("end")
        self.log.configure(state="disabled")

    def _drain_log_queue(self):
        try:
            while True:
                self._append_log(self.log_q.get_nowait())
        except queue.Empty:
            pass
        self.after(120, self._drain_log_queue)

    def _in_thread(self, fn):
        threading.Thread(target=fn, daemon=True).start()

    def _run_cmd_stream(self, cmd, cwd=None, env=None):
        self.log_q.put(f"$ {' '.join(cmd)}")
        creationflags = 0
        if os.name == "nt" and hasattr(subprocess, "CREATE_NO_WINDOW"):
            creationflags = subprocess.CREATE_NO_WINDOW
        p = subprocess.Popen(
            cmd,
            cwd=str(cwd) if cwd else None,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            shell=False,
            creationflags=creationflags,
            env=env,
        )
        for line in p.stdout:
            self.log_q.put(line.rstrip("\n"))
        return p.wait()

    def on_git_pull(self):
        def task():
            self._update_status("Running git pull...", "pending")
            if not (self.repo_dir / ".git").exists():
                messagebox.showinfo(
                    "Update not available",
                    "This folder is not a git clone (no .git folder).\n\n"
                    "If you downloaded a ZIP, re-download the latest ZIP to update.\n"
                    "If you want one-click updates, clone the repo with Git."
                )
                self._update_status("Update not available (not a git clone).", "warning")
                return

            git_exe = shutil.which("git")
            if git_exe is None:
                messagebox.showerror("Git not found", "Git is not installed or not in PATH.")
                self._update_status("Git not found in PATH.", "error")
                return

            env = os.environ.copy()
            env["GIT_TERMINAL_PROMPT"] = "0"
            if os.name == "nt":
                userprofile = env.get("USERPROFILE")
                if userprofile:
                    try:
                        userprofile_ok = Path(userprofile).exists()
                    except Exception:
                        userprofile_ok = False
                    if userprofile_ok:
                        home = env.get("HOME")
                        try:
                            home_ok = bool(home) and Path(home).exists()
                        except Exception:
                            home_ok = False
                        if not home_ok:
                            env["HOME"] = userprofile

                        homedrive = env.get("HOMEDRIVE")
                        homepath = env.get("HOMEPATH")
                        if homedrive and homepath:
                            try:
                                if not Path(homedrive + homepath).exists():
                                    env["HOMEDRIVE"] = userprofile[:2]
                                    env["HOMEPATH"] = userprofile[2:]
                            except Exception:
                                pass

            self.log_q.put(f"Using git: {git_exe}")
            rc = self._run_cmd_stream([git_exe, "-C", str(self.repo_dir), "pull"], env=env)
            if rc == 0:
                self._update_status("git pull completed.", "success")
            else:
                self._update_status(f"git pull failed (code {rc}).", "error")
                messagebox.showerror(
                    "git pull failed",
                    "git pull failed. Check the Command log for details.\n\n"
                    "Common fixes:\n"
                    "- Make sure you have internet access.\n"
                    "- If you have local edits, commit/stash them first.\n"
                    "- If the error mentions a missing drive, move the folder to a normal path like C:\\Users\\... and retry."
                )

        self._in_thread(task)

    def _venv_python(self) -> Path:
        if os.name == "nt":
            return self.repo_dir / ".venv" / "Scripts" / "python.exe"
        return self.repo_dir / ".venv" / "bin" / "python"

    def on_install_and_run(self):
        def task():
            self.log_q.put("Clicked: Install & Run")
            self.log_q.put(f"Launcher Python: {sys.executable} (Python {_format_py_version((sys.version_info[0], sys.version_info[1]))})")
            # If port already open, just open browser (server might already be running)
            if is_port_open(APP_HOST, APP_PORT):
                self._update_status("Server already running. Opening browser...", "success")
                self.log_q.put(f"Port {APP_PORT} is open, opening {APP_URL}")
                webbrowser.open(APP_URL)
                return

            req_path, run_path = find_project_files(self.repo_dir)

            # If missing, offer auto-extract from zip
            if not req_path or not run_path:
                zip_path = self.repo_dir / DEFAULT_ZIP_NAME
                if zip_path.exists():
                    self.log_q.put("App code not found in folder.")
                    do_extract = messagebox.askyesno(
                        "App code not found",
                        "I cannot find requirements.txt and run_app.py in this folder.\n\n"
                        "But I found answer-sheet-studio.zip.\n"
                        "Do you want me to extract it into this repo folder now?"
                    )
                    if do_extract:
                        try:
                            merge_zip_into_repo(self.repo_dir, zip_path, self.log_q.put)
                        except Exception as e:
                            messagebox.showerror("Extract failed", str(e))
                            self._update_status("Zip extraction failed.", "error")
                            return
                        # re-scan
                        req_path, run_path = find_project_files(self.repo_dir)

            if not req_path or not run_path:
                self.log_q.put("Missing project files after auto-fix attempt.")
                messagebox.showerror(
                    "Missing files",
                    "Still cannot find BOTH requirements.txt and run_app.py.\n\n"
                    "Make sure your repo folder contains the app code.\n"
                    "Tip: If you have a ZIP, unzip it and ensure run_app.py + requirements.txt exist."
                )
                self._update_status("Project files not found.", "error")
                return

            project_dir = run_path.parent
            self.log_q.put(f"Using project_dir: {project_dir}")
            self.log_q.put(f"requirements: {req_path}")
            self.log_q.put(f"runner: {run_path}")

            ok, reason = check_port_available(APP_HOST, APP_PORT)
            if not ok:
                self._update_status("Cannot use the selected port.", "error")
                messagebox.showerror("Port unavailable", reason)
                return

            venv_builder = _select_venv_builder_cmd()
            if venv_builder is None:
                detected = _format_py_version((sys.version_info[0], sys.version_info[1]))
                supported = f"{_format_py_version(SUPPORTED_PYTHON_MIN)}–3.{SUPPORTED_PYTHON_MAX_EXCLUSIVE[1] - 1}"
                messagebox.showerror(
                    "Unsupported Python",
                    "This project uses packages (NumPy/OpenCV/PyMuPDF/ReportLab) that require prebuilt wheels.\n\n"
                    f"Detected Python {detected}.\n"
                    f"Supported Python versions: {supported}.\n\n"
                    "Install Python 3.10 (recommended) or 3.11/3.12/3.13 from python.org, then re-run the launcher."
                )
                self._update_status("Unsupported Python version.", "error")
                return
            venv_builder_cmd, venv_builder_ver, venv_builder_exe = venv_builder
            self.log_q.put(f"Selected venv Python: {venv_builder_exe} (Python {_format_py_version(venv_builder_ver)})")

            # Create venv
            self._update_status("Creating virtual environment (if needed)...", "pending")
            venv_dir = self.repo_dir / ".venv"
            py = self._venv_python()

            existing_ver = _read_venv_version(venv_dir) if venv_dir.exists() else None
            if venv_dir.exists() and existing_ver and not _is_supported_python(existing_ver):
                do_recreate = messagebox.askyesno(
                    "Recreate virtual environment?",
                    f"Your existing .venv uses Python {_format_py_version(existing_ver)}, which is not supported.\n\n"
                    "Recreate .venv using a supported Python version now? (This will delete .venv and reinstall packages.)"
                )
                if not do_recreate:
                    self._update_status("Install canceled (unsupported .venv).", "warning")
                    return
                shutil.rmtree(venv_dir, ignore_errors=True)

            if not venv_dir.exists() or not py.exists():
                rc = self._run_cmd_stream(venv_builder_cmd + ["-m", "venv", str(venv_dir)], cwd=self.repo_dir)
                if rc != 0:
                    messagebox.showerror("Venv failed", "Failed to create virtual environment (.venv). Check log.")
                    self._update_status("Failed to create virtual environment.", "error")
                    return

            if not py.exists():
                messagebox.showerror("Venv error", f"Venv python not found: {py}")
                self._update_status("Virtual environment is incomplete.", "error")
                return

            # Install deps (only when needed)
            self._update_status("Checking dependencies...", "pending")
            marker_path = venv_dir / INSTALL_MARKER_NAME
            req_sha = _sha256_file(req_path)
            venv_ver = _read_venv_version(venv_dir)
            venv_ver_str = _format_py_version(venv_ver) if venv_ver else None

            marker = _load_install_marker(marker_path)
            marker_ok = (
                bool(marker)
                and marker.get("requirements_sha256") == req_sha
                and marker.get("python_version") == venv_ver_str
            )

            if marker_ok:
                self.log_q.put("Dependencies already installed (requirements unchanged).")
                rc = self._run_cmd_stream(
                    [str(py), "-c", "import fastapi, uvicorn, fitz, cv2, numpy, reportlab; print('deps ok')"],
                    cwd=self.repo_dir,
                )
                if rc != 0:
                    marker_ok = False
                    self.log_q.put("Dependency import check failed; reinstalling...")

            if not marker_ok:
                self._update_status("Installing dependencies (this can take a few minutes)...", "pending")
                rc = self._run_cmd_stream([str(py), "-m", "pip", "install", "--upgrade", "pip"], cwd=self.repo_dir)
                if rc != 0:
                    messagebox.showerror("pip failed", "pip upgrade failed. Check log.")
                    self._update_status("pip upgrade failed.", "error")
                    return

                rc = self._run_cmd_stream(
                    [str(py), "-m", "pip", "install", "--only-binary=:all:", "-r", str(req_path)],
                    cwd=self.repo_dir,
                )
                if rc != 0:
                    messagebox.showerror("pip failed", "pip install failed. Check log for the error.")
                    self._update_status("pip install failed.", "error")
                    return

                marker_data = {
                    "python_version": venv_ver_str or _format_py_version(venv_builder_ver),
                    "requirements_sha256": req_sha,
                    "requirements_file": str(req_path.name),
                    "installed_at": int(time.time()),
                }
                if not _write_install_marker(marker_path, marker_data):
                    self.log_q.put("WARNING: Failed to write install marker; next run may reinstall dependencies.")

            # Start server
            self._update_status("Starting server...", "pending")
            self.log_q.put(f"Starting: {run_path.name}")
            self.proc = subprocess.Popen(
                [str(py), str(run_path.name)],
                cwd=str(project_dir),
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
            )

            def stream():
                for line in self.proc.stdout:
                    self.log_q.put(line.rstrip("\n"))

            threading.Thread(target=stream, daemon=True).start()

            # Wait for server to become reachable
            for _ in range(60):  # ~30 seconds
                if is_port_open(APP_HOST, APP_PORT):
                    self._update_status("Server running. Opening browser...", "success")
                    self.log_q.put(f"Server is up: {APP_URL}")
                    webbrowser.open(APP_URL)
                    return
                if self.proc.poll() is not None:
                    self._update_status("Server exited (failed).", "error")
                    self.log_q.put(f"Server process exited with code {self.proc.returncode}")
                    messagebox.showerror("Server failed", "Server exited immediately. Check log for details.")
                    return
                time.sleep(0.5)

            self._update_status("Server not reachable yet.", "warning")
            messagebox.showwarning(
                "Not reachable",
                f"Server did not become reachable at {APP_URL}.\n"
                "Check log output to see what it is doing."
            )

        self._in_thread(task)

    def on_stop(self):
        if self.proc and self.proc.poll() is None:
            try:
                self.proc.terminate()
                self._update_status("Server stopped.", "info")
                self.log_q.put("Server stopped.")
            except Exception as e:
                messagebox.showerror("Stop failed", str(e))
        else:
            self._update_status("No server running.", "warning")

if __name__ == "__main__":
    LauncherApp().mainloop()
