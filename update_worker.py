import argparse
import os
import shutil
import socket
import subprocess
import sys
import time
import zipfile
from pathlib import Path


SKIP_DIR_NAMES = {
    ".git",
    ".venv",
    "__pycache__",
    "_incoming",
    "node_modules",
    "outputs",
    ".DS_Store",
    "Thumbs.db",
}


REPO_DIR = Path(__file__).resolve().parent
OUTPUTS_DIR = REPO_DIR / "outputs"
OUTPUTS_DIR.mkdir(exist_ok=True)
LOG_PATH = OUTPUTS_DIR / "update.log"


def _log(line: str) -> None:
    ts = time.strftime("%Y-%m-%d %H:%M:%S")
    try:
        with open(LOG_PATH, "a", encoding="utf-8") as f:
            f.write(f"[{ts}] {line}\n")
    except Exception:
        pass


def _subprocess_kwargs() -> dict:
    if os.name == "nt":
        return {"creationflags": getattr(subprocess, "CREATE_NO_WINDOW", 0)}
    return {}


def _run_capture(args: list[str]) -> tuple[int, str]:
    _log(f"$ {' '.join(args)}")
    p = subprocess.run(
        args,
        cwd=str(REPO_DIR),
        capture_output=True,
        text=True,
        errors="replace",
        **_subprocess_kwargs(),
    )
    out = (p.stdout or "") + (p.stderr or "")
    if out.strip():
        for line in out.rstrip("\n").splitlines():
            _log(line)
    return int(p.returncode), (p.stdout or "")


def is_port_open(host: str, port: int, timeout: float = 0.25) -> bool:
    try:
        with socket.create_connection((host, port), timeout=timeout):
            return True
    except OSError:
        return False


def merge_zip_into_repo(repo_dir: Path, zip_path: Path) -> None:
    incoming = repo_dir / "_incoming"
    if incoming.exists():
        shutil.rmtree(incoming, ignore_errors=True)
    incoming.mkdir(parents=True, exist_ok=True)

    _log(f"Extracting: {zip_path}")
    with zipfile.ZipFile(zip_path, "r") as z:
        z.extractall(incoming)

    children = [p for p in incoming.iterdir() if p.name != ".DS_Store"]
    root = children[0] if (len(children) == 1 and children[0].is_dir()) else incoming

    _log(f"Merging files from: {root}")
    for item in root.iterdir():
        if item.name in SKIP_DIR_NAMES:
            continue
        dest = repo_dir / item.name
        if item.is_dir():
            shutil.copytree(item, dest, dirs_exist_ok=True)
        else:
            shutil.copy2(item, dest)

    shutil.rmtree(incoming, ignore_errors=True)


def _spawn_launcher_headless() -> None:
    launcher = REPO_DIR / "launcher_headless.py"
    if not launcher.exists():
        raise SystemExit("launcher_headless.py not found after update.")

    env = os.environ.copy()
    env["ANSWER_SHEET_OPEN_BROWSER"] = "0"

    exe = sys.executable
    if os.name == "nt":
        pyw = Path(sys.executable).with_name("pythonw.exe")
        if pyw.exists():
            exe = str(pyw)

    kwargs: dict = {"cwd": str(REPO_DIR), "env": env}
    if os.name == "nt":
        creationflags = 0
        creationflags |= getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0)
        creationflags |= getattr(subprocess, "DETACHED_PROCESS", 0)
        creationflags |= getattr(subprocess, "CREATE_NO_WINDOW", 0)
        kwargs["creationflags"] = creationflags
    else:
        kwargs["start_new_session"] = True

    subprocess.Popen([exe, str(launcher)], **kwargs)


def main() -> None:
    ap = argparse.ArgumentParser()
    mode = ap.add_mutually_exclusive_group(required=True)
    mode.add_argument("--zip", help="Path to the update ZIP file")
    mode.add_argument("--git", action="store_true", help="Run git pull in the current repo")
    ap.add_argument("--host", default="127.0.0.1")
    ap.add_argument("--port", type=int, default=8000)
    ap.add_argument("--wait-seconds", type=int, default=120)
    args = ap.parse_args()

    _log("Update worker started.")
    _log(f"Waiting for server to stop at {args.host}:{args.port} ...")
    deadline = time.time() + max(1, int(args.wait_seconds))
    while time.time() < deadline:
        if not is_port_open(args.host, int(args.port)):
            break
        time.sleep(0.5)

    if is_port_open(args.host, int(args.port)):
        _log("WARNING: Server still reachable; applying update anyway.")
    else:
        _log("Server stopped. Applying update.")

    try:
        if args.git:
            git = shutil.which("git")
            if not git:
                _log("ERROR: git not found.")
            elif not (REPO_DIR / ".git").exists():
                _log("ERROR: .git not found. This is not a git repository.")
            else:
                rc, out = _run_capture([git, "status", "--porcelain"])
                if rc != 0:
                    _log("ERROR: git status failed.")
                elif out.strip():
                    _log("ERROR: Working tree not clean; aborting git update.")
                else:
                    rc, _ = _run_capture([git, "pull", "--ff-only"])
                    if rc != 0:
                        _log("ERROR: git pull failed.")
                    else:
                        _log("Git update complete.")
        else:
            zip_path = Path(args.zip or "")
            if not zip_path.exists():
                _log(f"ERROR: ZIP not found: {zip_path}")
            else:
                merge_zip_into_repo(REPO_DIR, zip_path)
                _log("ZIP update complete.")
    finally:
        _log("Starting launcher_headless.py ...")
    _spawn_launcher_headless()
    _log("Done.")


if __name__ == "__main__":
    try:
        main()
    except SystemExit:
        raise
    except Exception as exc:
        _log(f"ERROR: {exc!r}")
        raise
