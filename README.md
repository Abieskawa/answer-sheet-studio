# Answer Sheet Studio

English | 繁體中文: `README.zh-Hant.md`

Download (ZIP): https://github.com/Abieskawa/answer-sheet-studio/archive/refs/heads/main.zip
Docs (Read the Docs): https://answer-sheet-studio.readthedocs.io/zh-tw/latest/

Answer Sheet Studio lets teachers generate printable answer sheets and run local OMR recognition (no cloud upload):

- **Download page** – choose title, subject, class, number of questions (up to 100), and choices per question (`ABC` / `ABCD` / `ABCDE`). Generates a single-page PDF template (macOS/Windows compatible).
- **Upload page** – drop a multi-page PDF scan. Recognition exports `results.csv` (questions as rows, students as columns), `ambiguity.csv` (blank/ambiguous/multi picks), plus `annotated.pdf` that visualizes detections.
- **Launcher** – double-click (`start_mac.command` or `start_windows.vbs`) to create a virtual env, install dependencies, start the local server, and open `http://127.0.0.1:8000`.
- The web UI supports **English** and **Traditional Chinese**. Use the language tabs in the header.

## Requirements

- macOS ≥ 13 or Windows 11
- Python **3.10+** (3.11 recommended; 3.10–3.13 supported). If Python isn’t installed yet, the launcher can help download the official installer.
- On Windows, ensure “Add python.exe to PATH” during installation (and keep the `py` launcher enabled if offered).
- Internet access the first time to download Python packages (FastAPI, PyMuPDF, OpenCV, NumPy, etc.).

## Quick Start

### macOS
1. Double-click `start_mac.command` (or run `chmod +x start_mac.command` once if prompted). If Python isn’t installed, it will offer to download and open the installer.
2. First run creates `.venv` and installs requirements; later runs reuse the existing `.venv` (unless requirements changed).
3. Your browser opens `http://127.0.0.1:8000`. Close the browser when you’re done; the server auto-exits after a period of inactivity.

### Windows 11
1. Double-click `start_windows.vbs`. If Python isn’t installed, it will offer to download/install Python 3.11 automatically (recommended).
2. If you install Python manually, ensure “Add python.exe to PATH” during installation (and keep the `py` launcher enabled if offered).
3. First run installs requirements; later runs reuse the existing `.venv` (unless requirements changed). If Windows Defender prompts for network access, allow it so the server can bind to localhost.

## Important Notes

- For recognition, the **number of questions** and **choices per question** must match the generated answer sheet.
- For more stable results: use a darker pen/pencil, scan at **300dpi**, and avoid skew/rotation.

## Output Files

After recognition, files are written under `outputs/<job_id>/`:

- `results.csv`
- `ambiguity.csv`
- `annotated.pdf`
- `input.pdf` (original upload)

## Updating

- Open `http://127.0.0.1:8000/update` and upload the latest ZIP (from GitHub Releases or `main.zip`). The app will restart automatically.

## Debug Mode

- Open `http://127.0.0.1:8000/debug` and enter the Job ID (the folder name under `outputs/`) to download diagnostic files (including `ambiguity.csv`).

## Troubleshooting

- **pip install failed** – Check the launcher log; ensure network access and Python 3.10+. If you’re on Python 3.14, try Python 3.10–3.13.
- **Port already in use / permission denied** – The launcher checks port availability. If macOS firewall blocks Python, allow incoming connections in System Settings > Network > Firewall.
- **Server not opening** – Use the launcher log to identify crashes; you can also run `python run_app.py` inside `.venv` manually for debugging.
