# Answer Sheet Studio

Answer Sheet Studio lets teachers generate printable answer sheets and run local OMR recognition:

- **Download page** – choose title, subject, class and number of questions (up to 100). Generates a single-page PDF template (macOS/Windows compatible).
- **Upload page** – drop a multi-page PDF scan. Recognition exports `results.csv` plus `annotated.pdf` that visualizes detections (grade, class, seat, answers).
- **Launcher** – double-click launcher (`start_mac.command` or `start_windows.vbs`) to create a virtual env, install dependencies, start the FastAPI server, and open `http://127.0.0.1:8000`.

## Requirements

- macOS ≥ 13 or Windows 11
- Python **3.10–3.13** (3.13 recommended). Python 3.14 preview builds are not supported yet (NumPy/OpenCV/pandas/ReportLab wheels may be missing and pip will try to compile).
- On Windows, ensure “Add python.exe to PATH” during installation (and keep the `py` launcher enabled if offered).
- Internet access the first time to download Python packages (FastAPI, PyMuPDF, OpenCV, pandas, etc.).

## Quick Start

### macOS
1. Double-click `start_mac.command` (or run `chmod +x start_mac.command` once if prompted).
2. The Tk launcher appears; click **Install & Run**. It creates `.venv`, installs requirements, starts uvicorn, and opens the browser automatically.
3. Use **Open in Browser** to revisit `http://127.0.0.1:8000`. Use **Stop Server** to terminate.

### Windows 11
1. Install Python 3.13 (recommended) or 3.12/3.11/3.10 from python.org and check “Add python.exe to PATH”.
2. Double-click `start_windows.vbs`. The launcher GUI will open using `pythonw`.
3. Click **Install & Run**. If Windows Defender prompts for network access, allow it so the server can bind to localhost.

## Repository Layout

```
app/            FastAPI app (routes, templates, static)
omr/            Generator + recognition pipeline (ReportLab + OpenCV/PyMuPDF)
launcher_gui.py Cross-platform Tk launcher
start_mac.command / start_windows.vbs  OS-specific launchers
```

Generated files are written under `outputs/`.

## Customization

- Modify `omr/generator.py` to adjust layout (title, subject label, footer, coordinates).
- Update recognition thresholds or geometry in `omr/recognizer.py` / `omr/config.py`.
- Front-end templates live in `app/templates/`.

## Troubleshooting

- **pip install failed** – Check the launcher log; ensure network access and Python 3.10–3.13. If you’re on Python 3.14, install Python 3.13 and re-run.
- **Port already in use / permission denied** – The launcher checks port availability. If macOS firewall blocks Python, allow incoming connections in System Settings → Network → Firewall.
- **Server not opening** – Use the launcher log to identify crashes; you can also run `python run_app.py` inside `.venv` manually for debugging.

All processing stays on-device; no data leaves your computer.
