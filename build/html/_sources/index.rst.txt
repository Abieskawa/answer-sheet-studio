Answer Sheet Studio
===================

English | 繁體中文: ``README.zh-Hant.md``

Answer Sheet Studio lets teachers generate printable answer sheets and run local OMR recognition:

- **Download page** – choose title, subject, class, number of questions (up to 100), and choices per question (ABC/ABCD/ABCDE). Generates a single-page PDF template (macOS/Windows compatible).
- **Upload page** – drop a multi-page PDF scan. Recognition exports ``results.csv`` (questions as rows, students as columns), ``ambiguity.csv`` (blank/ambiguous picks), plus ``annotated.pdf`` that visualizes detections.
- **Launcher** – double-click (``start_mac.command`` or ``start_windows.vbs``) to create a virtual env, install dependencies, start the local server, and open ``http://127.0.0.1:8000``.
- The web UI supports **English** and **Traditional Chinese**. Use the language tabs in the header.

Requirements
------------

- macOS ≥ 13 or Windows 11
- Python **3.10+** (3.10–3.13 recommended). If you’re on Python 3.14 and installs fail, use Python 3.10–3.13.
- On Windows, ensure “Add python.exe to PATH” during installation (and keep the ``py`` launcher enabled if offered).
- Internet access the first time to download Python packages (FastAPI, PyMuPDF, OpenCV, NumPy, etc.).

Quick Start
-----------

macOS
~~~~~

#. Double-click ``start_mac.command`` (or run ``chmod +x start_mac.command`` once if prompted).
#. First run creates ``.venv`` and installs requirements; later runs reuse the existing ``.venv`` (unless requirements changed).
#. Your browser opens ``http://127.0.0.1:8000``. Close the browser when you’re done; the server auto-exits after a period of inactivity.

Windows 11
~~~~~~~~~~

#. Install Python 3.10+ (3.10–3.13 recommended) from python.org and check “Add python.exe to PATH”.
#. Double-click ``start_windows.vbs``.
#. First run installs requirements; later runs reuse the existing ``.venv`` (unless requirements changed). If Windows Defender prompts for network access, allow it so the server can bind to localhost.

Updating (No Git Required)
--------------------------

- Open ``http://127.0.0.1:8000/update`` and upload the latest ZIP (from GitHub Releases or ``main.zip``). The app will restart automatically.
- If you changed any project files locally, updating may overwrite your changes. Keep backups of customized files.

Debug Mode
----------

- Open ``http://127.0.0.1:8000/debug`` and enter the Job ID (the folder name under ``outputs/``) to download diagnostic files (including ``ambiguity.csv``).

Repository Layout
-----------------

.. code-block:: text

   app/            FastAPI app (routes, templates, static)
   omr/            Generator + recognition pipeline (ReportLab + OpenCV/PyMuPDF)
   launcher_headless.py Headless launcher (no GUI)
   update_worker.py Update helper (used by /update)
   start_mac.command / start_windows.vbs  OS-specific launchers

Generated files are written under ``outputs/``.

Customization
-------------

- Modify ``omr/generator.py`` to adjust layout (title, subject label, footer, coordinates).
- Update recognition thresholds or geometry in ``omr/recognizer.py`` / ``omr/config.py``.
- Front-end templates live in ``app/templates/``.

Troubleshooting
---------------

- **pip install failed** – Check the launcher log; ensure network access and Python 3.10+. If you’re on Python 3.14, try Python 3.10–3.13.
- **Port already in use / permission denied** – The launcher checks port availability. If macOS firewall blocks Python, allow incoming connections in System Settings > Network > Firewall.
- **Server not opening** – Use the launcher log to identify crashes; you can also run ``python run_app.py`` inside ``.venv`` manually for debugging.

Packaging Notes (for double-click to run)
-----------------------------------------

If you want a distribution that does not require teachers to install Python/Git:

- **Single** ``.exe`` **(recommended UX)**: build with PyInstaller or Nuitka and ship updates by replacing the ``.exe``.
- **Portable folder**: bundle Python embeddable + the repo + a ``.bat`` launcher (unzip and run).
- **Installer**: wrap the portable build with Inno Setup or NSIS (desktop shortcuts, uninstall, etc.).

All processing stays on-device; no data leaves your computer.
