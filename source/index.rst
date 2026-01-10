Answer Sheet Studio
===================

:ref:`English <english>` | :ref:`繁體中文 <zh-hant>`

.. _english:

English
-------

Answer Sheet Studio lets teachers generate printable answer sheets and run local OMR recognition:

- **Download page** – choose title, subject, class, number of questions (up to 100), and choices per question (ABC/ABCD/ABCDE). Generates a single-page PDF template (macOS/Windows compatible).
- **Upload page** – drop a multi-page PDF scan. Recognition exports ``results.csv`` (questions as rows, students as columns), ``ambiguity.csv`` (blank/ambiguous picks), plus ``annotated.pdf`` that visualizes detections.
- **Launcher** – double-click (``start_mac.command`` or ``start_windows.vbs``) to create a virtual env, install dependencies, start the local server, and open ``http://127.0.0.1:8000``.
- The web UI supports **English** and **Traditional Chinese**. Use the language tabs in the header.

Requirements
~~~~~~~~~~~~

- macOS ≥ 13 or Windows 11
- Python **3.10+** (3.10–3.13 recommended). If you’re on Python 3.14 and installs fail, use Python 3.10–3.13.
- On Windows, ensure “Add python.exe to PATH” during installation (and keep the ``py`` launcher enabled if offered).
- Internet access the first time to download Python packages (FastAPI, PyMuPDF, OpenCV, NumPy, etc.).

Quick Start
~~~~~~~~~~~

macOS
^^^^^

#. Double-click ``start_mac.command`` (or run ``chmod +x start_mac.command`` once if prompted).
#. First run creates ``.venv`` and installs requirements; later runs reuse the existing ``.venv`` (unless requirements changed).
#. Your browser opens ``http://127.0.0.1:8000``. Close the browser when you’re done; the server auto-exits after a period of inactivity.

Windows 11
^^^^^^^^^^

#. Install Python 3.10+ (3.10–3.13 recommended) from python.org and check “Add python.exe to PATH”.
#. Double-click ``start_windows.vbs``.
#. First run installs requirements; later runs reuse the existing ``.venv`` (unless requirements changed). If Windows Defender prompts for network access, allow it so the server can bind to localhost.

Updating (No Git Required)
~~~~~~~~~~~~~~~~~~~~~~~~~~

- Open ``http://127.0.0.1:8000/update`` and upload the latest ZIP (from GitHub Releases or ``main.zip``). The app will restart automatically.
- If you changed any project files locally, updating may overwrite your changes. Keep backups of customized files.

Debug Mode
~~~~~~~~~~

- Open ``http://127.0.0.1:8000/debug`` and enter the Job ID (the folder name under ``outputs/``) to download diagnostic files (including ``ambiguity.csv``).

Repository Layout
~~~~~~~~~~~~~~~~~

.. code-block:: text

   app/            FastAPI app (routes, templates, static)
   omr/            Generator + recognition pipeline (ReportLab + OpenCV/PyMuPDF)
   launcher_headless.py Headless launcher (no GUI)
   update_worker.py Update helper (used by /update)
   start_mac.command / start_windows.vbs  OS-specific launchers

Generated files are written under ``outputs/``.

Customization
~~~~~~~~~~~~~

- Modify ``omr/generator.py`` to adjust layout (title, subject label, footer, coordinates).
- Update recognition thresholds or geometry in ``omr/recognizer.py`` / ``omr/config.py``.
- Front-end templates live in ``app/templates/``.

Troubleshooting
~~~~~~~~~~~~~~~

- **pip install failed** – Check the launcher log; ensure network access and Python 3.10+. If you’re on Python 3.14, try Python 3.10–3.13.
- **Port already in use / permission denied** – The launcher checks port availability. If macOS firewall blocks Python, allow incoming connections in System Settings > Network > Firewall.
- **Server not opening** – Use the launcher log to identify crashes; you can also run ``python run_app.py`` inside ``.venv`` manually for debugging.

Packaging Notes (for double-click to run)
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

If you want a distribution that does not require teachers to install Python/Git:

- **Single** ``.exe`` **(recommended UX)**: build with PyInstaller or Nuitka and ship updates by replacing the ``.exe``.
- **Portable folder**: bundle Python embeddable + the repo + a ``.bat`` launcher (unzip and run).
- **Installer**: wrap the portable build with Inno Setup or NSIS (desktop shortcuts, uninstall, etc.).

All processing stays on-device; no data leaves your computer.

.. _zh-hant:

繁體中文
--------

Answer Sheet Studio 讓老師可以產生可列印的答案卡，並在本機進行 OMR 辨識：

- **下載答案卡**：輸入標題、科目、班級、題數（最多 100 題）、每題選項（ABC / ABCD / ABCDE），產生單頁 PDF。
- **上傳辨識**：上傳多頁 PDF（每頁一位學生）。會輸出 ``results.csv`` （題號為列、學生為欄）、``ambiguity.csv`` （空白/模稜兩可/多選）、以及 ``annotated.pdf`` （標註辨識結果）。
- **Launcher**：雙擊啟動器（``start_mac.command`` 或 ``start_windows.vbs``）即可建立虛擬環境、安裝套件、啟動伺服器並開啟 ``http://127.0.0.1:8000``。
- 網頁介面支援 **English / 繁體中文**，可用頁首的語言切換。

系統需求
~~~~~~~~

- macOS 13+ 或 Windows 11
- Python 3.10+（建議 3.10–3.13）
- Windows 安裝 Python 時請勾選「Add python.exe to PATH」（並保留 ``py`` launcher）
- 第一次安裝需要網路下載 Python 套件（FastAPI、PyMuPDF、OpenCV、NumPy 等）

快速開始
~~~~~~~~

macOS
^^^^^

#. 雙擊 ``start_mac.command`` （若提示權限，先執行一次 ``chmod +x start_mac.command``）。
#. 第一次會建立 ``.venv`` 並安裝依賴套件；之後會重用既有 ``.venv`` （除非 ``requirements.txt`` 有變更）。
#. 瀏覽器會自動開啟 ``http://127.0.0.1:8000``。用完關閉瀏覽器即可；伺服器會在一段時間無操作後自動結束。

Windows 11
^^^^^^^^^^

#. 從 python.org 安裝 Python 3.10+（建議 3.10–3.13），並勾選「Add python.exe to PATH」。
#. 雙擊 ``start_windows.vbs``。
#. 第一次會安裝依賴套件；之後會重用既有 ``.venv`` （除非 ``requirements.txt`` 有變更）。若 Windows Defender 詢問是否允許網路連線，請允許（只會綁定 localhost）。

更新（不需要 Git）
~~~~~~~~~~~~~~~~~~

- 開啟 ``http://127.0.0.1:8000/update``，上傳最新 ZIP（GitHub Releases 或 ``main.zip``）。更新後會自動重新啟動。
- 若你本機有改過專案檔案，更新可能會覆蓋你的修改，建議先備份。

Debug Mode（回報問題用）
~~~~~~~~~~~~~~~~~~~~~~~~

一般使用者不需要使用 Debug Mode。

- 開啟 ``http://127.0.0.1:8000/debug``，輸入 Job ID（``outputs/`` 底下的資料夾名稱）。
- 下載 ``results.csv``、``ambiguity.csv``、``annotated.pdf`` （必要時再下載 ``input.pdf``），提供給開發者協助排查。

專案結構
~~~~~~~~

.. code-block:: text

   app/            FastAPI app（路由、模板、靜態檔）
   omr/            產生與辨識流程（ReportLab + OpenCV/PyMuPDF）
   launcher_headless.py 無介面啟動器（不會跳出視窗）
   update_worker.py 更新輔助程式（/update 會用到）
   start_mac.command / start_windows.vbs  OS 啟動器

產生的檔案會寫入 ``outputs/``。

自訂
~~~~

- 修改 ``omr/generator.py`` 可調整版面（標題/科目/頁尾/座標）。
- 辨識相關的門檻或幾何參數可在 ``omr/recognizer.py`` / ``omr/config.py`` 調整。
- 前端模板在 ``app/templates/``。

排除問題
~~~~~~~~

- **pip install failed** – 檢查 launcher log，確認有網路、Python 版本為 3.10+；若你是 Python 3.14，建議改用 3.10–3.13。
- **Port already in use / permission denied** – 啟動器會檢查 port；若 macOS 防火牆阻擋 Python，請到 System Settings > Network > Firewall 放行。
- **Server not opening** – 先看 launcher log 找出錯誤；也可以在 ``.venv`` 裡直接跑 ``python run_app.py`` 方便除錯。

打包建議（給「非技術使用者」）
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

如果希望達到「下載後點兩下就能跑」，不需要安裝 Python/Git，可考慮：

- **單一執行檔** ``.exe``：用 PyInstaller 或 Nuitka 打包；更新時直接提供新版 ``.exe`` 覆蓋即可。
- **免安裝資料夾（Portable）**：把 Python embeddable + 專案放同一資料夾，用 ``.bat`` 啟動。
- **安裝精靈**：用 Inno Setup 或 NSIS 包成安裝程式，並建立桌面捷徑。

所有處理都在本機完成，不會上傳資料到雲端。
