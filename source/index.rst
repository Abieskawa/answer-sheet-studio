Answer Sheet Studio
===================

更新日期：|today_zh|

Last updated: |today_en|

語言 / Language：:ref:`繁體中文 <zh-hant>` | :ref:`简体中文 <zh-hans>` | :ref:`English <english>`

.. _zh-hant:

繁體中文
--------

`點此下載程式（ZIP） <https://github.com/Abieskawa/answer-sheet-studio/archive/refs/heads/main.zip>`_

Answer Sheet Studio 讓老師可以產生可列印的答案卡，並在本機進行影像辨識（不會上傳到雲端）。

最簡單開始（推薦）
~~~~~~~~~~~~~~~~~~

#. 下載 ZIP 並解壓縮。
#. macOS：雙擊 ``start_mac.command``；Windows：雙擊 ``start_windows.vbs``。
#. 若尚未安裝 Python，啟動器會提示下載/安裝官方 Python（依照畫面一路按確認即可）。
#. 安裝完成後再執行一次啟動器；瀏覽器會自動開啟 ``http://127.0.0.1:8000``。
#. 在網頁裡先下載並列印「答案卡 PDF」（給學生填答）；需要批改/分析時再下載 ``answer_key.xlsx``（老師答案檔）填入 ``correct/points``，到「上傳處理」上傳掃描 PDF + ``answer_key.xlsx``。

系統需求
~~~~~~~~

- macOS 13+ 或 Windows 11
- Python 3.10+（建議 3.11；支援 3.10–3.13）。若尚未安裝 Python，啟動器可協助下載官方安裝程式。
- Windows 安裝 Python 時請勾選「Add python.exe to PATH」（並保留 ``py`` launcher）
- 第一次安裝需要網路下載 Python 套件（FastAPI、PyMuPDF、OpenCV、NumPy 等）

功能說明（進階）
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

- **下載答案卡**：輸入標題、科目、題數（最多 100 題）、每題選項（ABC / ABCD / ABCDE），下載答案卡 PDF（給學生填答）。
- **下載老師答案檔**：下載 ``answer_key.xlsx``（Excel），填入 ``correct/points`` 後到「上傳處理」上傳。
- **上傳處理（辨識＋分析）**：上傳多頁 PDF（每頁一位學生）＋老師答案檔 ``answer_key.xlsx``。完成後會提供「開啟結果頁」按鈕（可用新分頁開啟），並輸出 ``results.csv``、``ambiguity.csv``、``annotated.pdf``，以及分析報表/圖表（圖表會依介面語言顯示；需要安裝 R 用 ``ggplot2`` 出圖）。
- **啟動器**：雙擊啟動器（``start_mac.command`` 或 ``start_windows.vbs``）即可建立虛擬環境、安裝套件、啟動伺服器並開啟 ``http://127.0.0.1:8000``。
  - Windows 會把虛擬環境放在 ``%LOCALAPPDATA%\\AnswerSheetStudio\\venvs\\<requirements-hash>``，即使重新下載/解壓縮專案也能重用，避免每次都重新安裝。
  - macOS 會把虛擬環境放在 ``~/Library/Application Support/AnswerSheetStudio/venvs/<requirements-hash>``，即使重新下載/解壓縮專案也能重用，避免每次都重新安裝。
  - 如需自訂位置，可設定環境變數 ``ANSWER_SHEET_VENV_DIR``（指向 venv 根目錄）。

使用注意事項
~~~~~~~~~~~~

- 進行影像辨識時，**題數** 與 **每題選項（ABC/ABCD/ABCDE）** 必須與答案卡一致。
- 想要結果更穩定：建議用較深的筆、掃描 **300dpi**，並避免歪斜/旋轉。
- 需要安裝 **R**（可執行 ``Rscript``）以產生試題分析報表與圖表（``ggplot2``）。
- R 套件需求：``readr``、``dplyr``、``tidyr``、``ggplot2``。

輸出檔案
~~~~~~~~

辨識完成後，檔案會寫入 ``outputs/<job_id>/``：

- ``results.csv``
- ``ambiguity.csv``
- ``annotated.pdf``
- ``input.pdf`` （原始上傳檔）
- ``answer_key.xlsx``（老師答案檔）
- ``showwrong.xlsx``（只顯示錯題：題號為列、學生為欄；最後一列為每位學生總分）
- ``analysis_template.csv``、``analysis_scores.csv``、``analysis_item.csv``、``analysis_summary.csv``
- ``analysis_score_hist.png``、``analysis_item_plot.png``
- ``analysis_showwrong.json``（結果頁互動圖表用）
- ``analysis_report.pdf``（分析結果 PDF；每班一頁）

圖表頁（/charts）
~~~~~~~~~~~~~~~~~

- 開啟 ``/result/<job_id>/charts`` 可在同一頁查看互動式表格/圖表與「試題分析檔案」下載連結。
- 互動功能：滑鼠移到表格/圖表可醒目提示題號，點一下可鎖定；支援學生篩選/排序、隱藏正確，以及鍵盤快捷鍵（←/→ 題號、↑/↓ 學生、Esc 清除、PgUp/PgDn 切換圖）。
- 可用頁面上的「列印／存成 PDF」按鈕，透過瀏覽器列印功能輸出成可列印的 PDF。
- 「試題分析數據（預覽）」為可展開表格（預設收合）；需要完整資料請下載 ``analysis_item.csv``。
- 若未來新增更多 ``analysis_*`` 輸出（CSV/XLSX/圖片/日誌），結果頁的「試題分析檔案」會自動列出下載連結。

更新
~~~~

- 開啟 ``http://127.0.0.1:8000/update``，上傳最新 ZIP（GitHub Releases 或 ``main.zip``）。更新後會自動重新啟動。

Debug Mode（回報問題用）
~~~~~~~~~~~~~~~~~~~~~~~~

一般使用者不需要使用 Debug Mode。

- 開啟 ``http://127.0.0.1:8000/debug``，輸入 Job ID（``outputs/`` 底下的資料夾名稱）。
- 下載 ``results.csv``、``ambiguity.csv``、``annotated.pdf``（必要時再下載 ``input.pdf``），提供給開發者協助排查。

排除問題
~~~~~~~~

- 若看到 ``ERR_CONNECTION_REFUSED`` /「127.0.0.1 拒絕連線」，請再雙擊啟動器（``start_mac.command`` / ``start_windows.vbs``）重新啟動。
- **pip install failed** – 檢查 launcher log，確認有網路、Python 版本為 3.10+；若你是 Python 3.14，建議改用 3.10–3.13。
- **Port already in use / permission denied** – 啟動器會檢查 port；若 macOS 防火牆阻擋 Python，請到 System Settings > Network > Firewall 放行。
- **Server not opening** – 先看 launcher log 找出錯誤；也可以在終端機執行 ``python run_app.py``（需先在同一環境安裝 requirements）方便除錯。

.. _english:

English
-------

`Download (ZIP) <https://github.com/Abieskawa/answer-sheet-studio/archive/refs/heads/main.zip>`_

Answer Sheet Studio lets teachers generate printable answer sheets and run local OMR recognition (no cloud upload).

Getting Started (recommended)
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#. Download and unzip the ZIP.
#. macOS: double-click ``start_mac.command``; Windows: double-click ``start_windows.vbs``.
#. If Python isn't installed yet, the launcher will guide you to download/install the official Python installer.
#. After installation finishes, run the launcher again; your browser opens ``http://127.0.0.1:8000``.
#. In the web UI: download/print the answer sheet PDF first; when grading, download ``answer_key.xlsx`` (teacher answer key), fill ``correct/points``, then upload the scanned PDF + ``answer_key.xlsx`` on the Upload page.

Requirements
~~~~~~~~~~~~

- macOS ≥ 13 or Windows 11
- Python **3.10+** (3.11 recommended; 3.10–3.13 supported). If Python isn’t installed yet, the launcher can help download the official installer.
- On Windows, ensure “Add python.exe to PATH” during installation (and keep the ``py`` launcher enabled if offered).
- Internet access the first time to download Python packages (FastAPI, PyMuPDF, OpenCV, NumPy, etc.).

Quick Start
~~~~~~~~~~~

macOS
^^^^^

#. Double-click ``start_mac.command`` (or run ``chmod +x start_mac.command`` once if prompted).
#. If Python isn’t installed, the launcher offers to download and open the installer (from python.org). After installation finishes, run ``start_mac.command`` again.
#. First run creates a virtual env and installs requirements; later runs reuse the existing env (unless requirements changed).
#. Your browser opens ``http://127.0.0.1:8000``. Close the browser when you’re done; the server auto-exits after a period of inactivity.

Windows 11
^^^^^^^^^^

#. Double-click ``start_windows.vbs``. If Python isn’t installed, it will offer to download/install Python 3.11 automatically (recommended).
   - For a traditional “Next/Next/Finish” installer wizard (Setup.exe), see ``installer/windows/README.md``.
#. If you install Python manually, check “Add python.exe to PATH” during installation (and keep the ``py`` launcher enabled if offered).
#. First run installs requirements; later runs reuse the existing environment (unless requirements changed). If Windows Defender prompts for network access, allow it so the server can bind to localhost.

Features
~~~~~~~~

- **Download page** – choose title, subject, number of questions (up to 100), and choices per question (ABC/ABCD/ABCDE). Download the answer sheet PDF (for students).
- **Teacher answer key** – download ``answer_key.xlsx`` (Excel), fill ``correct/points``, then upload it on the Upload page.
- **Upload page (recognize + analyze)** – upload a multi-page PDF scan (one student per page) plus ``answer_key.xlsx``. After processing, you’ll get an **Open result page** button (open in a new tab if desired). Exports ``results.csv``, ``ambiguity.csv``, ``annotated.pdf``, and analysis reports/plots (plots follow the selected UI language; R is required for ggplot2 plots).
- **Launcher** – sets up the virtual env, installs requirements, starts the local server, and opens the browser.
  - Windows reuses a stable venv at ``%LOCALAPPDATA%\\AnswerSheetStudio\\venvs\\<requirements-hash>``.
  - macOS reuses a stable venv at ``~/Library/Application Support/AnswerSheetStudio/venvs/<requirements-hash>``.
  - To override, set ``ANSWER_SHEET_VENV_DIR`` (points to the venv root dir).
- The web UI supports **English** and **Traditional Chinese**.

Important Notes
~~~~~~~~~~~~~~~

- For recognition, the **number of questions** and **choices per question** must match the generated answer sheet.
- For more stable results: use a darker pen/pencil, scan at **300dpi**, and avoid skew/rotation.
- Install **R** (``Rscript``) to generate item analysis reports and plots (via ``ggplot2``).
- R packages: ``readr``, ``dplyr``, ``tidyr``, ``ggplot2``.

Output Files
~~~~~~~~~~~~

After recognition, files are written under ``outputs/<job_id>/``:

- ``results.csv``
- ``ambiguity.csv``
- ``annotated.pdf``
- ``input.pdf`` (original upload)
- ``answer_key.xlsx`` (teacher answer key)
- ``showwrong.xlsx`` (wrong answers only; questions as rows, students as columns; last row is total score per student)
- ``analysis_template.csv``, ``analysis_scores.csv``, ``analysis_item.csv``, ``analysis_summary.csv``
- ``analysis_score_hist.png``, ``analysis_item_plot.png``
- ``analysis_showwrong.json`` (interactive report data)
- ``analysis_report.pdf`` (analysis report PDF; one page per class)

Charts Page (/charts)
~~~~~~~~~~~~~~~~~~~~~

- Open ``/result/<job_id>/charts`` to view the interactive table/charts and item-analysis downloads on one page.
- Interactions: hover the table/charts to highlight a question, click to lock; filter/sort students, hide correct, and use keyboard shortcuts (←/→ question, ↑/↓ student, Esc clear, PgUp/PgDn switch chart).
- Use the “Print / Save as PDF” button to export a printable PDF via your browser’s print dialog.
- “Item analysis data (preview)” is a collapsible table (collapsed by default); download ``analysis_item.csv`` for full data.
- If more ``analysis_*`` outputs are added later (CSV/XLSX/images/logs), the “Item analysis files” section auto-lists them.

Updating
~~~~~~~~

- Open ``http://127.0.0.1:8000/update`` and upload the latest ZIP (from GitHub Releases or ``main.zip``). The app will restart automatically.

Debug Mode
~~~~~~~~~~

- Open ``http://127.0.0.1:8000/debug`` and enter the Job ID to download diagnostic files.

Troubleshooting
~~~~~~~~~~~~~~~

- **ERR_CONNECTION_REFUSED / “127.0.0.1 refused to connect”** – Rerun the launcher (`start_mac.command` / `start_windows.vbs`) to start the local server again.
- **pip install failed** – Check the launcher log; ensure network access and Python 3.10+. If you’re on Python 3.14, try Python 3.10–3.13.
- **Port already in use / permission denied** – The launcher checks port availability.
- **Server not opening** – Use the launcher log or run ``python run_app.py`` after installing requirements in the same environment for debugging.

.. _zh-hans:

简体中文
--------

`点此下载程序（ZIP） <https://github.com/Abieskawa/answer-sheet-studio/archive/refs/heads/main.zip>`_

Answer Sheet Studio 让老师可以生成可打印的答案卡，并在本机进行图像识别（不会上传到云端）。

最简单开始（推荐）
~~~~~~~~~~~~~~~~~~

#. 下载 ZIP 并解压缩。
#. macOS：双击 ``start_mac.command``；Windows：双击 ``start_windows.vbs``。
#. 若尚未安装 Python，启动器会提示下载/安装官方 Python（依照画面一路确认即可）。
#. 安装完成后再执行一次启动器；浏览器会自动打开 ``http://127.0.0.1:8000``。用完关闭浏览器即可；若之后看到 ``ERR_CONNECTION_REFUSED`` /「127.0.0.1 拒绝连接」，请再双击启动器重新启动。
#. 在网页里先下载并打印「答案卡 PDF」（给学生填答）；需要批改/分析时再下载 ``answer_key.xlsx``（老师答案档）填写 ``correct/points``，到「上传处理」上传扫描 PDF + ``answer_key.xlsx``。

系统需求
~~~~~~~~

- macOS 13+ 或 Windows 11
- Python 3.10+（建议 3.11；支持 3.10–3.13）。若尚未安装 Python，启动器可协助下载官方安装程序。
- Windows 安装 Python 时请勾选「Add python.exe to PATH」（并保留 ``py`` launcher）
- 第一次安装需要网络下载 Python 套件（FastAPI、PyMuPDF、OpenCV、NumPy 等）
- 需要安装 **R**（可执行 ``Rscript``）以生成试题分析报表与图表（``ggplot2``）。
- R 套件需求：``readr``、``dplyr``、``tidyr``、``ggplot2``。

功能说明（进阶）
~~~~~~~~~~~~~~~~

- **下载答案卡**：输入标题、科目、题数（最多 100 题）、每题选项（ABC / ABCD / ABCDE），下载答案卡 PDF（给学生填答）。
- **下载老师答案档**：下载 ``answer_key.xlsx``（Excel），填入 ``correct/points`` 后到「上传处理」上传。
- **上传处理（识别＋分析）**：上传多页 PDF（每页一位学生）＋老师答案档 ``answer_key.xlsx``。完成后会提供「打开结果页」按钮，并输出 ``results.csv``、``ambiguity.csv``、``annotated.pdf``，以及分析报表/图表（图表会依界面语言显示；需要安装 R 用 ``ggplot2`` 出图）。

输出档案
~~~~~~~~

识别完成后，档案会写入 ``outputs/<job_id>/``：

- ``results.csv``
- ``ambiguity.csv``
- ``annotated.pdf``
- ``input.pdf``（原始上传档）
- ``answer_key.xlsx``（老师答案档）
- ``showwrong.xlsx``（只显示错题：题号为列、学生为栏；最后一列为每位学生总分）
- ``analysis_template.csv``、``analysis_scores.csv``、``analysis_item.csv``、``analysis_summary.csv``
- ``analysis_score_hist.png``、``analysis_item_plot.png``
- ``analysis_showwrong.json``（结果页互动图表用）
- ``analysis_report.pdf``（分析结果 PDF；每班一页）

图表页（/charts）
~~~~~~~~~~~~~~~~~

- 打开 ``/result/<job_id>/charts`` 可在同一页查看互动式表格/图表与「试题分析档案」下载链接。
- 可用页面上的「打印／存成 PDF」按钮，通过浏览器打印功能输出成可打印的 PDF。

更新
~~~~

- 打开 ``http://127.0.0.1:8000/update``，上传最新 ZIP（GitHub Releases 或 ``main.zip``）。更新后会自动重新启动。
