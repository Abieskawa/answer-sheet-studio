Answer Sheet Studio
===================

語言 / Language：:ref:`繁體中文 <zh-hant>` | :ref:`English <english>`

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
- **上傳處理（辨識＋分析）**：上傳多頁 PDF（每頁一位學生）＋老師答案檔 ``answer_key.xlsx``。完成後會在**新分頁**開啟結果頁（上傳頁不會被覆蓋），並輸出 ``results.csv``、``ambiguity.csv``、``annotated.pdf``，以及分析報表/圖表（建議安裝 R 用 ``ggplot2`` 出圖）。
- **啟動器**：雙擊啟動器（``start_mac.command`` 或 ``start_windows.vbs``）即可建立虛擬環境、安裝套件、啟動伺服器並開啟 ``http://127.0.0.1:8000``。

使用注意事項
~~~~~~~~~~~~

- 進行影像辨識時，**題數** 與 **每題選項（ABC/ABCD/ABCDE）** 必須與答案卡一致。
- 想要結果更穩定：建議用較深的筆、掃描 **300dpi**，並避免歪斜/旋轉。
- （推薦）安裝 **R**（可執行 ``Rscript``）以使用 ``ggplot2`` 產生更好看的圖表；若偵測不到 R，啟動器會自動下載並開啟安裝程式。安裝完成後請再執行一次啟動器。
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
- ``analysis_template.csv``、``analysis_scores.csv``、``analysis_item.csv``、``analysis_summary.csv``、``analysis_score_hist.png``、``analysis_item_plot.png``

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

- **pip install failed** – 檢查 launcher log，確認有網路、Python 版本為 3.10+；若你是 Python 3.14，建議改用 3.10–3.13。
- **Port already in use / permission denied** – 啟動器會檢查 port；若 macOS 防火牆阻擋 Python，請到 System Settings > Network > Firewall 放行。
- **Server not opening** – 先看 launcher log 找出錯誤；也可以在 ``.venv`` 裡直接跑 ``python run_app.py`` 方便除錯。

.. _english:

English
-------

`Download (ZIP) <https://github.com/Abieskawa/answer-sheet-studio/archive/refs/heads/main.zip>`_

Answer Sheet Studio lets teachers generate printable answer sheets and run local OMR recognition (no cloud upload).

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
#. First run creates ``.venv`` and installs requirements; later runs reuse the existing ``.venv`` (unless requirements changed).
#. Your browser opens ``http://127.0.0.1:8000``. Close the browser when you’re done; the server auto-exits after a period of inactivity.

Windows 11
^^^^^^^^^^

#. Double-click ``start_windows.vbs``. If Python isn’t installed, it will offer to download/install Python 3.11 automatically (recommended).
#. If you install Python manually, check “Add python.exe to PATH” during installation (and keep the ``py`` launcher enabled if offered).
#. First run installs requirements; later runs reuse the existing ``.venv`` (unless requirements changed). If Windows Defender prompts for network access, allow it so the server can bind to localhost.

Features
~~~~~~~~

- **Download page** – choose title, subject, number of questions (up to 100), and choices per question (ABC/ABCD/ABCDE). Download the answer sheet PDF (for students).
- **Teacher answer key** – download ``answer_key.xlsx`` (Excel), fill ``correct/points``, then upload it on the Upload page.
- **Upload page (recognize + analyze)** – upload a multi-page PDF scan (one student per page) plus ``answer_key.xlsx``. The result page opens in a **new tab** (so the upload form stays). Exports ``results.csv``, ``ambiguity.csv``, ``annotated.pdf``, and analysis reports/plots (install R for ggplot2 plots).
- **Launcher** – sets up ``.venv``, installs requirements, starts the local server, and opens the browser.
- The web UI supports **English** and **Traditional Chinese**.

Important Notes
~~~~~~~~~~~~~~~

- For recognition, the **number of questions** and **choices per question** must match the generated answer sheet.
- For more stable results: use a darker pen/pencil, scan at **300dpi**, and avoid skew/rotation.
- (Recommended) Install **R** (``Rscript``) for ggplot2 plots. If R is missing, the launcher auto-opens the official installer; run the launcher again after installing.
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
- ``analysis_template.csv``, ``analysis_scores.csv``, ``analysis_item.csv``, ``analysis_summary.csv``, ``analysis_score_hist.png``, ``analysis_item_plot.png``

Updating
~~~~~~~~

- Open ``http://127.0.0.1:8000/update`` and upload the latest ZIP (from GitHub Releases or ``main.zip``). The app will restart automatically.

Debug Mode
~~~~~~~~~~

- Open ``http://127.0.0.1:8000/debug`` and enter the Job ID to download diagnostic files.

Troubleshooting
~~~~~~~~~~~~~~~

- **pip install failed** – Check the launcher log; ensure network access and Python 3.10+. If you’re on Python 3.14, try Python 3.10–3.13.
- **Port already in use / permission denied** – The launcher checks port availability.
- **Server not opening** – Use the launcher log or run ``python run_app.py`` inside ``.venv`` for debugging.
