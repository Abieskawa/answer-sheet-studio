# Answer Sheet Studio

## 繁體中文

點此下載程式（ZIP）：https://github.com/Abieskawa/answer-sheet-studio/archive/refs/heads/main.zip  
GitHub Releases（推薦）：https://github.com/Abieskawa/answer-sheet-studio/releases/latest  
線上文件（Read the Docs）：https://answer-sheet-studio.readthedocs.io/zh-tw/latest/

Answer Sheet Studio 讓老師可以產生可列印的答案卡，並在本機進行影像辨識（不會上傳到雲端）：

- **下載答案卡**：輸入標題、科目、題數（最多 100 題）、每題選項（`ABC` / `ABCD` / `ABCDE`），下載答案卡 PDF（給學生填答）。
- **下載老師答案檔**：下載 `answer_key.xlsx`（Excel），填入 `correct/points` 之後到「上傳處理」上傳。
- **上傳處理（辨識＋分析）**：上傳多頁 PDF（每頁一位學生）＋老師答案檔 `answer_key.xlsx`。完成後會提供「開啟結果頁」按鈕（可用新分頁開啟），並輸出 `results.csv`（題號為列、學生為欄）、`ambiguity.csv`（空白/模稜兩可/多選）、`annotated.pdf`（標註辨識結果）與分析報表/圖表（圖表會依介面語言顯示）。
- **完整分析報告**：在圖表頁點擊「列印／存成 PDF」，即可一鍵下載包含所有圖表、成績分佈及詳細試題分析（含正確率/鑑別度）的完整報告。
- **啟動器**：雙擊啟動器（`start_mac.command` 或 `start_windows.vbs`）即可建立虛擬環境、安裝套件、啟動伺服器並開啟 `http://127.0.0.1:8000`。
  - Windows 會把虛擬環境放在 `%LOCALAPPDATA%\\AnswerSheetStudio\\venvs\\<requirements-hash>`，即使重新下載/解壓縮專案也能重用，避免每次都重新安裝。
  - macOS 會把虛擬環境放在 `~/Library/Application Support/AnswerSheetStudio/venvs/<requirements-hash>`，即使重新下載/解壓縮專案也能重用，避免每次都重新安裝。
  - 如需自訂位置，可設定環境變數 `ANSWER_SHEET_VENV_DIR`（指向 venv 根目錄）。
- 網頁介面支援 **English / 繁體中文**，可用頁首的語言切換。

### 系統需求

- macOS 13+ 或 Windows 11
- Python 3.10+（建議 3.11；支援 3.10–3.13）。若尚未安裝 Python，啟動器可協助下載官方安裝程式。
- Windows 安裝 Python 時請勾選「Add python.exe to PATH」（並保留 `py` launcher）
- 第一次安裝需要網路下載 Python 套件（FastAPI、PyMuPDF、OpenCV、NumPy 等）

### 快速開始

macOS
1. 雙擊 `start_mac.command`（若提示權限，先執行一次 `chmod +x start_mac.command`）。若尚未安裝 Python，會提示下載並開啟安裝程式。
2. 第一次會建立虛擬環境並安裝依賴套件；之後會重用既有環境（除非 `requirements.txt` 有變更）。
3. 瀏覽器會自動開啟 `http://127.0.0.1:8000`。用完關閉瀏覽器即可；伺服器會在一段時間無操作後自動結束。若之後看到 `ERR_CONNECTION_REFUSED` /「127.0.0.1 拒絕連線」，請再雙擊啟動器重新啟動。

Windows 11
1. 雙擊 `start_windows.vbs`。若尚未安裝 Python，會提示自動下載/安裝 Python 3.11（建議）。
   - 若要做「一直按下一步」的安裝精靈（Setup.exe），見 `installer/windows/README.md`。
2. 若你選擇手動安裝 Python，請勾選「Add python.exe to PATH」（並保留 `py` launcher）。
3. 第一次會安裝依賴套件；之後會重用既有環境（除非 `requirements.txt` 有變更）。若 Windows Defender 詢問是否允許網路連線，請允許（只會綁定 localhost）。

### 使用注意事項

- 進行辨識時，**題數** 與 **每題選項（ABC/ABCD/ABCDE）** 必須與答案卡一致。
- 想要結果更穩定：建議用較深的筆、掃描 **300dpi**，並避免歪斜/旋轉。

### 輸出檔案

辨識完成後，檔案會寫入 `outputs/<job_id>/`：

- `results.csv`（學生欄位預設為「年級-班級-座號」（班級支援 1–9；0 保留），例如 `8-1-01`；重複會自動加 `_2` / `_3`）
- `ambiguity.csv`
- `roster.csv`（從答案卡讀出的年級/班級/座號與頁碼）
- `annotated.pdf`
- `input.pdf`（原始上傳檔）
- `answer_key.xlsx`（老師答案檔）
- `showwrong.xlsx`（只顯示錯題：題號為列、學生為欄；最後一列為每位學生總分）
- `analysis_scores.csv`、`analysis_item.csv`、`analysis_summary.csv`（`analysis_template.csv` 預設隱藏）
- `analysis_scores_by_class.xlsx`（按班級分表的成績表）
- `analysis_score_hist.png`、`analysis_item_plot.png`
- `analysis_showwrong.json`（結果頁互動圖表用）
- `analysis_report.pdf`（分析結果 PDF；每班一頁）
- 若未來新增更多 `analysis_*` 輸出（CSV/XLSX/圖片/日誌），結果頁的「試題分析檔案」會自動列出下載連結。

### 圖表頁（/charts）

- `/result/<job_id>/charts` 會把互動式表格/圖表與「試題分析檔案」集中在同一頁。
- 互動功能：滑鼠移到表格/圖表可醒目提示題號，點一下可鎖定；支援學生篩選/排序、隱藏正確，以及鍵盤快捷鍵（←/→ 題號、↑/↓ 學生、Esc 清除、PgUp/PgDn 切換圖）。
- 可用頁面上的「列印／存成 PDF」按鈕，透過瀏覽器列印功能輸出成可列印的 PDF（包含所有圖表與下方詳細數據表格）。
- 「試題分析數據（預覽）」為可展開的表格（預設收合）；需要完整資料請下載 `analysis_item.csv`。

### 更新

- 開啟 `http://127.0.0.1:8000/update`，上傳最新 ZIP（GitHub Releases 或 `main.zip`）。更新後會自動重新啟動。

### Debug Mode（回報問題用）

一般使用者不需要使用 Debug Mode。

- 開啟 `http://127.0.0.1:8000/debug`，輸入 Job ID（`outputs/` 底下的資料夾名稱）。
- 下載 `results.csv`、`ambiguity.csv`、`annotated.pdf`（必要時再下載 `input.pdf`），提供給開發者協助排查。

### 疑難排解

- **pip install 失敗** – 檢查 launcher log；確認有網路、Python 版本為 3.10+。若你是 Python 3.14，建議改用 Python 3.10–3.13。
- **Port 已被佔用 / permission denied** – 啟動器會檢查 port；若 macOS 防火牆阻擋 Python，請到 System Settings > Network > Firewall 放行。
- **伺服器沒有自動開啟** – 先看 launcher log 找出錯誤；也可以在終端機執行 `python run_app.py`（需先在同一環境安裝 requirements）方便除錯。

### 文件（Sphinx / Read the Docs）

本機建置文件：

macOS / Linux
1. `python -m venv .venv && source .venv/bin/activate`
2. `pip install -r docs/requirements.txt`
3. `make html`

Windows（PowerShell）
1. `py -3.11 -m venv .venv`
2. `.\.venv\Scripts\Activate.ps1`
3. `pip install -r docs/requirements.txt`
4. `.\make.bat html`

Sphinx 原始檔在 `source/`；Read the Docs 使用相同設定（`source/conf.py`），並依 RTD 語系建置 `en` / `zh-tw`。

### E2E Demo（跑完整流程）

會自動產生「假老師答案檔」並把所有輸出寫到 `outputs/<job_id>/`：

- 有範例掃描檔時：`python scripts/e2e_demo.py --input test/八年級期末掃描.pdf`
- 沒有掃描檔時：`python scripts/e2e_demo.py --synthetic-pages 8`
- 依照程式輸出提示開啟 `http://127.0.0.1:8000/result/<job_id>/charts`（若伺服器已啟動）

---

## English

Download (ZIP): https://github.com/Abieskawa/answer-sheet-studio/archive/refs/heads/main.zip  
GitHub Releases (recommended): https://github.com/Abieskawa/answer-sheet-studio/releases/latest  
Docs (Read the Docs): https://answer-sheet-studio.readthedocs.io/en/latest/

Answer Sheet Studio lets teachers generate printable answer sheets and run local OMR recognition (no cloud upload):

- **Download page** – choose title, subject, number of questions (up to 100), and choices per question (`ABC` / `ABCD` / `ABCDE`). Download the answer sheet PDF (for students).
- **Teacher answer key** – download `answer_key.xlsx` (Excel), fill `correct/points`, then upload it on the Upload page.
- **Upload page (recognize + analyze)** – upload a multi-page PDF scan (one student per page) plus `answer_key.xlsx`. After processing, you’ll get an **Open result page** button (open in a new tab if desired). Exports `results.csv` (questions as rows, students as columns), `ambiguity.csv` (blank/ambiguous/multi picks), `annotated.pdf`, and analysis reports/plots (plots follow the selected UI language).
- **Launcher** – double-click (`start_mac.command` or `start_windows.vbs`) to create a virtual env, install dependencies, start the local server, and open `http://127.0.0.1:8000`.
- The launcher reuses a stable venv across re-downloads/unzips:
  - Windows: `%LOCALAPPDATA%\\AnswerSheetStudio\\venvs\\<requirements-hash>`
  - macOS: `~/Library/Application Support/AnswerSheetStudio/venvs/<requirements-hash>`
  - To override, set `ANSWER_SHEET_VENV_DIR` (points to the venv root dir).
- The web UI supports **English** and **Traditional Chinese**. Use the language tabs in the header.

### Requirements

- macOS ≥ 13 or Windows 11
- Python **3.10+** (3.11 recommended; 3.10–3.13 supported). If Python isn’t installed yet, the launcher can help download the official installer.
- On Windows, ensure “Add python.exe to PATH” during installation (and keep the `py` launcher enabled if offered).
- Internet access the first time to download Python packages (FastAPI, PyMuPDF, OpenCV, NumPy, etc.).

### Quick Start

macOS
1. Double-click `start_mac.command` (or run `chmod +x start_mac.command` once if prompted). If Python isn’t installed, it will offer to download and open the installer.
2. First run creates a virtual env and installs requirements; later runs reuse the existing env (unless requirements changed).
3. Your browser opens `http://127.0.0.1:8000`. Close the browser when you’re done; the server auto-exits after a period of inactivity. If you later see `ERR_CONNECTION_REFUSED` / “127.0.0.1 refused to connect”, rerun the launcher to start the local server again.

Windows 11
1. Double-click `start_windows.vbs`. If Python isn’t installed, it will offer to download/install Python 3.11 automatically (recommended).
   - For a traditional “Next/Next/Finish” installer wizard (Setup.exe), see `installer/windows/README.md`.
2. If you install Python manually, ensure “Add python.exe to PATH” during installation (and keep the `py` launcher enabled if offered).
3. First run installs requirements; later runs reuse the existing environment (unless requirements changed). If Windows Defender prompts for network access, allow it so the server can bind to localhost.

### Important Notes

- For recognition, the **number of questions** and **choices per question** must match the generated answer sheet.
- For more stable results: use a darker pen/pencil, scan at **300dpi**, and avoid skew/rotation.

### Output Files

After recognition, files are written under `outputs/<job_id>/`:

- `results.csv` (student columns default to `grade-class-seat` (class 1–9; 0 reserved), e.g. `8-1-01`; duplicates get `_2` / `_3`)
- `ambiguity.csv`
- `roster.csv` (grade/class/seat and page index extracted from the sheet)
- `annotated.pdf`
- `input.pdf` (original upload)
- `answer_key.xlsx` (teacher answer key)
- `showwrong.xlsx` (wrong answers only; questions as rows, students as columns; last row is total score per student)
- `analysis_scores.csv`, `analysis_item.csv`, `analysis_summary.csv` (`analysis_template.csv` is hidden by default)
- `analysis_scores_by_class.xlsx` (scores split by class)
- `analysis_score_hist.png`, `analysis_item_plot.png`
- `analysis_showwrong.json` (interactive report data)
- `analysis_report.pdf` (analysis report PDF; one page per class)
- If more `analysis_*` outputs are added later (CSV/XLSX/images/logs), the result page’s “Item analysis files” section will auto-list them.

### Charts Page (/charts)

- `/result/<job_id>/charts` puts plots and “Item analysis files” on one page.
- Interactions: hover the table/charts to highlight a question, click to lock; filter/sort students, hide correct, and use keyboard shortcuts (←/→ question, ↑/↓ student, Esc clear, PgUp/PgDn switch chart).
- Use the “Print / Save as PDF” button to export a printable PDF (includes all charts and the detailed data table).
- “Item analysis data (preview)” is a collapsible table (collapsed by default); download `analysis_item.csv` for full data.

### Updating

- Open `http://127.0.0.1:8000/update` and upload the latest ZIP (from GitHub Releases or `main.zip`). The app will restart automatically.

### Debug Mode

- Open `http://127.0.0.1:8000/debug` and enter the Job ID (the folder name under `outputs/`) to download diagnostic files (including `ambiguity.csv`).

### Troubleshooting

- **pip install failed** – Check the launcher log; ensure network access and Python 3.10+. If you’re on Python 3.14, try Python 3.10–3.13.
- **Port already in use / permission denied** – The launcher checks port availability. If macOS firewall blocks Python, allow incoming connections in System Settings > Network > Firewall.
- **Server not opening** – Use the launcher log to identify crashes; you can also run `python run_app.py` after installing requirements in the same environment for debugging.

### Documentation (Sphinx / Read the Docs)

Build the docs locally:

macOS / Linux
1. `python -m venv .venv && source .venv/bin/activate` (or your preferred venv)
2. `pip install -r docs/requirements.txt`
3. `make html`

Windows (PowerShell)
1. `py -3.11 -m venv .venv`
2. `.\.venv\Scripts\Activate.ps1`
3. `pip install -r docs/requirements.txt`
4. `.\make.bat html`

The Sphinx source lives in `source/`. Read the Docs uses the same config (`source/conf.py`) and builds `en` / `zh-tw` based on the RTD language setting.

### E2E Demo (run the full pipeline)

Creates a fake teacher answer key and writes all outputs under `outputs/<job_id>/`:

- With a sample scan: `python scripts/e2e_demo.py --input test/八年級期末掃描.pdf`
- Without a scan: `python scripts/e2e_demo.py --synthetic-pages 8`
- Open `http://127.0.0.1:8000/result/<job_id>/charts` (if the server is running)
