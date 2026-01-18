# Answer Sheet Studio

## 繁體中文

點此下載程式（ZIP）：https://github.com/Abieskawa/answer-sheet-studio/archive/refs/heads/main.zip  
線上文件（Read the Docs）：https://answer-sheet-studio.readthedocs.io/zh-tw/latest/

Answer Sheet Studio 讓老師可以產生可列印的答案卡，並在本機進行影像辨識（不會上傳到雲端）：

- **下載答案卡**：輸入標題、科目、題數（最多 100 題）、每題選項（`ABC` / `ABCD` / `ABCDE`），下載一個 ZIP（答案卡 PDF + 老師答案檔 `answer_key.xlsx`/`answer_key.csv`）。
- **上傳處理（辨識＋分析）**：上傳多頁 PDF（每頁一位學生）＋老師答案檔（`XLSX/CSV`）。會輸出 `results.csv`（題號為列、學生為欄）、`ambiguity.csv`（空白/模稜兩可/多選）、`annotated.pdf`（標註辨識結果）與分析報表/圖表（建議安裝 R 用 `ggplot2` 出圖）。
- **啟動器**：雙擊啟動器（`start_mac.command` 或 `start_windows.vbs`）即可建立虛擬環境、安裝套件、啟動伺服器並開啟 `http://127.0.0.1:8000`。
- 網頁介面支援 **English / 繁體中文**，可用頁首的語言切換。

### 系統需求

- macOS 13+ 或 Windows 11
- Python 3.10+（建議 3.11；支援 3.10–3.13）。若尚未安裝 Python，啟動器可協助下載官方安裝程式。
- Windows 安裝 Python 時請勾選「Add python.exe to PATH」（並保留 `py` launcher）
- 第一次安裝需要網路下載 Python 套件（FastAPI、PyMuPDF、OpenCV、NumPy 等）
- （推薦）安裝 **R**（可執行 `Rscript`）以使用 `ggplot2` 產生更好看的圖表；啟動器會在偵測不到 R 時自動下載並開啟安裝程式。安裝完成後請再執行一次啟動器。
- R 套件需求：`readr`、`dplyr`、`tidyr`、`ggplot2`。

### 快速開始

macOS
1. 雙擊 `start_mac.command`（若提示權限，先執行一次 `chmod +x start_mac.command`）。若尚未安裝 Python，會提示下載並開啟安裝程式。
2. 第一次會建立 `.venv` 並安裝依賴套件；之後會重用既有 `.venv`（除非 `requirements.txt` 有變更）。
3. 瀏覽器會自動開啟 `http://127.0.0.1:8000`。用完關閉瀏覽器即可；伺服器會在一段時間無操作後自動結束。

Windows 11
1. 雙擊 `start_windows.vbs`。若尚未安裝 Python，會提示自動下載/安裝 Python 3.11（建議）。
2. 若你選擇手動安裝 Python，請勾選「Add python.exe to PATH」（並保留 `py` launcher）。
3. 第一次會安裝依賴套件；之後會重用既有 `.venv`（除非 `requirements.txt` 有變更）。若 Windows Defender 詢問是否允許網路連線，請允許（只會綁定 localhost）。

### 使用注意事項

- 進行辨識時，**題數** 與 **每題選項（ABC/ABCD/ABCDE）** 必須與答案卡一致。
- 想要結果更穩定：建議用較深的筆、掃描 **300dpi**，並避免歪斜/旋轉。

### 輸出檔案

辨識完成後，檔案會寫入 `outputs/<job_id>/`：

- `results.csv`
- `ambiguity.csv`
- `annotated.pdf`
- `input.pdf`（原始上傳檔）
- `answer_key.xlsx`、`answer_key.csv`（老師答案檔）
- `analysis_template.csv`、`analysis_scores.csv`、`analysis_item.csv`、`analysis_summary.csv`、`analysis_score_hist.png`、`analysis_item_plot.png`

### 更新

- 開啟 `http://127.0.0.1:8000/update`，上傳最新 ZIP（GitHub Releases 或 `main.zip`）。更新後會自動重新啟動。

### Debug Mode（回報問題用）

一般使用者不需要使用 Debug Mode。

- 開啟 `http://127.0.0.1:8000/debug`，輸入 Job ID（`outputs/` 底下的資料夾名稱）。
- 下載 `results.csv`、`ambiguity.csv`、`annotated.pdf`（必要時再下載 `input.pdf`），提供給開發者協助排查。

### 疑難排解

- **pip install 失敗** – 檢查 launcher log；確認有網路、Python 版本為 3.10+。若你是 Python 3.14，建議改用 Python 3.10–3.13。
- **Port 已被佔用 / permission denied** – 啟動器會檢查 port；若 macOS 防火牆阻擋 Python，請到 System Settings > Network > Firewall 放行。
- **伺服器沒有自動開啟** – 先看 launcher log 找出錯誤；也可以在 `.venv` 裡直接跑 `python run_app.py` 方便除錯。

---

## English

Download (ZIP): https://github.com/Abieskawa/answer-sheet-studio/archive/refs/heads/main.zip  
Docs (Read the Docs): https://answer-sheet-studio.readthedocs.io/zh-tw/latest/

Answer Sheet Studio lets teachers generate printable answer sheets and run local OMR recognition (no cloud upload):

- **Download page** – choose title, subject, number of questions (up to 100), and choices per question (`ABC` / `ABCD` / `ABCDE`). Downloads a ZIP (answer sheet PDF + teacher answer key `answer_key.xlsx`/`answer_key.csv`).
- **Upload page (recognize + analyze)** – upload a multi-page PDF scan (one student per page) plus the teacher answer key (`XLSX/CSV`). Exports `results.csv` (questions as rows, students as columns), `ambiguity.csv` (blank/ambiguous/multi picks), `annotated.pdf`, and analysis reports/plots (install R for ggplot2 plots).
- **Launcher** – double-click (`start_mac.command` or `start_windows.vbs`) to create a virtual env, install dependencies, start the local server, and open `http://127.0.0.1:8000`.
- The web UI supports **English** and **Traditional Chinese**. Use the language tabs in the header.

### Requirements

- macOS ≥ 13 or Windows 11
- Python **3.10+** (3.11 recommended; 3.10–3.13 supported). If Python isn’t installed yet, the launcher can help download the official installer.
- On Windows, ensure “Add python.exe to PATH” during installation (and keep the `py` launcher enabled if offered).
- Internet access the first time to download Python packages (FastAPI, PyMuPDF, OpenCV, NumPy, etc.).
- (Recommended) Install **R** (`Rscript`) for ggplot2 plots. If R is missing, the launcher auto-downloads and opens the official installer; run the launcher again after installing.
- R packages: `readr`, `dplyr`, `tidyr`, `ggplot2`.

### Quick Start

macOS
1. Double-click `start_mac.command` (or run `chmod +x start_mac.command` once if prompted). If Python isn’t installed, it will offer to download and open the installer.
2. First run creates `.venv` and installs requirements; later runs reuse the existing `.venv` (unless requirements changed).
3. Your browser opens `http://127.0.0.1:8000`. Close the browser when you’re done; the server auto-exits after a period of inactivity.

Windows 11
1. Double-click `start_windows.vbs`. If Python isn’t installed, it will offer to download/install Python 3.11 automatically (recommended).
2. If you install Python manually, ensure “Add python.exe to PATH” during installation (and keep the `py` launcher enabled if offered).
3. First run installs requirements; later runs reuse the existing `.venv` (unless requirements changed). If Windows Defender prompts for network access, allow it so the server can bind to localhost.

### Important Notes

- For recognition, the **number of questions** and **choices per question** must match the generated answer sheet.
- For more stable results: use a darker pen/pencil, scan at **300dpi**, and avoid skew/rotation.

### Output Files

After recognition, files are written under `outputs/<job_id>/`:

- `results.csv`
- `ambiguity.csv`
- `annotated.pdf`
- `input.pdf` (original upload)
- `answer_key.xlsx`, `answer_key.csv` (teacher answer key)
- `analysis_template.csv`, `analysis_scores.csv`, `analysis_item.csv`, `analysis_summary.csv`, `analysis_score_hist.png`, `analysis_item_plot.png`

### Updating

- Open `http://127.0.0.1:8000/update` and upload the latest ZIP (from GitHub Releases or `main.zip`). The app will restart automatically.

### Debug Mode

- Open `http://127.0.0.1:8000/debug` and enter the Job ID (the folder name under `outputs/`) to download diagnostic files (including `ambiguity.csv`).

### Troubleshooting

- **pip install failed** – Check the launcher log; ensure network access and Python 3.10+. If you’re on Python 3.14, try Python 3.10–3.13.
- **Port already in use / permission denied** – The launcher checks port availability. If macOS firewall blocks Python, allow incoming connections in System Settings > Network > Firewall.
- **Server not opening** – Use the launcher log to identify crashes; you can also run `python run_app.py` inside `.venv` manually for debugging.
