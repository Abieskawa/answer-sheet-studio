# Answer Sheet Studio（繁體中文）

點此下載程式（ZIP）：https://github.com/Abieskawa/answer-sheet-studio/archive/refs/heads/main.zip
線上文件（Read the Docs）：https://answer-sheet-studio.readthedocs.io/zh-tw/latest/

Answer Sheet Studio 讓老師可以產生可列印的答案卡，並在本機進行影像辨識（不會上傳到雲端）：

- **下載答案卡**：輸入標題、科目、班級、題數（最多 100 題）、每題選項（`ABC` / `ABCD` / `ABCDE`），產生單頁 PDF。
- **上傳影像辨識**：上傳多頁 PDF（每頁一位學生）。會輸出 `results.csv`（題號為列、學生為欄）、`ambiguity.csv`（空白/模稜兩可/多選）與 `annotated.pdf`（標註辨識結果）。
- **啟動器**：雙擊啟動器（`start_mac.command` 或 `start_windows.vbs`）即可建立虛擬環境、安裝套件、啟動伺服器並開啟 `http://127.0.0.1:8000`。

網頁介面支援 **English / 繁體中文**，可用頁首的語言切換。

## 系統需求

- macOS 13+ 或 Windows 11
- Python 3.10+（建議 3.11；支援 3.10–3.13）。若尚未安裝 Python，啟動器可協助下載官方安裝程式。
- Windows 安裝 Python 時請勾選「Add python.exe to PATH」（並保留 `py` launcher）
- 第一次安裝需要網路下載 Python 套件（FastAPI、PyMuPDF、OpenCV、NumPy 等）

## 快速開始

### macOS
1. 雙擊 `start_mac.command`（若提示權限，先執行一次 `chmod +x start_mac.command`）。若尚未安裝 Python，會提示下載並開啟安裝程式。
2. 第一次會建立 `.venv` 並安裝依賴套件；之後會重用既有 `.venv`（除非 `requirements.txt` 有變更）。
3. 瀏覽器會自動開啟 `http://127.0.0.1:8000`。用完關閉瀏覽器即可；伺服器會在一段時間無操作後自動結束。

### Windows 11
1. 雙擊 `start_windows.vbs`。若尚未安裝 Python，會提示自動下載/安裝 Python 3.11（建議）。
2. 若你選擇手動安裝 Python，請勾選「Add python.exe to PATH」（並保留 `py` launcher）。
3. 第一次會安裝依賴套件；之後會重用既有 `.venv`（除非 `requirements.txt` 有變更）。若 Windows Defender 詢問是否允許網路連線，請允許（只會綁定 localhost）。

## 使用注意事項

- 進行辨識時，**題數** 與 **每題選項（ABC/ABCD/ABCDE）** 必須與答案卡一致。
- 想要結果更穩定：建議用較深的筆、掃描 **300dpi**，並避免歪斜/旋轉。

## 輸出檔案

辨識完成後，檔案會寫入 `outputs/<job_id>/`：

- `results.csv`
- `ambiguity.csv`
- `annotated.pdf`
- `input.pdf`（原始上傳檔）

## 更新

- 開啟 `http://127.0.0.1:8000/update`，上傳最新 ZIP（GitHub Releases 或 `main.zip`）。更新後會自動重新啟動。

## Debug Mode（回報問題用）

一般使用者不需要使用 Debug Mode。

- 開啟 `http://127.0.0.1:8000/debug`，輸入 Job ID（`outputs/` 底下的資料夾名稱）。
- 下載 `results.csv`、`ambiguity.csv`、`annotated.pdf`（必要時再下載 `input.pdf`），提供給開發者協助排查。
