# Answer Sheet Studio（繁體中文）

Answer Sheet Studio 讓老師可以產生可列印的答案卡，並在本機進行 OMR 辨識：

- **下載答案卡**：輸入標題、科目、班級、題數（最多 100 題），產生單頁 PDF。
- **上傳辨識**：上傳多頁 PDF（每頁一位學生）。會輸出 `results.csv`（題號為列、學生為欄）與 `annotated.pdf`（標註辨識結果）。
- **Launcher**：雙擊啟動器（`start_mac.command` 或 `start_windows.vbs`）即可建立虛擬環境、安裝套件、啟動伺服器並開啟 `http://127.0.0.1:8000`。

網頁介面支援 **English / 繁體中文**，可用頁首的語言切換。

## 系統需求

- macOS 13+ 或 Windows 11
- Python 3.10+（建議 3.10–3.13）
- Windows 安裝 Python 時請勾選「Add python.exe to PATH」（並保留 `py` launcher）
- 第一次安裝需要網路下載 Python 套件（FastAPI、PyMuPDF、OpenCV、NumPy 等）

## 快速開始

### macOS
1. 雙擊 `start_mac.command`（若提示權限，先執行一次 `chmod +x start_mac.command`）。
2. 出現 Launcher 視窗後，按 **Install & Run**。第一次會建立 `.venv` 並安裝依賴套件。
3. 之後用 **Open in Browser** 開啟 `http://127.0.0.1:8000`，用 **Stop Server** 停止服務。

### Windows 11
1. 從 python.org 安裝 Python 3.10+（建議 3.10–3.13），並勾選「Add python.exe to PATH」。
2. 雙擊 `start_windows.vbs` 開啟 Launcher（使用 `pythonw`）。
3. 按 **Install & Run**。若 Windows Defender 詢問是否允許網路連線，請允許（只會綁定 localhost）。

## 更新（不需要 Git）

- 在 Launcher 按 **Download/Update (GitHub)**：會從 GitHub 下載最新 ZIP 並覆蓋更新專案檔案。
- 若學校網路阻擋下載：可用瀏覽器先下載 ZIP，再用 Launcher 的 **Apply ZIP Update...** 套用更新。
- 若你本機有改過專案檔案，更新可能會覆蓋你的修改，建議先備份。

## Debug Mode（回報問題用）

一般使用者不需要使用 Debug Mode。

- 開啟 `http://127.0.0.1:8000/debug`，輸入處理完成頁面顯示的 Job ID。
- 下載 `results.csv`、`ambiguity.csv`、`annotated.pdf`（必要時再下載 `input.pdf`），提供給開發者協助排查。

## 專案結構

```
app/            FastAPI app（路由、模板、靜態檔）
omr/            產生與辨識流程（ReportLab + OpenCV/PyMuPDF）
launcher_gui.py 跨平台 Tk Launcher
start_mac.command / start_windows.vbs  OS 啟動器
```

產生的檔案會寫入 `outputs/`。

## 打包建議（給「非技術使用者」）

如果希望達到「下載後點兩下就能跑」，不需要安裝 Python/Git，可考慮：

- **單一執行檔 `.exe`**：用 PyInstaller 或 Nuitka 打包；更新時直接提供新版 `.exe` 覆蓋即可。
- **免安裝資料夾（Portable）**：把 Python embeddable + 專案放同一資料夾，用 `.bat` 啟動。
- **安裝精靈**：用 Inno Setup 或 NSIS 包成安裝程式，並建立桌面捷徑。

所有處理都在本機完成，不會上傳資料到雲端。
