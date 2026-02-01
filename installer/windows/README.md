# Windows 安裝精靈（Inno Setup）

這個資料夾提供一個「一直按下一步」的 Windows 安裝精靈，負責：
- 把整個專案安裝到 `Program Files`
- 建立開始功能表/桌面捷徑
- 安裝完成後可直接啟動（使用 `start_windows.vbs`）

## 需求
- Windows
- Inno Setup 6（安裝後會有 `ISCC.exe`）

## 產出安裝檔
在 repo 根目錄開 PowerShell：

`"%ProgramFiles(x86)%\\Inno Setup 6\\ISCC.exe" installer\\windows\\AnswerSheetStudio.iss`

成功後會在 `installer\\windows\\Output\\AnswerSheetStudio-Setup.exe`（或 Inno 的預設輸出資料夾）生成安裝檔。

## 使用者體驗
- 安裝過程是標準 Windows Wizard（下一步/下一步/完成）
- 完成後會建立捷徑並可直接啟動，啟動會開啟安裝進度頁，完成後自動導向到程式介面

## 備註
- 這個安裝精靈「不內建 Python」：但第一次啟動時會提供自動下載/安裝 Python 3.11（預設 3.11.8，來源 python.org），安裝完成後會自動重試啟動。
- 此專案 **不需要安裝 R**：試題分析報表/圖表由程式內建產生。
- 若要做成完全離線/免 Python 的安裝檔，通常需要把 Python + 依賴一起打包成單一執行檔（例如 PyInstaller）或使用 embeddable Python + wheels，流程會更複雜。
