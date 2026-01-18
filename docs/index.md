# Answer Sheet Studio

Answer Sheet Studio helps teachers generate printable answer sheets and run local OMR recognition (no cloud upload).

## Download

- GitHub Releases (recommended): https://github.com/Abieskawa/answer-sheet-studio/releases/latest
- ZIP (main branch): https://github.com/Abieskawa/answer-sheet-studio/archive/refs/heads/main.zip

## Quick Start

1. Double-click `start_mac.command` (macOS) or `start_windows.vbs` (Windows).
2. Your browser opens `http://127.0.0.1:8000`.
3. Close the browser when finished; the server auto-exits after a period of inactivity.

## Analysis (recommended)

Use the Download page to download the answer-sheet PDF and the teacher answer key template (`answer_key.xlsx`). Fill the answer key, then upload **both**:

- multi-page scanned PDF (one student per page)
- teacher answer key (`answer_key.xlsx`)

The Upload page runs recognition + analysis and lets you download tables and plots. For best-looking plots, install R (`Rscript`) and packages: `readr`, `dplyr`, `tidyr`, `ggplot2` (the launcher auto-opens the installer if R is missing).

## Docs / README

- README（繁體中文在前 / English below）: https://github.com/Abieskawa/answer-sheet-studio/blob/main/README.md
