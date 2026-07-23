#!/bin/bash
cd "$(dirname "$0")"

pick_python() {
  local candidate
  for candidate in python3.11 python3.10 python3; do
    if command -v "$candidate" >/dev/null 2>&1; then
      if "$candidate" -c 'import sys; raise SystemExit(0 if (3, 10) <= sys.version_info[:2] <= (3, 11) else 1)' >/dev/null 2>&1; then
        echo "$candidate"
        return 0
      fi
    fi
  done
  return 1
}

PYTHON_BIN="$(pick_python || true)"
if [ -z "$PYTHON_BIN" ]; then
  RECOMMENDED_PYTHON_VERSION="${ANSWER_SHEET_PYTHON_VERSION:-3.11.8}"
  PKG_URL="https://www.python.org/ftp/python/${RECOMMENDED_PYTHON_VERSION}/python-${RECOMMENDED_PYTHON_VERSION}-macos11.pkg"
  PKG_PATH="${HOME}/Downloads/answer_sheet_studio_python_${RECOMMENDED_PYTHON_VERSION}.pkg"
  CHOICE="$(osascript -e 'button returned of (display dialog "未偵測到系統安裝 Python 3.10 或 3.11。\n\n是否立即下載並開啟 Python 3.11.8 安裝程式？" buttons {"取消","下載"} default button "下載" with icon caution)' 2>/dev/null || true)"

  if [ "$CHOICE" = "下載" ]; then
    osascript -e 'display notification "正在下載 Python 3.11.8 安裝程式…" with title "Answer Sheet Studio"' >/dev/null 2>&1 || true
    if curl -L --fail --output "$PKG_PATH" "$PKG_URL" >/dev/null 2>&1; then
      PKG_SIZE=$(stat -f%z "$PKG_PATH" 2>/dev/null || echo 0)
      if [ "$PKG_SIZE" -gt 20000000 ]; then
        open "$PKG_PATH" >/dev/null 2>&1 || true
        osascript -e 'display dialog "Python 安裝程式已下載並開啟。\n\n請於安裝完成後，再次執行 Answer Sheet Studio。" buttons {"確定"} with icon note' >/dev/null 2>&1 || true
      else
        open "$PKG_URL" >/dev/null 2>&1 || true
        osascript -e 'display dialog "已開啟瀏覽器直接進行 Python 安裝檔下載。\n\n請於下載並安裝完成後，再次執行 Answer Sheet Studio。" buttons {"確定"} with icon note' >/dev/null 2>&1 || true
      fi
    else
      open "$PKG_URL" >/dev/null 2>&1 || true
      osascript -e 'display dialog "已開啟瀏覽器直接進行 Python 安裝檔下載。\n\n請於下載並安裝完成後，再次執行 Answer Sheet Studio。" buttons {"確定"} with icon note' >/dev/null 2>&1 || true
    fi
  else
    LINK_BTN=$(osascript -e 'button returned of (display dialog "此程式需要安裝 Python 3.11.8 才能執行。\n\n點擊「開啟連結」可直接在瀏覽器下載安裝檔（無需閱讀英文官方網頁）。" buttons {"關閉","開啟連結"} default button "開啟連結" with icon note)' 2>/dev/null || echo "")
    if [ "$LINK_BTN" = "開啟連結" ]; then
      open "$PKG_URL" >/dev/null 2>&1 || true
    fi
  fi
  exit 1
fi

ANSWER_SHEET_OPEN_BROWSER=0 nohup "$PYTHON_BIN" launcher_headless.py >/dev/null 2>&1 &

disown

# Open the local progress page written by launcher_headless.py (like Windows does).
if [ "${ANSWER_SHEET_OPEN_BROWSER:-1}" != "0" ]; then
  ( # subshell so we don't block Finder
    out_dir="$(pwd)/outputs"
    url_path="${out_dir}/progress_url.txt"
    i=0
    found=0
    while [ $i -lt 300 ]; do # ~30s
      if [ -f "$url_path" ]; then
        u="$(cat "$url_path" 2>/dev/null || true)"
        if [ -n "$u" ]; then
          if command -v curl >/dev/null 2>&1; then
            if ! curl -fsS --max-time 1 "${u}status" >/dev/null 2>&1; then
              u=""
            fi
          fi
          if [ -n "$u" ]; then
            open -a "Safari" "$u" >/dev/null 2>&1 || open "$u" >/dev/null 2>&1 || true
            found=1
            break
          fi
        fi
      fi
      sleep 0.1
      i=$((i+1))
    done
    if [ "$found" -eq 0 ]; then
      default_port="${ANSWER_SHEET_PORT:-8000}"
      fallback_url="http://127.0.0.1:${default_port}"
      open -a "Safari" "$fallback_url" >/dev/null 2>&1 || open "$fallback_url" >/dev/null 2>&1 || true
    fi
  ) >/dev/null 2>&1 &
fi

if [ "${ANSWER_SHEET_CLOSE_TERMINAL:-1}" != "0" ]; then
  osascript -e 'tell application "Terminal" to close (first window whose frontmost is true) saving no' >/dev/null 2>&1 || true
fi

exit 0
