#!/bin/bash
cd "$(dirname "$0")"

pick_python() {
  local candidate
  for candidate in python3.10 python3.11 python3.12 python3.13 python3; do
    if command -v "$candidate" >/dev/null 2>&1; then
      if "$candidate" -c 'import sys; raise SystemExit(0 if sys.version_info >= (3, 10) else 1)' >/dev/null 2>&1; then
        echo "$candidate"
        return 0
      fi
    fi
  done
  return 1
}

PYTHON_BIN="$(pick_python || true)"
if [ -z "$PYTHON_BIN" ]; then
  osascript -e 'display dialog "Python 3.10+ was not found.\n\nPlease install Python 3.11 (recommended) from python.org, then run Answer Sheet Studio again." buttons {"OK"} with icon stop' >/dev/null 2>&1 || true
  open "https://www.python.org/downloads/" >/dev/null 2>&1 || true
  exit 1
fi

nohup "$PYTHON_BIN" launcher_headless.py >/dev/null 2>&1 &

disown

if [ "${ANSWER_SHEET_CLOSE_TERMINAL:-1}" != "0" ]; then
  osascript -e 'tell application "Terminal" to close (first window whose frontmost is true) saving no' >/dev/null 2>&1 || true
fi

exit 0
