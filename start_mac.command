#!/bin/bash
cd "$(dirname "$0")"

if command -v python3.10 >/dev/null 2>&1; then
  nohup python3.10 launcher_headless.py >/dev/null 2>&1 &
elif command -v python3.11 >/dev/null 2>&1; then
  nohup python3.11 launcher_headless.py >/dev/null 2>&1 &
elif command -v python3.12 >/dev/null 2>&1; then
  nohup python3.12 launcher_headless.py >/dev/null 2>&1 &
elif command -v python3.13 >/dev/null 2>&1; then
  nohup python3.13 launcher_headless.py >/dev/null 2>&1 &
else
  nohup python3 launcher_headless.py >/dev/null 2>&1 &
fi

disown

if [ "${ANSWER_SHEET_CLOSE_TERMINAL:-1}" != "0" ]; then
  osascript -e 'tell application "Terminal" to close (first window whose frontmost is true) saving no' >/dev/null 2>&1 || true
fi

exit 0
