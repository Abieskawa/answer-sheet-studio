#!/bin/bash
cd "$(dirname "$0")"

<<<<<<< HEAD
pick_python() {
  local candidate
  for candidate in python3.10 python3.11 python3.12 python3.13 python3; do
=======
OUT_DIR="$(pwd)/outputs"
mkdir -p "$OUT_DIR"
START_LOG="${OUT_DIR}/start_mac.log"

log() {
  printf "[%s] %s\n" "$(date '+%Y-%m-%d %H:%M:%S')" "$*" >> "$START_LOG"
}

# Finder launches do not inherit the shell PATH.
export PATH="/opt/homebrew/bin:/opt/homebrew/sbin:/usr/local/bin:/usr/local/sbin:/opt/local/bin:/Library/Frameworks/Python.framework/Versions/3.12/bin:/Library/Frameworks/Python.framework/Versions/3.11/bin:/Library/Frameworks/Python.framework/Versions/3.10/bin:$PATH"

is_port_open() {
  nc -z -w 1 127.0.0.1 "$1" >/dev/null 2>&1
}

if is_port_open 8000; then
  log "App already running on port 8000. Opening browser."
  open "http://127.0.0.1:8000" >/dev/null 2>&1 || true
  exit 0
fi

pick_python() {
  local candidate
  for candidate in python3.11 python3.12 python3.10 python3.13 python3; do
>>>>>>> 82684a2 (Make start_mac.command executable)
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
<<<<<<< HEAD
=======
  log "Python 3.10+ not found. PATH=$PATH"
>>>>>>> 82684a2 (Make start_mac.command executable)
  RECOMMENDED_PYTHON_VERSION="${ANSWER_SHEET_PYTHON_VERSION:-3.11.8}"
  CHOICE="$(osascript -e 'button returned of (display dialog "Python 3.10+ was not found.\n\nDownload and open the Python installer now? (Recommended: Python 3.11)\n\nSource: python.org" buttons {"Cancel","Download"} default button "Download" with icon caution)' 2>/dev/null || true)"

  if [ "$CHOICE" = "Download" ]; then
    PKG_URL="https://www.python.org/ftp/python/${RECOMMENDED_PYTHON_VERSION}/python-${RECOMMENDED_PYTHON_VERSION}-macos11.pkg"
    PKG_PATH="${HOME}/Downloads/answer_sheet_studio_python_${RECOMMENDED_PYTHON_VERSION}.pkg"
    if curl -L --fail --output "$PKG_PATH" "$PKG_URL" >/dev/null 2>&1; then
      if pkgutil --check-signature "$PKG_PATH" 2>/dev/null | grep -q "Python Software Foundation"; then
        open "$PKG_PATH" >/dev/null 2>&1 || open "$PKG_URL" >/dev/null 2>&1 || true
        osascript -e 'display dialog "Python installer opened.\n\nAfter installation finishes, run Answer Sheet Studio again." buttons {"OK"} with icon note' >/dev/null 2>&1 || true
      else
        osascript -e 'display dialog "Downloaded installer signature could not be verified.\n\nWe will open python.org instead." buttons {"OK"} with icon stop' >/dev/null 2>&1 || true
        open "https://www.python.org/downloads/" >/dev/null 2>&1 || true
      fi
    else
      osascript -e 'display dialog "Failed to download the Python installer.\n\nWe will open python.org instead." buttons {"OK"} with icon stop' >/dev/null 2>&1 || true
      open "https://www.python.org/downloads/" >/dev/null 2>&1 || true
    fi
  else
    open "https://www.python.org/downloads/" >/dev/null 2>&1 || true
  fi
  exit 1
fi

<<<<<<< HEAD
ANSWER_SHEET_OPEN_BROWSER=0 nohup "$PYTHON_BIN" launcher_headless.py >/dev/null 2>&1 &
=======
log "Using Python: $PYTHON_BIN"
url_path="${OUT_DIR}/progress_url.txt"
rm -f "$url_path" >/dev/null 2>&1 || true

ANSWER_SHEET_OPEN_BROWSER=0 nohup "$PYTHON_BIN" launcher_headless.py >> "$START_LOG" 2>&1 &
log "Launcher PID: $!"
>>>>>>> 82684a2 (Make start_mac.command executable)

disown

# Open the local progress page written by launcher_headless.py (like Windows does).
if [ "${ANSWER_SHEET_OPEN_BROWSER:-1}" != "0" ]; then
  ( # subshell so we don't block Finder
<<<<<<< HEAD
    out_dir="$(pwd)/outputs"
    url_path="${out_dir}/progress_url.txt"
    i=0
    while [ $i -lt 50 ]; do # ~5s
=======
    opened="0"
    i=0
    while [ $i -lt 200 ]; do # ~20s
>>>>>>> 82684a2 (Make start_mac.command executable)
      if [ -f "$url_path" ]; then
        u="$(cat "$url_path" 2>/dev/null || true)"
        if [ -n "$u" ]; then
          open "$u" >/dev/null 2>&1 || true
<<<<<<< HEAD
=======
          log "Opened progress URL: $u"
          opened="1"
>>>>>>> 82684a2 (Make start_mac.command executable)
          break
        fi
      fi
      sleep 0.1
      i=$((i+1))
    done
<<<<<<< HEAD
  ) >/dev/null 2>&1 &
fi

if [ "${ANSWER_SHEET_CLOSE_TERMINAL:-1}" != "0" ]; then
  osascript -e 'tell application "Terminal" to close (first window whose frontmost is true) saving no' >/dev/null 2>&1 || true
fi

=======
    if [ "$opened" = "0" ]; then
      if is_port_open 8000; then
        open "http://127.0.0.1:8000" >/dev/null 2>&1 || true
        log "Progress URL not found; opened app URL instead."
        opened="1"
      else
        log "Progress URL not found after 20s. See outputs/launcher.log."
      fi
    fi
    if [ "${ANSWER_SHEET_CLOSE_TERMINAL:-1}" != "0" ] && [ "$opened" = "1" ]; then
      osascript -e 'tell application "Terminal" to close (first window whose frontmost is true) saving no' >/dev/null 2>&1 || true
    fi
  ) >/dev/null 2>&1 &
fi

>>>>>>> 82684a2 (Make start_mac.command executable)
exit 0
