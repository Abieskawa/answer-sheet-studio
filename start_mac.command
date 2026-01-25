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

ensure_r() {
  if command -v Rscript >/dev/null 2>&1; then
    return 0
  fi

  # R may be installed but Rscript not on PATH (e.g. when launched from Finder).
  local cand
  for cand in \
    "/Library/Frameworks/R.framework/Resources/bin/Rscript" \
    "/usr/local/bin/Rscript" \
    "/opt/homebrew/bin/Rscript"; do
    if [ -x "$cand" ]; then
      return 0
    fi
  done

  local arch base html pkg_name url
  arch="$(uname -m 2>/dev/null || echo "")"
  if [ "$arch" = "arm64" ]; then
    base="https://cran.r-project.org/bin/macosx/big-sur-arm64/base/"
  else
    base="https://cran.r-project.org/bin/macosx/big-sur-x86_64/base/"
  fi

  html="$(curl -fsSL "$base" 2>/dev/null || true)"
  pkg_name="$(printf "%s" "$html" | grep -Eo 'R-[0-9.]+(-arm64)?\\.pkg' | sort -V | tail -n 1)"
  if [ -z "$pkg_name" ]; then
    open "https://cran.r-project.org/bin/macosx/" >/dev/null 2>&1 || true
    osascript -e 'display dialog "R was not found.\n\nWe opened CRAN in your browser.\n\nPlease install R, then run Answer Sheet Studio again." buttons {"OK"} with icon caution' >/dev/null 2>&1 || true
    return 1
  fi

  url="${base}${pkg_name}"
  PKG_PATH="${HOME}/Downloads/answer_sheet_studio_R_${pkg_name}"
  if curl -L --fail --progress-bar --output "$PKG_PATH" "$url" >/dev/null 2>&1; then
    if pkgutil --check-signature "$PKG_PATH" 2>/dev/null | grep -qi "Status: signed"; then
      open "$PKG_PATH" >/dev/null 2>&1 || open "$url" >/dev/null 2>&1 || true
      osascript -e 'display dialog "R installer opened.\n\nAfter installation finishes, run Answer Sheet Studio again." buttons {"OK"} with icon note' >/dev/null 2>&1 || true
      return 1
    else
      osascript -e 'display dialog "Downloaded R installer signature could not be verified.\n\nWe will open CRAN instead.\n\nPlease install R, then run Answer Sheet Studio again." buttons {"OK"} with icon stop' >/dev/null 2>&1 || true
      open "https://cran.r-project.org/bin/macosx/" >/dev/null 2>&1 || true
      return 1
    fi
  else
    osascript -e 'display dialog "Failed to download the R installer.\n\nWe will open CRAN instead.\n\nPlease install R, then run Answer Sheet Studio again." buttons {"OK"} with icon stop' >/dev/null 2>&1 || true
    open "https://cran.r-project.org/bin/macosx/" >/dev/null 2>&1 || true
    return 1
  fi
  return 1
}

PYTHON_BIN="$(pick_python || true)"
if [ -z "$PYTHON_BIN" ]; then
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

ensure_r || exit 1

ANSWER_SHEET_OPEN_BROWSER=0 nohup "$PYTHON_BIN" launcher_headless.py >/dev/null 2>&1 &

disown

# Open the local progress page written by launcher_headless.py (like Windows does).
if [ "${ANSWER_SHEET_OPEN_BROWSER:-1}" != "0" ]; then
  ( # subshell so we don't block Finder
    out_dir="$(pwd)/outputs"
    url_path="${out_dir}/progress_url.txt"
    i=0
    while [ $i -lt 50 ]; do # ~5s
      if [ -f "$url_path" ]; then
        u="$(cat "$url_path" 2>/dev/null || true)"
        if [ -n "$u" ]; then
          open "$u" >/dev/null 2>&1 || true
          break
        fi
      fi
      sleep 0.1
      i=$((i+1))
    done
  ) >/dev/null 2>&1 &
fi

if [ "${ANSWER_SHEET_CLOSE_TERMINAL:-1}" != "0" ]; then
  osascript -e 'tell application "Terminal" to close (first window whose frontmost is true) saving no' >/dev/null 2>&1 || true
fi

exit 0
