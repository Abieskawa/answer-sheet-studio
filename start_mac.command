#!/bin/bash
cd "$(dirname "$0")"
export TK_SILENCE_DEPRECATION=1

if command -v python3.10 >/dev/null 2>&1; then
  python3.10 launcher_headless.py
elif command -v python3.11 >/dev/null 2>&1; then
  python3.11 launcher_headless.py
elif command -v python3.12 >/dev/null 2>&1; then
  python3.12 launcher_headless.py
elif command -v python3.13 >/dev/null 2>&1; then
  python3.13 launcher_headless.py
else
  python3 launcher_headless.py
fi
