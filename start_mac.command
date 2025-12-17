#!/bin/bash
cd "$(dirname "$0")"
export TK_SILENCE_DEPRECATION=1
python3 launcher_gui.py
