#!/usr/bin/env bash
# ===== SOA TOC/TOD tool - one-click launcher for macOS / Linux =====
# In Terminal:  bash run.sh   (or make executable: chmod +x run.sh && ./run.sh)
set -e
cd "$(dirname "$0")"

if ! command -v python3 >/dev/null 2>&1; then
  echo "Python 3 is not installed. Install it from https://www.python.org/downloads/"
  exit 1
fi

echo "Installing required libraries (first run only)..."
python3 -m pip install --quiet --disable-pip-version-check flask pdfplumber openpyxl

echo "Starting the SOA tool..."
( sleep 2; (open http://127.0.0.1:5000 2>/dev/null || xdg-open http://127.0.0.1:5000 2>/dev/null) ) &
python3 app.py
