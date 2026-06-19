@echo off
REM ===== SOA TOC/TOD tool - one-click launcher for Windows =====
REM Double-click this file. It installs what is needed and opens the app.

cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 (
  echo.
  echo Python is not installed. Please install Python 3.9+ from https://www.python.org/downloads/
  echo IMPORTANT: on the first install screen, tick "Add Python to PATH".
  echo.
  pause
  exit /b 1
)

echo Installing required libraries (first run only)...
python -m pip install --quiet --disable-pip-version-check flask pdfplumber openpyxl

echo Starting the SOA tool...
start "" http://127.0.0.1:5000
python app.py

pause
