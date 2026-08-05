@echo off
REM Super Filter - double-click launcher for Windows.
REM First run creates a local venv and installs the two dependencies. Later runs just start.
cd /d "%~dp0"

if not exist "venv\Scripts\python.exe" (
  echo First run: setting up. This takes about a minute.
  py -3 -m venv venv || python -m venv venv
  if not exist "venv\Scripts\python.exe" (
    echo Python 3 is not installed. Get it from python.org and tick "Add to PATH".
    pause
    exit /b 1
  )
  venv\Scripts\python.exe -m pip install --quiet --upgrade pip
  venv\Scripts\python.exe -m pip install --quiet openpyxl tkinterdnd2
)

start "" venv\Scripts\pythonw.exe app.py
