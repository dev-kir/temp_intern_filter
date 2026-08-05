#!/bin/bash
# Super Filter — double-click launcher for macOS.
# First run creates a local venv and installs the two dependencies. Later runs just start.
cd "$(dirname "$0")" || exit 1

if [ ! -x "venv/bin/python" ]; then
  echo "First run: setting up. This takes about a minute."
  python3 -m venv venv || { echo "Python 3 is not installed. Get it from python.org"; read -r; exit 1; }
  ./venv/bin/python -m pip install --quiet --upgrade pip
  ./venv/bin/python -m pip install --quiet openpyxl tkinterdnd2
fi

./venv/bin/python app.py
