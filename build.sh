#!/bin/bash
# Build a standalone macOS .app bundle for Super Filter.
# The friend can run this on their own Mac to get the .app.
#
# Requirements:
#   - Python 3.8+
#   - pip install pyinstaller openpyxl pandas tkinterdnd2

cd "$(dirname "$0")"

echo "Installing/updating dependencies..."
pip install -q -r requirements.txt
pip install -q pyinstaller

echo "Building Super Filter.app..."
rm -rf "dist/Super Filter" "build" "Super Filter.spec" 2>/dev/null

pyinstaller --windowed --onedir --name "Super Filter" \
    --icon icon.png \
    --add-data "report_template.xlsx:." \
    --add-data "icon.svg:." \
    --hidden-import=tkinterdnd2 \
    app.py

echo "Done!"
echo "The .app is at: dist/Super Filter.app"
echo "Double-click it to run — no terminal, no Python install needed."
