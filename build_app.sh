#!/bin/bash
# Builds "Invoice Converter.app" into the dist folder.
#   bash build_app.sh                 (uses python3)
#   bash build_app.sh /path/to/python (uses that Python)
set -e
cd "$(dirname "$0")"
PY="${1:-python3}"

echo "==> Setting up a build environment with $("$PY" --version)"
rm -rf .build-venv build dist
"$PY" -m venv .build-venv
source .build-venv/bin/activate
pip install --upgrade pip setuptools wheel -q
pip install -q -r requirements.txt "pyinstaller==6.11.1" "pyinstaller-hooks-contrib==2024.10"

echo "==> Building the app (takes a few minutes)"
pyinstaller --noconfirm --windowed --name "Invoice Converter" \
  --osx-bundle-identifier com.mllservices.invoiceconverter \
  --add-data "ui.html:." \
  --add-data "template.xlsx:." \
  --collect-all webview \
  --collect-all pymupdf \
  --hidden-import webview.platforms.cocoa \
  app.py

deactivate

echo "==> Checking the built app starts (self-test)"
if ! "dist/Invoice Converter.app/Contents/MacOS/Invoice Converter" --selftest; then
  echo
  echo "!! The app was built but fails its self-test (see the error above)."
  echo "   It was NOT installed. Copy this whole window and send it for help."
  exit 1
fi
echo
echo "Built: dist/Invoice Converter.app"
