#!/bin/bash
# Double-click this file to install Invoice Converter on this Mac.
# It installs what the app needs, builds it, puts it in Applications and opens it.
cd "$(dirname "$0")"
set -e
APP="Invoice Converter.app"

say_step() { echo; echo "==> $1"; }
fail() {
  echo; echo "!! Installation stopped: $1"
  echo "   Copy everything in this window and send it for help."
  read -n 1 -s -r -p "Press any key to close..."; exit 1
}
trap 'fail "a step above failed"' ERR

load_brew() {
  for b in /opt/homebrew/bin/brew /usr/local/bin/brew; do
    if [ -x "$b" ]; then eval "$("$b" shellenv)"; return 0; fi
  done
  return 1
}

echo "Invoice Converter installer"
echo "This takes about 5-10 minutes. Keep this window open."

say_step "1/5 Checking Homebrew"
if ! command -v brew >/dev/null 2>&1 && ! load_brew; then
  echo "Homebrew isn't installed. Installing it now - it will ask for this Mac's login password"
  echo "(nothing shows while you type the password, that's normal)."
  /bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"
  load_brew || fail "Homebrew didn't install."
fi
echo "Homebrew OK"

say_step "2/5 Installing Tesseract (reads the scans)"
brew list tesseract >/dev/null 2>&1 || brew install tesseract
echo "Tesseract OK"

say_step "3/5 Installing Python 3.12 (only used to build the app)"
brew list python@3.12 >/dev/null 2>&1 || brew install python@3.12
PY="$(brew --prefix python@3.12)/bin/python3.12"
[ -x "$PY" ] || fail "Python 3.12 not found at $PY"
echo "Python OK"

say_step "4/5 Building the app"
bash build_app.sh "$PY"
[ -d "dist/$APP" ] || fail "the app wasn't built."

say_step "5/5 Installing into Applications"
DEST="/Applications"
rm -rf "$DEST/$APP" 2>/dev/null || true
if ! cp -R "dist/$APP" "$DEST/" 2>/dev/null; then
  DEST="$HOME/Applications"; mkdir -p "$DEST"
  rm -rf "$DEST/$APP"; cp -R "dist/$APP" "$DEST/"
fi
xattr -cr "$DEST/$APP" 2>/dev/null || true
echo "Installed: $DEST/$APP"

trap - ERR
echo
echo "All done. Opening Invoice Converter..."
open "$DEST/$APP"
echo "From now on, open it from Applications (or Spotlight: Cmd+Space, type Invoice)."
read -n 1 -s -r -p "Press any key to close this window..."
