#!/bin/bash
# BioLegend CoA Lookup — One-click Setup & Run
# Double-click this file in Finder to set up and launch.
# A .venv is created inside this folder on first run; subsequent
# runs skip straight to launching the server.

set -euo pipefail
cd "$(dirname "$0")"          # always run from the folder this script lives in

echo ""
echo "🧬  BioLegend CoA Lookup"
echo "────────────────────────────────────────"

# ── 1. Find Python 3.9+ ───────────────────────────────────────────────────────
PYTHON=""
for cmd in python3 python3.13 python3.12 python3.11 python3.10 python3.9 python; do
  if command -v "$cmd" &>/dev/null; then
    # Check it is Python 3.9+
    OK=$("$cmd" -c \
      'import sys; print("yes" if sys.version_info >= (3,9) else "no")' 2>/dev/null || echo "no")
    if [ "$OK" = "yes" ]; then
      PYTHON="$cmd"
      echo "✓  Python: $("$cmd" --version)"
      break
    fi
  fi
done

if [ -z "$PYTHON" ]; then
  echo ""
  echo "❌  Python 3.9 or newer not found."
  echo ""
  echo "    Install with Homebrew:"
  echo "      /bin/bash -c \"\$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)\""
  echo "      brew install python"
  echo ""
  read -rp "Press Enter to exit…"
  exit 1
fi

# ── 2. Create or reuse .venv ──────────────────────────────────────────────────
VENV_DIR=".venv"

if [ ! -f "$VENV_DIR/bin/activate" ]; then
  echo ""
  echo "📦  Creating virtual environment in .venv …"
  "$PYTHON" -m venv "$VENV_DIR"
  echo "✓  Virtual environment created."
fi

# Activate the venv for the rest of this script
# shellcheck disable=SC1091
source "$VENV_DIR/bin/activate"
echo "✓  Virtual environment active: $(python --version)"

# ── 3. Install / upgrade packages ────────────────────────────────────────────
echo ""
echo "📦  Checking / installing packages…"
pip install --quiet --upgrade pip
pip install --quiet flask requests cloudscraper beautifulsoup4 lxml openpyxl pypdf
echo "✓  All packages ready."

# ── 4. Check for existing process on port 5050 ───────────────────────────────
PORT=5050
EXISTING_PID=$(lsof -ti tcp:$PORT 2>/dev/null || true)

if [ -n "$EXISTING_PID" ]; then
  echo ""
  echo "⚠️   Port $PORT is already in use (PID $EXISTING_PID)."
  echo ""
  read -rp "    Kill the existing process and restart? [y/N] " CONFIRM
  case "$CONFIRM" in
    [yY][eE][sS]|[yY])
      kill "$EXISTING_PID" 2>/dev/null \
        && echo "✓  Killed PID $EXISTING_PID." \
        || echo "⚠️  Could not kill PID $EXISTING_PID — it may have already stopped."
      sleep 0.5
      ;;
    *)
      echo ""
      echo "    Leaving existing process running. Opening browser…"
      open "http://localhost:$PORT"
      echo ""
      read -rp "Press Enter to close…"
      exit 0
      ;;
  esac
fi

# ── 5. Launch the server ─────────────────────────────────────────────────────
echo ""
echo "🚀  Starting server → http://localhost:$PORT"
echo "    (Close this Terminal window to stop the server)"
echo ""

# Open the browser after a short delay to let Flask start up
(sleep 1.5 && open "http://localhost:$PORT") &

python app.py

echo ""
read -rp "Server stopped. Press Enter to close…"
