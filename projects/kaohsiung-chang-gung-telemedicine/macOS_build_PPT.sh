#!/usr/bin/env bash

set -euo pipefail

# Resolve and enter the directory containing this script. The build never
# depends on the Terminal's current working directory.
SCRIPT_DIR="$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)"
cd "$SCRIPT_DIR"

VENV_DIR="$SCRIPT_DIR/.venv"
PYTHON_BIN="$VENV_DIR/bin/python"
OUTPUT_FILE="$SCRIPT_DIR/output/kaohsiung-chang-gung-telemedicine-v2.pptx"

if ! command -v python3 >/dev/null 2>&1; then
    echo "[ERROR] python3 was not found."
    echo "Install Python 3 from https://www.python.org/downloads/macos/ or Homebrew."
    exit 1
fi

echo "[1/4] Creating or reusing the local virtual environment..."
if [[ ! -x "$PYTHON_BIN" ]]; then
    python3 -m venv "$VENV_DIR"
fi

echo "[2/4] Installing required Python packages..."
"$PYTHON_BIN" -m pip install --disable-pip-version-check -r "$SCRIPT_DIR/requirements.txt"

echo "[3/4] Building the redesigned Clinical Calm V2 presentation..."
"$PYTHON_BIN" "$SCRIPT_DIR/build_deck_v2.py"

echo "[4/4] Validating the V2 presentation..."
"$PYTHON_BIN" "$SCRIPT_DIR/validate_deck_v2.py"

if [[ ! -f "$OUTPUT_FILE" ]]; then
    echo "[ERROR] The expected PPTX was not created:"
    echo "$OUTPUT_FILE"
    exit 1
fi

echo
echo "[SUCCESS] PowerPoint created:"
echo "$OUTPUT_FILE"

