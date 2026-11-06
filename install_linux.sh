#!/usr/bin/env bash
set -euo pipefail
cd "$(dirname "$0")"
PYTHON_BIN="${PYTHON_BIN:-python3}"
if ! command -v "$PYTHON_BIN" >/dev/null 2>&1; then
  echo "Python 3 was not found. Ask IT to install Python 3.10+ or set PYTHON_BIN." >&2
  exit 1
fi
"$PYTHON_BIN" -m venv .venv
.venv/bin/python -m pip install -r requirements.txt
echo "Dependencies installed. Review README_FA.md, then run ./run_linux.sh"
