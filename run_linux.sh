#!/usr/bin/env bash
set -euo pipefail
cd "$(dirname "$0")"
if [[ ! -x .venv/bin/gunicorn ]]; then
  echo "Virtual environment/dependencies not found. Run ./install_linux.sh first." >&2
  exit 1
fi
exec .venv/bin/gunicorn --workers 2 --bind 127.0.0.1:5000 --timeout 60 wsgi:app
