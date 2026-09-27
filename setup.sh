#!/usr/bin/env bash
# MyFinance - one-command setup (macOS / Linux)
# Installs the Python packages, then runs the first-time wizard:
#   ./setup.sh                                  (interactive)
#   ./setup.sh --yes --name "You" --passcode x --income 10000 --no-repo
set -e
cd "$(dirname "$0")"
echo "=== MyFinance setup ==="

if command -v python3 >/dev/null 2>&1; then
  PY=python3
elif command -v python >/dev/null 2>&1; then
  PY=python
else
  echo "[FAIL] Python 3 not found."
  echo "       Install: brew install python   (or https://www.python.org/downloads/)"
  exit 1
fi
echo "Python found: $($PY --version 2>&1)"

echo "Installing packages: openpyxl python-docx plotly"
if ! $PY -m pip install --quiet openpyxl python-docx plotly; then
  # PEP 668 (Homebrew/system Pythons block direct installs) - retry for this user
  if ! $PY -m pip install --quiet --user --break-system-packages openpyxl python-docx plotly; then
    # last resort: isolated virtualenv
    echo "pip blocked -> creating .venv"
    $PY -m venv .venv
    PY=".venv/bin/python"
    $PY -m pip install --quiet openpyxl python-docx plotly
  fi
fi

echo "Running first-time setup (config, sample SMS, build, verify)..."
$PY init_project.py "$@"

echo ""
echo "Setup done."
echo "Next step: install OpenCode from https://opencode.ai then run: opencode"
echo "OpenCode loads AGENTS.md + the myfinance skill automatically,"
echo "and system files are read-only for the agent (permission rules)."
