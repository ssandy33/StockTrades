#!/usr/bin/env bash
set -euo pipefail

# Run the Wheel Strategy Dashboard end-to-end
# - Creates a virtual environment (.venv) if missing
# - Installs dependencies from requirements.txt
# - Executes wheel_strategy_dashboard.py

# Configurable via env vars:
#   VENV_DIR (default: .venv)
#   PYTHON_BIN (default: python3)

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

VENV_DIR="${VENV_DIR:-.venv}"
PYTHON_BIN="${PYTHON_BIN:-python3}"

echo "[1/4] Ensuring virtual environment at: $VENV_DIR"
if [[ ! -d "$VENV_DIR" ]]; then
  "$PYTHON_BIN" -m venv "$VENV_DIR"
fi

echo "[2/4] Activating virtual environment"
source "$VENV_DIR/bin/activate"

echo "[3/4] Installing dependencies"
python -m pip install -U pip
if [[ -f requirements.txt ]]; then
  pip install -r requirements.txt
else
  pip install pandas matplotlib XlsxWriter
fi

echo "[4/4] Running wheel_strategy_dashboard.py"
python wheel_strategy_dashboard.py

echo "Done."

