#!/usr/bin/env bash
# Automated Setup and Launcher for Linux / macOS
set -e

DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" >/dev/null 2>&1 && pwd )"
cd "$DIR"

echo "==============================================================================="
echo "   BANK STATEMENT MULTI-FORMAT FILTER & INTER-BANK RECONCILIATION ENGINE"
echo "==============================================================================="

if ! command -v python3 &> /dev/null; then
    echo "[ERROR] Python 3 is required but not installed."
    exit 1
fi

if [ ! -d ".venv" ]; then
    echo "[*] Creating virtual environment in .venv..."
    python3 -m venv .venv
fi

echo "[*] Activating virtual environment..."
source .venv/bin/activate

echo "[*] Installing requirements..."
pip install --upgrade pip --quiet
pip install -r requirements.txt --quiet

echo "[*] Launching Streamlit application on http://localhost:8501..."
if command -v xdg-open &> /dev/null; then
    (sleep 2 && xdg-open http://localhost:8501) &
elif command -v open &> /dev/null; then
    (sleep 2 && open http://localhost:8501) &
fi

streamlit run app.py --server.port 8501 --server.headless false
