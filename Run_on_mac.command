#!/bin/bash
# Virtualization Sizing Calculator - macOS launcher
# First run: installs dependencies into a private .venv folder. Every run: launches the app.
cd "$(dirname "$0")" || exit 1
clear
echo "==========================================="
echo "   Virtualization Sizing Calculator"
echo "==========================================="
echo ""

# 1. Apple Command Line Tools (provides python3 on a clean Mac)
if ! xcode-select -p &>/dev/null; then
    echo "[setup] Apple Developer Tools missing. A popup will appear - click 'Install'."
    xcode-select --install
    echo "        Waiting for installation to finish..."
    until xcode-select -p &>/dev/null; do sleep 5; done
fi

# 2. Python
if ! command -v python3 &>/dev/null; then
    echo "Python 3 not found. Install it from https://www.python.org/downloads/ and re-run."
    read -r -p "Press [Enter] to close..."
    exit 1
fi

# 3. Private virtual environment (avoids 'externally-managed-environment' errors)
VENV=".venv"
if [ ! -x "$VENV/bin/python" ]; then
    echo "[setup] Creating Python environment (first run only)..."
    python3 -m venv "$VENV" || { echo "Failed to create virtual environment."; read -r -p "Press [Enter]..."; exit 1; }
fi

# 4. Install / update libraries only when requirements.txt changes
STAMP="$VENV/.requirements.installed"
if [ ! -f "$STAMP" ] || ! cmp -s requirements.txt "$STAMP"; then
    echo "[setup] Installing libraries (Streamlit, Pandas, OpenPyXL)..."
    "$VENV/bin/python" -m pip install --quiet --upgrade pip
    if "$VENV/bin/python" -m pip install --quiet -r requirements.txt; then
        cp requirements.txt "$STAMP"
    else
        echo "Library install failed. Check your network/proxy and re-run."
        read -r -p "Press [Enter] to close..."
        exit 1
    fi
fi

# 5. Launch
echo ""
echo "Launching... your browser will open shortly. Close this window to stop the app."
exec "$VENV/bin/python" -m streamlit run sizing_app.py
