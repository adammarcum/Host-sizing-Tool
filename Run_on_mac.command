#!/bin/bash
# =============================================================================
#  Host Sizer - macOS launcher
#  Double-click this file (or "Host Sizing Calculator.app").
#  If macOS says "permission denied", run it from Terminal instead:
#      bash Run_on_mac.command
#  First run: finds or installs Python, creates a private .venv, installs the
#  libraries. Every run: starts the app and opens your browser.
# =============================================================================
cd "$(dirname "$0")" || exit 1
clear
echo "==========================================="
echo "   Host Sizer  -  Powered by AHEAD"
echo "==========================================="
echo ""

finish() {   # finish <exit code>  - keep the window open so the message can be read
    echo ""
    read -r -p "Press [Enter] to close this window..."
    exit "${1:-1}"
}

py_ok() {    # true when $1 is a working Python 3.9 or newer
    "$1" -c 'import sys; sys.exit(0 if sys.version_info >= (3, 9) else 1)' >/dev/null 2>&1
}

# -----------------------------------------------------------------------------
# 1. Python 3.9+ (prefer a newer python.org / Homebrew install, else Apple's)
# -----------------------------------------------------------------------------
echo "[1/4] Checking for Python..."
PY=""
for c in python3.13 python3.12 python3.11 python3.10 /opt/homebrew/bin/python3 /usr/local/bin/python3; do
    if command -v "$c" >/dev/null 2>&1 && py_ok "$(command -v "$c")"; then
        PY="$(command -v "$c")"
        break
    fi
done

if [ -z "$PY" ]; then
    # Apple's python3 comes with the Command Line Tools. Calling /usr/bin/python3
    # without them only triggers a popup, so install the tools first.
    if ! xcode-select -p >/dev/null 2>&1; then
        echo ""
        echo "      Python is not installed yet. A macOS popup will ask to install the"
        echo "      Command Line Tools - click INSTALL and accept. This takes a few minutes."
        xcode-select --install >/dev/null 2>&1
        echo "      Waiting for the installation to finish..."
        until xcode-select -p >/dev/null 2>&1; do sleep 5; done
        echo "      Command Line Tools installed."
    fi
    if py_ok /usr/bin/python3; then
        PY=/usr/bin/python3
    fi
fi

if [ -z "$PY" ]; then
    echo ""
    echo "      Python 3.9 or newer could not be found."
    echo "      Opening python.org - install the latest Python for macOS, then double-click"
    echo "      this launcher again."
    open "https://www.python.org/downloads/macos/"
    finish 1
fi
echo "      Using $("$PY" -c 'import sys; print("Python %d.%d.%d" % sys.version_info[:3])') ($PY)"

# -----------------------------------------------------------------------------
# 2. Private environment for this tool (rebuilt automatically if it breaks)
# -----------------------------------------------------------------------------
VENV=".venv"
if [ -x "$VENV/bin/python" ] && ! "$VENV/bin/python" -c 'import sys' >/dev/null 2>&1; then
    echo "[2/4] The tool's Python environment is damaged - rebuilding it..."
    rm -rf "$VENV"
fi
if [ ! -x "$VENV/bin/python" ]; then
    echo "[2/4] Creating the tool's Python environment (first run only)..."
    if ! "$PY" -m venv "$VENV"; then
        echo ""
        echo "      Could not create the Python environment in:"
        echo "      $(pwd)"
        echo "      Move the tool folder somewhere you can write to (e.g. Documents) and try again."
        finish 1
    fi
else
    echo "[2/4] Python environment ready."
fi

# -----------------------------------------------------------------------------
# 3. Libraries - installed on first run and whenever requirements.txt changes
# -----------------------------------------------------------------------------
STAMP="$VENV/.requirements.installed"
if [ ! -f "$STAMP" ] || ! cmp -s requirements.txt "$STAMP"; then
    echo "[3/4] Installing libraries (Streamlit, Pandas, OpenPyXL). First run takes a few minutes..."
    "$VENV/bin/python" -m pip install --quiet --disable-pip-version-check --upgrade pip >/dev/null 2>&1
    if "$VENV/bin/python" -m pip install --quiet --disable-pip-version-check -r requirements.txt; then
        cp requirements.txt "$STAMP"
    else
        echo ""
        echo "      The libraries could not be downloaded from the Python Package Index (pypi.org)."
        echo "        - Check that you are connected to the internet (and VPN, if required)."
        echo "        - If your network uses a proxy, it may be blocking pypi.org."
        echo "          Ask your eTech team to allow pypi.org and files.pythonhosted.org."
        echo "      Then double-click this launcher again."
        finish 1
    fi
else
    echo "[3/4] Libraries up to date."
fi

# -----------------------------------------------------------------------------
# 4. Launch
# -----------------------------------------------------------------------------
echo "[4/4] Starting Host Sizer - your browser will open shortly."
echo "      Keep this window open while you use the tool. Close it to stop the tool."
echo ""
"$VENV/bin/python" sizing_app.py --setup >/dev/null 2>&1   # AHEAD theme + skip Streamlit's first-run prompt
exec "$VENV/bin/python" -m streamlit run sizing_app.py
