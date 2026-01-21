#!/bin/bash
clear
echo "==========================================="
echo "   RVTools Sizer - Automated Installer"
echo "==========================================="
echo ""

# 1. Check for Apple Command Line Tools (Required for Pip)
echo "[1/3] Checking Apple Developer Tools..."
if xcode-select -p &>/dev/null; then
    echo "      ✅ Tools found."
else
    echo "      ❌ Tools missing. Requesting install..."
    echo "      ⚠️ A POPUP WINDOW WILL APPEAR. PLEASE CLICK 'INSTALL'."
    xcode-select --install
    
    # Wait loop until user installs
    echo "      Waiting for installation to finish..."
    while ! xcode-select -p &>/dev/null; do
        sleep 5
    done
    echo "      ✅ Tools installed successfully."
fi

echo ""

# 2. Check Python
echo "[2/3] Checking Python Environment..."
if command -v python3 &>/dev/null; then
    echo "      ✅ Python3 found."
else
    echo "      ❌ Python3 not found. macOS usually comes with it."
    echo "      Please install Python from python.org if this fails."
    exit 1
fi

echo ""

# 3. Install Libraries
echo "[3/3] Installing Sizer Libraries (Streamlit, Pandas)..."
pip3 install --user streamlit pandas openpyxl

echo ""
echo "==========================================="
echo "   🎉 Installation Complete!"
echo "   You can now double-click 'Sizing Calculator' to run."
echo "==========================================="
echo ""
read -p "Press [Enter] to close..."