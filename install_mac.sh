#!/bin/bash
# =============================================================================
#  Host Sizer - Mac installer and updater
#
#  Open Terminal, paste this line and press Return:
#    curl -fsSL https://raw.githubusercontent.com/adammarcum/Host-sizing-Tool/main/install_mac.sh | bash
#
#  Downloads the latest version from GitHub into ~/HostSizer/app and starts it.
#  Run the same line again any time to update; your Python environment and
#  saved projects (~/HostSizer/Projects) are kept.
#  Files downloaded this way are not quarantined, so macOS doesn't block the
#  app or run it from a temporary copy.
# =============================================================================
set -euo pipefail

DEST="${HOSTSIZER_HOME:-$HOME/HostSizer}/app"
ZIP_URL="${HOSTSIZER_ZIP_URL:-https://github.com/adammarcum/Host-sizing-Tool/archive/refs/heads/main.zip}"
TMP="$(mktemp -d)"
trap 'rm -rf "$TMP"' EXIT

echo "==========================================="
echo "   Host Sizer  -  installer / updater"
echo "==========================================="
echo "Downloading the latest version..."
if ! curl -fsSL "$ZIP_URL" -o "$TMP/hostsizer.zip"; then
    echo ""
    echo "Download failed. Check your internet connection (and VPN, if required)"
    echo "and that github.com is reachable, then run the command again."
    exit 1
fi

unzip -q "$TMP/hostsizer.zip" -d "$TMP/unzipped"
SRC="$(find "$TMP/unzipped" -mindepth 1 -maxdepth 1 -type d | head -1)"
if [ -z "$SRC" ] || [ ! -f "$SRC/sizing_app.py" ]; then
    echo "The download didn't contain Host Sizer. Please try again later."
    exit 1
fi

mkdir -p "$DEST"
cp -R "$SRC/." "$DEST/"          # updates the app files, keeps .venv and settings
chmod +x "$DEST/Run_on_mac.command" 2>/dev/null || true
VERSION="$(sed -n 's/^APP_VERSION = "\(.*\)"/\1/p' "$DEST/sizing_app.py" | head -1)"

echo "Installed Host Sizer ${VERSION} in: $DEST"
echo "Next time, double-click 'Host Sizing Calculator.app' in that folder,"
echo "or run this same command again to update."
echo ""
open "$DEST" 2>/dev/null || true   # show the folder in Finder

# Start the tool. Read prompts from the keyboard, not from the curl pipe.
if [ -r /dev/tty ]; then
    exec bash "$DEST/Run_on_mac.command" < /dev/tty
else
    exec bash "$DEST/Run_on_mac.command"
fi
