#!/bin/bash
# Excel Auditor installer for macOS.
# Creates the Office add-in sideload folder if missing, downloads the
# manifest from GitHub Pages, places it in the right spot.
set -e

WEF="$HOME/Library/Containers/com.microsoft.Excel/Data/Documents/wef"
MANIFEST_URL="https://andresjz83.github.io/excel-auditor/manifest.xml"

echo ""
echo "==================================================="
echo "  Excel Auditor installer"
echo "==================================================="
echo ""

if [ ! -d "$HOME/Library/Containers/com.microsoft.Excel" ]; then
  echo "ERROR: Excel for Mac doesn't seem to be installed."
  echo "Install Excel from the Mac App Store or Microsoft 365 first."
  read -p "Press Enter to close this window..."
  exit 1
fi

echo "Creating sideload folder if missing..."
mkdir -p "$WEF"
echo "  Done."

echo ""
echo "Downloading the latest manifest..."
if curl -fsSL -o "$WEF/excel-auditor-manifest.xml" "$MANIFEST_URL"; then
  echo "  Downloaded to: $WEF/excel-auditor-manifest.xml"
else
  echo "ERROR: Could not download the manifest. Check your internet connection."
  read -p "Press Enter to close this window..."
  exit 1
fi

echo ""
echo "==================================================="
echo "  Installed!"
echo "==================================================="
echo ""
echo "Next steps:"
echo "  1. Cmd+Q to fully quit Excel if it's open."
echo "  2. Reopen Excel."
echo "  3. Open the auditor via Insert -> Add-ins -> Excel Auditor."
echo ""
echo "You will need to do step 3 every time you launch Excel."
echo "Microsoft does not pin sideloaded add-ins to the Mac ribbon."
echo ""
read -p "Press Enter to close this window..."
