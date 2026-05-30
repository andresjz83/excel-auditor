#!/bin/bash
# Excel Auditor — clean uninstall. Removes the add-in from Excel and the
# launchd dev server. Does not delete the source tree (do that yourself).
set -euo pipefail

WEF="$HOME/Library/Containers/com.microsoft.Excel/Data/Documents/wef"
PLIST="$HOME/Library/LaunchAgents/com.excel-auditor.dev-server.plist"

bold() { printf "\033[1m%s\033[0m\n" "$1"; }
ok()   { printf "  \033[32m✓\033[0m %s\n" "$1"; }

bold "Excel Auditor uninstaller"
echo

# Stop and remove launchd
if [ -f "$PLIST" ]; then
  launchctl bootout "gui/$UID" "$PLIST" 2>/dev/null || true
  rm -f "$PLIST"
  ok "Removed launchd job"
else
  ok "No launchd job to remove"
fi

# Kill any dev server still listening
if lsof -ti:3000 >/dev/null 2>&1; then
  lsof -ti:3000 | xargs kill -9 2>/dev/null || true
  ok "Killed server on :3000"
fi

# Remove sideloaded manifest
if [ -f "$WEF/excel-auditor-manifest.xml" ]; then
  rm -f "$WEF/excel-auditor-manifest.xml"
  ok "Removed sideloaded manifest"
fi

# (Optionally) clear the DevTools flag — leave it on by default, harmless.
echo
bold "Uninstalled."
echo "Logs preserved at \$HOME/.config/excel-auditor/logs/ — delete manually if you want."
echo "The dev cert at \$HOME/.office-addin-dev-certs/ is shared with other Office add-ins — left in place."
