#!/bin/bash
# Excel Auditor — one-shot installer for macOS.
# Idempotent: re-running fixes a broken install rather than duplicating one.
set -euo pipefail

ROOT="$(cd "$(dirname "$0")" && pwd)"
WEF="$HOME/Library/Containers/com.microsoft.Excel/Data/Documents/wef"
LAUNCHAGENTS="$HOME/Library/LaunchAgents"
PLIST_NAME="com.excel-auditor.dev-server.plist"
PLIST_DST="$LAUNCHAGENTS/$PLIST_NAME"
LOG_DIR="$HOME/.config/excel-auditor/logs"

bold()  { printf "\033[1m%s\033[0m\n" "$1"; }
ok()    { printf "  \033[32m✓\033[0m %s\n" "$1"; }
warn()  { printf "  \033[33m!\033[0m %s\n" "$1"; }
fail()  { printf "  \033[31m✗\033[0m %s\n" "$1" >&2; exit 1; }

bold "Excel Auditor installer"
echo "Install root: $ROOT"
echo

# ─── Prerequisite checks ──────────────────────────────────────────────────
bold "1. Checking prerequisites"

if [[ "$(uname)" != "Darwin" ]]; then
  fail "This add-in is macOS-only (Excel for Mac). You're on $(uname)."
fi
ok "macOS"

if [ ! -d "/Applications/Microsoft Excel.app" ]; then
  fail "Microsoft Excel for Mac not found at /Applications/Microsoft Excel.app. Install it first."
fi
ok "Excel for Mac"

if ! command -v node >/dev/null 2>&1; then
  fail "Node.js not found. Install with: brew install node"
fi
NODE_VERSION="$(node --version)"
ok "Node.js $NODE_VERSION"

if ! command -v npm >/dev/null 2>&1; then
  fail "npm not found (came with Node? path issue?)"
fi
ok "npm $(npm --version)"

echo

# ─── npm install ──────────────────────────────────────────────────────────
bold "2. Installing dependencies"
cd "$ROOT"
npm install --silent --no-fund --no-audit
ok "npm dependencies installed"
echo

# ─── Office.js dev certificate ────────────────────────────────────────────
bold "3. Ensuring Office.js dev certificate"
CERT_DIR="$HOME/.office-addin-dev-certs"
if [ -f "$CERT_DIR/localhost.crt" ] && [ -f "$CERT_DIR/localhost.key" ]; then
  ok "Certificate already installed at $CERT_DIR"
else
  echo "  Installing certificate (may prompt for your password to add to keychain)…"
  npx --yes office-addin-dev-certs install
  ok "Certificate installed"
fi
echo

# ─── Manifest: assign a fresh per-machine GUID and sideload ──────────────
bold "4. Sideloading manifest"
mkdir -p "$WEF"

# Reuse an existing per-machine GUID if we previously installed, so the
# add-in keeps the same identity across re-installs. Otherwise mint one.
GUID_FILE="$ROOT/.machine-guid"
if [ -f "$GUID_FILE" ]; then
  NEW_GUID="$(cat "$GUID_FILE")"
  ok "Reusing existing GUID $NEW_GUID"
else
  NEW_GUID="$(uuidgen)"
  echo "$NEW_GUID" > "$GUID_FILE"
  ok "Minted new GUID $NEW_GUID"
fi

# Replace the placeholder GUID in the source manifest. The source ships with
# a stable placeholder; we never modify the tracked file.
PLACEHOLDER="2EC926A4-B24F-4CBE-8D09-F5D721BCE55B"
INSTALLED_MANIFEST="$WEF/excel-auditor-manifest.xml"
sed "s|$PLACEHOLDER|$NEW_GUID|g" "$ROOT/manifest.xml" > "$INSTALLED_MANIFEST"
ok "Manifest written to $INSTALLED_MANIFEST"
echo

# ─── DevTools flag for the task pane ──────────────────────────────────────
bold "5. Enabling task-pane DevTools"
defaults write com.microsoft.Excel OfficeWebAddinDeveloperExtras -bool true
ok "Right-click → Inspect Element now works on the add-in pane"
echo

# ─── launchd service ──────────────────────────────────────────────────────
bold "6. Installing the launchd dev server"
mkdir -p "$LOG_DIR"
mkdir -p "$LAUNCHAGENTS"

# Materialise the plist from the template with the install root + home dir.
sed -e "s|__INSTALL_ROOT__|$ROOT|g" \
    -e "s|__HOME__|$HOME|g" \
    "$ROOT/server/com.excel-auditor.dev-server.plist.template" > "$PLIST_DST"
ok "Plist installed at $PLIST_DST"

chmod +x "$ROOT/server/run-server.sh"

# (Re)load the launchd service.
if launchctl print "gui/$UID/com.excel-auditor.dev-server" >/dev/null 2>&1; then
  launchctl bootout "gui/$UID" "$PLIST_DST" 2>/dev/null || true
fi
# Kill any stray http-server holding port 3000 (e.g. left over from a manual run).
if lsof -ti:3000 >/dev/null 2>&1; then
  warn "Port 3000 is in use — killing existing listener so launchd can bind."
  lsof -ti:3000 | xargs kill -9 2>/dev/null || true
  sleep 1
fi
launchctl bootstrap "gui/$UID" "$PLIST_DST"
sleep 2

# Sanity check the server is up.
if curl -ks -o /dev/null -w "%{http_code}" https://localhost:3000/taskpane.html | grep -q "^200$"; then
  ok "Dev server is up at https://localhost:3000"
else
  warn "Server not reachable yet. Check $LOG_DIR/server-error.log."
fi
echo

# ─── Done ─────────────────────────────────────────────────────────────────
bold "Installed."
cat <<EOF

Next steps:
  1. Quit Excel completely (Cmd+Q) if it's open, then reopen it.
  2. Open any workbook.
  3. Home ribbon → "Open Auditor" button (in the Auditor group).
  4. Select cells → "Show precedents" / "Show dependents".

The dev server auto-starts at every login. Uninstall with ./uninstall.sh.
Logs: $LOG_DIR
EOF
