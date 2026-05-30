#!/bin/bash
# Excel Auditor dev server. Started by launchd at login.
# Serves the add-in HTML/JS/CSS at https://localhost:3000 using the Office dev cert.
set -euo pipefail

# Resolve paths relative to this script — portable across install locations.
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
ROOT="$(cd "$SCRIPT_DIR/.." && pwd)"
CERT_DIR="$HOME/.office-addin-dev-certs"
LOG_DIR="$HOME/.config/excel-auditor/logs"
mkdir -p "$LOG_DIR"

if [ ! -f "$CERT_DIR/localhost.crt" ] || [ ! -f "$CERT_DIR/localhost.key" ]; then
  echo "[$(date)] ERROR: dev certs missing at $CERT_DIR — run install.sh or 'npx office-addin-dev-certs install'." >&2
  exit 1
fi

# Find npx via common locations (launchd PATH is minimal).
NPX_BIN=""
for p in /opt/homebrew/bin/npx /usr/local/bin/npx "$HOME/.nvm/versions/node"/*/bin/npx; do
  if [ -x "$p" ]; then NPX_BIN="$p"; break; fi
done
if [ -z "$NPX_BIN" ]; then
  echo "[$(date)] ERROR: npx not found. Install Node via 'brew install node'." >&2
  exit 1
fi

cd "$ROOT/src"
# Bind to 127.0.0.1 only — otherwise http-server listens on 0.0.0.0 and the
# add-in source becomes reachable from anyone on the same network.
exec "$NPX_BIN" http-server \
  -S \
  -a 127.0.0.1 \
  -C "$CERT_DIR/localhost.crt" \
  -K "$CERT_DIR/localhost.key" \
  -p 3000 \
  -c-1 \
  --cors \
  --silent \
  .
