#!/usr/bin/env bash
# ════════════════════════════════════════════════════════════════════════════════
#  Webiz Inventory Agent — macOS Installer
#  Single-file: just run this once.  Everything else is automatic.
#
#  Usage:  chmod +x WebizInventory_macOS.sh && ./WebizInventory_macOS.sh
# ════════════════════════════════════════════════════════════════════════════════

# ─── CONFIGURATION — edit these two lines before distributing ────────────────
APPS_SCRIPT_URL="https://script.google.com/macros/s/AKfycbxUVGyr5SuH7gjEc7zS5CcZkDV03qVGw7JbPHwvTFwLEUImY3xbRE8V8D4SQNalBMUdGw/exec"          # ← FILL IN
GITHUB_RAW_URL="https://raw.githubusercontent.com/Nikulina123/Check-in_agent/main/inventory_agent.py"
# ─────────────────────────────────────────────────────────────────────────────

set -uo pipefail   # note: -e removed so we can show real errors before exiting

# ── Error trap: show what failed and keep window open ────────────────────────
_die() {
    local line="$1"
    echo ""
    echo "╔══════════════════════════════════════════════════════════╗"
    echo "║  ✗  Installation failed (line $line)                    "
    echo "║     Check the error above, fix it, and re-run.          ║"
    echo "╚══════════════════════════════════════════════════════════╝"
    echo ""
    read -rp "Press Enter to close…" _
    exit 1
}
trap '_die $LINENO' ERR

AGENT_DIR="$HOME/Library/Application Support/WebizInventory"
AGENT_FILE="$AGENT_DIR/inventory_agent.py"
CONFIG_FILE="$AGENT_DIR/config.json"
PLIST_DIR="$HOME/Library/LaunchAgents"
PLIST_LABEL="com.webiz.inventory"
PLIST_FILE="$PLIST_DIR/$PLIST_LABEL.plist"

echo ""
echo "┌─────────────────────────────────────────┐"
echo "│   Webiz Inventory Agent – macOS Setup   │"
echo "└─────────────────────────────────────────┘"
echo ""

# ── Step 1: Locate Python 3 ───────────────────────────────────────────────────
echo "[1/5] Locating Python 3…"
PYTHON3=""
for candidate in python3 python; do
    if command -v "$candidate" &>/dev/null; then
        ver=$("$candidate" --version 2>&1 || true)
        if [[ "$ver" == Python\ 3* ]]; then
            PYTHON3=$(command -v "$candidate")
            echo "      Found: $PYTHON3  ($ver)"
            break
        fi
    fi
done

if [[ -z "$PYTHON3" ]]; then
    echo "      Python 3 not found — attempting to install via Homebrew…"
    if command -v brew &>/dev/null; then
        brew install python3
        PYTHON3=$(command -v python3)
    else
        echo ""
        echo "  ┌─ ACTION REQUIRED ─────────────────────────────────────────┐"
        echo "  │ Homebrew not found. Please install Python 3 manually:     │"
        echo "  │   https://www.python.org/downloads/                       │"
        echo "  │ Then re-run this script.                                  │"
        echo "  └───────────────────────────────────────────────────────────┘"
        exit 1
    fi
fi

# Check tkinter (bundled with python.org builds, may be absent in some Homebrew builds)
if ! "$PYTHON3" -c "import tkinter" 2>/dev/null; then
    echo "      tkinter missing — trying: brew install python-tk"
    brew install python-tk 2>/dev/null || true
    if ! "$PYTHON3" -c "import tkinter" 2>/dev/null; then
        echo "  [ERROR] tkinter still missing. Install python-tk manually."
        exit 1
    fi
fi
echo "      Python + tkinter: OK"

# ── Step 2: Download the agent ────────────────────────────────────────────────
echo ""
echo "[2/5] Downloading inventory agent from GitHub…"
echo "      URL: $GITHUB_RAW_URL"
mkdir -p "$AGENT_DIR"

DOWNLOAD_OK=false
if command -v curl &>/dev/null; then
    # -f: fail on HTTP errors  -L: follow redirects  --retry 3: retry on network glitch
    # -v removed from -s so errors print to terminal
    if curl -fL --retry 3 --retry-delay 2 --connect-timeout 15 \
            "$GITHUB_RAW_URL" -o "$AGENT_FILE" 2>&1; then
        DOWNLOAD_OK=true
    fi
fi

if [[ "$DOWNLOAD_OK" == false ]] && command -v wget &>/dev/null; then
    echo "      curl failed — trying wget…"
    if wget --tries=3 --timeout=15 "$GITHUB_RAW_URL" -O "$AGENT_FILE" 2>&1; then
        DOWNLOAD_OK=true
    fi
fi

if [[ "$DOWNLOAD_OK" == false ]]; then
    echo "      curl/wget failed — trying Python urllib…"
    "$PYTHON3" -c "
import urllib.request, sys
try:
    urllib.request.urlretrieve('$GITHUB_RAW_URL', '$AGENT_FILE')
    print('      Downloaded via Python urllib.')
except Exception as e:
    print(f'      urllib failed: {e}', file=sys.stderr)
    sys.exit(1)
" && DOWNLOAD_OK=true
fi

if [[ "$DOWNLOAD_OK" == false ]]; then
    echo ""
    echo "  [ERROR] Could not download inventory_agent.py."
    echo "  Possible causes:"
    echo "    • GitHub repo is private — make it public or check the URL"
    echo "    • File not yet pushed to the repo"
    echo "    • No internet connection"
    echo "  URL tried: $GITHUB_RAW_URL"
    exit 1
fi

# Validate: file must exist and not be empty HTML (GitHub 404 page)
if [[ ! -s "$AGENT_FILE" ]]; then
    echo "  [ERROR] Downloaded file is empty."
    exit 1
fi
if head -1 "$AGENT_FILE" | grep -qi "<!DOCTYPE\|<html"; then
    echo "  [ERROR] GitHub returned an HTML page instead of the Python file."
    echo "  The URL is probably wrong or the repo/file does not exist yet:"
    echo "  $GITHUB_RAW_URL"
    rm -f "$AGENT_FILE"
    exit 1
fi

chmod +x "$AGENT_FILE"
echo "      ✔  Agent saved to: $AGENT_FILE"

# ── Step 3: Write config ──────────────────────────────────────────────────────
echo ""
echo "[3/5] Writing configuration…"
cat > "$CONFIG_FILE" <<JSON
{
  "apps_script_url": "$APPS_SCRIPT_URL",
  "github_raw_url":  "$GITHUB_RAW_URL"
}
JSON
echo "      Config: $CONFIG_FILE"

# ── Step 4: First manual run ──────────────────────────────────────────────────
echo ""
echo "[4/5] Running agent for the first time…"
"$PYTHON3" "$AGENT_FILE" || true   # allow cancel without aborting installer

# ── Step 5: Install LaunchAgent ───────────────────────────────────────────────
echo ""
echo "[5/5] Installing LaunchAgent (runs at every login)…"
mkdir -p "$PLIST_DIR"

cat > "$PLIST_FILE" <<PLIST
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN"
    "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
    <key>Label</key>
    <string>${PLIST_LABEL}</string>

    <key>ProgramArguments</key>
    <array>
        <string>${PYTHON3}</string>
        <string>${AGENT_FILE}</string>
    </array>

    <key>WorkingDirectory</key>
    <string>${AGENT_DIR}</string>

    <!-- Run once at login (agent enforces the 6-month interval itself) -->
    <!-- The Python script sleeps 90 s internally when not running in a terminal -->
    <key>RunAtLoad</key>
    <true/>

    <key>StandardOutPath</key>
    <string>${AGENT_DIR}/launchd_stdout.log</string>
    <key>StandardErrorPath</key>
    <string>${AGENT_DIR}/launchd_stderr.log</string>

    <!-- Only launch when a GUI session is active -->
    <key>LimitLoadToSessionType</key>
    <string>Aqua</string>
</dict>
</plist>
PLIST

# Unload any existing version, then load new one
launchctl unload "$PLIST_FILE" 2>/dev/null || true
launchctl load -w "$PLIST_FILE"

# ── Ad-hoc code signing (prevents Gatekeeper quarantine) ─────────────────────
echo ""
echo "      Signing agent script (ad-hoc)…"
if command -v codesign &>/dev/null; then
    codesign --force --sign - "$AGENT_FILE" 2>/dev/null && \
        echo "      Signed: $AGENT_FILE" || \
        echo "      [WARN] codesign failed — agent will still run."
else
    echo "      [WARN] codesign not found (Xcode CLT not installed). Skipping signing."
fi

# Clear quarantine flag from this installer script itself
xattr -r -d com.apple.quarantine "$AGENT_FILE" 2>/dev/null || true

echo ""
echo "✔  Installation complete."
echo ""
echo "   The agent runs silently at every login."
echo "   It shows the form only when 6 months have passed since the last check-in."
echo ""
echo "   Useful commands:"
echo "   launchctl list | grep webiz            # check if loaded"
echo "   launchctl start $PLIST_LABEL           # trigger manually"
echo "   tail -f \"$AGENT_DIR/agent.log\"        # live log"
echo ""
echo "   To uninstall:"
echo "   launchctl unload ~/Library/LaunchAgents/$PLIST_LABEL.plist"
echo "   rm ~/Library/LaunchAgents/$PLIST_LABEL.plist"
echo ""
read -rp "Press Enter to close…" _
