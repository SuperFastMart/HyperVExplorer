#!/usr/bin/env bash
# Hypervisor Explorer installer / updater for macOS (Intel and Apple Silicon).
#
#   curl -fsSL https://raw.githubusercontent.com/SuperFastMart/HyperVisorExplorer/main/install.sh | bash
#
# Files fetched with curl aren't quarantined, so the unsigned app opens without a Gatekeeper prompt.
# Re-run at any time to update. Options (environment variables):
#   HVE_VERSION=v3.0.1      install a specific release instead of the latest
#   HVE_APP_DIR=/path       where to put "Hypervisor Explorer.app" (default /Applications, or ~/Applications)
#   HVE_BIN_DIR=/path       where to put the hvexplorer CLI (default ~/.local/bin); HVE_NO_CLI=1 to skip
#   HVE_NO_LAUNCH=1         don't open the app afterwards
#   HVE_WAIT_PID=1234       wait for this process to exit first (used by the app's own updater)
set -euo pipefail

REPO="SuperFastMart/HyperVisorExplorer"
APP_NAME="Hypervisor Explorer.app"

say() { printf '\033[1;34m==>\033[0m %s\n' "$*"; }
die() { printf '\033[1;31mError:\033[0m %s\n' "$*" >&2; exit 1; }

[[ "$(uname -s)" == "Darwin" ]] || die "This installer is for macOS. On Windows use install.ps1."

# ---------------------------------------------------------------- find the release asset
if [[ -n "${HVE_VERSION:-}" ]]; then
  API="https://api.github.com/repos/$REPO/releases/tags/${HVE_VERSION}"
else
  API="https://api.github.com/repos/$REPO/releases/latest"
fi
say "Looking up ${HVE_VERSION:-the latest} release..."
JSON="$(curl -fsSL -H 'Accept: application/vnd.github+json' "$API")" || die "Could not reach GitHub ($API)."
TAG="$(printf '%s' "$JSON" | grep -o '"tag_name": *"[^"]*"' | head -1 | sed 's/.*"\([^"]*\)"$/\1/')"
URL="$(printf '%s' "$JSON" | grep -o '"browser_download_url": *"[^"]*macos-universal\.zip"' | head -1 | sed 's/.*"\(https[^"]*\)"$/\1/')"
[[ -n "$URL" ]] || die "No macOS download found in release ${TAG:-?}."

# ---------------------------------------------------------------- choose install locations
APP_DIR="${HVE_APP_DIR:-}"
if [[ -z "$APP_DIR" ]]; then
  if [[ -w /Applications ]]; then APP_DIR=/Applications; else APP_DIR="$HOME/Applications"; fi
fi
mkdir -p "$APP_DIR" || die "Cannot create $APP_DIR."
BIN_DIR="${HVE_BIN_DIR:-$HOME/.local/bin}"

# ---------------------------------------------------------------- download and unpack
TMP="$(mktemp -d -t hve-install)"
trap 'rm -rf "$TMP"' EXIT
say "Downloading Hypervisor Explorer $TAG..."
curl -fL --progress-bar -o "$TMP/hve.zip" "$URL" || die "Download failed."
ditto -x -k "$TMP/hve.zip" "$TMP/x" || die "Could not unpack the download."
[[ -d "$TMP/x/$APP_NAME" ]] || die "The download doesn't contain $APP_NAME."

# ---------------------------------------------------------------- close the running app
if [[ -n "${HVE_WAIT_PID:-}" ]]; then
  say "Waiting for Hypervisor Explorer to close..."
  for _ in $(seq 1 60); do kill -0 "$HVE_WAIT_PID" 2>/dev/null || break; sleep 0.5; done
fi
# Only the copy being replaced is closed (matched by its path), never other copies of the app.
RUNNING="$(pgrep -f "$APP_DIR/$APP_NAME/Contents/MacOS/" || true)"
if [[ -n "$RUNNING" ]]; then
  say "Closing the running copy of Hypervisor Explorer..."
  kill $RUNNING 2>/dev/null || true
  for _ in $(seq 1 20); do pgrep -f "$APP_DIR/$APP_NAME/Contents/MacOS/" >/dev/null 2>&1 || break; sleep 0.5; done
  pkill -9 -f "$APP_DIR/$APP_NAME/Contents/MacOS/" >/dev/null 2>&1 || true
  sleep 1
fi

# ---------------------------------------------------------------- install
say "Installing to $APP_DIR/$APP_NAME"
rm -rf "$APP_DIR/$APP_NAME"
mv "$TMP/x/$APP_NAME" "$APP_DIR/$APP_NAME"
xattr -dr com.apple.quarantine "$APP_DIR/$APP_NAME" 2>/dev/null || true

if [[ -z "${HVE_NO_CLI:-}" ]]; then
  mkdir -p "$BIN_DIR"
  for f in hvexplorer hvexplorer-arm64 hvexplorer-x64; do
    [[ -f "$TMP/x/$f" ]] && cp "$TMP/x/$f" "$BIN_DIR/$f" && chmod +x "$BIN_DIR/$f"
  done
  xattr -d com.apple.quarantine "$BIN_DIR"/hvexplorer* 2>/dev/null || true
  say "Command-line tool installed to $BIN_DIR/hvexplorer"
  case ":$PATH:" in
    *":$BIN_DIR:"*) ;;
    *) echo "    (add it to your PATH: echo 'export PATH=\"$BIN_DIR:\$PATH\"' >> ~/.zshrc)" ;;
  esac
fi

say "Hypervisor Explorer $TAG is installed."
if [[ -z "${HVE_NO_LAUNCH:-}" ]]; then
  open "$APP_DIR/$APP_NAME"
fi
