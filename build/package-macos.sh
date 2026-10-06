#!/usr/bin/env bash
# Builds "Hypervisor Explorer.app" and the hvexplorer CLI for macOS.
#
# By default the app is universal: it contains self-contained Apple Silicon (arm64) and Intel (x64) builds, and a
# small launcher picks the right one at start-up. (.NET single-file executables can't be merged with lipo because the
# bundle is located by file offset, so the two builds sit side by side instead.)
#
# Usage: build/package-macos.sh [version] [universal|osx-arm64|osx-x64]
#        OUT_DIR=/some/dir ZIP=1 build/package-macos.sh v3.0.1
set -euo pipefail
VERSION="${1:-0.0.0-dev}"; TARGET="${2:-universal}"
SEMVER="${VERSION#v}"; NUMERIC="${SEMVER%%-*}"
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
OUT="${OUT_DIR:-$ROOT/artifacts/publish}"; TMP="$OUT/../tmp-macos-$$"
APP="$OUT/Hypervisor Explorer.app"
mkdir -p "$OUT"; rm -rf "$TMP"

case "$TARGET" in
  universal) RUNTIMES=(osx-arm64 osx-x64) ;;
  osx-arm64|osx-x64) RUNTIMES=("$TARGET") ;;
  *) echo "Unknown target $TARGET" >&2; exit 1 ;;
esac

common=(-c Release --self-contained -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true
        -p:DebugType=none "-p:Version=$NUMERIC" "-p:InformationalVersion=$SEMVER" -v q -nologo)
for rid in "${RUNTIMES[@]}"; do
  dotnet publish "$ROOT/src/HypervisorExplorer.App" "${common[@]}" -r "$rid" -o "$TMP/app-$rid"
  dotnet publish "$ROOT/src/HypervisorExplorer.Cli" "${common[@]}" -r "$rid" -o "$TMP/cli-$rid"
done

# Launcher: run the native build for this Mac's CPU. sysctl reports the hardware even under Rosetta.
write_launcher() { # $1 = output path, $2 = binary base name next to it
  cat > "$1" <<LAUNCHER
#!/bin/sh
DIR="\$(cd "\$(dirname "\$0")" && pwd)"
if [ "\$(/usr/sbin/sysctl -n hw.optional.arm64 2>/dev/null)" = "1" ] && [ -x "\$DIR/$2-arm64" ]; then
  exec "\$DIR/$2-arm64" "\$@"
fi
exec "\$DIR/$2-x64" "\$@"
LAUNCHER
  chmod +x "$1"
}

rm -rf "$APP"; mkdir -p "$APP/Contents/MacOS" "$APP/Contents/Resources"
if [[ ${#RUNTIMES[@]} -eq 2 ]]; then
  cp "$TMP/app-osx-arm64/HypervisorExplorer" "$APP/Contents/MacOS/HypervisorExplorer-arm64"
  cp "$TMP/app-osx-x64/HypervisorExplorer" "$APP/Contents/MacOS/HypervisorExplorer-x64"
  write_launcher "$APP/Contents/MacOS/HypervisorExplorer" HypervisorExplorer
else
  cp "$TMP/app-${RUNTIMES[0]}/HypervisorExplorer" "$APP/Contents/MacOS/HypervisorExplorer"
fi

cat > "$APP/Contents/Info.plist" <<PLIST
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
    <key>CFBundleName</key><string>Hypervisor Explorer</string>
    <key>CFBundleDisplayName</key><string>Hypervisor Explorer</string>
    <key>CFBundleIdentifier</key><string>com.superfastmart.hypervisorexplorer</string>
    <key>CFBundleVersion</key><string>$NUMERIC</string>
    <key>CFBundleShortVersionString</key><string>$NUMERIC</string>
    <key>CFBundleExecutable</key><string>HypervisorExplorer</string>
    <key>CFBundlePackageType</key><string>APPL</string>
    <key>LSMinimumSystemVersion</key><string>12.0</string>
    <key>NSHighResolutionCapable</key><true/>
</dict>
</plist>
PLIST
codesign --force --deep -s - "$APP" >/dev/null 2>&1 || true

# CLI: same arrangement (launcher + per-architecture binaries), or a single binary for one target.
CLI_FILES=()
if [[ ${#RUNTIMES[@]} -eq 2 ]]; then
  cp "$TMP/cli-osx-arm64/hvexplorer" "$OUT/hvexplorer-arm64"
  cp "$TMP/cli-osx-x64/hvexplorer" "$OUT/hvexplorer-x64"
  write_launcher "$OUT/hvexplorer" hvexplorer
  CLI_FILES=(hvexplorer hvexplorer-arm64 hvexplorer-x64)
else
  cp "$TMP/cli-${RUNTIMES[0]}/hvexplorer" "$OUT/hvexplorer"
  CLI_FILES=(hvexplorer)
fi

if [[ "${ZIP:-0}" == "1" ]]; then
  # Stage the app and CLI together, then zip the folder's contents (ditto keeps the bundle and exec bits intact).
  SUFFIX="$TARGET"; [[ "$TARGET" == "universal" ]] && SUFFIX="macos-universal"
  ZIPFILE="$OUT/HypervisorExplorer-$SEMVER-$SUFFIX.zip"
  STAGE="$TMP/zip"
  rm -rf "$STAGE" "$ZIPFILE"; mkdir -p "$STAGE"
  cp -R "$APP" "$STAGE/"
  for f in "${CLI_FILES[@]}"; do cp "$OUT/$f" "$STAGE/"; done
  ditto -c -k "$STAGE" "$ZIPFILE"
  echo "Created $ZIPFILE"
fi
rm -rf "$TMP"
echo "Built $APP"
