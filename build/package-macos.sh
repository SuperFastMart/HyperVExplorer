#!/usr/bin/env bash
# Builds "Hypervisor Explorer.app" (self-contained, ad-hoc signed) and the hvexplorer CLI for macOS.
# Usage: build/package-macos.sh [version] [runtime]     e.g. build/package-macos.sh v3.0.0 osx-arm64
set -euo pipefail
VERSION="${1:-0.0.0-dev}"; RUNTIME="${2:-osx-arm64}"
SEMVER="${VERSION#v}"; NUMERIC="${SEMVER%%-*}"
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
OUT="$ROOT/artifacts/publish"; TMP="$ROOT/artifacts/tmp-$RUNTIME"
APP="$OUT/Hypervisor Explorer.app"
rm -rf "$TMP"; mkdir -p "$OUT"

common=(-c Release -r "$RUNTIME" --self-contained -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true
        -p:DebugType=none "-p:Version=$NUMERIC" "-p:InformationalVersion=$SEMVER" -v q -nologo)
dotnet publish "$ROOT/src/HypervisorExplorer.App" "${common[@]}" -o "$TMP/app"
dotnet publish "$ROOT/src/HypervisorExplorer.Cli" "${common[@]}" -o "$TMP/cli"

rm -rf "$APP"; mkdir -p "$APP/Contents/MacOS" "$APP/Contents/Resources"
cp "$TMP/app/HypervisorExplorer" "$APP/Contents/MacOS/"
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
cp "$TMP/cli/hvexplorer" "$OUT/hvexplorer"
rm -rf "$TMP"

if [[ "${ZIP:-0}" == "1" ]]; then
  ZIPFILE="$OUT/HypervisorExplorer-$SEMVER-$RUNTIME.zip"
  rm -f "$ZIPFILE"
  (cd "$OUT" && ditto -c -k --keepParent "Hypervisor Explorer.app" "$ZIPFILE.tmp" && zip -qj "$ZIPFILE.tmp" hvexplorer && mv "$ZIPFILE.tmp" "$ZIPFILE")
  echo "Created $ZIPFILE"
fi
echo "Built $APP"
