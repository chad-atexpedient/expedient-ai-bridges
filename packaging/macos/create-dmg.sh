#!/usr/bin/env bash
set -euo pipefail
app="$1"; output="$2"; volume="Expedient AI Bridges"
[[ -d "$app" ]] || { echo "App bundle not found" >&2; exit 1; }
tmp="$(mktemp -d)"; trap 'rm -rf "$tmp"' EXIT
cp -R "$app" "$tmp/"
ln -s /Applications "$tmp/Applications"
hdiutil create -quiet -fs HFS+ -format UDZO -volname "$volume" -srcfolder "$tmp" "$output"
