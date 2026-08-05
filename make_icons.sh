#!/bin/bash
# Regenerate every icon format from icon.svg. Run this after editing the SVG.
#
#   ./make_icons.sh
#
# Produces:
#   icon.png    512px, loaded by Tk for the window icon at runtime
#   icon.icns   macOS app bundle icon
#   icon.ico    Windows .exe icon, multi-resolution
#
# Needs librsvg (brew install librsvg) and Pillow, which is already in the venv.
# Every size is rendered from the vector rather than resampled from one big PNG,
# so the 16px and 32px versions stay legible instead of turning to mush.
set -euo pipefail
cd "$(dirname "$0")"

command -v rsvg-convert >/dev/null || { echo "rsvg-convert not found: brew install librsvg"; exit 1; }

PY=./venv/bin/python
[ -x "$PY" ] || PY=python3

TMP=$(mktemp -d)
trap 'rm -rf "$TMP"' EXIT

for s in 16 32 64 128 256 512 1024; do
  rsvg-convert -w $s -h $s icon.svg -o "$TMP/r$s.png"
done

# ── macOS ──
IS="$TMP/icon.iconset"; mkdir -p "$IS"
cp "$TMP/r16.png"   "$IS/icon_16x16.png"
cp "$TMP/r32.png"   "$IS/icon_16x16@2x.png"
cp "$TMP/r32.png"   "$IS/icon_32x32.png"
cp "$TMP/r64.png"   "$IS/icon_32x32@2x.png"
cp "$TMP/r128.png"  "$IS/icon_128x128.png"
cp "$TMP/r256.png"  "$IS/icon_128x128@2x.png"
cp "$TMP/r256.png"  "$IS/icon_256x256.png"
cp "$TMP/r512.png"  "$IS/icon_256x256@2x.png"
cp "$TMP/r512.png"  "$IS/icon_512x512.png"
cp "$TMP/r1024.png" "$IS/icon_512x512@2x.png"
iconutil -c icns "$IS" -o icon.icns

# ── Windows and Tk ──
"$PY" - "$TMP" <<'PYEOF'
import sys
from PIL import Image
tmp = sys.argv[1]
pairs = [(16, 16), (24, 32), (32, 32), (48, 64), (64, 64), (128, 128), (256, 256)]
imgs = [Image.open(f"{tmp}/r{src}.png").convert("RGBA").resize((out, out), Image.LANCZOS)
        for out, src in pairs]
imgs[-1].save("icon.ico", format="ICO", sizes=[(o, o) for o, _ in pairs])
Image.open(f"{tmp}/r512.png").convert("RGBA").save("icon.png")
PYEOF

ls -lh icon.svg icon.png icon.ico icon.icns
echo "done"
