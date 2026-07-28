#!/bin/bash
# =====================================================================
# Slide Aid - build the distributable installer zip for colleagues.
#
# Run on YOUR Mac (needs osacompile, i.e. any macOS) after building
# the add-in (README steps 1-6):
#
#   ./tools/make_dist.sh
#
# Produces  dist/Slide Aid.zip  containing:
#   Slide Aid.ppam      the add-in (with ribbon + icons)
#   SlideAidUI.scpt     pre-compiled color-picker helper
#   slideaid.lua        Hammerspoon shortcut config
#   install.command     double-click installer (per-user, no admin)
#   uninstall.command   remover
#   INSTALL.md          two-line instructions for the recipient
# =====================================================================
set -euo pipefail
cd "$(dirname "$0")/.."          # repo root

if [ ! -f "Slide Aid.ppam" ]; then
  echo "ERROR: 'Slide Aid.ppam' not found - build the add-in first (README steps 1-6)." >&2
  exit 1
fi
# warn if the add-in looks older than the sources
newest_src=$(ls -t src/*.bas ribbon/customUI14.xml | head -1)
if [ "$newest_src" -nt "Slide Aid.ppam" ]; then
  echo "WARNING: '$newest_src' is newer than 'Slide Aid.ppam'." >&2
  echo "         The zip will ship an OUTDATED add-in - rebuild it first." >&2
  read -r -p "Continue anyway? [y/N] " a
  case "${a:-n}" in [Yy]*) ;; *) exit 1 ;; esac
fi

DIST="dist/Slide Aid"
rm -rf dist
mkdir -p "$DIST"

osacompile -o "$DIST/SlideAidUI.scpt" tools/SlideAidUI.applescript
cp "Slide Aid.ppam" hammerspoon/slideaid.lua "$DIST/"
cp tools/install.command tools/uninstall.command "$DIST/"
chmod +x "$DIST/install.command" "$DIST/uninstall.command"

cat > "$DIST/INSTALL.md" <<'EOF'
# Slide Aid - installation

1. Unzip this folder anywhere.
2. Right-click `install.command` -> Open -> Open. (Plain double-click may be
   blocked by macOS for downloaded files; right-click -> Open bypasses that.)
3. Follow the prompts. Keyboard shortcuts (via the free Hammerspoon utility)
   are optional - you will be asked.
4. Final step in PowerPoint: Tools -> PowerPoint Add-ins... -> tick
   "Slide Aid" -> Enable Macros -> restart PowerPoint.

You then have two new ribbon tabs: **Slide Aid** (align, size, color, text
tools - the "Master" is always the object you selected LAST) and **Chart Aid**
(table-driven charts - click "Sample Slides" there for
live examples).

No admin rights are needed; everything installs into your user account only.
To remove: run `uninstall.command`.
EOF

# ditto preserves the executable bits so .command files stay clickable
ditto -c -k --keepParent "$DIST" "dist/Slide Aid.zip"
echo "Built: dist/Slide Aid.zip ($(du -h "dist/Slide Aid.zip" | cut -f1 | tr -d ' '))"
echo "Share that zip - recipients follow INSTALL.md inside it."
