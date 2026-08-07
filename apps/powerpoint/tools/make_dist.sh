#!/bin/bash
# =====================================================================
# Slide Aid - build the distributable installer zip for colleagues.
#
# Run on YOUR Mac (needs osacompile, i.e. any macOS) after building
# the add-in (README steps 1-6):
#
#   cd apps/powerpoint
#   ./tools/make_dist.sh
#
# Produces  dist/Slide Aid.zip  containing:
#   Slide Aid.ppam      the add-in (with ribbon + icons)
#   SlideAidUI.scpt     pre-compiled native-dialogs helper (Chart Aid
#                       settings/colors panels + color picker)
#   slideaid.lua        Hammerspoon shortcut config
#   install.command     double-click installer (per-user, no admin)
#   uninstall.command   remover
#   INSTALL.md          two-line instructions for the recipient
# =====================================================================
set -euo pipefail
cd "$(dirname "$0")/.."          # PowerPoint app root

ADDIN="dist/Slide Aid.ppam"
if [ ! -f "$ADDIN" ]; then
  echo "ERROR: '$ADDIN' not found - build the add-in first (README steps 1-6)." >&2
  exit 1
fi
ICONDATA="data/icons.dat"
if [ ! -f "$ICONDATA" ]; then
  echo "ERROR: '$ICONDATA' not found - run scripts/build_iconaid_web.py first." >&2
  exit 1
fi
MANIFEST="../powerpoint-iconaid/manifest.xml"
if [ ! -f "$MANIFEST" ]; then
  echo "ERROR: '$MANIFEST' not found - the IconAid sidebar manifest is required." >&2
  exit 1
fi
# warn if the add-in looks older than the sources
newest_src=$(ls -t src/*.bas ribbon/customUI14.xml | head -1)
if [ "$newest_src" -nt "$ADDIN" ]; then
  echo "WARNING: '$newest_src' is newer than '$ADDIN'." >&2
  echo "         The zip will ship an OUTDATED add-in - rebuild it first." >&2
  read -r -p "Continue anyway? [y/N] " a
  case "${a:-n}" in [Yy]*) ;; *) exit 1 ;; esac
fi

DIST="dist/Slide Aid"
rm -rf "$DIST"
mkdir -p "$DIST"

osacompile -o "$DIST/SlideAidUI.scpt" tools/SlideAidUI.applescript
cp "$ADDIN" "$ICONDATA" "$MANIFEST" hammerspoon/slideaid.lua "$DIST/"
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

**Icons:** on the **Insert** tab, click **Insert Icons** to search a sidebar of
10,000+ vector icons; click one to drop it on the slide, then select it and click
**Make Editable** (also on the Insert tab) to turn it into an editable shape.
(The sidebar loads its catalog online, so it needs an internet connection.)

No admin rights are needed; everything installs into your user account only.
To remove: run `uninstall.command`.
EOF

# ditto preserves the executable bits so .command files stay clickable
ditto -c -k --norsrc --keepParent "$DIST" "dist/Slide Aid.zip"
echo "Built: dist/Slide Aid.zip ($(du -h "dist/Slide Aid.zip" | cut -f1 | tr -d ' '))"
echo "Share that zip - recipients follow INSTALL.md inside it."
