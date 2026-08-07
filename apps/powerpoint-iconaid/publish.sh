#!/bin/bash
# =====================================================================
# Publish the IconAid sidebar to the gh-pages branch, served at
#   https://esiivola.github.io/slide-aid/
# Re-run whenever the sidebar files or catalog.json change. Uses a
# throwaway worktree so the large generated catalog.json never lands on
# the main branch's history.
# =====================================================================
set -euo pipefail
cd "$(dirname "$0")"                        # apps/powerpoint-iconaid
REPO="$(git rev-parse --show-toplevel)"

for f in taskpane.html taskpane.js catalog.json manifest.xml; do
  [ -f "$f" ] || { echo "ERROR: missing $f (run scripts/build_iconaid_web.py first)." >&2; exit 1; }
done

WT="$(mktemp -d)"
cleanup() { git -C "$REPO" worktree remove --force "$WT" 2>/dev/null || true; }
trap cleanup EXIT

git -C "$REPO" fetch -q origin gh-pages 2>/dev/null || true
if git -C "$REPO" show-ref -q refs/remotes/origin/gh-pages; then
  git -C "$REPO" worktree add -qf "$WT" -B gh-pages origin/gh-pages
else                                        # first publish: empty orphan branch
  git -C "$REPO" worktree add -qf "$WT" --detach
  git -C "$WT" checkout -q --orphan gh-pages
fi

# replace the branch contents with exactly the site files
find "$WT" -mindepth 1 -maxdepth 1 ! -name .git -exec rm -rf {} +
cp taskpane.html taskpane.js catalog.json manifest.xml "$WT/"
cp -R assets "$WT/assets"
touch "$WT/.nojekyll"                        # serve files as-is (no Jekyll)

git -C "$WT" add -A
if git -C "$WT" diff --cached --quiet; then
  echo "gh-pages already up to date."
else
  git -C "$WT" commit -q -m "Publish IconAid sidebar"
  git -C "$WT" push -q origin gh-pages
  echo "Published. Live at https://esiivola.github.io/slide-aid/ in ~1 min."
fi
