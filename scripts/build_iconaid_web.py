#!/usr/bin/env python3
"""Build the full (~10.7k) IconAid data from unified-catalog.json:

  apps/powerpoint-iconaid/catalog.json  - for the web sidebar (metadata + normalized
                                           M/L/C/Z paths in a common 24 viewBox)
  apps/powerpoint/data/icons.dat         - for the add-in: "id|name|cat|tags|sub1|sub2|..."
                                           (same normalized paths; used by the gallery
                                           and by the 'Make Editable' convert, keyed by id)

Both are derived from the SAME normalized paths, so the sidebar preview and the
inserted PowerPoint freeform match. Curves are preserved; arcs finely flattened.
"""
from __future__ import annotations
import json
import re
from pathlib import Path

import svg_normalize as N

# Keep one consistent outline/regular style so every icon reads with the same
# 1.6 centerline stroke. Fill/solid/duotone paths are solid regions (stroking
# them looks thin/hollow); phosphor thin/light/bold are different weights.
# These are id suffixes, so legitimate names (e.g. "traffic-light") aren't hit
# unless the suffix is a real style variant.
_VARIANT = re.compile(r'-(fill|filled|solid|duotone)$')
_PH_WEIGHT = re.compile(r'^phosphor-.+-(thin|light|bold)$')


def is_variant(icon_id: str) -> bool:
    return bool(_VARIANT.search(icon_id) or _PH_WEIGHT.search(icon_id))

ROOT = Path(__file__).resolve().parent.parent
SRC = ROOT / "shared" / "iconaid" / "unified-catalog.json"
WEB_OUT = ROOT / "apps" / "powerpoint-iconaid" / "catalog.json"
DAT_OUT = ROOT / "apps" / "powerpoint" / "data" / "icons.dat"

# Native viewBox per source -> scale factor to a common 24-unit grid.
VIEWBOX = {"tabler": 24, "lucide": 24, "heroicons": 24, "bootstrap": 16, "phosphor": 256}


def main():
    cat = json.load(open(SRC, encoding="utf-8"))
    icons = cat["icons"]
    print(f"source icons: {len(icons)}")

    web, lines = [], []
    skipped = 0
    variants = 0
    by_source = {}
    for icon in icons:
        iid = icon["id"]
        if is_variant(iid):
            variants += 1
            continue
        src = icon.get("source", "")
        vb = VIEWBOX.get(src, 24)
        scale = 24.0 / vb
        subpaths = []
        for d in icon.get("paths", []):
            try:
                subpaths += N.normalize_field(d, scale)
            except Exception:
                pass
        if not subpaths:
            skipped += 1
            continue
        name = icon["name"]
        cate = icon.get("category", "General")
        tags = " ".join(icon.get("tags", [])[:12])
        web.append({"id": iid, "n": name, "c": cate, "s": src, "t": tags, "d": subpaths})
        # pipe-delimited line for the add-in (id|name|cat|tags|subpaths...)
        safe = lambda x: str(x).replace("|", "/").replace("\n", " ")
        lines.append("|".join([safe(iid), safe(name), safe(cate), safe(tags)] + subpaths))
        by_source[src] = by_source.get(src, 0) + 1

    WEB_OUT.parent.mkdir(parents=True, exist_ok=True)
    DAT_OUT.parent.mkdir(parents=True, exist_ok=True)
    # compact JSON (no spaces) to keep the download small
    WEB_OUT.write_text(json.dumps(web, separators=(",", ":"), ensure_ascii=False), encoding="utf-8")
    DAT_OUT.write_text("".join(l + "\n" for l in lines), encoding="utf-8")

    non_ascii = sum(1 for l in lines if not l.isascii())
    print(f"emitted: {len(web)} icons   style-variants dropped: {variants}   skipped (no path): {skipped}")
    print(f"by source: {by_source}")
    print(f"non-ASCII lines (VBA reads ASCII): {non_ascii}")
    print(f"{WEB_OUT}  ({WEB_OUT.stat().st_size/1024/1024:.2f} MB)")
    print(f"{DAT_OUT}  ({DAT_OUT.stat().st_size/1024/1024:.2f} MB)")


if __name__ == "__main__":
    main()
