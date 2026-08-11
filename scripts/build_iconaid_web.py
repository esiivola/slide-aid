#!/usr/bin/env python3
"""Build the full IconAid data from unified-catalog.json:

  apps/powerpoint-iconaid/catalog.json  - for the web sidebar (metadata + normalized
                                           M/L/C/Z paths in a common 24 viewBox)
  apps/powerpoint/data/icons.dat         - for the add-in: "id|name|cat|tags|sub1|sub2|..."
                                           (same normalized paths; used by the gallery
                                           and by the 'Make Editable' convert, keyed by id)
  shared/iconaid/slides/index.json       - for Google Slides: metadata only, embedded in
                                           the sidebar so search stays instant
  shared/iconaid/slides/paths-NN.json    - the same paths, sharded, fetched on demand

All three are derived from the SAME normalized paths, so the sidebar preview, the
inserted PowerPoint freeform and the Google Slides vectors match. Curves are
preserved; arcs finely flattened.

The Slides split exists because an Apps Script sidebar has no HTTP cache to lean
on the way the Office task pane does: shipping 4.6 MB of path data inside the
sidebar would re-download it on every open. Metadata alone is ~1 MB (~130 KB
gzipped) and carries the search, so only the paths actually being previewed are
fetched. Icons are sorted by id before sharding, which makes each shard a
contiguous id range - both sides can then resolve an id to its shard with a
binary search over a 20-odd entry boundary table, with no extra index to ship.
"""
from __future__ import annotations
import json
import re
from pathlib import Path

import svg_normalize as N

# Source packages select the styles we intend to ship. Do not discard generic
# fill/solid suffixes: those are valuable alternatives in the expanded library.
# The only legacy duplicates still filtered are non-regular Phosphor weights.
_PH_WEIGHT = re.compile(r'^phosphor-.+-(thin|light|bold)$')


def is_variant(icon_id: str) -> bool:
    return bool(_PH_WEIGHT.search(icon_id))

ROOT = Path(__file__).resolve().parent.parent
SRC = ROOT / "shared" / "iconaid" / "unified-catalog.json"
WEB_OUT = ROOT / "apps" / "powerpoint-iconaid" / "catalog.json"
DAT_OUT = ROOT / "apps" / "powerpoint" / "data" / "icons.dat"
SLIDES_DIR = ROOT / "shared" / "iconaid" / "slides"

# Icons per Google Slides path shard. ~450 keeps a shard near 200 KB, so one
# screenful of previews is one fetch and one small parse.
SLIDES_SHARD_SIZE = 450

# Bootstrap is drawn as solid regions rather than centreline strokes; the same is
# true of any -solid/-mini variant that survives the variant filter. The renderers
# need to know, because a solid glyph stroked as an outline reads as hollow.
_FILLED_SUFFIX = re.compile(r'-(solid|mini)$')


def is_filled(icon_id: str, source: str, render_mode: str = "") -> bool:
    if render_mode:
        return render_mode == "fill"
    return source == "bootstrap" or bool(_FILLED_SUFFIX.search(icon_id))


def write_slides_catalog(web: list[dict]) -> None:
    """Emit the Google Slides metadata index plus its sharded path data."""
    ordered = sorted(web, key=lambda icon: icon["id"])
    shards = []
    SLIDES_DIR.mkdir(parents=True, exist_ok=True)
    for old in SLIDES_DIR.glob("paths-*.json"):
        old.unlink()

    index = []
    for start in range(0, len(ordered), SLIDES_SHARD_SIZE):
        chunk = ordered[start:start + SLIDES_SHARD_SIZE]
        number = len(shards)
        (SLIDES_DIR / f"paths-{number:02d}.json").write_text(
            json.dumps({icon["id"]: icon["d"] for icon in chunk}, separators=(",", ":"), ensure_ascii=False),
            encoding="utf-8",
        )
        shards.append({"shard": number, "firstId": chunk[0]["id"], "lastId": chunk[-1]["id"]})
        for icon in chunk:
            entry = {"id": icon["id"], "n": icon["n"], "c": icon["c"], "s": icon["s"], "t": icon["t"], "k": number}
            if icon.get("f"):
                entry["f"] = 1
            index.append(entry)

    (SLIDES_DIR / "index.json").write_text(
        json.dumps(
            {
                "schema": 5,
                "viewBox": 24,
                "style": {"stroke": 1.6, "lineCap": "round", "lineJoin": "round"},
                "shardSize": SLIDES_SHARD_SIZE,
                "shards": shards,
                "icons": index,
            },
            separators=(",", ":"),
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )
    index_mb = (SLIDES_DIR / "index.json").stat().st_size / 1024 / 1024
    paths_mb = sum(p.stat().st_size for p in SLIDES_DIR.glob("paths-*.json")) / 1024 / 1024
    print(f"{SLIDES_DIR}  index {index_mb:.2f} MB + {len(shards)} path shards {paths_mb:.2f} MB")

# Legacy native viewBox per source -> scale factor to a common 24-unit grid.
# New sources carry viewBox per icon in their style metadata.
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
        style = icon.get("style", {})
        vb = style.get("viewBox", VIEWBOX.get(src, 24))
        render_mode = style.get("renderMode", "")
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
        tags = " ".join(icon.get("tags", []))
        entry = {"id": iid, "n": name, "c": cate, "s": src, "t": tags, "d": subpaths}
        if is_filled(iid, src, render_mode):
            entry["f"] = 1
        web.append(entry)
        # pipe-delimited line for the add-in (id|name|cat|tags|subpaths...)
        safe = lambda x: str(x).replace("|", "/").replace("\n", " ")
        lines.append("|".join([safe(iid), safe(name), safe(cate), safe(tags)] + subpaths))
        by_source[src] = by_source.get(src, 0) + 1

    WEB_OUT.parent.mkdir(parents=True, exist_ok=True)
    DAT_OUT.parent.mkdir(parents=True, exist_ok=True)
    # compact JSON (no spaces) to keep the download small
    WEB_OUT.write_text(json.dumps(web, separators=(",", ":"), ensure_ascii=False), encoding="utf-8")
    DAT_OUT.write_text("".join(l + "\n" for l in lines), encoding="utf-8")

    write_slides_catalog(web)

    non_ascii = sum(1 for l in lines if not l.isascii())
    print(f"emitted: {len(web)} icons   style-variants dropped: {variants}   skipped (no path): {skipped}")
    print(f"by source: {by_source}")
    print(f"non-ASCII lines (VBA reads ASCII): {non_ascii}")
    print(f"{WEB_OUT}  ({WEB_OUT.stat().st_size/1024/1024:.2f} MB)")
    print(f"{DAT_OUT}  ({DAT_OUT.stat().st_size/1024/1024:.2f} MB)")


if __name__ == "__main__":
    main()
