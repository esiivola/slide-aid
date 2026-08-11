#!/usr/bin/env python3
"""
Normalize all external icons into a unified IconAid catalog.

This script:
1. Loads all icons from external sources (Tabler, Lucide, Heroicons, Phosphor, Bootstrap)
2. Normalizes SVG paths to 24×24 viewBox
3. Standardizes stroke properties for consistent rendering
4. Creates a unified catalog where all icons can be rendered with customizable:
   - Line color
   - Stroke width
   - Export to SVG/PNG

Output: A single unified catalog with all icons as first-class IconAid citizens.
"""
from __future__ import annotations

import json
import hashlib
from pathlib import Path
from typing import Any

from enrich_icon_tags import enrich_icon_tags

ROOT = Path(__file__).resolve().parent.parent
EXTERNAL_DIR = ROOT / "shared" / "iconaid" / "external-sources"
OUTPUT_PATH = ROOT / "shared" / "iconaid" / "unified-catalog.json"
SOURCE_MANIFEST = ROOT / "shared" / "iconaid" / "sources.json"

# Standard IconAid style
ICONAID_STYLE = {
    "viewBox": 24,
    "strokeWidth": 1.6,
    "strokeLinecap": "round",
    "strokeLinejoin": "round",
    "fill": "none",
}


def load_source_file(filename: str) -> dict:
    """Load a normalized source file."""
    path = EXTERNAL_DIR / filename
    if not path.exists():
        return {"icons": []}
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def normalize_svg_path(path_d: str) -> str:
    """
    Normalize an SVG path string.
    - Trim whitespace
    - Normalize spacing around commands
    """
    if not path_d:
        return ""
    # Normalize whitespace
    path_d = " ".join(path_d.split())
    return path_d


def extract_path_bounds(paths: list[str]) -> tuple[float, float, float, float]:
    """
    Extract approximate bounds from SVG paths.
    Returns (min_x, min_y, max_x, max_y).
    
    This is a simplified parser - assumes paths are already on 24x24 grid.
    """
    # Most icons from these libraries are already on 24x24
    # We'll assume standard bounds
    return (0, 0, 24, 24)


def create_unified_icon(
    source_icon: dict,
    source: dict[str, Any],
    source_version: str = "",
) -> dict:
    """
    Create a unified IconAid icon from a source icon.
    """
    # Generate unique ID
    original_id = source_icon.get("id", source_icon.get("name", "unknown"))
    source_name = source["id"]
    render_mode = source_icon.get("renderMode", source.get("renderMode", "stroke"))
    icon_id = f"{source_name}-{original_id}".lower().replace(" ", "-").replace("_", "-")
    
    # Clean up duplicate dashes
    while "--" in icon_id:
        icon_id = icon_id.replace("--", "-")
    # All hosts have legacy-compatible filled-icon detection based on an id
    # suffix. Preserve that contract while carrying explicit renderMode too.
    if render_mode == "fill" and source_name != "bootstrap" and not icon_id.endswith(("-solid", "-mini")):
        icon_id += "-solid"
    elif render_mode != "fill" and source_name != "bootstrap" and icon_id.endswith(("-solid", "-mini")):
        # A few outline packs use "solid" as part of the concept name (for
        # example mail-out-solid) even though the SVG is a centerline stroke.
        # Avoid colliding with the legacy cross-host fill suffix contract.
        icon_id += "-outline"
    if len(icon_id) > 64:
        digest = hashlib.sha1(icon_id.encode("utf-8")).hexdigest()[:8]
        # The legacy VBA/editable paths infer filled geometry from the suffix.
        # Keep it at the very end even when long IDs require truncation.
        suffix = next(
            (value for value in ("-solid", "-mini") if render_mode == "fill" and icon_id.endswith(value)),
            "",
        )
        stem = icon_id[:-len(suffix)] if suffix else icon_id
        room = 64 - len(suffix) - len(digest) - 1
        icon_id = f"{stem[:room].rstrip('-')}-{digest}{suffix}"
    
    # Get SVG paths
    svg_paths = source_icon.get("svg_paths", [])
    if not svg_paths:
        return None
    
    # Normalize paths
    normalized_paths = [normalize_svg_path(p) for p in svg_paths if p]
    if not normalized_paths:
        return None
    
    # Get metadata
    name = source_icon.get("name", original_id)
    category = source_icon.get("category", "General")
    tags = list(dict.fromkeys(str(tag).strip().lower() for tag in source_icon.get("tags", []) if str(tag).strip()))
    
    # Build searchable text
    searchable = " ".join([
        name.lower(),
        category.lower(),
        " ".join(tags[:20]),
        source_name,
    ])
    
    icon = {
        "id": icon_id,
        "name": name,
        "category": category,
        "tags": tags,
        "source": source_name,
        "sourceName": source["name"],
        "sourceVersion": source_version,
        "sourceUrl": source["upstream"],
        "license": source["license"],
        "paths": normalized_paths,
        "searchable": searchable,
        "style": {
            "viewBox": source_icon.get("viewBox", source.get("viewBox", 24)),
            "renderMode": render_mode,
            "defaultStroke": 1.6,
            "linecap": "round",
            "linejoin": "round",
        },
    }
    return enrich_icon_tags(icon)


def load_all_sources() -> list[dict]:
    """Load and normalize all icons from all sources."""
    manifest = json.loads(SOURCE_MANIFEST.read_text(encoding="utf-8"))
    sources = [source for source in manifest["sources"] if source.get("enabled")]
    
    all_icons = []
    source_counts = {}
    
    for source in sources:
        filename = source["file"]
        source_name = source["id"]
        print(f"Loading {source_name}...")
        data = load_source_file(filename)
        icons = data.get("icons", [])
        source_version = data.get("version", "")
        
        count = 0
        for source_icon in icons:
            unified = create_unified_icon(source_icon, source, source_version)
            if unified:
                all_icons.append(unified)
                count += 1
        
        source_counts[source_name] = count
        print(f"  Loaded {count} icons from {source_name}")
    
    return all_icons, source_counts


def deduplicate_icons(icons: list[dict]) -> list[dict]:
    """
    Preserve every icon while making normalized ID collisions deterministic.

    The library deliberately keeps semantically equivalent icons from different
    families. A source can also contain names that normalize to the same legacy
    fill-suffixed ID (for example ``foo`` and ``foo-solid``). In that rare case,
    keep both and add a stable hash before the legacy fill suffix.
    """
    by_id: dict[str, dict] = {}
    for icon in icons:
        icon_id = icon["id"]
        if icon_id in by_id:
            suffix = next((value for value in ("-solid", "-mini") if icon_id.endswith(value)), "")
            stem = icon_id[:-len(suffix)] if suffix else icon_id
            seed = f"{icon.get('source', '')}:{icon.get('name', '')}:{icon_id}"
            digest = hashlib.sha1(seed.encode("utf-8")).hexdigest()[:8]
            room = 64 - len(suffix) - len(digest) - 1
            icon_id = f"{stem[:room].rstrip('-')}-{digest}{suffix}"
            attempt = 1
            while icon_id in by_id:
                digest = hashlib.sha1(f"{seed}:{attempt}".encode("utf-8")).hexdigest()[:8]
                icon_id = f"{stem[:room].rstrip('-')}-{digest}{suffix}"
                attempt += 1
            icon = dict(icon)
            icon["id"] = icon_id
            enrich_icon_tags(icon)
        by_id[icon_id] = icon
    return list(by_id.values())


def build_category_index(icons: list[dict]) -> dict[str, list[str]]:
    """Build an index of icons by category."""
    index = {}
    for icon in icons:
        cat = icon["category"]
        if cat not in index:
            index[cat] = []
        index[cat].append(icon["id"])
    return index


def build_tag_index(icons: list[dict]) -> dict[str, list[str]]:
    """Build an index of icons by tag."""
    index = {}
    for icon in icons:
        for tag in icon.get("tags", []):
            tag_lower = tag.lower()
            if tag_lower not in index:
                index[tag_lower] = []
            if icon["id"] not in index[tag_lower]:
                index[tag_lower].append(icon["id"])
    return index


def main():
    print("=" * 60)
    print("NORMALIZING ALL ICONS FOR ICONAID")
    print("=" * 60)
    print()
    
    # Load all sources
    all_icons, source_counts = load_all_sources()
    print(f"\nTotal icons loaded: {len(all_icons)}")
    
    # Deduplicate
    print("\nDeduplicating icons...")
    unique_icons = deduplicate_icons(all_icons)
    print(f"Unique icons after deduplication: {len(unique_icons)}")
    
    # Sort by name
    unique_icons.sort(key=lambda x: (x["category"], x["name"].lower()))
    
    # Build indexes
    print("\nBuilding search indexes...")
    category_index = build_category_index(unique_icons)
    tag_index = build_tag_index(unique_icons)
    
    # Analyze categories
    print(f"\nCategories: {len(category_index)}")
    cat_counts = {cat: len(ids) for cat, ids in category_index.items()}
    for cat, count in sorted(cat_counts.items(), key=lambda x: -x[1])[:15]:
        print(f"  {cat}: {count}")
    
    # Build unified catalog
    catalog = {
        "schema": 4,
        "name": "IconAid Unified Catalog",
        "description": "Permissively licensed icons normalized for offline IconAid search and rendering",
        "style": ICONAID_STYLE,
        "sources": source_counts,
        "totalIcons": len(unique_icons),
        "categories": list(category_index.keys()),
        "categoryIndex": category_index,
        "tagIndex": tag_index,
        "icons": unique_icons,
    }
    
    # Save
    print(f"\nSaving unified catalog to {OUTPUT_PATH}...")
    with open(OUTPUT_PATH, "w", encoding="utf-8") as f:
        json.dump(catalog, f, indent=2)
    
    file_size = OUTPUT_PATH.stat().st_size / (1024 * 1024)
    print(f"Saved! File size: {file_size:.2f} MB")
    
    print("\n" + "=" * 60)
    print("SUMMARY")
    print("=" * 60)
    print(f"Total normalized icons: {len(unique_icons)}")
    print(f"Categories: {len(category_index)}")
    print(f"Unique tags: {len(tag_index)}")
    print("\nAll icons now support:")
    print("  ✓ Customizable monochrome color")
    print("  ✓ Explicit stroke/fill rendering")
    print("  ✓ SVG export with any style")
    print("  ✓ PNG export at any resolution")
    print("  ✓ Offline usage")


if __name__ == "__main__":
    main()
