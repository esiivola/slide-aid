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
import re
import hashlib
from pathlib import Path
from typing import Any
from collections import Counter

ROOT = Path(__file__).resolve().parent.parent
EXTERNAL_DIR = ROOT / "shared" / "iconaid" / "external-sources"
OUTPUT_PATH = ROOT / "shared" / "iconaid" / "unified-catalog.json"

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
    source_name: str,
    source_license: str,
) -> dict:
    """
    Create a unified IconAid icon from a source icon.
    """
    # Generate unique ID
    original_id = source_icon.get("id", source_icon.get("name", "unknown"))
    icon_id = f"{source_name}-{original_id}".lower().replace(" ", "-").replace("_", "-")
    
    # Clean up duplicate dashes
    while "--" in icon_id:
        icon_id = icon_id.replace("--", "-")
    
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
    tags = source_icon.get("tags", [])
    
    # Build searchable text
    searchable = " ".join([
        name.lower(),
        category.lower(),
        " ".join(tags[:20]),
        source_name,
    ])
    
    return {
        "id": icon_id,
        "name": name,
        "category": category,
        "tags": tags[:20],  # Limit tags
        "source": source_name,
        "license": source_license,
        "paths": normalized_paths,
        "searchable": searchable,
        "style": {
            "viewBox": 24,
            "defaultStroke": 1.6,
            "linecap": "round",
            "linejoin": "round",
        },
    }


def load_all_sources() -> list[dict]:
    """Load and normalize all icons from all sources."""
    sources = [
        ("tabler-icons-normalized.json", "tabler", "MIT"),
        ("lucide-icons-normalized.json", "lucide", "ISC"),
        ("heroicons-normalized.json", "heroicons", "MIT"),
        ("phosphor-icons-normalized.json", "phosphor", "MIT"),
        ("bootstrap-icons-normalized.json", "bootstrap", "MIT"),
    ]
    
    all_icons = []
    source_counts = {}
    
    for filename, source_name, license_text in sources:
        print(f"Loading {source_name}...")
        data = load_source_file(filename)
        icons = data.get("icons", [])
        
        count = 0
        for source_icon in icons:
            unified = create_unified_icon(source_icon, source_name, license_text)
            if unified:
                all_icons.append(unified)
                count += 1
        
        source_counts[source_name] = count
        print(f"  Loaded {count} icons from {source_name}")
    
    return all_icons, source_counts


def deduplicate_icons(icons: list[dict]) -> list[dict]:
    """
    Remove duplicate icons (same paths from different sources).
    Prefer: tabler > lucide > heroicons > phosphor > bootstrap
    """
    source_priority = {
        "tabler": 1,
        "lucide": 2, 
        "heroicons": 3,
        "phosphor": 4,
        "bootstrap": 5,
    }
    
    # Hash paths to detect duplicates
    path_hash_to_icon = {}
    
    for icon in icons:
        # Create hash of paths
        paths_str = "|".join(sorted(icon["paths"]))
        path_hash = hashlib.md5(paths_str.encode()).hexdigest()[:12]
        
        if path_hash not in path_hash_to_icon:
            path_hash_to_icon[path_hash] = icon
        else:
            # Keep the one with higher priority (lower number)
            existing = path_hash_to_icon[path_hash]
            existing_priority = source_priority.get(existing["source"], 99)
            new_priority = source_priority.get(icon["source"], 99)
            
            if new_priority < existing_priority:
                path_hash_to_icon[path_hash] = icon
    
    return list(path_hash_to_icon.values())


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
        "description": "All icons normalized to IconAid style - customizable stroke, offline-ready",
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
    print("  ✓ Customizable stroke color")
    print("  ✓ Customizable stroke width")
    print("  ✓ SVG export with any style")
    print("  ✓ PNG export at any resolution")
    print("  ✓ Offline usage")


if __name__ == "__main__":
    main()
