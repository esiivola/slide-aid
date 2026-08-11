#!/usr/bin/env python3
"""
Build a combined search index from all external icon sources.

This creates a unified searchable database that can be queried for icon candidates.

Usage:
    python3 scripts/build_icon_search_index.py
    python3 scripts/build_icon_search_index.py --search "chart analytics"
    python3 scripts/build_icon_search_index.py --category "Finance"
    python3 scripts/build_icon_search_index.py --list-categories

Output:
    shared/iconaid/external-sources/combined-search-index.json
"""
from __future__ import annotations

import argparse
import json
import re
from pathlib import Path
from typing import Any

from enrich_icon_tags import enrich_icon_tags

ROOT = Path(__file__).resolve().parent.parent
SOURCES_DIR = ROOT / "shared" / "iconaid" / "external-sources"
OUTPUT_FILE = SOURCES_DIR / "combined-search-index.json"
SOURCE_MANIFEST = ROOT / "shared" / "iconaid" / "sources.json"


def load_source(filepath: Path) -> list[dict[str, Any]]:
    """Load icons from a source file."""
    if not filepath.exists():
        return []
    
    with open(filepath, "r", encoding="utf-8") as f:
        data = json.load(f)
    
    return data.get("icons", [])


def build_combined_index() -> dict[str, Any]:
    """Build combined search index from all sources."""
    all_icons = []
    source_counts = {}
    
    manifest = json.loads(SOURCE_MANIFEST.read_text(encoding="utf-8"))
    for source in manifest["sources"]:
        if not source.get("enabled"):
            continue
        icons = load_source(SOURCES_DIR / source["file"])
        for icon in icons:
            icon["source_id"] = f"{source['id']}:{icon['id']}"
            enrich_icon_tags(icon)
        all_icons.extend(icons)
        source_counts[source["id"]] = len(icons)

    # Preserve stylistic/source alternatives; only duplicate source IDs collapse.
    unique_icons = list({icon["source_id"]: icon for icon in all_icons}.values())
    
    # Sort by name
    unique_icons.sort(key=lambda x: x["name"].lower())
    
    # Build category index
    categories = {}
    for icon in unique_icons:
        cat = icon.get("category", "Unknown")
        if cat not in categories:
            categories[cat] = []
        categories[cat].append(icon["source_id"])
    
    # Build tag index
    tag_index = {}
    for icon in unique_icons:
        for tag in icon.get("tags", []):
            tag_lower = tag.lower()
            if tag_lower not in tag_index:
                tag_index[tag_lower] = []
            tag_index[tag_lower].append(icon["source_id"])
    
    return {
        "total_icons": len(unique_icons),
        "source_counts": source_counts,
        "categories": sorted(categories.keys()),
        "category_index": categories,
        "tag_index": tag_index,
        "icons": unique_icons,
    }


def search_icons(index: dict[str, Any], query: str) -> list[dict[str, Any]]:
    """Search icons by query."""
    query_lower = query.lower()
    query_terms = query_lower.split()
    
    results = []
    for icon in index["icons"]:
        searchable = icon.get("searchable", "").lower()
        name_lower = icon["name"].lower()
        
        # Calculate relevance score
        score = 0
        
        # Exact name match
        if query_lower == name_lower:
            score += 100
        
        # Name contains query
        elif query_lower in name_lower:
            score += 50
        
        # All terms match in searchable
        all_terms_match = all(term in searchable for term in query_terms)
        if all_terms_match:
            score += 25
        
        # Any term matches
        any_term_match = any(term in searchable for term in query_terms)
        if any_term_match:
            score += 10
        
        if score > 0:
            results.append((score, icon))
    
    # Sort by score (descending)
    results.sort(key=lambda x: (-x[0], x[1]["name"]))
    
    return [icon for _, icon in results]


def filter_by_category(index: dict[str, Any], category: str) -> list[dict[str, Any]]:
    """Filter icons by category."""
    category_lower = category.lower()
    
    results = []
    for icon in index["icons"]:
        if icon.get("category", "").lower() == category_lower:
            results.append(icon)
    
    results.sort(key=lambda x: x["name"])
    return results


def print_icon(icon: dict[str, Any], verbose: bool = False) -> None:
    """Print icon information."""
    print(f"  {icon['source_id']}")
    print(f"    Name: {icon['name']}")
    print(f"    Category: {icon.get('category', 'Unknown')}")
    print(f"    License: {icon.get('license', 'Unknown')}")
    if verbose:
        print(f"    Tags: {', '.join(icon.get('tags', []))}")
        paths = icon.get("svg_paths", [])
        if paths:
            print(f"    SVG elements: {len(paths)}")


def main():
    parser = argparse.ArgumentParser(description="Build and search icon index")
    parser.add_argument("--search", "-s", help="Search query")
    parser.add_argument("--category", "-c", help="Filter by category")
    parser.add_argument("--list-categories", action="store_true", help="List all categories")
    parser.add_argument("--limit", "-l", type=int, default=20, help="Limit results")
    parser.add_argument("--verbose", "-v", action="store_true", help="Verbose output")
    parser.add_argument("--rebuild", "-r", action="store_true", help="Force rebuild index")
    args = parser.parse_args()
    
    # Build or load index
    if args.rebuild or not OUTPUT_FILE.exists():
        print("Building combined search index...")
        index = build_combined_index()
        
        SOURCES_DIR.mkdir(parents=True, exist_ok=True)
        with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
            json.dump(index, f, indent=2, ensure_ascii=False)
        
        print(f"Written to: {OUTPUT_FILE}")
        print(f"Total unique icons: {index['total_icons']}")
        print(f"Source counts: {index['source_counts']}")
        print(f"Categories: {len(index['categories'])}")
    else:
        with open(OUTPUT_FILE, "r", encoding="utf-8") as f:
            index = json.load(f)
    
    # List categories
    if args.list_categories:
        print("\nCategories:")
        for cat in sorted(index["categories"]):
            count = len(index["category_index"].get(cat, []))
            print(f"  {cat}: {count} icons")
        return
    
    # Search
    if args.search:
        results = search_icons(index, args.search)
        print(f"\nSearch results for '{args.search}': {len(results)} matches")
        for icon in results[:args.limit]:
            print_icon(icon, args.verbose)
        if len(results) > args.limit:
            print(f"\n  ... and {len(results) - args.limit} more results")
        return
    
    # Filter by category
    if args.category:
        results = filter_by_category(index, args.category)
        print(f"\nCategory '{args.category}': {len(results)} icons")
        for icon in results[:args.limit]:
            print_icon(icon, args.verbose)
        if len(results) > args.limit:
            print(f"\n  ... and {len(results) - args.limit} more icons")
        return
    
    # Default: show summary
    print(f"\nCombined Icon Search Index")
    print(f"=" * 40)
    print(f"Total unique icons: {index['total_icons']}")
    print(f"Sources: {index['source_counts']}")
    print(f"\nTop categories:")
    cat_counts = [(cat, len(ids)) for cat, ids in index["category_index"].items()]
    cat_counts.sort(key=lambda x: -x[1])
    for cat, count in cat_counts[:15]:
        print(f"  {cat}: {count}")
    
    print(f"\nUsage:")
    print(f"  --search 'chart analytics'  Search for icons")
    print(f"  --category Finance          Filter by category")
    print(f"  --list-categories           List all categories")
    print(f"  --verbose                   Show more details")


if __name__ == "__main__":
    main()
