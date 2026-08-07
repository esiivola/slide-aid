#!/usr/bin/env python3
"""
Unified icon search across IconAid native icons and external references.

This script provides a single search interface for discovering icons from:
- IconAid native icons (70+ hand-crafted, ready to use)
- External references (1000+ from Tabler, Lucide, Heroicons, Phosphor, Bootstrap)

Usage:
    python3 scripts/search_icons.py "analytics"
    python3 scripts/search_icons.py "sustainability" --limit 20
    python3 scripts/search_icons.py "chart" --source native
    python3 scripts/search_icons.py "mail" --source external
"""
from __future__ import annotations

import argparse
import json
import re
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
CATALOG_PATH = ROOT / "shared" / "iconaid" / "catalog.json"
ENRICHED_CATALOG_PATH = ROOT / "shared" / "iconaid" / "catalog-with-external.json"


def load_catalogs():
    """Load both the native catalog and enriched catalog."""
    with open(CATALOG_PATH, "r", encoding="utf-8") as f:
        native_catalog = json.load(f)
    
    enriched_catalog = None
    if ENRICHED_CATALOG_PATH.exists():
        with open(ENRICHED_CATALOG_PATH, "r", encoding="utf-8") as f:
            enriched_catalog = json.load(f)
    
    return native_catalog, enriched_catalog


def normalize_query(query: str) -> list[str]:
    """Normalize search query into terms."""
    # Convert to lowercase and split on whitespace
    terms = query.lower().split()
    # Remove very short terms
    terms = [t for t in terms if len(t) > 1]
    return terms


def score_match(icon: dict, terms: list[str], source_type: str) -> int:
    """
    Score how well an icon matches the search terms.
    Higher score = better match.
    """
    score = 0
    name_lower = icon.get("name", "").lower()
    id_lower = icon.get("id", "").lower()
    category_lower = icon.get("category", "").lower()
    tags = [t.lower() for t in icon.get("tags", [])]
    aliases = [a.lower() for a in icon.get("aliases", [])]
    searchable = icon.get("searchable", "").lower()
    
    for term in terms:
        # Exact name match (highest priority)
        if term == name_lower or term == id_lower:
            score += 100
        # Name contains term
        elif term in name_lower:
            score += 50
        # ID contains term
        elif term in id_lower:
            score += 40
        # Alias exact match
        elif term in aliases:
            score += 35
        # Category match
        elif term == category_lower:
            score += 30
        # Tag exact match
        elif term in tags:
            score += 25
        # Searchable text contains term
        elif term in searchable:
            score += 10
        # Partial match in tags
        elif any(term in tag for tag in tags):
            score += 5
    
    # Bonus for native icons (they're ready to use)
    if source_type == "native":
        score += 20
    
    return score


def search_icons(
    query: str,
    native_catalog: dict,
    enriched_catalog: dict | None,
    source: str = "all",
    limit: int = 30,
) -> list[tuple[int, str, dict]]:
    """
    Search icons across all sources.
    
    Returns list of (score, source_type, icon) tuples.
    """
    terms = normalize_query(query)
    if not terms:
        return []
    
    results = []
    
    # Search native icons
    if source in ("all", "native"):
        for icon in native_catalog.get("icons", []):
            score = score_match(icon, terms, "native")
            if score > 0:
                results.append((score, "native", icon))
    
    # Search external references
    if source in ("all", "external") and enriched_catalog:
        ext_refs = enriched_catalog.get("externalReferences", {}).get("icons", [])
        for icon in ext_refs:
            score = score_match(icon, terms, "external")
            if score > 0:
                results.append((score, "external", icon))
    
    # Sort by score (descending), then by name
    results.sort(key=lambda x: (-x[0], x[2]["name"]))
    
    return results[:limit]


def format_native_icon(icon: dict, score: int, verbose: bool) -> str:
    """Format a native IconAid icon for display."""
    lines = [
        f"\n  🎯 [{score:3d}] {icon['name']} (id: {icon['id']})",
        f"         Category: {icon['category']}",
        f"         Status: ✅ Ready to use",
    ]
    if verbose:
        if icon.get("aliases"):
            lines.append(f"         Aliases: {', '.join(icon['aliases'][:5])}")
        if icon.get("tags"):
            lines.append(f"         Tags: {', '.join(icon['tags'][:8])}")
        primitives = icon.get("primitives", [])
        lines.append(f"         Primitives: {len(primitives)} elements")
    return "\n".join(lines)


def format_external_icon(icon: dict, score: int, verbose: bool) -> str:
    """Format an external reference icon for display."""
    source = icon.get("source", "unknown")
    license_text = icon.get("license", "Unknown")
    
    lines = [
        f"\n  📎 [{score:3d}] {icon['name']} (from {source})",
        f"         Category: {icon['category']}",
        f"         License: {license_text}",
        f"         Status: ⚠️ External reference (needs conversion)",
    ]
    if verbose:
        if icon.get("tags"):
            lines.append(f"         Tags: {', '.join(icon['tags'][:8])}")
        svg_paths = icon.get("svgPaths", [])
        lines.append(f"         SVG paths: {len(svg_paths)} elements")
        lines.append(f"         Source ID: {icon.get('sourceId', 'N/A')}")
    return "\n".join(lines)


def main():
    parser = argparse.ArgumentParser(
        description="Search IconAid icons and external references"
    )
    parser.add_argument("query", help="Search query (e.g., 'analytics', 'chart data')")
    parser.add_argument(
        "--source", "-s",
        choices=["all", "native", "external"],
        default="all",
        help="Search source: all, native (IconAid only), or external"
    )
    parser.add_argument("--limit", "-l", type=int, default=20, help="Maximum results")
    parser.add_argument("--verbose", "-v", action="store_true", help="Show detailed info")
    args = parser.parse_args()
    
    native_catalog, enriched_catalog = load_catalogs()
    
    print(f"Searching for: '{args.query}'")
    print(f"Source: {args.source}")
    print("-" * 60)
    
    results = search_icons(
        args.query,
        native_catalog,
        enriched_catalog,
        source=args.source,
        limit=args.limit,
    )
    
    if not results:
        print("No icons found matching your query.")
        return
    
    # Count by type
    native_count = sum(1 for _, t, _ in results if t == "native")
    external_count = sum(1 for _, t, _ in results if t == "external")
    
    print(f"Found {len(results)} icons ({native_count} native, {external_count} external)")
    
    for score, source_type, icon in results:
        if source_type == "native":
            print(format_native_icon(icon, score, args.verbose))
        else:
            print(format_external_icon(icon, score, args.verbose))
    
    if len(results) == args.limit:
        print(f"\n... showing first {args.limit} results. Use --limit to see more.")


if __name__ == "__main__":
    main()
