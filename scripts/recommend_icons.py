#!/usr/bin/env python3
"""
Recommend icons for IconAid expansion from the external sources.

This script suggests icons that would be good candidates for IconAid based on:
- Business/consulting relevance
- Simplicity (few SVG elements)
- Not already in IconAid catalog

Usage:
    python3 scripts/recommend_icons.py --category Finance --limit 20
    python3 scripts/recommend_icons.py --search "sustainability ESG"
    python3 scripts/recommend_icons.py --expansion-batch 6  # Communication/Document
    python3 scripts/recommend_icons.py --expansion-batch 7  # ESG
"""
from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parent.parent
INDEX_FILE = ROOT / "shared" / "iconaid" / "external-sources" / "combined-search-index.json"
CATALOG_FILE = ROOT / "shared" / "iconaid" / "catalog.json"

# Icons already in IconAid (to exclude from recommendations)
def get_existing_icon_names() -> set[str]:
    """Get names of icons already in IconAid."""
    if not CATALOG_FILE.exists():
        return set()
    
    with open(CATALOG_FILE, "r", encoding="utf-8") as f:
        data = json.load(f)
    
    names = set()
    for icon in data.get("icons", []):
        # Normalize name for comparison
        name = icon.get("name", "").lower().replace(" ", "-")
        names.add(name)
        # Also add ID
        icon_id = icon.get("id", "").lower()
        names.add(icon_id)
    
    return names


# Expansion batch definitions from EXPANSION_PLAN.md
EXPANSION_BATCHES = {
    6: {
        "name": "Communication and Document",
        "concepts": [
            "presentation", "mail", "chat", "notification", "video",
            "microphone", "send", "knowledge", "folder", "clipboard"
        ],
        "categories": ["Communication", "Document", "Media"],
    },
    7: {
        "name": "ESG",
        "concepts": [
            "solar", "wind", "water", "circular economy", "battery",
            "waste", "climate", "biodiversity", "reporting", "leaf",
            "plant", "tree", "recycle", "energy", "sustainability"
        ],
        "categories": ["ESG", "Nature", "Weather"],
    },
}


def load_index() -> dict[str, Any]:
    """Load the combined search index."""
    with open(INDEX_FILE, "r", encoding="utf-8") as f:
        return json.load(f)


def calculate_simplicity_score(icon: dict[str, Any]) -> int:
    """Calculate simplicity score (higher = simpler, better for IconAid)."""
    svg_paths = icon.get("svg_paths", [])
    num_elements = len(svg_paths)
    
    # Ideal is 4-10 elements
    if 4 <= num_elements <= 10:
        return 100
    elif num_elements < 4:
        return 80  # Too simple might be boring
    elif num_elements <= 16:
        return 60  # Acceptable but complex
    else:
        return 20  # Too complex


def search_icons(index: dict[str, Any], query: str, existing: set[str]) -> list[tuple[int, dict[str, Any]]]:
    """Search icons and score them."""
    query_lower = query.lower()
    query_terms = query_lower.split()
    
    results = []
    for icon in index["icons"]:
        # Skip existing icons
        name_normalized = icon["name"].lower().replace(" ", "-")
        if name_normalized in existing:
            continue
        
        searchable = icon.get("searchable", "").lower()
        name_lower = icon["name"].lower()
        
        # Calculate relevance score
        relevance = 0
        if query_lower in name_lower:
            relevance += 50
        
        all_match = all(term in searchable for term in query_terms)
        if all_match:
            relevance += 30
        
        any_match = any(term in searchable for term in query_terms)
        if any_match:
            relevance += 10
        
        if relevance > 0:
            simplicity = calculate_simplicity_score(icon)
            total_score = relevance + simplicity
            results.append((total_score, icon))
    
    results.sort(key=lambda x: (-x[0], x[1]["name"]))
    return results


def filter_by_category(index: dict[str, Any], categories: list[str], existing: set[str]) -> list[tuple[int, dict[str, Any]]]:
    """Filter icons by category and score them."""
    categories_lower = [c.lower() for c in categories]
    
    results = []
    for icon in index["icons"]:
        # Skip existing icons
        name_normalized = icon["name"].lower().replace(" ", "-")
        if name_normalized in existing:
            continue
        
        cat = icon.get("category", "").lower()
        if cat in categories_lower:
            simplicity = calculate_simplicity_score(icon)
            results.append((simplicity, icon))
    
    results.sort(key=lambda x: (-x[0], x[1]["name"]))
    return results


def get_expansion_recommendations(index: dict[str, Any], batch_num: int, existing: set[str]) -> list[tuple[int, dict[str, Any]]]:
    """Get recommendations for a specific expansion batch."""
    batch = EXPANSION_BATCHES.get(batch_num)
    if not batch:
        print(f"Unknown batch number: {batch_num}")
        print(f"Available batches: {list(EXPANSION_BATCHES.keys())}")
        return []
    
    results = []
    seen_ids = set()
    
    # Search for each concept
    for concept in batch["concepts"]:
        search_results = search_icons(index, concept, existing)
        for score, icon in search_results[:10]:  # Top 10 per concept
            if icon["source_id"] not in seen_ids:
                seen_ids.add(icon["source_id"])
                results.append((score, icon))
    
    # Also include icons from relevant categories
    cat_results = filter_by_category(index, batch["categories"], existing)
    for score, icon in cat_results[:50]:
        if icon["source_id"] not in seen_ids:
            seen_ids.add(icon["source_id"])
            results.append((score, icon))
    
    results.sort(key=lambda x: (-x[0], x[1]["name"]))
    return results


def print_recommendation(score: int, icon: dict[str, Any], verbose: bool = False) -> None:
    """Print a recommendation."""
    simplicity = calculate_simplicity_score(icon)
    elements = len(icon.get("svg_paths", []))
    
    print(f"\n  [{score:3d}] {icon['source_id']}")
    print(f"         Name: {icon['name']}")
    print(f"         Category: {icon.get('category', 'Unknown')}")
    print(f"         License: {icon.get('license', 'Unknown')}")
    print(f"         Elements: {elements} (simplicity: {simplicity})")
    
    if verbose:
        tags = icon.get("tags", [])
        print(f"         Tags: {', '.join(tags[:8])}")


def main():
    parser = argparse.ArgumentParser(description="Recommend icons for IconAid expansion")
    parser.add_argument("--search", "-s", help="Search query")
    parser.add_argument("--category", "-c", help="Filter by category")
    parser.add_argument("--expansion-batch", "-b", type=int, help="Get recommendations for expansion batch (6 or 7)")
    parser.add_argument("--limit", "-l", type=int, default=30, help="Limit results")
    parser.add_argument("--verbose", "-v", action="store_true", help="Verbose output")
    args = parser.parse_args()
    
    print("Loading icon index...")
    index = load_index()
    existing = get_existing_icon_names()
    print(f"Total icons in index: {index['total_icons']}")
    print(f"Icons already in IconAid: {len(existing)}")
    
    if args.expansion_batch:
        batch = EXPANSION_BATCHES.get(args.expansion_batch)
        if batch:
            print(f"\n{'='*60}")
            print(f"EXPANSION BATCH {args.expansion_batch}: {batch['name']}")
            print(f"{'='*60}")
            print(f"Target concepts: {', '.join(batch['concepts'])}")
            print(f"Related categories: {', '.join(batch['categories'])}")
        
        results = get_expansion_recommendations(index, args.expansion_batch, existing)
        print(f"\nFound {len(results)} recommendations")
        
        for score, icon in results[:args.limit]:
            print_recommendation(score, icon, args.verbose)
        
        if len(results) > args.limit:
            print(f"\n  ... and {len(results) - args.limit} more recommendations")
        return
    
    if args.search:
        results = search_icons(index, args.search, existing)
        print(f"\nSearch results for '{args.search}': {len(results)} matches")
        
        for score, icon in results[:args.limit]:
            print_recommendation(score, icon, args.verbose)
        
        if len(results) > args.limit:
            print(f"\n  ... and {len(results) - args.limit} more results")
        return
    
    if args.category:
        results = filter_by_category(index, [args.category], existing)
        print(f"\nCategory '{args.category}': {len(results)} icons")
        
        for score, icon in results[:args.limit]:
            print_recommendation(score, icon, args.verbose)
        
        if len(results) > args.limit:
            print(f"\n  ... and {len(results) - args.limit} more icons")
        return
    
    # Default: show available batches
    print("\nExpansion Batch Recommendations")
    print("=" * 40)
    for batch_num, batch in EXPANSION_BATCHES.items():
        print(f"\nBatch {batch_num}: {batch['name']}")
        print(f"  Concepts: {', '.join(batch['concepts'][:5])}...")
        print(f"  Run: python3 scripts/recommend_icons.py --expansion-batch {batch_num}")


if __name__ == "__main__":
    main()
