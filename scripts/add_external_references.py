#!/usr/bin/env python3
"""
Add external icon references to the IconAid catalog for searchability.

This script enriches the catalog.json with references to high-quality external icons
that match IconAid's design criteria. These serve as "suggested additions" that can
inspire manual icon creation or be converted with careful review.

External icons are marked with:
- source: The origin library (tabler, lucide, heroicons, phosphor, bootstrap)
- status: "external-reference" (not directly usable, needs conversion)
- svg_preview: SVG path data for preview purposes
- conversion_notes: Guidelines for adapting to IconAid style
"""
from __future__ import annotations

import json
from pathlib import Path
from collections import Counter

ROOT = Path(__file__).resolve().parent.parent
CATALOG_PATH = ROOT / "shared" / "iconaid" / "catalog.json"
INDEX_PATH = ROOT / "shared" / "iconaid" / "external-sources" / "combined-search-index.json"
OUTPUT_PATH = ROOT / "shared" / "iconaid" / "catalog-with-external.json"


def load_external_index() -> dict:
    """Load the combined search index."""
    with open(INDEX_PATH, "r", encoding="utf-8") as f:
        return json.load(f)


def load_catalog() -> dict:
    """Load the current IconAid catalog."""
    with open(CATALOG_PATH, "r", encoding="utf-8") as f:
        return json.load(f)


def get_existing_icon_names(catalog: dict) -> set[str]:
    """Get normalized names of icons already in IconAid."""
    names = set()
    for icon in catalog.get("icons", []):
        # Add ID and normalized name
        names.add(icon.get("id", "").lower())
        names.add(icon.get("name", "").lower().replace(" ", "-"))
        # Add aliases
        for alias in icon.get("aliases", []):
            names.add(alias.lower().replace(" ", "-"))
    return names


def calculate_quality_score(icon: dict) -> int:
    """
    Calculate a quality score for an icon.
    Higher = better fit for IconAid.
    
    Criteria:
    - Simplicity: 4-10 elements is ideal
    - License: MIT/ISC/Apache preferred
    - Tags: More tags = better searchability
    """
    score = 0
    
    # Simplicity score (based on number of SVG paths)
    num_paths = len(icon.get("svg_paths", []))
    if 4 <= num_paths <= 10:
        score += 50  # Ideal complexity
    elif num_paths < 4:
        score += 30  # Too simple
    elif num_paths <= 16:
        score += 40  # Acceptable
    else:
        score += 10  # Too complex
    
    # License score
    license_text = icon.get("license", "").upper()
    if "MIT" in license_text or "ISC" in license_text:
        score += 30
    elif "APACHE" in license_text:
        score += 25
    else:
        score += 10
    
    # Tag richness score
    num_tags = len(icon.get("tags", []))
    score += min(num_tags, 20)  # Cap at 20 points for tags
    
    return score


def is_business_relevant(icon: dict) -> bool:
    """
    Check if an icon is relevant for business/consulting use.
    """
    business_categories = {
        "business", "finance", "technology", "communication", 
        "document", "media", "security", "operations",
        "people", "esg", "nature", "charts", "devices"
    }
    
    category = icon.get("category", "").lower()
    if category in business_categories:
        return True
    
    # Check tags for business relevance
    business_keywords = {
        "analytics", "data", "chart", "graph", "report", "dashboard",
        "money", "finance", "budget", "cost", "revenue", "profit",
        "team", "people", "user", "employee", "customer",
        "document", "file", "folder", "mail", "email", "message",
        "cloud", "server", "database", "api", "code", "software",
        "security", "lock", "shield", "key", "compliance",
        "process", "workflow", "automation", "integration",
        "sustainability", "energy", "environment", "climate",
        "presentation", "meeting", "calendar", "schedule"
    }
    
    tags = set(t.lower() for t in icon.get("tags", []))
    name_words = set(icon.get("name", "").lower().split())
    
    return bool(tags & business_keywords) or bool(name_words & business_keywords)


def select_top_icons(index: dict, existing: set[str], limit: int = 500) -> list[dict]:
    """
    Select top icons from external sources.
    
    Criteria:
    - Not already in IconAid
    - Business/consulting relevant
    - High quality score
    """
    candidates = []
    
    for icon in index.get("icons", []):
        # Skip if already in IconAid
        name_normalized = icon["name"].lower().replace(" ", "-")
        if name_normalized in existing:
            continue
        
        # Skip if not business relevant
        if not is_business_relevant(icon):
            continue
        
        score = calculate_quality_score(icon)
        candidates.append((score, icon))
    
    # Sort by score and take top N
    candidates.sort(key=lambda x: (-x[0], x[1]["name"]))
    return [icon for _, icon in candidates[:limit]]


def create_external_reference(icon: dict) -> dict:
    """
    Create an external reference entry for the catalog.
    """
    source = icon.get("source", "unknown")
    source_id = icon.get("source_id", icon.get("id", "unknown"))
    
    return {
        "id": f"ext-{source_id.replace(':', '-')}",
        "name": icon["name"],
        "category": icon.get("category", "General"),
        "aliases": [],
        "tags": icon.get("tags", [])[:15],  # Limit to 15 tags
        "source": source,
        "sourceId": source_id,
        "license": icon.get("license", "Unknown"),
        "status": "external-reference",
        "svgPaths": icon.get("svg_paths", []),
        "searchable": icon.get("searchable", ""),
        "conversionNotes": (
            f"This icon from {source} may be adapted for IconAid. "
            f"Ensure 24×24 grid, 1.6px stroke, round caps/joins, and 4-10 visible elements."
        ),
    }


def main():
    print("Loading catalog and external index...")
    catalog = load_catalog()
    index = load_external_index()
    existing = get_existing_icon_names(catalog)
    
    print(f"IconAid icons: {len(catalog.get('icons', []))}")
    print(f"External icons: {index['total_icons']}")
    print(f"Existing icon names to exclude: {len(existing)}")
    
    print("\nSelecting top business-relevant icons...")
    top_icons = select_top_icons(index, existing, limit=1000)
    print(f"Selected {len(top_icons)} icons")
    
    # Analyze categories
    categories = Counter(icon.get("category") for icon in top_icons)
    print("\nCategories in selection:")
    for cat, count in categories.most_common(15):
        print(f"  {cat}: {count}")
    
    # Create external references
    external_refs = [create_external_reference(icon) for icon in top_icons]
    
    # Create enriched catalog
    enriched_catalog = dict(catalog)
    enriched_catalog["externalReferences"] = {
        "description": "Searchable icons from open-source libraries. Not directly usable - require manual conversion to IconAid style.",
        "sources": index.get("source_counts", {}),
        "total": len(external_refs),
        "icons": external_refs,
    }
    
    # Save enriched catalog
    with open(OUTPUT_PATH, "w", encoding="utf-8") as f:
        json.dump(enriched_catalog, f, indent=2)
    print(f"\nSaved enriched catalog to: {OUTPUT_PATH}")
    print(f"  IconAid icons: {len(catalog.get('icons', []))}")
    print(f"  External references: {len(external_refs)}")
    
    # Also create a summary
    print("\n" + "="*60)
    print("SUMMARY")
    print("="*60)
    print(f"Total searchable icons: {len(catalog.get('icons', [])) + len(external_refs)}")
    print(f"  - Native IconAid icons: {len(catalog.get('icons', []))}")
    print(f"  - External references: {len(external_refs)}")
    print("\nThe external references provide:")
    print("  - Rich search tags for discovery")
    print("  - Source attribution and licensing info")
    print("  - SVG path data for preview/conversion")
    print("  - Conversion guidelines for IconAid style")


if __name__ == "__main__":
    main()
