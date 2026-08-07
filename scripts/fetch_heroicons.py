#!/usr/bin/env python3
"""
Fetch and normalize Heroicons from their GitHub repository.

Heroicons: https://heroicons.com/
License: MIT
Repository: https://github.com/tailwindlabs/heroicons

Heroicons are designed for Tailwind CSS projects with outline (24x24) and solid (20x20) variants.

Usage:
    python3 scripts/fetch_heroicons.py
    python3 scripts/fetch_heroicons.py --output shared/iconaid/external-sources/heroicons-normalized.json

Output:
    JSON file with normalized icon data including:
    - id, name, category, tags
    - svg_paths (for simplicity scoring)
    - license, source information
"""
from __future__ import annotations

import argparse
import json
import re
import urllib.request
import urllib.error
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parent.parent
DEFAULT_OUTPUT = ROOT / "shared" / "iconaid" / "external-sources" / "heroicons-normalized.json"

# Heroicons GitHub raw URLs
# Use the outline (24x24) icons as they match IconAid's design
ICONS_API_URL = "https://api.github.com/repos/tailwindlabs/heroicons/contents/src/24/outline"
SVG_BASE_URL = "https://raw.githubusercontent.com/tailwindlabs/heroicons/master/src/24/outline"

# Category mapping based on icon naming patterns
def categorize_icon(name: str) -> tuple[str, list[str]]:
    """Categorize icon based on name patterns and return (category, tags)."""
    name_lower = name.lower()
    
    # Category patterns
    patterns = {
        "Arrows": ["arrow", "chevron", "switch"],
        "Communication": ["chat", "mail", "envelope", "phone", "inbox", "at-symbol", "megaphone", "bell", "speaker"],
        "Document": ["document", "clipboard", "folder", "newspaper", "book", "pencil", "paper"],
        "Finance": ["credit-card", "currency", "banknotes", "receipt", "calculator", "wallet"],
        "Media": ["camera", "photo", "film", "video", "microphone", "musical", "play", "pause", "stop", "speaker"],
        "Technology": ["computer", "device", "server", "cloud", "wifi", "signal", "code", "terminal", "cpu", "chip"],
        "Security": ["lock", "key", "shield", "eye", "fingerprint", "identification"],
        "People": ["user", "users", "face", "hand"],
        "Map": ["map", "globe", "location", "home", "building", "office"],
        "Business": ["briefcase", "chart", "presentation", "clipboard", "calendar", "clock", "archive"],
        "E-commerce": ["shopping", "cart", "gift", "truck", "cube", "qr-code", "tag", "ticket"],
        "Design": ["swatch", "paint", "scissors", "sparkles", "adjustments"],
        "System": ["cog", "wrench", "bolt", "power", "check", "x-mark", "plus", "minus", "magnifying"],
        "Nature": ["sun", "moon", "fire", "beaker", "bug"],
        "General": [],  # Fallback
    }
    
    category = "General"
    tags = []
    
    for cat, keywords in patterns.items():
        for keyword in keywords:
            if keyword in name_lower:
                category = cat
                tags.append(keyword.replace("-", " "))
                break
        if category != "General":
            break
    
    # Add name parts as tags
    name_parts = name.replace("-", " ").split()
    tags.extend([p.lower() for p in name_parts if p.lower() not in tags])
    
    return category, list(set(tags))


def fetch_url(url: str, as_json: bool = False) -> str | dict | list | None:
    """Fetch URL content."""
    try:
        req = urllib.request.Request(url, headers={
            "User-Agent": "IconAid/1.0",
            "Accept": "application/vnd.github.v3+json" if as_json else "text/plain",
        })
        with urllib.request.urlopen(req, timeout=30) as response:
            content = response.read().decode("utf-8")
            if as_json:
                return json.loads(content)
            return content
    except urllib.error.URLError as e:
        print(f"Failed to fetch {url}: {e}")
        return None


def list_icons_from_github() -> list[str]:
    """List all icon names from GitHub API."""
    data = fetch_url(ICONS_API_URL, as_json=True)
    if not data or not isinstance(data, list):
        return []
    
    icons = []
    for item in data:
        if item.get("type") == "file" and item.get("name", "").endswith(".svg"):
            icon_name = item["name"].replace(".svg", "")
            icons.append(icon_name)
    
    return icons


def fetch_svg_and_extract_paths(icon_name: str) -> list[str]:
    """Fetch SVG and extract elements for simplicity scoring."""
    url = f"{SVG_BASE_URL}/{icon_name}.svg"
    content = fetch_url(url)
    if not content:
        return []
    
    # Count SVG elements
    paths = re.findall(r'd="([^"]+)"', content)
    circles = re.findall(r'<circle[^>]*>', content)
    rects = re.findall(r'<rect[^>]*>', content)
    lines = re.findall(r'<line[^>]*>', content)
    polylines = re.findall(r'<polyline[^>]*>', content)
    
    elements = (
        paths +
        [f"circle{i}" for i in range(len(circles))] +
        [f"rect{i}" for i in range(len(rects))] +
        [f"line{i}" for i in range(len(lines))] +
        [f"polyline{i}" for i in range(len(polylines))]
    )
    
    return elements


def normalize_icon(icon_name: str, fetch_svg: bool = True) -> dict[str, Any]:
    """Normalize a Heroicon to IconAid format."""
    # Convert kebab-case to Title Case
    display_name = " ".join(word.capitalize() for word in icon_name.split("-"))
    
    # Get category and tags
    category, tags = categorize_icon(icon_name)
    
    # Build searchable text
    searchable_parts = [icon_name, display_name] + tags
    searchable = " ".join(searchable_parts).lower()
    
    # Fetch SVG for element count
    svg_paths = []
    if fetch_svg:
        svg_paths = fetch_svg_and_extract_paths(icon_name)
    
    return {
        "id": icon_name,
        "name": display_name,
        "category": category,
        "tags": tags,
        "aliases": [],
        "license": "MIT",
        "source": "Heroicons",
        "source_url": f"https://heroicons.com/?search={icon_name}",
        "svg_paths": svg_paths,
        "searchable": searchable,
    }


def main():
    parser = argparse.ArgumentParser(description="Fetch and normalize Heroicons")
    parser.add_argument("--output", "-o", type=Path, default=DEFAULT_OUTPUT,
                        help="Output JSON file path")
    parser.add_argument("--no-fetch-svgs", action="store_true",
                        help="Skip fetching SVG files (faster but no simplicity data)")
    parser.add_argument("--limit", "-l", type=int, default=0,
                        help="Limit number of icons to process (0 = all)")
    args = parser.parse_args()
    
    print("Fetching Heroicons list from GitHub...")
    
    icon_names = list_icons_from_github()
    if not icon_names:
        print("Failed to get icon list from GitHub")
        return
    
    print(f"Found {len(icon_names)} icons")
    
    # Apply limit if specified
    if args.limit > 0:
        icon_names = icon_names[:args.limit]
    
    # Normalize icons
    normalized = []
    fetch_svg = not args.no_fetch_svgs
    
    for i, icon_name in enumerate(icon_names):
        if i % 50 == 0:
            print(f"Processing {i}/{len(icon_names)}...")
        
        norm = normalize_icon(icon_name, fetch_svg=fetch_svg)
        normalized.append(norm)
    
    print(f"\nNormalized {len(normalized)} icons")
    
    # Sort by name
    normalized.sort(key=lambda x: x["name"].lower())
    
    # Build output
    output = {
        "source": "Heroicons",
        "source_url": "https://heroicons.com/",
        "license": "MIT",
        "license_url": "https://github.com/tailwindlabs/heroicons/blob/master/LICENSE",
        "total_icons": len(normalized),
        "generated_at": __import__("datetime").datetime.now().isoformat(),
        "icons": normalized,
    }
    
    # Write output
    args.output.parent.mkdir(parents=True, exist_ok=True)
    with open(args.output, "w", encoding="utf-8") as f:
        json.dump(output, f, indent=2, ensure_ascii=False)
    
    print(f"\nWritten to: {args.output}")
    
    # Show category distribution
    categories: dict[str, int] = {}
    for icon in normalized:
        cat = icon["category"]
        categories[cat] = categories.get(cat, 0) + 1
    
    print("\nCategory distribution:")
    for cat, count in sorted(categories.items(), key=lambda x: -x[1]):
        print(f"  {cat}: {count}")


if __name__ == "__main__":
    main()
