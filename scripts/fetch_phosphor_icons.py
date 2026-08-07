#!/usr/bin/env python3
"""
Fetch and normalize Phosphor Icons from their GitHub repository.

Phosphor Icons: https://phosphoricons.com/
License: MIT
Repository: https://github.com/phosphor-icons/core

The phosphor-icons/core repo contains catalog data with tags, categories in icons.ts.

Usage:
    python3 scripts/fetch_phosphor_icons.py
    python3 scripts/fetch_phosphor_icons.py --output shared/iconaid/external-sources/phosphor-icons-normalized.json

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
DEFAULT_OUTPUT = ROOT / "shared" / "iconaid" / "external-sources" / "phosphor-icons-normalized.json"

# Phosphor Icons core repository - TypeScript file with icon metadata
ICONS_TS_URL = "https://raw.githubusercontent.com/phosphor-icons/core/main/src/icons.ts"

# SVG base URL for regular weight
SVG_BASE_URL = "https://raw.githubusercontent.com/phosphor-icons/core/main/assets/regular"

# Category mapping to IconAid categories (Phosphor uses IconCategory enum)
CATEGORY_MAP = {
    "ARROWS": "Arrows",
    "BRAND": "Brand",
    "COMMERCE": "E-commerce",
    "COMMUNICATION": "Communication",
    "DESIGN": "Design",
    "DEVELOPMENT": "Technology",
    "EDITOR": "Document",
    "FINANCE": "Finance",
    "GAMES": "Games",
    "HEALTH": "Health",
    "MAP": "Map",
    "MEDIA": "Media",
    "NATURE": "Nature",
    "OBJECTS": "General",
    "OFFICE": "Business",
    "PEOPLE": "People",
    "SECURITY": "Security",
    "SYSTEM": "System",
    "WEATHER": "Weather",
}


def fetch_url(url: str) -> str | None:
    """Fetch URL content."""
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "IconAid/1.0"})
        with urllib.request.urlopen(req, timeout=30) as response:
            return response.read().decode("utf-8")
    except urllib.error.URLError as e:
        print(f"Failed to fetch {url}: {e}")
        return None


def parse_icons_ts(content: str) -> list[dict[str, Any]]:
    """Parse icons.ts TypeScript file to extract icon metadata."""
    icons = []
    
    # Find all icon entries using regex
    # Pattern matches objects like: { name: "icon-name", pascal_name: "IconName", categories: [...], tags: [...], ... }
    icon_pattern = re.compile(
        r'\{\s*'
        r'name:\s*"([^"]+)".*?'  # name
        r'pascal_name:\s*"([^"]+)".*?'  # pascal_name
        r'categories:\s*\[([^\]]*)\].*?'  # categories
        r'(?:figma_category:\s*\w+\.\w+,?\s*)?'  # optional figma_category
        r'tags:\s*\[([^\]]*)\]',  # tags
        re.DOTALL
    )
    
    for match in icon_pattern.finditer(content):
        name = match.group(1)
        pascal_name = match.group(2)
        categories_str = match.group(3)
        tags_str = match.group(4)
        
        # Parse categories (format: IconCategory.CATEGORY)
        categories = re.findall(r'IconCategory\.(\w+)', categories_str)
        
        # Parse tags (quoted strings)
        tags = re.findall(r'"([^"]+)"', tags_str)
        # Filter out special tags like "*new*"
        tags = [t for t in tags if not t.startswith("*")]
        
        icons.append({
            "name": name,
            "pascal_name": pascal_name,
            "categories": categories,
            "tags": tags,
        })
    
    return icons


def fetch_svg_paths(icon_name: str) -> list[str]:
    """Fetch SVG and extract path data for simplicity scoring."""
    url = f"{SVG_BASE_URL}/{icon_name}.svg"
    content = fetch_url(url)
    if not content:
        return []
    
    # Extract path d attributes
    paths = re.findall(r'd="([^"]+)"', content)
    # Also count other elements like circles, rects
    circles = re.findall(r'<circle[^>]*>', content)
    rects = re.findall(r'<rect[^>]*>', content)
    lines = re.findall(r'<line[^>]*>', content)
    
    return paths + [f"circle{i}" for i in range(len(circles))] + \
           [f"rect{i}" for i in range(len(rects))] + \
           [f"line{i}" for i in range(len(lines))]


def normalize_icon(icon: dict[str, Any], fetch_svgs: bool = False) -> dict[str, Any] | None:
    """Normalize a Phosphor icon to IconAid format."""
    name = icon.get("name", "")
    if not name:
        return None
    
    # Convert kebab-case to Title Case
    display_name = icon.get("pascal_name", "")
    if not display_name:
        display_name = " ".join(word.capitalize() for word in name.split("-"))
    else:
        # Add spaces before capital letters
        display_name = re.sub(r'([A-Z])', r' \1', display_name).strip()
    
    # Get tags
    tags = icon.get("tags", [])
    
    # Get categories and map to IconAid category
    categories = icon.get("categories", [])
    category = "General"
    for cat in categories:
        if cat in CATEGORY_MAP:
            category = CATEGORY_MAP[cat]
            break
    
    # Build searchable text
    searchable_parts = [name, display_name] + tags + [cat.lower() for cat in categories]
    searchable = " ".join(searchable_parts).lower()
    
    # Fetch SVG paths for simplicity scoring (optional, slower)
    svg_paths = []
    if fetch_svgs:
        svg_paths = fetch_svg_paths(name)
    
    return {
        "id": name,
        "name": display_name,
        "category": category,
        "tags": tags,
        "aliases": [],
        "license": "MIT",
        "source": "Phosphor Icons",
        "source_url": f"https://phosphoricons.com/?q={name}",
        "svg_paths": svg_paths,
        "searchable": searchable,
    }


def is_business_relevant(icon: dict[str, Any]) -> bool:
    """Check if icon is relevant for business/consulting use."""
    name = icon.get("name", "").lower()
    tags = [t.lower() for t in icon.get("tags", [])]
    categories = [c.upper() for c in icon.get("categories", [])]
    
    # Exclude patterns
    exclude_patterns = [
        "brand-", "smiley", "emoji", "game-controller",
        "gender", "skull", "alien", "ghost",
        "cat-", "dog-", "bird-", "fish-", "horse-",
        "moon-stars", "shooting-star", "rainbow",
        "baby", "toilet", "bathtub",
    ]
    
    for pattern in exclude_patterns:
        if pattern in name:
            return False
    
    # Exclude brand icons (too many, not business-relevant)
    if "BRAND" in categories and name.startswith("brand-"):
        return False
    
    return True


def main():
    parser = argparse.ArgumentParser(description="Fetch and normalize Phosphor Icons")
    parser.add_argument("--output", "-o", type=Path, default=DEFAULT_OUTPUT,
                        help="Output JSON file path")
    parser.add_argument("--fetch-svgs", action="store_true",
                        help="Fetch SVG files for simplicity scoring (slower)")
    parser.add_argument("--limit", "-l", type=int, default=0,
                        help="Limit number of icons to process (0 = all)")
    args = parser.parse_args()
    
    print("Fetching Phosphor Icons metadata from icons.ts...")
    
    # Fetch TypeScript file with icon metadata
    content = fetch_url(ICONS_TS_URL)
    if not content:
        print("Failed to fetch icons.ts")
        return
    
    print(f"Downloaded {len(content)} bytes")
    
    icons_data = parse_icons_ts(content)
    if not icons_data:
        print("Failed to parse icons.ts")
        return
    
    print(f"Found {len(icons_data)} icons in manifest")
    
    # Apply limit if specified
    if args.limit > 0:
        icons_data = icons_data[:args.limit]
    
    # Normalize icons
    normalized = []
    skipped = 0
    
    for i, icon in enumerate(icons_data):
        if i % 200 == 0:
            print(f"Processing {i}/{len(icons_data)}...")
        
        # Check business relevance
        if not is_business_relevant(icon):
            skipped += 1
            continue
        
        norm = normalize_icon(icon, fetch_svgs=args.fetch_svgs)
        if norm:
            normalized.append(norm)
    
    print(f"\nProcessed {len(icons_data)} icons")
    print(f"Skipped {skipped} non-business icons")
    print(f"Normalized {len(normalized)} business-relevant icons")
    
    # Sort by name
    normalized.sort(key=lambda x: x["name"].lower())
    
    # Build output
    output = {
        "source": "Phosphor Icons",
        "source_url": "https://phosphoricons.com/",
        "license": "MIT",
        "license_url": "https://github.com/phosphor-icons/core/blob/main/LICENSE",
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
    for cat, count in sorted(categories.items(), key=lambda x: -x[1])[:15]:
        print(f"  {cat}: {count}")


if __name__ == "__main__":
    main()
