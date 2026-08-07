#!/usr/bin/env python3
"""
Fetch and normalize Bootstrap Icons from their GitHub repository.

Bootstrap Icons: https://icons.getbootstrap.com/
License: MIT
Repository: https://github.com/twbs/icons

Usage:
    python3 scripts/fetch_bootstrap_icons.py

Output:
    JSON file with normalized icon data
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
DEFAULT_OUTPUT = ROOT / "shared" / "iconaid" / "external-sources" / "bootstrap-icons-normalized.json"

# Bootstrap Icons npm package contains metadata
NPM_PACKAGE_URL = "https://registry.npmjs.org/bootstrap-icons"
ICONS_JSON_URL = "https://raw.githubusercontent.com/twbs/icons/main/font/bootstrap-icons.json"

# Category patterns for classification
CATEGORY_PATTERNS = {
    "Arrows": ["arrow", "chevron", "caret", "sort"],
    "Communication": ["chat", "envelope", "phone", "megaphone", "bell", "inbox", "send", "reply"],
    "Document": ["file", "folder", "clipboard", "book", "journal", "newspaper"],
    "Finance": ["currency", "cash", "wallet", "credit-card", "bank", "coin", "piggy"],
    "Media": ["camera", "film", "image", "music", "play", "pause", "stop", "volume", "mic", "speaker"],
    "Technology": ["cpu", "gpu", "memory", "hdd", "ssd", "usb", "wifi", "bluetooth", "router", "server", "cloud", "database"],
    "Security": ["lock", "unlock", "key", "shield", "eye", "incognito"],
    "People": ["person", "people", "emoji"],
    "Map": ["geo", "map", "pin", "globe", "compass", "signpost", "house", "building"],
    "Business": ["briefcase", "calendar", "clock", "graph", "bar-chart", "pie-chart", "kanban"],
    "E-commerce": ["cart", "bag", "basket", "shop", "receipt", "gift", "tag", "box"],
    "Design": ["brush", "palette", "paint", "pencil", "pen", "vector", "bezier", "crop", "layers"],
    "Nature": ["tree", "flower", "sun", "moon", "cloud", "snow", "lightning", "wind", "water", "fire"],
    "Health": ["heart", "hospital", "thermometer", "bandaid", "capsule"],
    "Devices": ["phone", "tablet", "laptop", "display", "tv", "watch", "headset", "keyboard", "mouse", "printer"],
    "Social": ["facebook", "twitter", "instagram", "linkedin", "github", "youtube", "discord", "slack", "twitch"],
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


def categorize_icon(name: str) -> str:
    """Categorize icon based on name patterns."""
    name_lower = name.lower()
    
    for category, patterns in CATEGORY_PATTERNS.items():
        for pattern in patterns:
            if pattern in name_lower:
                return category
    
    return "General"


def generate_tags(name: str, category: str) -> list[str]:
    """Generate rich tags from icon name."""
    tags = []
    
    # Split name into words
    words = name.replace("-", " ").split()
    tags.extend([w.lower() for w in words])
    
    # Add category as tag
    tags.append(category.lower())
    
    # Add semantic enhancements based on name patterns
    name_lower = name.lower()
    
    # Business semantics
    if "chart" in name_lower or "graph" in name_lower:
        tags.extend(["analytics", "data", "visualization", "metrics", "reporting"])
    if "calendar" in name_lower:
        tags.extend(["schedule", "date", "planning", "event", "time"])
    if "clock" in name_lower:
        tags.extend(["time", "schedule", "deadline", "duration"])
    if "briefcase" in name_lower:
        tags.extend(["work", "business", "job", "professional", "career"])
    
    # Communication semantics
    if "chat" in name_lower or "message" in name_lower:
        tags.extend(["conversation", "discussion", "communication", "talk"])
    if "envelope" in name_lower or "mail" in name_lower:
        tags.extend(["email", "message", "letter", "inbox", "communication"])
    if "phone" in name_lower or "telephone" in name_lower:
        tags.extend(["call", "contact", "mobile", "communication"])
    
    # Document semantics
    if "file" in name_lower:
        tags.extend(["document", "paper", "attachment"])
    if "folder" in name_lower:
        tags.extend(["directory", "organize", "storage", "files"])
    
    # Technology semantics
    if "cloud" in name_lower:
        tags.extend(["storage", "internet", "hosting", "saas", "online"])
    if "server" in name_lower:
        tags.extend(["hosting", "backend", "infrastructure", "compute"])
    if "database" in name_lower:
        tags.extend(["data", "storage", "sql", "records"])
    if "code" in name_lower:
        tags.extend(["programming", "development", "software", "developer"])
    
    # Security semantics
    if "lock" in name_lower:
        tags.extend(["security", "privacy", "protection", "secure"])
    if "shield" in name_lower:
        tags.extend(["security", "protection", "defense", "safe"])
    if "key" in name_lower:
        tags.extend(["access", "authentication", "unlock", "password"])
    
    # Finance semantics
    if "currency" in name_lower or "cash" in name_lower or "money" in name_lower:
        tags.extend(["payment", "finance", "money", "cost", "price"])
    if "cart" in name_lower or "basket" in name_lower:
        tags.extend(["shopping", "purchase", "buy", "ecommerce"])
    
    # Nature/ESG semantics
    if "tree" in name_lower or "leaf" in name_lower:
        tags.extend(["nature", "environment", "green", "ecology", "sustainability"])
    if "sun" in name_lower:
        tags.extend(["solar", "energy", "bright", "day"])
    if "recycle" in name_lower:
        tags.extend(["sustainability", "environment", "green", "reuse", "eco"])
    
    return list(set(tags))


def normalize_icon(name: str, codepoint: int) -> dict[str, Any]:
    """Normalize a Bootstrap icon to IconAid format."""
    # Convert kebab-case to Title Case
    display_name = " ".join(word.capitalize() for word in name.split("-"))
    
    category = categorize_icon(name)
    tags = generate_tags(name, category)
    
    # Build searchable text
    searchable = f"{name} {display_name} {' '.join(tags)}".lower()
    
    return {
        "id": name,
        "name": display_name,
        "category": category,
        "tags": tags,
        "aliases": [],
        "license": "MIT",
        "source": "Bootstrap Icons",
        "source_url": f"https://icons.getbootstrap.com/icons/{name}/",
        "codepoint": codepoint,
        "svg_paths": [],  # Would need to fetch individual SVGs
        "searchable": searchable,
    }


def is_business_relevant(name: str) -> bool:
    """Check if icon is relevant for business use."""
    name_lower = name.lower()
    
    # Exclude patterns
    exclude_patterns = [
        "emoji", "suit-", "gender", "explicit",
    ]
    
    for pattern in exclude_patterns:
        if pattern in name_lower:
            return False
    
    return True


def main():
    parser = argparse.ArgumentParser(description="Fetch and normalize Bootstrap Icons")
    parser.add_argument("--output", "-o", type=Path, default=DEFAULT_OUTPUT)
    parser.add_argument("--limit", "-l", type=int, default=0)
    args = parser.parse_args()
    
    print("Fetching Bootstrap Icons metadata...")
    
    # Fetch icons.json from GitHub
    content = fetch_url(ICONS_JSON_URL)
    if not content:
        print("Failed to fetch icons.json")
        return
    
    try:
        icons_data = json.loads(content)
    except json.JSONDecodeError:
        print("Failed to parse icons.json")
        return
    
    print(f"Found {len(icons_data)} icons")
    
    # Normalize icons
    normalized = []
    skipped = 0
    
    icon_names = list(icons_data.keys())
    if args.limit > 0:
        icon_names = icon_names[:args.limit]
    
    for name in icon_names:
        if not is_business_relevant(name):
            skipped += 1
            continue
        
        codepoint = icons_data[name]
        norm = normalize_icon(name, codepoint)
        normalized.append(norm)
    
    print(f"Normalized {len(normalized)} icons (skipped {skipped})")
    
    # Sort by name
    normalized.sort(key=lambda x: x["name"].lower())
    
    # Build output
    output = {
        "source": "Bootstrap Icons",
        "source_url": "https://icons.getbootstrap.com/",
        "license": "MIT",
        "license_url": "https://github.com/twbs/icons/blob/main/LICENSE",
        "total_icons": len(normalized),
        "generated_at": __import__("datetime").datetime.now().isoformat(),
        "icons": normalized,
    }
    
    # Write output
    args.output.parent.mkdir(parents=True, exist_ok=True)
    with open(args.output, "w", encoding="utf-8") as f:
        json.dump(output, f, indent=2, ensure_ascii=False)
    
    print(f"Written to: {args.output}")
    
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
