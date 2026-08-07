#!/usr/bin/env python3
"""
Fetch and normalize Lucide Icons into a searchable dataset for IconAid.

Lucide Icons are ISC licensed (with some Feather-derived icons under MIT).

Usage:
    python3 scripts/fetch_lucide_icons.py

Output:
    shared/iconaid/external-sources/lucide-icons-normalized.json
"""
from __future__ import annotations

import json
import re
import subprocess
import sys
from pathlib import Path
from typing import Any
from concurrent.futures import ThreadPoolExecutor, as_completed

ROOT = Path(__file__).resolve().parent.parent
OUTPUT_DIR = ROOT / "shared" / "iconaid" / "external-sources"
OUTPUT_FILE = OUTPUT_DIR / "lucide-icons-normalized.json"

# Patterns to match business/consulting relevant icons
INCLUDE_PATTERNS = [
    r"chart", r"graph", r"analytics", r"target", r"goal", r"strategy",
    r"plan", r"report", r"document", r"file", r"folder", r"user", r"team",
    r"building", r"office", r"briefcase", r"presentation", r"calendar",
    r"clock", r"mail", r"message", r"phone", r"video", r"microphone",
    r"send", r"notification", r"bell", r"shield", r"lock", r"key",
    r"certificate", r"check", r"award", r"trophy", r"star", r"flag",
    r"bookmark", r"clipboard", r"list", r"table", r"database", r"server",
    r"cloud", r"network", r"api", r"code", r"terminal", r"settings",
    r"tool", r"wrench", r"factory", r"truck", r"package", r"warehouse",
    r"route", r"location", r"map", r"globe", r"world", r"currency",
    r"money", r"wallet", r"credit", r"bank", r"coin", r"cash", r"percent",
    r"calculator", r"receipt", r"invoice", r"solar", r"wind", r"battery",
    r"energy", r"leaf", r"plant", r"tree", r"recycle", r"droplet", r"water",
    r"temperature", r"lightning", r"bolt", r"bulb", r"idea", r"brain",
    r"puzzle", r"link", r"share", r"upload", r"download", r"arrow",
    r"trending", r"growth", r"filter", r"sort", r"search", r"zoom", r"eye",
    r"hierarchy", r"sitemap", r"flow", r"process", r"git", r"branch",
    r"merge", r"robot", r"cpu", r"chip", r"circuit", r"ai", r"sparkle",
    r"wand", r"magic", r"rocket", r"launch", r"speedometer", r"gauge",
    r"meter", r"dashboard", r"compass", r"binoculars", r"telescope",
    r"mountain", r"summit", r"diamond", r"gem", r"heart", r"handshake",
    r"hand", r"thumb", r"layers", r"stack", r"cube", r"box", r"container",
    r"grid", r"layout", r"home", r"activity", r"zap", r"alert", r"info",
    r"help", r"plus", r"minus", r"x", r"check", r"save", r"edit", r"pen",
    r"pencil", r"copy", r"paste", r"scissors", r"trash", r"archive",
    r"external", r"log", r"refresh", r"repeat", r"scan", r"print",
]


def fetch_icon_svg(icon_name: str) -> dict[str, Any] | None:
    """Fetch a single icon's SVG from GitHub raw content."""
    url = f"https://raw.githubusercontent.com/lucide-icons/lucide/main/icons/{icon_name}.svg"
    try:
        result = subprocess.run(
            ["curl", "-s", "-f", url],
            capture_output=True,
            text=True,
            timeout=30,
        )
        if result.returncode != 0:
            return None
        
        svg_content = result.stdout
        
        metadata = {
            "name": icon_name,
            "svg": extract_svg_paths(svg_content),
            "source": "lucide",
            "license": "ISC",
            "category": categorize_icon(icon_name),
            "tags": generate_tags(icon_name),
        }
        
        return metadata
    except Exception as e:
        print(f"Error fetching {icon_name}: {e}", file=sys.stderr)
        return None


def extract_svg_paths(svg_content: str) -> list[str]:
    """Extract path d attributes from SVG."""
    paths = []
    
    # Find all path elements
    path_matches = re.findall(r'<path[^>]*d="([^"]+)"', svg_content)
    for d in path_matches:
        paths.append(d)
    
    # Find circles
    circle_matches = re.findall(r'<circle[^>]*cx="([^"]+)"[^>]*cy="([^"]+)"[^>]*r="([^"]+)"', svg_content)
    for cx, cy, r in circle_matches:
        paths.append(f"CIRCLE:{cx},{cy},{r}")
    
    # Find rects
    rect_matches = re.findall(r'<rect[^>]*', svg_content)
    for rect in rect_matches:
        x = re.search(r'x="([^"]+)"', rect)
        y = re.search(r'y="([^"]+)"', rect)
        w = re.search(r'width="([^"]+)"', rect)
        h = re.search(r'height="([^"]+)"', rect)
        if x and y and w and h:
            paths.append(f"RECT:{x.group(1)},{y.group(1)},{w.group(1)},{h.group(1)}")
    
    # Find lines
    line_matches = re.findall(r'<line[^>]*x1="([^"]+)"[^>]*y1="([^"]+)"[^>]*x2="([^"]+)"[^>]*y2="([^"]+)"', svg_content)
    for x1, y1, x2, y2 in line_matches:
        paths.append(f"LINE:{x1},{y1},{x2},{y2}")
    
    # Find polylines
    polyline_matches = re.findall(r'<polyline[^>]*points="([^"]+)"', svg_content)
    for points in polyline_matches:
        paths.append(f"POLYLINE:{points}")
    
    # Find polygons
    polygon_matches = re.findall(r'<polygon[^>]*points="([^"]+)"', svg_content)
    for points in polygon_matches:
        paths.append(f"POLYGON:{points}")
    
    return paths


def categorize_icon(name: str) -> str:
    """Categorize icon based on name patterns."""
    categories = {
        "Finance": ["currency", "dollar", "euro", "pound", "yen", "bitcoin", "coin", "wallet", "credit", "bank", "receipt", "calculator", "percent", "money", "cash", "piggy"],
        "Business": ["briefcase", "presentation", "calendar", "clock", "flag", "target", "goal", "strategy", "award", "trophy", "medal", "badge", "ribbon", "certificate", "id-card", "building", "office", "company"],
        "Communication": ["mail", "message", "phone", "video", "mic", "speaker", "bell", "notification", "inbox", "send", "share", "at-sign"],
        "Technology": ["cloud", "server", "database", "code", "terminal", "cpu", "chip", "network", "wifi", "api", "git", "github", "gitlab", "bug", "bot", "robot", "brain", "sparkle", "wand", "webhook", "container", "docker"],
        "Security": ["shield", "lock", "unlock", "key", "eye", "eye-off", "alert", "scan", "fingerprint"],
        "Document": ["file", "folder", "clipboard", "book", "notebook", "document", "archive", "save", "copy", "paste", "edit", "pen", "pencil", "scissors", "trash", "print"],
        "People": ["user", "users", "person", "team", "contact", "accessibility", "baby", "smile", "frown"],
        "Operations": ["truck", "package", "box", "warehouse", "factory", "route", "map", "navigation", "compass", "location", "pin", "globe", "world", "earth"],
        "ESG": ["leaf", "plant", "tree", "trees", "flower", "sun", "sunrise", "sunset", "wind", "battery", "recycle", "droplet", "droplets", "thermometer", "flame", "zap", "bolt", "power", "plug", "solar"],
        "Charts": ["chart", "graph", "bar", "line", "pie", "trending", "activity", "area", "scatter", "radar", "gauge", "meter"],
        "Arrows": ["arrow", "chevron", "corner", "move", "maximize", "minimize", "expand", "shrink", "external"],
        "Editing": ["edit", "pen", "pencil", "type", "text", "italic", "bold", "underline", "strikethrough", "heading", "list", "align", "indent"],
        "Media": ["play", "pause", "stop", "skip", "rewind", "fast", "volume", "speaker", "music", "mic", "camera", "image", "photo", "video", "film", "tv", "monitor", "screen"],
    }
    
    name_lower = name.lower()
    for category, patterns in categories.items():
        for pattern in patterns:
            if pattern in name_lower:
                return category
    
    return "General"


def generate_tags(name: str) -> list[str]:
    """Generate tags from icon name."""
    # Split on common separators
    parts = re.split(r'[-_]', name)
    tags = [p.lower() for p in parts if len(p) > 1]
    
    # Add the full name
    tags.append(name.lower().replace("-", " "))
    
    return list(set(tags))


def get_icon_list() -> list[str]:
    """Get list of icon names from Lucide repository."""
    url = "https://api.github.com/repos/lucide-icons/lucide/git/trees/main?recursive=1"
    result = subprocess.run(
        ["curl", "-s", url],
        capture_output=True,
        text=True,
        timeout=120,
    )
    
    if result.returncode != 0:
        print("Failed to fetch icon list from GitHub API", file=sys.stderr)
        return []
    
    try:
        data = json.loads(result.stdout)
        tree = data.get("tree", [])
        icons = []
        for item in tree:
            path = item.get("path", "")
            if path.startswith("icons/") and path.endswith(".svg") and "/" not in path[6:]:
                icon_name = path.replace("icons/", "").replace(".svg", "")
                icons.append(icon_name)
        return icons
    except json.JSONDecodeError:
        print("Failed to parse GitHub API response", file=sys.stderr)
        return []


def is_relevant_icon(icon_name: str, metadata: dict[str, Any]) -> bool:
    """Check if icon is relevant for business/consulting context."""
    for pattern in INCLUDE_PATTERNS:
        if re.search(pattern, icon_name, re.IGNORECASE):
            return True
    
    tags = metadata.get("tags", [])
    tags_str = " ".join(tags).lower()
    for pattern in INCLUDE_PATTERNS:
        if re.search(pattern, tags_str, re.IGNORECASE):
            return True
    
    return False


def normalize_icon(icon: dict[str, Any]) -> dict[str, Any]:
    """Normalize icon data."""
    name = icon["name"]
    display_name = " ".join(word.capitalize() for word in name.split("-"))
    
    return {
        "id": name,
        "name": display_name,
        "source": "lucide",
        "license": "ISC",
        "category": icon.get("category", "General"),
        "tags": icon.get("tags", []),
        "svg_paths": icon.get("svg", []),
        "searchable": f"{display_name} {name} {' '.join(icon.get('tags', []))}".lower(),
    }


def main():
    """Main entry point."""
    print("Fetching Lucide Icons list...")
    all_icons = get_icon_list()
    print(f"Found {len(all_icons)} icons in Lucide repository")
    
    if not all_icons:
        print("Failed to fetch icon list", file=sys.stderr)
        return
    
    print(f"Fetching {len(all_icons)} icons...")
    
    icons_data = []
    with ThreadPoolExecutor(max_workers=10) as executor:
        futures = {executor.submit(fetch_icon_svg, name): name for name in all_icons}
        for i, future in enumerate(as_completed(futures)):
            icon_name = futures[future]
            try:
                result = future.result()
                if result:
                    icons_data.append(result)
                    if (i + 1) % 100 == 0:
                        print(f"  Processed {i + 1}/{len(all_icons)} icons...")
            except Exception as e:
                print(f"  Error processing {icon_name}: {e}", file=sys.stderr)
    
    print(f"Successfully fetched {len(icons_data)} icons")
    
    # Filter to relevant icons
    relevant_icons = [icon for icon in icons_data if is_relevant_icon(icon["name"], icon)]
    print(f"Filtered to {len(relevant_icons)} business-relevant icons")
    
    # Normalize icons
    normalized_icons = [normalize_icon(icon) for icon in relevant_icons]
    normalized_icons.sort(key=lambda x: (x["category"], x["name"]))
    
    output = {
        "source": "Lucide Icons",
        "source_url": "https://github.com/lucide-icons/lucide",
        "license": "ISC",
        "license_url": "https://github.com/lucide-icons/lucide/blob/main/LICENSE",
        "total_icons": len(normalized_icons),
        "categories": sorted(set(icon["category"] for icon in normalized_icons)),
        "generated_by": "scripts/fetch_lucide_icons.py",
        "icons": normalized_icons,
    }
    
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(output, f, indent=2, ensure_ascii=False)
    
    print(f"\nWritten to: {OUTPUT_FILE}")
    print(f"Total icons: {len(normalized_icons)}")
    print("Categories:", ", ".join(output["categories"]))
    
    print("\nIcons by category:")
    category_counts = {}
    for icon in normalized_icons:
        cat = icon["category"]
        category_counts[cat] = category_counts.get(cat, 0) + 1
    for cat, count in sorted(category_counts.items()):
        print(f"  {cat}: {count}")


if __name__ == "__main__":
    main()
