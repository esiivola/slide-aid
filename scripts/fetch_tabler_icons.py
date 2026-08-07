#!/usr/bin/env python3
"""
Fetch and normalize Tabler Icons into a searchable dataset for IconAid.

This script downloads icon metadata from Tabler Icons (MIT License) and creates
a normalized JSON dataset that can be used as reference material for IconAid expansion.

Usage:
    python3 scripts/fetch_tabler_icons.py

Output:
    shared/iconaid/external-sources/tabler-icons-normalized.json
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
OUTPUT_FILE = OUTPUT_DIR / "tabler-icons-normalized.json"

# Categories relevant to management consulting / business context
RELEVANT_CATEGORIES = {
    "Business",
    "Communication",
    "Design", 
    "Devices",
    "Document",
    "Finance",
    "Map",
    "Math",
    "Media",
    "Nature",
    "System",
    "Symbols",
    "Weather",
}

# Icon name patterns to include (business/consulting relevant)
INCLUDE_PATTERNS = [
    r"chart",
    r"graph",
    r"analytics",
    r"target",
    r"goal",
    r"strategy",
    r"plan",
    r"report",
    r"document",
    r"file",
    r"folder",
    r"user",
    r"team",
    r"building",
    r"office",
    r"briefcase",
    r"presentation",
    r"calendar",
    r"clock",
    r"mail",
    r"message",
    r"phone",
    r"video",
    r"microphone",
    r"send",
    r"notification",
    r"bell",
    r"shield",
    r"lock",
    r"key",
    r"certificate",
    r"check",
    r"award",
    r"trophy",
    r"star",
    r"flag",
    r"bookmark",
    r"clipboard",
    r"list",
    r"table",
    r"database",
    r"server",
    r"cloud",
    r"network",
    r"api",
    r"code",
    r"terminal",
    r"settings",
    r"tool",
    r"wrench",
    r"factory",
    r"truck",
    r"package",
    r"warehouse",
    r"route",
    r"location",
    r"map",
    r"globe",
    r"world",
    r"currency",
    r"money",
    r"wallet",
    r"credit",
    r"bank",
    r"coin",
    r"cash",
    r"percent",
    r"calculator",
    r"receipt",
    r"invoice",
    r"solar",
    r"wind",
    r"battery",
    r"energy",
    r"leaf",
    r"plant",
    r"tree",
    r"recycle",
    r"droplet",
    r"water",
    r"temperature",
    r"lightning",
    r"bolt",
    r"bulb",
    r"idea",
    r"brain",
    r"puzzle",
    r"link",
    r"share",
    r"upload",
    r"download",
    r"arrow",
    r"trending",
    r"growth",
    r"increase",
    r"decrease",
    r"filter",
    r"sort",
    r"search",
    r"zoom",
    r"eye",
    r"view",
    r"hierarchy",
    r"sitemap",
    r"flow",
    r"process",
    r"git",
    r"branch",
    r"merge",
    r"robot",
    r"cpu",
    r"chip",
    r"circuit",
    r"ai",
    r"sparkle",
    r"wand",
    r"magic",
    r"rocket",
    r"launch",
    r"speedometer",
    r"gauge",
    r"meter",
    r"dashboard",
    r"compass",
    r"binoculars",
    r"telescope",
    r"mountain",
    r"summit",
    r"diamond",
    r"gem",
    r"heart",
    r"handshake",
    r"hand",
    r"thumb",
    r"layers",
    r"stack",
    r"cube",
    r"box",
    r"container",
    r"grid",
    r"layout",
]


def fetch_icon_svg(icon_name: str) -> dict[str, Any] | None:
    """Fetch a single icon's SVG from GitHub raw content."""
    url = f"https://raw.githubusercontent.com/tabler/tabler-icons/main/icons/outline/{icon_name}.svg"
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
        
        # Parse metadata from SVG comments
        metadata = parse_svg_metadata(svg_content)
        metadata["name"] = icon_name
        metadata["svg"] = extract_svg_paths(svg_content)
        metadata["source"] = "tabler"
        metadata["license"] = "MIT"
        
        return metadata
    except Exception as e:
        print(f"Error fetching {icon_name}: {e}", file=sys.stderr)
        return None


def parse_svg_metadata(svg_content: str) -> dict[str, Any]:
    """Parse category, tags, and version from SVG comment header."""
    metadata = {
        "category": "Unknown",
        "tags": [],
        "version": "unknown",
    }
    
    # Extract comment block
    comment_match = re.search(r"<!--(.*?)-->", svg_content, re.DOTALL)
    if comment_match:
        comment = comment_match.group(1)
        
        # Parse category
        cat_match = re.search(r"category:\s*(.+)", comment)
        if cat_match:
            metadata["category"] = cat_match.group(1).strip()
        
        # Parse tags (array format)
        tags_match = re.search(r"tags:\s*\[(.*?)\]", comment, re.DOTALL)
        if tags_match:
            tags_str = tags_match.group(1)
            # Parse quoted and unquoted tags
            tags = re.findall(r'"([^"]+)"|([a-z0-9-]+)', tags_str)
            metadata["tags"] = [t[0] or t[1] for t in tags if t[0] or t[1]]
        
        # Parse version
        ver_match = re.search(r'version:\s*"?([0-9.]+)"?', comment)
        if ver_match:
            metadata["version"] = ver_match.group(1)
    
    return metadata


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
    rect_matches = re.findall(r'<rect[^>]*x="([^"]+)"[^>]*y="([^"]+)"[^>]*width="([^"]+)"[^>]*height="([^"]+)"', svg_content)
    for x, y, w, h in rect_matches:
        paths.append(f"RECT:{x},{y},{w},{h}")
    
    # Find lines
    line_matches = re.findall(r'<line[^>]*x1="([^"]+)"[^>]*y1="([^"]+)"[^>]*x2="([^"]+)"[^>]*y2="([^"]+)"', svg_content)
    for x1, y1, x2, y2 in line_matches:
        paths.append(f"LINE:{x1},{y1},{x2},{y2}")
    
    return paths


def get_icon_list() -> list[str]:
    """Get list of icon names from Tabler repository using Git tree API (no pagination limit)."""
    # Use Git tree API to get all files at once
    url = "https://api.github.com/repos/tabler/tabler-icons/git/trees/main?recursive=1"
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
            if path.startswith("icons/outline/") and path.endswith(".svg"):
                icon_name = path.replace("icons/outline/", "").replace(".svg", "")
                icons.append(icon_name)
        return icons
    except json.JSONDecodeError:
        print("Failed to parse GitHub API response", file=sys.stderr)
        return []


def is_relevant_icon(icon_name: str, metadata: dict[str, Any]) -> bool:
    """Check if icon is relevant for business/consulting context."""
    # Check name patterns first (most relevant)
    for pattern in INCLUDE_PATTERNS:
        if re.search(pattern, icon_name, re.IGNORECASE):
            return True
    
    # Check tags
    tags = metadata.get("tags", [])
    tags_str = " ".join(tags).lower()
    for pattern in INCLUDE_PATTERNS:
        if re.search(pattern, tags_str, re.IGNORECASE):
            return True
    
    # Check category (accept certain categories even without pattern match)
    category = metadata.get("category", "")
    if category in {"Business", "Finance", "Document", "Communication"}:
        return True
    
    return False


def normalize_icon(icon: dict[str, Any]) -> dict[str, Any]:
    """Normalize icon data into IconAid-compatible format."""
    name = icon["name"]
    
    # Convert kebab-case to Title Case
    display_name = " ".join(word.capitalize() for word in name.split("-"))
    
    # Create normalized structure
    normalized = {
        "id": name,
        "name": display_name,
        "source": "tabler",
        "license": "MIT",
        "category": icon.get("category", "Unknown"),
        "tags": icon.get("tags", []),
        "version": icon.get("version", "unknown"),
        "svg_paths": icon.get("svg", []),
        "searchable": f"{display_name} {name} {' '.join(icon.get('tags', []))}".lower(),
    }
    
    return normalized


def main():
    """Main entry point."""
    print("Fetching Tabler Icons list...")
    all_icons = get_icon_list()
    print(f"Found {len(all_icons)} icons in Tabler repository")
    
    if not all_icons:
        # Fallback to a curated list if API fails
        print("Using fallback curated list...")
        all_icons = FALLBACK_ICONS
    
    print(f"Fetching {len(all_icons)} icons...")
    
    # Fetch icons in parallel
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
    
    # Sort by category then name
    normalized_icons.sort(key=lambda x: (x["category"], x["name"]))
    
    # Create output structure
    output = {
        "source": "Tabler Icons",
        "source_url": "https://github.com/tabler/tabler-icons",
        "license": "MIT",
        "license_url": "https://github.com/tabler/tabler-icons/blob/main/LICENSE",
        "total_icons": len(normalized_icons),
        "categories": sorted(set(icon["category"] for icon in normalized_icons)),
        "generated_by": "scripts/fetch_tabler_icons.py",
        "icons": normalized_icons,
    }
    
    # Write output
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(output, f, indent=2, ensure_ascii=False)
    
    print(f"\nWritten to: {OUTPUT_FILE}")
    print(f"Total icons: {len(normalized_icons)}")
    print("Categories:", ", ".join(output["categories"]))
    
    # Print summary by category
    print("\nIcons by category:")
    category_counts = {}
    for icon in normalized_icons:
        cat = icon["category"]
        category_counts[cat] = category_counts.get(cat, 0) + 1
    for cat, count in sorted(category_counts.items()):
        print(f"  {cat}: {count}")


# Fallback list of business-relevant icons if API fails
FALLBACK_ICONS = [
    # Business
    "analytics", "briefcase", "target", "chart-bar", "chart-line", "chart-pie",
    "chart-area", "chart-donut", "chart-infographic", "presentation", "calendar",
    "clock", "clipboard", "checklist", "list", "list-check", "table", "layout-grid",
    "layout-dashboard", "sitemap", "hierarchy", "flag", "bookmark", "star",
    "award", "trophy", "medal", "certificate", "badge", "ribbon",
    
    # People
    "user", "users", "users-group", "user-plus", "user-check", "user-search",
    "friends", "building", "building-skyscraper", "home", "office-building",
    
    # Communication
    "mail", "mail-opened", "message", "message-circle", "messages", "phone",
    "phone-call", "video", "camera", "microphone", "send", "share", "bell",
    "notification", "speakerphone", "broadcast",
    
    # Finance
    "currency-dollar", "currency-euro", "coin", "coins", "cash", "wallet",
    "credit-card", "bank", "receipt", "file-invoice", "calculator", "percent",
    "chart-candle", "trending-up", "trending-down", "scale", "piggy-bank",
    
    # Technology
    "cloud", "cloud-computing", "database", "server", "server-2", "api",
    "code", "terminal", "terminal-2", "device-laptop", "device-desktop",
    "device-mobile", "wifi", "network", "cpu", "cpu-2", "circuit-board",
    "robot", "brain", "sparkles", "wand", "bolt", "rocket", "atom",
    
    # Security
    "shield", "shield-check", "shield-lock", "lock", "lock-open", "key",
    "fingerprint", "eye", "eye-off", "alert-triangle", "alert-circle",
    
    # Operations
    "truck", "truck-delivery", "package", "packages", "box", "archive",
    "building-warehouse", "forklift", "crane", "map-pin", "map", "route",
    "compass", "navigation", "world", "globe", "location",
    
    # Tools & Settings
    "settings", "settings-2", "adjustments", "tool", "tools", "wrench",
    "hammer", "screwdriver", "filter", "sort-ascending", "sort-descending",
    "search", "zoom-in", "zoom-out",
    
    # Documents
    "file", "file-text", "file-analytics", "file-report", "file-chart",
    "folder", "folder-open", "folders", "notebook", "book", "book-2",
    "report", "clipboard-text", "clipboard-list", "clipboard-check",
    "clipboard-data", "notes", "article",
    
    # Arrows & Navigation
    "arrow-up", "arrow-down", "arrow-left", "arrow-right", "arrows-diagonal",
    "arrows-horizontal", "arrows-vertical", "chevron-up", "chevron-down",
    "chevrons-up", "chevrons-down", "refresh", "reload", "repeat",
    
    # Sustainability / ESG
    "leaf", "plant", "plant-2", "tree", "trees", "flower", "seeding",
    "solar-panel", "solar-panel-2", "windmill", "wind", "wind-electricity",
    "battery", "battery-charging", "battery-eco", "plug", "charging-pile",
    "droplet", "droplets", "temperature", "thermometer", "sun", "cloud-rain",
    "recycle", "trash", "trash-x", "flame", "fire",
    
    # Misc business
    "puzzle", "puzzle-2", "link", "unlink", "attachment", "paperclip",
    "tag", "tags", "license", "id", "id-badge", "qrcode", "barcode",
    "bulb", "bulb-off", "lamp", "spotlight", "brain", "lightbulb",
    "question-mark", "info-circle", "help", "lifebuoy",
    "hand-click", "hand-finger", "hand-move", "hand-stop", "thumb-up",
    "thumb-down", "mood-happy", "mood-smile", "mood-sad",
]


if __name__ == "__main__":
    main()
