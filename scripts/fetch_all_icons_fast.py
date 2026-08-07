#!/usr/bin/env python3
"""
Fast bulk fetch of all icon libraries by downloading their npm packages or git archives.

This is MUCH faster than fetching SVGs one by one since we download the entire package once.

Libraries:
- Tabler Icons: npm package @tabler/icons
- Lucide: npm package lucide-static  
- Heroicons: npm package heroicons
- Phosphor: npm package @phosphor-icons/core
- Bootstrap Icons: npm package bootstrap-icons
"""
from __future__ import annotations

import json
import os
import re
import shutil
import subprocess
import tarfile
import tempfile
import urllib.request
from pathlib import Path
from typing import Any
from concurrent.futures import ThreadPoolExecutor, as_completed

ROOT = Path(__file__).resolve().parent.parent
OUTPUT_DIR = ROOT / "shared" / "iconaid" / "external-sources"

# NPM registry URL pattern
NPM_TARBALL_URL = "https://registry.npmjs.org/{package}/-/{name}-{version}.tgz"


def fetch_npm_package_info(package_name: str) -> dict | None:
    """Fetch package info from npm registry."""
    url = f"https://registry.npmjs.org/{package_name}"
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "IconAid/1.0"})
        with urllib.request.urlopen(req, timeout=30) as response:
            return json.loads(response.read().decode("utf-8"))
    except Exception as e:
        print(f"Failed to fetch npm info for {package_name}: {e}")
        return None


def download_and_extract_npm(package_name: str, dest_dir: Path) -> bool:
    """Download and extract npm package."""
    info = fetch_npm_package_info(package_name)
    if not info:
        return False
    
    latest_version = info.get("dist-tags", {}).get("latest")
    if not latest_version:
        print(f"No latest version found for {package_name}")
        return False
    
    tarball_url = info.get("versions", {}).get(latest_version, {}).get("dist", {}).get("tarball")
    if not tarball_url:
        print(f"No tarball URL found for {package_name}@{latest_version}")
        return False
    
    print(f"Downloading {package_name}@{latest_version}...")
    
    try:
        req = urllib.request.Request(tarball_url, headers={"User-Agent": "IconAid/1.0"})
        with urllib.request.urlopen(req, timeout=120) as response:
            with tempfile.NamedTemporaryFile(suffix=".tgz", delete=False) as tmp:
                tmp.write(response.read())
                tmp_path = tmp.name
        
        # Extract
        dest_dir.mkdir(parents=True, exist_ok=True)
        with tarfile.open(tmp_path, "r:gz") as tar:
            tar.extractall(dest_dir)
        
        os.unlink(tmp_path)
        print(f"Extracted {package_name} to {dest_dir}")
        return True
        
    except Exception as e:
        print(f"Failed to download {package_name}: {e}")
        return False


def extract_svg_paths(svg_content: str) -> list[str]:
    """Extract path d attributes and other elements from SVG."""
    paths = []
    
    # Extract path d attributes
    for match in re.finditer(r'<path[^>]*\sd="([^"]+)"', svg_content, re.IGNORECASE):
        paths.append(match.group(1))
    
    # Extract circle elements as simplified representation
    for match in re.finditer(r'<circle[^>]*cx="([^"]+)"[^>]*cy="([^"]+)"[^>]*r="([^"]+)"', svg_content, re.IGNORECASE):
        cx, cy, r = match.groups()
        paths.append(f"M {cx} {cy} m -{r} 0 a {r} {r} 0 1 0 {float(r)*2} 0 a {r} {r} 0 1 0 -{float(r)*2} 0")
    
    # Extract rect elements
    for match in re.finditer(r'<rect[^>]*x="([^"]+)"[^>]*y="([^"]+)"[^>]*width="([^"]+)"[^>]*height="([^"]+)"', svg_content, re.IGNORECASE):
        x, y, w, h = match.groups()
        paths.append(f"M {x} {y} h {w} v {h} h -{w} Z")
    
    # Extract line elements
    for match in re.finditer(r'<line[^>]*x1="([^"]+)"[^>]*y1="([^"]+)"[^>]*x2="([^"]+)"[^>]*y2="([^"]+)"', svg_content, re.IGNORECASE):
        x1, y1, x2, y2 = match.groups()
        paths.append(f"M {x1} {y1} L {x2} {y2}")
    
    # Extract polyline/polygon points
    for match in re.finditer(r'<(?:polyline|polygon)[^>]*points="([^"]+)"', svg_content, re.IGNORECASE):
        points = match.group(1).strip().split()
        if len(points) >= 2:
            path_d = f"M {points[0]}"
            for p in points[1:]:
                path_d += f" L {p}"
            paths.append(path_d)
    
    return paths


def process_tabler_icons(extract_dir: Path) -> list[dict]:
    """Process Tabler Icons from extracted npm package."""
    icons = []
    svg_dir = extract_dir / "package" / "icons" / "outline"
    
    if not svg_dir.exists():
        # Try alternative paths
        for alt in ["icons/outline", "svg", "icons"]:
            alt_dir = extract_dir / "package" / alt
            if alt_dir.exists():
                svg_dir = alt_dir
                break
    
    if not svg_dir.exists():
        print(f"SVG directory not found in Tabler package")
        return icons
    
    svg_files = list(svg_dir.glob("*.svg"))
    print(f"Found {len(svg_files)} Tabler SVGs")
    
    for svg_file in svg_files:
        try:
            content = svg_file.read_text(encoding="utf-8")
            paths = extract_svg_paths(content)
            
            if paths:
                name = svg_file.stem
                display_name = " ".join(word.capitalize() for word in name.split("-"))
                
                icons.append({
                    "id": name,
                    "name": display_name,
                    "category": categorize_by_name(name),
                    "tags": generate_tags(name),
                    "license": "MIT",
                    "source": "tabler",
                    "svg_paths": paths,
                    "searchable": f"{name} {display_name} {' '.join(generate_tags(name))}".lower(),
                })
        except Exception as e:
            print(f"Error processing {svg_file}: {e}")
    
    return icons


def process_lucide_icons(extract_dir: Path) -> list[dict]:
    """Process Lucide Icons from extracted npm package."""
    icons = []
    svg_dir = extract_dir / "package" / "icons"
    
    if not svg_dir.exists():
        print(f"SVG directory not found in Lucide package")
        return icons
    
    svg_files = list(svg_dir.glob("*.svg"))
    print(f"Found {len(svg_files)} Lucide SVGs")
    
    for svg_file in svg_files:
        try:
            content = svg_file.read_text(encoding="utf-8")
            paths = extract_svg_paths(content)
            
            if paths:
                name = svg_file.stem
                display_name = " ".join(word.capitalize() for word in name.split("-"))
                
                icons.append({
                    "id": name,
                    "name": display_name,
                    "category": categorize_by_name(name),
                    "tags": generate_tags(name),
                    "license": "ISC",
                    "source": "lucide",
                    "svg_paths": paths,
                    "searchable": f"{name} {display_name} {' '.join(generate_tags(name))}".lower(),
                })
        except Exception as e:
            print(f"Error processing {svg_file}: {e}")
    
    return icons


def process_heroicons(extract_dir: Path) -> list[dict]:
    """Process Heroicons from extracted npm package."""
    icons = []
    svg_dir = extract_dir / "package" / "24" / "outline"
    
    if not svg_dir.exists():
        # Try alternative
        svg_dir = extract_dir / "package" / "outline"
    
    if not svg_dir.exists():
        print(f"SVG directory not found in Heroicons package")
        return icons
    
    svg_files = list(svg_dir.glob("*.svg"))
    print(f"Found {len(svg_files)} Heroicons SVGs")
    
    for svg_file in svg_files:
        try:
            content = svg_file.read_text(encoding="utf-8")
            paths = extract_svg_paths(content)
            
            if paths:
                name = svg_file.stem
                display_name = " ".join(word.capitalize() for word in name.split("-"))
                
                icons.append({
                    "id": name,
                    "name": display_name,
                    "category": categorize_by_name(name),
                    "tags": generate_tags(name),
                    "license": "MIT",
                    "source": "heroicons",
                    "svg_paths": paths,
                    "searchable": f"{name} {display_name} {' '.join(generate_tags(name))}".lower(),
                })
        except Exception as e:
            print(f"Error processing {svg_file}: {e}")
    
    return icons


def process_phosphor_icons(extract_dir: Path) -> list[dict]:
    """Process Phosphor Icons from extracted npm package."""
    icons = []
    svg_dir = extract_dir / "package" / "assets" / "regular"
    
    if not svg_dir.exists():
        print(f"SVG directory not found in Phosphor package")
        return icons
    
    svg_files = list(svg_dir.glob("*.svg"))
    print(f"Found {len(svg_files)} Phosphor SVGs")
    
    for svg_file in svg_files:
        try:
            content = svg_file.read_text(encoding="utf-8")
            paths = extract_svg_paths(content)
            
            if paths:
                name = svg_file.stem
                display_name = " ".join(word.capitalize() for word in name.split("-"))
                
                icons.append({
                    "id": name,
                    "name": display_name,
                    "category": categorize_by_name(name),
                    "tags": generate_tags(name),
                    "license": "MIT",
                    "source": "phosphor",
                    "svg_paths": paths,
                    "searchable": f"{name} {display_name} {' '.join(generate_tags(name))}".lower(),
                })
        except Exception as e:
            print(f"Error processing {svg_file}: {e}")
    
    return icons


def process_bootstrap_icons(extract_dir: Path) -> list[dict]:
    """Process Bootstrap Icons from extracted npm package."""
    icons = []
    svg_dir = extract_dir / "package" / "icons"
    
    if not svg_dir.exists():
        print(f"SVG directory not found in Bootstrap package")
        return icons
    
    svg_files = list(svg_dir.glob("*.svg"))
    print(f"Found {len(svg_files)} Bootstrap SVGs")
    
    for svg_file in svg_files:
        try:
            content = svg_file.read_text(encoding="utf-8")
            paths = extract_svg_paths(content)
            
            if paths:
                name = svg_file.stem
                display_name = " ".join(word.capitalize() for word in name.split("-"))
                
                icons.append({
                    "id": name,
                    "name": display_name,
                    "category": categorize_by_name(name),
                    "tags": generate_tags(name),
                    "license": "MIT",
                    "source": "bootstrap",
                    "svg_paths": paths,
                    "searchable": f"{name} {display_name} {' '.join(generate_tags(name))}".lower(),
                })
        except Exception as e:
            print(f"Error processing {svg_file}: {e}")
    
    return icons


def categorize_by_name(name: str) -> str:
    """Categorize icon by name patterns."""
    name_lower = name.lower()
    
    categories = {
        "Arrows": ["arrow", "chevron", "caret", "sort", "move", "expand", "collapse"],
        "Communication": ["chat", "message", "mail", "envelope", "phone", "megaphone", "bell", "inbox", "send", "at-sign"],
        "Document": ["file", "folder", "clipboard", "book", "document", "page", "note", "copy", "paste"],
        "Finance": ["currency", "dollar", "euro", "cash", "wallet", "credit", "bank", "coin", "money", "receipt"],
        "Media": ["camera", "film", "image", "photo", "music", "play", "pause", "stop", "volume", "mic", "video", "speaker"],
        "Technology": ["cpu", "server", "cloud", "database", "code", "terminal", "wifi", "bluetooth", "usb", "chip", "api"],
        "Security": ["lock", "unlock", "key", "shield", "eye", "fingerprint", "scan"],
        "People": ["user", "users", "person", "people", "team", "group"],
        "Map": ["map", "pin", "location", "globe", "compass", "navigation", "route"],
        "Business": ["briefcase", "calendar", "clock", "chart", "graph", "analytics", "presentation", "meeting"],
        "E-commerce": ["cart", "bag", "basket", "shop", "store", "tag", "gift", "box", "package"],
        "Design": ["brush", "palette", "pencil", "pen", "edit", "crop", "layers", "grid"],
        "Nature": ["tree", "leaf", "flower", "sun", "moon", "star", "cloud", "rain", "snow", "wind", "fire", "water"],
        "Health": ["heart", "activity", "thermometer", "pill", "hospital"],
        "Devices": ["laptop", "desktop", "tablet", "monitor", "tv", "watch", "headphones", "keyboard", "mouse", "printer"],
        "Brand": ["brand", "logo", "facebook", "twitter", "instagram", "github", "google", "apple", "microsoft"],
    }
    
    for category, patterns in categories.items():
        for pattern in patterns:
            if pattern in name_lower:
                return category
    
    return "General"


def generate_tags(name: str) -> list[str]:
    """Generate tags from icon name."""
    tags = []
    
    # Split name into words
    words = re.split(r'[-_\s]+', name.lower())
    tags.extend(words)
    
    # Add semantic tags based on patterns
    name_lower = name.lower()
    
    semantics = {
        "chart": ["analytics", "data", "visualization", "metrics"],
        "graph": ["analytics", "data", "visualization", "statistics"],
        "calendar": ["schedule", "date", "planning", "event", "time"],
        "clock": ["time", "schedule", "deadline", "timer"],
        "mail": ["email", "message", "inbox", "communication"],
        "envelope": ["email", "message", "letter", "mail"],
        "chat": ["conversation", "message", "communication", "talk"],
        "phone": ["call", "contact", "mobile", "telephone"],
        "lock": ["security", "privacy", "protection", "secure"],
        "shield": ["security", "protection", "defense", "safe"],
        "cloud": ["storage", "internet", "hosting", "online", "saas"],
        "server": ["hosting", "backend", "infrastructure", "database"],
        "database": ["data", "storage", "sql", "records"],
        "code": ["programming", "development", "software", "developer"],
        "tree": ["nature", "environment", "ecology", "plant"],
        "leaf": ["nature", "environment", "green", "eco", "sustainability"],
        "sun": ["solar", "energy", "bright", "day", "light"],
        "user": ["person", "account", "profile", "customer"],
        "heart": ["love", "favorite", "like", "health"],
        "star": ["favorite", "rating", "bookmark", "important"],
        "file": ["document", "attachment", "paper"],
        "folder": ["directory", "organize", "files", "storage"],
        "cart": ["shopping", "ecommerce", "buy", "purchase"],
        "dollar": ["money", "finance", "payment", "cost", "price"],
        "settings": ["configuration", "options", "preferences", "gear"],
        "search": ["find", "lookup", "query", "discover"],
        "home": ["house", "main", "start", "dashboard"],
        "plus": ["add", "new", "create", "increase"],
        "minus": ["remove", "subtract", "decrease", "less"],
        "check": ["done", "complete", "success", "approve", "verify"],
        "x": ["close", "cancel", "delete", "remove", "error"],
    }
    
    for pattern, semantic_tags in semantics.items():
        if pattern in name_lower:
            tags.extend(semantic_tags)
    
    return list(set(tags))


def save_normalized_file(icons: list[dict], source: str, output_dir: Path) -> None:
    """Save normalized icons to JSON file."""
    output = {
        "source": source,
        "license": icons[0]["license"] if icons else "MIT",
        "total_icons": len(icons),
        "icons": sorted(icons, key=lambda x: x["name"].lower()),
    }
    
    output_file = output_dir / f"{source}-icons-normalized.json"
    with open(output_file, "w", encoding="utf-8") as f:
        json.dump(output, f, indent=2, ensure_ascii=False)
    
    print(f"Saved {len(icons)} {source} icons to {output_file}")


def main():
    print("=" * 60)
    print("FAST BULK FETCH OF ALL ICON LIBRARIES")
    print("=" * 60)
    print()
    
    # Create temp directory for downloads
    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp_path = Path(tmp_dir)
        
        # Define packages to download
        packages = [
            ("@tabler/icons", "tabler", process_tabler_icons),
            ("lucide-static", "lucide", process_lucide_icons),
            ("heroicons", "heroicons", process_heroicons),
            ("@phosphor-icons/core", "phosphor", process_phosphor_icons),
            ("bootstrap-icons", "bootstrap", process_bootstrap_icons),
        ]
        
        all_icons = []
        source_counts = {}
        
        for package_name, source_name, processor in packages:
            print(f"\n{'='*40}")
            print(f"Processing {source_name.upper()}")
            print(f"{'='*40}")
            
            extract_dir = tmp_path / source_name
            
            if download_and_extract_npm(package_name, extract_dir):
                icons = processor(extract_dir)
                
                if icons:
                    save_normalized_file(icons, source_name, OUTPUT_DIR)
                    all_icons.extend(icons)
                    source_counts[source_name] = len(icons)
                    print(f"✓ {source_name}: {len(icons)} icons")
                else:
                    print(f"✗ {source_name}: No icons extracted")
            else:
                print(f"✗ {source_name}: Download failed")
        
        # Build combined search index
        print(f"\n{'='*40}")
        print("BUILDING COMBINED INDEX")
        print(f"{'='*40}")
        
        # Deduplicate by name similarity
        seen_names = set()
        unique_icons = []
        for icon in all_icons:
            key = icon["name"].lower().replace(" ", "")
            if key not in seen_names:
                seen_names.add(key)
                unique_icons.append(icon)
        
        print(f"Total icons before dedup: {len(all_icons)}")
        print(f"Total icons after dedup: {len(unique_icons)}")
        
        # Build category and tag indexes
        category_index = {}
        tag_index = {}
        
        for icon in unique_icons:
            cat = icon["category"]
            if cat not in category_index:
                category_index[cat] = []
            category_index[cat].append(icon["id"])
            
            for tag in icon.get("tags", []):
                tag_lower = tag.lower()
                if tag_lower not in tag_index:
                    tag_index[tag_lower] = []
                tag_index[tag_lower].append(icon["id"])
        
        # Save combined index
        combined = {
            "total_icons": len(unique_icons),
            "source_counts": source_counts,
            "categories": list(category_index.keys()),
            "category_index": category_index,
            "tag_index": tag_index,
            "icons": sorted(unique_icons, key=lambda x: x["name"].lower()),
        }
        
        combined_file = OUTPUT_DIR / "combined-search-index.json"
        with open(combined_file, "w", encoding="utf-8") as f:
            json.dump(combined, f, indent=2, ensure_ascii=False)
        
        print(f"\nSaved combined index with {len(unique_icons)} icons")
        
        # Summary
        print(f"\n{'='*60}")
        print("SUMMARY")
        print(f"{'='*60}")
        for source, count in source_counts.items():
            print(f"  {source}: {count} icons")
        print(f"  TOTAL: {len(unique_icons)} unique icons")
        print(f"\nAll icons now have SVG path data for rendering!")


if __name__ == "__main__":
    main()
