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
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parent.parent
OUTPUT_DIR = ROOT / "shared" / "iconaid" / "external-sources"
SOURCE_MANIFEST = ROOT / "shared" / "iconaid" / "sources.json"

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


def download_and_extract_npm(
    package_name: str, dest_dir: Path, requested_version: str | None = None
) -> str | None:
    """Download and extract an npm package, returning its resolved version."""
    info = fetch_npm_package_info(package_name)
    if not info:
        return None
    
    resolved_version = requested_version or info.get("dist-tags", {}).get("latest")
    if not resolved_version:
        print(f"No version found for {package_name}")
        return None
    
    tarball_url = info.get("versions", {}).get(resolved_version, {}).get("dist", {}).get("tarball")
    if not tarball_url:
        print(f"No tarball URL found for {package_name}@{resolved_version}")
        return None
    
    print(f"Downloading {package_name}@{resolved_version}...")
    
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
        return resolved_version
        
    except Exception as e:
        print(f"Failed to download {package_name}: {e}")
        return None


def load_source_manifest() -> dict[str, Any]:
    """Load the reviewed source allowlist used by fetching and normalization."""
    return json.loads(SOURCE_MANIFEST.read_text(encoding="utf-8"))


def _number(value: str | None, default: float = 0.0) -> float:
    """Parse a simple SVG numeric attribute (Iconify data has no CSS units)."""
    if value is None:
        return default
    match = re.match(r"[-+]?(?:\d*\.\d+|\d+\.?)(?:[eE][-+]?\d+)?", value)
    return float(match.group(0)) if match else default


def _fmt_number(value: float) -> str:
    text = f"{value:.6f}".rstrip("0").rstrip(".")
    return text or "0"


def _points_to_path(points: str, close: bool) -> str | None:
    values = [_number(value) for value in re.findall(
        r"[-+]?(?:\d*\.\d+|\d+\.?)(?:[eE][-+]?\d+)?", points
    )]
    if len(values) < 4 or len(values) % 2:
        return None
    pairs = list(zip(values[::2], values[1::2]))
    path = f"M{_fmt_number(pairs[0][0])} {_fmt_number(pairs[0][1])}"
    path += "".join(f" L{_fmt_number(x)} {_fmt_number(y)}" for x, y in pairs[1:])
    return path + (" Z" if close else "")


def extract_iconify_paths(body: str) -> tuple[list[str], str]:
    """Convert safe Iconify body geometry to paths and infer stroke/fill mode.

    Iconify has already removed scripts, external resources and unsupported
    elements.  This importer handles the geometry elements used by the selected
    monochrome packs. Icons with transformations are skipped instead of being
    silently distorted; the build summary makes such losses visible.
    """
    try:
        root = ET.fromstring(f'<svg xmlns="http://www.w3.org/2000/svg">{body}</svg>')
    except ET.ParseError:
        return [], "fill"

    paths: list[str] = []
    has_stroke = False
    unsupported_transform = False

    def walk(element: ET.Element, inherited: dict[str, str]) -> None:
        nonlocal has_stroke, unsupported_transform
        attrs = dict(inherited)
        attrs.update({key.split("}")[-1]: value for key, value in element.attrib.items()})
        if attrs.get("transform"):
            unsupported_transform = True
            return
        stroke = attrs.get("stroke", "")
        if stroke and stroke.lower() != "none":
            has_stroke = True

        tag = element.tag.split("}")[-1]
        if tag == "path" and attrs.get("d"):
            paths.append(attrs["d"])
        elif tag == "line":
            paths.append(
                f"M{attrs.get('x1', '0')} {attrs.get('y1', '0')} "
                f"L{attrs.get('x2', '0')} {attrs.get('y2', '0')}"
            )
        elif tag in {"polyline", "polygon"}:
            path = _points_to_path(attrs.get("points", ""), tag == "polygon")
            if path:
                paths.append(path)
        elif tag == "circle":
            cx, cy, radius = (_number(attrs.get("cx")), _number(attrs.get("cy")), _number(attrs.get("r")))
            if radius > 0:
                c, y, r = map(_fmt_number, (cx, cy, radius))
                paths.append(f"M{_fmt_number(cx-radius)} {y} A{r} {r} 0 1 0 {_fmt_number(cx+radius)} {y} A{r} {r} 0 1 0 {_fmt_number(cx-radius)} {y} Z")
        elif tag == "ellipse":
            cx, cy = _number(attrs.get("cx")), _number(attrs.get("cy"))
            rx, ry = _number(attrs.get("rx")), _number(attrs.get("ry"))
            if rx > 0 and ry > 0:
                paths.append(
                    f"M{_fmt_number(cx-rx)} {_fmt_number(cy)} A{_fmt_number(rx)} {_fmt_number(ry)} 0 1 0 "
                    f"{_fmt_number(cx+rx)} {_fmt_number(cy)} A{_fmt_number(rx)} {_fmt_number(ry)} 0 1 0 "
                    f"{_fmt_number(cx-rx)} {_fmt_number(cy)} Z"
                )
        elif tag == "rect":
            x, y = _number(attrs.get("x")), _number(attrs.get("y"))
            width, height = _number(attrs.get("width")), _number(attrs.get("height"))
            if width > 0 and height > 0:
                paths.append(
                    f"M{_fmt_number(x)} {_fmt_number(y)} H{_fmt_number(x+width)} "
                    f"V{_fmt_number(y+height)} H{_fmt_number(x)} Z"
                )

        for child in element:
            walk(child, attrs)

    walk(root, {})
    return ([], "stroke" if has_stroke else "fill") if unsupported_transform else (paths, "stroke" if has_stroke else "fill")


def process_iconify_json(extract_dir: Path, source: dict[str, Any]) -> list[dict]:
    """Process one reviewed @iconify-json package into normalized source data."""
    package_dir = extract_dir / "package"
    icons_path = package_dir / "icons.json"
    if not icons_path.exists():
        print(f"Iconify icons.json not found for {source['id']}")
        return []

    data = json.loads(icons_path.read_text(encoding="utf-8"))
    default_width = data.get("width", source.get("viewBox", 24))
    default_height = data.get("height", default_width)
    category_by_icon: dict[str, list[str]] = {}
    for category, names in data.get("categories", {}).items():
        for icon_name in names:
            category_by_icon.setdefault(icon_name, []).append(category)
    aliases_by_parent: dict[str, list[str]] = {}
    for alias, alias_data in data.get("aliases", {}).items():
        aliases_by_parent.setdefault(alias_data.get("parent", ""), []).append(alias)

    icons = []
    skipped = 0
    filtered = 0
    include_pattern = source.get("includePattern")
    exclude_pattern = source.get("excludePattern")
    category_override = source.get("categoryOverride")
    source_tags = source.get("sourceTags", [])
    for icon_id, icon_data in data.get("icons", {}).items():
        if include_pattern and not re.search(include_pattern, icon_id):
            filtered += 1
            continue
        if exclude_pattern and re.search(exclude_pattern, icon_id):
            filtered += 1
            continue
        paths, inferred_mode = extract_iconify_paths(icon_data.get("body", ""))
        if not paths:
            skipped += 1
            continue
        words = re.split(r"[-_\s]+", icon_id.lower())
        categories = category_by_icon.get(icon_id, [])
        aliases = aliases_by_parent.get(icon_id, [])
        category = category_override or (
            categories[0].replace("-", " ").title()
            if categories else categorize_by_name(icon_id)
        )
        tags = generate_tags(icon_id) + words + categories + aliases + source_tags
        tags = list(dict.fromkeys(tag.lower() for tag in tags if tag))
        render_mode = source.get("renderMode", "auto")
        if render_mode == "auto":
            render_mode = inferred_mode
        width = icon_data.get("width", default_width)
        height = icon_data.get("height", default_height)
        if width != height:
            skipped += 1
            continue
        display_name = " ".join(word.capitalize() for word in re.split(r"[-_]", icon_id))
        icons.append({
            "id": icon_id,
            "name": display_name,
            "category": category,
            "tags": tags,
            "license": source["license"],
            "source": source["id"],
            "svg_paths": paths,
            "viewBox": width,
            "renderMode": render_mode,
            "searchable": " ".join([icon_id, display_name, category, *tags]).lower(),
        })
    summary = f"Found {len(icons)} {source['name']} icons; skipped {skipped} unsupported icons"
    if filtered:
        summary += f"; filtered {filtered} non-canonical variants"
    print(summary)
    return icons


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
    normalized_name = "-" + re.sub(r"[^a-z0-9]+", "-", name_lower).strip("-") + "-"

    def has_pattern(pattern: str) -> bool:
        normalized_pattern = re.sub(r"[^a-z0-9]+", "-", pattern.lower()).strip("-")
        if f"-{normalized_pattern}-" in normalized_name:
            return True
        # Common plural variants should retain their singular category without
        # reintroducing substring bugs such as classifying "clock" as "lock".
        if "-" not in normalized_pattern and len(normalized_pattern) > 3:
            return f"-{normalized_pattern}s-" in normalized_name or f"-{normalized_pattern}es-" in normalized_name
        return False
    
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
            if has_pattern(pattern):
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
    name_tokens = set(re.split(r'[-_\s]+', name_lower))
    
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
        if pattern in name_tokens:
            tags.extend(semantic_tags)
    
    return list(set(tags))


def save_normalized_file(
    icons: list[dict], source: dict[str, Any], version: str, output_dir: Path
) -> None:
    """Save normalized icons to JSON file."""
    output = {
        "source": source["id"],
        "sourceName": source["name"],
        "license": source["license"],
        "upstream": source["upstream"],
        "package": source["package"],
        "version": version,
        "total_icons": len(icons),
        "icons": sorted(icons, key=lambda x: x["name"].lower()),
    }

    output_file = output_dir / source["file"]
    with open(output_file, "w", encoding="utf-8") as f:
        json.dump(output, f, indent=2, ensure_ascii=False)
    
    print(f"Saved {len(icons)} {source['id']} icons to {output_file}")


def main():
    print("=" * 60)
    print("FAST BULK FETCH OF ALL ICON LIBRARIES")
    print("=" * 60)
    print()
    
    # Create temp directory for downloads
    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp_path = Path(tmp_dir)
        
        processors = {
            "tabler": process_tabler_icons,
            "lucide": process_lucide_icons,
            "heroicons": process_heroicons,
            "phosphor": process_phosphor_icons,
            "bootstrap": process_bootstrap_icons,
        }
        sources = [source for source in load_source_manifest()["sources"] if source.get("enabled")]
        
        all_icons = []
        source_counts = {}
        
        for source in sources:
            package_name = source["package"]
            source_name = source["id"]
            print(f"\n{'='*40}")
            print(f"Processing {source_name.upper()}")
            print(f"{'='*40}")
            
            extract_dir = tmp_path / source_name
            
            version = download_and_extract_npm(package_name, extract_dir, source.get("version"))
            if version:
                if source["format"] == "iconify-json":
                    icons = process_iconify_json(extract_dir, source)
                else:
                    icons = processors[source_name](extract_dir)
                
                if icons:
                    save_normalized_file(icons, source, version, OUTPUT_DIR)
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
        
        # Preserve alternative artwork from every source; collapse only an
        # accidental repeated record from the same source.
        seen_names = set()
        unique_icons = []
        for icon in all_icons:
            key = f"{icon.get('source', '')}:{icon['id']}"
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
