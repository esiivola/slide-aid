#!/usr/bin/env python3
"""
Render IconAid icons to SVG or PNG with customizable styling.

This script renders any icon from the unified catalog with:
- Custom stroke color (any hex color or named color)
- Custom stroke width (1.0 to 4.0 recommended)
- Custom size (for PNG export)
- SVG or PNG output format

Usage:
    # Render single icon to SVG
    python3 scripts/render_icon.py tabler-chart-bar --color "#2563eb" --stroke 2.0

    # Render to PNG at 256px
    python3 scripts/render_icon.py lucide-mail --format png --size 256

    # Render all icons in a category
    python3 scripts/render_icon.py --category Business --output ./exports/

    # Search and render
    python3 scripts/render_icon.py --search "analytics" --limit 10
"""
from __future__ import annotations

import argparse
import json
import re
import subprocess
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parent.parent
CATALOG_PATH = ROOT / "shared" / "iconaid" / "unified-catalog.json"
DEFAULT_OUTPUT_DIR = ROOT / "shared" / "iconaid" / "rendered"

# Default IconAid style
DEFAULT_STYLE = {
    "viewBox": 24,
    "strokeWidth": 1.6,
    "strokeColor": "#1f2937",  # Dark gray
    "strokeLinecap": "round",
    "strokeLinejoin": "round",
    "fill": "none",
    "backgroundColor": "transparent",
}

# Named colors
NAMED_COLORS = {
    "black": "#000000",
    "white": "#ffffff",
    "gray": "#6b7280",
    "red": "#ef4444",
    "orange": "#f97316",
    "yellow": "#eab308",
    "green": "#22c55e",
    "blue": "#3b82f6",
    "indigo": "#6366f1",
    "purple": "#a855f7",
    "pink": "#ec4899",
    # Business colors
    "primary": "#2563eb",
    "secondary": "#64748b",
    "success": "#16a34a",
    "warning": "#d97706",
    "danger": "#dc2626",
    "info": "#0891b2",
}


def load_catalog() -> dict:
    """Load the unified icon catalog."""
    with open(CATALOG_PATH, "r", encoding="utf-8") as f:
        return json.load(f)


def get_icon_by_id(catalog: dict, icon_id: str) -> dict | None:
    """Get icon by ID from catalog."""
    for icon in catalog["icons"]:
        if icon["id"] == icon_id:
            return icon
    return None


def search_icons(catalog: dict, query: str, limit: int = 20) -> list[dict]:
    """Search icons by query."""
    query_terms = query.lower().split()
    results = []
    
    for icon in catalog["icons"]:
        searchable = icon.get("searchable", "").lower()
        name = icon.get("name", "").lower()
        
        score = 0
        for term in query_terms:
            if term in name:
                score += 10
            if term in searchable:
                score += 1
        
        if score > 0:
            results.append((score, icon))
    
    results.sort(key=lambda x: (-x[0], x[1]["name"]))
    return [icon for _, icon in results[:limit]]


def resolve_color(color: str) -> str:
    """Resolve named color to hex."""
    if color.startswith("#"):
        return color
    return NAMED_COLORS.get(color.lower(), color)


def render_svg(
    icon: dict,
    stroke_color: str = DEFAULT_STYLE["strokeColor"],
    stroke_width: float = DEFAULT_STYLE["strokeWidth"],
    view_box: int = DEFAULT_STYLE["viewBox"],
    background_color: str = "transparent",
) -> str:
    """Render icon to SVG string."""
    color = resolve_color(stroke_color)
    bg_attr = "" if background_color == "transparent" else f'style="background-color: {background_color}"'
    
    paths_svg = []
    for path_d in icon.get("paths", icon.get("svg_paths", [])):
        # Determine if this is a filled element (ends with Z and is small)
        # Most icons use stroke-only, but some have filled elements
        is_filled = False  # Default to stroke-only for consistency
        
        fill = "none"
        stroke = color
        
        paths_svg.append(
            f'  <path d="{path_d}" fill="{fill}" stroke="{stroke}" '
            f'stroke-width="{stroke_width}" stroke-linecap="round" stroke-linejoin="round"/>'
        )
    
    svg = f'''<?xml version="1.0" encoding="UTF-8"?>
<svg xmlns="http://www.w3.org/2000/svg" width="{view_box}" height="{view_box}" viewBox="0 0 {view_box} {view_box}" {bg_attr}>
  <!-- {icon["name"]} - Source: {icon.get("source", "iconaid")} - License: {icon.get("license", "MIT")} -->
{chr(10).join(paths_svg)}
</svg>'''
    
    return svg


def render_png(
    icon: dict,
    output_path: Path,
    size: int = 256,
    stroke_color: str = DEFAULT_STYLE["strokeColor"],
    stroke_width: float = DEFAULT_STYLE["strokeWidth"],
    background_color: str = "transparent",
) -> bool:
    """Render icon to PNG using system tools or cairosvg if available."""
    # First render to SVG
    svg = render_svg(
        icon,
        stroke_color=stroke_color,
        stroke_width=stroke_width,
        background_color=background_color,
    )
    
    # Try using cairosvg (Python library)
    try:
        import cairosvg
        cairosvg.svg2png(
            bytestring=svg.encode("utf-8"),
            write_to=str(output_path),
            output_width=size,
            output_height=size,
        )
        return True
    except ImportError:
        pass
    
    # Try using rsvg-convert (common on macOS/Linux)
    try:
        svg_path = output_path.with_suffix(".svg")
        svg_path.write_text(svg, encoding="utf-8")
        
        result = subprocess.run(
            ["rsvg-convert", "-w", str(size), "-h", str(size), "-o", str(output_path), str(svg_path)],
            capture_output=True,
        )
        
        svg_path.unlink()  # Clean up temp SVG
        
        if result.returncode == 0:
            return True
    except FileNotFoundError:
        pass
    
    # Try using Inkscape
    try:
        svg_path = output_path.with_suffix(".svg")
        svg_path.write_text(svg, encoding="utf-8")
        
        result = subprocess.run(
            ["inkscape", "-w", str(size), "-h", str(size), str(svg_path), "-o", str(output_path)],
            capture_output=True,
        )
        
        svg_path.unlink()  # Clean up temp SVG
        
        if result.returncode == 0:
            return True
    except FileNotFoundError:
        pass
    
    # Fallback: just save SVG
    print(f"Note: PNG conversion requires cairosvg, rsvg-convert, or inkscape.")
    print(f"Saving as SVG instead: {output_path.with_suffix('.svg')}")
    output_path.with_suffix(".svg").write_text(svg, encoding="utf-8")
    return False


def main():
    parser = argparse.ArgumentParser(
        description="Render IconAid icons with custom styling"
    )
    parser.add_argument("icon_id", nargs="?", help="Icon ID to render")
    parser.add_argument("--search", "-s", help="Search for icons")
    parser.add_argument("--category", "-c", help="Render all icons in category")
    parser.add_argument("--color", default="#1f2937", help="Stroke color (hex or name)")
    parser.add_argument("--stroke", type=float, default=1.6, help="Stroke width")
    parser.add_argument("--format", "-f", choices=["svg", "png"], default="svg", help="Output format")
    parser.add_argument("--size", type=int, default=256, help="PNG size in pixels")
    parser.add_argument("--output", "-o", type=Path, help="Output file or directory")
    parser.add_argument("--limit", "-l", type=int, default=20, help="Limit for search results")
    parser.add_argument("--list-colors", action="store_true", help="List available named colors")
    parser.add_argument("--list-categories", action="store_true", help="List available categories")
    args = parser.parse_args()
    
    if args.list_colors:
        print("Available named colors:")
        for name, hex_code in sorted(NAMED_COLORS.items()):
            print(f"  {name}: {hex_code}")
        return
    
    catalog = load_catalog()
    
    if args.list_categories:
        print(f"Available categories ({len(catalog['categories'])}):")
        for cat in sorted(catalog["categories"]):
            count = len(catalog["categoryIndex"].get(cat, []))
            print(f"  {cat}: {count} icons")
        return
    
    # Determine icons to render
    icons_to_render = []
    
    if args.search:
        icons_to_render = search_icons(catalog, args.search, args.limit)
        print(f"Found {len(icons_to_render)} icons matching '{args.search}'")
        for icon in icons_to_render:
            print(f"  - {icon['id']}: {icon['name']} ({icon['source']})")
        
        if not args.output:
            print("\nUse --output to render these icons")
            return
    
    elif args.category:
        icon_ids = catalog.get("categoryIndex", {}).get(args.category, [])
        icons_to_render = [icon for icon in catalog["icons"] if icon["id"] in icon_ids]
        print(f"Found {len(icons_to_render)} icons in category '{args.category}'")
    
    elif args.icon_id:
        icon = get_icon_by_id(catalog, args.icon_id)
        if not icon:
            # Try partial match
            matches = [i for i in catalog["icons"] if args.icon_id.lower() in i["id"].lower()]
            if matches:
                print(f"Icon '{args.icon_id}' not found. Did you mean:")
                for m in matches[:10]:
                    print(f"  - {m['id']}")
            else:
                print(f"Icon '{args.icon_id}' not found")
            return
        icons_to_render = [icon]
    
    else:
        parser.print_help()
        return
    
    # Render icons
    output_dir = args.output or DEFAULT_OUTPUT_DIR
    if len(icons_to_render) > 1:
        output_dir = Path(output_dir)
        output_dir.mkdir(parents=True, exist_ok=True)
    
    for icon in icons_to_render:
        if len(icons_to_render) == 1 and args.output:
            output_path = Path(args.output)
        else:
            output_path = output_dir / f"{icon['id']}.{args.format}"
        
        output_path.parent.mkdir(parents=True, exist_ok=True)
        
        if args.format == "svg":
            svg = render_svg(
                icon,
                stroke_color=args.color,
                stroke_width=args.stroke,
            )
            output_path.write_text(svg, encoding="utf-8")
            print(f"✓ {icon['id']} -> {output_path}")
        
        else:  # png
            success = render_png(
                icon,
                output_path,
                size=args.size,
                stroke_color=args.color,
                stroke_width=args.stroke,
            )
            if success:
                print(f"✓ {icon['id']} -> {output_path}")
    
    if len(icons_to_render) > 1:
        print(f"\nRendered {len(icons_to_render)} icons to {output_dir}")


if __name__ == "__main__":
    main()
