# IconAid Icon Sources Research

This document catalogs open-source icon libraries that can be used as reference or source material for extending the IconAid catalog. All listed libraries have licenses compatible with free commercial use.

## Implementation Status ✅

A normalized, searchable database of **5,930 unique icons** from 4 open-source libraries has been created:

| Source | Icons | Status |
|--------|-------|--------|
| Tabler Icons | 3,990 | ✅ Fetched and normalized |
| Phosphor Icons | 1,480 | ✅ Fetched and normalized |
| Lucide | 920 | ✅ Fetched and normalized |
| Heroicons | 324 | ✅ Fetched and normalized |
| **Total Unique** | **5,930** | ✅ Combined search index |

### Search Tools

```bash
# Rebuild combined index
python3 scripts/build_icon_search_index.py --rebuild

# Search for icons
python3 scripts/build_icon_search_index.py --search "analytics chart"
python3 scripts/build_icon_search_index.py --search "sustainability ESG"

# List all categories
python3 scripts/build_icon_search_index.py --list-categories

# Get expansion recommendations
python3 scripts/recommend_icons.py --expansion-batch 6  # Communication/Document
python3 scripts/recommend_icons.py --expansion-batch 7  # ESG

# Search specific terms
python3 scripts/recommend_icons.py --search "solar energy"
```

### Files

- `scripts/fetch_tabler_icons.py` - Fetch and normalize Tabler Icons
- `scripts/fetch_lucide_icons.py` - Fetch and normalize Lucide Icons  
- `scripts/fetch_heroicons.py` - Fetch and normalize Heroicons
- `scripts/fetch_phosphor_icons.py` - Fetch and normalize Phosphor Icons
- `scripts/build_icon_search_index.py` - Build combined searchable index
- `scripts/recommend_icons.py` - Recommend icons for IconAid expansion
- `external-sources/*.json` - Normalized icon data files

## Design Requirements Summary

IconAid icons must:
- 24×24 grid with 2-unit safe area
- 1.6px stroke width (SVG preview), round caps and joins
- Monochrome, no gradients or shadows
- 4-10 visible strokes or shapes (max 16)
- Management consulting style: simple, clean, minimalistic
- Read well at 24px, 48px, and 72pt

## Current Catalog Status (70 icons)

| Category | Count | Status |
|----------|-------|--------|
| Business | 15 | Core complete |
| Technology | 13 | Core complete |
| Finance | 13 | Core complete |
| Operations | 13 | Core complete |
| Security | 12 | Core complete |
| People | 2 | Minimal |
| Communication | 1 | Needs expansion |
| ESG | 1 | Needs expansion |

## Recommended Open-Source Icon Libraries

### 1. Tabler Icons (Primary Recommendation)

- **License**: MIT
- **URL**: https://tabler.io/icons / https://github.com/tabler/tabler-icons
- **Count**: 6,100+ icons
- **Grid**: 24×24
- **Stroke**: 2px (adjustable)
- **Style**: Outline-first, clean, modern
- **Compatibility**: Excellent - same grid and similar stroke style

**Why Tabler is ideal:**
- MIT license allows unrestricted use with simple attribution
- Same 24×24 grid as IconAid
- 2px stroke is close to IconAid's 1.6px (easy to adjust)
- Extensive business, finance, and technology categories
- Consistent design language across all icons

**Relevant Tabler icon categories for IconAid expansion:**
- **ESG/Sustainability**: solar-panel, solar-electricity, wind, wind-electricity, battery-eco, recycle, plant, tree, droplet, flame
- **Communication**: mail, message, phone, video, microphone, send, bell, notification
- **Documents**: file, folder, clipboard, bookmark, archive, paperclip
- **Charts/Business**: chart-bar, chart-line, chart-pie, chart-donut, chart-area, trending-up, trending-down
- **People**: user, users, user-plus, user-check, user-x, building, home

### 2. Lucide Icons

- **License**: ISC (with some Feather-derived icons under MIT)
- **URL**: https://lucide.dev / https://github.com/lucide-icons/lucide
- **Count**: 1,400+ icons
- **Grid**: 24×24
- **Stroke**: 2px
- **Style**: Clean outline, Feather-derived

**Notes:**
- Community-maintained fork of Feather Icons
- ISC license is permissive (similar to MIT)
- When using Feather-derived icons, MIT attribution required
- Strong consistency but smaller catalog than Tabler

### 3. Heroicons

- **License**: MIT
- **URL**: https://heroicons.com / https://github.com/tailwindlabs/heroicons
- **Count**: 300+ icons (in multiple sizes: 16, 20, 24)
- **Grid**: 24×24 (outline), 20×20 (solid), 16×16 (micro)
- **Stroke**: 1.5px
- **Style**: Minimal, balanced

**Notes:**
- Created by Tailwind Labs
- Smaller catalog but high optical quality
- Good for when cleaner, simpler alternatives are needed

### 4. Phosphor Icons

- **License**: MIT
- **URL**: https://phosphoricons.com / https://github.com/phosphor-icons
- **Count**: 1,200+ base icons × 6 weights = 7,200+ variants
- **Grid**: 256×256 (scales to any size)
- **Weights**: Thin, Light, Regular, Bold, Fill, Duotone

**Notes:**
- Multiple weight variants provide flexibility
- Good for presentation-friendly breadth
- Tags and categories built-in

### 5. Material Symbols

- **License**: Apache 2.0
- **URL**: https://fonts.google.com/icons
- **Count**: 2,500+ icons
- **Grid**: Various (typically 24×24)
- **Style**: Material Design 3

**Notes:**
- Apache 2.0 requires license notice preservation
- Very broad concept coverage
- Style may be too "app-like" for consulting use

### 6. Fluent UI System Icons

- **License**: MIT
- **URL**: https://github.com/microsoft/fluentui-system-icons
- **Count**: 2,500+ icons
- **Sizes**: 16, 20, 24, 28, 32, 48
- **Style**: Modern rounded forms

**Notes:**
- Familiar Microsoft metaphors
- Good for business/enterprise contexts

## Icons NOT Suitable

### Streamline
- **License**: Proprietary
- **Status**: Reference only, do NOT vendor or trace

### Consulting Firm Artwork (McKinsey, BCG, Bain, etc.)
- **License**: Proprietary
- **Status**: Visual language reference only, do NOT copy or derive

### Font Awesome Free (SVG icons)
- **License**: CC BY 4.0 (requires attribution)
- **Status**: Can use but requires visible attribution - may not fit presentation context

## Next Expansion Priorities (per EXPANSION_PLAN.md)

### Batch 6: Communication and Document Family
Target icons to find/adapt from open-source libraries:
- `presentation` - slides/presentation icon
- `mail` - email/envelope
- `chat` - speech bubble/message
- `notification` - bell/alert
- `video` - video camera/play
- `microphone` - mic/audio
- `send` - paper plane/arrow
- `knowledge` - book/lightbulb
- `folder` - folder/directory
- `clipboard` - clipboard/paste

### Batch 7: ESG Family
Target icons to find/adapt from open-source libraries:
- `solar` - solar panel/sun
- `wind` - wind turbine
- `water` - droplet/wave
- `circular-economy` - recycle/arrows
- `battery` - battery/power
- `waste` - trash/bin
- `climate` - thermometer/earth
- `biodiversity` - tree/plant/animal
- `reporting` - document/chart

## Attribution Requirements by License

### MIT License (Tabler, Heroicons, Phosphor, Fluent)
Minimal requirements:
- Include MIT license text in project
- Preserve copyright notice

### ISC License (Lucide)
Minimal requirements:
- Include ISC license text in project
- Preserve copyright notice

### Apache 2.0 (Material Symbols, Carbon)
More requirements:
- Include Apache 2.0 license notice
- State changes if modified
- Preserve copyright/patent/trademark notices

## Implementation Notes

When adapting icons from these libraries:

1. **Geometry Translation**: Convert SVG paths to IconAid primitives (lines, rects, ellipses, polylines)
2. **Stroke Adjustment**: Scale from 2px to 1.6px stroke width
3. **Grid Verification**: Ensure all coordinates align to IconAid's 24×24 grid
4. **Simplification**: Reduce complexity to 4-10 strokes if source is more detailed
5. **Style Matching**: Ensure round caps/joins, no filled regions except for small anchors
6. **Documentation**: Record source library in icon's designNotes field

## License Documentation Updates

When adding icons derived from external libraries, update `/shared/iconaid/LICENSES.md` with:
- Source library name
- License type
- Version/date sourced
- Number of icons derived
- Link to original license file
