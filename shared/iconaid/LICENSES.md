# IconAid Licensing

IconAid's curated pilot artwork is original Slide Aid geometry and is distributed under the project MIT license in `shared/iconaid/LICENSE`.

The full offline search library also redistributes normalized SVG path data from
the reviewed open-source projects listed in `shared/iconaid/sources.json`. Each
record retains its source, source version, upstream URL, and license. Geometry is
scaled and converted to the common IconAid path representation; those technical
modifications do not replace or remove the upstream license.

## Searchable Icon Database

The IconAid project maintains a searchable offline database of **54,250+ insertable icons**
from approved open-source libraries. The catalog is generated from pinned npm
package versions, enriched with search terms, normalized to a 24×24 coordinate
system, and shared by the PowerPoint and Google Slides integrations.

### Database Files

| Source | Icons | License | File |
|--------|-------|---------|------|
| Tabler Icons | 5,130 | MIT | `external-sources/tabler-icons-normalized.json` |
| Lucide | 1,993 | ISC | `external-sources/lucide-icons-normalized.json` |
| Heroicons | 324 | MIT | `external-sources/heroicons-normalized.json` |
| Phosphor Icons | 1,512 | MIT | `external-sources/phosphor-icons-normalized.json` |
| Bootstrap Icons | 2,078 | MIT | `external-sources/bootstrap-icons-normalized.json` |
| Iconoir | 1,676 | MIT | `external-sources/iconoir-icons-normalized.json` |
| Hugeicons Free | 5,080 | MIT | `external-sources/hugeicons-icons-normalized.json` |
| IconPark Outline | 2,586 | Apache-2.0 | `external-sources/icon-park-outline-icons-normalized.json` |
| MingCute Core | 3,334 | Apache-2.0 | `external-sources/mingcute-icons-normalized.json` |
| Carbon Icons | 2,713 | Apache-2.0 | `external-sources/carbon-icons-normalized.json` |
| Material Symbols | 16,278 | Apache-2.0 | `external-sources/material-symbols-icons-normalized.json` |
| Fluent System Icons (24px regular and filled) | 5,136 | MIT | `external-sources/fluent-icons-normalized.json` |
| Simple Icons | 3,723 | CC0-1.0 | `external-sources/simple-icons-normalized.json` |
| Health Icons | 2,691 | MIT | `external-sources/healthicons-normalized.json` |
| **Combined source records** | **54,254** | Various | `unified-catalog.json` |

Four legacy Lucide records contain no insertable path segments, so the host
catalogs contain 54,250 icons. Source records are retained for auditability.

### Search Quality

All icons are enriched at build time with their complete upstream terms plus
domain, business, technical, visual, abbreviation, and synonym mappings. The
normalizer no longer truncates icons to 20 tags, and the shipped catalogs no
longer truncate them to 12 tags.

Tags include business concepts (ROI, merger, agile, kanban), technical terms (API, cloud, devops), 
ESG terminology (carbon, sustainability, renewable), and visual descriptors (animal, shape, color).

### Search Tools

```bash
# Search for icons
python3 scripts/build_icon_search_index.py --search "analytics chart"

# List categories
python3 scripts/build_icon_search_index.py --list-categories

# Get expansion recommendations
python3 scripts/recommend_icons.py --expansion-batch 6  # Communication/Document
python3 scripts/recommend_icons.py --expansion-batch 7  # ESG
```

## Approved Source Libraries

The source manifest is the authoritative allowlist. Adding a source requires a
permissive license, upstream and package provenance, a defined render mode, and
tests covering import and search behavior. Sources outside the manifest are not
part of the distributable library.

The expansion adds:

- **Iconoir** — MIT, https://github.com/iconoir-icons/iconoir
- **Hugeicons Free** — MIT, https://github.com/hugeicons/hugeicons
- **IconPark Outline** — Apache-2.0, https://github.com/bytedance/IconPark
- **MingCute Core Regular and Filled** — Apache-2.0,
  https://github.com/mingcute-design/mingcute-icons
- **Carbon Icons** — Apache-2.0,
  https://github.com/carbon-design-system/carbon
- **Material Symbols** — Apache-2.0,
  https://github.com/google/material-design-icons
- **Fluent System Icons, canonical 24px regular and filled variants** — MIT,
  https://github.com/microsoft/fluentui-system-icons
- **Simple Icons** — CC0-1.0,
  https://github.com/simple-icons/simple-icons
- **Health Icons** — MIT,
  https://github.com/resolvetosavelives/healthicons

Only MingCute's public Core collection is included. MingCute Pro artwork and
packages are not approved.

Simple Icons contains third-party brand marks. Its icon data is CC0-1.0, but
names and logos can remain subject to trademark rules. Inclusion does not imply
endorsement, and users are responsible for brand-compliant use.

### Tabler Icons (Recommended Primary Source)

- **License**: MIT
- **URL**: https://github.com/tabler/tabler-icons
- **Grid**: 24×24
- **Stroke**: 2px (adjustable)
- **Count**: 6,100+ icons
- **Copyright**: Copyright (c) 2020-2026 Paweł Kuna

The MIT License permits use, modification, and redistribution provided the copyright notice and license are included. When deriving icons from Tabler:
1. Record "Derived from Tabler Icons" in the icon's `designNotes` field
2. The original Tabler MIT license is satisfied by this project's MIT license

### Phosphor Icons

- **License**: MIT
- **URL**: https://phosphoricons.com/
- **Repository**: https://github.com/phosphor-icons/core
- **Grid**: 256×256 (scalable)
- **Count**: 1,500+ icons (6 weights)
- **Copyright**: Copyright (c) 2020-2026 Helena Zhang & Tobias Fried

When deriving from Phosphor Icons:
1. Record "Derived from Phosphor Icons" in the icon's `designNotes` field
2. The original Phosphor MIT license is satisfied by this project's MIT license

### Heroicons

- **License**: MIT
- **URL**: https://heroicons.com/
- **Repository**: https://github.com/tailwindlabs/heroicons
- **Grid**: 24×24 (outline) / 20×20 (solid)
- **Count**: 320+ icons
- **Copyright**: Copyright (c) 2020-2026 Tailwind Labs, Inc.

When deriving from Heroicons:
1. Record "Derived from Heroicons" in the icon's `designNotes` field
2. The original Heroicons MIT license is satisfied by this project's MIT license

### Other Approved MIT-Licensed Libraries

- **Microsoft Fluent UI System Icons**: MIT, https://github.com/microsoft/fluentui-system-icons
- **Iconoir**: MIT, https://github.com/iconoir-icons/iconoir
- **Hugeicons Free**: MIT, https://github.com/hugeicons/hugeicons
- **Health Icons**: MIT, https://github.com/resolvetosavelives/healthicons

### Bootstrap Icons

- **License**: MIT
- **URL**: https://icons.getbootstrap.com/
- **Repository**: https://github.com/twbs/icons
- **Grid**: Various (scalable)
- **Count**: 2,000+ icons
- **Copyright**: Copyright (c) 2019-2026 The Bootstrap Authors

When deriving from Bootstrap Icons:
1. Record "Derived from Bootstrap Icons" in the icon's `designNotes` field
2. The original Bootstrap MIT license is satisfied by this project's MIT license

### ISC-Licensed Libraries

- **Lucide**: ISC (with selected Feather-derived icons under MIT), https://github.com/lucide-icons/lucide
  - When using Feather-derived icons, MIT attribution is also required

### Apache 2.0 Licensed Libraries

- **Material Symbols and Material Icons**: Apache 2.0, https://fonts.google.com/icons
- **IBM Carbon Icons**: Apache 2.0, https://github.com/carbon-design-system/carbon
- **IconPark Outline**: Apache 2.0, https://github.com/bytedance/IconPark
- **MingCute Core**: Apache 2.0, https://github.com/mingcute-design/mingcute-icons

When deriving from Apache 2.0 sources:
- Preserve the Apache 2.0 notice in this file
- State modifications made
- The Apache 2.0 license notice is included below for reference

## Libraries Requiring Attribution (Use with Caution)

- **Font Awesome Free**: SVG icons are CC BY 4.0 (requires visible attribution), fonts are SIL OFL 1.1, code is MIT. Only use if attribution can be provided in presentation context.

## Prohibited Sources

- **Remix Icon (current releases)**: Remix Icon License v1.0 prohibits using
  the icons to create a competing icon library. Do not ingest current releases,
  even if an older aggregator record still labels the project Apache-2.0.
- **Streamline**: Proprietary license. Do not vendor, trace, or derive.
- **Consulting Firm Artwork**: McKinsey, BCG, Bain, Deloitte, PwC, and Accenture report graphics are proprietary. Visual language reference only - do not trace, copy, or derive matching artwork.

## License Texts

### Tabler Icons MIT License

```
MIT License

Copyright (c) 2020-2026 Paweł Kuna

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
```

### Phosphor Icons MIT License

```
MIT License

Copyright (c) 2020-2026 Helena Zhang & Tobias Fried

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
```

### Heroicons MIT License

```
MIT License

Copyright (c) 2020-2026 Tailwind Labs, Inc.

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
```

### Apache License 2.0 (for Material/Carbon icons if used)

When deriving icons from Apache 2.0 sources, include this notice:
```
Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
```
