# IconAid Licensing

IconAid pilot artwork is original Slide Aid geometry and is distributed under the project MIT license in `shared/iconaid/LICENSE`.

No third-party SVG path data has been copied into the pilot catalog.

## Searchable Icon Database

The IconAid project maintains a searchable database of **7,500+ icons** from approved open-source libraries. These icons are stored in normalized JSON format for search and discovery. When adapting icons for IconAid, the geometry must be redrawn to match IconAid design specifications.

### Database Files

| Source | Icons | License | File |
|--------|-------|---------|------|
| Tabler Icons | 3,990 | MIT | `external-sources/tabler-icons-normalized.json` |
| Bootstrap Icons | 2,029 | MIT | `external-sources/bootstrap-icons-normalized.json` |
| Phosphor Icons | 1,480 | MIT | `external-sources/phosphor-icons-normalized.json` |
| Lucide | 920 | ISC | `external-sources/lucide-icons-normalized.json` |
| Heroicons | 324 | MIT | `external-sources/heroicons-normalized.json` |
| **Combined** | **7,519** | Various | `external-sources/combined-search-index.json` |

### Search Quality

All icons have been enriched with semantic tags for better searchability:
- **Total searchable tags**: 113,000+
- **Average tags per icon**: 15.1
- **Icons with 5+ tags**: 94%
- **Icons with 10+ tags**: 84%
- **Icons with 15+ tags**: 54%

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

The following open-source icon libraries are approved for deriving IconAid icons. When adapting icons from these sources, the geometry must be redrawn to match IconAid design specifications (24×24 grid, 1.6px stroke, round caps/joins, simplified to 4-10 elements).

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

When deriving from Apache 2.0 sources:
- Preserve the Apache 2.0 notice in this file
- State modifications made
- The Apache 2.0 license notice is included below for reference

## Libraries Requiring Attribution (Use with Caution)

- **Font Awesome Free**: SVG icons are CC BY 4.0 (requires visible attribution), fonts are SIL OFL 1.1, code is MIT. Only use if attribution can be provided in presentation context.

## Prohibited Sources

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
