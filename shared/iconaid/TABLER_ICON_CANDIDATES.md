# Tabler Icons Candidates for IconAid Expansion

This document lists specific Tabler icons suitable for adapting into IconAid, organized by the expansion batch priorities from EXPANSION_PLAN.md.

All icons below are from Tabler Icons (MIT License, https://github.com/tabler/tabler-icons).

## How to Use This Document

1. Each icon entry shows the Tabler icon name and SVG path data
2. Review the original icon at `https://tabler.io/icons/icon/{icon-name}`
3. Simplify/redraw to match IconAid design system (1.6px stroke, round caps, 4-10 elements)
4. Add to `build_icon_catalog.py` with `designNotes: "Derived from Tabler Icons"`

---

## Batch 6: Communication and Document Family

### presentation
```
Tabler: presentation
Category: Document
SVG:
  <path d="M3 4l18 0" />
  <path d="M4 4v10a2 2 0 0 0 2 2h12a2 2 0 0 0 2 -2v-10" />
  <path d="M12 16l0 4" />
  <path d="M9 20l6 0" />
  <path d="M8 12l3 -3l2 2l3 -3" />

Adaptation notes:
- Clean presentation screen metaphor
- Has embedded chart which fits consulting context
- May simplify chart line to match IconAid style
```

### mail
```
Tabler: mail
Category: Communication
SVG:
  <path d="M3 7a2 2 0 0 1 2 -2h14a2 2 0 0 1 2 2v10a2 2 0 0 1 -2 2h-14a2 2 0 0 1 -2 -2v-10" />
  <path d="M3 7l9 6l9 -6" />

Adaptation notes:
- Classic envelope with V-fold
- Simple, 2-element design
- Excellent fit for IconAid
```

### chat (message)
```
Tabler: message
Category: Communication
SVG:
  <path d="M8 9h8" />
  <path d="M8 13h6" />
  <path d="M18 4a3 3 0 0 1 3 3v8a3 3 0 0 1 -3 3h-5l-5 3v-3h-2a3 3 0 0 1 -3 -3v-8a3 3 0 0 1 3 -3h12" />

Adaptation notes:
- Speech bubble with text lines
- Rounded corners (use rx on rect)
- May simplify to remove interior text lines
```

### notification (bell)
```
Tabler: bell
Category: System
SVG:
  <path d="M10 5a2 2 0 1 1 4 0a7 7 0 0 1 4 6v3a4 4 0 0 0 2 3h-16a4 4 0 0 0 2 -3v-3a7 7 0 0 1 4 -6" />
  <path d="M9 17v1a3 3 0 0 0 6 0v-1" />

Adaptation notes:
- Classic bell silhouette
- May need path simplification to primitives
- Clear recognizable metaphor
```

### video
```
Tabler: video
Category: Media
SVG:
  <path d="M15 10l4.553 -2.276a1 1 0 0 1 1.447 .894v6.764a1 1 0 0 1 -1.447 .894l-4.553 -2.276v-4" />
  <path d="M3 8a2 2 0 0 1 2 -2h8a2 2 0 0 1 2 2v8a2 2 0 0 1 -2 2h-8a2 2 0 0 1 -2 -2l0 -8" />

Adaptation notes:
- Camera with viewfinder shape
- Two clear elements (screen + viewfinder)
- Good for video conferencing context
```

### microphone
```
Tabler: microphone
Category: Media
SVG:
  <path d="M9 5a3 3 0 0 1 3 -3a3 3 0 0 1 3 3v5a3 3 0 0 1 -3 3a3 3 0 0 1 -3 -3l0 -5" />
  <path d="M5 10a7 7 0 0 0 14 0" />
  <path d="M8 21l8 0" />
  <path d="M12 17l0 4" />

Adaptation notes:
- Classic microphone silhouette
- 4 elements, good complexity
- Recognizable audio metaphor
```

### send
```
Tabler: send
Category: Communication
SVG:
  <path d="M10 14l11 -11" />
  <path d="M21 3l-6.5 18a.55 .55 0 0 1 -1 0l-3.5 -7l-7 -3.5a.55 .55 0 0 1 0 -1l18 -6.5" />

Adaptation notes:
- Paper airplane metaphor
- Slightly complex path, may simplify
- Strong "send/share" recognition
```

### knowledge (book)
```
Tabler: book
Category: Document
SVG:
  <path d="M3 19a9 9 0 0 1 9 0a9 9 0 0 1 9 0" />
  <path d="M3 6a9 9 0 0 1 9 0a9 9 0 0 1 9 0" />
  <path d="M3 6l0 13" />
  <path d="M12 6l0 13" />
  <path d="M21 6l0 13" />

Adaptation notes:
- Open book with spine
- 5 elements, perfect complexity
- Good "knowledge/learning" metaphor
```

### folder
```
Tabler: folder
Category: Document
SVG:
  <path d="M5 4h4l3 3h7a2 2 0 0 1 2 2v8a2 2 0 0 1 -2 2h-14a2 2 0 0 1 -2 -2v-11a2 2 0 0 1 2 -2" />

Adaptation notes:
- Single path folder shape
- Very clean, minimal design
- May need to convert to rect + lines for primitives
```

### clipboard
```
Tabler: clipboard
Category: Document
SVG:
  <path d="M9 5h-2a2 2 0 0 0 -2 2v12a2 2 0 0 0 2 2h10a2 2 0 0 0 2 -2v-12a2 2 0 0 0 -2 -2h-2" />
  <path d="M9 5a2 2 0 0 1 2 -2h2a2 2 0 0 1 2 2a2 2 0 0 1 -2 2h-2a2 2 0 0 1 -2 -2" />

Adaptation notes:
- Clipboard with clip detail
- 2 paths, clean design
- Good "checklist/tasks" metaphor
```

---

## Batch 7: ESG Family

### solar (solar-panel)
```
Tabler: solar-panel
Category: System
Tags: energy, sun, power, ecology, electricity, solar
SVG:
  <path d="M4.28 14h15.44a1 1 0 0 0 .97 -1.243l-1.5 -6a1 1 0 0 0 -.97 -.757h-12.44a1 1 0 0 0 -.97 .757l-1.5 6a1 1 0 0 0 .97 1.243" />
  <path d="M4 10h16" />
  <path d="M10 6l-1 8" />
  <path d="M14 6l1 8" />
  <path d="M12 14v4" />
  <path d="M7 18h10" />

Adaptation notes:
- Tilted solar panel with grid lines
- 6 elements, acceptable complexity
- Strong renewable energy metaphor
- May simplify to rect with lines
```

### wind (windmill)
```
Tabler: windmill
Category: Map
Tags: generate, power, blade, energy, electricity
SVG:
  <path d="M12 12c2.76 0 5 -2.01 5 -4.5s-2.24 -4.5 -5 -4.5v9" />
  <path d="M12 12c0 2.76 2.01 5 4.5 5s4.5 -2.24 4.5 -5h-9" />
  <path d="M12 12c-2.76 0 -5 2.01 -5 4.5s2.24 4.5 5 4.5v-9" />
  <path d="M12 12c0 -2.76 -2.01 -5 -4.5 -5s-4.5 2.24 -4.5 5h9" />

Adaptation notes:
- Wind turbine blades (pinwheel)
- 4 symmetrical paths
- May need significant simplification
- Alternative: use simpler wind-electricity icon
```

### wind-electricity (alternative)
```
Tabler: wind-electricity
Category: Nature
Tags: turbine, renewable, sustainable, breeze, airflow, power
SVG:
  <path d="M20 7l-3 5h4l-3 5" />
  <path d="M3 16h4a2 2 0 1 1 0 4" />
  <path d="M3 12h8a2 2 0 1 0 0 -4" />
  <path d="M3 8h3a2 2 0 1 0 0 -4" />

Adaptation notes:
- Abstract wind lines
- Simpler than windmill, 4 paths
- Good "airflow/wind" metaphor
```

### water (droplet)
```
Tabler: droplet
Category: Design
Tags: water, rain, liquid
SVG:
  <path d="M7.502 19.423c2.602 2.105 6.395 2.105 8.996 0c2.602 -2.105 3.262 -5.708 1.566 -8.546l-4.89 -7.26c-.42 -.625 -1.287 -.803 -1.936 -.397a1.376 1.376 0 0 0 -.41 .397l-4.893 7.26c-1.695 2.838 -1.035 6.441 1.567 8.546" />

Adaptation notes:
- Single droplet path
- Complex Bezier curves need simplification
- May use ellipse + pointed top instead
```

### circular-economy (recycle)
```
Tabler: recycle
Category: Symbols
Tags: recyclable, reuse, waste
SVG:
  <path d="M12 17l-2 2l2 2" />
  <path d="M10 19h9a2 2 0 0 0 1.75 -2.75l-.55 -1" />
  <path d="M8.536 11l-.732 -2.732l-2.732 .732" />
  <path d="M7.804 8.268l-4.5 7.794a2 2 0 0 0 1.506 2.89l1.141 .024" />
  <path d="M15.464 11l2.732 .732l.732 -2.732" />
  <path d="M18.196 11.732l-4.5 -7.794a2 2 0 0 0 -3.256 -.14l-.591 .976" />

Adaptation notes:
- Classic recycle/Möbius loop symbol
- 6 paths with arrows
- Strong circular economy metaphor
- May simplify arrow heads
```

### battery (battery-eco)
```
Tabler: battery-eco
Category: Devices
Tags: ecology, charge, energy, power
SVG:
  <path d="M4 9a2 2 0 0 1 2 -2h11a2 2 0 0 1 2 2v.5a.5 .5 0 0 0 .5 .5a.5 .5 0 0 1 .5 .5v3a.5 .5 0 0 1 -.5 .5a.5 .5 0 0 0 -.5 .5v.5a2 2 0 0 1 -2 2h-5.5" />
  <path d="M3 16.143c0 -2.84 2.09 -5.143 4.667 -5.143h2.333v.857c0 2.84 -2.09 5.143 -4.667 5.143h-2.333v-.857" />
  <path d="M3 20v-3" />

Adaptation notes:
- Battery with leaf/eco indicator
- Combines energy + sustainability
- Good complexity but may need simplification
```

### waste (trash)
```
Tabler: trash
Category: System
Tags: garbage, delete, remove, bin
SVG:
  <path d="M4 7l16 0" />
  <path d="M10 11l0 6" />
  <path d="M14 11l0 6" />
  <path d="M5 7l1 12a2 2 0 0 0 2 2h8a2 2 0 0 0 2 -2l1 -12" />
  <path d="M9 7v-3a1 1 0 0 1 1 -1h4a1 1 0 0 1 1 1v3" />

Adaptation notes:
- Classic trash can
- 5 paths, good complexity
- Clear waste/disposal metaphor
```

### climate (temperature)
```
Tabler: temperature
Category: Weather
Tags: weather, celsius, fahrenheit, cold, hot
SVG:
  <path d="M10 13.5a4 4 0 1 0 4 0v-8.5a2 2 0 0 0 -4 0v8.5" />
  <path d="M10 9l4 0" />

Adaptation notes:
- Thermometer shape
- 2 paths, very simple
- Good climate/temperature metaphor
```

### biodiversity (plant/tree)
```
Tabler: plant
Category: Nature
Tags: nature, green, flower, pot, tree, leaf
SVG:
  <path d="M7 15h10v4a2 2 0 0 1 -2 2h-6a2 2 0 0 1 -2 -2v-4" />
  <path d="M12 9a6 6 0 0 0 -6 -6h-3v2a6 6 0 0 0 6 6h3" />
  <path d="M12 11a6 6 0 0 1 6 -6h3v1a6 6 0 0 1 -6 6h-3" />
  <path d="M12 15l0 -6" />

Adaptation notes:
- Plant in pot with leaves
- 4 paths, good complexity
- Good growth/nature/biodiversity metaphor
```

### reporting (report)
```
Tabler: report
Category: Document
Tags: time, timesheet, analysis, results, business
SVG:
  <path d="M8 5h-2a2 2 0 0 0 -2 2v12a2 2 0 0 0 2 2h5.697" />
  <path d="M18 14v4h4" />
  <path d="M18 11v-4a2 2 0 0 0 -2 -2h-2" />
  <path d="M8 5a2 2 0 0 1 2 -2h2a2 2 0 0 1 2 2a2 2 0 0 1 -2 2h-2a2 2 0 0 1 -2 -2" />
  <path d="M14 18a4 4 0 1 0 8 0a4 4 0 1 0 -8 0" />
  <path d="M8 11h4" />
  <path d="M8 15h3" />

Adaptation notes:
- Document with clock/time indicator
- 7 paths, may need simplification
- Good ESG reporting metaphor
- Alternative: simpler document + chart combo
```

---

## Additional Business/Consulting Icons Worth Considering

### briefcase
```
Tabler: briefcase
Good for: work, consulting, portfolio
Simple rect + handle design
```

### chart-bar / chart-line / chart-pie
```
Tabler: chart-bar, chart-line, chart-pie
Good for: analytics, metrics, KPIs
Various chart visualizations
```

### users-group
```
Tabler: users-group
Good for: team, stakeholders, organization
Multiple people silhouette
```

### building
```
Tabler: building
Good for: office, corporate, headquarters
Building silhouette with windows
```

### hand-shake
```
Tabler: hand-shake
Good for: partnership, deal, agreement
Two hands meeting
```

### bulb
```
Tabler: bulb
Good for: idea, innovation, insight
Lightbulb silhouette
```

### compass
```
Tabler: compass
Good for: direction, strategy, navigation
Compass with needle
```

---

## Implementation Workflow

1. **Select icon** from candidates above
2. **Download SVG** from `https://raw.githubusercontent.com/tabler/tabler-icons/main/icons/outline/{name}.svg`
3. **Simplify paths** to IconAid primitives (line, rect, ellipse, polyline)
4. **Adjust coordinates** to fit 24×24 grid with 2-unit safe area
5. **Scale stroke** from 2px to 1.6px equivalent
6. **Add to build_icon_catalog.py** with proper metadata
7. **Run build** and verify contact sheet
8. **Test insertion** in PowerPoint and Google Slides
