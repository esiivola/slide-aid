# IconAid Design System

IconAid is a monochrome management-consulting icon system for presentation work. The icons should read as quiet exhibit symbols, not app UI decoration and not miniature diagrams assembled from office shapes.

## Grid And Safe Area

- Master grid: 24 by 24 units.
- Safe area: 2 units on every side. Artwork may extend into the safe area only for optical compensation, such as arrowheads, round caps, or wide circular forms.
- Preferred drawing area: 4 to 20 units.
- Coordinates should usually land on whole or half units. Smaller decimals are allowed only for optical centering, not as a default construction habit.

## Stroke

- Default stroke: 1.6 units in SVG preview.
- Line caps: round.
- Joins: round.
- Strokes must feel equal at 24 px, 48 px, and 72 pt. Filled accents are allowed only when they stabilize small-size recognition.
- Avoid doubled strokes from overlapping primitives. If two shapes meet, they should either deliberately share a seam or maintain clear negative space.

## Corners And Curves

- Rectangular objects use restrained corners: 0 to 2 units depending on the metaphor.
- Buildings, documents, devices, and process blocks should have low-radius corners.
- People, clouds, leaves, and speech bubbles need real curves in preview geometry. Do not approximate organic forms with jagged polygon chains unless the fallback insertion layer requires it.

## Optical Alignment

- Center icons by apparent mass, not by bounding-box arithmetic.
- Circular icons may be slightly larger than rectangular icons to appear equal.
- Dense interiors should be pulled inward to preserve exterior clarity.
- Arrowheads should land on the grid and leave enough space from containers to avoid tangencies.

## Filled Versus Outline

- The default language is outline.
- Filled elements are limited to small anchors: chart bars, data points, status dots, or stable interiors.
- No duotone, gradients, shadows, or decorative fills.
- Icons must recolor cleanly to dark ink, white, and common consulting blues.

## Complexity

- Target: 4 to 10 visible strokes or shapes.
- Maximum: 16 visible elements unless the concept genuinely requires more.
- Avoid labels, letters, tiny people, tiny badges, or duplicated micro-symbols.
- If a concept cannot be recognized at 24 px, simplify or choose a different metaphor.

## Family Treatments

- People: circular heads, restrained shoulder curves, no facial details.
- Organizations: same people geometry plus hierarchy or grouping only when needed.
- Documents: consistent page height, folded corner, three or fewer content lines.
- Devices and technology: simple frames, sparse internals, no decorative ports.
- Data and finance: axes, bars, ledgers, currency marks, and cylinders use shared spacing rhythm.
- Operations: rectangular process blocks, factories, nodes, and logistics flows use clear orthogonal structure.
- ESG: organic forms may use curves, but remain monochrome and low detail.
- Security: shield is the family anchor; lock/warning/control details sit inside with clear negative space.
- Arrows: short, deliberate, grid-aligned. Avoid chart junk and oversized directional clutter.
- Badges and status variants: mechanical base-plus-badge variants are disabled. A status icon must be composed as a new concept or omitted.

## Rendering Model

Schema 3 separates artwork from insertion fallback:

- `elements`: the artwork. Paths and polylines are allowed here for clean curves, and this is what both the preview and the insert now draw.
- `primitives`: a coarse rect/ellipse/line approximation, kept for compatibility.

`primitives` is no longer an insertion path. It existed as fallback geometry while
native path insertion was unproven, and the two layers drifting apart was a real
defect: Google Slides previewed `elements` but inserted `primitives`, so 57 of the
70 pilot icons landed as a visibly different shape than the thumbnail the user
clicked. Both hosts now render `elements` — PowerPoint through its freeform
convert, Google Slides by flattening paths into native lines
(`apps/google-slides/src/core/icon-path.ts`).

Keep the two visually close anyway while `primitives` still ships, and review
insertion quality from real host output before expanding beyond the pilot.
