import type { IconPolyline } from "./icon-path";

/**
 * Solid icons (all of Bootstrap, plus any -solid/-mini variant) are drawn as
 * filled regions rather than centreline strokes. Google Slides has no freeform
 * path, so a filled glyph cannot be handed to it directly, and stroking its
 * outline instead would render it hollow - a filled "person" would come out as
 * an outline of a person, which is the wrong icon.
 *
 * The region is therefore scan-converted: horizontal slices across the shape,
 * each becoming a filled rectangle. That is the same trick the chart builders
 * use for areas Slides will not fill, and it reproduces holes correctly because
 * the spans follow SVG's even-odd rule - which is exactly what the PowerPoint
 * task pane asks for when it renders these icons with fill-rule="evenodd".
 */

/** Which icons are solid regions. Mirrors is_filled() in build_iconaid_web.py. */
export function isFilledIcon(id: string): boolean {
  return id.startsWith("bootstrap-") || /-(solid|mini)$/.test(id);
}

export interface FillSpan {
  top: number;
  height: number;
  left: number;
  width: number;
}

function edgesOf(polygons: readonly IconPolyline[]): { x1: number; y1: number; x2: number; y2: number }[] {
  const edges: { x1: number; y1: number; x2: number; y2: number }[] = [];
  for (const polygon of polygons) {
    const points = polygon.points;
    if (points.length < 2) continue;
    for (let index = 0; index + 1 < points.length; index += 1) {
      edges.push({ x1: points[index]![0], y1: points[index]![1], x2: points[index + 1]![0], y2: points[index + 1]![1] });
    }
    // A fill always treats each subpath as closed, whether or not it ended in Z.
    const first = points[0]!;
    const last = points[points.length - 1]!;
    if (Math.abs(first[0] - last[0]) > 1e-9 || Math.abs(first[1] - last[1]) > 1e-9) {
      edges.push({ x1: last[0], y1: last[1], x2: first[0], y2: first[1] });
    }
  }
  return edges;
}

/**
 * Scan-converts closed polygons into filled rectangles.
 *
 * `slices` trades object count against smoothness: every slice becomes at least
 * one rectangle on the slide, so this stays deliberately modest.
 */
export function fillSpans(polygons: readonly IconPolyline[], slices = 56): FillSpan[] {
  if (slices < 1) throw new Error("A filled icon needs at least one slice.");
  const edges = edgesOf(polygons);
  if (!edges.length) return [];
  const top = Math.min(...edges.map((edge) => Math.min(edge.y1, edge.y2)));
  const bottom = Math.max(...edges.map((edge) => Math.max(edge.y1, edge.y2)));
  const height = bottom - top;
  if (!(height > 0)) return [];

  const step = height / slices;
  const spans: FillSpan[] = [];
  for (let index = 0; index < slices; index += 1) {
    // Sample down the middle of the slice so thin features are not clipped away.
    const y = top + (index + 0.5) * step;
    const crossings: number[] = [];
    for (const edge of edges) {
      const { x1, y1, x2, y2 } = edge;
      if (y1 === y2) continue;
      const lower = Math.min(y1, y2);
      const upper = Math.max(y1, y2);
      // Half-open test: a vertex shared by two edges must count once, not twice.
      if (y < lower || y >= upper) continue;
      crossings.push(x1 + ((y - y1) / (y2 - y1)) * (x2 - x1));
    }
    if (crossings.length < 2) continue;
    crossings.sort((a, b) => a - b);
    for (let pair = 0; pair + 1 < crossings.length; pair += 2) {
      const left = crossings[pair]!;
      const width = crossings[pair + 1]! - left;
      if (width <= 1e-6) continue;
      // Slices overlap slightly so no hairline gap shows between them.
      spans.push({ top: y - step / 2, height: step * 1.04, left, width });
    }
  }
  return spans;
}
