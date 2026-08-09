import type { IconElement, IconShapePrimitive } from "./integrations";

// Google Slides has no freeform-path API, so a schema-3 icon's `path` and
// `polyline` elements cannot be inserted verbatim. They are flattened here into
// straight-line runs, which the insert path then draws as native Slides lines.
// That keeps what the user clicked and what lands on the slide identical - the
// previous code fell back to each icon's coarse `primitives` list, so 57 of the
// 70 catalog icons inserted as a visibly different shape than their thumbnail.
//
// Curves are sampled rather than approximated analytically: at 24x24 design
// scale a cubic never spans more than a few points, so uniform sampling is
// indistinguishable from an exact flattening and costs nothing to reason about.
// Six keeps the worst catalog icon around 40 grouped objects, which stays
// comfortable to ungroup and edit by hand.
const CUBIC_SEGMENTS = 6;

export interface IconPolyline {
  points: [number, number][];
  closed: boolean;
}

export interface FlattenedIcon {
  /** Rects and ellipses, which stay real Slides shapes. */
  shapes: IconShapePrimitive[];
  /** Everything stroke-shaped, reduced to point runs. */
  polylines: IconPolyline[];
}

function cubicPoint(
  from: [number, number], c1: [number, number], c2: [number, number], to: [number, number], t: number,
): [number, number] {
  const u = 1 - t;
  const a = u * u * u;
  const b = 3 * u * u * t;
  const c = 3 * u * t * t;
  const d = t * t * t;
  return [
    a * from[0] + b * c1[0] + c * c2[0] + d * to[0],
    a * from[1] + b * c1[1] + c * c2[1] + d * to[1],
  ];
}

function quadraticPoint(from: [number, number], control: [number, number], to: [number, number], t: number): [number, number] {
  const u = 1 - t;
  return [
    u * u * from[0] + 2 * u * t * control[0] + t * t * to[0],
    u * u * from[1] + 2 * u * t * control[1] + t * t * to[1],
  ];
}

interface PathCursor {
  command: string;
  values: number[];
}

const SUPPORTED_COMMANDS = "MmLlHhVvCcQqZz";

// Splits "M6 15 L9.5 12 C…" into command/argument pairs. Both cases are handled
// so a relative path from a future catalog revision does not silently misdraw.
// Any letter is captured as a command, including ones we do not support: matching
// only the supported set would let an arc's "A" fall through into the previous
// command's number list and be drawn as a line, which is worse than refusing it.
function tokenizePath(d: string): PathCursor[] {
  const tokens: PathCursor[] = [];
  const pattern = /([A-Za-z])([^A-Za-z]*)/g;
  let match = pattern.exec(d);
  while (match) {
    const command = match[1]!;
    if (!SUPPORTED_COMMANDS.includes(command)) throw new Error(`The icon path uses an unsupported command: ${command}`);
    const values = (match[2]?.match(/-?\d*\.?\d+/g) ?? []).map(Number);
    if (values.some((value) => !Number.isFinite(value))) throw new Error("The icon path contains an invalid number.");
    tokens.push({ command, values });
    match = pattern.exec(d);
  }
  if (!tokens.length) throw new Error("The icon path is empty.");
  return tokens;
}

/**
 * Flattens one SVG path string into straight-line runs. Supports the command
 * set the catalog actually uses (M, L, H, V, C, Z) plus Q and the relative
 * forms, and treats anything else as unsupported rather than guessing.
 */
export function flattenPath(d: string): IconPolyline[] {
  const runs: IconPolyline[] = [];
  let current: [number, number][] = [];
  let cursor: [number, number] = [0, 0];
  let start: [number, number] = [0, 0];

  const push = (closed: boolean): void => {
    if (current.length > 1) runs.push({ points: current, closed });
    current = [];
  };
  const lineTo = (point: [number, number]): void => {
    if (!current.length) current.push([cursor[0], cursor[1]]);
    current.push(point);
    cursor = point;
  };

  for (const { command, values } of tokenizePath(d)) {
    const relative = command === command.toLowerCase();
    const upper = command.toUpperCase();
    const base = (): [number, number] => (relative ? cursor : [0, 0]);

    if (upper === "Z") {
      if (current.length > 1) {
        current.push([start[0], start[1]]);
        cursor = [start[0], start[1]];
      }
      push(true);
      continue;
    }
    if (upper === "M") {
      if (values.length < 2) throw new Error("The icon path has an incomplete move command.");
      push(false);
      const origin = base();
      cursor = [origin[0] + values[0]!, origin[1] + values[1]!];
      start = [cursor[0], cursor[1]];
      // Extra coordinate pairs after a moveto are implicit linetos.
      for (let index = 2; index + 1 < values.length; index += 2) {
        const from = relative ? cursor : ([0, 0] as [number, number]);
        lineTo([from[0] + values[index]!, from[1] + values[index + 1]!]);
      }
      continue;
    }
    if (upper === "L") {
      if (values.length < 2) throw new Error("The icon path has an incomplete line command.");
      for (let index = 0; index + 1 < values.length; index += 2) {
        const from = base();
        lineTo([from[0] + values[index]!, from[1] + values[index + 1]!]);
      }
      continue;
    }
    if (upper === "H" || upper === "V") {
      if (!values.length) throw new Error("The icon path has an incomplete axis command.");
      for (const value of values) {
        const horizontal = upper === "H";
        const shift = relative ? (horizontal ? cursor[0] : cursor[1]) : 0;
        lineTo(horizontal ? [shift + value, cursor[1]] : [cursor[0], shift + value]);
      }
      continue;
    }
    if (upper === "C") {
      if (values.length < 6) throw new Error("The icon path has an incomplete curve command.");
      for (let index = 0; index + 5 < values.length; index += 6) {
        const from: [number, number] = [cursor[0], cursor[1]];
        const origin = base();
        const c1: [number, number] = [origin[0] + values[index]!, origin[1] + values[index + 1]!];
        const c2: [number, number] = [origin[0] + values[index + 2]!, origin[1] + values[index + 3]!];
        const to: [number, number] = [origin[0] + values[index + 4]!, origin[1] + values[index + 5]!];
        for (let step = 1; step <= CUBIC_SEGMENTS; step += 1) lineTo(cubicPoint(from, c1, c2, to, step / CUBIC_SEGMENTS));
      }
      continue;
    }
    if (upper === "Q") {
      if (values.length < 4) throw new Error("The icon path has an incomplete curve command.");
      for (let index = 0; index + 3 < values.length; index += 4) {
        const from: [number, number] = [cursor[0], cursor[1]];
        const origin = base();
        const control: [number, number] = [origin[0] + values[index]!, origin[1] + values[index + 1]!];
        const to: [number, number] = [origin[0] + values[index + 2]!, origin[1] + values[index + 3]!];
        for (let step = 1; step <= CUBIC_SEGMENTS; step += 1) lineTo(quadraticPoint(from, control, to, step / CUBIC_SEGMENTS));
      }
      continue;
    }
    throw new Error(`The icon path uses an unsupported command: ${command}`);
  }
  push(false);
  if (!runs.length) throw new Error("The icon path produced no drawable segments.");
  return runs;
}

/**
 * Splits an icon's drawable elements into the two things the Slides insert path
 * can create: native rect/ellipse shapes, and line runs. Rects and ellipses stay
 * real shapes so they remain fillable and resizable the way PowerPoint's
 * editable icons are.
 */
export function flattenIconElements(elements: readonly IconElement[]): FlattenedIcon {
  const shapes: IconShapePrimitive[] = [];
  const polylines: IconPolyline[] = [];
  for (const element of elements) {
    if (element.kind === "path") {
      polylines.push(...flattenPath(element.d));
      continue;
    }
    if (element.kind === "polyline") {
      const points = element.points.map((point) => [point[0], point[1]] as [number, number]);
      if (element.closed && points.length > 1) points.push([points[0]![0], points[0]![1]]);
      polylines.push({ points, closed: element.closed });
      continue;
    }
    if (element.kind === "line") {
      polylines.push({ points: [[element.x1, element.y1], [element.x2, element.y2]], closed: false });
      continue;
    }
    shapes.push(element);
  }
  if (!shapes.length && !polylines.length) throw new Error("The icon contains no drawable elements.");
  return { shapes, polylines };
}

/** Straight segments an icon needs, in insertion order. */
export function polylineSegments(polylines: readonly IconPolyline[]): { x1: number; y1: number; x2: number; y2: number }[] {
  const segments: { x1: number; y1: number; x2: number; y2: number }[] = [];
  for (const run of polylines) {
    for (let index = 0; index + 1 < run.points.length; index += 1) {
      const [x1, y1] = run.points[index]!;
      const [x2, y2] = run.points[index + 1]!;
      if (Math.abs(x1 - x2) < 1e-6 && Math.abs(y1 - y2) < 1e-6) continue;
      segments.push({ x1, y1, x2, y2 });
    }
  }
  return segments;
}
