/**
 * A stand-in for the Apps Script host.
 *
 * The geometry and chart maths are covered by pure unit tests, but everything
 * that actually reaches Google - the shape batches, alt-text tags, image
 * insertion, property storage - only runs inside Apps Script. This fake records
 * what the code asks the host to do, so those paths can be exercised offline and
 * asserted on: a wrong field mask or a missing request shows up here rather than
 * on a live slide.
 *
 * It is deliberately thin. It models what Slide Aid uses and nothing else.
 */
import { resetCatalogCache } from "../src/slides/icon-catalog";

export interface RecordedRequest {
  [key: string]: Record<string, unknown>;
}

export interface FakeElement {
  id: string;
  type: string;
  left: number;
  top: number;
  width: number;
  height: number;
  rotation: number;
  title: string;
  description: string;
  removed: boolean;
  selected: boolean;
  image?: { bytes: number[]; contentType: string };
  /** TABLE only: the grid a chart builder reads. */
  cells?: string[][];
  /** TABLE only: geometry Snap to Table measures against. */
  columnWidths?: number[];
  rowHeights?: number[];
  /** SHAPE only: text runs, so font scaling and text tools have something real. */
  runs?: { text: string; fontSize: number }[];
}

export interface HostState {
  /** Every batchUpdate request, flattened in submission order. */
  requests: RecordedRequest[];
  /** Elements the code created or was given, by object id. */
  elements: Map<string, FakeElement>;
  /** Document + user properties, as the storage layer sees them. */
  documentProperties: Map<string, string>;
  userProperties: Map<string, string>;
  /** Contents of the project's icon data files, keyed by file name. */
  projectFiles: Map<string, string>;
  batchUpdates: number;
  /** Font sizes per element after any scaling, for assertions. */
  fontSizes: Map<string, number[]>;
}

let uuidCounter = 0;

function element(state: HostState, partial: Partial<FakeElement> & { id: string; type: string }): FakeElement {
  const created: FakeElement = {
    left: 0, top: 0, width: 10, height: 10, rotation: 0,
    title: "", description: "", removed: false, selected: false,
    ...partial,
  };
  state.elements.set(created.id, created);
  return created;
}

function wrapElement(state: HostState, item: FakeElement): Record<string, unknown> {
  const api: Record<string, unknown> = {
    getObjectId: () => item.id,
    getPageElementType: () => item.type,
    getLeft: () => item.left,
    getTop: () => item.top,
    getWidth: () => item.width,
    getHeight: () => item.height,
    getRotation: () => item.rotation,
    getTitle: () => item.title,
    getDescription: () => item.description,
    setTitle: (value: string) => { item.title = value; return api; },
    setDescription: (value: string) => { item.description = value; return api; },
    setLeft: (value: number) => { item.left = value; return api; },
    setTop: (value: number) => { item.top = value; return api; },
    setWidth: (value: number) => { item.width = value; return api; },
    setHeight: (value: number) => { item.height = value; return api; },
    setRotation: (value: number) => { item.rotation = value; return api; },
    select: () => { item.selected = true; },
    remove: () => { item.removed = true; state.elements.delete(item.id); },
    getParentGroup: () => null,
    asTable: () => api,
    asShape: () => api,
    asGroup: () => api,
    getChildren: () => [],
  };
  if (item.type === "SHAPE") {
    // Shapes default to a single 12pt run so font scaling has something to move.
    item.runs ??= [{ text: "", fontSize: 12 }];
    const syncSizes = (): void => { state.fontSizes.set(item.id, item.runs!.map((run) => run.fontSize)); };
    syncSizes();
    const textStyle = (run: { fontSize: number }) => ({
      getFontSize: () => run.fontSize,
      setFontSize: (value: number) => { run.fontSize = value; syncSizes(); },
      setForegroundColor: () => undefined,
      setFontFamily: () => undefined,
      setBold: () => undefined,
      setItalic: () => undefined,
      getForegroundColor: () => null,
      getFontFamily: () => null,
      isBold: () => null,
      isItalic: () => null,
    });
    Object.assign(api, {
      getText: () => ({
        asString: () => item.runs!.map((run) => run.text).join(""),
        setText: (value: string) => { item.runs = [{ text: value, fontSize: item.runs![0]?.fontSize ?? 12 }]; syncSizes(); },
        getRuns: () => item.runs!.map((run) => ({ getTextStyle: () => textStyle(run) })),
        getTextStyle: () => textStyle(item.runs![0]!),
        getParagraphStyle: () => ({ getParagraphAlignment: () => null, setParagraphAlignment: () => undefined }),
      }),
      getFill: () => ({ getType: () => "SOLID", getSolidFill: () => ({ getColor: () => ({ getColorType: () => "RGB", asRgbColor: () => ({ asHexString: () => "#FFFFFF" }) }) }), setSolidFill: () => undefined }),
      getBorder: () => ({
        getLineFill: () => ({ getFillType: () => "SOLID", getSolidFill: () => ({ getColor: () => ({ getColorType: () => "RGB", asRgbColor: () => ({ asHexString: () => "#000000" }) }) }), setSolidFill: () => undefined }),
        getWeight: () => 1, setWeight: () => undefined, getDashStyle: () => null, setDashStyle: () => undefined,
      }),
      getShapeType: () => "RECTANGLE",
      getContentAlignment: () => null,
      setContentAlignment: () => undefined,
    });
  }
  if (item.type === "TABLE") {
    // Tables have to answer the same way whether the code got them from
    // insertTable or from getPageElements, or the chart reader sees an empty grid.
    item.cells ??= [];
    Object.assign(api, {
      getNumRows: () => item.cells!.length,
      getNumColumns: () => Math.max(0, ...item.cells!.map((row) => row.length)),
      getColumn: (index: number) => ({
        getWidth: () => item.columnWidths?.[index] ?? item.width / Math.max(1, Math.max(0, ...item.cells!.map((row) => row.length))),
      }),
      getRow: (index: number) => ({
        getMinimumHeight: () => item.rowHeights?.[index] ?? item.height / Math.max(1, item.cells!.length),
      }),
      getCell: (row: number, column: number) => ({
        getText: () => ({
          asString: () => item.cells![row]?.[column] ?? "",
          setText: (value: string) => {
            while (item.cells!.length <= row) item.cells!.push([]);
            item.cells![row]![column] = value;
          },
        }),
      }),
    });
  }
  return api;
}

export interface HostOptions {
  /** Elements already on the slide when the code runs. */
  existing?: (Partial<FakeElement> & { id: string; type: string })[];
  /** Which of those the user has selected. */
  selectedIds?: string[];
  pageWidth?: number;
  pageHeight?: number;
  /** Icon path data, keyed by project file name, e.g. IconPaths00. */
  projectFiles?: Record<string, string>;
}

/**
 * Installs the fake host globals and returns the recorded state. Call
 * `restoreHost()` afterwards so tests stay independent.
 */
export function installHost(options: HostOptions = {}): HostState {
  const state: HostState = {
    requests: [],
    elements: new Map(),
    documentProperties: new Map(),
    userProperties: new Map(),
    projectFiles: new Map(Object.entries(options.projectFiles ?? {})),
    batchUpdates: 0,
    fontSizes: new Map(),
  };
  uuidCounter = 0;
  // The icon catalog memoises shards for the life of an execution; each test is
  // a fresh "execution", so the cache must not survive between them.
  resetCatalogCache();

  for (const item of options.existing ?? []) element(state, item);
  const selectedIds = options.selectedIds ?? (options.existing ?? []).map((item) => item.id);
  const pageWidth = options.pageWidth ?? 720;
  const pageHeight = options.pageHeight ?? 405;

  const slide: Record<string, unknown> = {
    getObjectId: () => "slide1",
    getPageType: () => "SLIDE",
    asSlide: () => slide,
    getPageElements: () => [...state.elements.values()].map((item) => wrapElement(state, item)),
    insertImage: (blob: { getBytes: () => number[]; getContentType: () => string }, left: number, top: number, width: number, height: number) => {
      const created = element(state, {
        id: `image${state.elements.size + 1}`, type: "IMAGE", left, top, width, height,
        image: { bytes: blob.getBytes(), contentType: blob.getContentType() },
      });
      return wrapElement(state, created);
    },
    insertTable: (rows: number, columns: number, left: number, top: number, width: number, height: number) => {
      const created = element(state, {
        id: `table${state.elements.size + 1}`, type: "TABLE", left, top, width, height,
        cells: Array.from({ length: rows }, () => Array.from({ length: columns }, () => "")),
      });
      return wrapElement(state, created);
    },
  };

  const presentation = {
    getId: () => "presentation1",
    getName: () => "Test deck",
    getPageWidth: () => pageWidth,
    getPageHeight: () => pageHeight,
    getSlides: () => [slide],
    getPageElementById: (id: string) => {
      const item = state.elements.get(id);
      return item ? wrapElement(state, item) : null;
    },
    getSelection: () => ({
      getCurrentPage: () => slide,
      getPageElementRange: () => ({
        getPageElements: () => selectedIds
          .map((id) => state.elements.get(id))
          .filter((item): item is FakeElement => Boolean(item))
          .map((item) => wrapElement(state, item)),
      }),
      getTextRange: () => null,
    }),
  };

  (globalThis as Record<string, unknown>).SlidesApp = {
    getActivePresentation: () => presentation,
    openById: () => presentation,
    PageType: { SLIDE: "SLIDE" },
    PageElementType: { SHAPE: "SHAPE", IMAGE: "IMAGE", TABLE: "TABLE", GROUP: "GROUP", LINE: "LINE" },
    ShapeType: { ELLIPSE: "ELLIPSE", RECTANGLE: "RECTANGLE" },
    LineCategory: { STRAIGHT: "STRAIGHT" },
  };

  (globalThis as Record<string, unknown>).Slides = {
    Presentations: {
      get: () => ({ revisionId: "revision1" }),
      batchUpdate: (body: { requests: RecordedRequest[] }) => {
        state.batchUpdates += 1;
        state.requests.push(...body.requests);
        // Model grouping well enough that callers can look the result up.
        for (const request of body.requests) {
          const group = request.groupObjects as { groupObjectId?: string } | undefined;
          if (group?.groupObjectId) element(state, { id: group.groupObjectId, type: "GROUP" });
        }
        return {};
      },
    },
  };

  (globalThis as Record<string, unknown>).Utilities = {
    getUuid: () => `uuid-${++uuidCounter}`,
    newBlob: (bytes: number[], contentType: string, name: string) => ({
      getBytes: () => bytes,
      getContentType: () => contentType,
      getName: () => name,
    }),
    base64Decode: (value: string) => Array.from(value).map((character) => character.charCodeAt(0)),
  };

  const properties = (store: Map<string, string>) => ({
    getProperty: (key: string) => store.get(key) ?? null,
    setProperty: (key: string, value: string) => { store.set(key, value); },
    setProperties: (values: Record<string, string>) => { for (const key of Object.keys(values)) store.set(key, values[key]!); },
    deleteProperty: (key: string) => { store.delete(key); },
  });
  (globalThis as Record<string, unknown>).PropertiesService = {
    getDocumentProperties: () => properties(state.documentProperties),
    getUserProperties: () => properties(state.userProperties),
  };
  (globalThis as Record<string, unknown>).LockService = { getDocumentLock: () => null };
  (globalThis as Record<string, unknown>).HtmlService = {
    createHtmlOutputFromFile: (name: string) => {
      const content = state.projectFiles.get(name);
      if (content == null) throw new Error(`missing project file ${name}`);
      return { getContent: () => content };
    },
  };
  (globalThis as Record<string, unknown>).console = console;

  return state;
}

export function restoreHost(): void {
  for (const name of ["SlidesApp", "Slides", "Utilities", "PropertiesService", "LockService", "HtmlService"]) {
    delete (globalThis as Record<string, unknown>)[name];
  }
}

/** All requests of one kind, e.g. every createShape. */
export function requestsOfKind(state: HostState, kind: string): Record<string, unknown>[] {
  return state.requests.filter((request) => kind in request).map((request) => request[kind] as Record<string, unknown>);
}

/** Text of every createShape+insertText pair, in order. */
export function insertedText(state: HostState): string[] {
  return requestsOfKind(state, "insertText").map((request) => String(request.text));
}
