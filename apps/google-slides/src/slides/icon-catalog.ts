/**
 * Server-side access to the ~10k icon path data.
 *
 * The paths ship as numbered HTML files in the Apps Script project (Apps Script
 * only stores .gs and .html, so a JSON payload rides inside an .html file).
 * Icons are sorted by id and sharded into contiguous ranges, so an id resolves to
 * its shard with a binary search over the small boundary table - no id-to-shard
 * map has to be shipped or kept in sync, and Make Editable keeps working on icons
 * inserted before a catalog rebuild.
 */

interface ShardBoundary {
  shard: number;
  firstId: string;
  lastId: string;
}

const SHARD_FILE_PREFIX = "IconPaths";
const BOUNDARY_FILE = "IconShards";

// Parsed shards are memoised for the life of one execution: converting several
// selected icons at once then costs one parse per shard rather than per icon.
// Apps Script gives every execution a fresh module, so this never goes stale in
// production; resetCatalogCache exists so tests can swap deployments.
let loaded: Record<number, Record<string, string[]>> = {};
let boundaryTable: ShardBoundary[] | null = null;

export function resetCatalogCache(): void {
  loaded = {};
  boundaryTable = null;
}

function readProjectJson<T>(name: string): T {
  const content = HtmlService.createHtmlOutputFromFile(name).getContent();
  return JSON.parse(content) as T;
}

function boundaries(): ShardBoundary[] {
  if (!boundaryTable) {
    try {
      boundaryTable = readProjectJson<ShardBoundary[]>(BOUNDARY_FILE);
    } catch {
      throw new Error("The icon catalog is missing from this deployment. Run npm run build and push again.");
    }
  }
  return boundaryTable;
}

function shardFor(id: string): number | null {
  const table = boundaries();
  let low = 0;
  let high = table.length - 1;
  while (low <= high) {
    const middle = (low + high) >> 1;
    const entry = table[middle]!;
    if (id < entry.firstId) high = middle - 1;
    else if (id > entry.lastId) low = middle + 1;
    else return entry.shard;
  }
  return null;
}

function loadShard(shard: number): Record<string, string[]> {
  const cached = loaded[shard];
  if (cached) return cached;
  const name = `${SHARD_FILE_PREFIX}${String(shard).padStart(2, "0")}`;
  let data: Record<string, string[]>;
  try {
    data = readProjectJson<Record<string, string[]>>(name);
  } catch {
    throw new Error(`Icon data file ${name} is missing from this deployment.`);
  }
  loaded[shard] = data;
  return data;
}

/** Subpath strings for one icon, or null when the catalog does not have it. */
export function iconPaths(id: string): string[] | null {
  const shard = shardFor(id);
  if (shard == null) return null;
  return loadShard(shard)[id] ?? null;
}

/**
 * Path data for a batch of icons, for the sidebar's preview grid. Ids are grouped
 * by shard first so a screenful of neighbouring icons costs a single read.
 */
export function iconPathsFor(ids: readonly string[]): Record<string, string[]> {
  if (!Array.isArray(ids)) throw new Error("Ask for icons by id.");
  if (ids.length > 500) throw new Error("Too many icons requested at once.");
  const wanted: Record<number, string[]> = {};
  for (const raw of ids) {
    const id = String(raw);
    if (!/^[a-z0-9-]{1,64}$/.test(id)) continue;
    const shard = shardFor(id);
    if (shard == null) continue;
    (wanted[shard] ??= []).push(id);
  }
  const result: Record<string, string[]> = {};
  for (const key of Object.keys(wanted)) {
    const shard = loadShard(Number(key));
    for (const id of wanted[Number(key)]!) {
      const paths = shard[id];
      if (paths) result[id] = paths;
    }
  }
  return result;
}
