/**
 * OneDrive v2 source: platform-owned Microsoft OAuth connect (root folders
 * chosen in the platform's shared folder-picker), a per-root walk over
 * Microsoft Graph's `/drive/items/{id}/delta` feed for both the initial
 * backfill AND every later poll, and a pure toDocument.
 *
 * Ported from the v1 connector (`git show
 * main:src/main/connectors/onedrive/<file>.ts`): backfill.ts's per-root
 * enumerate-to-deltaLink loop, delta.ts's per-root poll + 410-resync
 * recovery, ingest.ts's routing / hash-skip / downloadUrl-refresh-and-retry,
 * path-resolver.ts's pure path building (ported verbatim in path-
 * resolver.ts), tree.ts's picker listings (ported in tree.ts). Platform
 * pieces (SQLite tables, tracked-roots UI, local converter pipeline,
 * safeStorage token blobs, scheduler) are replaced by the v2 engine.
 *
 * KEY STRUCTURAL DIFFERENCE FROM GOOGLE DRIVE: Microsoft Graph's per-item
 * `/delta` endpoint IS both the initial full-recursive-enumeration AND the
 * later incremental poll for the SAME root — calling it with no token
 * returns every descendant across pages ending in `@odata.deltaLink`;
 * calling it again with that link's `token` returns only what changed since.
 * There is no separate folder-BFS listing call (unlike Drive's files.list),
 * so this port has no ancestor/folder-index walk at all — Graph's
 * `parentReference.path` gives the full path for free (path-resolver.ts).
 *
 * Deliberate v2 changes from v1 (see README):
 *  - Extraction is DEFERRED to the v2 engine: v1's ingest.ts downloaded bytes
 *    and ran a local converter itself, storing markdown directly. v2 instead
 *    emits `binary: { bytes, mime, filename }` with `markdown: null` for
 *    every downloadable file — the engine's convert/vision pipeline extracts
 *    it (mirrors the google-docs-kia-connector template's binary route). See
 *    mime-route.ts for the routing gate this replaces v1's local
 *    `isConvertibleMime` with — broadened to include images (v1 left images
 *    unsupported; the v2 vision worker OCRs them via `fetchBytes`).
 *  - Content eligibility is now decided by kiagent-core's canonical
 *    `decideFileIndexing` policy (via the SDK's `chooseRoute` — see
 *    mime-route.ts), consulted in `pageChunks` BEFORE any downloadUrl
 *    refresh or download. Archives (any size), cloud audio/video (any
 *    size), unsupported types, and oversized files (converter/PDF over
 *    `MAX_CLOUD_BINARY_BYTES` = 25 MiB, images over `MAX_CLOUD_IMAGE_BYTES` =
 *    20 MiB) are ignored outright — no metadata-only row is produced for
 *    them, and an already-indexed file the policy now excludes is archived
 *    (see the ignore branch of `pageChunks`, and `buildItem`'s post-download
 *    re-check for the declared-size-absent case — this connector has no
 *    reconcile pass, so these two sites are the only places that will ever
 *    clean up such a row). v1 had no size cap at all (a
 *    gap the google-docs-kia-connector template called out and fixed for
 *    Drive first). A file whose declared `size` is absent is admitted
 *    provisionally and re-checked against the same policy once its real
 *    byte length is known post-download.
 *  - Tracked-root OVERLAP VALIDATION (v1's overlap.ts) is DROPPED: v1 needed
 *    it because its own folder-picker UI let a user select a folder inside
 *    an already-tracked one. The v2 contract's shared `pickFolders` already
 *    guarantees "covering roots — never both a node and its own descendant"
 *    (see `AuthChannel.pickFolders` in @kiagent/connector-sdk), so there is
 *    no place left in this connector for overlap.ts's check to run. As
 *    defense in depth against a hand-edited config with overlapping roots
 *    anyway (or a future picker relaxation), `processed` (a per-pull-call
 *    `Set<string>` of item ids already emitted) still guarantees an item
 *    reachable from two configured roots is only ever ingested once per
 *    pull — the same rationale as the google-docs-kia-connector template's
 *    `walked` folder set.
 *  - No RECONCILE: v1 never had a periodic full-listing reconciler for
 *    OneDrive (unlike Drive, whose `changes.list` token can go stale in a
 *    way serious enough that v1 shipped `runFullRescan` — later found buggy
 *    and REMOVED in the google-docs v2 port). Graph's delta feed reports
 *    deletions directly and a resync (410 `resyncRequired`) triggers an
 *    inline full re-establish of just that one root (see `isResyncRequired`
 *    below) — so nothing here is missing a reconciliation path v1 had.
 *  - A root that never gets a delta token (interrupted mid-establish, or a
 *    resync-reprime that gets interrupted) SELF-HEALS on the next pull(): v1
 *    gap — its delta.ts logged a warning and permanently skipped a tokenless
 *    root ("must run backfill first") with no automatic recovery. This port's
 *    `backfillDone` check re-enters `backfill()` for any root still lacking
 *    a token, on every pull cycle, until it succeeds.
 *  - Per-batch cursor commits: v1 only persisted `cursor_json` after an
 *    entire `runDelta` tick across ALL roots finished (a crash mid-root lost
 *    that WHOLE tick's progress for every root). v2's `Batch.cursor` commits
 *    transactionally with each page's items, so a crash only replays the
 *    page in flight (idempotent via hash-skip).
 */
import type {
  Account,
  AuthChannel,
  Batch,
  Credentials,
  Document,
  DocumentInput,
  ExternalRef,
  FileIgnoreReason,
  FolderScopeUpdate,
  FolderScopedConfig,
  FolderSelectionChannel,
  HostFor,
  Query,
  Session,
  Source,
} from '@kiagent/connector-sdk';
import { MAX_CLOUD_BINARY_BYTES, SourcePermanentError } from '@kiagent/connector-sdk';
import {
  GRAPH_BASE,
  GraphApiError,
  GraphClient,
  OneDriveAuthError,
  isAuthError,
  type GraphClientDeps,
} from './client';
import { buildDisplayPath } from './path-resolver';
import { chooseRoute } from './mime-route';
import { countChildren, listChildFolders, listSharedRoots } from './tree';

export { GRAPH_BASE };

/** v1 FILES_SCOPES minus `openid`/`email`/`profile`/`offline_access` — the
 *  platform's Microsoft OAuth provider appends those itself (see brief: the
 *  extension only requests Graph RESOURCE scopes). */
export const FILES_SCOPES = ['Files.Read.All', 'User.Read'];

/** v1 had NO cap (whole file into a Buffer) — v2 binds one (see module doc).
 *  Re-exported from the SDK's canonical `MAX_CLOUD_BINARY_BYTES` under this
 *  connector's original name so existing callers/tests are undisturbed; the
 *  actual cap enforcement lives entirely in `decideFileIndexing` now (via
 *  `chooseRoute`), not in a local comparison against this constant. */
export const MAX_BINARY_BYTES = MAX_CLOUD_BINARY_BYTES;

/**
 * Per-batch flush budget. A Graph delta page is server-sized (typically a
 * few hundred items) and every convertible file in it — images included, for
 * vision OCR — carries up to MAX_BINARY_BYTES of downloaded bytes. Holding a
 * whole page's bytes at once (then structured-cloning them over IPC in ONE
 * message) put a photo-heavy personal OneDrive past the extension process's
 * heap, and because the cursor only advances per page the same page replayed
 * on every retry: a deterministic crash loop ("extension process exited").
 * So a page is flushed to the engine in sub-page chunks once the accumulated
 * bytes or entry count cross these budgets — see `pageChunks`. Both are soft
 * ceilings checked AFTER an entry is added, so a chunk can overshoot by one
 * file (≤ MAX_BINARY_BYTES).
 */
export const BATCH_BYTE_BUDGET = 32 * 1024 * 1024;
export const BATCH_ITEM_LIMIT = 250;

export interface BatchBudget {
  bytes: number;
  items: number;
}

export interface RootConfig {
  rootFolderId: string;
  rootName: string;
}

export interface OneDriveCursor {
  delta_tokens: Record<string, string>;
  /** The configured root ids this cursor's coverage was built for. ABSENT on
   *  every cursor written before folder scope existed — core's v3 migration
   *  deliberately leaves `Account.cursor` untouched — which
   *  `normalizeCursor` treats as a mismatch. Deliberately NOT derived from
   *  `delta_tokens` keys: nothing in this connector ever deletes a key
   *  (`backfill()` and `delta()` both seed from a spread of the incoming
   *  map), so those keys are a SUPERSET of the configured roots and would
   *  claim coverage of roots that are long gone (spec-reality-diff A5a). */
  scope_roots?: string[];
  /** Present only while establishing a root (initial backfill, or a root
   *  added after the account went live) for the first time. */
  backfill?: { root_index: number; next_link?: string };
}

/** Subset of Microsoft Graph's driveItem consumed here. */
export interface DriveItem {
  id: string;
  name: string;
  webUrl?: string;
  size?: number;
  eTag?: string;
  cTag?: string;
  createdDateTime?: string;
  lastModifiedDateTime?: string;
  file?: { mimeType?: string };
  folder?: { childCount?: number };
  deleted?: { state: string };
  parentReference?: { driveId?: string; path?: string; parentId?: string };
  '@microsoft.graph.downloadUrl'?: string;
}

export interface OneDriveItem {
  file: DriveItem;
  /** '' for the one remaining metadata-only case (downloadUrl unobtainable,
   *  `extractionStatus: 'failed'`) — a file the eligibility policy has
   *  already ruled IN never reaches this type at all when the policy rules
   *  it OUT (see `pageChunks`'s ignore branch, which never calls
   *  `buildItem`). null for binary items (the engine converts). */
  markdown: string | null;
  bytes?: Uint8Array;
  extractionStatus: 'ok' | 'failed';
  displayPath: string;
  rootFolderId: string;
}

interface GraphDeltaPage {
  value?: DriveItem[];
  '@odata.nextLink'?: string;
  '@odata.deltaLink'?: string;
}

const errText = (e: unknown): string => (e instanceof Error ? e.message : String(e));

/** Graph's alias for the caller's own drive root — the "whole OneDrive"
 *  catch-all. This ONE literal is the id the My-files picker root carries in
 *  both `connect()` and `manageFolders`, the id `rootsConfig` falls back to,
 *  and the id `resolveRootLocation` short-circuits. It is also the id core's
 *  v3 migration must treat as OneDrive's catch-all (DECISIONS R6): a live row
 *  whose `metadata.root_folder_id` matches no other selected root is
 *  attributed to THIS id instead of being archived. Core cannot import it
 *  (different package), so a test pins the literal on this side. */
export const ONEDRIVE_DRIVE_ROOT_ID = 'root';

/** Normalize account config to the tracked roots.
 *
 *  Reads the CANONICAL `folderRoots: [{id, name}]` written by `connect()`,
 *  `manageFolders` and core's v3 migration. The legacy
 *  `roots: [{rootFolderId, rootName}]` shape is a read-only FALLBACK, used
 *  only when `folderRoots` is absent or yields nothing usable: the R1 mirror
 *  runs the other way (core writes `roots` alongside `folderRoots` for one
 *  release train so the *installed* 2.0.5 build keeps working), so canonical
 *  always wins and the two shapes are never merged.
 *
 *  Deduped by id — first entry wins. Falls back to the whole OneDrive when
 *  neither key holds anything usable (defensive only — `connect()` and
 *  `manageFolders` both reject an empty selection). Order is SEMANTIC: it
 *  drives `backfill()`'s `root_index` and the first-root-wins attribution. */
export function rootsConfig(session: Session): RootConfig[] {
  const cfg = session.account.config as { folderRoots?: unknown; roots?: unknown };
  const parsed: RootConfig[] = [];
  if (Array.isArray(cfg.folderRoots)) {
    for (const raw of cfg.folderRoots) {
      const r = raw as { id?: unknown; name?: unknown } | null;
      if (r && typeof r.id === 'string' && r.id) {
        const name = typeof r.name === 'string' && r.name ? r.name : r.id;
        parsed.push({ rootFolderId: r.id, rootName: name });
      }
    }
  }
  // TODO(folder-scope-train-2): drop the legacy `roots` fallback (R1).
  if (parsed.length === 0 && Array.isArray(cfg.roots)) {
    for (const raw of cfg.roots) {
      const r = raw as { rootFolderId?: unknown; rootName?: unknown } | null;
      if (r && typeof r.rootFolderId === 'string' && r.rootFolderId) {
        const name =
          typeof r.rootName === 'string' && r.rootName ? r.rootName : r.rootFolderId;
        parsed.push({ rootFolderId: r.rootFolderId, rootName: name });
      }
    }
  }
  if (parsed.length === 0) {
    parsed.push({ rootFolderId: ONEDRIVE_DRIVE_ROOT_ID, rootName: 'OneDrive' });
  }
  const seen = new Set<string>();
  const deduped: RootConfig[] = [];
  for (const r of parsed) {
    if (seen.has(r.rootFolderId)) continue;
    seen.add(r.rootFolderId);
    deduped.push(r);
  }
  return deduped;
}

async function requireToken(session: Session): Promise<string> {
  const creds = await session.credentials();
  if (!creds?.accessToken) {
    throw new OneDriveAuthError('onedrive: no credentials available — reconnect the account');
  }
  return creds.accessToken;
}

/** Aggregate ignore-reason counter for ONE `pull()` call — never holds a
 *  filename or path, only counts by `FileIgnoreReason`. Shared (via
 *  `ItemDeps.ignored`) across every root and page a single `backfill()` or
 *  `delta()` walks, so `logIfAny` fires at most once per completed pull. */
class IgnoreTally {
  private readonly counts = new Map<FileIgnoreReason, number>();

  add(reason: FileIgnoreReason): void {
    this.counts.set(reason, (this.counts.get(reason) ?? 0) + 1);
  }

  get total(): number {
    let sum = 0;
    for (const n of this.counts.values()) sum += n;
    return sum;
  }

  private summary(): string {
    return [...this.counts.entries()]
      .sort((a, b) => b[1] - a[1])
      .map(([reason, n]) => `${reason}=${n}`)
      .join(' ');
  }

  /** No-op when nothing was ignored — a pull with full eligibility logs
   *  nothing extra. */
  logIfAny(session: Session): void {
    if (this.total === 0) return;
    session.log('info', `onedrive: ignored ${this.total} file(s) this pull (${this.summary()})`);
  }
}

interface ItemDeps {
  client: GraphClient;
  session: Session;
  query: Query;
  fetchFn: (url: string, init?: unknown) => Promise<unknown>;
  /** Item ids already emitted THIS pull() call — guards overlapping
   *  configured roots (see module doc). */
  processed: Set<string>;
  budget: BatchBudget;
  /** Eligibility-ignore counts for this pull() call (see `pageChunks`'s
   *  ignore branch and `buildItem`'s post-download re-check). */
  ignored: IgnoreTally;
}

/** Query-first content-hash skip: an unchanged, still-live document is never
 *  re-downloaded (v1 ingest.ts's `canShortcut`, minus the metadata-only
 *  `last_seen_at` refresh — the v2 engine owns row freshness). An ARCHIVED
 *  match does NOT skip: re-emitting is what un-archives a doc that moved back
 *  into scope. Only an `'ok'` row (real, already-fetched content) is ever
 *  pinned behind an unchanged eTag. Anything else bypasses the pin: a
 *  'failed' row (v1's exact rationale — possibly just a quota storm; must
 *  retry) AND, just as importantly, a row carrying one of the PRE-migration
 *  statuses this connector used to write for an ineligible file
 *  (`'unsupported'`/`'too-large'`) — `pageChunks` only ever calls `hashSkip`
 *  for an item the CURRENT policy has judged eligible, so reaching here at
 *  all means a legacy row that used to be ineligible was just newly rescued
 *  by a policy change (e.g. an octet-stream-mime `.pdf` rescued by its
 *  extension). Pinning it on an unchanged eTag would strand it as
 *  contentless metadata forever, since nothing else in this connector will
 *  ever revisit it. */
async function hashSkip(deps: ItemDeps, itemId: string, etag: string | undefined): Promise<boolean> {
  if (!etag) return false;
  const existing = await deps.query.byExternalId(deps.session.account.id, itemId, 'file');
  if (!existing || existing.archivedAt) return false;
  const meta = existing.metadata as Record<string, unknown>;
  if (meta.extraction_status !== 'ok') return false;
  return meta.etag === etag;
}

/** The one remaining metadata-only case: an eligible file whose downloadUrl
 *  could not be obtained even after a refresh (v1 parity — see `buildItem`).
 *  Ineligible files never reach here at all (`pageChunks` gates them out
 *  before `buildItem` is ever called). */
function failedItem(raw: DriveItem, displayPath: string, rootFolderId: string): OneDriveItem {
  return { file: raw, markdown: '', extractionStatus: 'failed', displayPath, rootFolderId };
}

/** Do NOT pass `$select` here (v1 `refreshItemDownloadUrl` parity):
 *  `@microsoft.graph.downloadUrl` is an OData annotation, and Graph's
 *  `$select` clause strips annotations even when listed explicitly — a bare
 *  GET returns the full driveItem with the annotation populated. */
async function refreshDownloadUrl(client: GraphClient, itemId: string): Promise<string | undefined> {
  const item = await client.request<{ '@microsoft.graph.downloadUrl'?: string }>(
    `${GRAPH_BASE}/me/drive/items/${itemId}`,
  );
  return item['@microsoft.graph.downloadUrl'];
}

/** One-shot pre-signed-URL download (v1 `ingest.ts`'s `downloadBytes`
 *  parity): no bearer header (the URL is self-authenticating), no retry loop
 *  — the caller (`buildItem`) handles exactly one 403-refresh-and-retry, v1's
 *  own contract. */
async function downloadBytes(
  fetchFn: (url: string, init?: unknown) => Promise<unknown>,
  url: string,
): Promise<Uint8Array> {
  const res = (await fetchFn(url)) as { status: number; body: Uint8Array };
  if (res.status >= 200 && res.status < 300) return res.body;
  const body = new TextDecoder().decode(res.body).slice(0, 500);
  throw new GraphApiError(res.status, url, body);
}

/**
 * Route one already-eligible (`chooseRoute` in `pageChunks` already ruled it
 * IN) listed/changed item into an OneDriveItem, or null to skip it — via
 * hash-skip, or because a declared-size-absent file turns out oversized once
 * its real byte length is known. May do I/O (downloadUrl refresh / bytes
 * download) — toDocument stays pure. Never called for a route `chooseRoute`
 * ruled OUT (see `pageChunks`'s ignore branch).
 *
 * `deletions` is the caller's (`pageChunks`'s) in-progress chunk array,
 * passed through so the post-download re-check below can push into it
 * directly — the exact same "was this id already a live row?" pattern
 * `pageChunks`'s own pre-check ignore branch uses. This connector has no
 * periodic reconcile pass (see module doc), so this is the ONLY chance a
 * post-download-oversize file with a pre-existing live row ever gets to
 * archive that stale row; skipping it here would strand it forever.
 */
async function buildItem(
  raw: DriveItem,
  root: RootConfig,
  deps: ItemDeps,
  deletions: ExternalRef[],
): Promise<OneDriveItem | null> {
  const mime = raw.file?.mimeType ?? 'application/octet-stream';
  const displayPath = buildDisplayPath({ rootName: root.rootName }, raw);

  if (await hashSkip(deps, raw.id, raw.eTag)) return null;

  // v1 ingest.ts: Graph's per-item /delta response routinely omits
  // @microsoft.graph.downloadUrl (especially on MSA personal OneDrive) —
  // fall back to a fresh GET before giving up.
  let downloadUrl = raw['@microsoft.graph.downloadUrl'];
  if (!downloadUrl) {
    downloadUrl = await refreshDownloadUrl(deps.client, raw.id);
  }
  if (!downloadUrl) {
    return failedItem(raw, displayPath, root.rootFolderId);
  }

  let bytes: Uint8Array;
  try {
    bytes = await downloadBytes(deps.fetchFn, downloadUrl);
  } catch (e) {
    if (e instanceof GraphApiError && e.status === 403) {
      const refreshed = await refreshDownloadUrl(deps.client, raw.id);
      if (!refreshed) throw e;
      bytes = await downloadBytes(deps.fetchFn, refreshed);
    } else {
      throw e;
    }
  }
  // Post-download re-check: `pageChunks`'s pre-check already applied the
  // policy's declared-size cap when Graph's `size` field was present. This
  // only matters when `size` was absent (admitted provisionally) and the
  // real byte length now exceeds the pipeline's cap — re-running the SAME
  // policy (rather than a flat local constant) keeps the check correct per
  // pipeline (25 MiB converter/PDF vs 20 MiB images) without duplicating the
  // SDK's caps here.
  const postRoute = chooseRoute(mime, raw.name, bytes.byteLength);
  if (postRoute.kind === 'ignore') {
    deps.ignored.add(postRoute.reason);
    const existing = await deps.query.byExternalId(deps.session.account.id, raw.id, 'file');
    if (existing && !existing.archivedAt) {
      deletions.push({ externalId: raw.id, type: 'file' });
    }
    return null;
  }

  return {
    file: raw,
    markdown: null,
    bytes,
    extractionStatus: 'ok',
    displayPath,
    rootFolderId: root.rootFolderId,
  };
}

interface PageChunk {
  items: OneDriveItem[];
  deletions: ExternalRef[];
  /** True on the LAST chunk of the page — the only one after which the
   *  caller may advance its cursor past this page. Intermediate chunks must
   *  be committed under the cursor that fetched the page, so a crash after
   *  them replays the page (idempotent via hash-skip) rather than skipping
   *  its unflushed remainder. */
  pageComplete: boolean;
}

/**
 * Build the items/deletions for one Graph delta page, yielding them in
 * sub-page chunks as soon as the accumulated downloaded bytes or entry count
 * cross `deps.budget` (see BATCH_BYTE_BUDGET). Always ends with exactly one
 * `pageComplete` chunk (possibly empty) unless the session is aborted.
 * Folders (no `.file` facet) and the root item itself are never ingested
 * (v1 parity). Every live `.file` item is classified via `chooseRoute`
 * BEFORE `hashSkip`/`buildItem` even run: an ineligible file never reaches a
 * downloadUrl refresh or a download, and — if it was previously indexed —
 * is surfaced as a deletion instead. Deletions (upstream tombstones and
 * policy transitions alike) are cheap but count against the entry limit so
 * a mass-delete page is still bounded.
 */
async function* pageChunks(
  page: GraphDeltaPage,
  root: RootConfig,
  deps: ItemDeps,
  opts: { includeDeletions: boolean },
): AsyncGenerator<PageChunk> {
  let items: OneDriveItem[] = [];
  let deletions: ExternalRef[] = [];
  let bytes = 0;
  const full = (): boolean => bytes >= deps.budget.bytes || items.length + deletions.length >= deps.budget.items;
  const flush = (pageComplete: boolean): PageChunk => {
    const chunk = { items, deletions, pageComplete };
    items = [];
    deletions = [];
    bytes = 0;
    return chunk;
  };

  for (const raw of page.value ?? []) {
    if (deps.session.signal.aborted) return;
    if (raw.deleted) {
      // Same overlapping-roots guard as live items: a deletion reported by two
      // feeds is surfaced once per pull.
      if (!opts.includeDeletions || deps.processed.has(raw.id)) continue;
      deps.processed.add(raw.id);
      const existing = await deps.query.byExternalId(deps.session.account.id, raw.id, 'file');
      if (existing) deletions.push({ externalId: raw.id, type: 'file' });
    } else if (!raw.file || raw.id === root.rootFolderId || deps.processed.has(raw.id)) {
      // folder (not ingested, v1 parity) / the root item itself / overlapping
      // tracked roots — first root wins.
      continue;
    } else {
      deps.processed.add(raw.id);
      const mime = raw.file.mimeType ?? 'application/octet-stream';
      const sizeBytes = typeof raw.size === 'number' ? raw.size : undefined;
      const route = chooseRoute(mime, raw.name, sizeBytes);
      if (route.kind === 'ignore') {
        // A local eligibility transition — NOT gated by `opts.includeDeletions`
        // (that flag scopes only upstream `raw.deleted` tombstones): an
        // already-indexed file the policy now excludes must still be
        // archived out, including one encountered during backfill. No
        // downloadUrl refresh or download ever happens for it.
        const existing = await deps.query.byExternalId(deps.session.account.id, raw.id, 'file');
        if (existing && !existing.archivedAt) {
          deletions.push({ externalId: raw.id, type: 'file' });
        }
        deps.ignored.add(route.reason);
      } else {
        try {
          const item = await buildItem(raw, root, deps, deletions);
          if (item) {
            items.push(item);
            bytes += item.bytes?.byteLength ?? 0;
          }
        } catch (e) {
          if (isAuthError(e)) throw e;
          // One unreadable item must not abort the walk (v1 parity: backfill.ts
          // / delta.ts both log-and-continue on a per-item error).
          deps.session.log('warn', `onedrive: item ${raw.id} skipped: ${errText(e)}`);
        }
      }
    }
    if (full()) yield flush(false);
  }
  if (deps.session.signal.aborted) return;
  yield flush(true);
}

const tokenFromDeltaLink = (link: string): string => new URL(link).searchParams.get('token') ?? '';

/** v1 delta.ts's resync check, reproduced against this port's `graph <status>
 *  <url> <body>` message contract (client.ts). Graph signals an expired/
 *  invalid delta token as HTTP 410 with an error code of `resyncRequired`
 *  (or, per Graph's docs, `invalidToken`/`syncStateNotFound`). */
function isResyncRequired(msg: string): boolean {
  return /410/.test(msg) && /(resyncRequired|invalidToken|syncStateNotFound)/i.test(msg);
}

interface WalkStep {
  items: OneDriveItem[];
  deletions: ExternalRef[];
  /** False for an intermediate chunk of a page (budget flush) — commit under
   *  the cursor that fetched the page; neither `token` nor `nextLink` is set. */
  pageComplete: boolean;
  /** Set once the walk reaches `@odata.deltaLink` — the new live token. */
  token?: string;
  /** Set on the final chunk of every non-final page — `@odata.nextLink` to
   *  resume from. */
  nextLink?: string;
}

/**
 * Walks ONE root's Graph `/delta` feed from `startUrl` to completion (paging
 * via `@odata.nextLink` until a page carries `@odata.deltaLink`), yielding
 * one step per page. The SAME mechanism serves both a root's initial
 * establishment (backfill) and a resync reprime (delta) — v1 shared the
 * identical `walkGraphDelta` call for both (`backfill.ts`'s `enumerateRoot` /
 * `delta.ts`'s `walkRootDelta` without a token). A page over the batch
 * budget yields several steps; only its last one (`pageComplete`) carries
 * the link/token to advance by.
 */
async function* walkRootToDeltaLink(
  client: GraphClient,
  root: RootConfig,
  deps: ItemDeps,
  startUrl: string,
  includeDeletions: boolean,
): AsyncGenerator<WalkStep> {
  let url = startUrl;
  for (;;) {
    if (deps.session.signal.aborted) return;
    const page = await client.request<GraphDeltaPage>(url);
    if (!page['@odata.deltaLink'] && !page['@odata.nextLink']) {
      // v1 parity (backfill.ts): a delta feed that ends without either link
      // is a genuine upstream contract violation, not a recoverable case.
      throw new Error(
        `onedrive: root ${root.rootFolderId} delta feed ended without a nextLink or deltaLink`,
      );
    }
    for await (const chunk of pageChunks(page, root, deps, { includeDeletions })) {
      if (!chunk.pageComplete) {
        yield { ...chunk };
      } else if (page['@odata.deltaLink']) {
        yield { ...chunk, token: tokenFromDeltaLink(page['@odata.deltaLink']) };
      } else {
        yield { ...chunk, nextLink: page['@odata.nextLink'] };
      }
    }
    if (page['@odata.deltaLink']) return;
    url = page['@odata.nextLink']!;
  }
}

/** Order-independent set equality on root ids: a merely REORDERED root list
 *  must never force a re-walk. `undefined` (a pre-folder-scope cursor) is
 *  always a mismatch. */
function sameScope(a: string[] | undefined, b: string[]): boolean {
  if (!a || a.length !== b.length) return false;
  const seen = new Set(a);
  return b.every((id) => seen.has(id));
}

/**
 * Reconcile an incoming cursor with the CURRENT configured roots before any
 * walking happens. On a match the cursor rides through untouched (a
 * mid-backfill resume keeps its `next_link`). On a mismatch — a scope change
 * that did not come through `manageFolders`, or any cursor written before
 * `scope_roots` existed — it is rebuilt:
 *
 *  - `delta_tokens` is PRUNED to the configured roots. This is the only
 *    place, besides `manageFolders`, where a token key is ever removed.
 *  - the `backfill` key is DELETED OUTRIGHT, never reset to
 *    `{root_index: 0}`: `root_index` is a position in the roots array and
 *    `next_link` is one specific root's opaque page URL, so a stale pair
 *    resumes paging the WRONG root's result set (spec-reality-diff A5c). A
 *    root that loses its resume link simply re-enumerates its pages from the
 *    start — bounded, and idempotent via `hashSkip`.
 *  - a configured root left with no token is then established by
 *    `backfill()` on this very pull, because `backfillDone` sees it lacking
 *    one. Established roots keep polling; nothing is re-downloaded.
 */
export function normalizeCursor(
  cursor: OneDriveCursor | null,
  roots: RootConfig[],
): OneDriveCursor | null {
  if (!cursor) return null;
  const configured = roots.map((r) => r.rootFolderId);
  if (sameScope(cursor.scope_roots, configured)) return cursor;
  const delta_tokens: Record<string, string> = {};
  for (const id of configured) {
    const token = cursor.delta_tokens?.[id];
    if (typeof token === 'string') delta_tokens[id] = token;
  }
  return { delta_tokens, scope_roots: configured };
}

function backfillDone(cursor: OneDriveCursor | null, roots: RootConfig[]): boolean {
  if (!cursor) return roots.length === 0;
  if (cursor.backfill) return false;
  return roots.every((r) => cursor.delta_tokens[r.rootFolderId] != null);
}

/**
 * Establishes any root lacking a delta token, one Graph delta page per
 * batch. A crash mid-root resumes via `cursor.backfill.next_link` (v1
 * parity); a crash between roots resumes at `cursor.backfill.root_index`. A
 * root already holding a token is skipped — reached only when a NEW root is
 * added after some roots are already live (v1 had no such path: it always
 * restarts every root from index 0 on a config change).
 */
async function* backfill(
  client: GraphClient,
  session: Session,
  query: Query,
  fetchFn: (url: string, init?: unknown) => Promise<unknown>,
  cursor: OneDriveCursor | null,
  roots: RootConfig[],
  processed: Set<string>,
  budget: BatchBudget,
): AsyncGenerator<Batch<OneDriveCursor, OneDriveItem>> {
  const deps: ItemDeps = { client, session, query, fetchFn, processed, budget, ignored: new IgnoreTally() };
  if (roots.length === 0) {
    yield { phase: 'live', items: [], cursor: { delta_tokens: cursor?.delta_tokens ?? {} } };
    return;
  }

  const deltaTokens: Record<string, string> = { ...(cursor?.delta_tokens ?? {}) };
  let startIdx: number;
  let resumeNextLink: string | undefined;
  if (cursor?.backfill) {
    startIdx = Math.min(cursor.backfill.root_index, roots.length - 1);
    resumeNextLink = cursor.backfill.next_link;
  } else {
    const idx = roots.findIndex((r) => deltaTokens[r.rootFolderId] == null);
    startIdx = idx === -1 ? roots.length : idx;
  }

  for (let i = startIdx; i < roots.length; i++) {
    if (session.signal.aborted) return;
    const root = roots[i];
    if (deltaTokens[root.rootFolderId] != null) continue; // already established

    const resuming = i === startIdx && resumeNextLink;
    const startUrl = resuming
      ? resumeNextLink!
      : `${GRAPH_BASE}/me/drive/items/${root.rootFolderId}/delta`;
    // The cursor that fetched the CURRENT page — what an intermediate
    // (budget-flushed) chunk commits under, so a crash replays this page.
    let pageCursor: OneDriveCursor = {
      delta_tokens: { ...deltaTokens },
      backfill: resuming ? { root_index: i, next_link: resumeNextLink } : { root_index: i },
    };

    for await (const step of walkRootToDeltaLink(client, root, deps, startUrl, false)) {
      if (session.signal.aborted) return;
      if (!step.pageComplete) {
        yield { phase: 'backfill', items: step.items, deletions: step.deletions, cursor: pageCursor };
      } else if (step.token != null) {
        deltaTokens[root.rootFolderId] = step.token;
        const more = i + 1 < roots.length;
        const nextCursor: OneDriveCursor = more
          ? { delta_tokens: { ...deltaTokens }, backfill: { root_index: i + 1 } }
          : { delta_tokens: { ...deltaTokens } };
        yield { phase: more ? 'backfill' : 'live', items: step.items, deletions: step.deletions, cursor: nextCursor };
      } else {
        pageCursor = {
          delta_tokens: { ...deltaTokens },
          backfill: { root_index: i, next_link: step.nextLink },
        };
        yield { phase: 'backfill', items: step.items, deletions: step.deletions, cursor: pageCursor };
      }
    }
  }
  deps.ignored.logIfAny(session);
}

/** Establishes ONE root fully, always yielding `phase: 'live'` batches (no
 *  `cursor.backfill` marker) — used by delta()'s resync reprime (v1 parity:
 *  a resync never flips account status back to 'backfilling'). */
async function* establishRootLive(
  client: GraphClient,
  root: RootConfig,
  deps: ItemDeps,
  deltaTokens: Record<string, string>,
): AsyncGenerator<Batch<OneDriveCursor, OneDriveItem>> {
  const startUrl = `${GRAPH_BASE}/me/drive/items/${root.rootFolderId}/delta`;
  for await (const step of walkRootToDeltaLink(client, root, deps, startUrl, false)) {
    if (deps.session.signal.aborted) return;
    if (step.token != null) deltaTokens[root.rootFolderId] = step.token;
    yield {
      phase: 'live',
      items: step.items,
      deletions: step.deletions,
      cursor: { delta_tokens: { ...deltaTokens } },
    };
  }
}

/** Polls one root's already-established delta feed (v1 `delta.ts`'s
 *  `walkRootDelta` with a token). Deletions ARE surfaced here (unlike
 *  backfill/establish). */
async function* pollRoot(
  client: GraphClient,
  root: RootConfig,
  deps: ItemDeps,
  startUrl: string,
  deltaTokens: Record<string, string>,
): AsyncGenerator<Batch<OneDriveCursor, OneDriveItem>> {
  let url = startUrl;
  for (;;) {
    if (deps.session.signal.aborted) return;
    const page = await client.request<GraphDeltaPage>(url);
    // Intermediate pages (and intermediate chunks of any page) commit under
    // the PRIOR token — the new one is only stored once the deltaLink page's
    // final chunk is out. A feed that ends without a link keeps the prior
    // token (v1 `walkRootDelta` parity: `return token ?? ''`) but still
    // surfaces its items.
    for await (const chunk of pageChunks(page, root, deps, { includeDeletions: true })) {
      if (chunk.pageComplete && page['@odata.deltaLink']) {
        deltaTokens[root.rootFolderId] = tokenFromDeltaLink(page['@odata.deltaLink']);
      }
      yield {
        phase: 'live',
        items: chunk.items,
        deletions: chunk.deletions,
        cursor: { delta_tokens: { ...deltaTokens } },
      };
    }
    if (page['@odata.deltaLink'] || !page['@odata.nextLink']) return;
    url = page['@odata.nextLink'];
  }
}

/**
 * Incremental poll across every configured root (v1 `delta.ts`). A root
 * whose token is rejected (410 resync) is re-primed inline via a full
 * re-establish (v1 parity) — self-contained per root, so one root's resync
 * never blocks another's poll.
 */
async function* delta(
  client: GraphClient,
  session: Session,
  query: Query,
  fetchFn: (url: string, init?: unknown) => Promise<unknown>,
  cursor: OneDriveCursor,
  roots: RootConfig[],
  processed: Set<string>,
  budget: BatchBudget,
): AsyncGenerator<Batch<OneDriveCursor, OneDriveItem>> {
  const deps: ItemDeps = { client, session, query, fetchFn, processed, budget, ignored: new IgnoreTally() };
  const deltaTokens: Record<string, string> = { ...cursor.delta_tokens };

  for (const root of roots) {
    if (session.signal.aborted) return;
    // Invariant enforced by pull()'s dispatch: delta() only runs once
    // backfillDone(cursor, roots) is true, which requires EVERY root here to
    // already hold a token — so `token` is never undefined in practice. A
    // root added later without one is caught by backfillDone and routed back
    // through backfill() instead (see module doc's self-heal note).
    const token = deltaTokens[root.rootFolderId]!;
    const startUrl = `${GRAPH_BASE}/me/drive/items/${root.rootFolderId}/delta?token=${encodeURIComponent(token)}`;
    try {
      yield* pollRoot(client, root, deps, startUrl, deltaTokens);
    } catch (e) {
      if (isAuthError(e)) throw e;
      if (!isResyncRequired(errText(e))) throw e;
      session.log(
        'warn',
        `onedrive delta: token invalid for root ${root.rootFolderId} (${errText(e)}) — re-priming`,
      );
      yield* establishRootLive(client, root, deps, deltaTokens);
    }
  }
  deps.ignored.logIfAny(session);
}

/** Where a tracked root sits, for containment comparison. `drive` is the
 *  literal `'me'` for the caller's own OneDrive — Graph spells those item
 *  paths `/drive/root:` — and the remote drive id for a "Shared with me"
 *  root that reports `/drives/{id}/root:`. The PATH FORM, not
 *  `parentReference.driveId`, is the discriminator, because the drive-root
 *  catch-all is answered without an item lookup at all and therefore has no
 *  `parentReference` to read. `path` is the folder's absolute path inside
 *  that drive with a LEADING and TRAILING slash, so `startsWith` compares
 *  whole segments (`/Alpha/` never prefixes `/AlphaBeta/`). */
export interface RootLocation {
  drive: string;
  path: string;
}

const MY_DRIVE = 'me';

/**
 * Resolve one tracked root to a comparable location, or `null` when it is
 * INCOMPARABLE — it can neither cover nor be covered by anything.
 *
 * `null` is returned for exactly two cases, both of which make archiving the
 * correct answer: a "Shared with me" root whose `parentReference.path` is
 * absent or relative to a drive this connector never resolves (the same
 * limitation `buildDisplayPath` documents in path-resolver.ts), and a root
 * that is gone upstream (404/410).
 *
 * ANY OTHER failure RETHROWS, aborting the whole save — including a malformed
 * percent-escape out of `decodeURIComponent`, which is deliberately not
 * caught. With no `reconcile()` in this connector, guessing "incomparable" on
 * a transient Graph 5xx would archive a root that a retained ancestor
 * actually covers, and nothing would ever restore it: the covering root keeps
 * its delta token, and `hashSkip` never re-emits a LIVE row. A failed save is
 * recoverable by clicking Save again; a wrong one is not.
 *
 * Note the client retries a 5xx four times before throwing
 * (`client.ts:91-93`, `:156-165`), while a 404 throws on the first response.
 */
export async function resolveRootLocation(
  client: GraphClient,
  rootFolderId: string,
): Promise<RootLocation | null> {
  if (rootFolderId === ONEDRIVE_DRIVE_ROOT_ID) return { drive: MY_DRIVE, path: '/' };
  type RootItem = { name?: string; parentReference?: { driveId?: string; path?: string } };
  let item: RootItem;
  try {
    item = await client.request<RootItem>(
      `${GRAPH_BASE}/me/drive/items/${encodeURIComponent(rootFolderId)}?$select=id,name,parentReference`,
    );
  } catch (e) {
    if (e instanceof GraphApiError && (e.status === 404 || e.status === 410)) return null;
    throw e;
  }
  const raw = item.parentReference?.path;
  const name = item.name;
  if (!raw || !name) return null;
  const mine = /^\/drive\/root:?/.exec(raw);
  const foreign = /^\/drives\/([^/]+)\/root:?/.exec(raw);
  let drive: string;
  let parent: string;
  if (mine) {
    drive = MY_DRIVE;
    parent = raw.slice(mine[0].length);
  } else if (foreign) {
    drive = foreign[1];
    parent = raw.slice(foreign[0].length);
  } else {
    return null; // an unknown/relative path form — incomparable
  }
  // Graph percent-encodes path SEGMENTS but returns `name` raw, so the parent
  // must be decoded before the two are joined or `/My%20Docs` would never
  // prefix `/My Docs/`.
  return { drive, path: `${decodeURIComponent(parent)}/${name}/`.replace(/\/+/g, '/') };
}

/**
 * `p` is `root` itself or lives anywhere under it, respecting the separator
 * BOUNDARY: `/Docs` does not contain `/DocsBackup`. Same shape as core's
 * `@shared/folder-paths.ts:isUnder`, which is where this belongs
 * conceptually — it is duplicated rather than imported because a connector
 * is a standalone bundle with no access to core's internals.
 *
 * C-46/D4: `covers`/`overlaps` used a bare `startsWith`, so a removed sibling
 * whose NAME EXTENDS a retained root read as contained and was silently
 * exempted from archival. Measured reach of that defect through
 * `manageFolders` TODAY: none — every path `resolveRootLocation` builds
 * already ends in `/` (it appends one, and the drive root is the literal
 * `'/'`), and `/DocsBackup/` does not `startsWith` `/Docs/`. The boundary
 * check makes the predicates correct INDEPENDENT of that undocumented
 * convention instead of resting on it, which is the point: nothing in the
 * types says a `RootLocation.path` is slash-terminated, and both predicates
 * are exported.
 *
 * Single separator only (`/`): these are Graph paths, never OS paths, so
 * core's extra `\\` handling has nothing to match here.
 */
const isUnderPath = (p: string, root: string): boolean => {
  if (p === root) return true;
  if (!p.startsWith(root)) return false;
  return root.endsWith('/') || p.charAt(root.length) === '/';
};

/** `retained` is an ancestor-or-self of `removed`: NOTHING under the removed
 *  root leaves scope, so it must not be archived (DECISIONS R8). Exported for
 *  the C-46/D4 boundary unit tests, which are the only way to exercise a
 *  non-slash-terminated path. */
export const covers = (retained: RootLocation | null, removed: RootLocation | null): boolean =>
  retained !== null &&
  removed !== null &&
  retained.drive === removed.drive &&
  isUnderPath(removed.path, retained.path);

/** The two roots share a subtree in EITHER direction — the precondition for
 *  first-root-wins attribution to have named the wrong root. Boundary-aware
 *  for the same reason as `covers` (C-46/D4). */
export const overlaps = (a: RootLocation | null, b: RootLocation | null): boolean =>
  a !== null &&
  b !== null &&
  a.drive === b.drive &&
  (isUnderPath(a.path, b.path) || isUnderPath(b.path, a.path));

export interface OneDriveTestSeams extends Partial<Pick<GraphClientDeps, 'sleep' | 'random'>> {
  batchByteBudget?: number;
  batchItemLimit?: number;
}

export function createOneDriveSource(
  host: HostFor<'net' | 'query'>,
  // Test seam only: GraphClient's sleep/random are injectable so retry tests
  // never actually wait, and the batch budgets shrink so chunking is
  // testable with tiny fixtures; production callers omit this.
  seams?: OneDriveTestSeams,
): Source<OneDriveCursor, OneDriveItem> {
  const { batchByteBudget, batchItemLimit, ...clock } = seams ?? {};
  const budget: BatchBudget = {
    bytes: batchByteBudget ?? BATCH_BYTE_BUDGET,
    items: batchItemLimit ?? BATCH_ITEM_LIMIT,
  };
  const clientFor = (session: Session): GraphClient =>
    new GraphClient({
      fetch: host.net.fetch,
      getToken: () => requireToken(session),
      ...clock,
    });

  return {
    descriptor: {
      id: 'onedrive',
      name: 'OneDrive',
      documentTypes: ['file'],
      auth: 'oauth',
      multiAccount: true,
      cadence: { every: '15m' },
      /** Scoped to a user-selected set of folder roots: enables the Tracked
       *  folders card and `accounts:start-manage-folders`. A descriptor with
       *  this flag MUST implement `manageFolders` — it does, below. */
      folderScope: true,
    },

    async connect(auth: AuthChannel) {
      auth.status('Waiting for Microsoft sign-in…');
      const creds: Credentials = await auth.oauth(FILES_SCOPES);
      const accessToken = creds.accessToken;
      if (!accessToken) {
        throw new Error('onedrive: Microsoft sign-in returned no access token');
      }
      const client = new GraphClient({
        fetch: host.net.fetch,
        getToken: async () => accessToken,
        ...clock,
      });

      auth.status('Fetching Microsoft profile…');
      const me = await client.request<{ mail?: string; userPrincipalName?: string }>(
        `${GRAPH_BASE}/me?$select=mail,userPrincipalName,id`,
      );
      const identifier = me.mail ?? me.userPrincipalName;
      if (!identifier) {
        throw new Error('onedrive: Graph /me response missing mail and userPrincipalName');
      }

      // No needsReauth lookup any more: re-auth is `reauthenticate(account,
      // auth)`, which is handed the exact account and verifies the returned
      // identity against it. The deleted heuristic matched on
      // (source, identifier, status) with `.find()` — first match wins — so
      // two accounts sharing an identifier were indistinguishable, and an
      // account stored from `mail` never matched a `/me` that answered only
      // `userPrincipalName`. `connect()` is now purely "add a NEW account".
      //
      // The platform's shared folder-picker: lazy tree over the connect-time
      // client, multi-select with covering roots. A user cancel rejects —
      // let that propagate out of connect().
      const picked = await auth.pickFolders({
        modes: [
          { key: 'my-files', label: 'My files' },
          { key: 'shared', label: 'Shared with me' },
        ],
        multiSelect: true,
        roots: async (mode) =>
          mode === 'my-files'
            ? [{ id: ONEDRIVE_DRIVE_ROOT_ID, name: 'OneDrive', hasChildren: true }]
            : listSharedRoots(client),
        children: (id) => listChildFolders(client, id),
        count: (id) => countChildren(client, id),
        purpose: 'connect',
      });
      if (picked.length === 0) throw new Error('onedrive: no folders selected');
      return {
        identifier,
        config: { folderRoots: picked.map((n) => ({ id: n.id, name: n.name })) },
      };
    },

    async *pull(session: Session, cursor: OneDriveCursor | null) {
      const client = clientFor(session);
      const roots = rootsConfig(session);
      const scopeRoots = roots.map((r) => r.rootFolderId);
      const start = normalizeCursor(cursor, roots);
      const processed = new Set<string>();
      const walk = !backfillDone(start, roots)
        ? backfill(client, session, host.query, host.net.fetch, start, roots, processed, budget)
        : delta(client, session, host.query, host.net.fetch, start!, roots, processed, budget);
      // ONE stamping site. There are SEVEN internal cursor literals across
      // backfill()/establishRootLive()/pollRoot() (`src/source.ts:595`,
      // `:621`, `:634`, `:635`, `:638`, `:666`, `:698`); stamping each of
      // them would make a single missed site look like a permanent scope
      // mismatch, i.e. a token prune on every pull forever. Stamping the
      // yielded batch instead is impossible to get partially wrong.
      for await (const batch of walk) {
        yield { ...batch, cursor: { ...batch.cursor, scope_roots: scopeRoots } };
      }
    },

    toDocument(item: OneDriveItem): DocumentInput {
      const f = item.file;
      const mime = f.file?.mimeType ?? 'application/octet-stream';
      return {
        externalId: f.id,
        type: 'file',
        title: f.name,
        markdown: item.markdown,
        ...(item.bytes ? { binary: { bytes: item.bytes, mime, filename: f.name } } : {}),
        url: f.webUrl ?? 'https://onedrive.live.com/',
        /** Folder scope: the tracked root whose delta feed emitted this item
         *  (first-root-wins under overlap — `deps.processed`, `:463-468`).
         *  ALWAYS set: every walk carries its `RootConfig` down to
         *  `buildItem`/`failedItem`, so this connector never reaches the
         *  engine's NULL-and-warn path (DECISIONS R5). `metadata.
         *  root_folder_id` below is kept as-is — core's v3 migration reads
         *  that spelling to backfill `scope_root_id` on existing rows. */
        scopeRootId: item.rootFolderId,
        metadata: {
          drive_item_id: f.id,
          drive_id: f.parentReference?.driveId ?? '',
          mime_type: mime,
          size_bytes: f.size ?? null,
          etag: f.eTag ?? null,
          ctag: f.cTag ?? null,
          display_path: item.displayPath,
          root_folder_id: item.rootFolderId,
          extraction_status: item.extractionStatus,
          modified_time: f.lastModifiedDateTime ?? null,
          // Engine vision/classify aliases: kiagent-core's vision pipeline
          // reads metadata.mime / filename / sizeBytes, not the v1-named keys
          // above. Every OneDrive doc is type 'file' (no native-doc split),
          // so these are always emitted (mirrors the google-docs-kia-
          // connector template's conditional block, unconditional here).
          mime,
          filename: f.name,
          ...(f.size != null ? { sizeBytes: f.size } : {}),
        },
        createdAt: f.createdDateTime ?? f.lastModifiedDateTime ?? null,
      };
    },

    /** Random-access bytes for the engine's deep-extraction passes (the
     *  vision worker's OCR/VLM two-pass pulls image/pdf bytes back through
     *  here). Pre-signed download URLs expire, so this always fetches a
     *  fresh one rather than trusting anything cached in metadata.
     *  Classifies from stored metadata FIRST — an ignored route returns null
     *  before any client/OAuth/network work happens at all (no
     *  `clientFor(session)`, no downloadUrl refresh). */
    async fetchBytes(session: Session, doc: Document): Promise<Uint8Array | null> {
      const meta = doc.metadata as Record<string, unknown>;
      const itemId = meta.drive_item_id;
      if (typeof itemId !== 'string' || !itemId) return null;
      const mime = typeof meta.mime_type === 'string' ? meta.mime_type : '';
      const filename = typeof meta.filename === 'string' ? meta.filename : (doc.title ?? '');
      const size = typeof meta.size_bytes === 'number' ? meta.size_bytes : undefined;
      if (chooseRoute(mime, filename, size).kind === 'ignore') return null;
      const client = clientFor(session);
      try {
        const downloadUrl = await refreshDownloadUrl(client, itemId);
        if (!downloadUrl) return null;
        return await downloadBytes(host.net.fetch, downloadUrl);
      } catch (e) {
        if (e instanceof GraphApiError && (e.status === 404 || e.status === 410)) {
          return null; // gone upstream — the next delta poll will archive it
        }
        throw e;
      }
    },

    /**
     * Edit this account's tracked roots using its EXISTING credentials. Never
     * authenticates — `FolderSelectionChannel` has no `oauth` verb — and
     * persists nothing: core's `applyFolderScope` writes config, cursor and
     * archival in ONE transaction.
     *
     * The Graph client is `clientFor(session)`, which calls
     * `session.credentials()` fresh per attempt. `connect()`'s client closes
     * over a FROZEN access token and would die on a long-open picker
     * (spec-reality-diff D9); do not copy that wiring here.
     */
    async manageFolders(
      session: Session,
      channel: FolderSelectionChannel,
    ): Promise<FolderScopeUpdate<OneDriveCursor>> {
      const client = clientFor(session);
      const prior = rootsConfig(session);
      const priorIds = new Set(prior.map((r) => r.rootFolderId));

      channel.status('Loading your OneDrive folders…');
      const picked = await channel.pickFolders({
        modes: [
          { key: 'my-files', label: 'My files' },
          { key: 'shared', label: 'Shared with me' },
        ],
        multiSelect: true,
        roots: async (mode) =>
          mode === 'my-files'
            ? [{ id: ONEDRIVE_DRIVE_ROOT_ID, name: 'OneDrive', hasChildren: true }]
            : listSharedRoots(client),
        children: (id) => listChildFolders(client, id),
        count: (id) => countChildren(client, id),
        // The complete current covering set, pre-checked AND removable.
        selected: prior.map((r) => ({ id: r.rootFolderId, name: r.rootName, hasChildren: true })),
        purpose: 'manage',
      });
      if (picked.length === 0) throw new Error('onedrive: no folders selected');

      // Root ORDER is semantic — it drives backfill()'s root_index and the
      // first-root-wins attribution — and the picker's return order is not
      // stable. Keep retained roots in their prior config order; append what
      // is new.
      const byId = new Map(picked.map((n) => [n.id, n]));
      const ordered = [
        ...prior.filter((r) => byId.has(r.rootFolderId)).map((r) => byId.get(r.rootFolderId)!),
        ...picked.filter((n) => !priorIds.has(n.id)),
      ];
      const nextIds = ordered.map((n) => n.id);
      const removed = prior.filter((r) => !byId.has(r.rootFolderId));
      const retained = prior.filter((r) => byId.has(r.rootFolderId));

      const archiveScopeRootIds: string[] = [];
      const dropTokens = new Set<string>();
      if (removed.length > 0) {
        channel.status('Checking which folders are still covered…');
        const where = new Map<string, RootLocation | null>();
        for (const id of [
          ...removed.map((r) => r.rootFolderId),
          ...retained.map((r) => r.rootFolderId),
        ]) {
          where.set(id, await resolveRootLocation(client, id));
        }
        for (const r of removed) {
          const loc = where.get(r.rootFolderId) ?? null;
          let covered = false;
          for (const k of retained) {
            const kloc = where.get(k.rootFolderId) ?? null;
            // A retained ANCESTOR still contains the removed root's whole
            // subtree, so nothing leaves scope — the empty archive set is the
            // correct answer, not a fudge (DECISIONS R8).
            if (covers(kloc, loc)) covered = true;
            // Overlap in EITHER direction means first-root-wins attribution
            // (a config-ORDER artifact, `deps.processed`, `:463-468`) may
            // have stamped documents in the shared region with the wrong
            // root's id. Drop the retained root's delta token so its next
            // pull re-walks the region (spec-reality-diff A5b). Cost:
            // `hashSkip` gates before any download, so every LIVE row in the
            // shared region is re-read as metadata pages and never
            // re-downloaded. Rows this save ARCHIVES are the exception:
            // `hashSkip` returns false for an archived row
            // (`src/source.ts:296`), so the re-walk re-emits — hence
            // un-archives — any of them a retained root genuinely still
            // contains, and those DO cost a re-download. That is the price of
            // attribution that self-heals, and it is the only such mechanism
            // this connector has (there is no `reconcile()`).
            if (overlaps(kloc, loc)) dropTokens.add(k.rootFolderId);
          }
          if (!covered) archiveScopeRootIds.push(r.rootFolderId);
        }
      }

      const prev = (session.account.cursor ?? null) as OneDriveCursor | null;
      const delta_tokens: Record<string, string> = {};
      for (const id of nextIds) {
        if (dropTokens.has(id)) continue;
        const token = prev?.delta_tokens?.[id];
        if (typeof token === 'string') delta_tokens[id] = token;
      }
      // Pruned to the configured set in this SAME write — the append-only
      // `delta_tokens` leak is closed here and in `normalizeCursor`, nowhere
      // else. NO `backfill` key, ever: a stale root_index/next_link pair
      // would resume paging the wrong root (A5c).
      const cursor: OneDriveCursor = { delta_tokens, scope_roots: [...nextIds] };

      // A-3 / R5. `archiveNullScoped` asks core to archive this account's live
      // rows whose `scope_root_id` is NULL — rows core's v3 migration could
      // not attribute (a mass-archive refusal, unreadable metadata). This
      // connector cannot SEE them: `Query` offers `byExternalId`, `document`,
      // `children`, `search`, `count` and `countBy({field:'from'|'label'})`,
      // none of them keyed on scope. So the flag is derived from the CURSOR
      // WE RETURN rather than from row evidence, and the derivation IS A-3's
      // pairing condition stated as an invariant: request the repair exactly
      // when this cursor re-establishes EVERY configured root from scratch —
      // no delta token survives, and this branch never writes a `backfill`
      // key. Under that condition the very next pull re-walks every root and,
      // because `hashSkip` does not skip an archived row, re-emits (hence
      // un-archives) every NULL row still genuinely in scope. Archiving NULL
      // rows WITHOUT that re-walk is permanent loss — `contentHash` excludes
      // scope and this connector hashSkips — which is exactly what A-3
      // forbids.
      const archiveNullScoped = Object.keys(delta_tokens).length === 0;
      // ⚠️ C-34 — CORE DECLINES TO ACT ON THIS FLAG IN THIS TRAIN, BY
      // DESIGN. Keep computing and emitting it: it is part of the frozen
      // `FolderScopeUpdate` and it is how a source states intent. But expect
      // NO effect — Task 3 drops `archiveNullScoped` from
      // `applyFolderScope`'s store input type and Task 7 does not forward it
      // (it warns that a source asked and was refused). Do NOT "fix" this
      // connector when you discover the field has no effect, and do not
      // delete the derivation: it is the invariant a later, safe repair path
      // would key on.
      //
      // Why core refuses the pairing argued for above: the archive would land
      // BEFORE there is any proof the re-establishing walk actually LISTED
      // the row. An archived row IS re-emitted (hashSkip's
      // `if (!existing || existing.archivedAt) return false;` at
      // `src/source.ts:296`) — but only for rows the walk reaches, and a LIVE
      // NULL-scoped row has no other re-stamp path, because core's
      // upsertDocument early-returns on `content_hash === hash &&
      // archived_at === null` (kiagent-core write-tx.ts:170-176). Anything
      // the walk misses stays archived for good, and this connector has NO
      // `reconcile()` (`src/source.ts:62`) — no later pass ever notices. On a
      // `needsReauth` account the walk does not even start (core
      // boot.ts:194-202: 'needsReauth' is a RESTING state; only the user's
      // explicit Retry or a fresh connect restarts the loop). What would have
      // to exist before core could honour the flag is an archive-AFTER-proof
      // predicate shaped like core's `reconcile` (write-tx.ts:512-538:
      // `seq <= ?` AND `NOT EXISTS (… reconcile_listing …)`), plus the
      // listing pass this connector does not have — not a boolean.

      // A-2: the legacy R1 `roots` mirror is CORE's to write (in the v3
      // migration and in applyFolderScope) and core's alone. This connector
      // neither writes nor strips it — any existing `roots` key rides through
      // untouched and core overwrites it, derived from `folderRoots`, inside
      // the same transaction. Stripping it here is what left the installed,
      // non-auto-updating 2.0.5 build with no `roots` at all after the first
      // Save, which is precisely the failure R1 exists to prevent.
      const config: Record<string, unknown> & FolderScopedConfig = {
        ...session.account.config,
        folderRoots: ordered.map((n) => ({ id: n.id, name: n.name })),
      };

      return { config, cursor, archiveScopeRootIds, archiveNullScoped };
    },

    /**
     * Re-authenticate THIS account. Returns nothing: reconnect never changes
     * scope, and config/cursor are untouched.
     *
     * Identity is verified against BOTH Graph identity fields, trimmed and
     * case-folded. `/me` may answer with `mail` on one sign-in and only
     * `userPrincipalName` on the next; that is the same person, not a
     * mismatch. A genuine mismatch is a SourcePermanentError — retrying the
     * same wrong Microsoft account can never help, and silently repointing
     * an account at a different identity is how a corpus gets mixed.
     */
    async reauthenticate(account: Account, auth: AuthChannel): Promise<void> {
      auth.status('Waiting for Microsoft sign-in…');
      const creds: Credentials = await auth.oauth(FILES_SCOPES);
      const accessToken = creds.accessToken;
      if (!accessToken) {
        throw new Error('onedrive: Microsoft sign-in returned no access token');
      }
      const client = new GraphClient({
        fetch: host.net.fetch,
        getToken: async () => accessToken,
        ...clock,
      });

      auth.status('Fetching Microsoft profile…');
      const me = await client.request<{ mail?: string; userPrincipalName?: string }>(
        `${GRAPH_BASE}/me?$select=mail,userPrincipalName,id`,
      );
      const identities = [me.mail, me.userPrincipalName].filter(
        (v): v is string => typeof v === 'string' && v.trim() !== '',
      );
      if (identities.length === 0) {
        throw new Error('onedrive: Graph /me response missing mail and userPrincipalName');
      }
      const want = account.identifier.trim().toLowerCase();
      if (!identities.some((v) => v.trim().toLowerCase() === want)) {
        throw new SourcePermanentError(
          `onedrive: signed in as ${identities[0]}, but this account is ${account.identifier} — sign in with the original Microsoft account`,
        );
      }
    },
  };
}
