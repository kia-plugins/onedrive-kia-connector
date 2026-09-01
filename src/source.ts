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
 *    mime-route.ts for the mime gate this replaces v1's `isConvertibleMime`
 *    with — broadened to include images (v1 left images unsupported; the v2
 *    vision worker OCRs them via `fetchBytes`).
 *  - A 25 MiB download cap (`MAX_BINARY_BYTES`) is ADDED: v1 had none (v1 gap
 *    — the same one the google-docs-kia-connector template called out and
 *    fixed for Drive). An oversized file is indexed as metadata-only
 *    (`extraction_status: 'too-large'`) instead of downloaded.
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
  AuthChannel,
  Batch,
  Credentials,
  Document,
  DocumentInput,
  ExternalRef,
  HostFor,
  Query,
  Session,
  Source,
} from '@kiagent/connector-sdk';
import {
  GRAPH_BASE,
  GraphApiError,
  GraphClient,
  OneDriveAuthError,
  isAuthError,
  type GraphClientDeps,
} from './client';
import { buildDisplayPath } from './path-resolver';
import { isConvertibleMime } from './mime-route';
import { countChildren, listChildFolders, listSharedRoots } from './tree';

export { GRAPH_BASE };

/** v1 FILES_SCOPES minus `openid`/`email`/`profile`/`offline_access` — the
 *  platform's Microsoft OAuth provider appends those itself (see brief: the
 *  extension only requests Graph RESOURCE scopes). */
export const FILES_SCOPES = ['Files.Read.All', 'User.Read'];

/** v1 had NO cap (whole file into a Buffer) — v2 binds one (see module doc). */
export const MAX_BINARY_BYTES = 25 * 1024 * 1024;

/**
 * Per-batch flush budget. A Graph delta page is server-sized (typically a
 * few hundred items) and every convertible file in it — images included, for
 * vision OCR — carries up to MAX_BINARY_BYTES of downloaded bytes. Holding a
 * whole page's bytes at once (then structured-cloning them over IPC in ONE
 * message) put a photo-heavy personal OneDrive past the extension process's
 * heap, and because the cursor only advances per page the same page replayed
 * on every retry: a deterministic crash loop ("extension process exited").
 * So a page is flushed to the engine in sub-page chunks once the accumulated
 * bytes or entry count cross these budgets — see `pageChunks`.
 */
export const BATCH_BYTE_BUDGET = 32 * 1024 * 1024;
export const BATCH_ITEM_LIMIT = 100;

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
  /** '' for metadata-only rows (no binary, no conversion enrollment); null
   *  for binary items (the engine converts). */
  markdown: string | null;
  bytes?: Uint8Array;
  extractionStatus: 'ok' | 'unsupported' | 'too-large' | 'failed';
  displayPath: string;
  rootFolderId: string;
}

interface GraphDeltaPage {
  value?: DriveItem[];
  '@odata.nextLink'?: string;
  '@odata.deltaLink'?: string;
}

const errText = (e: unknown): string => (e instanceof Error ? e.message : String(e));

/** Normalize account config to the tracked roots. Deduped by rootFolderId —
 *  first entry wins. Falls back to the OneDrive root when the picker somehow
 *  stored nothing (defensive only — `connect()` always rejects an empty
 *  selection, so this path is not expected to run in practice). */
export function rootsConfig(session: Session): RootConfig[] {
  const cfg = session.account.config as { roots?: unknown };
  const parsed: RootConfig[] = [];
  if (Array.isArray(cfg.roots)) {
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
    parsed.push({ rootFolderId: 'root', rootName: 'OneDrive' });
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

interface ItemDeps {
  client: GraphClient;
  session: Session;
  query: Query;
  fetchFn: (url: string, init?: unknown) => Promise<unknown>;
  /** Item ids already emitted THIS pull() call — guards overlapping
   *  configured roots (see module doc). */
  processed: Set<string>;
  budget: BatchBudget;
}

/** Query-first content-hash skip: an unchanged, still-live document is never
 *  re-downloaded (v1 ingest.ts's `canShortcut`, minus the metadata-only
 *  `last_seen_at` refresh — the v2 engine owns row freshness). An ARCHIVED
 *  match does NOT skip: re-emitting is what un-archives a doc that moved back
 *  into scope. A 'failed' row is never pinned behind an unchanged eTag —
 *  v1's exact rationale (possibly just a quota storm; must retry). */
async function hashSkip(deps: ItemDeps, itemId: string, etag: string | undefined): Promise<boolean> {
  if (!etag) return false;
  const existing = await deps.query.byExternalId(deps.session.account.id, itemId, 'file');
  if (!existing || existing.archivedAt) return false;
  const meta = existing.metadata as Record<string, unknown>;
  if (meta.extraction_status === 'failed') return false;
  return meta.etag === etag;
}

function metadataOnly(
  raw: DriveItem,
  extractionStatus: 'unsupported' | 'too-large' | 'failed',
  displayPath: string,
  rootFolderId: string,
): OneDriveItem {
  return { file: raw, markdown: '', extractionStatus, displayPath, rootFolderId };
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
 * Route one listed/changed item into an OneDriveItem (or null to skip it via
 * hash-skip). May do I/O (downloadUrl refresh / bytes download) — toDocument
 * stays pure.
 */
async function buildItem(raw: DriveItem, root: RootConfig, deps: ItemDeps): Promise<OneDriveItem | null> {
  const mime = raw.file?.mimeType ?? 'application/octet-stream';
  const displayPath = buildDisplayPath({ rootName: root.rootName }, raw);

  if (await hashSkip(deps, raw.id, raw.eTag)) return null;

  if (!isConvertibleMime(mime)) {
    return metadataOnly(raw, 'unsupported', displayPath, root.rootFolderId);
  }

  // v1 ingest.ts: Graph's per-item /delta response routinely omits
  // @microsoft.graph.downloadUrl (especially on MSA personal OneDrive) —
  // fall back to a fresh GET before giving up.
  let downloadUrl = raw['@microsoft.graph.downloadUrl'];
  if (!downloadUrl) {
    downloadUrl = await refreshDownloadUrl(deps.client, raw.id);
  }
  if (!downloadUrl) {
    return metadataOnly(raw, 'failed', displayPath, root.rootFolderId);
  }
  if (Number(raw.size ?? 0) > MAX_BINARY_BYTES) {
    return metadataOnly(raw, 'too-large', displayPath, root.rootFolderId);
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
  // Post-download cap: OneDrive items virtually always carry `size`, but the
  // cap is the guarantee — an unknown-size file must not slip past it.
  if (bytes.byteLength > MAX_BINARY_BYTES) {
    return metadataOnly(raw, 'too-large', displayPath, root.rootFolderId);
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
 * (v1 parity). Deletions are cheap but count against the entry limit so a
 * mass-delete page is still bounded.
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
      if (opts.includeDeletions) {
        const existing = await deps.query.byExternalId(deps.session.account.id, raw.id, 'file');
        if (existing) deletions.push({ externalId: raw.id, type: 'file' });
      }
    } else if (!raw.file || raw.id === root.rootFolderId || deps.processed.has(raw.id)) {
      // folder (not ingested, v1 parity) / the root item itself / overlapping
      // tracked roots — first root wins.
      continue;
    } else {
      deps.processed.add(raw.id);
      try {
        const item = await buildItem(raw, root, deps);
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
  const deps: ItemDeps = { client, session, query, fetchFn, processed, budget };
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
        yield { phase: 'backfill', items: step.items, cursor: pageCursor };
      } else if (step.token != null) {
        deltaTokens[root.rootFolderId] = step.token;
        const more = i + 1 < roots.length;
        const nextCursor: OneDriveCursor = more
          ? { delta_tokens: { ...deltaTokens }, backfill: { root_index: i + 1 } }
          : { delta_tokens: { ...deltaTokens } };
        yield { phase: more ? 'backfill' : 'live', items: step.items, cursor: nextCursor };
      } else {
        pageCursor = {
          delta_tokens: { ...deltaTokens },
          backfill: { root_index: i, next_link: step.nextLink },
        };
        yield { phase: 'backfill', items: step.items, cursor: pageCursor };
      }
    }
  }
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
    yield { phase: 'live', items: step.items, cursor: { delta_tokens: { ...deltaTokens } } };
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
  const deps: ItemDeps = { client, session, query, fetchFn, processed, budget };
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
}

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

      // Reconnect after an auth failure restores the stored selection: the
      // picker opens BLANK (no preselection), so re-asking here invites a
      // careless confirm that silently shrinks the corpus. The config rides
      // back VERBATIM. A healthy re-connect still runs the picker: it is the
      // only way to change the selection; the engine upserts on
      // (source, identifier) either way, so documents and cursor survive.
      const prior = (await host.query.accounts()).find(
        (a) => a.source === 'onedrive' && a.identifier === identifier && a.status === 'needsReauth',
      );
      if (prior) {
        auth.status('Restoring previous folder selection…');
        return { identifier, config: prior.config };
      }

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
          mode === 'my-files' ? [{ id: 'root', name: 'OneDrive', hasChildren: true }] : listSharedRoots(client),
        children: (id) => listChildFolders(client, id),
        count: (id) => countChildren(client, id),
      });
      if (picked.length === 0) throw new Error('onedrive: no folders selected');
      return {
        identifier,
        config: { roots: picked.map((n) => ({ rootFolderId: n.id, rootName: n.name })) },
      };
    },

    async *pull(session: Session, cursor: OneDriveCursor | null) {
      const client = clientFor(session);
      const roots = rootsConfig(session);
      const processed = new Set<string>();
      if (!backfillDone(cursor, roots)) {
        yield* backfill(client, session, host.query, host.net.fetch, cursor, roots, processed, budget);
      } else {
        yield* delta(client, session, host.query, host.net.fetch, cursor!, roots, processed, budget);
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
     *  fresh one rather than trusting anything cached in metadata. */
    async fetchBytes(session: Session, doc: Document): Promise<Uint8Array | null> {
      const meta = doc.metadata as Record<string, unknown>;
      const itemId = meta.drive_item_id;
      if (typeof itemId !== 'string' || !itemId) return null;
      const mime = typeof meta.mime_type === 'string' ? meta.mime_type : '';
      if (!isConvertibleMime(mime)) return null;
      const size = Number(meta.size_bytes ?? 0);
      if (Number.isFinite(size) && size > MAX_BINARY_BYTES) return null;
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
  };
}
