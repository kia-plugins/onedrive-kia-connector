/**
 * Shared offline test harness. The GENERIC plumbing — host-shaped JSON
 * responses, the exact-URL scripted fetch, the instant clock, and the fakes
 * for Session / AuthChannel — comes from `@kiagent/connector-sdk/testing`;
 * what stays here is the part no other connector can use: the fake Microsoft
 * Graph world (delta feeds, picker listings, pre-signed downloads), the
 * driveItem fixture builders, and the Query fake the ingest hash-skip needs.
 *
 * `jsonRes` / `instantClock` / the `HostResponse` type are re-exported so the
 * suites keep importing every fixture helper from this one module.
 *
 * Lives outside src/__tests__ so jest's default testMatch does not treat it
 * as a suite. Never bundled: build.mjs only follows imports from index.ts, so
 * the SDK's testing entrypoint (which pulls in node:child_process) is
 * reachable only from the suites.
 */
import type {
  Account,
  AuthChannel,
  Credentials,
  Document,
  FolderNode,
  FolderPickerSpec,
  HostFor,
  Query,
  Session,
} from '@kiagent/connector-sdk';
import type { HostResponse } from '@kiagent/connector-sdk/http';
import {
  fakeAuthChannel,
  fakeSession,
  instantClock,
  jsonRes,
  scriptedFetch,
} from '@kiagent/connector-sdk/testing';
import { GRAPH_BASE, type NetFetch } from '../client';
import type { DriveItem } from '../source';

export { instantClock, jsonRes };
export type { HostResponse };

export const bytesRes = (status: number, body: Uint8Array): HostResponse => ({
  status,
  statusText: '',
  headers: {},
  body,
});

const isHostResponse = (v: unknown): v is HostResponse =>
  typeof v === 'object' &&
  v !== null &&
  typeof (v as HostResponse).status === 'number' &&
  (v as HostResponse).body instanceof Uint8Array;

export interface GraphChildItemFx {
  id: string;
  name: string;
  folder?: { childCount?: number };
}

export interface GraphDeltaPageFx {
  value?: DriveItem[];
  '@odata.nextLink'?: string;
  '@odata.deltaLink'?: string;
}

/** Build the exact delta-walk start URL the connector requests (`token`
 *  omitted for a fresh backfill/establish walk). Tests use this to key
 *  `world.deltaPages` and to assert `calls` contents. */
export const deltaUrl = (rootId: string, token?: string): string =>
  token
    ? `${GRAPH_BASE}/me/drive/items/${rootId}/delta?token=${encodeURIComponent(token)}`
    : `${GRAPH_BASE}/me/drive/items/${rootId}/delta`;

/** The fake upstream Microsoft Graph. Every field optional; unhandled
 *  requests throw loudly (which the client treats as a network error and
 *  retries — tests run with instant sleep, so a genuinely missing fixture
 *  still fails fast). */
export interface GraphWorld {
  about?: { mail?: string; userPrincipalName?: string; id?: string };
  /** id → child-folder listing page(s) (picker `children`/root listings);
   *  page arrays paginate automatically via a synthesized `_page` param. */
  children?: Record<string, GraphChildItemFx[] | GraphChildItemFx[][]>;
  /** "Shared with me" picker tab listing. */
  sharedRoots?: GraphChildItemFx[] | GraphChildItemFx[][];
  /** id → bare item GET result — serves BOTH the picker's per-row count
   *  (`folder.childCount`) and `refreshDownloadUrl`
   *  (`@microsoft.graph.downloadUrl`) from the SAME fixture entry, since a
   *  real Graph GET without `$select` returns the full item either way. */
  items?: Record<string, (GraphChildItemFx & { '@microsoft.graph.downloadUrl'?: string }) | HostResponse>;
  /** EXACT delta-walk URL → page. Populate with `deltaUrl()` for the start
   *  URL of each root/token combination, and chain continuation pages by
   *  setting `@odata.nextLink`/`@odata.deltaLink` to further keys of this
   *  same map (mirrors how Graph's own opaque nextLink/deltaLink URLs work —
   *  and how the legacy v1 connector tests mocked `graphFetch`). */
  deltaPages?: Record<string, GraphDeltaPageFx | HostResponse>;
  /** Pre-signed downloadUrl → raw bytes (no bearer auth, no Graph host). */
  downloads?: Record<string, Uint8Array | HostResponse>;
  /** Checked first; return undefined to fall through to the world tables. */
  custom?: (url: URL, count: number) => HostResponse | undefined;
}

function paginate(url: URL, raw: GraphChildItemFx[] | GraphChildItemFx[][]): HostResponse {
  const pages: GraphChildItemFx[][] = Array.isArray(raw[0])
    ? (raw as GraphChildItemFx[][])
    : [raw as GraphChildItemFx[]];
  const idx = Number(url.searchParams.get('_page') ?? '0');
  const body: { value: GraphChildItemFx[]; '@odata.nextLink'?: string } = {
    value: pages[idx] ?? [],
  };
  if (idx + 1 < pages.length) {
    const next = new URL(url.toString());
    next.searchParams.set('_page', String(idx + 1));
    body['@odata.nextLink'] = next.toString();
  }
  return jsonRes(200, body);
}

/** The SDK kit's `scriptedFetch` with Graph's routing layered on top: the
 *  exact-URL fixtures (`deltaPages` / `downloads`) ARE the kit's own `urls`
 *  table, and the kit's `custom` hook carries the world's own hook plus the
 *  `/me`, `sharedWithMe`, `children` and item-GET endpoints — the domain
 *  knowledge that deliberately stays out of the shared kit. A URL matched by
 *  neither throws (the kit's `unhandled url`), which is what keeps a missing
 *  fixture from silently passing as an empty listing.
 *
 *  The two layers cannot collide: every path route above is a `graph.
 *  microsoft.com` path WITHOUT a `/delta` segment, while `deltaPages` keys
 *  all carry one and `downloads` keys live on another host. */
export function graphFetch(world: GraphWorld = {}): { fetchFn: NetFetch; calls: string[] } {
  const urls: Record<string, HostResponse | unknown> = {
    // Plain pages are JSON-wrapped by the kit; a HostResponse fixture (a 401
    // page, say) is passed through by it untouched — same as before.
    ...world.deltaPages,
    ...Object.fromEntries(
      Object.entries(world.downloads ?? {}).map(([url, fx]) => [
        url,
        // Raw bytes must be wrapped HERE: the kit JSON-encodes any plain
        // value, which would turn a Uint8Array body into `{"0":1,…}`.
        fx instanceof Uint8Array ? bytesRes(200, fx) : fx,
      ]),
    ),
  };

  const { fetchFn, calls } = scriptedFetch({
    urls,
    custom: (url, count) => {
      if (world.custom) {
        const r = world.custom(url, count);
        if (r) return r;
      }

      const p = url.pathname;
      if (p === '/v1.0/me') {
        return jsonRes(200, world.about ?? { mail: 'user@example.com' });
      }
      if (p === '/v1.0/me/drive/sharedWithMe') {
        return paginate(url, world.sharedRoots ?? []);
      }
      const childrenM = /^\/v1\.0\/me\/drive\/items\/([^/]+)\/children$/.exec(p);
      if (childrenM) {
        const id = childrenM[1];
        const raw = world.children?.[id];
        if (raw === undefined) throw new Error(`fake graph: no children listing for ${id}`);
        return paginate(url, raw);
      }
      const itemM = /^\/v1\.0\/me\/drive\/items\/([^/]+)$/.exec(p);
      if (itemM) {
        const id = itemM[1];
        const v = world.items?.[id];
        if (v === undefined) throw new Error(`fake graph: no item GET for ${id}`);
        return isHostResponse(v) ? v : jsonRes(200, v);
      }
      return undefined; // → the exact-URL tables above
    },
  });
  return { fetchFn, calls };
}

export function fakeQuery(docs: Document[] = [], accounts: Account[] = []): Query & {
  byExternalIdCalls: Array<{ account: string; externalId: string; type: string }>;
} {
  const byExternalIdCalls: Array<{ account: string; externalId: string; type: string }> = [];
  const unused = () => {
    throw new Error('unused Query surface');
  };
  return {
    byExternalIdCalls,
    byExternalId: async (account, externalId, type) => {
      byExternalIdCalls.push({ account: String(account), externalId, type });
      return (
        docs.find(
          (d) => d.accountId === account && d.externalId === externalId && d.type === type,
        ) ?? null
      );
    },
    document: unused,
    children: unused,
    search: unused,
    count: unused,
    countBy: unused,
    accounts: async () => accounts,
  };
}

/** A stored account as `host.query.accounts()` reports it — defaults to this
 *  connector's identity in the state the reconnect flow gates on. */
export function fakeAccount(over: Partial<Account> = {}): Account {
  return {
    id: 'acc-prior' as Account['id'],
    source: 'onedrive',
    identifier: 'ed@example.com',
    config: {},
    status: 'needsReauth',
    cursor: null,
    createdAt: '2026-01-01T00:00:00.000Z',
    ...over,
  };
}

export function makeHost(fetchFn: NetFetch, query: Query = fakeQuery()): HostFor<'net' | 'query'> {
  return {
    self: { id: 'kia.onedrive', dataDir: '/tmp' },
    log: () => {},
    net: { fetch: fetchFn },
    query,
  };
}

export function makeSession(
  opts: {
    creds?: Credentials | null;
    config?: Record<string, unknown>;
    signal?: AbortSignal;
  } = {},
): { session: Session; logs: { level: string; msg: string }[] } {
  const logs: { level: string; msg: string }[] = [];
  const base = fakeSession({
    account: {
      id: 'acc-1',
      source: 'onedrive',
      identifier: 'user@example.com',
      config: opts.config ?? {},
      status: 'live',
      cursor: null,
      createdAt: '2026-01-01T00:00:00Z',
    },
    // undefined = the default live token; an explicit null is the
    // no-credentials case the auth-error suites need.
    credentials: opts.creds === undefined ? { accessToken: 'ms-test-token-deadbeef' } : opts.creds,
    signal: opts.signal,
  });
  return {
    // Only `log` is re-pointed off the kit's session: it collects
    // [level, msg] tuples, and these suites assert on {level, msg}.
    session: {
      account: base.account,
      signal: base.signal,
      credentials: base.credentials,
      log: (level, msg) => logs.push({ level, msg }),
    },
    logs,
  };
}

export function makeAuth(
  opts: {
    creds?: Credentials;
    answers?: Record<string, unknown>;
    /** pickFolders resolves this selection (default: whole OneDrive root) —
     *  or, when a function, drives the spec itself (e.g. to reject as a user
     *  cancel). */
    picked?: FolderNode[] | ((spec: FolderPickerSpec) => Promise<FolderNode[]>);
  } = {},
): {
  auth: AuthChannel;
  statuses: string[];
  getScopes: () => string[] | undefined;
  getSchema: () => unknown;
  getPickerSpec: () => FolderPickerSpec | undefined;
} {
  let scopes: string[] | undefined;
  let schema: unknown;
  let pickerSpec: FolderPickerSpec | undefined;
  // Every interactive verb is scripted: connect() always reaches for oauth
  // and the picker, so the kit's reject-if-unscripted default would only ever
  // fire on a genuinely unexpected call.
  const auth = fakeAuthChannel({
    oauth: async (s) => {
      scopes = s;
      return opts.creds ?? { accessToken: 'ms-test-token-deadbeef' };
    },
    prompt: async (s) => {
      schema = s;
      return opts.answers ?? {};
    },
    pickFolders: async (spec) => {
      pickerSpec = spec;
      if (typeof opts.picked === 'function') return opts.picked(spec);
      return opts.picked ?? [{ id: 'root', name: 'OneDrive', hasChildren: true }];
    },
  });
  return {
    auth,
    statuses: auth.statuses,
    getScopes: () => scopes,
    getSchema: () => schema,
    getPickerSpec: () => pickerSpec,
  };
}

export function fakeDoc(
  externalId: string,
  type: string,
  metadata: Record<string, unknown>,
  over: Partial<Document> = {},
): Document {
  return {
    id: `id-${externalId}-${type}`,
    accountId: 'acc-1',
    parentId: null,
    externalId,
    type,
    title: externalId,
    markdown: '',
    metadata,
    createdAt: null,
    contentHash: 'hash-x',
    seq: 1,
    archivedAt: null,
    languages: [],
    ingestedAt: '2026-01-01T00:00:00Z',
    updatedAt: '2026-01-01T00:00:00Z',
    ...over,
  };
}

// ── Graph driveItem fixture builders ────────────────────────────────────────

export function driveFile(id: string, name: string, over: Partial<DriveItem> = {}): DriveItem {
  return {
    id,
    name,
    webUrl: `https://onedrive.live.com/?id=${id}`,
    size: 2048,
    eTag: `etag-${id}-1`,
    cTag: `ctag-${id}-1`,
    createdDateTime: '2026-04-01T10:00:00Z',
    lastModifiedDateTime: '2026-05-01T10:00:00Z',
    file: { mimeType: 'application/pdf' },
    parentReference: { driveId: 'drive-1', path: '/drive/root:' },
    '@microsoft.graph.downloadUrl': `https://download.example/${id}`,
    ...over,
  };
}

export function driveFolder(id: string, name: string, over: Partial<DriveItem> = {}): DriveItem {
  return {
    id,
    name,
    folder: { childCount: 0 },
    parentReference: { driveId: 'drive-1', path: '/drive/root:' },
    ...over,
  };
}

export function deletedItem(id: string): DriveItem {
  return { id, name: '', deleted: { state: 'deleted' }, parentReference: {} };
}

export async function collect<T>(iter: AsyncIterable<T>): Promise<T[]> {
  const out: T[] = [];
  for await (const v of iter) out.push(v);
  return out;
}
