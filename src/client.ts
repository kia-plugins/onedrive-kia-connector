/**
 * v2 Microsoft Graph client — ported from alpha-cent's shared bearer-fetch
 * core (`git show main:src/main/connectors/http-shared/bearer-fetch.ts`) as
 * pinned by `ms-shared/graph-fetch.ts`'s `graphFetch` wrapper.
 *
 * Preserved verbatim from v1:
 *  - Retry policy: up to MAX_RETRIES=4 retries after the initial request
 *    (v1's MAX_ATTEMPTS loop shape), backoff `min(60000, 1000*2^attempt) +
 *    jitter*250`, `Retry-After` (seconds) honored when finite and > 0.
 *  - Retryable = 429 or >=500 ONLY (`ms-shared/graph-fetch.ts`'s `isRetryable`
 *    is status-only — MS Graph throttles via 429/503 and never uses Google's
 *    403-with-quota-reason body pattern; do not add a body regex here).
 *  - Thrown message CONTRACT: `graph <status> <url> <body>` — v1's
 *    `errorPrefix: 'graph'`, matched verbatim by the legacy delta resync
 *    check (`/410/` + `/resyncRequired|invalidToken|syncStateNotFound/i`),
 *    reproduced in this port's `isResyncRequired` (source.ts). Do not
 *    reformat.
 *  - Token fetched fresh per attempt via the `getToken` seam (in pull this is
 *    `session.credentials()` per request — the platform refreshes OAuth
 *    tokens near expiry; in connect it is the accessToken from `auth.oauth`).
 *
 * Deltas from v1 (mirrors the google-docs-kia-connector template's client.ts):
 *  1. All I/O goes through `deps.fetch` — the host's `net.fetch` surface —
 *     never global fetch. The host resolves to a plain object (status /
 *     statusText / headers with lowercase keys / body: Uint8Array), so
 *     responses are decoded manually and there is no `.ok`.
 *  2. The v1 90s per-attempt AbortController timeout is DROPPED:
 *     `host.net.fetch` owns the transport (platform-level retry/backoff and
 *     socket hygiene), so the connector no longer arms its own timers.
 *  3. HTTP 401 throws `OneDriveAuthError` (message ends "— reconnect the
 *     account"), is NEVER retried, and always propagates — the engine flips
 *     the account to needsReauth on auth errors. v1 had no 401 special-case.
 *  4. `sleep`/`random` are injectable so tests never actually wait.
 *
 * Downloading file bytes is NOT part of this client: OneDrive's pre-signed
 * `@microsoft.graph.downloadUrl` carries its own signature (no bearer header,
 * no retry loop) — v1's `ingest.ts` fetched it directly and only special-cased
 * a single 403-refresh-and-retry. That one-shot download lives in
 * `source.ts`'s `downloadBytes`, alongside this client's `GraphApiError` for
 * a uniform error shape.
 */

export type NetFetch = (url: string, init?: unknown) => Promise<unknown>;

/** Graph API root. Lives here (rather than source.ts) so tree.ts can import
 *  it without a source.ts ↔ tree.ts circular dependency. */
export const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';

/** Max retries AFTER the initial request (v1 bearer-fetch MAX_ATTEMPTS). */
const MAX_RETRIES = 4;
/** Error bodies are truncated to this many chars in thrown messages. */
const BODY_SNIPPET_CHARS = 500;

/** The host `net.fetch` surface resolves to this shape — header keys are
 *  lowercase (built via Object.fromEntries(res.headers.entries())). */
export interface HostResponse {
  status: number;
  statusText: string;
  headers: Record<string, string>;
  body: Uint8Array;
}

/** Non-2xx Graph response (except 401). Message format is load-bearing — see
 *  the module doc. */
export class GraphApiError extends Error {
  constructor(
    public readonly status: number,
    public readonly url: string,
    body: string,
  ) {
    super(`graph ${status} ${url} ${body}`);
    this.name = 'GraphApiError';
  }
}

/** HTTP 401 (or missing credentials). Never retried, always propagated —
 *  every later call would fail identically. */
export class OneDriveAuthError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'OneDriveAuthError';
  }
}

export const isAuthError = (e: unknown): e is OneDriveAuthError =>
  e instanceof OneDriveAuthError;

/** MS Graph retry predicate (`ms-shared/graph-fetch.ts`'s `isRetryable`,
 *  verbatim semantics). 429 = throttled; 5xx = transient. Other 4xx are
 *  caller errors, deliberately NOT retried. */
export function isRetryableGraphFailure(status: number): boolean {
  return status === 429 || status >= 500;
}

export interface GraphClientDeps {
  fetch: NetFetch;
  /** Fresh token per attempt — see module doc. */
  getToken: () => Promise<string>;
  sleep?: (ms: number) => Promise<void>;
  /** Jitter source for backoff (default Math.random) — injectable so retry
   *  tests can assert exact delays. */
  random?: () => number;
}

export class GraphClient {
  private readonly fetchFn: NetFetch;

  private readonly getToken: () => Promise<string>;

  private readonly sleepFn: (ms: number) => Promise<void>;

  private readonly random: () => number;

  constructor(deps: GraphClientDeps) {
    this.fetchFn = deps.fetch;
    this.getToken = deps.getToken;
    this.sleepFn =
      deps.sleep ?? ((ms) => new Promise((resolve) => setTimeout(resolve, ms)));
    this.random = deps.random ?? Math.random;
  }

  /** GET a Graph JSON endpoint with bearer auth, retry, and backoff. */
  async request<T>(url: string): Promise<T> {
    for (let attempt = 0; ; attempt++) {
      const token = await this.getToken(); // fresh per attempt
      let res: HostResponse | undefined;
      let netError: Error | undefined;
      try {
        res = (await this.fetchFn(url, {
          headers: { authorization: `Bearer ${token}`, accept: 'application/json' },
        })) as HostResponse;
      } catch (e) {
        netError = e instanceof Error ? e : new Error(String(e));
      }

      if (netError) {
        if (attempt < MAX_RETRIES) {
          await this.sleepFn(this.backoff(attempt));
          continue;
        }
        throw netError;
      }

      const r = res!;
      if (r.status >= 200 && r.status < 300) {
        const text = new TextDecoder().decode(r.body);
        return JSON.parse(text) as T;
      }

      const body = new TextDecoder().decode(r.body).slice(0, BODY_SNIPPET_CHARS);
      if (r.status === 401) {
        throw new OneDriveAuthError(
          `graph 401 ${url} ${body} — reconnect the account`,
        );
      }
      if (attempt < MAX_RETRIES && isRetryableGraphFailure(r.status)) {
        const retryAfterS = Number(r.headers['retry-after']);
        const delay =
          Number.isFinite(retryAfterS) && retryAfterS > 0
            ? retryAfterS * 1000
            : this.backoff(attempt);
        await this.sleepFn(delay);
        continue;
      }
      throw new GraphApiError(r.status, url, body);
    }
  }

  private backoff(attempt: number): number {
    return Math.min(60_000, 1000 * 2 ** attempt) + this.random() * 250;
  }
}
