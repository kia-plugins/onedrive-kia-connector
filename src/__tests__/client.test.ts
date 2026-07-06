/**
 * Retry-matrix suite for the v2 GraphClient — the v1 contracts (bearer-fetch
 * retry predicate, backoff shape, Retry-After, error-message format) plus the
 * v2 additions (401 → OneDriveAuthError, host-shaped responses). Unlike
 * Drive, Graph's retry predicate is status-only (no quota-body regex).
 */
import { GraphApiError, GraphClient, OneDriveAuthError, type NetFetch } from '../client';
import { jsonRes, type HostResponse } from '../testing/harness';

const URL_X = 'https://graph.microsoft.com/v1.0/me/drive/items/x';

function scripted(responses: Array<HostResponse | Error>): {
  fetchFn: NetFetch;
  calls: string[];
} {
  const calls: string[] = [];
  let i = 0;
  const fetchFn: NetFetch = async (url) => {
    calls.push(String(url));
    const r = responses[i];
    i += 1;
    if (r === undefined) throw new Error(`scripted: no response for call #${i}`);
    if (r instanceof Error) throw r;
    return r;
  };
  return { fetchFn, calls };
}

function makeClient(fetchFn: NetFetch, tokens?: string[]) {
  const sleeps: number[] = [];
  const tokenCalls: number[] = [];
  let n = 0;
  const client = new GraphClient({
    fetch: fetchFn,
    getToken: async () => {
      tokenCalls.push(n);
      const t = tokens?.[n] ?? 'ms-test-token-deadbeef';
      n += 1;
      return t;
    },
    sleep: async (ms) => {
      sleeps.push(ms);
    },
    random: () => 0,
  });
  return { client, sleeps, tokenCalls };
}

describe('GraphClient retry matrix', () => {
  it('retries 429 with exponential backoff and succeeds', async () => {
    const { fetchFn, calls } = scripted([
      jsonRes(429, {}),
      jsonRes(429, {}),
      jsonRes(200, { id: 'x' }),
    ]);
    const { client, sleeps } = makeClient(fetchFn);
    await expect(client.request(URL_X)).resolves.toEqual({ id: 'x' });
    expect(calls).toHaveLength(3);
    expect(sleeps).toEqual([1000, 2000]); // min(60000, 1000*2^attempt) + 0 jitter
  });

  it('retries 5xx', async () => {
    const { fetchFn, calls } = scripted([jsonRes(503, {}), jsonRes(200, { ok: true })]);
    const { client } = makeClient(fetchFn);
    await expect(client.request(URL_X)).resolves.toEqual({ ok: true });
    expect(calls).toHaveLength(2);
  });

  it('does NOT retry a non-throttle 403 (Graph has no quota-body regex, unlike Google)', async () => {
    const { fetchFn, calls } = scripted([
      jsonRes(403, { error: { message: 'Access denied' } }),
    ]);
    const { client } = makeClient(fetchFn);
    await expect(client.request(URL_X)).rejects.toThrow(GraphApiError);
    expect(calls).toHaveLength(1);
  });

  it('does NOT retry 400 and throws the "graph <status> <url> <body>" format', async () => {
    const { fetchFn, calls } = scripted([jsonRes(400, { error: { message: 'Bad Request' } })]);
    const { client } = makeClient(fetchFn);
    const err = (await client.request(URL_X).catch((e) => e)) as Error;
    expect(err).toBeInstanceOf(GraphApiError);
    expect(err.message).toMatch(new RegExp(`^graph 400 ${URL_X.replace(/[?.]/g, '\\$&')} `));
    expect(err.message).toContain('Bad Request');
    expect(calls).toHaveLength(1);
  });

  it('401 → OneDriveAuthError, never retried, message ends "— reconnect the account"', async () => {
    const { fetchFn, calls } = scripted([
      jsonRes(401, { error: { message: 'InvalidAuthenticationToken' } }),
      jsonRes(200, { never: 'reached' }),
    ]);
    const { client, sleeps } = makeClient(fetchFn);
    const err = (await client.request(URL_X).catch((e) => e)) as Error;
    expect(err).toBeInstanceOf(OneDriveAuthError);
    expect(err.message).toMatch(/— reconnect the account$/);
    expect(calls).toHaveLength(1);
    expect(sleeps).toEqual([]);
  });

  it('honors Retry-After (seconds) over the backoff curve', async () => {
    const { fetchFn } = scripted([
      jsonRes(429, {}, { 'retry-after': '7' }),
      jsonRes(200, { ok: true }),
    ]);
    const { client, sleeps } = makeClient(fetchFn);
    await client.request(URL_X);
    expect(sleeps).toEqual([7000]);
  });

  it('retries network errors, then rethrows the last one after 4 retries (5 requests)', async () => {
    const boom = new Error('socket hang up');
    const { fetchFn, calls } = scripted([boom, boom, boom, boom, boom]);
    const { client, sleeps } = makeClient(fetchFn);
    await expect(client.request(URL_X)).rejects.toThrow('socket hang up');
    expect(calls).toHaveLength(5);
    expect(sleeps).toEqual([1000, 2000, 4000, 8000]);
  });

  it('gives up on persistent 429 after 4 retries (5 requests)', async () => {
    const r = jsonRes(429, {});
    const { fetchFn, calls } = scripted([r, r, r, r, r]);
    const { client } = makeClient(fetchFn);
    await expect(client.request(URL_X)).rejects.toThrow(/^graph 429 /);
    expect(calls).toHaveLength(5);
  });

  it('recovers mid-sequence: network error then 500 then 200', async () => {
    const { fetchFn, calls } = scripted([
      new Error('ECONNRESET'),
      jsonRes(500, {}),
      jsonRes(200, { done: true }),
    ]);
    const { client } = makeClient(fetchFn);
    await expect(client.request(URL_X)).resolves.toEqual({ done: true });
    expect(calls).toHaveLength(3);
  });

  it('fetches a fresh token per attempt (expiry-safe retries)', async () => {
    const { fetchFn } = scripted([jsonRes(500, {}), jsonRes(200, {})]);
    const { client, tokenCalls } = makeClient(fetchFn, ['ms-test-old', 'ms-test-new']);
    await client.request(URL_X);
    expect(tokenCalls).toHaveLength(2);
  });
});
