/**
 * Ingest/routing suite (v1 ingest.ts parity), exercised through the public
 * pull() API: hash-skip (etag unchanged), the 'failed' row exemption,
 * unsupported-mime metadata-only routing, the 25 MiB size cap (declared and
 * post-download), the downloadUrl-refresh fallback (delta payloads routinely
 * omit the annotation), and the one-shot 403-refresh-and-retry.
 */
import { createOneDriveSource, GRAPH_BASE, MAX_BINARY_BYTES, type OneDriveItem } from '../source';
import type { Batch } from '../kiagent-contracts';
import {
  collect,
  deltaUrl,
  driveFile,
  fakeDoc,
  fakeQuery,
  graphFetch,
  instantClock,
  jsonRes,
  makeHost,
  makeSession,
} from '../testing/harness';

type B = Batch<unknown, OneDriveItem>;
const ids = (b: B) => b.items.map((i) => i.file.id);

function makeSource(world: Parameters<typeof graphFetch>[0], query = fakeQuery()) {
  const { fetchFn, calls } = graphFetch(world);
  const source = createOneDriveSource(makeHost(fetchFn, query), instantClock);
  return { source, calls, query };
}

const oneRoot = { roots: [{ rootFolderId: 'FA', rootName: 'Alpha' }] };
const finalLink = (root: string, token: string) => `${GRAPH_BASE}/me/drive/items/${root}/delta?token=${token}`;

describe('ingest routing', () => {
  it('hash-skip: unchanged eTag on a live doc → no download, no item', async () => {
    const query = fakeQuery([fakeDoc('f1', 'file', { etag: 'etag-f1-1' })]);
    const { source, calls } = makeSource(
      {
        deltaPages: {
          [deltaUrl('FA')]: {
            value: [driveFile('f1', 'a.pdf')],
            '@odata.deltaLink': finalLink('FA', 'TOK1'),
          },
        },
      },
      query,
    );
    const { session } = makeSession({ config: oneRoot });

    const batches = (await collect(source.pull(session, null))) as B[];

    expect(batches.flatMap(ids)).toEqual([]);
    expect(calls.some((u) => u.includes('download.example'))).toBe(false);
    expect(query.byExternalIdCalls).toContainEqual({ account: 'acc-1', externalId: 'f1', type: 'file' });
  });

  it('hash-skip does NOT apply to an archived match (re-emitting un-archives it)', async () => {
    const query = fakeQuery([fakeDoc('f1', 'file', { etag: 'etag-f1-1' }, { archivedAt: '2026-06-01T00:00:00Z' })]);
    const { source } = makeSource(
      {
        deltaPages: {
          [deltaUrl('FA')]: {
            value: [driveFile('f1', 'a.pdf')],
            '@odata.deltaLink': finalLink('FA', 'TOK1'),
          },
        },
        downloads: { 'https://download.example/f1': new Uint8Array([1]) },
      },
      query,
    );
    const { session } = makeSession({ config: oneRoot });

    const batches = (await collect(source.pull(session, null))) as B[];

    expect(batches.flatMap(ids)).toEqual(['f1']);
  });

  it("hash-skip does NOT apply to a 'failed' extraction row — download is re-attempted", async () => {
    const query = fakeQuery([fakeDoc('f1', 'file', { etag: 'etag-f1-1', extraction_status: 'failed' })]);
    const { source, calls } = makeSource(
      {
        deltaPages: {
          [deltaUrl('FA')]: {
            value: [driveFile('f1', 'a.pdf')],
            '@odata.deltaLink': finalLink('FA', 'TOK1'),
          },
        },
        downloads: { 'https://download.example/f1': new Uint8Array([1]) },
      },
      query,
    );
    const { session } = makeSession({ config: oneRoot });

    const batches = (await collect(source.pull(session, null))) as B[];

    expect(calls.some((u) => u.includes('download.example'))).toBe(true);
    expect(batches.flatMap(ids)).toEqual(['f1']);
    expect(batches[0].items[0].extractionStatus).toBe('ok');
  });

  it('routes an unsupported mime to a metadata-only row (no download)', async () => {
    const { source, calls } = makeSource({
      deltaPages: {
        [deltaUrl('FA')]: {
          value: [driveFile('z1', 'archive.zip', { file: { mimeType: 'application/zip' } })],
          '@odata.deltaLink': finalLink('FA', 'TOK1'),
        },
      },
    });
    const { session } = makeSession({ config: oneRoot });

    const batches = (await collect(source.pull(session, null))) as B[];

    expect(batches[0].items[0]).toMatchObject({ markdown: '', extractionStatus: 'unsupported' });
    expect(batches[0].items[0].bytes).toBeUndefined();
    expect(calls.some((u) => u.includes('download.example'))).toBe(false);
  });

  it('skips the download for a declared-too-large file (size field over the cap)', async () => {
    const { source, calls } = makeSource({
      deltaPages: {
        [deltaUrl('FA')]: {
          value: [driveFile('big1', 'huge.pdf', { size: MAX_BINARY_BYTES + 1 })],
          '@odata.deltaLink': finalLink('FA', 'TOK1'),
        },
      },
    });
    const { session } = makeSession({ config: oneRoot });

    const batches = (await collect(source.pull(session, null))) as B[];

    expect(batches[0].items[0]).toMatchObject({ markdown: '', extractionStatus: 'too-large' });
    expect(calls.some((u) => u.includes('download.example'))).toBe(false);
  });

  it('caps an unknown-size file AFTER download (post-hoc too-large, no bytes emitted)', async () => {
    const { source } = makeSource({
      deltaPages: {
        [deltaUrl('FA')]: {
          value: [driveFile('nosize1', 'n.pdf', { size: undefined })],
          '@odata.deltaLink': finalLink('FA', 'TOK1'),
        },
      },
      downloads: { 'https://download.example/nosize1': new Uint8Array(MAX_BINARY_BYTES + 1) },
    });
    const { session } = makeSession({ config: oneRoot });

    const batches = (await collect(source.pull(session, null))) as B[];

    expect(batches[0].items[0]).toMatchObject({ markdown: '', extractionStatus: 'too-large' });
    expect(batches[0].items[0].bytes).toBeUndefined();
  });

  it('refreshes the downloadUrl when the delta payload omits it, then downloads successfully', async () => {
    const { source } = makeSource({
      deltaPages: {
        [deltaUrl('FA')]: {
          value: [driveFile('f1', 'a.pdf', { '@microsoft.graph.downloadUrl': undefined })],
          '@odata.deltaLink': finalLink('FA', 'TOK1'),
        },
      },
      items: { f1: { id: 'f1', name: 'a.pdf', '@microsoft.graph.downloadUrl': 'https://download.example/f1-fresh' } },
      downloads: { 'https://download.example/f1-fresh': new Uint8Array([7]) },
    });
    const { session } = makeSession({ config: oneRoot });

    const batches = (await collect(source.pull(session, null))) as B[];

    expect(batches[0].items[0].extractionStatus).toBe('ok');
    expect(batches[0].items[0].bytes).toEqual(new Uint8Array([7]));
  });

  it("marks extraction_status='failed' when the refreshed downloadUrl is also missing", async () => {
    const { source } = makeSource({
      deltaPages: {
        [deltaUrl('FA')]: {
          value: [driveFile('f1', 'a.pdf', { '@microsoft.graph.downloadUrl': undefined })],
          '@odata.deltaLink': finalLink('FA', 'TOK1'),
        },
      },
      items: { f1: { id: 'f1', name: 'a.pdf' } },
    });
    const { session } = makeSession({ config: oneRoot });

    const batches = (await collect(source.pull(session, null))) as B[];

    expect(batches[0].items[0]).toMatchObject({ markdown: '', extractionStatus: 'failed' });
  });

  it('retries once on a pre-signed 403 by re-fetching the downloadUrl', async () => {
    let calls403 = 0;
    const { source } = makeSource({
      deltaPages: {
        [deltaUrl('FA')]: {
          value: [driveFile('f1', 'a.pdf')],
          '@odata.deltaLink': finalLink('FA', 'TOK1'),
        },
      },
      items: { f1: { id: 'f1', name: 'a.pdf', '@microsoft.graph.downloadUrl': 'https://download.example/f1-fresh' } },
      custom: (url) => {
        if (url.toString() === 'https://download.example/f1') {
          calls403++;
          return jsonRes(403, { error: 'expired' });
        }
        return undefined;
      },
      downloads: { 'https://download.example/f1-fresh': new Uint8Array([9]) },
    });
    const { session } = makeSession({ config: oneRoot });

    const batches = (await collect(source.pull(session, null))) as B[];

    expect(calls403).toBe(1);
    expect(batches[0].items[0]).toMatchObject({ extractionStatus: 'ok' });
    expect(batches[0].items[0].bytes).toEqual(new Uint8Array([9]));
  });

  it('a 403 that persists even after refresh propagates', async () => {
    const { source } = makeSource({
      deltaPages: {
        [deltaUrl('FA')]: {
          value: [driveFile('f1', 'a.pdf')],
          '@odata.deltaLink': finalLink('FA', 'TOK1'),
        },
      },
      items: { f1: { id: 'f1', name: 'a.pdf', '@microsoft.graph.downloadUrl': 'https://download.example/f1-fresh' } },
      downloads: {
        'https://download.example/f1': jsonRes(403, { error: 'expired' }),
        'https://download.example/f1-fresh': jsonRes(403, { error: 'still expired' }),
      },
    });
    const { session, logs } = makeSession({ config: oneRoot });

    const batches = (await collect(source.pull(session, null))) as B[];

    // Per-item fault tolerance (v1 parity): the 403 is caught at the item
    // level and warn-logged, not thrown out of pull().
    expect(batches.flatMap(ids)).toEqual([]);
    expect(logs.some((l) => l.level === 'warn' && /item f1 skipped/.test(l.msg))).toBe(true);
  });
});
