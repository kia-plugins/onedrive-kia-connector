/**
 * Ingest/routing suite (v1 ingest.ts parity), exercised through the public
 * pull() API: hash-skip (etag unchanged), the 'failed' row exemption (and,
 * post-review, the same non-pin exemption for LEGACY pre-migration
 * 'unsupported'/'too-large' rows a policy change newly rescues), the SDK
 * decideFileIndexing eligibility gate (cloud-media/archives/unsupported
 * ignored before any content I/O, the aggregate per-pull ignore-summary
 * log), the 25 MiB size cap (declared and post-download, including the
 * post-download-oversize deletion for a pre-existing live row — this
 * connector has no reconcile pass), the downloadUrl-refresh fallback (delta
 * payloads routinely omit the annotation), and the one-shot
 * 403-refresh-and-retry.
 */
import { createOneDriveSource, GRAPH_BASE, MAX_BINARY_BYTES, type OneDriveItem } from '../source';
import { MAX_FETCH_BYTES, type Batch, type DocumentInput } from '@kiagent/connector-sdk';
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
  it('hash-skip: unchanged eTag on a live, already-successfully-fetched (extraction_status "ok") doc → no download, no item', async () => {
    const query = fakeQuery([fakeDoc('f1', 'file', { etag: 'etag-f1-1', extraction_status: 'ok' })]);
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

  it("hash-skip does NOT apply to a legacy pre-migration row (extraction_status 'unsupported'/'too-large') — a policy change that newly rescues the file must re-fetch it, not strand it behind an unchanged eTag", async () => {
    // 'unsupported' and 'too-large' are statuses the OLD (pre-SDK) code used
    // to write for a file it judged ineligible. pageChunks's pre-check only
    // ever routes an item into buildItem/hashSkip once the CURRENT policy has
    // ruled it eligible — so reaching hashSkip for a legacy non-'ok' row means
    // a policy change (e.g. an octet-stream-mime .pdf rescued by its
    // extension) just newly rescued a file that used to be ineligible.
    const query = fakeQuery([
      fakeDoc('resc1', 'file', { etag: 'etag-resc1-1', extraction_status: 'unsupported' }),
    ]);
    const { source, calls } = makeSource(
      {
        deltaPages: {
          [deltaUrl('FA')]: {
            value: [driveFile('resc1', 'rescued.pdf', { file: { mimeType: 'application/octet-stream' } })],
            '@odata.deltaLink': finalLink('FA', 'TOK1'),
          },
        },
        downloads: { 'https://download.example/resc1': new Uint8Array([9]) },
      },
      query,
    );
    const { session } = makeSession({ config: oneRoot });

    const batches = (await collect(source.pull(session, null))) as B[];

    expect(calls.some((u) => u.includes('download.example'))).toBe(true);
    expect(batches.flatMap(ids)).toEqual(['resc1']);
    expect(batches[0].items[0].extractionStatus).toBe('ok');
  });

  it("hash-skip does NOT pin a legacy 'too-large' row either — same legacy-status class as 'unsupported'", async () => {
    const query = fakeQuery([
      fakeDoc('resc2', 'file', { etag: 'etag-resc2-1', extraction_status: 'too-large' }),
    ]);
    const { source, calls } = makeSource(
      {
        deltaPages: {
          [deltaUrl('FA')]: {
            value: [driveFile('resc2', 'rescued2.pdf')],
            '@odata.deltaLink': finalLink('FA', 'TOK1'),
          },
        },
        downloads: { 'https://download.example/resc2': new Uint8Array([1]) },
      },
      query,
    );
    const { session } = makeSession({ config: oneRoot });

    const batches = (await collect(source.pull(session, null))) as B[];

    expect(calls.some((u) => u.includes('download.example'))).toBe(true);
    expect(batches.flatMap(ids)).toEqual(['resc2']);
  });

  describe('policy-ineligible files are gated out before any content I/O (SDK decideFileIndexing, cloud-drive)', () => {
    const table: Array<{ label: string; id: string; name: string; mime: string; size?: number }> = [
      { label: 'MP3 (cloud-media)', id: 'mp3-1', name: 'song.mp3', mime: 'audio/mpeg' },
      { label: 'MP4 (cloud-media)', id: 'mp4-1', name: 'clip.mp4', mime: 'video/mp4' },
      { label: 'ZIP (archive)', id: 'zip-1', name: 'bundle.zip', mime: 'application/zip' },
      { label: '26 MiB ZIP (archive, ignored regardless of size)', id: 'zip-2', name: 'huge.zip', mime: 'application/zip', size: 26 * 1024 * 1024 },
      { label: 'EXE (unsupported)', id: 'exe-1', name: 'setup.exe', mime: 'application/x-msdownload' },
      { label: 'octet-stream (unsupported)', id: 'oct-1', name: 'blob.bin', mime: 'application/octet-stream' },
    ];

    it.each(table)('$label: zero items, zero item-GET, zero download', async ({ id, name, mime, size }) => {
      const { source, calls } = makeSource({
        deltaPages: {
          [deltaUrl('FA')]: {
            value: [driveFile(id, name, { file: { mimeType: mime }, ...(size != null ? { size } : {}) })],
            '@odata.deltaLink': finalLink('FA', 'TOK1'),
          },
        },
      });
      const { session } = makeSession({ config: oneRoot });

      const batches = (await collect(source.pull(session, null))) as B[];

      expect(batches.flatMap((b) => b.items)).toEqual([]);
      expect(calls.some((u) => u.includes(`/me/drive/items/${id}`) && !u.includes('/delta'))).toBe(false);
      expect(calls.some((u) => u.includes('download.example'))).toBe(false);
    });
  });

  it('logs exactly one aggregate ignore summary per pull, with no filenames — and logs nothing when a pull ignores nothing', async () => {
    const { source: mixedSource } = makeSource({
      deltaPages: {
        [deltaUrl('FA')]: {
          value: [
            driveFile('mp3-x', 'secret-song.mp3', { file: { mimeType: 'audio/mpeg' } }),
            driveFile('zip-x', 'confidential.zip', { file: { mimeType: 'application/zip' } }),
          ],
          '@odata.deltaLink': finalLink('FA', 'TOK1'),
        },
      },
    });
    const { session: mixedSession, logs: mixedLogs } = makeSession({ config: oneRoot });

    await collect(mixedSource.pull(mixedSession, null));

    const summaryLogs = mixedLogs.filter((l) => /ignored \d+ file/.test(l.msg));
    expect(summaryLogs).toHaveLength(1);
    expect(summaryLogs[0].level).toBe('info');
    expect(summaryLogs[0].msg).toBe('onedrive: ignored 2 file(s) this pull (cloud-media=1 archive=1)');
    expect(summaryLogs[0].msg).not.toMatch(/secret-song|confidential\.zip/);

    const { source: cleanSource } = makeSource({
      deltaPages: {
        [deltaUrl('FA')]: {
          value: [driveFile('f1', 'a.pdf')],
          '@odata.deltaLink': finalLink('FA', 'TOK1'),
        },
      },
      downloads: { 'https://download.example/f1': new Uint8Array([1]) },
    });
    const { session: cleanSession, logs: cleanLogs } = makeSession({ config: oneRoot });

    await collect(cleanSource.pull(cleanSession, null));

    expect(cleanLogs.some((l) => /ignored \d+ file/.test(l.msg))).toBe(false);
  });

  const MiB = 1024 * 1024;
  const one = (id: string, name: string, over = {}) => ({
    [deltaUrl('FA')]: { value: [driveFile(id, name, over)], '@odata.deltaLink': finalLink('FA', 'TOK1') },
  });
  const DOCX = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';

  it('a 60 MiB PDF becomes a metadata-only deferred item: no item-GET, no download', async () => {
    const { source, calls } = makeSource({ deltaPages: one('big', 'big.pdf', { size: 60 * MiB }) });
    const { session } = makeSession({ config: oneRoot });
    const items = ((await collect(source.pull(session, null))) as B[]).flatMap((b) => b.items);
    expect(items).toHaveLength(1);
    expect(items[0]).toMatchObject({ markdown: '', extractionStatus: 'deferred' });
    expect(items[0].bytes).toBeUndefined();
    expect((source.toDocument(items[0]) as DocumentInput).metadata).toMatchObject({
      extraction_status: 'deferred',
      mime: 'application/pdf',
      filename: 'big.pdf',
      sizeBytes: 60 * MiB,
    });
    expect(calls.some((u) => u.includes('download.example'))).toBe(false);
    expect(calls.some((u) => u.includes('/me/drive/items/big') && !u.includes('/delta'))).toBe(false);
  });

  it('the eager cap is inclusive: size === cap downloads, cap+1 is deferred', async () => {
    const { source } = makeSource({
      deltaPages: {
        [deltaUrl('FA')]: {
          value: [
            driveFile('atcap', 'exact.pdf', { size: MAX_BINARY_BYTES }),
            driveFile('overcap', 'over.pdf', { size: MAX_BINARY_BYTES + 1 }),
          ],
          '@odata.deltaLink': finalLink('FA', 'TOK1'),
        },
      },
      downloads: { 'https://download.example/atcap': new Uint8Array([1]) },
    });
    const { session } = makeSession({ config: oneRoot });
    const items = ((await collect(source.pull(session, null))) as B[]).flatMap((b) => b.items);
    expect(items.map((i) => [i.file.id, i.extractionStatus])).toEqual([
      ['atcap', 'ok'],
      ['overcap', 'deferred'],
    ]);
  });

  it('a PDF over MAX_FETCH_BYTES and a 30 MiB docx are metadata-only "none" items', async () => {
    const { source } = makeSource({
      deltaPages: {
        [deltaUrl('FA')]: {
          value: [
            driveFile('huge', 'huge.pdf', { size: MAX_FETCH_BYTES + 1 }),
            driveFile('doc', 'big.docx', { size: 30 * MiB, file: { mimeType: DOCX } }),
          ],
          '@odata.deltaLink': finalLink('FA', 'TOK1'),
        },
      },
    });
    const { session } = makeSession({ config: oneRoot });
    const items = ((await collect(source.pull(session, null))) as B[]).flatMap((b) => b.items);
    expect(items.map((i) => [i.file.id, i.extractionStatus])).toEqual([
      ['huge', 'none'],
      ['doc', 'none'],
    ]);
  });

  it('a 30 MiB PNG is still ignored, and its pre-existing live row is archived', async () => {
    const query = fakeQuery([fakeDoc('img', 'file', { etag: 'old', extraction_status: 'ok' })]);
    const { source } = makeSource(
      { deltaPages: one('img', 'big.png', { size: 30 * MiB, file: { mimeType: 'image/png' } }) },
      query,
    );
    const { session } = makeSession({ config: oneRoot });
    const batches = (await collect(source.pull(session, null))) as B[];
    expect(batches.flatMap((b) => b.items)).toEqual([]);
    expect(batches.flatMap((b) => b.deletions ?? [])).toEqual([{ externalId: 'img', type: 'file' }]);
  });

  it('unknown declared size, 40 MiB after download → deferred item without bytes, no deletion', async () => {
    const query = fakeQuery([fakeDoc('ns', 'file', { etag: 'etag-OLD', extraction_status: 'ok' })]);
    const { source } = makeSource(
      {
        deltaPages: one('ns', 'n.pdf', { size: undefined }),
        downloads: { 'https://download.example/ns': new Uint8Array(40 * MiB) },
      },
      query,
    );
    const { session } = makeSession({ config: oneRoot });
    const batches = (await collect(source.pull(session, null))) as B[];
    const items = batches.flatMap((b) => b.items);
    expect(items.map((i) => i.extractionStatus)).toEqual(['deferred']);
    expect(items[0].bytes).toBeUndefined();
    expect(batches.flatMap((b) => b.deletions ?? [])).toEqual([]);
  });

  it('unknown declared size, 30 MiB docx after download → none, measured size persisted, fetchBytes refuses it', async () => {
    const { source, calls } = makeSource({
      deltaPages: one('nd', 'n.docx', { size: undefined, file: { mimeType: DOCX } }),
      downloads: { 'https://download.example/nd': new Uint8Array(30 * MiB) },
    });
    const { session } = makeSession({ config: oneRoot });
    const items = ((await collect(source.pull(session, null))) as B[]).flatMap((b) => b.items);
    expect(items.map((i) => i.extractionStatus)).toEqual(['none']);
    const doc = fakeDoc('nd', 'file', (source.toDocument(items[0]) as DocumentInput).metadata as Record<string, unknown>);
    expect(doc.metadata).toMatchObject({ size_bytes: 30 * MiB, sizeBytes: 30 * MiB });
    const before = calls.length;
    expect(await source.fetchBytes!(session, doc)).toBeNull();
    expect(calls.length).toBe(before);
  });

  it('a size-less file measured earlier as deferred is NOT re-downloaded on an unchanged re-listing', async () => {
    const query = fakeQuery([
      fakeDoc('ns2', 'file', { etag: 'etag-ns2-1', extraction_status: 'deferred', size_bytes: 40 * MiB }),
    ]);
    const { source, calls } = makeSource({ deltaPages: one('ns2', 'n2.pdf', { size: undefined }) }, query);
    const { session } = makeSession({ config: oneRoot });
    expect(((await collect(source.pull(session, null))) as B[]).flatMap(ids)).toEqual([]);
    expect(calls.some((u) => u.includes('download.example'))).toBe(false);
  });

  it('unknown declared size, 30 MiB PNG after download → still ignored + archived', async () => {
    const query = fakeQuery([fakeDoc('ni', 'file', { etag: 'etag-OLD', extraction_status: 'ok' })]);
    const { source } = makeSource(
      {
        deltaPages: one('ni', 'n.png', { size: undefined, file: { mimeType: 'image/png' } }),
        downloads: { 'https://download.example/ni': new Uint8Array(30 * MiB) },
      },
      query,
    );
    const { session } = makeSession({ config: oneRoot });
    const batches = (await collect(source.pull(session, null))) as B[];
    expect(batches.flatMap((b) => b.items)).toEqual([]);
    expect(batches.flatMap((b) => b.deletions ?? [])).toEqual([{ externalId: 'ni', type: 'file' }]);
  });

  it('hash-skip pins a deferred row whose eTag is unchanged (keeps the worker-converted text)', async () => {
    const query = fakeQuery([fakeDoc('big', 'file', { etag: 'etag-big-1', extraction_status: 'deferred' })]);
    const { source } = makeSource({ deltaPages: one('big', 'big.pdf', { size: 60 * MiB }) }, query);
    const { session } = makeSession({ config: oneRoot });
    expect(((await collect(source.pull(session, null))) as B[]).flatMap(ids)).toEqual([]);
  });

  it('hash-skip does NOT pin when the route changed (stored deferred, now eager after a shrink)', async () => {
    const query = fakeQuery([fakeDoc('f', 'file', { etag: 'etag-f-1', extraction_status: 'deferred' })]);
    const { source } = makeSource(
      {
        deltaPages: one('f', 'f.pdf', { size: 1000 }),
        downloads: { 'https://download.example/f': new Uint8Array([1]) },
      },
      query,
    );
    const { session } = makeSession({ config: oneRoot });
    expect(
      ((await collect(source.pull(session, null))) as B[]).flatMap((b) => b.items).map((i) => i.extractionStatus),
    ).toEqual(['ok']);
  });

  it('a .msg reported as octet-stream downloads and emits binary', async () => {
    const { source, calls } = makeSource({
      deltaPages: one('m1', 'Fwd.msg', { file: { mimeType: 'application/octet-stream' } }),
      downloads: { 'https://download.example/m1': new Uint8Array([0xd0, 0xcf, 0x11, 0xe0]) },
    });
    const { session } = makeSession({ config: oneRoot });
    const items = ((await collect(source.pull(session, null))) as B[]).flatMap((b) => b.items);
    expect(items.map((i) => [i.file.id, i.extractionStatus])).toEqual([['m1', 'ok']]);
    expect((source.toDocument(items[0]) as DocumentInput).binary?.filename).toBe('Fwd.msg');
    expect(calls).toContain('https://download.example/m1');
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
