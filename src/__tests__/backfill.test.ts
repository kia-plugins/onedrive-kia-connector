/**
 * Backfill suite: per-root walk over Graph's own /delta endpoint (no token)
 * to full enumeration, page-aligned batches, cursor.backfill resume, the
 * multi-root sequencing (phase 'backfill' until the LAST root's deltaLink),
 * overlapping-root dedup, abort handling, and the "feed ended without a
 * link" contract violation.
 */
import { createOneDriveSource, GRAPH_BASE, type OneDriveCursor, type OneDriveItem } from '../source';
import { OneDriveAuthError } from '../client';
import type { Batch } from '@kiagent/connector-sdk';
import {
  collect,
  deltaUrl,
  driveFile,
  driveFolder,
  fakeDoc,
  fakeQuery,
  graphFetch,
  instantClock,
  jsonRes,
  makeHost,
  makeSession,
} from '../testing/harness';
import type { Query } from '@kiagent/connector-sdk';

type B = Batch<OneDriveCursor, OneDriveItem>;
const ids = (b: B) => b.items.map((i) => i.file.id);

function makeSource(world: Parameters<typeof graphFetch>[0], query?: Query) {
  const { fetchFn, calls } = graphFetch(world);
  const source = createOneDriveSource(makeHost(fetchFn, query ?? fakeQuery()), instantClock);
  return { source, calls };
}

describe('backfill', () => {
  it('walks a single root across 2 delta pages: one batch per page, final live flip with the captured token', async () => {
    const start = deltaUrl('FA');
    const page2 = `${GRAPH_BASE}/me/drive/items/FA/delta?$skiptoken=p2`;
    const finalLink = `${GRAPH_BASE}/me/drive/items/FA/delta?token=TOK1`;
    const { source, calls } = makeSource({
      deltaPages: {
        [start]: { value: [driveFile('f1', 'a.pdf')], '@odata.nextLink': page2 },
        [page2]: { value: [driveFile('f2', 'b.pdf')], '@odata.deltaLink': finalLink },
      },
      downloads: {
        'https://download.example/f1': new Uint8Array([1, 1]),
        'https://download.example/f2': new Uint8Array([2, 2]),
      },
    });
    const { session } = makeSession({ config: { roots: [{ rootFolderId: 'FA', rootName: 'Alpha' }] } });

    const batches = (await collect(source.pull(session, null))) as B[];

    expect(batches).toHaveLength(2);
    expect(batches.map((b) => b.phase)).toEqual(['backfill', 'live']);
    expect(calls[0]).toBe(start);
    expect(batches[0].cursor).toEqual({
      delta_tokens: {},
      backfill: { root_index: 0, next_link: page2 },
    });
    expect(batches[1].cursor).toEqual({ delta_tokens: { FA: 'TOK1' } });
    expect(ids(batches[0])).toEqual(['f1']);
    expect(ids(batches[1])).toEqual(['f2']);
    expect(batches[0].items[0]).toMatchObject({
      markdown: null,
      extractionStatus: 'ok',
      displayPath: 'Alpha / a.pdf',
      rootFolderId: 'FA',
    });
    expect(batches[0].items[0].bytes).toEqual(new Uint8Array([1, 1]));
  });

  it('resumes an interrupted walk via cursor.backfill.next_link without recapturing the start URL', async () => {
    const page2 = `${GRAPH_BASE}/me/drive/items/FA/delta?$skiptoken=p2`;
    const { source, calls } = makeSource({
      deltaPages: {
        [page2]: {
          value: [driveFile('f2', 'b.pdf')],
          '@odata.deltaLink': `${GRAPH_BASE}/me/drive/items/FA/delta?token=TOK1`,
        },
      },
      downloads: { 'https://download.example/f2': new Uint8Array([2]) },
    });
    const { session } = makeSession({ config: { roots: [{ rootFolderId: 'FA', rootName: 'Alpha' }] } });

    const cursor: OneDriveCursor = { delta_tokens: {}, backfill: { root_index: 0, next_link: page2 } };
    const batches = (await collect(source.pull(session, cursor))) as B[];

    expect(calls[0]).toBe(page2);
    expect(calls.some((u) => u === deltaUrl('FA'))).toBe(false);
    expect(batches[batches.length - 1].cursor).toEqual({ delta_tokens: { FA: 'TOK1' } });
  });

  it('multi-root: phase stays backfill until the LAST root reaches its deltaLink', async () => {
    const { source } = makeSource({
      deltaPages: {
        [deltaUrl('FA')]: {
          value: [driveFile('a1', 'A doc.pdf')],
          '@odata.deltaLink': `${GRAPH_BASE}/me/drive/items/FA/delta?token=TOK_A`,
        },
        [deltaUrl('FB')]: {
          value: [driveFile('b1', 'B doc.pdf')],
          '@odata.deltaLink': `${GRAPH_BASE}/me/drive/items/FB/delta?token=TOK_B`,
        },
      },
      downloads: {
        'https://download.example/a1': new Uint8Array([1]),
        'https://download.example/b1': new Uint8Array([2]),
      },
    });
    const { session } = makeSession({
      config: {
        roots: [
          { rootFolderId: 'FA', rootName: 'Alpha' },
          { rootFolderId: 'FB', rootName: 'Beta' },
        ],
      },
    });

    const batches = (await collect(source.pull(session, null))) as B[];

    expect(batches.map((b) => b.phase)).toEqual(['backfill', 'live']);
    expect(batches[0].cursor).toEqual({ delta_tokens: { FA: 'TOK_A' }, backfill: { root_index: 1 } });
    expect(batches[1].cursor).toEqual({ delta_tokens: { FA: 'TOK_A', FB: 'TOK_B' } });
    expect(ids(batches[0])).toEqual(['a1']);
    expect(ids(batches[1])).toEqual(['b1']);
  });

  it('an item reachable from two overlapping configured roots is ingested exactly once (first root wins)', async () => {
    const shared = driveFile('shared1', 'x.pdf');
    const { source, calls } = makeSource({
      deltaPages: {
        [deltaUrl('FA')]: {
          value: [shared],
          '@odata.deltaLink': `${GRAPH_BASE}/me/drive/items/FA/delta?token=TOK_A`,
        },
        [deltaUrl('FB')]: {
          value: [shared],
          '@odata.deltaLink': `${GRAPH_BASE}/me/drive/items/FB/delta?token=TOK_B`,
        },
      },
      downloads: { 'https://download.example/shared1': new Uint8Array([9]) },
    });
    const { session } = makeSession({
      config: {
        roots: [
          { rootFolderId: 'FA', rootName: 'Alpha' },
          { rootFolderId: 'FB', rootName: 'Beta' },
        ],
      },
    });

    const batches = (await collect(source.pull(session, null))) as B[];

    expect(batches.flatMap(ids)).toEqual(['shared1']);
    expect(batches[0].items[0].rootFolderId).toBe('FA');
    expect(calls.filter((u) => u === 'https://download.example/shared1')).toHaveLength(1);
  });

  it('a policy-ineligible file (MP3) with a pre-existing live row is archived via a deletion during BACKFILL — the deletion survives from pageChunks through the yielded batch', async () => {
    const query = fakeQuery([fakeDoc('f1', 'file', { etag: 'etag-old' })]);
    const { source } = makeSource(
      {
        deltaPages: {
          [deltaUrl('FA')]: {
            value: [driveFile('f1', 'song.mp3', { file: { mimeType: 'audio/mpeg' } })],
            '@odata.deltaLink': `${GRAPH_BASE}/me/drive/items/FA/delta?token=TOK1`,
          },
        },
      },
      query,
    );
    const { session } = makeSession({ config: { roots: [{ rootFolderId: 'FA', rootName: 'Alpha' }] } });

    const batches = (await collect(source.pull(session, null))) as B[];

    expect(batches.flatMap(ids)).toEqual([]);
    expect(batches.flatMap((b) => b.deletions ?? [])).toEqual([{ externalId: 'f1', type: 'file' }]);
  });

  it('folders and the root item itself are never ingested', async () => {
    const { source } = makeSource({
      deltaPages: {
        [deltaUrl('FA')]: {
          value: [
            { id: 'FA', name: 'Alpha', folder: { childCount: 1 } }, // the root itself
            driveFolder('SUB1', 'Sub'),
            driveFile('f1', 'a.pdf'),
          ],
          '@odata.deltaLink': `${GRAPH_BASE}/me/drive/items/FA/delta?token=TOK1`,
        },
      },
      downloads: { 'https://download.example/f1': new Uint8Array([1]) },
    });
    const { session } = makeSession({ config: { roots: [{ rootFolderId: 'FA', rootName: 'Alpha' }] } });

    const batches = (await collect(source.pull(session, null))) as B[];

    expect(batches.flatMap(ids)).toEqual(['f1']);
  });

  it('stops cleanly when aborted mid-walk', async () => {
    const controller = new AbortController();
    const page2 = `${GRAPH_BASE}/me/drive/items/FA/delta?$skiptoken=p2`;
    const { source, calls } = makeSource({
      deltaPages: {
        [deltaUrl('FA')]: { value: [], '@odata.nextLink': page2 },
        [page2]: { value: [], '@odata.deltaLink': `${GRAPH_BASE}/me/drive/items/FA/delta?token=T` },
      },
    });
    const { session } = makeSession({
      config: { roots: [{ rootFolderId: 'FA', rootName: 'Alpha' }] },
      signal: controller.signal,
    });

    const iter = source.pull(session, null)[Symbol.asyncIterator]();
    const first = await iter.next();
    expect(first.done).toBe(false);
    controller.abort();
    const second = await iter.next();
    expect(second.done).toBe(true);
    expect(calls).toHaveLength(1);
  });

  it('throws when a delta feed ends without a nextLink or deltaLink (upstream contract violation)', async () => {
    const { source } = makeSource({
      deltaPages: { [deltaUrl('FA')]: { value: [] } },
    });
    const { session } = makeSession({ config: { roots: [{ rootFolderId: 'FA', rootName: 'Alpha' }] } });

    await expect(collect(source.pull(session, null))).rejects.toThrow(
      /ended without a nextLink or deltaLink/,
    );
  });

  it('propagates auth errors out of the walk (401)', async () => {
    const { source } = makeSource({
      deltaPages: { [deltaUrl('FA')]: jsonRes(401, { error: { message: 'InvalidAuthenticationToken' } }) },
    });
    const { session } = makeSession({ config: { roots: [{ rootFolderId: 'FA', rootName: 'Alpha' }] } });

    await expect(collect(source.pull(session, null))).rejects.toThrow(OneDriveAuthError);
  });

  it('propagates missing credentials as an auth error', async () => {
    const { source } = makeSource({});
    const { session } = makeSession({ creds: null, config: { roots: [{ rootFolderId: 'FA', rootName: 'A' }] } });

    await expect(collect(source.pull(session, null))).rejects.toThrow(/reconnect the account/);
  });

  it('one unreadable item is warn-skipped and the walk continues', async () => {
    const { source } = makeSource({
      deltaPages: {
        [deltaUrl('FA')]: {
          value: [driveFile('bad1', 'bad.pdf'), driveFile('f1', 'a.pdf')],
          '@odata.deltaLink': `${GRAPH_BASE}/me/drive/items/FA/delta?token=TOK1`,
        },
      },
      downloads: {
        'https://download.example/bad1': jsonRes(404, { error: { message: 'itemNotFound' } }),
        'https://download.example/f1': new Uint8Array([1]),
      },
    });
    const { session, logs } = makeSession({ config: { roots: [{ rootFolderId: 'FA', rootName: 'Alpha' }] } });

    const batches = (await collect(source.pull(session, null))) as B[];

    expect(batches.flatMap(ids)).toEqual(['f1']);
    expect(logs.some((l) => l.level === 'warn' && /item bad1 skipped/.test(l.msg))).toBe(true);
  });
});
