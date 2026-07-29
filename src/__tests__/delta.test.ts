/**
 * Delta suite: incremental poll of an already-established root (v1
 * delta.ts), deletions, following @odata.nextLink across multiple poll
 * pages, and the 410 resyncRequired re-prime.
 */
import { createOneDriveSource, GRAPH_BASE, type OneDriveCursor, type OneDriveItem } from '../source';
import type { Batch } from '@kiagent/connector-sdk';
import {
  collect,
  deletedItem,
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

type B = Batch<OneDriveCursor, OneDriveItem>;
const ids = (b: B) => b.items.map((i) => i.file.id);

function makeSource(world: Parameters<typeof graphFetch>[0], query = fakeQuery()) {
  const { fetchFn, calls } = graphFetch(world);
  const source = createOneDriveSource(makeHost(fetchFn, query), instantClock);
  return { source, calls, query };
}

describe('delta', () => {
  it('polls with the stored token and saves the new token from the deltaLink', async () => {
    const pollUrl = deltaUrl('FA', 'OLD_TOKEN');
    const { source, calls } = makeSource({
      deltaPages: {
        [pollUrl]: {
          value: [driveFile('f1', 'a.pdf')],
          '@odata.deltaLink': `${GRAPH_BASE}/me/drive/items/FA/delta?token=NEW_TOKEN`,
        },
      },
      downloads: { 'https://download.example/f1': new Uint8Array([1]) },
    });
    const { session } = makeSession({ config: { roots: [{ rootFolderId: 'FA', rootName: 'Alpha' }] } });
    const cursor: OneDriveCursor = { delta_tokens: { FA: 'OLD_TOKEN' } };

    const batches = (await collect(source.pull(session, cursor))) as B[];

    expect(calls[0]).toBe(pollUrl);
    expect(batches).toHaveLength(1);
    expect(batches[0].phase).toBe('live');
    expect(ids(batches[0])).toEqual(['f1']);
    expect(batches[0].cursor).toEqual({ delta_tokens: { FA: 'NEW_TOKEN' } });
  });

  it('follows @odata.nextLink across poll pages and saves the final token', async () => {
    const pollUrl = deltaUrl('FA', 'OLD_TOKEN');
    const page2 = `${GRAPH_BASE}/me/drive/items/FA/delta?$skiptoken=p2`;
    const { source } = makeSource({
      deltaPages: {
        [pollUrl]: { value: [driveFile('f1', 'a.pdf')], '@odata.nextLink': page2 },
        [page2]: {
          value: [driveFile('f2', 'b.pdf')],
          '@odata.deltaLink': `${GRAPH_BASE}/me/drive/items/FA/delta?token=NEW_TOKEN`,
        },
      },
      downloads: {
        'https://download.example/f1': new Uint8Array([1]),
        'https://download.example/f2': new Uint8Array([2]),
      },
    });
    const { session } = makeSession({ config: { roots: [{ rootFolderId: 'FA', rootName: 'Alpha' }] } });
    const cursor: OneDriveCursor = { delta_tokens: { FA: 'OLD_TOKEN' } };

    const batches = (await collect(source.pull(session, cursor))) as B[];

    expect(batches).toHaveLength(2);
    expect(batches.flatMap(ids)).toEqual(['f1', 'f2']);
    expect(batches[1].cursor).toEqual({ delta_tokens: { FA: 'NEW_TOKEN' } });
  });

  it('deletes: a removed item with an existing doc is surfaced via Batch.deletions', async () => {
    const query = fakeQuery([fakeDoc('gone1', 'file', { etag: 'e-old' })]);
    const pollUrl = deltaUrl('FA', 'T');
    const { source } = makeSource(
      {
        deltaPages: {
          [pollUrl]: {
            value: [deletedItem('gone1')],
            '@odata.deltaLink': `${GRAPH_BASE}/me/drive/items/FA/delta?token=NEW`,
          },
        },
      },
      query,
    );
    const { session } = makeSession({ config: { roots: [{ rootFolderId: 'FA', rootName: 'Alpha' }] } });
    const cursor: OneDriveCursor = { delta_tokens: { FA: 'T' } };

    const batches = (await collect(source.pull(session, cursor))) as B[];

    expect(batches[0].items).toEqual([]);
    expect(batches[0].deletions).toEqual([{ externalId: 'gone1', type: 'file' }]);
  });

  it('a removed item with NO existing doc emits no deletion ref', async () => {
    const pollUrl = deltaUrl('FA', 'T');
    const { source } = makeSource({
      deltaPages: {
        [pollUrl]: {
          value: [deletedItem('never-existed')],
          '@odata.deltaLink': `${GRAPH_BASE}/me/drive/items/FA/delta?token=NEW`,
        },
      },
    });
    const { session } = makeSession({ config: { roots: [{ rootFolderId: 'FA', rootName: 'Alpha' }] } });
    const cursor: OneDriveCursor = { delta_tokens: { FA: 'T' } };

    const batches = (await collect(source.pull(session, cursor))) as B[];

    expect(batches[0].deletions).toEqual([]);
  });

  it('does NOT surface deletions during backfill (only delta polls)', async () => {
    const { source } = makeSource({
      deltaPages: {
        [deltaUrl('FA')]: {
          value: [deletedItem('gone1')],
          '@odata.deltaLink': `${GRAPH_BASE}/me/drive/items/FA/delta?token=TOK1`,
        },
      },
    });
    const { session } = makeSession({ config: { roots: [{ rootFolderId: 'FA', rootName: 'Alpha' }] } });

    const batches = (await collect(source.pull(session, null))) as B[];

    expect(batches[0].deletions ?? []).toEqual([]);
  });

  it('on 410 resyncRequired, re-primes that root with a fresh full enumerate', async () => {
    const staleUrl = deltaUrl('FA', 'STALE');
    const repriming = deltaUrl('FA');
    const { source, calls } = makeSource({
      deltaPages: {
        [staleUrl]: jsonRes(410, { error: { code: 'resyncRequired' } }),
        [repriming]: {
          value: [],
          '@odata.deltaLink': `${GRAPH_BASE}/me/drive/items/FA/delta?token=REPRIMED`,
        },
      },
    });
    const { session, logs } = makeSession({ config: { roots: [{ rootFolderId: 'FA', rootName: 'Alpha' }] } });
    const cursor: OneDriveCursor = { delta_tokens: { FA: 'STALE' } };

    const batches = (await collect(source.pull(session, cursor))) as B[];

    expect(calls).toEqual([staleUrl, repriming]);
    expect(batches[batches.length - 1].cursor).toEqual({ delta_tokens: { FA: 'REPRIMED' } });
    expect(logs.some((l) => l.level === 'warn' && /token invalid for root FA/.test(l.msg))).toBe(true);
  });

  it('a non-resync error is NOT swallowed (e.g. a plain 400)', async () => {
    const pollUrl = deltaUrl('FA', 'T');
    const { source } = makeSource({
      deltaPages: { [pollUrl]: jsonRes(400, { error: { message: 'Bad Request' } }) },
    });
    const { session } = makeSession({ config: { roots: [{ rootFolderId: 'FA', rootName: 'Alpha' }] } });
    const cursor: OneDriveCursor = { delta_tokens: { FA: 'T' } };

    await expect(collect(source.pull(session, cursor))).rejects.toThrow(/^graph 400 /);
  });

  it('multiple roots: one root resyncing does not block another root from polling normally', async () => {
    const staleUrl = deltaUrl('FA', 'STALE');
    const repriming = deltaUrl('FA');
    const pollB = deltaUrl('FB', 'OK_TOKEN');
    const { source } = makeSource({
      deltaPages: {
        [staleUrl]: jsonRes(410, { error: { code: 'resyncRequired' } }),
        [repriming]: {
          value: [],
          '@odata.deltaLink': `${GRAPH_BASE}/me/drive/items/FA/delta?token=REPRIMED`,
        },
        [pollB]: {
          value: [driveFile('b1', 'b.pdf')],
          '@odata.deltaLink': `${GRAPH_BASE}/me/drive/items/FB/delta?token=NEW_B`,
        },
      },
      downloads: { 'https://download.example/b1': new Uint8Array([1]) },
    });
    const { session } = makeSession({
      config: {
        roots: [
          { rootFolderId: 'FA', rootName: 'Alpha' },
          { rootFolderId: 'FB', rootName: 'Beta' },
        ],
      },
    });
    const cursor: OneDriveCursor = { delta_tokens: { FA: 'STALE', FB: 'OK_TOKEN' } };

    const batches = (await collect(source.pull(session, cursor))) as B[];

    expect(batches.flatMap(ids)).toEqual(['b1']);
    const last = batches[batches.length - 1];
    expect(last.cursor).toEqual({ delta_tokens: { FA: 'REPRIMED', FB: 'NEW_B' } });
  });
});
