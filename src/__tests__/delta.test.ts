/**
 * Delta suite: incremental poll of an already-established root (v1
 * delta.ts), deletions, following @odata.nextLink across multiple poll
 * pages, and the 410 resyncRequired re-prime.
 */
import { createOneDriveSource, GRAPH_BASE, type OneDriveCursor, type OneDriveItem } from '../source';
import { FILE_POLICY_VERSION, type Batch } from '@kiagent/connector-sdk';
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
    const cursor: OneDriveCursor = { delta_tokens: { FA: 'OLD_TOKEN' }, policy_version: FILE_POLICY_VERSION };

    const batches = (await collect(source.pull(session, cursor))) as B[];

    expect(calls[0]).toBe(pollUrl);
    expect(batches).toHaveLength(1);
    expect(batches[0].phase).toBe('live');
    expect(ids(batches[0])).toEqual(['f1']);
    expect(batches[0].cursor).toEqual({ delta_tokens: { FA: 'NEW_TOKEN' }, scope_roots: ['FA'], policy_version: FILE_POLICY_VERSION });
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
    const cursor: OneDriveCursor = { delta_tokens: { FA: 'OLD_TOKEN' }, policy_version: FILE_POLICY_VERSION };

    const batches = (await collect(source.pull(session, cursor))) as B[];

    expect(batches).toHaveLength(2);
    expect(batches.flatMap(ids)).toEqual(['f1', 'f2']);
    expect(batches[1].cursor).toEqual({ delta_tokens: { FA: 'NEW_TOKEN' }, scope_roots: ['FA'], policy_version: FILE_POLICY_VERSION });
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
    const cursor: OneDriveCursor = { delta_tokens: { FA: 'T' }, policy_version: FILE_POLICY_VERSION };

    const batches = (await collect(source.pull(session, cursor))) as B[];

    expect(batches[0].items).toEqual([]);
    expect(batches[0].deletions).toEqual([{ externalId: 'gone1', type: 'file' }]);
  });

  it('a live doc that turns policy-ineligible (changed to MP3) is archived via exactly one deletion — no item, no download', async () => {
    const query = fakeQuery([fakeDoc('f1', 'file', { etag: 'etag-f1-1' })]);
    const pollUrl = deltaUrl('FA', 'T');
    const { source, calls } = makeSource(
      {
        deltaPages: {
          [pollUrl]: {
            value: [driveFile('f1', 'song.mp3', { file: { mimeType: 'audio/mpeg' }, eTag: 'etag-f1-2' })],
            '@odata.deltaLink': `${GRAPH_BASE}/me/drive/items/FA/delta?token=NEW`,
          },
        },
      },
      query,
    );
    const { session } = makeSession({ config: { roots: [{ rootFolderId: 'FA', rootName: 'Alpha' }] } });
    const cursor: OneDriveCursor = { delta_tokens: { FA: 'T' }, policy_version: FILE_POLICY_VERSION };

    const batches = (await collect(source.pull(session, cursor))) as B[];

    expect(batches.flatMap(ids)).toEqual([]);
    expect(batches.flatMap((b) => b.deletions ?? [])).toEqual([{ externalId: 'f1', type: 'file' }]);
    expect(calls.some((u) => u.includes('/me/drive/items/f1') && !u.includes('/delta'))).toBe(false);
    expect(calls.some((u) => u.includes('download.example'))).toBe(false);
  });

  it('a live doc that turns policy-ineligible (changed to ZIP) is archived via exactly one deletion — no item, no download', async () => {
    const query = fakeQuery([fakeDoc('f2', 'file', { etag: 'etag-f2-1' })]);
    const pollUrl = deltaUrl('FA', 'T');
    const { source, calls } = makeSource(
      {
        deltaPages: {
          [pollUrl]: {
            value: [driveFile('f2', 'bundle.zip', { file: { mimeType: 'application/zip' }, eTag: 'etag-f2-2' })],
            '@odata.deltaLink': `${GRAPH_BASE}/me/drive/items/FA/delta?token=NEW`,
          },
        },
      },
      query,
    );
    const { session } = makeSession({ config: { roots: [{ rootFolderId: 'FA', rootName: 'Alpha' }] } });
    const cursor: OneDriveCursor = { delta_tokens: { FA: 'T' }, policy_version: FILE_POLICY_VERSION };

    const batches = (await collect(source.pull(session, cursor))) as B[];

    expect(batches.flatMap(ids)).toEqual([]);
    expect(batches.flatMap((b) => b.deletions ?? [])).toEqual([{ externalId: 'f2', type: 'file' }]);
    expect(calls.some((u) => u.includes('download.example'))).toBe(false);
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
    const cursor: OneDriveCursor = { delta_tokens: { FA: 'T' }, policy_version: FILE_POLICY_VERSION };

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
    const cursor: OneDriveCursor = { delta_tokens: { FA: 'STALE' }, policy_version: FILE_POLICY_VERSION };

    const batches = (await collect(source.pull(session, cursor))) as B[];

    expect(calls).toEqual([staleUrl, repriming]);
    expect(batches[batches.length - 1].cursor).toEqual({ delta_tokens: { FA: 'REPRIMED' }, scope_roots: ['FA'], policy_version: FILE_POLICY_VERSION });
    expect(logs.some((l) => l.level === 'warn' && /token invalid for root FA/.test(l.msg))).toBe(true);
  });

  it('a non-resync error is NOT swallowed (e.g. a plain 400)', async () => {
    const pollUrl = deltaUrl('FA', 'T');
    const { source } = makeSource({
      deltaPages: { [pollUrl]: jsonRes(400, { error: { message: 'Bad Request' } }) },
    });
    const { session } = makeSession({ config: { roots: [{ rootFolderId: 'FA', rootName: 'Alpha' }] } });
    const cursor: OneDriveCursor = { delta_tokens: { FA: 'T' }, policy_version: FILE_POLICY_VERSION };

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
    const cursor: OneDriveCursor = { delta_tokens: { FA: 'STALE', FB: 'OK_TOKEN' }, policy_version: FILE_POLICY_VERSION };

    const batches = (await collect(source.pull(session, cursor))) as B[];

    expect(batches.flatMap(ids)).toEqual(['b1']);
    const last = batches[batches.length - 1];
    expect(last.cursor).toEqual({ delta_tokens: { FA: 'REPRIMED', FB: 'NEW_B' }, scope_roots: ['FA', 'FB'], policy_version: FILE_POLICY_VERSION });
  });
});

describe('policy_version re-enumeration', () => {
  const oneRoot = { roots: [{ rootFolderId: 'FA', rootName: 'Alpha' }] };
  const finalLink = (root: string, token: string) => `${GRAPH_BASE}/me/drive/items/${root}/delta?token=${token}`;

  it('an old cursor without policy_version re-enumerates from scratch once, then polls its OLD tokens for deletions', async () => {
    const query = fakeQuery([
      fakeDoc('gone', 'file', { etag: 'e', extraction_status: 'ok' }),
      fakeDoc('kept', 'file', { etag: 'etag-kept-1', extraction_status: 'ok' }),
    ]);
    const { source, calls } = makeSource(
      {
        deltaPages: {
          // from scratch: an already-indexed file (hash-skipped) + a 60 MiB PDF the old policy ignored
          [deltaUrl('FA')]: {
            value: [driveFile('kept', 'k.pdf'), driveFile('big', 'big.pdf', { size: 60 * 1024 * 1024 })],
            '@odata.deltaLink': finalLink('FA', 'NEW'),
          },
          // old token, polled AFTER the re-walk: reports the deletion
          [deltaUrl('FA', 'OLD')]: {
            value: [deletedItem('gone')],
            '@odata.deltaLink': finalLink('FA', 'OLD2'),
          },
        },
      },
      query,
    );
    const { session } = makeSession({ config: oneRoot });
    const old: OneDriveCursor = { delta_tokens: { FA: 'OLD' }, scope_roots: ['FA'] };

    const batches = (await collect(source.pull(session, old))) as B[];

    expect(calls.filter((u) => u.includes('/delta'))).toEqual([deltaUrl('FA'), deltaUrl('FA', 'OLD')]);
    expect(batches.flatMap((b) => b.deletions ?? [])).toEqual([{ externalId: 'gone', type: 'file' }]);
    expect(batches.flatMap(ids)).toEqual(['big']);
    expect(calls.some((u) => u.includes('download.example'))).toBe(false);
    // the re-walk keeps the old token (its own NEW token is discarded)
    expect(batches[0].cursor).toEqual({ delta_tokens: { FA: 'OLD' }, scope_roots: ['FA'], policy_version: FILE_POLICY_VERSION });
    expect(batches.at(-1)!.cursor).toEqual({
      delta_tokens: { FA: 'OLD2' },
      scope_roots: ['FA'],
      policy_version: FILE_POLICY_VERSION,
    });
  });

  it('a partially backfilled old cursor re-walks every root, keeps its old token, and polls only that one', async () => {
    const twoRoots = {
      roots: [
        { rootFolderId: 'FA', rootName: 'Alpha' },
        { rootFolderId: 'FB', rootName: 'Beta' },
      ],
    };
    const query = fakeQuery([fakeDoc('gone', 'file', { etag: 'e', extraction_status: 'ok' })]);
    const { source } = makeSource(
      {
        deltaPages: {
          [deltaUrl('FA')]: { value: [], '@odata.deltaLink': finalLink('FA', 'NA') },
          [deltaUrl('FB')]: { value: [], '@odata.deltaLink': finalLink('FB', 'NB') },
          [deltaUrl('FA', 'OLD')]: { value: [deletedItem('gone')], '@odata.deltaLink': finalLink('FA', 'OLD2') },
        },
      },
      query,
    );
    const { session } = makeSession({ config: twoRoots });
    const old: OneDriveCursor = { delta_tokens: { FA: 'OLD' }, scope_roots: ['FA', 'FB'], backfill: { root_index: 1 } };
    const batches = (await collect(source.pull(session, old))) as B[];
    expect(batches.flatMap((b) => b.deletions ?? [])).toEqual([{ externalId: 'gone', type: 'file' }]);
    expect(batches.at(-1)!.cursor).toEqual({
      delta_tokens: { FA: 'OLD2', FB: 'NB' },
      scope_roots: ['FA', 'FB'],
      policy_version: FILE_POLICY_VERSION,
    });
  });

  it('a re-walk interrupted between roots still reports a deletion made while the app was closed', async () => {
    const twoRoots = {
      roots: [
        { rootFolderId: 'FA', rootName: 'Alpha' },
        { rootFolderId: 'FB', rootName: 'Beta' },
      ],
    };
    const query = fakeQuery([fakeDoc('gone-b', 'file', { etag: 'e', extraction_status: 'ok' })]);
    const { source, calls } = makeSource(
      {
        deltaPages: {
          // B, enumerated from scratch, no longer lists the deleted file
          [deltaUrl('FB')]: { value: [], '@odata.deltaLink': finalLink('FB', 'NB') },
          [deltaUrl('FA', 'OA')]: { value: [], '@odata.deltaLink': finalLink('FA', 'OA2') },
          [deltaUrl('FB', 'OB')]: { value: [deletedItem('gone-b')], '@odata.deltaLink': finalLink('FB', 'OB2') },
        },
      },
      query,
    );
    const { session } = makeSession({ config: twoRoots });
    // as committed after root A's re-walk finished; the app was then closed for days
    const mid: OneDriveCursor = {
      delta_tokens: { FA: 'OA', FB: 'OB' },
      scope_roots: ['FA', 'FB'],
      policy_version: FILE_POLICY_VERSION,
      rewalk: { root_index: 1 },
    };
    const batches = (await collect(source.pull(session, mid))) as B[];
    expect(calls.filter((u) => u.includes('/delta'))).toEqual([deltaUrl('FB'), deltaUrl('FA', 'OA'), deltaUrl('FB', 'OB')]);
    expect(batches.flatMap((b) => b.deletions ?? [])).toEqual([{ externalId: 'gone-b', type: 'file' }]);
    expect(batches.at(-1)!.cursor).toEqual({
      delta_tokens: { FA: 'OA2', FB: 'OB2' },
      scope_roots: ['FA', 'FB'],
      policy_version: FILE_POLICY_VERSION,
    });
  });

  it('a current cursor polls delta normally (no re-enumeration)', async () => {
    const { source, calls } = makeSource({
      deltaPages: { [deltaUrl('FA', 'T')]: { value: [], '@odata.deltaLink': finalLink('FA', 'T2') } },
    });
    const { session } = makeSession({ config: oneRoot });
    await collect(
      source.pull(session, { delta_tokens: { FA: 'T' }, scope_roots: ['FA'], policy_version: FILE_POLICY_VERSION }),
    );
    expect(calls.some((u) => u === deltaUrl('FA'))).toBe(false);
  });

  it('a crash mid re-enumeration resumes the re-walk page, then polls the kept token', async () => {
    const page2 = `${GRAPH_BASE}/me/drive/items/FA/delta?$skiptoken=p2`;
    // cursor as stamped by a re-walk batch
    const mid: OneDriveCursor = {
      delta_tokens: { FA: 'OLD' },
      scope_roots: ['FA'],
      policy_version: FILE_POLICY_VERSION,
      rewalk: { root_index: 0, next_link: page2 },
    };
    const { source, calls } = makeSource({
      deltaPages: {
        [page2]: { value: [], '@odata.deltaLink': finalLink('FA', 'NEW') },
        [deltaUrl('FA', 'OLD')]: { value: [], '@odata.deltaLink': finalLink('FA', 'OLD2') },
      },
    });
    const { session } = makeSession({ config: oneRoot });
    await collect(source.pull(session, mid));
    expect(calls.filter((u) => u.includes('/delta'))).toEqual([page2, deltaUrl('FA', 'OLD')]);
  });
});
