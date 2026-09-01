/**
 * Batch byte-budget suite: a Graph delta page is flushed to the engine in
 * sub-page chunks once the accumulated downloaded bytes (or item count)
 * cross a budget, instead of holding every file of the page in memory at
 * once. Intermediate chunks repeat the cursor that fetched the page; only the
 * page's final chunk advances it (nextLink / delta token) — so a crash
 * mid-page replays that page (idempotent via hash-skip) and never skips it.
 */
import {
  BATCH_BYTE_BUDGET,
  BATCH_ITEM_LIMIT,
  createOneDriveSource,
  GRAPH_BASE,
  MAX_BINARY_BYTES,
  type OneDriveCursor,
  type OneDriveItem,
} from '../source';
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
  makeHost,
  makeSession,
} from '../testing/harness';

type B = Batch<OneDriveCursor, OneDriveItem>;
const ids = (b: B) => b.items.map((i) => i.file.id);

function makeSource(
  world: Parameters<typeof graphFetch>[0],
  seams: { batchByteBudget?: number; batchItemLimit?: number } = {},
  query = fakeQuery(),
) {
  const { fetchFn, calls } = graphFetch(world);
  const source = createOneDriveSource(makeHost(fetchFn, query), { ...instantClock, ...seams });
  return { source, calls, query };
}

const bytes = (n: number, fill = 1) => new Uint8Array(n).fill(fill);

describe('batch byte budget', () => {
  it('defaults: budget is a fraction of the process, well below a page of max-size files', () => {
    expect(BATCH_BYTE_BUDGET).toBeGreaterThanOrEqual(MAX_BINARY_BYTES);
    expect(BATCH_BYTE_BUDGET).toBeLessThanOrEqual(64 * 1024 * 1024);
    expect(BATCH_ITEM_LIMIT).toBeGreaterThan(0);
  });

  it('delta: flushes mid-page once bytes cross the budget; only the final chunk carries the new token', async () => {
    const pollUrl = deltaUrl('FA', 'OLD');
    const { source } = makeSource(
      {
        deltaPages: {
          [pollUrl]: {
            value: [driveFile('f1', 'a.pdf'), driveFile('f2', 'b.pdf'), driveFile('f3', 'c.pdf')],
            '@odata.deltaLink': `${GRAPH_BASE}/me/drive/items/FA/delta?token=NEW`,
          },
        },
        downloads: {
          'https://download.example/f1': bytes(3),
          'https://download.example/f2': bytes(3),
          'https://download.example/f3': bytes(3),
        },
      },
      { batchByteBudget: 5 },
    );
    const { session } = makeSession({ config: { roots: [{ rootFolderId: 'FA', rootName: 'Alpha' }] } });

    const batches = (await collect(source.pull(session, { delta_tokens: { FA: 'OLD' } }))) as B[];

    expect(batches.map(ids)).toEqual([['f1', 'f2'], ['f3']]);
    expect(batches.map((b) => b.phase)).toEqual(['live', 'live']);
    // Intermediate chunk: cursor that fetched the page (old token) — a crash
    // after this commit replays the page, never skips it.
    expect(batches[0].cursor).toEqual({ delta_tokens: { FA: 'OLD' } });
    expect(batches[1].cursor).toEqual({ delta_tokens: { FA: 'NEW' } });
  });

  it('delta: a budget hit on the last item still advances the token via a trailing empty chunk', async () => {
    const pollUrl = deltaUrl('FA', 'OLD');
    const { source } = makeSource(
      {
        deltaPages: {
          [pollUrl]: {
            value: [driveFile('f1', 'a.pdf'), driveFile('f2', 'b.pdf')],
            '@odata.deltaLink': `${GRAPH_BASE}/me/drive/items/FA/delta?token=NEW`,
          },
        },
        downloads: {
          'https://download.example/f1': bytes(3),
          'https://download.example/f2': bytes(3),
        },
      },
      { batchByteBudget: 5 },
    );
    const { session } = makeSession({ config: { roots: [{ rootFolderId: 'FA', rootName: 'Alpha' }] } });

    const batches = (await collect(source.pull(session, { delta_tokens: { FA: 'OLD' } }))) as B[];

    expect(batches.map(ids)).toEqual([['f1', 'f2'], []]);
    expect(batches[0].cursor).toEqual({ delta_tokens: { FA: 'OLD' } });
    expect(batches[1].cursor).toEqual({ delta_tokens: { FA: 'NEW' } });
  });

  it('delta: the item limit flushes metadata-only items too, and each deletion lands in exactly one chunk', async () => {
    const pollUrl = deltaUrl('FA', 'OLD');
    const unsupported = { file: { mimeType: 'application/x-msdownload' } };
    const { source } = makeSource(
      {
        deltaPages: {
          [pollUrl]: {
            value: [
              deletedItem('gone1'),
              driveFile('f1', 'a.exe', unsupported),
              driveFile('f2', 'b.exe', unsupported),
              driveFile('f3', 'c.exe', unsupported),
              deletedItem('gone2'),
            ],
            '@odata.deltaLink': `${GRAPH_BASE}/me/drive/items/FA/delta?token=NEW`,
          },
        },
      },
      { batchItemLimit: 2 },
      fakeQuery([fakeDoc('gone1', 'file', {}), fakeDoc('gone2', 'file', {})]),
    );
    const { session } = makeSession({ config: { roots: [{ rootFolderId: 'FA', rootName: 'Alpha' }] } });

    const batches = (await collect(source.pull(session, { delta_tokens: { FA: 'OLD' } }))) as B[];

    // Deletions count against the limit like items: [gone1,f1] [f2,f3] [gone2]
    expect(batches.map((b) => [...(b.deletions ?? []).map((d) => d.externalId), ...ids(b)])).toEqual([
      ['gone1', 'f1'],
      ['f2', 'f3'],
      ['gone2'],
    ]);
    expect(batches.slice(0, -1).every((b) => b.cursor.delta_tokens.FA === 'OLD')).toBe(true);
    expect(batches.at(-1)!.cursor).toEqual({ delta_tokens: { FA: 'NEW' } });
  });

  it('delta: a deletion reported by two overlapping roots is surfaced once per pull', async () => {
    const urlA = deltaUrl('FA', 'TA');
    const urlB = deltaUrl('FB', 'TB');
    const { source } = makeSource(
      {
        deltaPages: {
          [urlA]: {
            value: [deletedItem('gone1')],
            '@odata.deltaLink': `${GRAPH_BASE}/me/drive/items/FA/delta?token=TA2`,
          },
          [urlB]: {
            value: [deletedItem('gone1')],
            '@odata.deltaLink': `${GRAPH_BASE}/me/drive/items/FB/delta?token=TB2`,
          },
        },
      },
      {},
      fakeQuery([fakeDoc('gone1', 'file', {})]),
    );
    const { session } = makeSession({
      config: {
        roots: [
          { rootFolderId: 'FA', rootName: 'Alpha' },
          { rootFolderId: 'FB', rootName: 'Beta' },
        ],
      },
    });

    const batches = (await collect(source.pull(session, { delta_tokens: { FA: 'TA', FB: 'TB' } }))) as B[];

    expect(batches.flatMap((b) => (b.deletions ?? []).map((d) => d.externalId))).toEqual(['gone1']);
    expect(batches.at(-1)!.cursor).toEqual({ delta_tokens: { FA: 'TA2', FB: 'TB2' } });
  });

  it('backfill: intermediate chunks repeat the cursor that fetched the page; the final chunk stores the nextLink', async () => {
    const start = deltaUrl('FA');
    const page2 = `${GRAPH_BASE}/me/drive/items/FA/delta?$skiptoken=p2`;
    const { source } = makeSource(
      {
        deltaPages: {
          [start]: { value: [driveFile('f1', 'a.pdf'), driveFile('f2', 'b.pdf')], '@odata.nextLink': page2 },
          [page2]: {
            value: [driveFile('f3', 'c.pdf'), driveFile('f4', 'd.pdf'), driveFile('f5', 'e.pdf')],
            '@odata.deltaLink': `${GRAPH_BASE}/me/drive/items/FA/delta?token=TOK1`,
          },
        },
        downloads: {
          'https://download.example/f1': bytes(3),
          'https://download.example/f2': bytes(3),
          'https://download.example/f3': bytes(3),
          'https://download.example/f4': bytes(3),
          'https://download.example/f5': bytes(3),
        },
      },
      { batchByteBudget: 5 },
    );
    const { session } = makeSession({ config: { roots: [{ rootFolderId: 'FA', rootName: 'Alpha' }] } });

    const batches = (await collect(source.pull(session, null))) as B[];

    expect(batches.map(ids)).toEqual([['f1', 'f2'], [], ['f3', 'f4'], ['f5']]);
    expect(batches.map((b) => b.phase)).toEqual(['backfill', 'backfill', 'backfill', 'live']);
    // Page 1 fetched from the start URL → its intermediate chunk has NO next_link.
    expect(batches[0].cursor).toEqual({ delta_tokens: {}, backfill: { root_index: 0 } });
    expect(batches[1].cursor).toEqual({ delta_tokens: {}, backfill: { root_index: 0, next_link: page2 } });
    // Page 2's intermediate chunk repeats page 2's own link — a replay refetches page 2.
    expect(batches[2].cursor).toEqual({ delta_tokens: {}, backfill: { root_index: 0, next_link: page2 } });
    expect(batches[3].cursor).toEqual({ delta_tokens: { FA: 'TOK1' } });
  });

  it('backfill resume: intermediate chunks of a resumed page repeat the resume link', async () => {
    const page2 = `${GRAPH_BASE}/me/drive/items/FA/delta?$skiptoken=p2`;
    const { source } = makeSource(
      {
        deltaPages: {
          [page2]: {
            value: [driveFile('f3', 'c.pdf'), driveFile('f4', 'd.pdf'), driveFile('f5', 'e.pdf')],
            '@odata.deltaLink': `${GRAPH_BASE}/me/drive/items/FA/delta?token=TOK1`,
          },
        },
        downloads: {
          'https://download.example/f3': bytes(3),
          'https://download.example/f4': bytes(3),
          'https://download.example/f5': bytes(3),
        },
      },
      { batchByteBudget: 5 },
    );
    const { session } = makeSession({ config: { roots: [{ rootFolderId: 'FA', rootName: 'Alpha' }] } });
    const cursor: OneDriveCursor = { delta_tokens: {}, backfill: { root_index: 0, next_link: page2 } };

    const batches = (await collect(source.pull(session, cursor))) as B[];

    expect(batches.map(ids)).toEqual([['f3', 'f4'], ['f5']]);
    expect(batches[0].cursor).toEqual(cursor);
    expect(batches[1].cursor).toEqual({ delta_tokens: { FA: 'TOK1' } });
  });

  it('never splits a page when the budget is not reached (one batch per page, unchanged behaviour)', async () => {
    const pollUrl = deltaUrl('FA', 'OLD');
    const { source } = makeSource({
      deltaPages: {
        [pollUrl]: {
          value: [driveFile('f1', 'a.pdf'), driveFile('f2', 'b.pdf')],
          '@odata.deltaLink': `${GRAPH_BASE}/me/drive/items/FA/delta?token=NEW`,
        },
      },
      downloads: {
        'https://download.example/f1': bytes(3),
        'https://download.example/f2': bytes(3),
      },
    });
    const { session } = makeSession({ config: { roots: [{ rootFolderId: 'FA', rootName: 'Alpha' }] } });

    const batches = (await collect(source.pull(session, { delta_tokens: { FA: 'OLD' } }))) as B[];

    expect(batches.map(ids)).toEqual([['f1', 'f2']]);
    expect(batches[0].cursor).toEqual({ delta_tokens: { FA: 'NEW' } });
  });
});
