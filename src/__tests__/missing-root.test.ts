/**
 * A tracked root that no longer exists upstream (Graph 404 itemNotFound on
 * its /delta feed — folder deleted, or access lost) is skipped with a warn,
 * never fatal: the other roots keep syncing, and a backfill or policy
 * re-walk still completes. The root is re-probed on every pull, so a folder
 * restored from the recycle bin is picked up again.
 */
import { createOneDriveSource, GRAPH_BASE, type OneDriveCursor, type OneDriveItem } from '../source';
import { FILE_POLICY_VERSION, type Batch } from '@kiagent/connector-sdk';
import {
  collect,
  deltaUrl,
  driveFile,
  fakeQuery,
  graphFetch,
  instantClock,
  jsonRes,
  makeHost,
  makeSession,
} from '../testing/harness';

type B = Batch<OneDriveCursor, OneDriveItem>;
const ids = (b: B) => b.items.map((i) => i.file.id);
const notFound = () => jsonRes(404, { error: { code: 'itemNotFound', message: 'The resource could not be found.' } });
const deltaLink = (root: string, token: string) => `${GRAPH_BASE}/me/drive/items/${root}/delta?token=${token}`;

function makeSource(world: Parameters<typeof graphFetch>[0]) {
  const { fetchFn, calls } = graphFetch(world);
  const source = createOneDriveSource(makeHost(fetchFn, fakeQuery()), instantClock);
  return { source, calls };
}

const roots = [
  { rootFolderId: 'GONE', rootName: 'Deleted folder' },
  { rootFolderId: 'FB', rootName: 'Beta' },
];
const downloads = { 'https://download.example/b1': new Uint8Array([1]) };

describe('a tracked root that is gone upstream', () => {
  it('backfill skips it with a warn and establishes the other roots', async () => {
    const { source } = makeSource({
      deltaPages: {
        [deltaUrl('GONE')]: notFound(),
        [deltaUrl('FB')]: { value: [driveFile('b1', 'b.pdf')], '@odata.deltaLink': deltaLink('FB', 'TOKB') },
      },
      downloads,
    });
    const { session, logs } = makeSession({ config: { roots } });

    const batches = (await collect(source.pull(session, null))) as B[];

    expect(batches.flatMap(ids)).toEqual(['b1']);
    expect(batches[batches.length - 1].cursor.delta_tokens).toEqual({ FB: 'TOKB' });
    expect(batches[batches.length - 1].cursor.backfill).toBeUndefined();
    expect(logs).toContainEqual({ level: 'warn', msg: expect.stringContaining('"Deleted folder" (GONE) no longer exists') });
  });

  it('a live poll skips it and still polls the other roots', async () => {
    const { source, calls } = makeSource({
      deltaPages: {
        [deltaUrl('GONE', 'TOKG')]: notFound(),
        [deltaUrl('FB', 'TOKB')]: { value: [driveFile('b1', 'b.pdf')], '@odata.deltaLink': deltaLink('FB', 'TOKB2') },
      },
      downloads,
    });
    const { session } = makeSession({ config: { roots } });
    const cursor: OneDriveCursor = { delta_tokens: { GONE: 'TOKG', FB: 'TOKB' }, policy_version: FILE_POLICY_VERSION };

    const batches = (await collect(source.pull(session, cursor))) as B[];

    expect(calls).toContain(deltaUrl('FB', 'TOKB'));
    expect(batches.flatMap(ids)).toEqual(['b1']);
    expect(batches[batches.length - 1].cursor.delta_tokens).toEqual({ GONE: 'TOKG', FB: 'TOKB2' });
  });

  it('a tokenless gone root does not stall polling of the established roots', async () => {
    const { source, calls } = makeSource({
      deltaPages: {
        [deltaUrl('GONE')]: notFound(),
        [deltaUrl('FB', 'TOKB')]: { value: [driveFile('b1', 'b.pdf')], '@odata.deltaLink': deltaLink('FB', 'TOKB2') },
      },
      downloads,
    });
    const { session } = makeSession({ config: { roots } });
    const cursor: OneDriveCursor = {
      delta_tokens: { FB: 'TOKB' },
      scope_roots: ['GONE', 'FB'],
      policy_version: FILE_POLICY_VERSION,
    };

    const batches = (await collect(source.pull(session, cursor))) as B[];

    expect(calls.filter((c) => c.includes('/delta'))).toEqual([deltaUrl('GONE'), deltaUrl('FB', 'TOKB')]);
    expect(batches[0].cursor.backfill).toBeUndefined();
    expect(batches.flatMap(ids)).toEqual(['b1']);
    expect(batches[batches.length - 1].cursor.delta_tokens).toEqual({ FB: 'TOKB2' });
  });

  it('the policy re-walk completes past it and polls the kept tokens', async () => {
    const { source, calls } = makeSource({
      deltaPages: {
        [deltaUrl('GONE')]: notFound(),
        [deltaUrl('FB')]: { value: [driveFile('b1', 'b.pdf')], '@odata.deltaLink': deltaLink('FB', 'FRESH') },
        [deltaUrl('FB', 'OLDB')]: { value: [], '@odata.deltaLink': deltaLink('FB', 'TOKB2') },
      },
      downloads,
    });
    const { session } = makeSession({ config: { roots } });
    const cursor: OneDriveCursor = { delta_tokens: { GONE: 'OLDG', FB: 'OLDB' }, policy_version: 1 };

    const batches = (await collect(source.pull(session, cursor))) as B[];

    expect(calls).toContain(deltaUrl('FB', 'OLDB'));
    expect(calls).not.toContain(deltaUrl('GONE', 'OLDG'));
    const last = batches[batches.length - 1].cursor;
    expect(last.rewalk).toBeUndefined();
    expect(last.policy_version).toBe(FILE_POLICY_VERSION);
    expect(last.delta_tokens.FB).toBe('TOKB2');
  });

  it('any other Graph error still fails the pull', async () => {
    const { source } = makeSource({
      deltaPages: {
        [deltaUrl('GONE')]: jsonRes(403, { error: { code: 'accessDenied' } }),
        [deltaUrl('FB')]: { value: [], '@odata.deltaLink': deltaLink('FB', 'TOKB') },
      },
    });
    const { session } = makeSession({ config: { roots } });

    await expect(collect(source.pull(session, null))).rejects.toThrow(/graph 403/);
  });
});
