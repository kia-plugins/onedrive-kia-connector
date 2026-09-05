/**
 * Folder scope: the canonical `folderRoots` config (with the R1 legacy
 * `roots` mirror as a read-only fallback), the `scopeRootId` stamp on every
 * emitted document, and the `scope_roots` cursor field that lets a pull tell
 * a scope CHANGE from a plain resume.
 */
import {
  GRAPH_BASE,
  ONEDRIVE_DRIVE_ROOT_ID,
  createOneDriveSource,
  normalizeCursor,
  rootsConfig,
  type OneDriveCursor,
  type OneDriveItem,
} from '../source';
import type { Batch, DocumentInput } from '@kiagent/connector-sdk';
import {
  collect,
  deltaUrl,
  driveFile,
  graphFetch,
  instantClock,
  makeAuth,
  makeHost,
  makeSession,
} from '../testing/harness';

type B = Batch<OneDriveCursor, OneDriveItem>;
const R = (...ids: string[]) => ids.map((id) => ({ rootFolderId: id, rootName: id }));
const cfg = (config: Record<string, unknown>) => rootsConfig(makeSession({ config }).session);

describe('canonical folderRoots config', () => {
  it('reads the canonical folderRoots shape, in order', () => {
    expect(
      cfg({
        folderRoots: [
          { id: 'FA', name: 'Alpha' },
          { id: 'SH1', name: 'Shared specs' },
        ],
      }),
    ).toEqual([
      { rootFolderId: 'FA', rootName: 'Alpha' },
      { rootFolderId: 'SH1', rootName: 'Shared specs' },
    ]);
  });

  it('prefers canonical folderRoots over the legacy roots mirror and never merges the two', () => {
    expect(
      cfg({
        folderRoots: [{ id: 'FA', name: 'Alpha' }],
        roots: [{ rootFolderId: 'OLD', rootName: 'Stale mirror' }],
      }),
    ).toEqual([{ rootFolderId: 'FA', rootName: 'Alpha' }]);
  });

  it('falls back to the legacy roots shape when folderRoots is absent or yields nothing usable', () => {
    expect(cfg({ roots: [{ rootFolderId: 'OLD', rootName: 'Kept' }] })).toEqual([
      { rootFolderId: 'OLD', rootName: 'Kept' },
    ]);
    expect(cfg({ folderRoots: [], roots: [{ rootFolderId: 'OLD', rootName: 'Kept' }] })).toEqual([
      { rootFolderId: 'OLD', rootName: 'Kept' },
    ]);
    expect(cfg({ folderRoots: [{ id: 7 }], roots: [{ rootFolderId: 'OLD', rootName: 'Kept' }] })).toEqual([
      { rootFolderId: 'OLD', rootName: 'Kept' },
    ]);
  });

  it('the whole-OneDrive fallback and the picker root node are the SAME literal as the exported drive-root id', async () => {
    expect(ONEDRIVE_DRIVE_ROOT_ID).toBe('root');
    expect(cfg({})).toEqual([{ rootFolderId: ONEDRIVE_DRIVE_ROOT_ID, rootName: 'OneDrive' }]);

    const { fetchFn } = graphFetch({});
    const source = createOneDriveSource(makeHost(fetchFn), instantClock);
    const { auth, getPickerSpec } = makeAuth();
    await source.connect(auth);
    await expect(getPickerSpec()!.roots('my-files')).resolves.toEqual([
      { id: ONEDRIVE_DRIVE_ROOT_ID, name: 'OneDrive', hasChildren: true },
    ]);
  });
});

describe('scopeRootId', () => {
  it('stamps the emitting root, alongside the legacy metadata.root_folder_id the v3 migration reads', () => {
    const { fetchFn } = graphFetch({});
    const source = createOneDriveSource(makeHost(fetchFn), instantClock);
    const item: OneDriveItem = {
      file: driveFile('f1', 'a.pdf'),
      markdown: null,
      bytes: new Uint8Array([1]),
      extractionStatus: 'ok',
      displayPath: 'Alpha / a.pdf',
      rootFolderId: 'FA',
    };

    const d = source.toDocument(item) as DocumentInput;

    expect(d.scopeRootId).toBe('FA');
    expect((d.metadata as Record<string, unknown>).root_folder_id).toBe('FA');
  });
});

describe('normalizeCursor', () => {
  it('a matching scope rides through UNTOUCHED — order-independent, backfill resume preserved', () => {
    const c: OneDriveCursor = {
      delta_tokens: { FA: 'TA', FB: 'TB' },
      scope_roots: ['FB', 'FA'],
      backfill: { root_index: 1, next_link: 'https://graph.microsoft.com/page2' },
    };
    expect(normalizeCursor(c, R('FA', 'FB'))).toBe(c);
  });

  it('an ABSENT scope_roots is a mismatch: tokens are pruned and the backfill key is DELETED, not reset', () => {
    expect(
      normalizeCursor(
        {
          delta_tokens: { FA: 'TA', GONE: 'TG' },
          backfill: { root_index: 0, next_link: 'https://graph.microsoft.com/page2' },
        },
        R('FA'),
      ),
    ).toEqual({ delta_tokens: { FA: 'TA' }, scope_roots: ['FA'] });
  });

  it('a de-selected root loses its delta token — the append-only leak, closed', () => {
    expect(
      normalizeCursor({ delta_tokens: { FA: 'TA', FB: 'TB' }, scope_roots: ['FA', 'FB'] }, R('FA')),
    ).toEqual({ delta_tokens: { FA: 'TA' }, scope_roots: ['FA'] });
  });

  it('a null cursor stays null', () => {
    expect(normalizeCursor(null, R('FA'))).toBeNull();
  });
});

describe('scope_roots is stamped on every persisted cursor', () => {
  it('every batch of a multi-page backfill carries the configured root ids', async () => {
    const page2 = `${GRAPH_BASE}/me/drive/items/FA/delta?$skiptoken=p2`;
    const { fetchFn } = graphFetch({
      deltaPages: {
        [deltaUrl('FA')]: { value: [driveFile('f1', 'a.pdf')], '@odata.nextLink': page2 },
        [page2]: { value: [], '@odata.deltaLink': `${GRAPH_BASE}/me/drive/items/FA/delta?token=TOK1` },
      },
      downloads: { 'https://download.example/f1': new Uint8Array([1]) },
    });
    const source = createOneDriveSource(makeHost(fetchFn), instantClock);
    const { session } = makeSession({ config: { folderRoots: [{ id: 'FA', name: 'Alpha' }] } });

    const batches = (await collect(source.pull(session, null))) as B[];

    expect(batches.map((b) => b.cursor.scope_roots)).toEqual([['FA'], ['FA']]);
  });

  it('a pre-folder-scope cursor re-enumerates instead of resuming a stale next_link (A5c)', async () => {
    const page2 = `${GRAPH_BASE}/me/drive/items/FA/delta?$skiptoken=p2`;
    const { fetchFn, calls } = graphFetch({
      deltaPages: {
        [deltaUrl('FA')]: {
          value: [],
          '@odata.deltaLink': `${GRAPH_BASE}/me/drive/items/FA/delta?token=TOK1`,
        },
      },
    });
    const source = createOneDriveSource(makeHost(fetchFn), instantClock);
    const { session } = makeSession({ config: { folderRoots: [{ id: 'FA', name: 'Alpha' }] } });

    const batches = (await collect(
      source.pull(session, { delta_tokens: {}, backfill: { root_index: 0, next_link: page2 } }),
    )) as B[];

    expect(calls).toEqual([deltaUrl('FA')]);
    expect(batches.at(-1)!.cursor).toEqual({ delta_tokens: { FA: 'TOK1' }, scope_roots: ['FA'] });
  });
});
