/**
 * Folders that live on ANOTHER user's drive ("Shared with me", or a shortcut
 * added to My files). Graph rejects `/me/drive/items/{id}` for them with 400
 * invalidRequest "The provided item ID is not valid for the requested
 * drive", which in 2.3.1 failed every pull of the account and every
 * Manage-folders save (field report 2026-10-11). The picker now hands out
 * drive-qualified ids (`drive:{driveId}/{itemId}`, item-ref.ts) and every
 * Graph call for such a root goes to `/drives/{driveId}/items/{itemId}`; a
 * root saved the old way (bare id) is skipped instead of failing the pull.
 */
import { createOneDriveSource, GRAPH_BASE, type OneDriveCursor, type OneDriveItem } from '../source';
import { FILE_POLICY_VERSION, type Batch, type DocumentInput, type FolderNode } from '@kiagent/connector-sdk';
import { GraphClient } from '../client';
import { encodeRef, itemUrl, parseRef, siblingRef } from '../item-ref';
import { countChildren, listChildFolders, listSharedRoots } from '../tree';
import { buildDisplayPath } from '../path-resolver';
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
  makeFolderChannel,
  makeHost,
  makeSession,
  type GraphWorld,
} from '../testing/harness';

type B = Batch<OneDriveCursor, OneDriveItem>;

// Sebastian's ids, verbatim from the log.
const OWNER_DRIVE = '34ef3bb966e5a116';
const SHARED_ITEM = '34EF3BB966E5A116!s9b34d1f040e5494998fb1de0c4d33503';
const SHARED_ROOT = `drive:${OWNER_DRIVE}/${SHARED_ITEM}`;
const sharedItemUrl = (itemId: string) => `${GRAPH_BASE}/drives/${OWNER_DRIVE}/items/${itemId}`;
const invalidForDrive = () =>
  jsonRes(400, {
    error: { code: 'invalidRequest', message: 'The provided item ID is not valid for the requested drive' },
  });

const makeClient = (world: GraphWorld) => {
  const { fetchFn, calls } = graphFetch(world);
  return {
    client: new GraphClient({ fetch: fetchFn, getToken: async () => 'ms-test-token-deadbeef', ...instantClock }),
    calls,
  };
};

describe('item-ref', () => {
  it('a plain id stays on /me/drive; a drive-qualified id round-trips to /drives/{driveId}', () => {
    expect(itemUrl('root')).toBe(`${GRAPH_BASE}/me/drive/items/root`);
    expect(itemUrl('FC8FB60D276EEE45!158028')).toBe(`${GRAPH_BASE}/me/drive/items/FC8FB60D276EEE45!158028`);
    expect(encodeRef({ driveId: OWNER_DRIVE, itemId: SHARED_ITEM })).toBe(SHARED_ROOT);
    expect(parseRef(SHARED_ROOT)).toEqual({ driveId: OWNER_DRIVE, itemId: SHARED_ITEM });
    expect(itemUrl(SHARED_ROOT)).toBe(sharedItemUrl(SHARED_ITEM));
    // A business drive id (`b!…`) survives too.
    expect(parseRef('drive:b!Xy-_9/01ABC')).toEqual({ driveId: 'b!Xy-_9', itemId: '01ABC' });
  });

  it('siblings of a shared root inherit its drive; siblings of a plain root stay plain', () => {
    expect(siblingRef(SHARED_ROOT, 'CHILD')).toBe(`drive:${OWNER_DRIVE}/CHILD`);
    expect(siblingRef('root', 'f1')).toBe('f1');
    expect(siblingRef('', 'f1')).toBe('f1');
  });
});

describe('picker', () => {
  it('"Shared with me" hands out the remoteItem drive-qualified id (folder facet on remoteItem)', async () => {
    const { client, calls } = makeClient({
      sharedRoots: [
        {
          id: SHARED_ITEM,
          name: 'Familie',
          remoteItem: { id: SHARED_ITEM, parentReference: { driveId: OWNER_DRIVE }, folder: { childCount: 4 } },
        } as never,
        { id: 'plain', name: 'Legacy', folder: { childCount: 0 } },
      ],
    });

    expect(await listSharedRoots(client)).toEqual([
      { id: SHARED_ROOT, name: 'Familie', hasChildren: true },
      { id: 'plain', name: 'Legacy', hasChildren: false },
    ]);
    expect(new URL(calls[0]).searchParams.get('$select')).toContain('remoteItem');
  });

  it('children of a shared folder are listed on its drive and keep the drive in their ids', async () => {
    const { client, calls } = makeClient({
      custom: (url) =>
        url.pathname === `/v1.0/drives/${OWNER_DRIVE}/items/${SHARED_ITEM}/children`
          ? jsonRes(200, { value: [{ id: 'SUB', name: 'Steuern', folder: { childCount: 2 } }] })
          : undefined,
    });

    expect(await listChildFolders(client, SHARED_ROOT)).toEqual([
      { id: `drive:${OWNER_DRIVE}/SUB`, name: 'Steuern', hasChildren: true },
    ]);
    expect(new URL(calls[0]).pathname).toBe(`/v1.0/drives/${OWNER_DRIVE}/items/${SHARED_ITEM}/children`);
  });

  it('a shortcut in My files is addressed on its owner drive', async () => {
    const { client } = makeClient({
      children: {
        root: [
          {
            id: 'SHORTCUT',
            name: 'Familie',
            remoteItem: { id: SHARED_ITEM, parentReference: { driveId: OWNER_DRIVE }, folder: { childCount: 1 } },
          } as never,
        ],
      },
    });

    expect(await listChildFolders(client, 'root')).toEqual([{ id: SHARED_ROOT, name: 'Familie', hasChildren: true }]);
  });

  it('counts a shared folder on its drive', async () => {
    const { client } = makeClient({
      custom: (url) =>
        url.pathname === `/v1.0/drives/${OWNER_DRIVE}/items/${SHARED_ITEM}`
          ? jsonRes(200, { id: SHARED_ITEM, folder: { childCount: 9 } })
          : undefined,
    });
    await expect(countChildren(client, SHARED_ROOT)).resolves.toEqual({ count: 9, capped: false });
  });
});

describe('syncing a shared root', () => {
  it('walks its delta on the owner drive, skips the root item, and refreshes downloads on that drive', async () => {
    const own = { rootFolderId: 'FB', rootName: 'Beta' };
    const shared = { rootFolderId: SHARED_ROOT, rootName: 'Familie' };
    const { fetchFn, calls } = graphFetch({
      deltaPages: {
        [`${sharedItemUrl(SHARED_ITEM)}/delta`]: {
          value: [
            driveFolder(SHARED_ITEM, 'Familie', { parentReference: { driveId: OWNER_DRIVE } }),
            driveFile('S1', 'scan.pdf', {
              '@microsoft.graph.downloadUrl': undefined,
              parentReference: { driveId: OWNER_DRIVE, path: `/drives/${OWNER_DRIVE}/root:/Familie` },
            }),
          ],
          '@odata.deltaLink': `${sharedItemUrl(SHARED_ITEM)}/delta?token=TS`,
        },
        [deltaUrl('FB')]: { value: [driveFile('b1', 'b.pdf')], '@odata.deltaLink': `${GRAPH_BASE}/me/drive/items/FB/delta?token=TB` },
      },
      downloads: {
        'https://download.example/S1-fresh': new Uint8Array([7]),
        'https://download.example/b1': new Uint8Array([1]),
      },
      custom: (url) =>
        url.pathname === `/v1.0/drives/${OWNER_DRIVE}/items/S1`
          ? jsonRes(200, { id: 'S1', '@microsoft.graph.downloadUrl': 'https://download.example/S1-fresh' })
          : undefined,
    });
    const source = createOneDriveSource(makeHost(fetchFn, fakeQuery()), instantClock);
    const { session } = makeSession({ config: { roots: [own, shared] } });

    const batches = (await collect(source.pull(session, null))) as B[];
    const items = batches.flatMap((b) => b.items);

    expect(items.map((i) => i.file.id).sort()).toEqual(['S1', 'b1']);
    const s1 = items.find((i) => i.file.id === 'S1')!;
    expect(s1.rootFolderId).toBe(SHARED_ROOT);
    expect(s1.displayPath).toBe('Familie / Familie / scan.pdf');
    expect((source.toDocument(s1) as DocumentInput).scopeRootId).toBe(SHARED_ROOT);
    expect(calls).toContain(`${GRAPH_BASE}/drives/${OWNER_DRIVE}/items/S1`);
    expect(calls.some((c) => c.includes('/me/drive/items/' + SHARED_ITEM))).toBe(false);
    expect(batches[batches.length - 1].cursor.delta_tokens).toEqual({ FB: 'TB', [SHARED_ROOT]: 'TS' });
  });

  it('fetchBytes re-fetches a shared-root file on the root drive', async () => {
    const { fetchFn, calls } = graphFetch({
      custom: (url) =>
        url.pathname === `/v1.0/drives/${OWNER_DRIVE}/items/S1`
          ? jsonRes(200, { id: 'S1', '@microsoft.graph.downloadUrl': 'https://download.example/S1' })
          : undefined,
      downloads: { 'https://download.example/S1': new Uint8Array([3]) },
    });
    const source = createOneDriveSource(makeHost(fetchFn), instantClock);
    const { session } = makeSession();
    const doc = fakeDoc('S1', 'file', {
      drive_item_id: 'S1',
      mime_type: 'application/pdf',
      size_bytes: 10,
      root_folder_id: SHARED_ROOT,
    });

    await expect(source.fetchBytes!(session, doc)).resolves.toEqual(new Uint8Array([3]));
    expect(calls[0]).toBe(`${GRAPH_BASE}/drives/${OWNER_DRIVE}/items/S1`);
  });
});

describe('a shared root saved by ≤2.3.1 as a bare id (the field report)', () => {
  const roots = [
    { rootFolderId: SHARED_ITEM, rootName: 'Familie' },
    { rootFolderId: 'FB', rootName: 'Beta' },
  ];

  it('is skipped with a warn — the other roots keep syncing', async () => {
    const { fetchFn } = graphFetch({
      deltaPages: {
        [deltaUrl(SHARED_ITEM)]: invalidForDrive(),
        [deltaUrl('FB', 'TOKB')]: {
          value: [driveFile('b1', 'b.pdf')],
          '@odata.deltaLink': `${GRAPH_BASE}/me/drive/items/FB/delta?token=TOKB2`,
        },
      },
      downloads: { 'https://download.example/b1': new Uint8Array([1]) },
    });
    const source = createOneDriveSource(makeHost(fetchFn, fakeQuery()), instantClock);
    const { session, logs } = makeSession({ config: { roots } });
    const cursor: OneDriveCursor = {
      delta_tokens: { FB: 'TOKB' },
      scope_roots: [SHARED_ITEM, 'FB'],
      policy_version: FILE_POLICY_VERSION,
    };

    const batches = (await collect(source.pull(session, cursor))) as B[];

    expect(batches.flatMap((b) => b.items.map((i) => i.file.id))).toEqual(['b1']);
    expect(batches[batches.length - 1].cursor.delta_tokens).toEqual({ FB: 'TOKB2' });
    expect(logs).toContainEqual({
      level: 'warn',
      msg: expect.stringContaining('"Familie" (' + SHARED_ITEM + ') is shared from another OneDrive'),
    });
  });

  it('can be removed under Manage folders (the save no longer fails)', async () => {
    const { fetchFn } = graphFetch({
      items: {
        [SHARED_ITEM]: invalidForDrive(),
        FB: driveFolder('FB', 'Beta'),
      },
      // The same folder, re-added from "Shared with me", resolves on its drive.
      custom: (url) =>
        url.pathname === `/v1.0/drives/${OWNER_DRIVE}/items/${SHARED_ITEM}`
          ? jsonRes(200, {
              id: SHARED_ITEM,
              name: 'Familie',
              parentReference: { driveId: OWNER_DRIVE, path: `/drives/${OWNER_DRIVE}/root:` },
            })
          : undefined,
    });
    const source = createOneDriveSource(makeHost(fetchFn), instantClock);
    const { session } = makeSession({
      config: { folderRoots: [{ id: 'FB', name: 'Beta' }, { id: SHARED_ITEM, name: 'Familie' }] },
      cursor: { delta_tokens: { FB: 'TB' }, scope_roots: ['FB', SHARED_ITEM] } as OneDriveCursor,
    });
    const picked: FolderNode[] = [
      { id: 'FB', name: 'Beta', hasChildren: true },
      { id: SHARED_ROOT, name: 'Familie', hasChildren: true },
    ];
    const { channel } = makeFolderChannel({ picked });

    const out = await source.manageFolders!(session, channel);

    expect(out.archiveScopeRootIds).toEqual([SHARED_ITEM]);
    expect((out.config as { folderRoots: unknown }).folderRoots).toEqual([
      { id: 'FB', name: 'Beta' },
      { id: SHARED_ROOT, name: 'Familie' },
    ]);
  });
});

describe('display path', () => {
  it("strips a shared folder's owner-drive prefix", () => {
    expect(
      buildDisplayPath(
        { rootName: 'Familie' },
        { name: 'a.pdf', parentReference: { path: `/drives/${OWNER_DRIVE}/root:/Familie/Steuern` } },
      ),
    ).toBe('Familie / Familie / Steuern / a.pdf');
  });
});

describe('review fixes', () => {
  const SUB = `drive:${OWNER_DRIVE}/SUB`;
  const subItem = (url: URL) =>
    url.pathname === `/v1.0/drives/${OWNER_DRIVE}/items/SUB`
      ? jsonRes(200, {
          id: 'SUB',
          name: 'Steuern',
          // Relative to the drive it was requested on — the OWNER's.
          parentReference: { driveId: OWNER_DRIVE, id: 'SHARED', path: '/drive/root:/Familie' },
        })
      : url.pathname === `/v1.0/drives/${OWNER_DRIVE}/items/SHARED`
        ? jsonRes(200, {
            id: 'SHARED',
            name: 'Familie',
            parentReference: { driveId: OWNER_DRIVE, id: 'OWNERROOT', path: '/drive/root:' },
          })
        : undefined;

  it('a removed shared root is NOT covered by the retained whole-OneDrive root (archived, not re-attributed)', async () => {
    const { fetchFn } = graphFetch({ custom: subItem });
    const source = createOneDriveSource(makeHost(fetchFn), instantClock);
    const { session } = makeSession({
      config: { folderRoots: [{ id: 'root', name: 'OneDrive' }, { id: SUB, name: 'Steuern' }] },
      cursor: { delta_tokens: { root: 'TR', [SUB]: 'TS' }, scope_roots: ['root', SUB] } as OneDriveCursor,
    });
    const { channel } = makeFolderChannel({ picked: [{ id: 'root', name: 'OneDrive', hasChildren: true }] });

    const out = await source.manageFolders!(session, channel);

    expect(out.archiveScopeRootIds).toEqual([SUB]);
    expect(out.reattributeScopeRoots).toEqual([]);
  });

  it('Manage folders reveals a shared sub-folder through its drive-qualified parent, never via My files', async () => {
    const { fetchFn, calls } = graphFetch({ custom: subItem });
    const source = createOneDriveSource(makeHost(fetchFn), instantClock);
    const { session } = makeSession({ config: { folderRoots: [{ id: SUB, name: 'Steuern' }] } });
    const { channel, getPickerSpec } = makeFolderChannel({ picked: [{ id: SUB, name: 'Steuern', hasChildren: true }] });

    await source.manageFolders!(session, channel);

    expect(getPickerSpec()!.expand).toEqual([`drive:${OWNER_DRIVE}/SHARED`]);
    expect(calls.some((c) => c.includes('/me/drive/items/'))).toBe(false);
  });

  it('only the wrong-drive 400 on a BARE root is skipped; other 400s still fail the pull', async () => {
    const pullWith = async (rootId: string, res: ReturnType<typeof jsonRes>) => {
      const url = `${itemUrl(rootId)}/delta`;
      const { fetchFn } = graphFetch({ deltaPages: { [url]: res } });
      const source = createOneDriveSource(makeHost(fetchFn, fakeQuery()), instantClock);
      const { session } = makeSession({ config: { roots: [{ rootFolderId: rootId, rootName: 'X' }] } });
      return collect(source.pull(session, null));
    };
    await expect(
      pullWith('BARE', jsonRes(400, { error: { code: 'invalidRequest', message: 'Invalid query option' } })),
    ).rejects.toThrow(/graph 400/);
    await expect(pullWith(SHARED_ROOT, invalidForDrive())).rejects.toThrow(/graph 400/);
  });

  it('a sharedWithMe entry whose name is only on remoteItem keeps it', async () => {
    const { client } = makeClient({
      sharedRoots: [
        {
          id: SHARED_ITEM,
          remoteItem: {
            id: SHARED_ITEM,
            name: 'Familie',
            parentReference: { driveId: OWNER_DRIVE },
            folder: { childCount: 0 },
          },
        } as never,
      ],
    });
    expect(await listSharedRoots(client)).toEqual([{ id: SHARED_ROOT, name: 'Familie', hasChildren: false }]);
  });
});
