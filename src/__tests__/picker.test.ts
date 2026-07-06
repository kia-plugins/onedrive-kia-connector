/**
 * Folder-picker callback suite (tree.ts): shared-with-me root listing
 * (pagination), child-folder listing (folders only, root alias), and the
 * per-row childCount (v1 parity — immediate children, not a recursive
 * count; Graph error → null).
 */
import { countChildren, listChildFolders, listSharedRoots } from '../tree';
import { GraphClient } from '../client';
import type { NetFetch } from '../client';
import { graphFetch, instantClock, jsonRes } from '../testing/harness';

const makeClient = (fetchFn: NetFetch) =>
  new GraphClient({ fetch: fetchFn, getToken: async () => 'ms-test-token-deadbeef', ...instantClock });

describe('listSharedRoots', () => {
  it('lists shared-with-me folders across pages as hasChildren nodes (files filtered out)', async () => {
    const { fetchFn, calls } = graphFetch({
      sharedRoots: [
        [
          { id: 'SH1', name: 'Alpha', folder: { childCount: 3 } },
          { id: 'SH2', name: 'Beta', folder: { childCount: 0 } },
          { id: 'f1', name: 'not-a-folder.txt' },
        ],
        [{ id: 'SH3', name: 'Gamma', folder: { childCount: 5 } }],
      ],
    });

    const nodes = await listSharedRoots(makeClient(fetchFn));

    expect(nodes).toEqual([
      { id: 'SH1', name: 'Alpha', hasChildren: true },
      { id: 'SH2', name: 'Beta', hasChildren: false },
      { id: 'SH3', name: 'Gamma', hasChildren: true },
    ]);
    expect(calls).toHaveLength(2);
    const u = new URL(calls[0]);
    expect(u.pathname).toBe('/v1.0/me/drive/sharedWithMe');
    expect(u.searchParams.get('$top')).toBe('200');
  });
});

describe('listChildFolders', () => {
  it('lists only child FOLDERS of the given id, hasChildren from childCount', async () => {
    const { fetchFn, calls } = graphFetch({
      children: {
        F1: [
          { id: 'C1', name: 'Sub A', folder: { childCount: 1 } },
          { id: 'd1', name: 'Doc.docx' },
        ],
      },
    });

    const nodes = await listChildFolders(makeClient(fetchFn), 'F1');

    expect(nodes).toEqual([{ id: 'C1', name: 'Sub A', hasChildren: true }]);
    const u = new URL(calls[0]);
    expect(u.pathname).toBe('/v1.0/me/drive/items/F1/children');
  });

  it("uses Graph's 'root' alias for the OneDrive-root node uniformly with any other id", async () => {
    const { fetchFn, calls } = graphFetch({
      children: { root: [{ id: 'C1', name: 'Documents', folder: { childCount: 0 } }] },
    });

    const nodes = await listChildFolders(makeClient(fetchFn), 'root');

    expect(nodes).toEqual([{ id: 'C1', name: 'Documents', hasChildren: false }]);
    const u = new URL(calls[0]);
    expect(u.pathname).toBe('/v1.0/me/drive/items/root/children');
  });
});

describe('countChildren', () => {
  it("resolves the item's immediate childCount (not recursive — v1 parity)", async () => {
    const { fetchFn, calls } = graphFetch({
      items: { F1: { id: 'F1', name: 'F1', folder: { childCount: 7 } } },
    });

    await expect(countChildren(makeClient(fetchFn), 'F1')).resolves.toEqual({
      count: 7,
      capped: false,
    });
    const u = new URL(calls[0]);
    expect(u.pathname).toBe('/v1.0/me/drive/items/F1');
  });

  it('resolves 0 when folder facet/childCount is absent', async () => {
    const { fetchFn } = graphFetch({ items: { F1: { id: 'F1', name: 'F1' } } });
    await expect(countChildren(makeClient(fetchFn), 'F1')).resolves.toEqual({
      count: 0,
      capped: false,
    });
  });

  it('resolves null on a Graph API error instead of rejecting', async () => {
    const { fetchFn } = graphFetch({
      custom: (url) =>
        url.pathname === '/v1.0/me/drive/items/F1'
          ? jsonRes(404, { error: { message: 'itemNotFound' } })
          : undefined,
    });

    await expect(countChildren(makeClient(fetchFn), 'F1')).resolves.toBeNull();
  });
});
