/**
 * Folder-picker callbacks (connect-time UI), ported from v1 `tree.ts` (alpha-
 * cent `git show main:src/main/connectors/onedrive/tree.ts`) onto the v2
 * `FolderPickerSpec` shape.
 *
 * v1 exposed a lazy `pageToken` parameter to its own renderer-side picker
 * component; v2's `FolderPickerSpec.children`/`roots` return the FULL node
 * list for one row in a single call (no page-token surface), so this port
 * follows `@odata.nextLink` internally to completion — same simplification
 * the google-docs-kia-connector template made for Drive's picker listings.
 *
 * v1 used three distinct pseudo-parent-ids ('me' → `/me/drive/root/children`,
 * 'sharedWithMe' → `/me/drive/sharedWithMe`, any real id →
 * `/me/drive/items/{id}/children`). Microsoft Graph accepts the literal
 * string `root` as a valid special item id everywhere a real id is accepted
 * (`/me/drive/items/root/...` ≡ `/me/drive/root/...`), so this port collapses
 * the 'me' case into the general `items/{id}/...` form using `root` as the id
 * — the same alias convention the google-docs-kia-connector template uses for
 * Drive's 'root'. Functionally identical to v1; only 'sharedWithMe' keeps its
 * own endpoint (Graph has no item-id alias for it).
 */
import type { FolderCount, FolderNode } from '@kiagent/connector-sdk';
import { GRAPH_BASE, type GraphClient } from './client';
import { itemUrl, listedRef, type MaybeRemote } from './item-ref';

/** Picker listing page size (v1's `$top=200`). */
const PICKER_TOP = 200;
// `remoteItem` must be selected explicitly: it carries the owner's drive id
// for a shared folder / shortcut (item-ref.ts).
const PICKER_SELECT = '$select=id,name,folder,parentReference,childCount,remoteItem';

interface GraphChildItem extends MaybeRemote {
  name?: string;
  folder?: { childCount?: number };
  remoteItem?: MaybeRemote['remoteItem'] & { name?: string; folder?: { childCount?: number } };
}

interface GraphChildPage {
  value?: GraphChildItem[];
  '@odata.nextLink'?: string;
}

/** `parentId` null = a listing with no parent folder ("Shared with me"). */
async function listFolderNodes(
  client: GraphClient,
  startUrl: string,
  parentId: string | null,
): Promise<FolderNode[]> {
  const nodes: FolderNode[] = [];
  let url: string | undefined = startUrl;
  while (url) {
    const page: GraphChildPage = await client.request<GraphChildPage>(url);
    for (const it of page.value ?? []) {
      // A shared folder / shortcut carries its folder facet on `remoteItem`.
      const folder = it.folder ?? it.remoteItem?.folder;
      if (!folder) continue; // files aren't shown in the picker (v1 parity)
      // A shortcut's own (possibly renamed) name wins; a sharedWithMe entry
      // may carry its name only on `remoteItem`.
      const name = it.name ?? it.remoteItem?.name ?? it.id;
      nodes.push({ id: listedRef(parentId, it), name, hasChildren: (folder.childCount ?? 0) > 0 });
    }
    url = page['@odata.nextLink'];
  }
  return nodes;
}

/** Roots of the picker's "Shared with me" tab. */
export const listSharedRoots = (client: GraphClient): Promise<FolderNode[]> =>
  listFolderNodes(client, `${GRAPH_BASE}/me/drive/sharedWithMe?${PICKER_SELECT}&$top=${PICKER_TOP}`, null);

/** Child FOLDERS of one picker row (`root` is Graph's My-Files alias; a
 *  shared row's id is drive-qualified, and so are its children's). */
export const listChildFolders = (client: GraphClient, id: string): Promise<FolderNode[]> =>
  listFolderNodes(client, `${itemUrl(id)}/children?${PICKER_SELECT}&$top=${PICKER_TOP}`, id);

/**
 * Per-row file count for the picker. v1 surfaced Graph's `folder.childCount`
 * — the item's IMMEDIATE child count (files + subfolders), free on every
 * listing — rather than a recursive file count (unlike the google-docs-kia-
 * connector template's budgeted BFS `countFilesUnder`, which Drive's API
 * requires because Drive listings don't carry a descendant count). This is a
 * faithful v1 port, not a simplification of it: v1 never computed a
 * recursive count either. Any Graph error resolves `null` (uncounted row) —
 * a count must never kill the picker.
 */
export async function countChildren(client: GraphClient, id: string): Promise<FolderCount | null> {
  try {
    const item = await client.request<{ folder?: { childCount?: number } }>(
      `${itemUrl(id)}?$select=id,folder`,
    );
    return { count: item.folder?.childCount ?? 0, capped: false };
  } catch {
    return null;
  }
}
