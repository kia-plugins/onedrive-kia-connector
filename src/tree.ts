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

/** Picker listing page size (v1's `$top=200`). */
const PICKER_TOP = 200;
const PICKER_SELECT = '$select=id,name,folder,parentReference,childCount';

interface GraphChildItem {
  id: string;
  name: string;
  folder?: { childCount?: number };
}

interface GraphChildPage {
  value?: GraphChildItem[];
  '@odata.nextLink'?: string;
}

async function listFolderNodes(client: GraphClient, startUrl: string): Promise<FolderNode[]> {
  const nodes: FolderNode[] = [];
  let url: string | undefined = startUrl;
  while (url) {
    const page: GraphChildPage = await client.request<GraphChildPage>(url);
    for (const it of page.value ?? []) {
      if (!it.folder) continue; // files aren't shown in the picker (v1 parity)
      nodes.push({ id: it.id, name: it.name, hasChildren: (it.folder.childCount ?? 0) > 0 });
    }
    url = page['@odata.nextLink'];
  }
  return nodes;
}

/** Roots of the picker's "Shared with me" tab. */
export const listSharedRoots = (client: GraphClient): Promise<FolderNode[]> =>
  listFolderNodes(client, `${GRAPH_BASE}/me/drive/sharedWithMe?${PICKER_SELECT}&$top=${PICKER_TOP}`);

/** Child FOLDERS of one picker row (`root` is Graph's My-Files alias). */
export const listChildFolders = (client: GraphClient, id: string): Promise<FolderNode[]> =>
  listFolderNodes(
    client,
    `${GRAPH_BASE}/me/drive/items/${encodeURIComponent(id)}/children?${PICKER_SELECT}&$top=${PICKER_TOP}`,
  );

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
      `${GRAPH_BASE}/me/drive/items/${encodeURIComponent(id)}?$select=id,folder`,
    );
    return { count: item.folder?.childCount ?? 0, capped: false };
  } catch {
    return null;
  }
}
