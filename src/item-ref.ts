/**
 * Folder ids for items that live on ANOTHER user's drive.
 *
 * `/me/drive/items/{id}` only addresses the signed-in user's own drive. A
 * folder from "Shared with me" (or a "Add shortcut to My files" shortcut)
 * lives on its owner's drive, and Graph answers every `/me/drive/items/{id}`
 * call for it with 400 `invalidRequest` "The provided item ID is not valid
 * for the requested drive" — the whole account then stopped syncing (2.3.1).
 * Such an item must be addressed as `/drives/{driveId}/items/{itemId}`, with
 * the ids from its `remoteItem` facet.
 *
 * Core keeps a tracked root as an opaque `{id, name}` and rebuilds the config
 * from exactly those two fields on every Manage-folders save, so a separate
 * `driveId` field would be dropped. The drive therefore travels INSIDE the
 * id: `drive:{driveId}/{itemId}`. Neither Graph drive ids (personal hex,
 * business `b!…` base64url) nor item ids contain `/`, and no Graph item id
 * starts with `drive:`, so the encoding is unambiguous. A plain id (every
 * root saved before this, and everything in the user's own drive) keeps
 * meaning "on `/me/drive`" — existing configs and delta tokens are untouched.
 *
 * The prefix of a personal-account item id happens to be its drive id, but
 * that does not hold for work/school accounts, so the drive is never derived
 * from the id — only taken from Graph's `remoteItem.parentReference.driveId`
 * or `parentReference.driveId`.
 */
import { GRAPH_BASE } from './client';

const PREFIX = 'drive:';

export interface ItemRef {
  /** undefined = the signed-in user's own drive (`/me/drive`). */
  driveId?: string;
  itemId: string;
}

export function encodeRef(ref: ItemRef): string {
  return ref.driveId ? `${PREFIX}${ref.driveId}/${ref.itemId}` : ref.itemId;
}

export function parseRef(id: string): ItemRef {
  if (id.startsWith(PREFIX)) {
    const rest = id.slice(PREFIX.length);
    const slash = rest.indexOf('/');
    if (slash > 0 && slash < rest.length - 1) {
      return { driveId: rest.slice(0, slash), itemId: rest.slice(slash + 1) };
    }
  }
  return { itemId: id };
}

/** The Graph item URL for an encoded id (no trailing slash; append
 *  `/delta`, `/children` or a query string). */
export function itemUrl(id: string): string {
  const { driveId, itemId } = parseRef(id);
  return driveId
    ? `${GRAPH_BASE}/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(itemId)}`
    : `${GRAPH_BASE}/me/drive/items/${encodeURIComponent(itemId)}`;
}

/** Id of another item on the SAME drive as `anchorId` — the children of a
 *  shared folder, or a file inside a shared root's delta feed, come back
 *  from Graph with bare item ids. */
export function siblingRef(anchorId: string, itemId: string): string {
  const { driveId } = parseRef(anchorId);
  return encodeRef({ driveId, itemId });
}

/** Graph's own id for an encoded id — what Graph puts in `item.id`. */
export const graphItemId = (id: string): string => parseRef(id).itemId;

/** Shape of a listing entry that may carry a `remoteItem` facet. */
export interface MaybeRemote {
  id: string;
  remoteItem?: { id?: string; parentReference?: { driveId?: string } };
}

/** Encoded id for a picker listing entry. A `remoteItem` (shared-with-me
 *  entry, or a shortcut in My files) is addressed on its owner's drive;
 *  anything else is a sibling of the listed parent. */
export function listedRef(parentId: string | null, it: MaybeRemote): string {
  const remote = it.remoteItem;
  if (remote?.id && remote.parentReference?.driveId) {
    return encodeRef({ driveId: remote.parentReference.driveId, itemId: remote.id });
  }
  return parentId === null ? it.id : siblingRef(parentId, it.id);
}
