/**
 * v2 port of v1 `path-resolver.ts` (v1 repo `git show
 * main:src/main/connectors/onedrive/path-resolver.ts`), verbatim algorithm.
 *
 * Pure-local: Microsoft Graph carries `parentReference.path` on every
 * DriveItem, so building the human-readable path needs no network call and
 * no folder index (unlike Google Drive, whose v2 port walks an ancestor
 * chain — Graph gives the whole path string for free).
 */

export interface DisplayPathRoot {
  rootName: string;
}

export interface DisplayPathItem {
  name: string;
  parentReference?: { path?: string };
}

/**
 * Build the user-facing path for an item under a tracked root.
 *
 * For items returned via the "Shared with me" root, `parentReference.path`
 * is either missing or relative to a different drive; in those cases this
 * falls back to "<root> / <name>" so the path is at least informative (v1
 * parity).
 */
export function buildDisplayPath(root: DisplayPathRoot, item: DisplayPathItem): string {
  const raw = item.parentReference?.path;
  if (!raw) return `${root.rootName} / ${item.name}`;
  // Own drive: `/drive/root:/…`; a shared folder's owner drive:
  // `/drives/{driveId}/root:/…`.
  const stripped = raw.replace(/^\/(?:drive|drives\/[^/]+)\/root:?/, '');
  const segments = stripped
    .split('/')
    .filter((s) => s.length > 0)
    .map(decodeSegment);
  return [root.rootName, ...segments, item.name].join(' / ');
}

/** Graph does not always percent-encode a literal `%` in a folder name
 *  (`/100% Rabatt`), and one such segment must not cost the item its row —
 *  it keeps its raw spelling instead. */
function decodeSegment(segment: string): string {
  try {
    return decodeURIComponent(segment);
  } catch {
    return segment;
  }
}
