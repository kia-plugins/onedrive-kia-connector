/**
 * v2 port of v1 `path-resolver.ts` (alpha-cent `git show
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
  const stripped = raw.replace(/^\/drive\/root:?/, '');
  const segments = decodeURIComponent(stripped)
    .split('/')
    .filter((s) => s.length > 0);
  return [root.rootName, ...segments, item.name].join(' / ');
}
