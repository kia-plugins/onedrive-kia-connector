# OneDrive connector for KIAgent

Indexes your Microsoft OneDrive into your local KIAgent digital memory: files
under the folders you pick are downloaded and extracted or OCR'd by the
platform, kept in sync automatically via Microsoft Graph's delta feed.

## Install

Install **OneDrive** from the KIAgent marketplace (Settings → Extensions →
Marketplace → OneDrive → Install). KIAgent will ask for the two grants this
connector needs before it activates:

- `net` — to talk to `graph.microsoft.com` (and to the pre-signed download
  URLs Graph hands back for file content);
- `query` — to check what is already indexed, so unchanged files are never
  re-downloaded.

## Connect your OneDrive

1. Add a OneDrive account. A Microsoft sign-in window opens — the OAuth flow
   is owned entirely by the platform. The connector requests the
   `Files.Read.All` and `User.Read` Graph scopes and never sees your
   Microsoft password; tokens live in KIAgent's encrypted vault and are
   refreshed by the platform.
2. Pick the folders to index in KIAgent's folder picker. Two tabs — **My
   files** and **Shared with me** — let you browse the folder tree lazily and
   select **multiple folders** (selecting a folder covers its whole subtree,
   so a selected folder's descendants can't be selected again). Each row
   shows Graph's own per-folder item count (files + subfolders directly
   inside it — not a recursive count) as an orientation aid. To index
   everything, just pick **My files** itself.
3. The account shows up under your Microsoft account's email address (or
   user principal name) and backfills from there, then checks for changes
   every 15 minutes.

You can connect multiple Microsoft accounts side by side.

### Tracked-roots config

The picker writes the account config as

```json
{ "roots": [ { "rootFolderId": "<OneDrive item id>", "rootName": "<name>" } ] }
```

with one entry per selected folder (`"root"` is Graph's alias for the whole
OneDrive). Duplicate root ids are ignored (first entry wins). If a hand-edited
config ends up with one tracked root nested inside another, the overlapping
item is still only indexed once (attributed to whichever root's walk reaches
it first) — in the normal case the picker itself never lets you select a
folder that overlaps an already-picked one.

## What gets indexed

- Microsoft Graph's delta feed still enumerates every item's metadata under
  the picked folders, but **content is requested only for files the
  platform can actually extract**: PDF, Word (.docx), Excel (.xlsx),
  CSV/HTML/plain-text and other `text/*` files are downloaded (up to 25 MiB)
  and converted locally by the engine; images (PNG/JPEG/…) up to 20 MiB go
  through the local OCR / vision pipeline.
- **Ignored** — no download, no index entry of any kind: cloud audio and
  video files (regardless of size), archives (.zip, .rar, .7z, and similar,
  regardless of size), files over the size caps above, and any other
  unsupported or unrecognized file type. An already-indexed file that a
  later policy update excludes this way is removed from your index the next
  time it's seen.
- **Folder paths** — every document records a human-readable `display_path`
  from the tracked root it was found under (and that root's id as
  `root_folder_id`), taken directly from Graph's own `parentReference.path`.
- Files deleted, trashed, or moved out of every indexed folder are archived
  from the local index, reported by Graph's own delta feed.
- Folders themselves are never indexed as documents.

Unchanged files are skipped by content hash (OneDrive's `eTag`), so re-syncs
are cheap.

## Privacy

- Read-only Graph scope (`Files.Read.All`); nothing is ever written to your
  OneDrive.
- All content stays on your machine — extraction and OCR run locally in
  KIAgent; this extension ships no Microsoft client credentials and stores no
  tokens itself (the platform's Microsoft OAuth provider owns the app
  registration and refresh).
- Only the folders you choose (or the whole OneDrive) are read.

## Limitations

- SharePoint document libraries and Teams-only shared drives are not
  indexed — OneDrive "My files" and "Shared with me" folders only.
- Binary downloads are capped at 25 MiB (20 MiB for images); larger files
  are ignored, not indexed.
- A resync (Microsoft Graph occasionally invalidates a delta token) re-walks
  that one root from scratch; already-indexed unchanged files are skipped by
  eTag, so this is cheap but not instant on a large root.
- Folders shared from another user's drive are addressed through
  `/me/drive/items/{id}` without resolving the remote drive id (same as the
  legacy connector); whether Graph transparently follows the remote item
  depends on the account type, so some work/school shared folders may not
  enumerate.
