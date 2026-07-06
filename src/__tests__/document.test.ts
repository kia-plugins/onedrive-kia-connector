/**
 * toDocument (pure) and fetchBytes (deep-extraction re-fetch) suite.
 */
import { createOneDriveSource, type OneDriveItem } from '../source';
import { GraphApiError } from '../client';
import type { DocumentInput } from '../kiagent-contracts';
import { driveFile, fakeDoc, graphFetch, instantClock, jsonRes, makeHost, makeSession } from '../testing/harness';

function makeSource(world: Parameters<typeof graphFetch>[0] = {}) {
  const { fetchFn, calls } = graphFetch(world);
  const source = createOneDriveSource(makeHost(fetchFn), instantClock);
  return { source, calls };
}

describe('toDocument', () => {
  it('maps a binary item to a file DocumentInput with binary bytes and vision aliases', () => {
    const { source } = makeSource();
    const item: OneDriveItem = {
      file: driveFile('f1', 'report.pdf', { size: 4096 }),
      markdown: null,
      bytes: new Uint8Array([1, 2, 3]),
      extractionStatus: 'ok',
      displayPath: 'Alpha / report.pdf',
      rootFolderId: 'FA',
    };

    const doc = source.toDocument(item);
    expect(Array.isArray(doc)).toBe(false);
    expect(doc).not.toBeNull();
    const d = doc as DocumentInput;
    expect(d).toMatchObject({
      externalId: 'f1',
      type: 'file',
      title: 'report.pdf',
      markdown: null,
      binary: { bytes: new Uint8Array([1, 2, 3]), mime: 'application/pdf', filename: 'report.pdf' },
      url: expect.stringContaining('f1'),
    });
    expect(d.metadata).toMatchObject({
      drive_item_id: 'f1',
      mime_type: 'application/pdf',
      size_bytes: 4096,
      display_path: 'Alpha / report.pdf',
      root_folder_id: 'FA',
      extraction_status: 'ok',
      mime: 'application/pdf',
      filename: 'report.pdf',
      sizeBytes: 4096,
    });
  });

  it('maps a metadata-only item with markdown="" and no binary field', () => {
    const { source } = makeSource();
    const item: OneDriveItem = {
      file: driveFile('z1', 'archive.zip', { file: { mimeType: 'application/zip' } }),
      markdown: '',
      extractionStatus: 'unsupported',
      displayPath: 'Alpha / archive.zip',
      rootFolderId: 'FA',
    };

    const d = source.toDocument(item) as { binary?: unknown; markdown: string | null };
    expect(d.markdown).toBe('');
    expect(d.binary).toBeUndefined();
  });

  it('falls back to the onedrive.live.com URL when webUrl is absent', () => {
    const { source } = makeSource();
    const item: OneDriveItem = {
      file: driveFile('f1', 'a.pdf', { webUrl: undefined }),
      markdown: null,
      bytes: new Uint8Array([1]),
      extractionStatus: 'ok',
      displayPath: 'Alpha / a.pdf',
      rootFolderId: 'FA',
    };
    const d = source.toDocument(item) as { url?: string };
    expect(d.url).toBe('https://onedrive.live.com/');
  });
});

describe('fetchBytes', () => {
  it('re-fetches fresh bytes via a refreshed downloadUrl for a convertible doc', async () => {
    const { source } = makeSource({
      items: { f1: { id: 'f1', name: 'a.pdf', '@microsoft.graph.downloadUrl': 'https://download.example/f1-fresh' } },
      downloads: { 'https://download.example/f1-fresh': new Uint8Array([5, 5]) },
    });
    const { session } = makeSession();
    const doc = fakeDoc('f1', 'file', { drive_item_id: 'f1', mime_type: 'application/pdf', size_bytes: 100 });

    await expect(source.fetchBytes!(session, doc)).resolves.toEqual(new Uint8Array([5, 5]));
  });

  it('returns null for a non-convertible mime without fetching', async () => {
    const { source, calls } = makeSource();
    const { session } = makeSession();
    const doc = fakeDoc('z1', 'file', { drive_item_id: 'z1', mime_type: 'application/zip', size_bytes: 10 });

    await expect(source.fetchBytes!(session, doc)).resolves.toBeNull();
    expect(calls).toHaveLength(0);
  });

  it('returns null when the size in metadata exceeds the cap', async () => {
    const { source, calls } = makeSource();
    const { session } = makeSession();
    const doc = fakeDoc('f1', 'file', {
      drive_item_id: 'f1',
      mime_type: 'application/pdf',
      size_bytes: 26 * 1024 * 1024,
    });

    await expect(source.fetchBytes!(session, doc)).resolves.toBeNull();
    expect(calls).toHaveLength(0);
  });

  it('returns null when the item is gone upstream (404/410)', async () => {
    const { source } = makeSource({
      custom: (url) =>
        url.pathname === '/v1.0/me/drive/items/f1'
          ? jsonRes(404, { error: { message: 'itemNotFound' } })
          : undefined,
    });
    const { session } = makeSession();
    const doc = fakeDoc('f1', 'file', { drive_item_id: 'f1', mime_type: 'application/pdf', size_bytes: 100 });

    await expect(source.fetchBytes!(session, doc)).resolves.toBeNull();
  });

  it('propagates other Graph errors', async () => {
    const { source } = makeSource({
      custom: (url) =>
        url.pathname === '/v1.0/me/drive/items/f1'
          ? jsonRes(500, { error: { message: 'boom' } })
          : undefined,
    });
    const { session } = makeSession();
    const doc = fakeDoc('f1', 'file', { drive_item_id: 'f1', mime_type: 'application/pdf', size_bytes: 100 });

    await expect(source.fetchBytes!(session, doc)).rejects.toThrow(GraphApiError);
  });

  it('returns null when metadata has no drive_item_id', async () => {
    const { source } = makeSource();
    const { session } = makeSession();
    const doc = fakeDoc('f1', 'file', { mime_type: 'application/pdf' });

    await expect(source.fetchBytes!(session, doc)).resolves.toBeNull();
  });
});
