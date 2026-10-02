/**
 * Binary-mime routing for OneDrive items. v1 (`ingest.ts`) delegated this to
 * `@main/converter`'s `isConvertibleMime` and ran the extraction ITSELF
 * (download bytes, call the local converter, store markdown). v2 dropped
 * that: every downloadable file is emitted with `binary: { bytes, mime,
 * filename }` and `markdown: null`, and the v2 ENGINE's convert/vision
 * pipeline does the extraction — this module only decides which files are
 * worth downloading at all.
 *
 * v3 (this file) replaces the connector-local `isConvertibleMime` allowlist
 * with kiagent-core's canonical `decideFileIndexing` policy (via the SDK):
 * the SAME "can this file be indexed?" gate the local-folder source, the
 * vision worker, the audio worker and the google-docs-kia-connector template
 * all derive from, so this connector can no longer drift from them. Under
 * `profile: 'cloud-drive'` the policy ignores archives at any size, all
 * cloud audio/video regardless of size, and anything outside its allowlist
 * — and caps images at `MAX_CLOUD_IMAGE_BYTES` (20 MiB). Documents over the
 * eager cap (`MAX_CLOUD_BINARY_BYTES`, 25 MiB) still get a row: the policy's
 * `bytes` says `'deferred'` (core's convert worker fetches it later, up to
 * `MAX_FETCH_BYTES`) or `'none'` (metadata only); only media is
 * `too-large`. See `source.ts`'s `pageChunks` for where `ignore` routes gate
 * OUT of any content I/O entirely (no downloadUrl refresh, no download).
 */
import { decideFileIndexing, type FileIgnoreReason } from '@kiagent/connector-sdk';

export type OneDriveRoute =
  | { kind: 'binary'; pipeline: 'converter' | 'vision'; bytes: 'eager' | 'deferred' | 'none' }
  | { kind: 'ignore'; reason: FileIgnoreReason };

export function chooseRoute(mimeType: string, filename: string, sizeBytes?: number): OneDriveRoute {
  const d = decideFileIndexing({
    profile: 'cloud-drive',
    filename,
    mime: mimeType,
    sizeBytes,
  });
  return d.kind === 'ignore'
    ? d
    : { kind: 'binary', pipeline: d.pipeline === 'vision' ? 'vision' : 'converter', bytes: d.bytes };
}

/** The `extraction_status` a row built from this route carries. */
export function statusFor(route: Extract<OneDriveRoute, { kind: 'binary' }>): 'ok' | 'deferred' | 'none' {
  return route.bytes === 'eager' ? 'ok' : route.bytes;
}
