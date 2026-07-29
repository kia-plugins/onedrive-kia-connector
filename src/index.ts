import type { ExtensionModule } from '@kiagent/connector-sdk';
import { createOneDriveSource } from './source';

const mod = {
  async activate(host) {
    return { sources: [createOneDriveSource(host)] };
  },
} satisfies ExtensionModule<'net' | 'query'>;

export default mod;
module.exports = mod; // dual export — the host child require()s CJS
