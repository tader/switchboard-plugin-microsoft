import { build } from 'esbuild';
import fs from 'node:fs/promises';
await build({
  entryPoints: ['scripts/vendor/xml.js'],
  bundle: true,
  platform: 'node',
  target: 'node24',
  format: 'esm',
  outfile: 'plugins/microsoft/vendor/xml.js',
  legalComments: 'inline',
});
await fs.copyFile(
  'node_modules/@xmldom/xmldom/LICENSE',
  'plugins/microsoft/vendor/XML-LICENSE',
);
