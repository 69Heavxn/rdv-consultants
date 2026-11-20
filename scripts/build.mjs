// Construit docs/ (GitHub Pages) : la page de prise de RDV, avec le code du
// favori Salesforce dedans (installé depuis « Réglages »).
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { build } from 'esbuild';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const out = path.join(root, 'docs');
fs.mkdirSync(out, { recursive: true });
const pkg = JSON.parse(fs.readFileSync(path.join(root, 'package.json'), 'utf8'));
const version = `v${pkg.version} (${new Date().toISOString().slice(0, 10)})`;
const common = { bundle: true, minify: true, format: 'iife', target: ['chrome100'], charset: 'ascii', legalComments: 'none', write: false };

// 1. le favori : une adresse javascript: en ASCII pur
const bm = await build({ ...common, entryPoints: [path.join(root, 'src/bookmarklet.js')] });
const bmJs = bm.outputFiles[0].text.trim()
  .replace(/[^\x00-\x7f]/g, (c) => '\\u' + c.charCodeAt(0).toString(16).padStart(4, '0'));
if (!bmJs.includes('__RDV_URL__')) throw new Error('le favori doit contenir __RDV_URL__');
// une adresse javascript: est décodée (%XX) avant exécution et le navigateur
// supprime les retours à la ligne : on échappe ce qui doit l'être
const bmUrl = 'javascript:' + bmJs
  .replace(/%/g, '%25').replace(/\r/g, '%0D').replace(/\n/g, '%0A').replace(/\t/g, '%09').replace(/#/g, '%23');

// 2. la page
const app = await build({
  ...common,
  entryPoints: [path.join(root, 'src/app.js')],
  define: { __VERSION__: JSON.stringify(version), __BOOKMARKLET__: JSON.stringify(bmUrl) },
});
const appJs = app.outputFiles[0].text;
fs.writeFileSync(path.join(out, 'app.js'), appJs);
fs.writeFileSync(path.join(out, 'index.html'), fs.readFileSync(path.join(root, 'web/index.html'), 'utf8'));
fs.writeFileSync(path.join(out, '.nojekyll'), '');
console.log(`RDV consultants ${version}`);
console.log(`  docs/app.js  ${(Buffer.byteLength(appJs) / 1024).toFixed(0)} KB (favori ${(bmUrl.length / 1024).toFixed(1)} KB)`);
