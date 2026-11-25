import { build } from 'esbuild';
import { copyFileSync, renameSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs';
const js = '../ControlAddin/scripts/vendor';
const css = '../ControlAddin/stylesheets/vendor';
mkdirSync(js, { recursive: true }); mkdirSync(css, { recursive: true });
await build({
  entryPoints: ['entry.js'], bundle: true, minify: true, format: 'iife', globalName: 'AKUniver',
  target: ['es2020'], outfile: `${js}/univer.bundle.js`, legalComments: 'none',
  define: { 'process.env.NODE_ENV': '"production"' }, loader: { '.woff': 'dataurl', '.woff2': 'dataurl', '.ttf': 'dataurl', '.svg': 'dataurl', '.png': 'dataurl' },
});
// Excel wording for a few Univer menu labels (Paste special menu).
const LABELS = { 'Paste Value': 'Paste Values', 'Paste Format': 'Paste Formatting', 'Paste Column Width': 'Paste Column Widths', 'Paste Besides Border Styles': 'Paste All Except Borders' };
let bundle = readFileSync(`${js}/univer.bundle.js`, 'utf8');
for (const [from, to] of Object.entries(LABELS)) bundle = bundle.replace(`"${from}"`, `"${to}"`);
writeFileSync(`${js}/univer.bundle.js`, bundle);
renameSync(`${js}/univer.bundle.css`, `${css}/univer.bundle.css`);
copyFileSync('node_modules/exceljs/dist/exceljs.min.js', `${js}/exceljs.min.js`);
console.log('Vendor bundle built.');
