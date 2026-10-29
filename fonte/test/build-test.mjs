import { build } from 'esbuild';
import { readFileSync, writeFileSync, mkdirSync, cpSync, readdirSync, existsSync } from 'node:fs';
// Junta main.js com os módulos de src/partes (inseridos antes da inicialização do app).
const _main = readFileSync('src/main.js', 'utf8'); const _k = _main.lastIndexOf("if ('serviceWorker' in navigator && !EMU)");
const _partes = existsSync('src/partes') ? readdirSync('src/partes').filter(f => f.endsWith('.js')).sort().map(f => `/* ---- partes/${f} ---- */\n` + readFileSync('src/partes/' + f, 'utf8')).join('\n') : '';
writeFileSync('src/_app.js', _main.slice(0, _k) + _partes + '\n' + _main.slice(_k));
mkdirSync('test/www/icons', { recursive: true });
const fake = { name: 'fake', setup(b) { b.onResolve({ filter: /^firebase\/(app|auth|firestore)$/ }, () => ({ path: new URL('./fake.js', import.meta.url).pathname })); } };
await build({ entryPoints: ['src/_app.js'], bundle: true, format: 'iife', outfile: 'test/www/app.js', plugins: [fake], logLevel: 'warning' });
writeFileSync('test/www/app.js', readFileSync('test/www/app.js', 'utf8').replaceAll('__VERSAO__', 'teste'));
writeFileSync('test/www/index.html', readFileSync('src/index.html', 'utf8').replaceAll('__VERSAO__', 'teste'));
for (const n of ['xlsx', 'chart', 'pdf']) await build({ entryPoints: ['src/libs/' + n + '.js'], bundle: true, format: 'iife', outfile: 'test/www/lib-' + n + '.js', logLevel: 'warning' });
await build({ entryPoints: ['src/leitor.js'], bundle: true, format: 'iife', outfile: 'test/www/leitor.js', logLevel: 'warning' });
cpSync('node_modules/pdfjs-dist/legacy/build/pdf.worker.min.mjs', 'test/www/pdf.worker.min.mjs');
{ const cfgTxt = readFileSync('src/config.js', 'utf8'); const pega = n => (cfgTxt.match(new RegExp(n + ':\\s*"([^"]+)"')) || [])[1] || ''; writeFileSync('test/www/cliente.html', readFileSync('src/cliente.html', 'utf8').replaceAll('__PROJETO__', pega('projectId')).replaceAll('__CHAVE__', pega('apiKey'))); }
cpSync('src/icons', 'test/www/icons', { recursive: true });
cpSync('src/fonts', 'test/www/fonts', { recursive: true });
console.log('ok');
