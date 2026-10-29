// Gera o app publicado na raiz do repositório (GitHub Pages). Rode de dentro de fonte/: npm install && node build.mjs
const OUT = process.env.OUT || '..';
import { build } from 'esbuild';
import { readFileSync, writeFileSync, mkdirSync, cpSync, readdirSync, existsSync } from 'node:fs';
// Junta main.js com os módulos de src/partes (inseridos antes da inicialização do app).
const _main = readFileSync('src/main.js', 'utf8'); const _k = _main.lastIndexOf("if ('serviceWorker' in navigator && !EMU)");
const _partes = existsSync('src/partes') ? readdirSync('src/partes').filter(f => f.endsWith('.js')).sort().map(f => `/* ---- partes/${f} ---- */\n` + readFileSync('src/partes/' + f, 'utf8')).join('\n') : '';
writeFileSync('src/_app.js', _main.slice(0, _k) + _partes + '\n' + _main.slice(_k));

const versao = new Date().toISOString().slice(0, 16).replace(/[-:T]/g, '');
mkdirSync(OUT + '/icons', { recursive: true });
await build({
  entryPoints: ['src/_app.js'], bundle: true, minify: true, format: 'iife', target: ['safari15', 'chrome100'],
  outfile: OUT + '/app.js', define: { __VERSAO__: JSON.stringify(versao) }, legalComments: 'none', logLevel: 'warning'
});
writeFileSync(OUT + '/app.js', readFileSync(OUT + '/app.js', 'utf8').replaceAll('__VERSAO__', versao));
writeFileSync(OUT + '/index.html', readFileSync('src/index.html', 'utf8').replaceAll('__VERSAO__', versao));
writeFileSync(OUT + '/sw.js', readFileSync('src/sw.js', 'utf8').replaceAll('__VERSAO__', versao));
for (const n of ['xlsx', 'chart', 'pdf']) await build({ entryPoints: ['src/libs/' + n + '.js'], bundle: true, format: 'iife', minify: true, target: ['safari15', 'chrome100'], legalComments: 'none', outfile: OUT + '/lib-' + n + '.js', logLevel: 'warning' });
await build({ entryPoints: ['src/leitor.js'], bundle: true, format: 'iife', minify: true, target: ['safari15', 'chrome100'], legalComments: 'none', outfile: OUT + '/leitor.js', logLevel: 'warning' });
cpSync('node_modules/pdfjs-dist/legacy/build/pdf.worker.min.mjs', OUT + '/pdf.worker.min.mjs');
{ const cfgTxt = readFileSync('src/config.js', 'utf8'); const pega = n => (cfgTxt.match(new RegExp(n + ':\\s*"([^"]+)"')) || [])[1] || ''; writeFileSync(OUT + '/cliente.html', readFileSync('src/cliente.html', 'utf8').replaceAll('__PROJETO__', pega('projectId')).replaceAll('__CHAVE__', pega('apiKey'))); }
cpSync('src/icons', OUT + '/icons', { recursive: true });
cpSync('src/fonts', OUT + '/fonts', { recursive: true });
writeFileSync(OUT + '/manifest.webmanifest', JSON.stringify({
  name: 'SFG Financeiro', short_name: 'SFG', start_url: './', scope: './', display: 'standalone',
  background_color: '#FFFFFF', theme_color: '#0B2046', lang: 'pt-BR',
  icons: [{ src: 'icons/icon-192.png', sizes: '192x192', type: 'image/png' }, { src: 'icons/icon-512.png', sizes: '512x512', type: 'image/png' }, { src: 'icons/icon-512.png', sizes: '512x512', type: 'image/png', purpose: 'maskable' }]
}, null, 1));
writeFileSync(OUT + '/.nojekyll', '');
console.log('build', versao);
