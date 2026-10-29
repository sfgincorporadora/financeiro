// Gera o app publicado na raiz do repositório (GitHub Pages). Rode de dentro de fonte/: npm install && node build.mjs
const OUT = process.env.OUT || '..';
import { build } from 'esbuild';
import { readFileSync, writeFileSync, mkdirSync, cpSync } from 'node:fs';

const versao = new Date().toISOString().slice(0, 16).replace(/[-:T]/g, '');
mkdirSync(OUT + '/icons', { recursive: true });
await build({
  entryPoints: ['src/main.js'], bundle: true, minify: true, format: 'iife', target: ['safari15', 'chrome100'],
  outfile: OUT + '/app.js', define: { __VERSAO__: JSON.stringify(versao) }, legalComments: 'none', logLevel: 'warning'
});
writeFileSync(OUT + '/app.js', readFileSync(OUT + '/app.js', 'utf8').replaceAll('__VERSAO__', versao));
writeFileSync(OUT + '/index.html', readFileSync('src/index.html', 'utf8').replaceAll('__VERSAO__', versao));
writeFileSync(OUT + '/sw.js', readFileSync('src/sw.js', 'utf8').replaceAll('__VERSAO__', versao));
cpSync('src/icons', OUT + '/icons', { recursive: true });
cpSync('src/fonts', OUT + '/fonts', { recursive: true });
writeFileSync(OUT + '/manifest.webmanifest', JSON.stringify({
  name: 'SFG Financeiro', short_name: 'SFG', start_url: './', scope: './', display: 'standalone',
  background_color: '#0B2046', theme_color: '#0B2046', lang: 'pt-BR',
  icons: [{ src: 'icons/icon-192.png', sizes: '192x192', type: 'image/png' }, { src: 'icons/icon-512.png', sizes: '512x512', type: 'image/png' }, { src: 'icons/icon-512.png', sizes: '512x512', type: 'image/png', purpose: 'maskable' }]
}, null, 1));
writeFileSync(OUT + '/.nojekyll', '');
console.log('build', versao);
