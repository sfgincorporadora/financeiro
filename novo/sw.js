// Service worker: guarda o app no aparelho para abrir sem internet.
const VERSAO = '202610292253';
const CACHE = 'sfg-' + VERSAO;
const ARQUIVOS = ['./', './index.html', './app.js?v=' + VERSAO, './manifest.webmanifest', './icons/icon-192.png', './icons/icon-512.png', './icons/apple-touch-icon.png', './icons/logo-branco.svg', './icons/logo-azul.svg', './icons/logo-azul-incorporadora.svg', './icons/logo-branco-incorporadora.svg', './fonts/archivo-var.woff2', './lib-xlsx.js?v=' + VERSAO, './lib-chart.js?v=' + VERSAO, './lib-pdf.js?v=' + VERSAO];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(ARQUIVOS)).then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k.startsWith('sfg-') && k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});
self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);
  if (url.origin !== location.origin) return; // Firebase e Google passam direto
  if (url.pathname.includes('/novo/') && !self.registration.scope.includes('/novo/')) return; // prévia tem o próprio cache
  if (req.mode === 'navigate') {
    // página: tenta a rede (para pegar versão nova), cai no cache se offline
    e.respondWith(fetch(req).then(r => { const c = r.clone(); caches.open(CACHE).then(x => x.put('./index.html', c)); return r; }).catch(() => caches.match('./index.html').then(r => r || caches.match('./'))));
    return;
  }
  e.respondWith(caches.match(req).then(r => r || fetch(req).then(res => { if (res.ok) { const c = res.clone(); caches.open(CACHE).then(x => x.put(req, c)); } return res; })));
});
