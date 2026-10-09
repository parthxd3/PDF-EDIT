/* XD3 PDF Editor — service worker: after the first visit the editor keeps working with no connection, and it can be installed like an app. */
const VERSION = 'xd3-pdf-editor-v3';
const CORE = ['./', 'index.html', 'manifest.webmanifest', 'css/style.css', 'js/app.js', 'js/ocr.js', 'js/insert.js', 'js/imaging.js', 'js/convert.js', 'js/forms.js', 'js/extras.js', 'js/reader.js', 'js/organize.js', 'js/secure.js', 'js/ui.js', 'js/bundle.js',
  'lib/pdf.min.js', 'lib/pdf.worker.min.js', 'lib/pdf-lib.min.js', 'lib/fontkit.umd.min.js', 'lib/jszip.min.js', 'icons/icon-192.png', 'icons/icon-512.png'];

self.addEventListener('install', e => { e.waitUntil(caches.open(VERSION).then(c => Promise.all(CORE.map(u => c.add(u).catch(() => { /* a file that this build does not have */ })))).then(() => self.skipWaiting())); });
self.addEventListener('activate', e => { e.waitUntil(caches.keys().then(keys => Promise.all(keys.filter(k => k !== VERSION).map(k => caches.delete(k)))).then(() => self.clients.claim())); });
self.addEventListener('fetch', e => {
  const url = new URL(e.request.url);
  if (e.request.method !== 'GET' || url.origin !== location.origin || url.pathname.includes('__open__') || url.pathname.includes('/tests/')) return;
  if (/\/lib\//.test(url.pathname)) { // bundled libraries, fonts, OCR data: serve what we have, refresh quietly in the background
    e.respondWith(caches.open(VERSION).then(async c => {
      const hit = await c.match(e.request), net = fetch(e.request).then(r => { if (r.ok) c.put(e.request, r.clone()); return r; }).catch(() => hit);
      return hit || net;
    }));
    return;
  }
  // the page and its own code: the network first, so an update shows at once; the cache when offline
  e.respondWith(fetch(new Request(e.request, { cache: 'no-cache' })).then(r => { if (r.ok) caches.open(VERSION).then(c => c.put(e.request, r.clone())); return r; }) // (always re-checked with the server, so an update is never missed)
    .catch(() => caches.match(e.request).then(hit => hit || (e.request.mode === 'navigate' ? caches.match('index.html') : undefined))));
});
