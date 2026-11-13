/* XD3 PDF Editor — service worker: after the first visit the editor keeps working with no connection, and it can be installed like an app. */
const VERSION = 'xd3-pdf-editor-v19'; // (raise on every release: it replaces the cached libraries and fonts)
// the page and its code, exactly as this build has it (tools/bundle-js.ps1 rewrites this list in the hosted ZIP, where the scripts are joined into js/bundle.js)
const CORE = ['./', 'index.html', 'manifest.webmanifest', 'css/style.css', 'css/flow.css', 'js/app.js', 'js/native.js', 'js/ocr.js', 'js/insert.js', 'js/imaging.js', 'js/convert.js', 'js/pdf2word.js', 'js/home.js', 'js/forms.js', 'js/extras.js', 'js/reader.js', 'js/writing.js', 'js/download.js', 'js/desktop.js', 'js/ai.js', 'js/ai-worker.js', 'js/docx.js', 'js/letterhead.js', 'js/tools.js', 'js/organize.js', 'js/flow.js', 'js/secure.js', 'js/ui.js',
  'lib/pdf.min.js', 'lib/pdf.worker.min.js', 'lib/pdf-lib.min.js', 'lib/fontkit.umd.min.js', 'lib/jszip.min.js', 'icons/icon-192.png', 'icons/icon-512.png', 'icons/logo.svg', 'icons/logo-small.svg', 'icons/apple-touch-icon.png'];
// the UI font and the free twins of the everyday Office fonts (Word to PDF keeps real, selectable text offline): 4.4 MB, kept from the first visit
const FONT_CORE = ['Inter', 'Arimo', 'Tinos', 'Cousine', 'Carlito', 'Caladea', 'Gelasio', 'DejaVuSans'].reduce((a, f) => a.concat(`lib/fonts/${f}-Regular.ttf`, `lib/fonts/${f}-Bold.ttf`), []);

self.addEventListener('install', e => { e.waitUntil(caches.open(VERSION).then(c => Promise.all(CORE.concat(FONT_CORE).map(u => c.add(u).catch(() => { /* a file that this build does not have */ })))).then(() => self.skipWaiting())); });
const AI_CACHE = 'xd3-ai-qwen2.5-0.5b-q4f16'; // the offline AI model (480 MB) is kept across app updates; a new model gets a new name
self.addEventListener('activate', e => { e.waitUntil(caches.keys().then(keys => Promise.all(keys.filter(k => k !== VERSION && k !== AI_CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim())); });

/* the rest of the fonts and the OCR engine are fetched quietly in the background once the editor has loaded (the page sends the list),
   one file at a time, so they are on this device before they are first needed offline; files already kept are skipped, so a later visit finishes the job */
let warming = null, queue = [];
const WARMABLE = /^lib\/(fonts|ocr)\/[\w.+-]+$/;
function warm(urls, first) { // first: a tool that needs these was just opened, so they go to the front of the queue
  const add = urls.filter(u => typeof u === 'string' && WARMABLE.test(u)), set = new Set(add);
  queue = first ? [...new Set(add)].concat(queue.filter(u => !set.has(u))) : queue.concat(add.filter(u => !queue.includes(u)));
  return warming || (warming = (async () => {
    const c = await caches.open(VERSION);
    let misses = 0;
    while (queue.length) {
      const u = queue.shift();
      try {
        if (await c.match(u)) continue;
        const r = await fetch(u, { cache: 'no-cache' }); if (r.ok) await c.put(u, r); else misses++;
      } catch (err) { if (++misses > 3) { queue = []; break; } } // (offline or the server is gone: try again on the next visit)
    }
  })().finally(() => { warming = null; }));
}
self.addEventListener('message', e => { const d = e.data; if (d && d.type === 'warm' && Array.isArray(d.urls)) e.waitUntil(warm(d.urls, d.first)); });

self.addEventListener('fetch', e => {
  const url = new URL(e.request.url);
  if (e.request.method !== 'GET' || url.origin !== location.origin || url.pathname.includes('__open__') || url.pathname.includes('/tests/') || url.pathname.includes('/downloads/')) return; // (the Windows installers are never cached)
  if (/\/lib\/ai\/models\//.test(url.pathname)) { // the AI model: from this device once it has been loaded the first time
    e.respondWith(caches.open(AI_CACHE).then(async c => (await c.match(e.request)) || fetch(e.request).then(r => { if (r.ok) c.put(e.request, r.clone()).catch(() => {}); return r; })));
    return;
  }
  if (/\/lib\//.test(url.pathname)) { // bundled libraries, fonts, OCR data: serve what we have, refresh quietly in the background
    e.respondWith(caches.open(VERSION).then(async c => {
      const hit = await c.match(e.request), net = fetch(e.request).then(r => { if (r.ok && r.status === 200) c.put(e.request, r.clone()).catch(() => {}); return r; }).catch(err => { if (hit) return hit; throw err; });
      if (hit) { e.waitUntil(net.catch(() => {})); return hit; }
      return net;
    }));
    return;
  }
  // the page and its own code: the network first, so an update shows at once; the cache when offline
  e.respondWith(fetch(new Request(e.request, { cache: 'no-cache' })).then(r => { if (r.ok) caches.open(VERSION).then(c => c.put(e.request, r.clone())); return r; }) // (always re-checked with the server, so an update is never missed)
    .catch(() => caches.match(e.request).then(hit => hit || (e.request.mode === 'navigate' ? caches.match('index.html') : undefined))));
});
