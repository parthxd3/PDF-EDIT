/* Helpers for checking "PDF → Word" by eye and by numbers (load in the editor page: a <script src="tests/convert_check.js"> tag).
   Flow: CC.convert(pdf, name) saves tests/out/name.docx → powershell -File tests\word_pdf.ps1 tests\out\name.docx (Word lays it out and
   exports name.word.pdf) → CC.positions(pdf, 'tests/out/name.word.pdf') compares where every piece of text ended up;
   CC.side / CC.zoom show original and Word's version next to each other. */
(() => {
const X = window.XD3, S = X.S;
const open = async url => pdfjsLib.getDocument({ url, isEvalSupported: false, cMapUrl: 'lib/cmaps/', cMapPacked: true, standardFontDataUrl: 'lib/standard_fonts/' }).promise;
const render = async (url, page = 1, k = 1) => {
  const pg = await (await open(url)).getPage(page), vp = pg.getViewport({ scale: k }), c = document.createElement('canvas'); c.width = vp.width; c.height = vp.height;
  const x = c.getContext('2d'); x.fillStyle = '#fff'; x.fillRect(0, 0, c.width, c.height); await pg.render({ canvasContext: x, viewport: vp, intent: 'print' }).promise; return c;
};
const show = c => {
  let img = document.getElementById('ccimg'); if (!img) { img = document.createElement('img'); img.id = 'ccimg'; img.onclick = () => img.remove(); document.body.appendChild(img); }
  img.style.cssText = 'position:fixed;left:0;top:0;z-index:99999;width:100vw;background:#888'; img.src = c.toDataURL(); return 'shown (click it to close)';
};
async function convert(pdf, name, mode = 'same') {
  const b = await (await fetch(pdf)).blob(); S.dirty = false; await X.openFiles([new File([b], pdf.split('/').pop(), { type: 'application/pdf' })], false); S.pages.forEach(p => X.visible.add(p.id));
  const t0 = performance.now(), blob = await X.buildDocx(S.pages, mode), ms = Math.round(performance.now() - t0), r = await fetch('tests/out/' + name + '.docx', { method: 'PUT', body: blob });
  return { pages: S.pages.length, size: blob.size, ms, saved: r.status === 204 };
}
async function positions(a, b, maxPages = 5) { // each piece of text in `a` → the nearest piece in `b` that starts the same way: how far it moved (points)
  const items = async url => { const doc = await open(url), out = []; for (let i = 1; i <= Math.min(doc.numPages, maxPages); i++) { const tc = await (await doc.getPage(i)).getTextContent(); for (const it of tc.items) if (it.str.trim().length > 3) out.push({ pg: i, s: it.str.trim(), x: it.transform[4], y: it.transform[5] }); } return out; };
  const A = await items(a), B = await items(b), dx = [], dy = [], off = [];
  for (const t of A) {
    const m = B.filter(u => u.pg === t.pg && (u.s.startsWith(t.s.slice(0, 12)) || t.s.startsWith(u.s.slice(0, 12)))).sort((p, q) => Math.hypot(p.x - t.x, p.y - t.y) - Math.hypot(q.x - t.x, q.y - t.y))[0];
    if (!m || Math.hypot(m.x - t.x, m.y - t.y) > 40) continue;
    dx.push(Math.abs(m.x - t.x)); dy.push(Math.abs(m.y - t.y)); if (Math.abs(m.x - t.x) > 1 || Math.abs(m.y - t.y) > 1) off.push([t.pg, t.s.slice(0, 24), +(m.x - t.x).toFixed(2), +(m.y - t.y).toFixed(2)]);
  }
  const st = v => (v.length ? { avg: +(v.reduce((s, x) => s + x, 0) / v.length).toFixed(2), max: +Math.max(...v).toFixed(2) } : null);
  return { pieces: A.length, matched: dx.length, dx: st(dx), dy: st(dy), off };
}
async function side(pairs, k = 0.5, page = 1) { // [[original, word], …] next to each other
  const cs = []; for (const [a, b] of pairs) cs.push(await render(a, page, k), await render(b, page, k));
  const out = document.createElement('canvas'); out.width = cs.reduce((s, c, i) => s + c.width + (i % 2 ? 24 : 4), 0); out.height = Math.max(...cs.map(c => c.height));
  const o = out.getContext('2d'); o.fillStyle = '#f0f'; o.fillRect(0, 0, out.width, out.height); let x = 0; cs.forEach((c, i) => { o.drawImage(c, x, 0); x += c.width + (i % 2 ? 24 : 4); });
  return show(out);
}
async function zoom(a, b, page, rect, k = 3) { // rect [x, y, w, h] in points: original left, Word right
  const A = await render(a, page, k), B = await render(b, page, k), [rx, ry, rw, rh] = rect.map(v => v * k), out = document.createElement('canvas'); out.width = rw * 2 + 8; out.height = rh;
  const o = out.getContext('2d'); o.fillStyle = '#f0f'; o.fillRect(0, 0, out.width, out.height); o.drawImage(A, rx, ry, rw, rh, 0, 0, rw, rh); o.drawImage(B, rx, ry, rw, rh, rw + 8, 0, rw, rh);
  return show(out);
}
async function text(url, page = 1) { const tc = await (await (await open(url)).getPage(page)).getTextContent(); return tc.items.map(i => i.str).join(' '); }
window.CC = { convert, positions, side, zoom, text, render };
})();
