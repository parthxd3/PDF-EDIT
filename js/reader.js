/* XD3 PDF Editor — reading mode: a clean viewer with selectable text, working links, bookmarks in the side panel, night colours,
   full screen, and a small pop-up to highlight, copy or comment on selected text. Esc (or the Read button) returns to editing. */
(() => {
'use strict';
const X = window.XD3, { S, $, $$, PE, toast, icons, uid } = X, body = document.body, viewer = X.viewer;
const store = { get: k => { try { return localStorage.getItem(k); } catch (e) { return null; } }, set: (k, v) => { try { localStorage.setItem(k, v); } catch (e) { /* not remembered */ } } };
const on = () => body.classList.contains('reader');

/* ───────── text + link layers (pdf.js), rendered once per page; zoom is handled by --scale-factor ───────── */
async function layers(p) {
  const pe = PE.get(p.id); if (!pe || p.src < 0 || !on()) return;
  const key = 'r' + p.rot; if (pe.textKey === key) return; pe.textKey = key;
  $$('.textLayer, .linkLayer', pe.wrap).forEach(e => e.remove());
  try {
    const pg = await X.getPdfPage(p), sc = +pe.wrap.style.getPropertyValue('--scale-factor') || 1, vp = pg.getViewport({ scale: sc, rotation: (pg.rotate + p.rot) % 360 }), [rw, rh] = X.rdims(p).map(v => v * sc); // (pdf.js wants the layer rendered at the current zoom; later zooms only change --scale-factor)
    if (PE.get(p.id) !== pe || pe.textKey !== key) return;
    const tl = document.createElement('div'); tl.className = 'textLayer'; pe.wrap.insertBefore(tl, pe.svg);
    await pdfjsLib.renderTextLayer({ textContentSource: await pg.getTextContent(), container: tl, viewport: vp, textDivs: [] }).promise;
    const end = document.createElement('div'); end.className = 'endOfContent'; tl.appendChild(end);
    tl.addEventListener('mousedown', () => tl.classList.add('selecting')); tl.addEventListener('mouseup', () => tl.classList.remove('selecting'));
    const ll = document.createElement('div'); ll.className = 'linkLayer'; pe.wrap.insertBefore(ll, pe.svg);
    for (const a of await pg.getAnnotations()) {
      if (a.subtype !== 'Link' || !a.rect || (!a.url && !a.dest)) continue;
      const r = vp.convertToViewportRectangle(a.rect), x = Math.min(r[0], r[2]), y = Math.min(r[1], r[3]), el = document.createElement('a');
      el.className = 'plink'; el.style.cssText = `left:${x / rw * 100}%;top:${y / rh * 100}%;width:${Math.abs(r[2] - r[0]) / rw * 100}%;height:${Math.abs(r[3] - r[1]) / rh * 100}%`;
      el.title = a.url || 'Go to the linked page'; el.href = a.url || '#';
      el.addEventListener('click', async e => {
        e.preventDefault();
        if (a.url) { if (/^https?:|^mailto:/i.test(a.url)) window.open(a.url, '_blank', 'noopener'); return; }
        try { let d = a.dest; if (typeof d === 'string') d = await S.sources[p.src].pdf.getDestination(d); const idx = Array.isArray(d) && d[0] != null ? (typeof d[0] === 'object' ? await S.sources[p.src].pdf.getPageIndex(d[0]) : d[0]) : -1;
          const target = S.pages.findIndex(q => q.src === p.src && q.idx === idx); if (target >= 0) X.goToPage(target); else toast('The linked page is not in this document'); } catch (err) { toast('This link could not be followed'); }
      });
      ll.appendChild(el);
    }
  } catch (e) { console.warn('reader layers', e); }
}
X.onPageVisible = p => { if (on()) layers(p); };

/* ───────── the mode itself ───────── */
function setReader(v) {
  body.classList.toggle('reader', v); store.set('xd3.reader', v ? '1' : '0');
  $$('[data-act=reader]').forEach(b => b.classList.toggle('active', v));
  if (v) { X.commitEdit(); X.setTool('hand'); X.select(null); X.visible.forEach(id => { const p = S.pages.find(q => q.id === id); if (p) layers(p); }); viewer.focus(); $('#hint').textContent = 'Reading mode — select text to highlight or copy it · click links · ← → change page · Esc returns to editing'; }
  else { hidePop(); X.setTool('select'); if (document.fullscreenElement) document.exitFullscreen().catch(() => {}); }
}
X.ACT.reader = () => { if (S.pages.length) setReader(!on()); };
X.ACT.night = () => { const v = !body.classList.contains('night'); body.classList.toggle('night', v); store.set('xd3.night', v ? '1' : '0'); $$('[data-act=night]').forEach(b => b.classList.toggle('active', v)); };
X.ACT.fullscreen = () => { if (document.fullscreenElement) document.exitFullscreen().catch(() => {}); else document.documentElement.requestFullscreen().catch(() => toast('Full screen is not available here')); };
X.NO_DOC.add('night'); X.NO_DOC.add('fullscreen');
if (store.get('xd3.night') === '1') X.ACT.night();

/* ───────── selection pop-up: highlight · comment · copy ───────── */
const pop = document.createElement('div'); pop.id = 'selpop'; pop.hidden = true;
pop.innerHTML = '<button data-p="hl" title="Highlight the selected text"><i data-icon="hl"></i>Highlight</button><button data-p="note" title="Attach a sticky note"><i data-icon="note"></i>Note</button><button data-p="copy" title="Copy (Ctrl+C)"><i data-icon="copy"></i>Copy</button>';
document.body.appendChild(pop); icons(pop);
const hidePop = () => { pop.hidden = true; };
function selRects() { // the selected text as rectangles in page coordinates, grouped per page and merged line by line
  const sel = getSelection(); if (!sel || sel.isCollapsed || !sel.rangeCount) return null;
  const range = sel.getRangeAt(0), tl = range.commonAncestorContainer.parentElement && range.commonAncestorContainer.parentElement.closest('.textLayer, .page'); if (!tl) return null;
  const byPage = new Map();
  for (const r of range.getClientRects()) {
    if (r.width < 1 || r.height < 1) continue; const cx = r.left + r.width / 2, cy = r.top + r.height / 2; let wrap = null;
    for (const [, q] of PE) { const b = q.wrap.getBoundingClientRect(); if (cx >= b.left && cx <= b.right && cy >= b.top && cy <= b.bottom) { wrap = q.wrap; break; } } if (!wrap) continue;
    const pe = PE.get(wrap.dataset.id); if (!pe) continue; const M = pe.g.getScreenCTM().inverse(), a = new DOMPoint(r.left, r.top).matrixTransform(M), b = new DOMPoint(r.right, r.bottom).matrixTransform(M);
    const box = { x: Math.min(a.x, b.x), y: Math.min(a.y, b.y), w: Math.abs(b.x - a.x), h: Math.abs(b.y - a.y) }, list = byPage.get(wrap.dataset.id) || [];
    const last = list[list.length - 1];
    if (last && Math.abs(last.y - box.y) < box.h * 0.5 && box.x <= last.x + last.w + box.h) { const x1 = Math.max(last.x + last.w, box.x + box.w), y1 = Math.max(last.y + last.h, box.y + box.h); last.x = Math.min(last.x, box.x); last.y = Math.min(last.y, box.y); last.w = x1 - last.x; last.h = y1 - last.y; }
    else list.push(box);
    byPage.set(wrap.dataset.id, list);
  }
  return byPage.size ? { byPage, text: sel.toString(), rect: range.getBoundingClientRect() } : null;
}
function showPop() {
  if (!on()) return hidePop(); const s = selRects(); if (!s) return hidePop();
  pop.hidden = false; const w = pop.offsetWidth, x = Math.max(6, Math.min(innerWidth - w - 6, s.rect.left + s.rect.width / 2 - w / 2)), y = s.rect.top - pop.offsetHeight - 8;
  pop.style.left = x + 'px'; pop.style.top = (y > 50 ? y : s.rect.bottom + 8) + 'px';
}
document.addEventListener('selectionchange', () => { clearTimeout(showPop.t); showPop.t = setTimeout(showPop, 120); });
pop.addEventListener('mousedown', e => e.preventDefault()); // keep the selection
pop.addEventListener('click', e => {
  const b = e.target.closest('[data-p]'); if (!b) return; const s = selRects(); if (!s) return hidePop();
  if (b.dataset.p === 'copy') { (navigator.clipboard && navigator.clipboard.writeText ? navigator.clipboard.writeText(s.text) : Promise.reject()).then(() => toast('Copied'), () => document.execCommand('copy')); hidePop(); return; }
  X.pushUndo();
  for (const [pid, boxes] of s.byPage) {
    const p = S.pages.find(q => q.id === pid); if (!p) continue;
    if (b.dataset.p === 'hl') for (const r of boxes) p.objs.push({ id: uid(), type: 'rect', x: r.x - 1, y: r.y - 1, w: r.w + 2, h: r.h + 2, color: 'none', fill: '#facc15', width: 0, opacity: 0.4, blend: true });
    else { const last = boxes[boxes.length - 1], o = X.textObj('', 12, { x: last.x + last.w + 6, y: last.y, color: '#1f2937', bg: '#fef08a', note: true }); p.objs.push(o); X.drawObjs(p); getSelection().removeAllRanges(); hidePop(); X.startEdit(p, o, X.snap()); return; }
    X.drawObjs(p);
  }
  getSelection().removeAllRanges(); hidePop(); toast('Highlighted — it is saved as a PDF comment');
});

/* ───────── keys and bookmarks in the side panel ───────── */
window.addEventListener('keydown', e => {
  if (!on() || $('#dlg').open || /INPUT|TEXTAREA|SELECT/.test(e.target.tagName) || e.target.isContentEditable) return;
  const k = e.key;
  if (k === 'Escape' && !S.editing) { e.stopImmediatePropagation(); setReader(false); }
  else if (k === 'ArrowRight' || k === 'PageDown' && e.shiftKey) { e.preventDefault(); X.goToPage(S.cur + 1); }
  else if (k === 'ArrowLeft' || k === 'PageUp' && e.shiftKey) { e.preventDefault(); X.goToPage(S.cur - 1); }
  else if (k === 'Home') { e.preventDefault(); X.goToPage(0); } else if (k === 'End') { e.preventDefault(); X.goToPage(S.pages.length - 1); }
}, true);
const panel = document.createElement('div'); panel.id = 'bmPanel'; panel.hidden = true; $('#thumbs').parentNode.insertBefore(panel, $('#thumbs').nextSibling);
function renderBookmarks() {
  const list = (S.outline || []).filter(e => S.pages.some(p => p.id === e.pid));
  panel.innerHTML = `<div class="sbtabs"><button data-sb="pages">Pages</button><button data-sb="bm" class="on">Bookmarks</button></div>` + (list.length
    ? list.map((e, i) => `<button class="bmitem" data-i="${i}" style="padding-left:${10 + e.level * 14}px" title="Page ${S.pages.findIndex(p => p.id === e.pid) + 1}">${X.esc(e.title)}</button>`).join('')
    : '<div class="empty">No bookmarks in this document.</div>') + '<button class="btn line" data-act="bookmarks" style="margin:10px"><i data-icon="bookmark"></i>Edit bookmarks…</button>';
  icons(panel);
}
X.renderBookmarks = renderBookmarks;
function showPanel(which) { const bm = which === 'bm'; panel.hidden = !bm; $('#thumbs').classList.toggle('offscreen', bm); if (bm) renderBookmarks(); $$('[data-sb]').forEach(b => b.classList.toggle('on', b.dataset.sb === which)); }
panel.addEventListener('click', e => { const it = e.target.closest('.bmitem'); if (it) { const en = (S.outline || []).filter(x => S.pages.some(p => p.id === x.pid))[+it.dataset.i]; if (en) { X.goToPage(S.pages.findIndex(p => p.id === en.pid)); body.classList.remove('sbopen'); } return; } const t = e.target.closest('[data-sb]'); if (t) showPanel(t.dataset.sb); });
const tabs = document.createElement('div'); tabs.className = 'sbtabs'; tabs.innerHTML = '<button data-sb="pages" class="on">Pages</button><button data-sb="bm">Bookmarks</button>';
$('#thumbs').insertBefore(tabs, $('#thumbs').firstChild); tabs.addEventListener('click', e => { const t = e.target.closest('[data-sb]'); if (t) showPanel(t.dataset.sb); });
X.showSidePanel = showPanel;

if (store.get('xd3.reader') === '1') X.onDocOpened = () => setReader(true); // the last mode is remembered: readers land in reading mode
})();
