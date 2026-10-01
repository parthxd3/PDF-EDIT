/* XD3 PDF Editor — page organizer (multi-select reorder / delete / replace / extract) and page resizing. */
(() => {
'use strict';
const X = window.XD3, { S, $, $$, toast, busy, dialog, uid, clone, download } = X;
const org = $('#org'), grid = $('#orgGrid'), sel = new Set();
let lastIdx = -1, tio = null, dragging = false;

const idxs = () => S.pages.map((p, i) => sel.has(p.id) ? i : -1).filter(i => i >= 0);
const need = () => { if (!sel.size) { toast('Select one or more pages first (click them)'); return false; } return true; };
function pick(accept, multiple) { // → Promise<File[]>
  return new Promise(res => { const i = document.createElement('input'); i.type = 'file'; i.accept = accept; i.multiple = !!multiple; i.onchange = () => res([...i.files]); i.click(); });
}
async function pagesFrom(files) { // PDF / Word / image files → page records (sources are registered as needed)
  const out = [];
  for (const f of files) {
    if (X.isPdf(f)) out.push(...await X.loadSource(new Uint8Array(await f.arrayBuffer())));
    else if (X.isDocx(f)) out.push(...await X.docxToPages(f));
    else if (f.type.startsWith('image/')) out.push(X.imagePage(await X.readImage(f)));
  }
  if (!out.length) throw new Error('No pages found in the selected file(s)');
  return out;
}
const ACCEPT = 'application/pdf,image/*,.docx';

function render() {
  if (tio) tio.disconnect();
  tio = new IntersectionObserver(ens => ens.forEach(async en => {
    if (!en.isIntersecting) return; tio.unobserve(en.target); const p = S.pages.find(q => q.id === en.target.dataset.id); if (!p) return;
    try { $('.othumb', en.target).replaceChildren(await X.renderPageImage(p, 220 / X.rdims(p)[0])); } catch (e) { /* page re-rendered meanwhile */ }
  }), { root: grid, rootMargin: '300px 0px' });
  for (const id of [...sel]) if (!S.pages.some(p => p.id === id)) sel.delete(id);
  grid.textContent = '';
  S.pages.forEach((p, i) => {
    const [rw, rh] = X.rdims(p), c = document.createElement('div'); c.className = 'ocard' + (sel.has(p.id) ? ' sel' : ''); c.dataset.id = p.id; c.draggable = true;
    c.innerHTML = `<div class="othumb" style="aspect-ratio:${rw}/${rh}"></div><div class="onum">${i + 1}</div><div class="ocheck">✓</div>`;
    grid.appendChild(c); tio.observe(c);
  });
  info();
}
const info = () => { $('#orgInfo').textContent = `${S.pages.length} page(s)${sel.size ? ` · ${sel.size} selected` : ' · click pages to select, drag to reorder, double-click to open'}`; $$('.ocard', grid).forEach(c => c.classList.toggle('sel', sel.has(c.dataset.id))); };
function apply(fn) { X.pushUndo(); fn(); const cur = S.cur; X.buildPages(); S.cur = Math.min(cur, S.pages.length - 1); render(); }
function open() { if (!S.pages.length) return; sel.clear(); lastIdx = -1; org.hidden = false; document.body.classList.add('modal'); render(); }
function close(goTo) { org.hidden = true; document.body.classList.remove('modal'); if (tio) tio.disconnect(); grid.textContent = ''; X.goToPage(goTo == null ? S.cur : goTo); }

grid.addEventListener('click', e => {
  const c = e.target.closest('.ocard'); if (!c) return; const i = S.pages.findIndex(p => p.id === c.dataset.id);
  if (e.shiftKey && lastIdx >= 0) { for (let k = Math.min(i, lastIdx); k <= Math.max(i, lastIdx); k++) sel.add(S.pages[k].id); }
  else if (sel.has(c.dataset.id)) sel.delete(c.dataset.id); else sel.add(c.dataset.id);
  lastIdx = i; info();
});
grid.addEventListener('dblclick', e => { const c = e.target.closest('.ocard'); if (c) close(S.pages.findIndex(p => p.id === c.dataset.id)); });
grid.addEventListener('dragstart', e => { const c = e.target.closest('.ocard'); if (!c) return; if (!sel.has(c.dataset.id)) { sel.clear(); sel.add(c.dataset.id); info(); } dragging = true; e.dataTransfer.effectAllowed = 'move'; e.dataTransfer.setData('text/plain', 'pages'); });
grid.addEventListener('dragover', e => { if (!dragging) return; e.preventDefault(); $$('.dropt', grid).forEach(x => x.classList.remove('dropt')); const c = e.target.closest('.ocard'); if (c && !sel.has(c.dataset.id)) c.classList.add('dropt'); });
grid.addEventListener('dragend', () => { dragging = false; $$('.dropt', grid).forEach(x => x.classList.remove('dropt')); });
grid.addEventListener('drop', e => {
  if (!dragging) return; e.preventDefault(); e.stopPropagation(); dragging = false;
  const c = e.target.closest('.ocard'); if (!c || sel.has(c.dataset.id)) return;
  const b = c.getBoundingClientRect(), after = e.clientX > b.left + b.width / 2;
  apply(() => { const moving = S.pages.filter(p => sel.has(p.id)), rest = S.pages.filter(p => !sel.has(p.id)), at = rest.findIndex(p => p.id === c.dataset.id) + (after ? 1 : 0); rest.splice(at, 0, ...moving); S.pages.splice(0, S.pages.length, ...rest); });
});

const OPS = {
  all: () => { S.pages.forEach(p => sel.add(p.id)); info(); },
  none: () => { sel.clear(); info(); },
  rotl: () => need() && apply(() => idxs().forEach(i => { S.pages[i].rot = (S.pages[i].rot + 270) % 360; })),
  rotr: () => need() && apply(() => idxs().forEach(i => { S.pages[i].rot = (S.pages[i].rot + 90) % 360; })),
  dup: () => need() && apply(() => idxs().reverse().forEach(i => { const n = clone(S.pages[i]); n.id = uid(); n.objs.forEach(o => o.id = uid()); S.pages.splice(i + 1, 0, n); })),
  del: () => { if (!need()) return; if (sel.size >= S.pages.length) return toast('A document needs at least one page', true); const n = sel.size; apply(() => { S.pages.splice(0, S.pages.length, ...S.pages.filter(p => !sel.has(p.id))); }); toast(`${n} page(s) deleted — Ctrl+Z to undo`); },
  left: () => need() && idxs()[0] > 0 && apply(() => idxs().forEach(i => { const [p] = S.pages.splice(i, 1); S.pages.splice(i - 1, 0, p); })),
  right: () => need() && idxs().pop() < S.pages.length - 1 && apply(() => idxs().reverse().forEach(i => { const [p] = S.pages.splice(i, 1); S.pages.splice(i + 1, 0, p); })),
  reverse: () => apply(() => S.pages.reverse()),
  blank: () => apply(() => { const at = sel.size ? idxs().pop() + 1 : S.pages.length, ref = S.pages[Math.max(0, at - 1)], [w, h] = ref ? X.rdims(ref) : X.A4; S.pages.splice(at, 0, { id: uid(), src: -1, w, h, rot: 0, objs: [] }); }),
  extract: () => need() && busy(async () => download(new Blob([await X.buildPdf(idxs().map(i => S.pages[i]))], { type: 'application/pdf' }), `${X.baseName()}-pages.pdf`), 'Extracting…'),
  add: async () => { const files = await pick(ACCEPT, true); if (!files.length) return; await busy(async () => { const np = await pagesFrom(files); apply(() => S.pages.splice(sel.size ? idxs().pop() + 1 : S.pages.length, 0, ...np)); toast(`${np.length} page(s) added`); }, 'Adding pages…'); },
  replace: async () => { if (need()) await replacePages(idxs()); },
  resize: () => resizeDlg(sel.size ? idxs() : null),
  numbers: async () => { await X.pageNumDlg(); render(); },
  done: () => close(),
};
org.querySelector('header').addEventListener('click', e => { const b = e.target.closest('[data-o]'); if (b && OPS[b.dataset.o]) OPS[b.dataset.o](); });
window.addEventListener('keydown', e => {
  if (org.hidden || $('#dlg').open) return; const k = e.key.toLowerCase();
  if (k === 'escape') close(); else if (k === 'delete' || k === 'backspace') OPS.del();
  else if ((e.ctrlKey || e.metaKey) && k === 'a') { e.preventDefault(); OPS.all(); }
  else if ((e.ctrlKey || e.metaKey) && k === 'z') { e.preventDefault(); $('#bUndo').click(); render(); }
});

async function replacePages(ix) { // swap the given pages for pages from another file
  const files = await pick(ACCEPT, true); if (!files.length) return;
  await busy(async () => {
    const np = await pagesFrom(files), first = ix[0], drop = new Set(ix.map(i => S.pages[i].id));
    X.pushUndo(); const rest = S.pages.filter(p => !drop.has(p.id)); rest.splice(Math.min(first, rest.length), 0, ...np); S.pages.splice(0, S.pages.length, ...rest);
    X.buildPages(); if (!org.hidden) render(); else X.goToPage(first);
    toast(`Replaced ${ix.length} page(s) with ${np.length} new page(s)`);
  }, 'Replacing…');
}

/* ───────── change page size ───────── */
const SIZES = { A4: [595.28, 841.89], Letter: [612, 792], Legal: [612, 1008], A3: [841.89, 1190.55], A5: [419.53, 595.28] };
function mapObj(o, p, s, ox, oy) { // carry an edit from the old page (base space, extra rotation p.rot) onto the resized page
  const map = ([x, y]) => { const q = p.rot === 90 ? [p.h - y, x] : p.rot === 180 ? [p.w - x, p.h - y] : p.rot === 270 ? [y, p.w - x] : [x, y]; return [q[0] * s + ox, q[1] * s + oy]; };
  const n = clone(o);
  if (X.isLine(n)) { [n.x1, n.y1] = map([o.x1, o.y1]); [n.x2, n.y2] = map([o.x2, o.y2]); n.width *= s; return n; }
  const c0 = X.center(X.bbox(o)), c1 = map(c0);
  if (n.type === 'path') { n.pts = o.pts.map(([x, y]) => [c1[0] + (x - c0[0]) * s, c1[1] + (y - c0[1]) * s]); n.width *= s; }
  else if (n.type === 'text') { n.size = +(o.size * s).toFixed(2); n.x = 0; n.y = 0; const b = X.bbox(n); n.x = c1[0] - b.w / 2; n.y = c1[1] - b.h / 2; }
  else { n.w = o.w * s; n.h = o.h * s; n.x = c1[0] - n.w / 2; n.y = c1[1] - n.h / 2; if (n.width) n.width *= s; if (n.rx) n.rx *= s; }
  n.rot = ((o.rot || 0) + p.rot) % 360; return n;
}
async function resizePages(ix, W, H, orient) {
  await busy(async () => {
    const L = PDFLib, out = await L.PDFDocument.create(), docs = {}, plan = [];
    for (const i of ix) {
      const p = S.pages[i], [rw, rh] = X.rdims(p), land = orient === 'auto' ? rw > rh : orient === 'land', PW = land ? Math.max(W, H) : Math.min(W, H), PH = land ? Math.min(W, H) : Math.max(W, H);
      const s = Math.min(PW / rw, PH / rh), ox = (PW - rw * s) / 2, oy = (PH - rh * s) / 2, page = out.addPage([PW, PH]);
      if (p.src >= 0) {
        if (!docs[p.src]) { try { docs[p.src] = await L.PDFDocument.load(S.sources[p.src].bytes); } catch (e) { throw new Error('This PDF is password-protected and its pages cannot be resized'); } }
        const pg = await X.getPdfPage(p), [x1, y1, x2, y2] = pg.view, R = (pg.rotate + p.rot) % 360, X0 = ox, Y0 = PH - oy - rh * s;
        const emb = await out.embedPage(docs[p.src].getPage(p.idx), { left: x1, bottom: y1, right: x2, top: y2 });
        const at = R === 90 ? [X0, Y0 + rh * s] : R === 180 ? [X0 + rw * s, Y0 + rh * s] : R === 270 ? [X0 + rw * s, Y0] : [X0, Y0]; // where the page's own origin lands once turned upright
        page.drawPage(emb, { x: at[0], y: at[1], xScale: s, yScale: s, rotate: L.degrees(-R) });
      }
      plan.push({ i, PW, PH, objs: p.objs.map(o => mapObj(o, p, s, ox, oy)) });
    }
    const si = await X.addSource(await out.save());
    X.pushUndo(); plan.forEach((pl, k) => { S.pages[pl.i] = { id: uid(), src: si, idx: k, w: pl.PW, h: pl.PH, rot: 0, objs: pl.objs }; });
    X.buildPages(); if (!org.hidden) render(); toast(`${ix.length} page(s) resized`);
  }, 'Resizing pages…');
}
async function resizeDlg(ix) {
  const r = await dialog('Change page size', `
    <div class="row"><label>Size<select name="size">${Object.keys(SIZES).map(k => `<option>${k}</option>`).join('')}<option value="custom">Custom…</option></select></label>
    <label>Orientation<select name="orient"><option value="auto">Keep each page's</option><option value="port">Portrait</option><option value="land">Landscape</option></select></label></div>
    <div class="row"><label>Custom width (mm)<input type="number" name="w" value="210" min="20" max="2000"></label><label>Custom height (mm)<input type="number" name="h" value="297" min="20" max="2000"></label></div>
    <label>Apply to<select name="scope">${ix ? `<option value="sel">${ix.length} selected page(s)</option>` : ''}<option value="all">All pages</option><option value="cur">Current page</option></select></label>
    <p class="mut">Content is scaled to fit and centred; your edits move with it. Clickable links and form fields on resized pages are not kept.</p>`, { ok: 'Resize' });
  if (!r) return;
  const [W, H] = r.size === 'custom' ? [(+r.w || 210) * 72 / 25.4, (+r.h || 297) * 72 / 25.4] : SIZES[r.size];
  await resizePages(r.scope === 'sel' ? ix : r.scope === 'cur' ? [S.cur] : S.pages.map((_, i) => i), W, H, r.orient);
}

/* ───────── merge dialog: pick files, put them in order, combine ───────── */
async function mergeDlg() {
  const files = [];
  const r = await dialog('Merge files', `
    <p class="mut">Add PDF, Word (.docx) and image files, put them in the order you want, then merge them into one PDF.</p>
    <div class="mglist" id="mgList"></div>
    <div class="row"><button type="button" class="btn line" id="mgAdd">+ Add files…</button></div>
    ${S.pages.length ? '<label class="inl"><input type="checkbox" name="append" checked>Add them to the end of the open document (untick to start a new one)</label>' : ''}`, {
    ok: 'Merge', wide: true, onOpen: d => {
      const list = $('#mgList', d);
      const draw = () => { list.innerHTML = files.map((f, i) => `<div class="mgrow"><b>${i + 1}</b><span>${X.esc(f.name)}</span><small>${(f.size / 1048576).toFixed(2)} MB</small><button type="button" data-u="${i}" title="Move up">▲</button><button type="button" data-d="${i}" title="Move down">▼</button><button type="button" data-x="${i}" title="Remove">✕</button></div>`).join(''); };
      list.onclick = e => {
        const b = e.target.closest('button'); if (!b) return; const { u, d: dn, x } = b.dataset;
        if (x !== undefined) files.splice(+x, 1); else if (u !== undefined && +u > 0) files.splice(+u - 1, 0, files.splice(+u, 1)[0]); else if (dn !== undefined && +dn < files.length - 1) files.splice(+dn + 1, 0, files.splice(+dn, 1)[0]);
        draw();
      };
      $('#mgAdd', d).onclick = async () => { files.push(...(await pick(ACCEPT, true)).filter(f => X.isPdf(f) || X.isDocx(f) || f.type.startsWith('image/'))); draw(); };
      draw();
    },
  });
  if (!r) return; if (!files.length) return toast('Add at least one file to merge');
  await X.openFiles(files, !!r.append); if (S.pages.length) toast(`Merged ${files.length} file(s) — the document now has ${S.pages.length} page(s)`);
}

/* ───────── right-click menu inside the organizer ───────── */
X.orgCtx = e => {
  const c = e.target.closest('.ocard'), o = (label, icon, op, extra) => X.mi(label, icon, 'org', { d: { o: op }, ...extra });
  if (c && !sel.has(c.dataset.id)) { sel.clear(); sel.add(c.dataset.id); lastIdx = S.pages.findIndex(p => p.id === c.dataset.id); info(); }
  X.showCtx(e, c ? [
    X.mi('Open this page', 'pages', 'orgopen', { d: { id: c.dataset.id } }), '-',
    o('Rotate right', 'rotate', 'rotr'), o('Rotate left', 'rotatel', 'rotl'), o('Duplicate', 'copy', 'dup'), o('Move earlier', 'left', 'left'), o('Move later', 'right', 'right'), '-',
    o('Extract as new PDF', 'split', 'extract'), o('Replace with another file…', 'swap', 'replace'), o('Insert blank page after', 'plus', 'blank'), o('Insert files after…', 'merge', 'add'), o('Change page size…', 'fit', 'resize'), '-',
    o('Select all', 'selall', 'all'), o('Clear selection', 'x', 'none'), '-', o(sel.size > 1 ? `Delete ${sel.size} pages` : 'Delete page', 'trash', 'del', { danger: 1, disabled: sel.size >= S.pages.length }),
  ] : [o('Select all', 'selall', 'all'), o('Insert blank page', 'plus', 'blank'), o('Insert files…', 'merge', 'add'), o('Reverse page order', 'swap', 'reverse'), o('Change page size…', 'fit', 'resize'), o('Add page numbers…', 'hash', 'numbers'), '-', o('Done', 'check', 'done')]);
};

Object.assign(X.ACT, { mergedlg: mergeDlg, org: b => OPS[b.dataset.o] && OPS[b.dataset.o](), orgopen: b => close(S.pages.findIndex(p => p.id === b.dataset.id)), organize: open, resize: () => resizeDlg(null), replacepage: () => replacePages([S.cur]) });
})();
