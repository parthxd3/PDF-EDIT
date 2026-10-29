/* XD3 PDF Editor — document tools: searchable PDFs (OCR layer), bookmarks, comparing two PDFs, batch processing. */
(() => {
'use strict';
const X = window.XD3, { S, $, $$, toast, busy, dialog, download, clamp, esc, uid, icons } = X;
const slim = it => ({ str: it.str, x: X.r2 ? X.r2(it.x) : +it.x.toFixed(2), y: +it.y.toFixed(2), w: +it.w.toFixed(2), h: +it.h.toFixed(2), base: +it.base.toFixed(2), size: +it.size.toFixed(2), ...(it.rot ? { rot: it.rot } : {}) });

/* ───────── searchable PDF: recognise every page, keep the words as an invisible layer ───────── */
async function makeSearchable() {
  if (!X.ocrItems) return toast('Text recognition is not available (lib/ocr is missing)', true);
  const r = await dialog('Make searchable (OCR)', `
    <p>Reads the printed text on scanned pages and keeps it as an invisible layer, so the saved PDF can be <b>searched, copied and indexed</b> in any reader. The pages still look exactly the same.</p>
    <label>Pages<select name="which"><option value="need">Only pages that have no text yet (scans, pictures)</option><option value="all">All pages — also read text inside pictures on normal pages</option></select></label>
    <p class="mut">Recognition language: <b>${esc($('#pOcrLang option:checked') ? $('#pOcrLang option:checked').textContent : 'English')}</b> — change it in the Edit text toolbar. About 2–5 seconds per page.</p>`, { ok: 'Recognise' });
  if (!r) return;
  await busy(async () => {
    const before = X.snap(); let pages = 0, lines = 0;
    for (let i = 0; i < S.pages.length; i++) {
      const p = S.pages[i]; $('#busyTxt').textContent = `Reading page ${i + 1} of ${S.pages.length}…`;
      if (r.which !== 'all' && (await X.getLines(p)).length) continue;
      const found = await X.ocrItems(p, true); p.ocr = found.map(slim); pages++; lines += found.length;
    }
    if (X.snap() !== before) X.pushUndo(before);
    toast(lines ? `Recognised ${lines} line(s) on ${pages} page(s) — click Save PDF to keep the searchable copy` : 'No printed text was found on those pages', lines ? 'success' : 'info');
  }, 'Reading pages…');
}

/* ───────── bookmarks (the reader's outline panel) ───────── */
async function readOutline(pdf, pages) { // the bookmarks of an opened PDF → [{ title, pid, level }]
  const out = [];
  try {
    const items = await pdf.getOutline(); if (!items) return out;
    const walk = async (list, level) => {
      for (const it of list) {
        if (out.length > 600) return;
        let idx = -1;
        try { let d = it.dest; if (typeof d === 'string') d = await pdf.getDestination(d); if (Array.isArray(d) && d[0] != null) idx = typeof d[0] === 'object' ? await pdf.getPageIndex(d[0]) : d[0]; } catch (e) { /* broken link */ }
        if (idx >= 0 && pages[idx]) out.push({ title: String(it.title || '').trim() || 'Untitled', pid: pages[idx].id, level });
        if (it.items && it.items.length) await walk(it.items, level + 1);
      }
    };
    await walk(items, 0);
  } catch (e) { console.warn('outline', e); }
  return out;
}
function writeOutline(out, pages, made) { // called while saving: S.outline → the PDF's /Outlines tree
  const L = PDFLib, ctx = out.context, list = (S.outline || []).filter(e => pages.some(p => p.id === e.pid) && made.get(pages.find(p => p.id === e.pid)));
  if (!list.length) return;
  const rootKids = [], stack = [{ level: -1, kids: rootKids }];
  for (const e of list) { const node = { e, kids: [] }; while (stack.length > 1 && stack[stack.length - 1].level >= e.level) stack.pop(); stack[stack.length - 1].kids.push(node); stack.push({ level: e.level, kids: node.kids }); }
  const rootRef = ctx.nextRef();
  const build = (nodes, parent) => {
    const refs = nodes.map(() => ctx.nextRef()); let total = 0;
    nodes.forEach((n, i) => {
      const page = made.get(pages.find(p => p.id === n.e.pid)), d = { Title: L.PDFHexString.fromText(n.e.title), Parent: parent, Dest: [page.ref, 'XYZ', null, null, null] };
      if (i) d.Prev = refs[i - 1]; if (i < nodes.length - 1) d.Next = refs[i + 1];
      const sub = build(n.kids, refs[i]); if (sub.count) { d.First = sub.first; d.Last = sub.last; d.Count = sub.count; }
      total += 1 + sub.count; ctx.assign(refs[i], ctx.obj(d));
    });
    return { first: refs[0], last: refs[refs.length - 1], count: total };
  };
  const r = build(rootKids, rootRef);
  ctx.assign(rootRef, ctx.obj({ Type: 'Outlines', First: r.first, Last: r.last, Count: r.count }));
  out.catalog.set(L.PDFName.of('Outlines'), rootRef); out.catalog.set(L.PDFName.of('PageMode'), L.PDFName.of('UseOutlines'));
}
async function bookmarksDlg() {
  S.outline = (S.outline || []).filter(e => S.pages.some(p => p.id === e.pid));
  const row = (e, i) => { const pi = S.pages.findIndex(p => p.id === e.pid); return `<div class="bm" data-i="${i}" style="padding-left:${e.level * 18}px"><input type="text" value="${esc(e.title)}" title="Title">
    <button type="button" class="btn ico" data-b="go" title="Go to page ${pi + 1}"><b>p.${pi + 1}</b></button><button type="button" class="btn ico" data-b="out" title="Nest under the one above">→</button><button type="button" class="btn ico" data-b="in" title="Move out one level">←</button><button type="button" class="btn ico danger" data-b="del" title="Remove"><i data-icon="trash"></i></button></div>`; };
  const r = await dialog('Bookmarks', `<p class="mut">Bookmarks appear in the side panel of PDF readers. Click a page number to jump there.</p>
    <div class="bmlist">${S.outline.map(row).join('') || '<div class="empty">No bookmarks yet.</div>'}</div>
    <button type="button" class="btn line" data-b="add"><i data-icon="plus"></i>Add bookmark for page ${S.cur + 1}</button>`, {
    ok: 'Save bookmarks', wide: true, onOpen: d => {
      icons(d);
      const list = $('.bmlist', d), redraw = () => { list.innerHTML = S.outline.map(row).join('') || '<div class="empty">No bookmarks yet.</div>'; icons(list); };
      const sync = () => $$('.bm', list).forEach(el => { S.outline[+el.dataset.i].title = $('input', el).value; });
      d.addEventListener('click', e => {
        const b = e.target.closest('[data-b]'); if (!b) return; const el = b.closest('.bm'), i = el ? +el.dataset.i : -1, act = b.dataset.b; sync();
        if (act === 'add') { const p = S.pages[S.cur], lvl = S.outline.length ? S.outline[S.outline.length - 1].level : 0; S.outline.push({ title: `Page ${S.cur + 1}`, pid: p.id, level: lvl }); redraw(); const last = $$('.bm input', list).pop(); if (last) { last.focus(); last.select(); } }
        else if (act === 'del') { S.outline.splice(i, 1); redraw(); }
        else if (act === 'out') { if (i > 0) S.outline[i].level = Math.min(S.outline[i - 1].level + 1, S.outline[i].level + 1); redraw(); }
        else if (act === 'in') { S.outline[i].level = Math.max(0, S.outline[i].level - 1); redraw(); }
        else if (act === 'go') { d.returnValue = 'ok'; d.close(); X.goToPage(S.pages.findIndex(p => p.id === S.outline[i].pid)); }
      });
      d.addEventListener('close', sync, { once: true });
    } });
  if (r) { S.dirty = true; X.updateUI(); X.scheduleSave(); }
}

/* ───────── compare with another PDF ───────── */
function pickFile(accept, multiple) { return new Promise(res => { const i = document.createElement('input'); i.type = 'file'; i.accept = accept; i.multiple = !!multiple; i.onchange = () => res([...i.files]); i.click(); }); }
async function compareDlg() {
  const r = await dialog('Compare with another PDF', `<p>Choose the other version. A new document is built where each page shows the two versions on top of each other: <b style="color:#dc2626">red</b> = only in this document, <b style="color:#15803d">green</b> = only in the other one, grey = the same in both.</p>
    <p class="mut">Your current document is closed first — save it if it has changes.</p>`, { ok: 'Choose PDF…' });
  if (!r) return; const [file] = await pickFile('application/pdf'); if (!file) return;
  if (S.dirty && !confirm('Discard unsaved changes and open the comparison?')) return;
  await compareWith(new Uint8Array(await file.arrayBuffer()));
}
async function compareWith(bytes) { // the open document vs. another PDF → a new document of difference pages
  await busy(async () => {
    const other = await pdfjsLib.getDocument({ data: bytes, cMapUrl: 'lib/cmaps/', cMapPacked: true, standardFontDataUrl: 'lib/standard_fonts/', isEvalSupported: false }).promise;
    const A = S.pages.slice(), n = Math.max(A.length, other.numPages), pages = [], k = 1.6; let changed = 0;
    for (let i = 0; i < n; i++) {
      $('#busyTxt').textContent = `Comparing page ${i + 1} of ${n}…`;
      const pa = A[i], [w, h] = pa ? X.rdims(pa) : [595.28, 841.89];
      const ca = pa ? await X.renderPageImage(pa, k) : null, c = document.createElement('canvas'); c.width = Math.round(w * k); c.height = Math.round(h * k);
      const ctx = c.getContext('2d', { willReadFrequently: true }); ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, c.width, c.height);
      if (i < other.numPages) { const pg = await other.getPage(i + 1), vp0 = pg.getViewport({ scale: 1 }); await pg.render({ canvasContext: ctx, intent: 'print', viewport: pg.getViewport({ scale: c.width / vp0.width }) }).promise; }
      const da = ca ? ca.getContext('2d').getImageData(0, 0, c.width, c.height).data : null, img = ctx.getImageData(0, 0, c.width, c.height), db = img.data; let diff = 0;
      for (let j = 0; j < db.length; j += 4) {
        const ia = da ? (da[j] + da[j + 1] + da[j + 2]) / 3 < 190 : false, ib = (db[j] + db[j + 1] + db[j + 2]) / 3 < 190;
        if (ia && ib) { db[j] = db[j + 1] = db[j + 2] = 110; } else if (ia) { db[j] = 225; db[j + 1] = 45; db[j + 2] = 45; diff++; } else if (ib) { db[j] = 25; db[j + 1] = 150; db[j + 2] = 70; diff++; } else { db[j] = db[j + 1] = db[j + 2] = 255; }
      }
      ctx.putImageData(img, 0, 0); if (diff > c.width * c.height * 0.0005) changed++;
      const label = X.textObj(`Page ${i + 1} — ${diff > c.width * c.height * 0.0005 ? 'differences found' : 'identical'}${!pa ? ' (only in the other file)' : i >= other.numPages ? ' (only in this document)' : ''}`, 9, { color: diff ? '#b91c1c' : '#15803d', bg: '#ffffff' });
      label.x = 8; label.y = 6;
      pages.push({ id: uid(), src: -1, w, h, rot: 0, objs: [{ id: uid(), type: 'image', url: c.toDataURL('image/jpeg', 0.8), x: 0, y: 0, w, h, opacity: 1 }, label] });
    }
    await other.destroy();
    X.resetDoc(); S.fileName = 'comparison.pdf'; for (const pg of pages) { pg.objs[0].asset = X.addAsset(pg.objs[0].url); delete pg.objs[0].url; } S.pages.push(...pages); // (assets are registered after the reset, which clears them) X.buildPages(); S.zoom = clamp(X.fitZoom(), 0.25, 1.25); X.layout(); X.viewer.scrollTop = 0; X.setTool('select'); X.syncControls(); X.updateUI();
    toast(changed ? `${changed} of ${n} page(s) differ — red is only in your document, green only in the other file` : 'The two documents look identical', changed ? 'warn' : 'success');
  }, 'Comparing…');
}

/* ───────── batch: the same action on many files, delivered as a ZIP ───────── */
async function batchDlg() {
  const r = await dialog('Batch process', `<p>Apply one action to many PDFs at once. The results are downloaded together as a ZIP (a single file is downloaded as it is).</p>
    <label>Files<input type="file" name="files" accept="application/pdf" multiple required></label>
    <label>Action<select name="act"><option value="compress">Compress (re-encode pictures)</option><option value="ocr">Make searchable (OCR text layer)</option><option value="watermark">Add a text watermark</option><option value="pagenum">Add page numbers (bottom centre)</option><option value="flatten">Flatten to pictures (permanent redaction, 150 dpi)</option></select></label>
    <label>Watermark text (for the watermark action)<input type="text" name="wm" value="CONFIDENTIAL"></label>`, { ok: 'Run' });
  if (!r) return;
  const files = [...($('#dlg input[name=files]') || { files: [] }).files]; if (!files.length) return toast('Choose one or more PDF files', true);
  await busy(async () => {
    const outs = []; let i = 0;
    for (const f of files) {
      i++; $('#busyTxt').textContent = `${f.name} (${i} of ${files.length})…`;
      try { outs.push([f.name.replace(/\.pdf$/i, '') + '-' + r.act + '.pdf', await batchOne(new Uint8Array(await f.arrayBuffer()), r.act, r.wm)]); }
      catch (e) { console.error(e); toast(`${f.name}: ${e.message || e}`, true); }
    }
    if (!outs.length) return;
    if (outs.length === 1) return download(new Blob([outs[0][1]], { type: 'application/pdf' }), outs[0][0]);
    const JSZip = await X.loadLib('jszip'), zip = new JSZip(); for (const [n, b] of outs) zip.file(n, b);
    download(await zip.generateAsync({ type: 'blob' }), 'xd3-batch-' + r.act + '.zip'); toast(`${outs.length} file(s) processed and saved as a ZIP`);
  }, 'Processing…');
}
async function batchOne(bytes, act, wm) { // one file through the chosen action; the file is loaded as a temporary source and removed again afterwards
  const L = PDFLib;
  if (act === 'compress') {
    return (await X.compressPdf(bytes, { ...X.COMPRESS_LEVELS.balanced, strip: true })).bytes;
  }
  const pages = await X.loadSource(bytes), si = pages[0] ? pages[0].src : -1, assets = [];
  try {
    if (act === 'ocr') { for (const p of pages) if (!(await X.getLines(p)).length) p.ocr = (await X.ocrItems(p, true)).map(slim); }
    else if (act === 'watermark') {
      const w = X.makeWatermark(wm || 'CONFIDENTIAL', '#dc2626', 64), asset = X.addAsset(w.url); assets.push(asset);
      for (const p of pages) { const f = Math.min(1, Math.hypot(p.w, p.h) * 0.8 / w.w), ww = w.w * f, hh = w.h * f; p.objs.push({ id: uid(), type: 'image', asset, x: (p.w - ww) / 2, y: (p.h - hh) / 2, w: ww, h: hh, rot: 325, opacity: 0.18 }); }
    } else if (act === 'pagenum') pages.forEach((p, i) => { const o = X.textObj(`${i + 1} / ${pages.length}`, 10, { color: '#333333', align: 'center' }), b = X.bbox(o); o.x = (p.w - b.w) / 2; o.y = p.h - 28 - b.h; p.objs.push(o); });
    else if (act === 'flatten') {
      const out = await L.PDFDocument.create();
      for (const p of pages) { const c = await X.renderPageImage(p, 150 / 72); out.addPage([p.w, p.h]).drawImage(await out.embedJpg(c.toDataURL('image/jpeg', 0.72)), { x: 0, y: 0, width: p.w, height: p.h }); }
      return out.save();
    }
    return await X.buildPdf(pages, { noAnnots: true });
  } finally {
    assets.forEach(a => delete S.assets[a]);
    if (si >= 0 && si === S.sources.length - 1) { try { S.sources[si].pdf.destroy(); } catch (e) { /* already gone */ } S.sources.pop(); S.srcVer++; }
  }
}

/* ───────── recent files: the last few PDFs opened on this device, for one-click reopening ───────── */
const RECENT_MAX = 8, RECENT_BYTES = 160e6, RECENT_FILE = 40e6;
async function noteRecent(files) {
  try {
    const f = [...files].find(x => X.isPdf(x)); if (!f || f.size > RECENT_FILE) return;
    const bytes = new Uint8Array(await f.arrayBuffer()); let list = (await X.DB.get('recent')) || [];
    list = list.filter(r => r.name !== f.name); list.unshift({ name: f.name, t: Date.now(), bytes });
    let total = 0; list = list.filter(r => (total += r.bytes.length) <= RECENT_BYTES).slice(0, RECENT_MAX);
    await X.DB.put('recent', list); renderRecent(list);
  } catch (e) { /* storage unavailable */ }
}
async function renderRecent(list) {
  const host = $('#recent'); if (!host) return;
  try { list = list || (await X.DB.get('recent')) || []; } catch (e) { list = []; }
  host.hidden = !list.length; if (!list.length) return;
  const ago = t => { const d = Math.round((Date.now() - t) / 60000); return d < 2 ? 'just now' : d < 60 ? d + ' min ago' : d < 1440 ? Math.round(d / 60) + ' h ago' : Math.round(d / 1440) + ' d ago'; };
  host.innerHTML = '<h3><i data-icon="clock"></i>Recent files</h3>' + list.map((r, i) => `<button class="recent" data-r="${i}" title="Open again"><i data-icon="pages"></i><span>${esc(r.name)}</span><small>${(r.bytes.length / 1048576).toFixed(1)} MB · ${ago(r.t)}</small></button>`).join('') + '<button class="btn line" data-r="clear" style="margin-top:6px">Clear list</button>';
  icons(host);
}
$('#welcome').addEventListener('click', async e => {
  const b = e.target.closest('[data-r]'); if (!b) return;
  if (b.dataset.r === 'clear') { await X.DB.put('recent', []); renderRecent([]); return; }
  const list = (await X.DB.get('recent')) || [], r = list[+b.dataset.r]; if (r) X.openFiles([new File([r.bytes], r.name, { type: 'application/pdf' })], false);
});
renderRecent();
X.noteRecent = noteRecent;

/* ───────── links and cropping ───────── */
async function linkDlg(p, o, isNew) {
  const pi = o.page ? S.pages.findIndex(q => q.id === o.page) + 1 : '';
  const r = await dialog(isNew ? 'Add link' : 'Edit link', `
    <label>Link to<select name="kind"><option value="url"${o.page ? '' : ' selected'}>A website or e-mail</option><option value="page"${o.page ? ' selected' : ''}>A page in this document</option></select></label>
    <label>Web address (https://… or mailto:…)<input type="text" name="url" value="${esc(o.url || '')}" placeholder="https://example.com" autocomplete="off"></label>
    <label>Page number (1–${S.pages.length})<input type="number" name="page" min="1" max="${S.pages.length}" value="${pi}"></label>
    <p class="mut">The box stays invisible in the saved PDF; clicking it in any reader follows the link.</p>`, { ok: isNew ? 'Add link' : 'Save' });
  if (!r) { if (isNew) { p.objs = p.objs.filter(x => x !== o); X.select(null); X.drawObjs(p); } return; }
  if (r.kind === 'page') { const n = clamp(+r.page || 1, 1, S.pages.length); o.page = S.pages[n - 1].id; o.url = ''; }
  else { let u = r.url.trim(); if (!u) return toast('Type a web address', true); if (!/^[a-z][a-z0-9+.-]*:/i.test(u)) u = (/@/.test(u) && !/\//.test(u) ? 'mailto:' : 'https://') + u; o.url = u; delete o.page; }
  if (!isNew) X.pushUndo(); X.drawObjs(p); toast('Link saved — it works in the PDF you save');
}
async function cropDlg(p, b) {
  const x = clamp(b.x, 0, p.w), y = clamp(b.y, 0, p.h), box = { x, y, w: Math.min(b.x + b.w, p.w) - x, h: Math.min(b.y + b.h, p.h) - y };
  if (box.w < 20 || box.h < 20) return toast('Drag a larger box to crop');
  const r = await dialog('Crop pages', `<p>Keep only the marked area (${Math.round(box.w)} × ${Math.round(box.h)} pt). The rest is hidden in every PDF reader; nothing is cut out of the file.</p>
    <label>Apply to<select name="scope"><option value="this">This page</option><option value="all">All pages</option><option value="same">All pages of the same size</option></select></label>`, { ok: 'Crop' });
  if (!r) return; X.pushUndo(); let n = 0;
  for (const q of S.pages) {
    if (r.scope === 'this' ? q !== p : r.scope === 'same' ? Math.abs(q.w - p.w) > 1 || Math.abs(q.h - p.h) > 1 : false) continue;
    const cx = Math.min(box.x, q.w - 20), cy = Math.min(box.y, q.h - 20); q.crop = { x: cx, y: cy, w: Math.min(box.w, q.w - cx), h: Math.min(box.h, q.h - cy) }; X.drawObjs(q); n++;
  }
  X.setTool('select'); toast(`Cropped ${n} page(s) — right-click a page and choose Remove crop to undo it`);
}
X.ACT.linktool = () => X.setTool('link'); X.ACT.croptool = () => X.setTool('crop');

X.NO_DOC.add('batch'); X.NO_DOC.add('annots');
X.ACT.annots = () => { const on = localStorage.getItem('xd3.annots') === 'off'; try { localStorage.setItem('xd3.annots', on ? 'on' : 'off'); } catch (e) { /* not remembered */ } syncAnnots(); toast(on ? 'Shapes, highlights and notes will be saved as PDF comments' : 'Shapes, highlights and notes will be saved as flat page content'); };
function syncAnnots() { let on = true; try { on = localStorage.getItem('xd3.annots') !== 'off'; } catch (e) { /* default */ } $$('[data-act=annots]').forEach(b => b.classList.toggle('on', on)); }
syncAnnots();
Object.assign(X.ACT, { ocrall: makeSearchable, bookmarks: bookmarksDlg, compare: compareDlg, batch: batchDlg });
Object.assign(X, { linkDlg, cropDlg, readOutline, writeOutline, makeSearchable, compareWith, batchOne });
})();
