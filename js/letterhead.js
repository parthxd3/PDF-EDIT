/* XD3 PDF Editor — letterheads: design one (logo, name, tagline, contact details, colour, layout), upload your own, or take it from the
   top of a PDF page; keep several; put one on PDF pages or into a Word document's header and footer. Stored on this device only. */
(() => {
'use strict';
const X = window.XD3, { S, $, toast, dialog, esc, uid } = X;
const KEY = 'xd3.letterheads';
const list = () => { try { return JSON.parse(localStorage.getItem(KEY) || '[]'); } catch (e) { return []; } };
function save(all) {
  try { localStorage.setItem(KEY, JSON.stringify(all)); return true; }
  catch (e) { toast('There is no room left to keep another letterhead on this device — delete one first (Manage letterheads)', true); return false; }
}
const loadImg = src => new Promise((res, rej) => { const i = new Image(); i.onload = () => res(i); i.onerror = rej; i.src = src; });
const fileUrl = f => new Promise((res, rej) => { const r = new FileReader(); r.onload = () => res(r.result); r.onerror = rej; r.readAsDataURL(f); });
async function shrink(url, maxW = 2400) { // big photos are scaled down so the letterhead stays small enough to keep
  const im = await loadImg(url), f = Math.min(1, maxW / im.width), c = document.createElement('canvas'); c.width = Math.round(im.width * f); c.height = Math.round(im.height * f);
  c.getContext('2d').drawImage(im, 0, 0, c.width, c.height); return c.toDataURL('image/png');
}

/* ───────── drawing a designed letterhead ───────── */
const W = 2400; // pixels across (an A4 / Letter text width at about 300 dpi)
async function draw(o) { // → { header, footer } data URLs
  const logo = o.logo ? await loadImg(o.logo).catch(() => null) : null, acc = o.color || '#1d4ed8', font = o.font || 'Segoe UI, Calibri, Arial, sans-serif';
  const c = document.createElement('canvas'), x = c.getContext('2d'); c.width = W; c.height = 420;
  const lines = [o.address, [o.phone, o.email].filter(Boolean).join('  ·  '), o.web].filter(Boolean);
  x.textBaseline = 'alphabetic'; const fit = (t, px, wt, max) => { let s = px; x.font = `${wt} ${s}px ${font}`; while (x.measureText(t).width > max && s > 20) { s -= 2; x.font = `${wt} ${s}px ${font}`; } return s; };
  let h = 0;
  if (o.layout === 'center') {
    let y = 20; if (logo) { const lh = 150, lw = logo.width * lh / logo.height; x.drawImage(logo, (W - lw) / 2, y, lw, lh); y += lh + 30; }
    if (o.name) { fit(o.name, 92, 700, W - 200); x.fillStyle = acc; x.textAlign = 'center'; y += 80; x.fillText(o.name, W / 2, y); }
    if (o.tagline) { fit(o.tagline, 44, 400, W - 200); x.fillStyle = '#555'; y += 58; x.fillText(o.tagline, W / 2, y); }
    if (lines.length && !o.footerContact) { x.font = `400 36px ${font}`; x.fillStyle = '#333'; y += 56; x.fillText(lines.join('   |   '), W / 2, y); }
    h = y + 40;
  } else {
    const right = o.layout === 'right', pad = 0; let tx = pad, top = 30;
    if (logo) { const lh = 220, lw = Math.min(700, logo.width * lh / logo.height); x.drawImage(logo, right ? W - lw : 0, top, lw, lh); tx = (right ? 0 : lw + 50); h = top + lh; }
    x.textAlign = right ? 'right' : 'left'; const nx = right ? W - (logo ? Math.min(700, logo.width * 220 / logo.height) + 50 : 0) : tx;
    let y = top + 95; if (o.name) { fit(o.name, 96, 700, 1300); x.fillStyle = acc; x.fillText(o.name, nx, y); }
    if (o.tagline) { fit(o.tagline, 44, 400, 1300); x.fillStyle = '#555'; y += 64; x.fillText(o.tagline, nx, y); }
    h = Math.max(h, y + 30);
    if (lines.length && !o.footerContact) { x.textAlign = right ? 'left' : 'right'; x.font = `400 34px ${font}`; x.fillStyle = '#333'; let cy = top + 50; for (const l of lines) { x.fillText(l, right ? 0 : W, cy); cy += 50; } h = Math.max(h, cy); }
  }
  h += 16; x.fillStyle = acc; x.fillRect(0, h, W, 8); h += 8;
  const out = document.createElement('canvas'); out.width = W; out.height = h + 4; out.getContext('2d').drawImage(c, 0, 0);
  let footer = null;
  if (o.footerContact && lines.length) {
    const f = document.createElement('canvas'), fx = f.getContext('2d'); f.width = W; f.height = 110; fx.fillStyle = acc; fx.fillRect(0, 0, W, 5);
    fx.font = `400 34px ${font}`; fx.fillStyle = '#444'; fx.textAlign = 'center'; fx.textBaseline = 'alphabetic'; let s = 34; while (fx.measureText(lines.join('   |   ')).width > W - 40 && s > 18) { s -= 2; fx.font = `400 ${s}px ${font}`; }
    fx.fillText(lines.join('   |   '), W / 2, 70); footer = f.toDataURL('image/png');
  }
  return { header: out.toDataURL('image/png'), footer };
}

/* ───────── making one ───────── */
async function design(prev) {
  const o = prev || { layout: 'left', color: '#1d4ed8' };
  const html = `<div class="lh-form">
    <div class="row"><label>Company or your name<input type="text" name="name" value="${esc(o.name || '')}" autocomplete="off"></label><label>Tagline<input type="text" name="tagline" value="${esc(o.tagline || '')}" autocomplete="off" placeholder="optional"></label></div>
    <label>Address<input type="text" name="address" value="${esc(o.address || '')}" autocomplete="off"></label>
    <div class="row"><label>Phone<input type="text" name="phone" value="${esc(o.phone || '')}" autocomplete="off"></label><label>E-mail<input type="text" name="email" value="${esc(o.email || '')}" autocomplete="off"></label><label>Website<input type="text" name="web" value="${esc(o.web || '')}" autocomplete="off"></label></div>
    <div class="row"><label>Logo<input type="file" name="logofile" accept="image/*"></label><label>Layout<select name="layout"><option value="left">Logo left</option><option value="center">Centred</option><option value="right">Logo right</option></select></label><label>Colour<input type="color" name="color" value="${esc(o.color || '#1d4ed8')}"></label></div>
    <label class="inl"><input type="checkbox" name="footerContact"${o.footerContact ? ' checked' : ''}>Put the contact details in a footer line</label>
    <label>Save as<input type="text" name="label" value="${esc(o.label || o.name || 'My letterhead')}" autocomplete="off"></label>
    <div class="lh-prev"><img id="lhPrev" alt="Preview"><img id="lhPrevF" alt="" hidden></div></div>`;
  let logo = o.logo || null, last = null;
  const refresh = async form => { const f = Object.fromEntries(new FormData(form)); const r = await draw({ ...f, logo, footerContact: !!f.footerContact }); last = r; const a = $('#lhPrev'), b = $('#lhPrevF'); if (a) a.src = r.header; if (b) { b.hidden = !r.footer; if (r.footer) b.src = r.footer; } };
  setTimeout(() => { const form = $('#dlg form'); if (!form) return; form.querySelector('[name=layout]').value = o.layout || 'left'; refresh(form);
    form.addEventListener('input', e => { if (e.target.name === 'logofile') return; refresh(form); });
    form.querySelector('[name=logofile]').addEventListener('change', async e => { const fl = e.target.files[0]; if (fl) { logo = await shrink(await fileUrl(fl), 900); refresh(form); } }); }, 50);
  const r = await dialog('Design a letterhead', html, { ok: 'Save letterhead' }); if (!r) return null;
  const data = { ...r, logo, footerContact: !!r.footerContact }; delete data.logofile; const img = await draw(data);
  const lh = { id: uid(), label: (r.label || r.name || 'Letterhead').trim(), header: img.header, footer: img.footer, src: data, created: Date.now() };
  return keep(lh);
}
function keep(lh) { const all = list().filter(x => x.id !== lh.id); all.unshift(lh); if (!save(all.slice(0, 20))) return null; toast(`Letterhead “${lh.label}” saved`, 'success'); return lh; }
async function upload() {
  const r = await dialog('Use your own letterhead', `<p>Choose a picture of your letterhead (PNG or JPG) — the top part of your company paper, as wide as the page. A footer picture is optional.</p>
    <label>Letterhead picture<input type="file" name="h" accept="image/png,image/jpeg"></label><label>Footer picture (optional)<input type="file" name="f" accept="image/png,image/jpeg"></label>
    <label>Save as<input type="text" name="label" value="My letterhead" autocomplete="off"></label>`, { ok: 'Save letterhead' });
  if (!r || !r.h || !r.h.size) return null;
  const lh = { id: uid(), label: (r.label || 'Letterhead').trim(), header: await shrink(await fileUrl(r.h)), footer: r.f && r.f.size ? await shrink(await fileUrl(r.f)) : null, created: Date.now() };
  return keep(lh);
}
async function fromPage() { // the top of the PDF page that is open → a letterhead
  const p = S.pages[S.cur]; if (!p) return toast('Open a PDF first'); const c = await X.renderPageImage(p, 200 / 72);
  const html = `<p>The top of this page becomes a letterhead you can put on other PDFs and Word documents.</p><label>How much of the page<input type="range" name="pct" min="5" max="40" value="16"></label><div class="lh-prev"><img id="lhCrop" alt="Preview"></div><label>Save as<input type="text" name="label" value="${esc(X.baseName())} letterhead" autocomplete="off"></label>`;
  const crop = pct => { const h = Math.round(c.height * pct / 100), o = document.createElement('canvas'); o.width = c.width; o.height = h; o.getContext('2d').drawImage(c, 0, 0, c.width, h, 0, 0, c.width, h); return o.toDataURL('image/png'); };
  setTimeout(() => { const form = $('#dlg form'); if (!form) return; const go = () => { const i = $('#lhCrop'); if (i) i.src = crop(+form.querySelector('[name=pct]').value); }; go(); form.addEventListener('input', go); }, 50);
  const r = await dialog('Save a letterhead from this page', html, { ok: 'Save letterhead' }); if (!r) return null;
  return keep({ id: uid(), label: (r.label || 'Letterhead').trim(), header: await shrink(crop(+r.pct)), footer: null, created: Date.now() });
}
async function manage() {
  const all = list(); if (!all.length) return toast('No letterheads saved yet', 'info');
  const r = await dialog('Manage letterheads', `<div class="lh-list">${all.map(l => `<label class="lh-item"><input type="checkbox" name="del_${l.id}"><img src="${l.header}" alt=""><span>${esc(l.label)}</span></label>`).join('')}</div><p class="mut">Tick the ones to delete.</p>`, { ok: 'Delete ticked' });
  if (!r) return; const del = new Set(Object.keys(r).filter(k => k.startsWith('del_')).map(k => k.slice(4))); if (!del.size) return;
  save(all.filter(l => !del.has(l.id))); toast(`${del.size} letterhead(s) deleted`, 'success');
}

/* ───────── using one ───────── */
async function onPdf(lh) {
  if (!S.pages.length) return toast('Open a PDF first');
  const r = await dialog('Letterhead on the PDF', `<label>Pages<select name="where"><option value="all">Every page</option><option value="first">First page only</option><option value="cur">This page</option></select></label><label class="inl"><input type="checkbox" name="foot"${lh.footer ? ' checked' : ' disabled'}>Footer too</label><p class="mut">It is added as a picture you can move, resize or delete.</p>`, { ok: 'Add letterhead' });
  if (!r) return; const pages = r.where === 'all' ? S.pages : r.where === 'first' ? S.pages.slice(0, 1) : [S.pages[S.cur]];
  const im = await loadImg(lh.header), hAsset = X.addAsset(lh.header), fIm = r.foot && lh.footer ? await loadImg(lh.footer) : null, fAsset = fIm ? X.addAsset(lh.footer) : null;
  X.pushUndo();
  for (const p of pages) {
    const m = Math.min(36, p.w * 0.06), w = p.w - 2 * m, h = w * im.height / im.width;
    p.objs.push({ id: uid(), type: 'image', asset: hAsset, x: m, y: Math.min(24, m), w, h, opacity: 1 });
    if (fIm) { const fh = w * fIm.height / fIm.width; p.objs.push({ id: uid(), type: 'image', asset: fAsset, x: m, y: p.h - Math.min(24, m) - fh, w, h: fh, opacity: 1 }); }
    X.drawObjs(p);
  }
  toast(`Letterhead added to ${pages.length} page(s)`, 'success');
}
async function onDocx(lh) { if (X.docx && X.docx.applyLetterhead) await X.docx.applyLetterhead(lh); }
function menu(at, target) {
  const mi = X.mi, all = list();
  X.LH_PICK = { target, all };
  X.showCtx(at, [...all.slice(0, 12).map((l, i) => mi(l.label, 'stamp', 'lh_use', { d: { i }, keep: 1 })), all.length ? '-' : null,
    mi('Design a new letterhead…', 'brush', 'lh_new', { keep: 1 }), mi('Use my own letterhead picture…', 'image', 'lh_upload', { keep: 1 }),
    target === 'pdf' && S.pages.length && mi('Save the top of this page as a letterhead…', 'crop', 'lh_page', { keep: 1 }),
    all.length && mi('Manage letterheads…', 'sliders', 'lh_manage', { keep: 1 })]);
}
const use = async lh => { if (!lh) return; if (X.LH_PICK && X.LH_PICK.target === 'docx') await onDocx(lh); else await onPdf(lh); };
Object.assign(X.ACT, {
  letterhead: b => menu(b && b.getBoundingClientRect ? (() => { const r = b.getBoundingClientRect(); return { clientX: r.left, clientY: r.bottom + 4, target: b }; })() : { clientX: innerWidth / 2, clientY: 120, target: document.body }, 'pdf'),
  lh_use: b => use(X.LH_PICK.all[+b.dataset.i]), lh_new: async () => use(await design()), lh_upload: async () => use(await upload()), lh_page: async () => use(await fromPage()), lh_manage: manage,
});
['lh_use', 'lh_new', 'lh_upload', 'lh_manage', 'lh_page'].forEach(a => X.NO_DOC.add(a));
X.letterheads = { list, draw, design, upload, fromPage, manage, menu, onPdf };
})();
