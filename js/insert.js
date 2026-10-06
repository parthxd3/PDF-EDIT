/* XD3 PDF Editor — things you insert: date, stamps, watermark, page numbers, header / footer, signature. */
(() => {
'use strict';
const X = window.XD3, { S, $, toast, dialog, clamp, esc, uid, textObj, bbox, placeObj, insertImage, addAsset, pushUndo, drawObjs, drawSel, parseRange, baseName } = X;

function insertDate() { const o = textObj(new Date().toLocaleDateString(undefined, { day: '2-digit', month: 'short', year: 'numeric' }), 12, { color: '#111111' }), b = bbox(o); placeObj(o, b.w, b.h); }
function makeStamp(text, color, plain) {
  const k = 3, fs = plain ? 40 : 26, c = document.createElement('canvas'); let x = c.getContext('2d');
  const font = `800 ${fs * k}px Arial, sans-serif`; x.font = font;
  c.width = Math.ceil(x.measureText(text).width + (plain ? 8 : 36) * k); c.height = (fs + (plain ? 8 : 22)) * k;
  x = c.getContext('2d'); x.font = font; x.fillStyle = x.strokeStyle = color; x.textAlign = 'center'; x.textBaseline = 'middle';
  if (!plain) { x.lineWidth = 3 * k; x.beginPath(); if (x.roundRect) x.roundRect(3 * k, 3 * k, c.width - 6 * k, c.height - 6 * k, 7 * k); else x.rect(3 * k, 3 * k, c.width - 6 * k, c.height - 6 * k); x.stroke(); }
  x.fillText(text, c.width / 2, c.height / 2 + k);
  return { url: c.toDataURL('image/png'), w: c.width / k, h: c.height / k };
}
function makeWatermark(text, color, size) {
  const k = 2, font = `bold ${size * k}px Arial, sans-serif`, c = document.createElement('canvas'); let x = c.getContext('2d'); x.font = font;
  c.width = Math.ceil(x.measureText(text).width + 8 * k); c.height = Math.ceil(size * k * 1.3);
  x = c.getContext('2d'); x.font = font; x.fillStyle = color; x.textAlign = 'center'; x.textBaseline = 'middle'; x.fillText(text, c.width / 2, c.height / 2);
  return { url: c.toDataURL('image/png'), w: c.width / k, h: c.height / k };
}
async function watermarkDlg() {
  const r = await dialog('Add watermark', `
    <label>Text<input type="text" name="text" value="CONFIDENTIAL" required></label>
    <div class="row"><label>Size<input type="number" name="size" value="64" min="8" max="300"></label>
    <label>Angle°<input type="number" name="angle" value="-35" min="-180" max="180"></label>
    <label>Opacity<input type="number" name="op" value="0.18" min="0.02" max="1" step="0.02"></label></div>
    <div class="row"><label>Color<input type="color" name="color" value="#dc2626"></label>
    <label>Apply to<select name="scope"><option value="all">All pages</option><option value="cur">Current page</option></select></label></div>`, { ok: 'Add watermark' });
  if (!r || !r.text.trim()) return;
  const wm = makeWatermark(r.text.trim(), r.color, +r.size || 64), asset = addAsset(wm.url), rot = (((+r.angle || 0) % 360) + 360) % 360; pushUndo();
  for (const p of r.scope === 'all' ? S.pages : [S.pages[S.cur]]) {
    const f = Math.min(1, Math.hypot(p.w, p.h) * 0.8 / wm.w), w = wm.w * f, h = wm.h * f;
    p.objs.push({ id: uid(), type: 'image', asset, x: (p.w - w) / 2, y: (p.h - h) / 2, w, h, rot, opacity: clamp(+r.op || 0.18, 0.02, 1) }); drawObjs(p);
  }
  toast('Watermark added — select it to move, resize, rotate or delete');
}
function stampText(p, o, pos, m) { // pos: [t|b][l|c|r]
  const b = bbox(o); o.x = pos[1] === 'c' ? (p.w - b.w) / 2 : pos[1] === 'r' ? p.w - m - b.w : m; o.y = pos[0] === 'b' ? p.h - m - b.h : m;
  o.align = { l: 'left', c: 'center', r: 'right' }[pos[1]]; p.objs.push(o); drawObjs(p);
}
async function pageNumDlg() {
  const r = await dialog('Add page numbers', `
    <div class="row"><label>Format<select name="fmt"><option value="{n}">1</option><option value="Page {n}">Page 1</option><option value="{n} / {N}">1 / N</option><option value="Page {n} of {N}">Page 1 of N</option><option value="- {n} -">- 1 -</option></select></label>
    <label>Position<select name="pos"><option value="bc">Bottom center</option><option value="br">Bottom right</option><option value="bl">Bottom left</option><option value="tc">Top center</option><option value="tr">Top right</option><option value="tl">Top left</option></select></label></div>
    <label>Custom text — overrides Format ({n} = number, {N} = total), e.g. DOC-{n}<input type="text" name="custom" autocomplete="off"></label>
    <div class="row"><label>Font size<input type="number" name="size" value="10" min="5" max="40"></label>
    <label>Start at<input type="number" name="start" value="1" min="0"></label>
    <label>Margin<input type="number" name="m" value="28" min="4" max="200"></label>
    <label style="flex:0 0 60px">Color<input type="color" name="color" value="#333333"></label></div>
    <div class="row"><label>Font<select name="font"><option>Helvetica</option><option>Times</option><option>Courier</option></select></label>
    <label>Only pages (blank = all), e.g. 2-10<input type="text" name="range" autocomplete="off"></label></div>
    <label class="inl"><input type="checkbox" name="skip1">Skip the first page (cover)</label>
    <label class="inl"><input type="checkbox" name="replace" checked>Replace page numbers added earlier</label>`, { ok: 'Add numbers' });
  if (!r) return;
  let idx = r.range.trim() ? parseRange(r.range, S.pages.length) : S.pages.map((_, i) => i);
  if (!idx) return toast(`Invalid page range. Use numbers 1–${S.pages.length}, e.g. 2-10`, true);
  if (r.skip1) idx = idx.filter(i => i > 0);
  pushUndo();
  if (r.replace) S.pages.forEach(p => { if (p.objs.some(o => o.pn)) { p.objs = p.objs.filter(o => !o.pn); drawObjs(p); } });
  const st = Number.isFinite(+r.start) ? +r.start : 1, N = idx.length + st - 1, fmt = r.custom.trim() || r.fmt;
  idx.forEach((pi, k) => {
    const text = fmt.replace(/\{n\}/g, k + st).replace(/\{N\}/g, N);
    stampText(S.pages[pi], textObj(text, +r.size || 10, { color: r.color, font: r.font, pn: true }), r.pos, +r.m || 28);
  });
  S.sel = null; drawSel(); toast(`Page numbers added to ${idx.length} page(s)`);
}
async function headFootDlg() {
  const r = await dialog('Header / footer', `
    <label>Header text<input type="text" name="head" placeholder="e.g. Company name — {date}"></label>
    <label>Footer text<input type="text" name="foot" placeholder="e.g. Page {page} of {pages}"></label>
    <div class="row"><label>Align<select name="al"><option value="c">Center</option><option value="l">Left</option><option value="r">Right</option></select></label>
    <label>Font size<input type="number" name="size" value="10" min="5" max="40"></label>
    <label>Margin<input type="number" name="m" value="28" min="4" max="200"></label></div>
    <p style="color:var(--mut)">Placeholders: {page} {pages} {date} {file}</p>`, { ok: 'Add to all pages' });
  if (!r || (!r.head.trim() && !r.foot.trim())) return; pushUndo();
  const N = S.pages.length, fill = (t, i) => t.replace(/\{page\}/g, i + 1).replace(/\{pages\}/g, N).replace(/\{date\}/g, new Date().toLocaleDateString()).replace(/\{file\}/g, baseName());
  S.pages.forEach((p, i) => {
    if (r.head.trim()) stampText(p, textObj(fill(r.head.trim(), i), +r.size || 10), 't' + r.al, +r.m || 28);
    if (r.foot.trim()) stampText(p, textObj(fill(r.foot.trim(), i), +r.size || 10), 'b' + r.al, +r.m || 28);
  });
  toast('Header / footer added');
}
async function signDlg() {
  let pad; const last = (() => { try { return localStorage.getItem('xd3.sig'); } catch (e) { return null; } })();
  const r = await dialog('Signature', `
    <canvas class="sigpad" width="1120" height="400"></canvas>
    <div class="row"><label>…or type it<input type="text" name="typed" placeholder="Your name" autocomplete="off"></label>
    <label style="flex:0 0 70px">Ink<input type="color" name="ink" value="#0b2a7a"></label></div>
    <div class="row"><button type="button" class="btn" data-s="clear" style="border:1px solid var(--line)">Clear</button>
    ${last ? '<button type="button" class="btn" data-s="last" style="border:1px solid var(--line)">Use last signature</button>' : ''}</div>`, {
    ok: 'Insert', onOpen: d => {
      const c = $('.sigpad', d), x = c.getContext('2d'), ink = $('[name=ink]', d), typed = $('[name=typed]', d); let on = false, used = false;
      const pos = e => { const b = c.getBoundingClientRect(); return [(e.clientX - b.left) * c.width / b.width, (e.clientY - b.top) * c.height / b.height]; };
      const clear = () => { x.clearRect(0, 0, c.width, c.height); used = false; };
      c.onpointerdown = e => { on = true; used = true; c.setPointerCapture(e.pointerId); x.strokeStyle = ink.value; x.lineWidth = 5; x.lineCap = x.lineJoin = 'round'; x.beginPath(); x.moveTo(...pos(e)); };
      c.onpointermove = e => { if (on) { x.lineTo(...pos(e)); x.stroke(); } };
      c.onpointerup = c.onpointercancel = () => on = false;
      const type = () => { clear(); if (!typed.value) return; used = true; x.fillStyle = ink.value; x.font = 'italic 150px "Segoe Script", "Brush Script MT", "Lucida Handwriting", cursive'; x.textAlign = 'center'; x.textBaseline = 'middle'; x.fillText(typed.value, c.width / 2, c.height / 2, c.width - 40); };
      typed.oninput = type; ink.oninput = () => { if (typed.value) type(); };
      $('[data-s=clear]', d).onclick = () => { typed.value = ''; clear(); };
      const lb = $('[data-s=last]', d); if (lb) lb.onclick = () => { const im = new Image(); im.onload = () => { clear(); used = true; const f = Math.min(c.width / im.width, c.height / im.height, 1); x.drawImage(im, (c.width - im.width * f) / 2, (c.height - im.height * f) / 2, im.width * f, im.height * f); }; im.src = last; };
      pad = () => { // crop to ink bounds
        if (!used) return null; const d2 = x.getImageData(0, 0, c.width, c.height).data; let a = c.width, b = c.height, e2 = 0, f = 0;
        for (let yy = 0; yy < c.height; yy++) for (let xx = 0; xx < c.width; xx++) if (d2[(yy * c.width + xx) * 4 + 3] > 8) { if (xx < a) a = xx; if (xx > e2) e2 = xx; if (yy < b) b = yy; if (yy > f) f = yy; }
        if (e2 <= a || f <= b) return null; const o = document.createElement('canvas'); o.width = e2 - a + 12; o.height = f - b + 12;
        o.getContext('2d').drawImage(c, a - 6, b - 6, o.width, o.height, 0, 0, o.width, o.height); return { url: o.toDataURL('image/png'), w: o.width, h: o.height };
      };
    }
  });
  if (!r) return; const sig = pad(); if (!sig) return toast('Signature is empty', true);
  try { localStorage.setItem('xd3.sig', sig.url); } catch (e) { /* storage unavailable */ }
  const w = Math.min(190, sig.w / 3); insertImage(sig.url, w, w * sig.h / sig.w);
}

Object.assign(X.ACT, { sign: signDlg, watermark: watermarkDlg, pagenum: pageNumDlg, headfoot: headFootDlg, datestamp: insertDate,
  stamp: b => { const s = makeStamp(b.dataset.t, b.dataset.c, !!b.dataset.plain); insertImage(s.url, s.w, s.h); } });
Object.assign(X, { pageNumDlg, makeStamp, makeWatermark });
})();
