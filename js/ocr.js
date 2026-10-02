/* XD3 PDF Editor — offline text recognition (Tesseract OCR, bundled in lib/ocr): makes the text in scans, flattened pages and pictures editable,
   and matches a typeface, size and colour to printed text that has no font of its own. */
(() => {
'use strict';
const X = window.XD3, { S, $, PE, toast, clamp } = X;
const DIR = new URL('lib/ocr/', location.href).href;
const LANGS = { eng: 'English', 'eng+hin': 'English + Hindi' };
const ocrLang = () => { try { const v = localStorage.getItem('xd3.ocr'); return LANGS[v] ? v : 'eng'; } catch (e) { return 'eng'; } };

/* ───────── engine: one background worker, started on first use and released when idle ───────── */
let eng = null, engLang = '', idle = 0, chain = Promise.resolve(), onProg = null;
const loadScript = src => new Promise((res, rej) => { const s = document.createElement('script'); s.src = src; s.onload = res; s.onerror = () => rej(new Error('the OCR engine files are missing (lib/ocr)')); document.head.appendChild(s); });
async function engine(lang) {
  clearTimeout(idle);
  if (eng && engLang !== lang) { const old = eng; eng = null; await old.terminate().catch(() => {}); }
  if (!eng) {
    if (!window.Tesseract) await loadScript('lib/ocr/tesseract.min.js');
    eng = await window.Tesseract.createWorker(lang.split('+'), 1, { workerPath: DIR + 'worker.min.js', corePath: DIR, langPath: DIR, workerBlobURL: false, gzip: true, cacheMethod: 'none',
      logger: m => { if (onProg && m.status === 'recognizing text') onProg(m.progress || 0); } });
    await eng.setParameters({ user_defined_dpi: '300' }); engLang = lang;
  }
  return eng;
}
const job = (canvas, lang, prog, psm = '3', skip) => chain = chain.catch(() => {}).then(async () => { // recognitions run one after another
  try {
    if (skip && skip()) throw new Error('skip');
    const w = await engine(lang); onProg = prog; await w.setParameters({ tessedit_pageseg_mode: psm });
    return (await w.recognize(canvas, {}, { blocks: true, text: false })).data;
  } finally { onProg = null; clearTimeout(idle); idle = setTimeout(() => { if (eng) { eng.terminate().catch(() => {}); eng = null; } }, 120000); }
});

function linesOf(data, k, ox, oy) { // Tesseract's result → editable lines in page coordinates (k = pixels per point of the recognised picture)
  const out = [];
  for (const b of data.blocks || []) for (const pa of b.paragraphs || []) for (const ln of pa.lines || []) {
    const lh = ln.bbox.y1 - ln.bbox.y0, bl = ln.baseline || {}; let run = [];
    const flush = () => {
      if (!run.length) return; const ws = run; run = [];
      const str = ws.map(w => w.text.trim()).join(' '), conf = ws.reduce((a, w) => a + w.confidence, 0) / ws.length;
      if (conf < 45 || !/[\p{L}\p{N}]/u.test(str)) return;
      const x0 = Math.min(...ws.map(w => w.bbox.x0)), x1 = Math.max(...ws.map(w => w.bbox.x1)), y0 = Math.min(...ws.map(w => w.bbox.y0)), y1 = Math.max(...ws.map(w => w.bbox.y1));
      if (y1 - y0 < 6 || x1 - x0 < 4) return;
      let by = bl.x1 > bl.x0 ? bl.y0 + (bl.y1 - bl.y0) * ((x0 + x1) / 2 - bl.x0) / (bl.x1 - bl.x0) : y1; by = clamp(by, y0 + (y1 - y0) * 0.55, y1);
      const size = Math.max(4, (by - y0) / 0.72 / k), top = y0 / k - 0.06 * size, bot = Math.max(y1 / k, by / k + 0.22 * size) + 0.06 * size;
      out.push({ str, x: ox + x0 / k, w: (x1 - x0) / k, y: oy + top, h: bot - top, base: oy + by / k, size, font: 'Helvetica', ocr: true, conf });
    };
    for (const w of ln.words || []) {
      if (!w.text || !w.text.trim() || w.confidence < 20) continue;
      if (run.length && w.bbox.x0 - run[run.length - 1].bbox.x1 > 2.2 * lh) flush(); // a wide gap = another column or table cell
      run.push(w);
    }
    flush();
  }
  return out;
}

/* ───────── which parts of a page need reading ───────── */
async function targets(p) {
  const lines = p.src >= 0 ? await X.getLines(p) : [], full = [{ x: 0, y: 0, w: p.w, h: p.h }];
  const imgs = [...(p.src >= 0 ? (await X.scanPage(p)).all : []), ...p.objs.filter(X.isBaseImg).map(o => X.aabb({ ...X.bbox(o), rot: o.rot }))];
  if (p.src >= 0 && !lines.length) { full.full = true; return full; } // no text layer at all: a scan, or text flattened to shapes
  let regs = imgs.map(b => { const x = clamp(b.x, 0, p.w), y = clamp(b.y, 0, p.h); return { x, y, w: Math.min(b.x + b.w, p.w) - x, h: Math.min(b.y + b.h, p.h) - y }; })
    .filter(b => b.w > 36 && b.h > 12 && !lines.some(l => { const cx = l.x + l.w / 2, cy = l.y + l.h / 2; return cx > b.x && cx < b.x + b.w && cy > b.y && cy < b.y + b.h; })); // (pictures that already carry real text need no reading)
  for (let merged = true; merged;) { // overlapping pictures are read as one
    merged = false;
    for (let i = 0; i < regs.length && !merged; i++) for (let j = i + 1; j < regs.length; j++) if (X.hitsRect(regs[i], regs[j])) {
      const a = regs[i], b = regs[j], x = Math.min(a.x, b.x), y = Math.min(a.y, b.y); regs[i] = { x, y, w: Math.max(a.x + a.w, b.x + b.w) - x, h: Math.max(a.y + a.h, b.y + b.h) - y }; regs.splice(j, 1); merged = true; break;
    }
  }
  if (regs.length > 6) { full.full = true; return full; }
  return regs;
}
function badge(p) {
  const pe = PE.get(p.id); if (!pe) return { set() {}, done() {} };
  const b = document.createElement('div'); b.className = 'ocrbadge'; b.innerHTML = '<i></i><span>Reading text…</span>'; pe.wrap.appendChild(b);
  return { set: t => { b.lastChild.textContent = t; }, done: () => b.remove() };
}
async function run(p, regs, lang, force) {
  const b = badge(p), out = [], gone = () => !force && !X.visible.has(p.id); // pages scrolled out of view give up their turn (they are read when they come back)
  try {
    const cv = await X.baseCanvas(p, 3), k = cv.k;
    for (let i = 0; i < regs.length; i++) {
      const r = regs[i], pad = 6, x = clamp(Math.floor(r.x * k) - pad, 0, cv.c.width - 1), y = clamp(Math.floor(r.y * k) - pad, 0, cv.c.height - 1);
      const w = Math.min(cv.c.width - x, Math.ceil(r.w * k) + 2 * pad), h = Math.min(cv.c.height - y, Math.ceil(r.h * k) + 2 * pad); if (w < 24 || h < 16) continue;
      const s = k < 2.6 && w * h < 4e6 ? 3 / k : 1, c = document.createElement('canvas'); c.width = Math.round(w * s); c.height = Math.round(h * s); // small print is enlarged first
      const g = c.getContext('2d'); g.fillStyle = '#fff'; g.fillRect(0, 0, c.width, c.height); g.imageSmoothingQuality = 'high'; g.drawImage(cv.c, x, y, w, h, 0, 0, c.width, c.height);
      const prog = f => b.set(`Reading text… ${Math.round((i + f) / regs.length * 100)}%`);
      let found = linesOf(await job(c, lang, prog, '3', gone), k * s, x / k, y / k);
      if (!found.length && !regs.full && r.h < 160) found = linesOf(await job(c, lang, prog, '6', gone), k * s, x / k, y / k); // a small label: read it as a single block
      out.push(...found);
    }
  } finally { b.done(); }
  const nat = p.src >= 0 ? await X.getLines(p) : []; // real text wins over a reading of the same spot
  return out.filter(o => { const cx = o.x + o.w / 2, cy = o.y + o.h / 2; return !nat.some(l => cx > l.x && cx < l.x + l.w && cy > l.y && cy < l.y + l.h); });
}
const cache = new Map(), keyOf = p => X.baseSig(p) + '|' + ocrLang();
function ocrItems(p, force) { // lines of text recognised in the page's pictures (or, with `force`, on the whole page); remembered per page
  const lang = ocrLang(), key = keyOf(p); let e = cache.get(key);
  if (e && (e.full || !force)) return e.promise;
  e = { full: !!force, lines: null };
  e.promise = (async () => {
    const regs = force ? Object.assign([{ x: 0, y: 0, w: p.w, h: p.h }], { full: true }) : await targets(p);
    if (regs.full) e.full = true; if (!regs.length) return e.lines = [];
    const lines = e.lines = await run(p, regs, lang, force);
    if (!force && lines.length && S.tool === 'edittext') toast(`Recognised ${lines.length} line(s) of text in ${regs.full ? 'this scanned page' : 'the pictures on this page'} — click one to edit it`, 'success');
    return lines;
  })().catch(err => { cache.delete(key); if (err && err.message === 'skip') return []; console.error(err); toast('Text recognition did not work: ' + (err && err.message || err), true); return []; });
  cache.set(key, e); if (cache.size > 80) cache.delete(cache.keys().next().value);
  return e.promise;
}
const ocrCached = p => { const e = cache.get(keyOf(p)); return (e && e.lines) || []; };

/* ───────── matching a typeface to printed text ───────── */
const STD = ['Helvetica', 'Times', 'Courier'], FAMS = [...STD, 'Roboto', 'Open Sans', 'Lato', 'Montserrat', 'Poppins', 'Raleway', 'Nunito', 'Ubuntu', 'Oswald', 'Lora', 'PT Serif', 'Libre Baskerville', 'Playfair Display', 'Crimson Text', 'Roboto Slab', 'Roboto Mono', 'Source Code Pro'], CANDS = [];
for (const font of FAMS) for (const bold of [false, true]) for (const italic of STD.includes(font) ? [false, true] : [false]) CANDS.push({ font, bold, italic });
const fctx = document.createElement('canvas').getContext('2d', { willReadFrequently: true }), tctx = document.createElement('canvas').getContext('2d', { willReadFrequently: true });
const MONO = new Set(['Courier', 'Roboto Mono', 'Source Code Pro']);
const css = (c, px) => `${c.italic ? 'italic ' : ''}${c.bold ? 'bold ' : ''}${px}px ${X.FONTS[c.font]}`;
function inkBox(cv, it) { // where the ink of a line really is, plus its colours
  const { c, k } = cv, px = it.h * 0.12, py = it.h * 0.18, X0 = clamp(Math.floor((it.x - px) * k), 0, c.width - 2), Y0 = clamp(Math.floor((it.y - py) * k), 0, c.height - 2);
  const W = clamp(Math.ceil((it.w + 2 * px) * k), 2, c.width - X0), H = clamp(Math.ceil((it.h + 2 * py) * k), 2, c.height - Y0), d = c.getContext('2d', { willReadFrequently: true }).getImageData(X0, Y0, W, H).data;
  const cnt = {}; let best = 0, bk = 0; for (let i = 0; i < d.length; i += 4) { const q = (d[i] >> 4) << 8 | (d[i + 1] >> 4) << 4 | d[i + 2] >> 4, n = cnt[q] = (cnt[q] || 0) + 1; if (n > best) { best = n; bk = q; } }
  const B = [0, 0, 0]; let bn = 0; for (let i = 0; i < d.length; i += 4) if (((d[i] >> 4) << 8 | (d[i + 1] >> 4) << 4 | d[i + 2] >> 4) === bk) { B[0] += d[i]; B[1] += d[i + 1]; B[2] += d[i + 2]; bn++; }
  for (let q = 0; q < 3; q++) B[q] = Math.round(B[q] / bn);
  const dev = new Uint16Array(W * H); let max = 0; for (let i = 0, j = 0; i < d.length; i += 4, j++) { const v = dev[j] = Math.abs(d[i] - B[0]) + Math.abs(d[i + 1] - B[1]) + Math.abs(d[i + 2] - B[2]); if (v > max) max = v; }
  if (max < 70) return null;
  const thr = Math.max(48, 0.45 * max), rows = new Uint32Array(H), need = Math.max(1, W * 0.004), gapTol = Math.max(1, Math.round(H * 0.06));
  for (let y = 0; y < H; y++) for (let x = 0; x < W; x++) if (dev[y * W + x] > thr) rows[y]++;
  let mid = clamp(Math.round((it.y + it.h * 0.55) * k) - Y0, 0, H - 1);
  for (let o = 0; rows[mid] < need && o < H / 3; o++) { if (rows[clamp(mid + o, 0, H - 1)] >= need) mid = clamp(mid + o, 0, H - 1); else if (rows[clamp(mid - o, 0, H - 1)] >= need) mid = clamp(mid - o, 0, H - 1); }
  if (rows[mid] < need) return null;
  let t = mid, b = mid;
  for (let y = mid - 1, gap = 0; y >= 0; y--) { if (rows[y] >= need) { t = y; gap = 0; } else if (++gap > gapTol) break; }
  for (let y = mid + 1, gap = 0; y < H; y++) { if (rows[y] >= need) { b = y; gap = 0; } else if (++gap > gapTol) break; }
  let l = W, r = -1; const F = [0, 0, 0]; let fn = 0;
  for (let y = t; y <= b; y++) for (let x = 0; x < W; x++) { const v = dev[y * W + x]; if (v > thr) { if (x < l) l = x; if (x > r) r = x; if (v > 0.8 * max) { const i = (y * W + x) * 4; F[0] += d[i]; F[1] += d[i + 1]; F[2] += d[i + 2]; fn++; } } }
  if (r - l < 3 || b - t < 3) return null;
  return { l: X0 + l, t: Y0 + t, w: r - l + 1, h: b - t + 1, B, thr, color: X.hex(...F.map(v => Math.round(v / (fn || 1)))) };
}
function wordSpans(bits, ww, hh) { // runs of columns that hold ink, split where the gap is about a space wide
  const out = [], gap = Math.max(3, Math.round(hh * 0.2)); let s = -1, last = -1;
  for (let x = 0; x < ww; x++) {
    let on = false; for (let y = 0; y < hh; y++) if (bits[y * ww + x]) { on = true; break; }
    if (!on) continue; if (s < 0) s = x; else if (x - last > gap) { out.push([s, last + 1]); s = x; } last = x;
  }
  if (s >= 0) out.push([s, last + 1]); return out;
}
function shapeDiff(A, wa, B, wb, hh) { // share of pixels that differ between two lines drawn at the same height, compared word by word (from each word's left edge) so that letter-spacing does not count against a face
  const sa = wordSpans(A, wa, hh), sb = wordSpans(B, wb, hh), pairs = sa.length === sb.length && sa.length ? sa.map((a, i) => [a, sb[i]]) : [[[0, wa], [0, wb]]]; let xor = 0, n = 0;
  for (const [[a0, a1], [b0, b1]] of pairs) {
    const aw = a1 - a0, bw = b1 - b0, close = Math.abs(aw - bw) <= 0.07 * Math.max(aw, bw), w = close ? aw : Math.max(aw, bw); // a small difference in width is only spacing: stretch over it
    for (let x = 0; x < w; x++) { const ax = a0 + x < a1 ? a0 + x : -1, bx = close ? b0 + Math.min(bw - 1, Math.floor((x + 0.5) * bw / aw)) : b0 + x < b1 ? b0 + x : -1; for (let y = 0; y < hh; y++) if ((ax < 0 ? 0 : A[y * wa + ax]) !== (bx < 0 ? 0 : B[y * wb + bx])) xor++; }
    n += w * hh;
  }
  return xor / n + (pairs.length === sa.length ? 0 : 0.04);
}
function measure(c, str) { fctx.font = css(c, 100); fctx.fontKerning = 'none'; const m = fctx.measureText(str); return { A: m.actualBoundingBoxAscent, D: m.actualBoundingBoxDescent, L: m.actualBoundingBoxLeft, R: m.actualBoundingBoxRight }; }
async function fitLine(cv, it, known) { // → { font, bold, italic, size, x, base, color } for a printed line, or null when no ink is found there
  await Promise.all(FAMS.map(X.ensureFont));
  const str = it.str.trim(), ink = str && inkBox(cv, it); if (!ink) return null;
  const k = cv.k, hh = 48, ww = clamp(Math.round(ink.w * hh / ink.h), 8, 2000);
  let pick = known ? { font: known.font, bold: known.bold, italic: known.italic } : null;
  if (!pick) { // compare the ink with each candidate face drawn into the same box
    const tc = tctx.canvas; tc.width = ww; tc.height = hh; tctx.imageSmoothingQuality = 'high'; tctx.drawImage(cv.c, ink.l, ink.t, ink.w, ink.h, 0, 0, ww, hh);
    const td = tctx.getImageData(0, 0, ww, hh).data, mask = new Uint8Array(ww * hh);
    for (let i = 0, j = 0; j < mask.length; i += 4, j++) mask[j] = Math.abs(td[i] - ink.B[0]) + Math.abs(td[i + 1] - ink.B[1]) + Math.abs(td[i + 2] - ink.B[2]) > ink.thr * 0.7 ? 1 : 0;
    const fc = fctx.canvas; let bestScore = 1e9;
    for (const c of CANDS) {
      const m = measure(c, str), iw = m.L + m.R, ih = m.A + m.D; if (iw <= 0 || ih <= 0) continue;
      const sc = hh / ih, cw = clamp(Math.ceil(iw * sc) + 1, 8, 2600); fc.width = cw; fc.height = hh; // same height as the ink; the width is whatever this face needs
      fctx.fillStyle = '#fff'; fctx.fillRect(0, 0, cw, hh);
      fctx.setTransform(sc, 0, 0, sc, m.L * sc, m.A * sc); fctx.fillStyle = '#000'; fctx.font = css(c, 100); fctx.fontKerning = 'none'; fctx.fillText(str, 0, 0); fctx.setTransform(1, 0, 0, 1, 0, 0);
      const cd = fctx.getImageData(0, 0, cw, hh).data, bits = new Uint8Array(cw * hh); for (let i = 0, j = 0; j < bits.length; i += 4, j++) bits[j] = cd[i] < 128 ? 1 : 0;
      const score = shapeDiff(mask, ww, bits, cw, hh) + 0.2 * Math.abs(Math.log((ink.w / iw) / (ink.h / ih))) + (MONO.has(c.font) ? 0.03 : 0); // (a typewriter face has to earn its place)
      if (score < bestScore) { bestScore = score; pick = c; }
    }
    fctx.setTransform(1, 0, 0, 1, 0, 0);
    if (!pick) return null;
  }
  const m = measure(pick, str), iw = m.L + m.R, ih = m.A + m.D; if (iw <= 0 || ih <= 0) return null;
  const sh = ink.h / k / ih * 100, sw = ink.w / k / iw * 100, size = Math.round((Math.abs(Math.log(sw / sh)) < 0.15 ? Math.sqrt(sh * sw) : sh) * 10) / 10;
  if (!(size >= 3 && size < 600)) return null;
  return { ...pick, size, x: ink.l / k + m.L * size / 100, base: (ink.t + ink.h) / k - m.D * size / 100, color: ink.color };
}

/* ───────── UI ───────── */
X.ACT.ocr = async () => { // read the whole current page, whatever it contains
  const p = S.pages[S.cur]; if (!p) return; if (S.tool !== 'edittext') X.setTool('edittext');
  const n = (await ocrItems(p, true)).length; X.drawTI(p);
  toast(n ? `Recognised ${n} line(s) of text on this page — click one to edit it` : 'No more text was found on this page', n ? 'success' : 'info');
};
const sel = $('#pOcrLang');
if (sel) {
  sel.innerHTML = Object.entries(LANGS).map(([v, l]) => `<option value="${v}">${l}</option>`).join(''); sel.value = ocrLang();
  sel.addEventListener('change', () => { try { localStorage.setItem('xd3.ocr', sel.value); } catch (e) { /* not remembered */ } if (S.pages.length) X.setTool('edittext'); });
  sel.addEventListener('keydown', e => e.stopPropagation());
}
Object.assign(X, { ocrItems, ocrCached, fitLine });
})();
