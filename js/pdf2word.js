/* XD3 PDF Editor — PDF → Word, "exact" mode: a flowing, editable Word document laid out so every line lands where it is in the PDF.
   It rests on Word layout rules measured through Word itself (Word 2013+, compatibilityMode 15):
   - exact line spacing L puts the baseline 0.8·L below the line top, so a paragraph's first baseline = last baseline + 0.2·L(prev) + before + 0.8·L;
   - a table starts at the previous line box bottom, and the next paragraph starts at the table bottom;
   - left / centred / right lines break where the summed letter widths pass the room; a justified line may squeeze each space by about 20%.
   One Word section per PDF page. The page's drawing (shapes, pictures) sits behind the text in that section's own header,
   so editing the body never loses it; rules under text become underlines. Rotated text is the only text box. */
(() => {
'use strict';
const X = window.XD3;
const cv = () => X._cv; // shared helpers from convert.js
const S = () => X.S;
const clamp = (v, a, b) => Math.min(b, Math.max(a, v));
const twip = v => Math.round(v * 20), emu = v => Math.round(v * 12700);
const xesc = s => String(s).replace(/[\x00-\x08\x0b\x0c\x0e-\x1f\ufffe\uffff]/g, '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const median = a => { if (!a.length) return 0; const s = [...a].sort((x, y) => x - y); return s[s.length >> 1]; };
const fnv = u8 => { let h = 0x811c9dc5; for (let i = 0; i < u8.length; i++) { h ^= u8[i]; h = Math.imul(h, 0x01000193); } return (h >>> 0).toString(36) + '.' + u8.length; };
const toBlob = (c, type, q) => new Promise(res => c.toBlob(res, type, q));
const tick = () => new Promise(r => setTimeout(r, 0));
const say = t => { const e = document.getElementById('busyTxt'); if (e) e.textContent = t; };
const K = 2.5; // render scale (180 dpi) for text colours and the page drawing
const SQUEEZE = 0.25; // Word may shrink each space of a justified line by up to this share of its width (measured: 0.24–0.25; less on lines with many spaces)
const GHOST = '<w14:textFill><w14:solidFill><w14:srgbClr w14:val="000000"><w14:alpha w14:val="100000"/></w14:srgbClr></w14:solidFill></w14:textFill>';
const LIG = { '\ufb00': 'ff', '\ufb01': 'fi', '\ufb02': 'fl', '\ufb03': 'ffi', '\ufb04': 'ffl', '\ufb05': 'st', '\ufb06': 'st' };
const unlig = s => s.replace(/[\ufb00-\ufb06]/g, c => LIG[c]);
const BUL = /^[•●○◦▪■□►▸‣⁃∙·\-–—*✓✔➢➤\uf0b7\uf0a7\uf076\uf0d8\uf0fc\uf0a8]$/, NUMM = /^\(?(\d{1,3}|[a-zA-Z])[.)]$/;

/* ───────── fonts: the face Word will use, and a face the canvas can measure it with ───────── */
const OFFICE_OF = { Carlito: 'Calibri', Caladea: 'Cambria', Gelasio: 'Georgia', 'DejaVu Sans': 'Verdana' }; // free metric twins → the Office name (never embedded)
const normName = s => String(s || '').toLowerCase().replace(/[^a-z0-9]/g, '');
function bundledFamily(ps) {
  const raw = String(ps || '').replace(/^[A-Z]{6}\+/, ''), base = raw.split(/[-,]/)[0].replace(/(PS)?MT$|PS$/, '');
  const bare = base.replace(/(Bold|Italic|Oblique|Regular|Book|Normal|Medium|Light|Semi[Bb]old|Black)+$/, '') || base, n = normName(bare);
  return n ? Object.keys(X.BUNDLED).find(k => normName(k) === n) || null : null;
}
const fontMemo = new Map();
function fontOf(ps, cls) { // → { name: Word's font, meas: CSS family measuring like it, exact, weight, bundle: BUNDLED face to load, embed: family to embed }
  const key = ps + '|' + cls; if (fontMemo.has(key)) return fontMemo.get(key);
  const { wordFont, hasFont } = cv(), wf = wordFont(ps, cls), q = n => `"${n.replace(/["\\]/g, '')}"`, fam = bundledFamily(ps); let r;
  const own = fam && !OFFICE_OF[fam] && X.BUNDLED[fam][3] !== 'Popular' ? fam : null; // a free family Word won't have: embed it
  if (wf.exact) r = { name: wf.name, meas: q(wf.name), exact: true, weight: wf.weight, embed: own && normName(own) === normName(wf.name) ? own : null };
  else if (fam) { const off = OFFICE_OF[fam] || (X.BUNDLED[fam][3] === 'Popular' ? fam : null); r = { name: off || fam, meas: q('XD3 ' + fam), exact: true, bundle: fam, embed: off ? null : fam }; }
  else r = { name: wf.name, meas: hasFont(wf.name) || !X.BUNDLED[wf.name] ? q(wf.name) : q('XD3 ' + wf.name), exact: false, bundle: !hasFont(wf.name) && X.BUNDLED[wf.name] ? wf.name : null };
  fontMemo.set(key, r); return r;
}
const probe = document.createElement('canvas').getContext('2d'); try { probe.fontKerning = 'none'; probe.textRendering = 'optimizeSpeed'; } catch (e) { /* older Chromium */ }
let probeFont = ''; const wcache = new Map();
const cssOf = st => `${st.i ? 'italic ' : ''}${st.b ? 'bold ' : ''}100px ${st.meas}`;
const wsz = s => Math.max(1, Math.round(s * 2) / 2); // Word keeps sizes in half points
function w100(text, css) { const k = css + '\u0001' + text; let w = wcache.get(k); if (w == null) { if (probeFont !== css) { probe.font = css; probeFont = css; } w = probe.measureText(text).width; if (wcache.size > 300000) wcache.clear(); wcache.set(k, w); } return w; }
const measure = (text, st) => w100(text, cssOf(st)) * wsz(st.size) / 100; // Word's natural width (no kerning, no ligatures)
const loaded = new Set();
async function loadFonts(sts) { // the canvas measures a face only once it has loaded (before that it silently uses a fallback)
  for (const st of sts) {
    const css = cssOf(st); if (loaded.has(css)) continue; loaded.add(css);
    if (st.bundle) { try { await X.ensureFont(st.bundle); } catch (e) { /* fallback measures */ } }
    try { await document.fonts.load(css.replace('100px', '24px'), 'AaZz09'); } catch (e) { /* system font */ }
  }
}

/* ───────── a page: text pieces with font, colour, and the page's drawing without text ───────── */
async function noTextRender(p) { // the page without its text — also text inside transparency groups, which pdf.js draws on canvases of its own
  let cf = null; try { cf = p.src >= 0 ? (await X.getPdfPage(p))._transport.canvasFactory : null; } catch (e) { cf = null; }
  if (!cf || typeof cf.create !== 'function') return X.renderPageImage(p, K, { noText: true, annots: false });
  const orig = cf.create; cf.create = function (w, h) { const r = orig.call(this, w, h); try { r.context.fillText = r.context.strokeText = () => {}; } catch (e) { /* plain canvas */ } return r; };
  try { return await X.renderPageImage(p, K, { noText: true, annots: false }); } finally { cf.create = orig; }
}
const SANS = '100px Arial, Helvetica, sans-serif';
function originOf(q) { // where a piece's line starts on the page (a turned piece is kept unturned around its pivot)
  if (!q.rot || !q.pv) return [q.x, q.base];
  const a = q.rot * Math.PI / 180, dx = q.x - q.pv[0], dy = q.base - q.pv[1];
  return [q.pv[0] + dx * Math.cos(a) - dy * Math.sin(a), q.pv[1] + dx * Math.sin(a) + dy * Math.cos(a)];
}
function uncover(s, covers) { // the letters under a cover (black-out, white-out) are left out: they must not come back in Word
  if (!covers.length) return [s];
  const [ox, oy] = originOf(s), a = (s.rot || 0) * Math.PI / 180, ca = Math.cos(a), sa = Math.sin(a), up = 0.3 * s.size;
  const xs = [ox, ox + s.w * ca], ys = [oy, oy + s.w * sa], x0 = Math.min(...xs) - s.size, x1 = Math.max(...xs) + s.size, y0 = Math.min(...ys) - s.size, y1 = Math.max(...ys) + s.size;
  const near = covers.map(o => X.bbox(o)).filter(b => b.x < x1 && b.x + b.w > x0 && b.y < y1 && b.y + b.h > y0); if (!near.length) return [s];
  const ch = [...s.str], ws = ch.map(c => w100(c, SANS) || 50), k = s.w / (ws.reduce((t, v) => t + v, 0) || 1), out = []; let cur = null, cum = 0;
  ch.forEach((c, i) => {
    const w = ws[i] * k, m = cum + w / 2, x = ox + m * ca + up * sa, y = oy + m * sa - up * ca;
    if (near.some(b => x > b.x && x < b.x + b.w && y > b.y && y < b.y + b.h)) cur = null;
    else { if (!cur) { cur = { ...s, str: '', x: s.x + cum, w: 0 }; out.push(cur); } cur.str += c; cur.w += w; }
    cum += w;
  });
  return out.filter(o => o.str.trim());
}
function turnPiece(s, p, prot) { // the page as stored → the page as shown, when the user has turned it
  if (!prot) return s;
  const view = prot === 90 ? (x, y) => [p.h - y, x] : prot === 180 ? (x, y) => [p.w - x, p.h - y] : (x, y) => [y, p.w - x];
  const pv = s.pv || [s.x, s.base], nv = view(pv[0], pv[1]);
  let r = ((s.rot || 0) + prot) % 360; if (r > 180) r -= 360; if (r <= -180) r += 360; if (Math.abs(r) < 0.05) r = 0;
  return { ...s, x: s.x - pv[0] + nv[0], base: s.base - pv[1] + nv[1], rot: r, pv: r ? nv : null };
}
function untrack(str, gl, x, base, w, size) { // pdf.js puts a space wherever letters stand apart, so tracked capitals read "S K I L L S": the glyphs say where the word gaps are
  const g = gl.filter(q => Math.abs(q.oy - base) < 0.2 * size && q.cx > x - 0.5 && q.cx < x + w + 0.5 && Math.abs(q.size - size) < 0.2 * size).sort((a, b) => a.ox - b.ox);
  const letters = g.filter(q => q.uni && q.uni.trim()), flat = s => unlig(s).replace(/\s+/g, '');
  if (letters.length < 3 || flat(letters.map(q => q.uni).join('')) !== flat(str)) return str;
  const LN = /[\p{L}\p{N}]/u, gaps = []; for (let i = 1; i < letters.length; i++) { const a = letters[i - 1], b = letters[i]; gaps.push({ d: b.ox - a.ex, sp: g.some(q => !(q.uni && q.uni.trim()) && q.ox >= a.ox && q.ox <= b.ox), ln: LN.test(a.uni) && LN.test(b.uni) }); }
  const lg = gaps.filter(q => !q.sp && q.ln); if (lg.length < 2) return str; const med = median(lg.map(q => q.d)); // (the usual gap between letters: dot leaders do not count)
  let out = unlig(letters[0].uni); for (let i = 1; i < letters.length; i++) out += (gaps[i - 1].sp || gaps[i - 1].d > med + 0.15 * size ? ' ' : '') + unlig(letters[i].uni);
  return str.match(/^\s*/)[0] + out + str.match(/\s*$/)[0];
}
async function pageModel(p) {
  const [PW, PH] = X.rdims(p);
  const bg = await noTextRender(p), ref = await X.renderPageImage(p, K, { annots: false });
  const W = bg.width, H = bg.height, bx = bg.getContext('2d', { willReadFrequently: true }), bd = bx.getImageData(0, 0, W, H), d = bd.data, rd = ref.getContext('2d', { willReadFrequently: true }).getImageData(0, 0, W, H).data;
  if (p.src >= 0) { // form fields and other annotations keep their look as part of the drawing
    let an = []; try { an = await (await X.getPdfPage(p)).getAnnotations({ intent: 'print' }); } catch (e) { /* none */ }
    if (an.length) {
      const fd = (await X.renderPageImage(p, K)).getContext('2d', { willReadFrequently: true }).getImageData(0, 0, W, H).data; let ch = 0;
      for (let i = 0; i < d.length; i += 4) if (fd[i] !== rd[i] || fd[i + 1] !== rd[i + 1] || fd[i + 2] !== rd[i + 2]) { d[i] = fd[i]; d[i + 1] = fd[i + 1]; d[i + 2] = fd[i + 2]; ch++; }
      if (ch) bx.putImageData(bd, 0, 0);
    }
  }
  const pieces = [], fonts = {}, covers = p.objs.filter(o => o.cover); let type3 = 0, chars = 0;
  const prot = (((p.rot || 0) % 360) + 360) % 360, segs = [];
  if (p.src >= 0) {
    const pg = await X.getPdfPage(p), vt = pg.getViewport({ scale: 1 }).transform, tc = await pg.getTextContent();
    let A = null; if (X.native && tc.items.some(it => /(^|\s)\S{1,2} \S{1,2} \S{1,2}(\s|$)/.test(it.str))) { try { A = await X.native.analyze(p); } catch (e) { A = null; } } // (glyph positions: to undo spaces pdf.js puts between tracked letters)
    for (const it of tc.items) {
      if (!it.str || !it.str.trim() || !(it.width > 0) || (tc.styles[it.fontName] || {}).vertical) continue;
      const m = pdfjsLib.Util.transform(vt, it.transform), size = Math.hypot(m[2], m[3]), ang = Math.atan2(m[1], m[0]); if (size < 1) continue; // (text under 1 pt is a hiding trick)
      const rot = Math.abs(ang) > 0.02 ? Math.round(ang * 1800 / Math.PI) / 10 : 0, c = [m[4] + Math.cos(ang) * it.width / 2 + Math.sin(ang) * 0.3 * size, m[5] + Math.sin(ang) * it.width / 2 - Math.cos(ang) * 0.3 * size];
      if (X.covered(p, { x: c[0] - 0.5, y: c[1] - 0.5, w: 1, h: 1 })) continue; // replaced by an edit
      let str = unlig(it.str); if (A && !rot && / /.test(str.trim())) str = untrack(str, A.glyphs, m[4], m[5], it.width, size);
      if (/^\s*\p{L}{1,2}( \p{L}{1,2}){3,}\s*$/u.test(str) && str.trim().split(' ').filter(t => t.length === 1).length * 2 >= str.trim().split(' ').length) str = str.replace(/(\S) (?=\S)/g, '$1'); // (still spaced out, its glyphs out of reach: most such lines are one word)
      segs.push(...uncover({ str, x: m[4], base: m[5], w: it.width, size, rot, pv: rot ? [m[4], m[5]] : null, fontName: it.fontName }, covers));
    }
  }
  for (const l of X.ocrCached ? X.ocrCached(p) : []) if (l.str && l.str.trim() && !X.covered(p, l)) segs.push(...uncover({ str: l.str, x: l.x, base: l.base, w: l.w, size: l.size, rot: 0, pv: null, ocr: true }, covers));
  for (const s0 of segs) {
    const s = turnPiece(s0, p, prot), f = s.fontName ? await (fonts[s.fontName] || (fonts[s.fontName] = X.origFont(p, { fontName: s.fontName }, true).catch(() => ({ font: 'Helvetica' })))) : { font: 'Helvetica' };
    const fi = fontOf(f.ps || f.real || '', f.font); if (f.type3) type3 += s.str.length; chars += s.str.length;
    pieces.push({ str: s.str, x: s.x, base: s.base, w: s.w, size: s.size, rot: s.rot, pv: s.pv, font: fi.name, exact: fi.exact, bold: !!f.bold && !fi.weight, italic: !!f.italic, fi, sample: true });
  }
  p.objs.forEach((o, oi) => { // text added in the editor (covers drawn after it hide part of it)
    if (o.type !== 'text' || !o.text || !o.text.trim()) return;
    const xf = S().xfonts[o.font], BF = X.BUNDLED[o.font], cls = xf ? xf.fb : BF ? ({ serif: 'Times', monospace: 'Courier' }[BF[2]] || 'Helvetica') : o.font;
    const fi = fontOf(xf ? xf.ps || xf.name : BF ? o.font : o.font, cls), pv = o.rot ? X.center(X.bbox(o)) : null, cv2 = p.objs.slice(oi + 1).filter(c => c.cover);
    for (const l of X.textLines(o)) if (l.ln.trim()) for (const s0 of uncover({ str: l.ln, x: l.x, base: l.y, w: l.w, size: o.size, rot: o.rot || 0, pv }, cv2)) {
      const s = turnPiece(s0, p, prot);
      pieces.push({ str: s.str, x: s.x, base: s.base, w: s.w, size: o.size, rot: s.rot, pv: s.pv, font: fi.name, exact: fi.exact, bold: !!o.bold && !fi.weight, italic: !!o.italic, underline: !!o.underline, strike: !!o.strike, color: o.color, fi, sample: false });
    }
  });
  const hex3 = c => '#' + c.map(v => Math.round(v).toString(16).padStart(2, '0')).join('');
  for (const q of pieces) { // colour of the text: what the page with it has and the page without it lacks (nothing → hidden text)
    if (!q.sample) continue;
    let md = 0; const hits = [];
    const at = i => { const dd = Math.abs(rd[i] - d[i]) + Math.abs(rd[i + 1] - d[i + 1]) + Math.abs(rd[i + 2] - d[i + 2]); if (dd > 40) hits.push(dd, i); if (dd > md) md = dd; };
    if (q.rot) { // turned text: walk its box along the line
      const [ox, oy] = originOf(q), a = q.rot * Math.PI / 180, ca = Math.cos(a), sa = Math.sin(a), st = 1 / K, seen = new Set();
      for (let u = 0; u <= q.w; u += st) for (let v = -0.2 * q.size; v <= 0.8 * q.size; v += st) {
        const x = Math.round((ox + u * ca + v * sa) * K), y = Math.round((oy + u * sa - v * ca) * K); if (x < 0 || y < 0 || x >= W || y >= H) continue;
        const i = (y * W + x) * 4; if (!seen.has(i)) { seen.add(i); at(i); }
      }
    } else {
      const xa = clamp(Math.floor(q.x * K), 0, W - 1), xb = clamp(Math.ceil((q.x + q.w) * K), 0, W - 1), ya = clamp(Math.floor((q.base - 0.8 * q.size) * K), 0, H - 1), yb = clamp(Math.ceil((q.base + 0.2 * q.size) * K), 0, H - 1);
      for (let y = ya; y <= yb; y++) for (let x = xa, i = (y * W + xa) * 4; x <= xb; x++, i += 4) { const dd = Math.abs(rd[i] - d[i]) + Math.abs(rd[i + 1] - d[i + 1]) + Math.abs(rd[i + 2] - d[i + 2]); if (dd > 40) hits.push(dd, i); if (dd > md) md = dd; } // (inline: the hot loop)
    }
    if (md < 60) { q.ghost = true; continue; }
    let r = 0, g = 0, bl = 0, n = 0; for (let k = 0; k < hits.length; k += 2) if (hits[k] >= 0.8 * md) { const i = hits[k + 1]; r += rd[i]; g += rd[i + 1]; bl += rd[i + 2]; n++; }
    const c = [r / n, g / n, bl / n].map(Math.round); q.color = Math.max(...c) < 70 ? '#000000' : hex3(c);
  }
  return { p, PW, PH, W, H, bg, d, pieces, type3: chars ? type3 / chars : 0 };
}

/* ───────── the page's drawing → rectangles (shapes) and picture patches, in points ───────── */
function drawing(M) {
  const { W, H, d } = M, N = W * H, key = new Int32Array(N); let any = 0;
  for (let i = 0, j = 0; i < N; i++, j += 4) { const r = d[j], g = d[j + 1], b = d[j + 2]; if (r > 249 && g > 249 && b > 249) continue; key[i] = 1 + ((r >> 3) << 10 | (g >> 3) << 5 | (b >> 3)); any++; }
  if (!any) return { rects: [], pics: [] };
  const lab = new Int32Array(N), stack = new Int32Array(N), C = [null];
  for (let s = 0; s < N; s++) { // parts of one colour that touch
    const k = key[s]; if (!k || lab[s]) continue;
    const id = C.length; let sp = 0, x0 = W, y0 = H, x1 = -1, y1 = -1, n = 0; stack[sp++] = s; lab[s] = id;
    while (sp) {
      const i = stack[--sp], x = i % W, y = (i - x) / W; n++; if (x < x0) x0 = x; if (x > x1) x1 = x; if (y < y0) y0 = y; if (y > y1) y1 = y;
      if (x > 0 && !lab[i - 1] && key[i - 1] === k) { lab[i - 1] = id; stack[sp++] = i - 1; }
      if (x < W - 1 && !lab[i + 1] && key[i + 1] === k) { lab[i + 1] = id; stack[sp++] = i + 1; }
      if (y > 0 && !lab[i - W] && key[i - W] === k) { lab[i - W] = id; stack[sp++] = i - W; }
      if (y < H - 1 && !lab[i + W] && key[i + W] === k) { lab[i + W] = id; stack[sp++] = i + W; }
    }
    C.push({ id, x0, y0, x1, y1, n, seed: s, rect: null });
    if (C.length > 400000) return { rects: [], pics: [bboxPic(key, W, H)] }; // a photo-like page: one picture
  }
  const col = c => { const j = c.seed * 4; return [d[j], d[j + 1], d[j + 2]]; };
  const lum = c => 0.3 * c[0] + 0.59 * c[1] + 0.11 * c[2];
  const rects = [], resid = new Uint8Array(C.length);
  for (let id = 1; id < C.length; id++) {
    const c = C[id], w = c.x1 - c.x0 + 1, h = c.y1 - c.y0 + 1, A = w * h;
    if (c.n >= 0.96 * A && (Math.min(w, h) >= 4 || Math.max(w, h) >= 16)) { c.rect = { x0: c.x0, y0: c.y0, x1: c.x1 + 1, y1: c.y1 + 1, col: col(c), comp: id, geom: 'rect' }; rects.push(c.rect); continue; } // (specks of a photo go into its picture)
    if (w >= 6 && h >= 6 && c.n >= 0.74 * A && rounded(c, w, h)) { rects.push(c.rect); continue; }
    if (Math.max(w, h) >= 16 && c.n >= 12 && lineSegs(c, w, h)) continue;
    resid[id] = 1;
  }
  function rounded(c, w, h) { // a filled rounded box or an ellipse: corners empty, middle row and column full
    const at = (x, y) => lab[y * W + x] === c.id, my = (c.y0 + c.y1) >> 1, mx = (c.x0 + c.x1) >> 1;
    if (at(c.x0, c.y0) || at(c.x1, c.y0) || at(c.x0, c.y1) || at(c.x1, c.y1)) return false;
    for (let x = c.x0; x <= c.x1; x++) if (!at(x, my)) return false;
    for (let y = c.y0; y <= c.y1; y++) if (!at(mx, y)) return false;
    let r = 0; while (r < w / 2 && !at(c.x0 + r, c.y0)) r++;
    let rv = 0; while (rv < h / 2 && !at(c.x0, c.y0 + rv)) rv++;
    const ell = r > 0.3 * w && rv > 0.3 * h && Math.abs(w - h) > 2;
    c.rect = { x0: c.x0, y0: c.y0, x1: c.x1 + 1, y1: c.y1 + 1, col: col(c), comp: c.id, geom: ell ? 'ellipse' : 'roundRect', adj: Math.round(clamp(Math.max(r, rv) / Math.min(w, h), 0, 0.5) * 100000) };
    return true;
  }
  function lineSegs(c, w, h) { // a frame or a grid: the part's pixels all lie on long horizontal or vertical lines
    const rc = new Int32Array(h), cc = new Int32Array(w);
    for (let y = c.y0; y <= c.y1; y++) for (let x = c.x0, i = y * W + c.x0; x <= c.x1; x++, i++) if (lab[i] === c.id) { rc[y - c.y0]++; cc[x - c.x0]++; }
    const hr = w >= 8 ? Array.from(rc, v => v >= 0.85 * w) : [], vc = h >= 8 ? Array.from(cc, v => v >= 0.85 * h) : [];
    let cov = 0; for (let y = c.y0; y <= c.y1; y++) for (let x = c.x0, i = y * W + c.x0; x <= c.x1; x++, i++) if (lab[i] === c.id && (hr[y - c.y0] || vc[x - c.x0])) cov++;
    if (cov < 0.97 * c.n) return false;
    const cl = col(c), out = [];
    for (let a = 0; a < h; a++) if (hr[a]) { let b = a; while (b + 1 < h && hr[b + 1]) b++; out.push({ x0: c.x0, y0: c.y0 + a, x1: c.x1 + 1, y1: c.y0 + b + 1, col: cl, comp: c.id, geom: 'rect' }); a = b; }
    for (let a = 0; a < w; a++) if (vc[a]) { let b = a; while (b + 1 < w && vc[b + 1]) b++; out.push({ x0: c.x0 + a, y0: c.y0, x1: c.x0 + b + 1, y1: c.y1 + 1, col: cl, comp: c.id, geom: 'rect' }); a = b; }
    rects.push(...out); return true;
  }
  // anti-aliased edges: a thin, lighter strip along a rectangle's side is part of it (extended by the strip's coverage)
  const cell = 24, grid = new Map(), cellsOf = r => { const a = []; for (let gy = Math.floor((r.y0 - 2) / cell); gy <= Math.floor((r.y1 + 2) / cell); gy++) for (let gx = Math.floor((r.x0 - 2) / cell); gx <= Math.floor((r.x1 + 2) / cell); gx++) a.push(gy * 100000 + gx); return a; };
  rects.forEach((r, i) => { for (const k of cellsOf(r)) { if (!grid.has(k)) grid.set(k, []); grid.get(k).push(i); } });
  const dead = new Uint8Array(rects.length);
  rects.forEach((t, ti) => {
    const tw = t.x1 - t.x0, th = t.y1 - t.y0; if (Math.min(tw, th) > 2 || Math.max(tw, th) < 3) return;
    const lt = lum(t.col); let best = null;
    for (const k of cellsOf(t)) for (const bi of grid.get(k) || []) {
      if (bi === ti || dead[bi]) continue; const b = rects[bi]; if (lum(b.col) >= lt - 6) continue;
      if (th <= 2 && tw >= 3 && (t.y1 === b.y0 || t.y0 === b.y1) && Math.min(t.x1, b.x1) - Math.max(t.x0, b.x0) >= 0.8 * tw) best = b;
      else if (tw <= 2 && th >= 3 && (t.x1 === b.x0 || t.x0 === b.x1) && Math.min(t.y1, b.y1) - Math.max(t.y0, b.y0) >= 0.8 * th) best = b;
      if (best) break;
    }
    if (!best) return;
    { // a real blend lies between the fill and what is on its other side; a lighter strip with the same fill on both sides is a rule
      const hz = th <= 2 && tw >= 3, ox = hz ? (t.x0 + t.x1) >> 1 : t.x1 === best.x0 ? t.x0 - 1 : t.x1, oy = hz ? (t.y1 === best.y0 ? t.y0 - 1 : t.y1) : (t.y0 + t.y1) >> 1;
      const j = (oy * W + ox) * 4, opp = ox < 0 || oy < 0 || ox >= W || oy >= H ? [255, 255, 255] : [d[j], d[j + 1], d[j + 2]];
      if (lum(opp) < lt - 3) return;
    }
    const lb = lum(best.col), a = lb < 245 ? clamp((255 - lt) / (255 - lb), 0, 1) : 0.5;
    if (th <= 2 && tw >= 3) { if (t.y1 === best.y0) best.y0 -= th * a; else best.y1 += th * a; } else { if (t.x1 === best.x0) best.x0 -= tw * a; else best.x1 += tw * a; }
    dead[ti] = 1; if (t.comp && C[t.comp]) C[t.comp].absorbed = true;
  });
  // gradients: neighbouring rectangles of nearly the same colour are one smooth fill → a picture
  const near = (a, b) => Math.abs(a[0] - b[0]) <= 14 && Math.abs(a[1] - b[1]) <= 14 && Math.abs(a[2] - b[2]) <= 14 && (a[0] !== b[0] || a[1] !== b[1] || a[2] !== b[2]);
  const up = rects.map((r, i) => i), root = i => { while (up[i] !== i) i = up[i] = up[up[i]]; return i; };
  rects.forEach((r, ri) => {
    if (dead[ri]) return;
    for (const k of cellsOf(r)) for (const bi of grid.get(k) || []) {
      if (bi === ri || dead[bi]) continue; const b = rects[bi];
      const touchY = Math.abs(r.y1 - b.y0) <= 2 || Math.abs(b.y1 - r.y0) <= 2, touchX = Math.abs(r.x1 - b.x0) <= 2 || Math.abs(b.x1 - r.x0) <= 2;
      const ovx = Math.min(r.x1, b.x1) - Math.max(r.x0, b.x0), ovy = Math.min(r.y1, b.y1) - Math.max(r.y0, b.y0);
      if (near(r.col, b.col) && ((touchY && ovx > 0.8 * Math.max(r.x1 - r.x0, b.x1 - b.x0)) || (touchX && ovy > 0.8 * Math.max(r.y1 - r.y0, b.y1 - b.y0)))) up[root(ri)] = root(bi); // (bands of a gradient lie side by side along their length; a rule touching a frame's side is not one)
    }
  });
  const gsize = new Map(); rects.forEach((r, i) => { if (!dead[i]) { const k = root(i); gsize.set(k, (gsize.get(k) || 0) + 1); } });
  rects.forEach((r, ri) => { if (!dead[ri] && gsize.get(root(ri)) >= 3) { dead[ri] = 1; resid[r.comp] = 1; } }); // (three or more bands: a gradient)
  // what is left (curves, photos, text drawn as shapes, gradients) → picture patches cut from the drawing
  const G = 2, gw = Math.ceil(W / G), gh = Math.ceil(H / G), on = new Uint8Array(gw * gh), cnt = new Uint32Array(gw * gh);
  let nres = 0; for (let i = 0; i < N; i++) { const l = lab[i]; if (l && resid[l] && !(C[l].absorbed)) { const x = i % W, y = (i - x) / W, g = (y / G | 0) * gw + (x / G | 0); on[g] = 1; cnt[g]++; nres++; } }
  const pics = [];
  if (nres) {
    const R = Math.max(1, Math.round(3 * K / G)), dil = new Uint8Array(gw * gh); // things closer than ~3 pt go in one picture
    for (let y = 0; y < gh; y++) for (let x = 0; x < gw; x++) if (on[y * gw + x]) for (let yy = Math.max(0, y - R); yy <= Math.min(gh - 1, y + R); yy++) { const row = yy * gw; for (let xx = Math.max(0, x - R); xx <= Math.min(gw - 1, x + R); xx++) dil[row + xx] = 1; }
    const seen = new Uint8Array(gw * gh), st = [];
    for (let s0 = 0; s0 < gw * gh; s0++) {
      if (!dil[s0] || seen[s0]) continue; let x0 = gw, y0 = gh, x1 = -1, y1 = -1, px = 0; st.push(s0); seen[s0] = 1;
      while (st.length) { const s = st.pop(), x = s % gw, y = (s - x) / gw; if (on[s]) { px += cnt[s]; if (x < x0) x0 = x; if (x > x1) x1 = x; if (y < y0) y0 = y; if (y > y1) y1 = y; }
        for (const n of [s - 1, s + 1, s - gw, s + gw]) if (n >= 0 && n < gw * gh && !seen[n] && dil[n] && Math.abs((n % gw) - x) <= 1) { seen[n] = 1; st.push(n); } }
      if (x1 < 0 || px < 10) continue; // (stray anti-aliasing specks)
      pics.push({ x0: Math.max(0, x0 * G - 1), y0: Math.max(0, y0 * G - 1), x1: Math.min(W, (x1 + 1) * G + 1), y1: Math.min(H, (y1 + 1) * G + 1) });
    }
  }
  const pt = r => ({ x: r.x0 / K, y: r.y0 / K, w: (r.x1 - r.x0) / K, h: (r.y1 - r.y0) / K });
  const shown = (r, i) => !dead[i] && (r.x1 - r.x0) * (r.y1 - r.y0) >= 3 && !pics.some(q => r.x0 >= q.x0 && r.x1 <= q.x1 && r.y0 >= q.y0 && r.y1 <= q.y1); // (a shape under a picture patch is hidden by it)
  const hex = r => '#' + r.col.map(v => v.toString(16).padStart(2, '0')).join(''), inPic = r => pics.findIndex(q => r.x0 >= q.x0 && r.x1 <= q.x1 && r.y0 >= q.y0 && r.y1 <= q.y1);
  return {
    rects: rects.filter(shown).map(r => ({ ...pt(r), fill: hex(r), geom: r.geom, adj: r.adj })),
    pics: pics.map(r => ({ ...pt(r), px: r })),
    hidden: rects.map((r, i) => (!dead[i] && (r.x1 - r.x0) * (r.y1 - r.y0) >= 3 ? [r, inPic(r)] : null)).filter(v => v && v[1] >= 0).map(([r, k]) => ({ ...pt(r), fill: hex(r), geom: r.geom, adj: r.adj, pk: k, px: r })), // (shapes inside a picture patch: a button under its label may come out of it)
  };
}
function bboxPic(key, W, H) { let x0 = W, y0 = H, x1 = -1, y1 = -1; for (let i = 0; i < key.length; i++) if (key[i]) { const x = i % W, y = (i - x) / W; if (x < x0) x0 = x; if (x > x1) x1 = x; if (y < y0) y0 = y; if (y > y1) y1 = y; } return { x: x0 / K, y: y0 / K, w: (x1 - x0 + 1) / K, h: (y1 - y0 + 1) / K, px: { x0, y0, x1: x1 + 1, y1: y1 + 1 } }; }

/* ───────── text: rows → fragments → runs ───────── */
const styleKey = st => [st.name, st.b ? 1 : 0, st.i ? 1 : 0, Math.round(st.size * 100), st.color || '', st.ghost ? 1 : 0, st.pos || 0, st.u || '', st.strike ? 1 : 0, st.shd || ''].join('|');
const textOf = f => f.runs.map(r => r.t).join('');
function styleOfPiece(q, ghost) { return { name: q.font, meas: q.fi.meas, exact: q.fi.exact, bundle: q.fi.bundle, embed: q.fi.embed, b: !!q.bold, i: !!q.italic, size: q.size, color: ghost ? null : q.color || '#000000', ghost: !!ghost, u: q.underline ? 'single' : '', strike: !!q.strike }; }
function rowsOf(pieces, ghost) {
  const ps = [], seen = new Map();
  for (const q of [...pieces].sort((a, b) => a.base - b.base || a.x - b.x)) { // text drawn twice (fake bold, shadows) is kept once
    const k = q.str + '|' + Math.round(q.base), c = seen.get(k), dup = c && c.find(o => Math.abs(o.x - q.x) < 0.15 * q.size && Math.abs(o.base - q.base) < 0.1 * q.size && Math.abs(o.size - q.size) < 0.1);
    if (dup) { if (Math.abs(dup.x - q.x) > 0.01 * q.size) dup.bold = true; continue; }
    if (c) c.push(q); else seen.set(k, [q]); ps.push(q);
  }
  for (const q of ps) { q.st = styleOfPiece(q, ghost); q.key = styleKey(q.st); }
  // a few smaller letters touching bigger text, raised or lowered from its baseline (x², wᵢ, a footnote mark) → part of that text's line, raised or lowered
  const scr = new Map();
  for (const q of ps) {
    const n = q.str.trim().length; if (q.rot || !n || n > 6) continue;
    let host = null, bd = 1e9;
    for (const h of ps) {
      if (h === q || h.rot || h.size < 1.1 * q.size || h.size > 2.2 * q.size || scr.has(h)) continue;
      const dy = h.base - q.base; if (Math.abs(dy) < 0.12 * h.size || Math.abs(dy) > 0.6 * h.size) continue;
      const gx = Math.max(q.x - (h.x + h.w), h.x - (q.x + q.w)); if (gx > 1.5 || gx < -0.5 * q.w) continue;
      const d = Math.abs(dy) + Math.max(0, gx); if (d < bd) { bd = d; host = h; }
    }
    if (host) scr.set(q, host);
  }
  let rows = [];
  for (const q of ps) { if (scr.has(q)) continue; let r = null; for (let i = rows.length - 1; i >= 0 && i >= rows.length - 4; i--) if (Math.abs(rows[i].base - q.base) < 0.35 * Math.min(rows[i].size, q.size)) { r = rows[i]; break; } if (r) { r.ps.push(q); r.size = Math.max(r.size, q.size); } else rows.push({ base: q.base, size: q.size, ps: [q] }); }
  for (const [q, h] of scr) { const r = rows.find(r => r.ps.includes(h)); if (r) { r.ps.push(q); q.scr = true; } else rows.push({ base: q.base, size: q.size, ps: [q] }); }
  // superscripts and subscripts: a short row of smaller or raised text touching a row's text → pieces of that row, raised or lowered
  const drop = new Set();
  for (const r of rows) {
    const a0 = Math.min(...r.ps.map(q => q.x)), a1 = Math.max(...r.ps.map(q => q.x + q.w)); if (r.ps.reduce((n, q) => n + q.str.trim().length, 0) > 12) continue;
    let host = null, best = 1e9;
    for (const h of rows) { // the best host: text it touches, raised above its baseline, the nearest
      if (h === r || drop.has(h) || h.size < 0.95 * r.size) continue;
      const dy = Math.abs(h.base - r.base); if (dy > 0.75 * h.size || (h.size < 1.25 * r.size && dy < 0.2 * h.size)) continue;
      const sp = r.size < 0.85 * h.size ? 0.6 * h.size : 1.5; // (a footnote mark may stand a space before its text)
      const touch = h.ps.some(q => (a0 - (q.x + q.w) > -1.5 && a0 - (q.x + q.w) < sp) || (q.x - a1 > -1.5 && q.x - a1 < sp)), inside = h.ps.some(q => q.x < a0 && a0 < q.x + q.w);
      if (!touch && !inside) continue;
      const sc = (touch ? 0 : 10) + (r.base < h.base ? 0 : 2) + dy / h.size; if (sc < best) { best = sc; host = h; }
    }
    if (!host) continue;
    const pos = Math.round((host.base - r.base) * 2);
    for (const q of r.ps) { q.st = { ...q.st, pos }; q.key = styleKey(q.st); host.ps.push(q); } drop.add(r);
  }
  rows = rows.filter(r => !drop.has(r));
  for (const r of rows) {
    r.base = r.ps.reduce((a, b) => (b.size > a.size ? b : a)).base; r.ps.sort((a, b) => a.x - b.x); r.frags = []; let f = null;
    for (const q of r.ps) if (q.scr) { q.st = { ...q.st, pos: Math.round((r.base - q.base) * 2) }; q.key = styleKey(q.st); }
    const one = q => q.str.trim().length === 1, gapOf = (a, b) => b.x - (a.x + a.w);
    for (let i = 0; i < r.ps.length;) { // tracked letters drawn one by one: their usual gap is letter spacing, not a word gap
      let j = i; while (j + 1 < r.ps.length && one(r.ps[j]) && one(r.ps[j + 1]) && gapOf(r.ps[j], r.ps[j + 1]) < 0.9 * r.ps[j].size) j++;
      if (j - i >= 3) { const gs = []; for (let k = i + 1; k <= j; k++) gs.push(gapOf(r.ps[k - 1], r.ps[k])); const m = median(gs); for (let k = i + 1; k <= j; k++) r.ps[k].trk = m; }
      i = j + 1;
    }
    for (const q of r.ps) {
      let t = q.str, x0 = q.x, x1 = q.x + q.w; const nat = measure(t, q.st), k = nat > 0.1 ? q.w / nat : 1;
      const lead = t.length - t.trimStart().length; if (lead) { x0 += measure(t.slice(0, lead), q.st) * k; t = t.slice(lead); }
      const trail = t.length - t.trimEnd().length; if (trail) { x1 -= measure(t.slice(t.length - trail), q.st) * k; t = t.slice(0, t.length - trail); }
      t = t.replace(/[\t\r\n]/g, ' ');
      const gap = f ? x0 - f.x1 : 0;
      if (f && gap < 0.9 * Math.max(q.size, f.size) && gap > -0.5 * q.size) {
        let last = f.runs[f.runs.length - 1];
        if (gap > Math.max(0.15 * q.size, q.trk != null ? q.trk + 0.15 * q.size : 0) && !/\s$/.test(last.t)) { // (a gap with no space glyph is still a word gap)
          if (q.trk == null && gap > 1.6 * measure(' ', last.st) + 0.3) { last = { t: ' ', st: last.st, key: last.key, x0: f.x1, x1: x0, gap: true }; f.runs.push(last); } // (a wide word gap keeps its width: a space of its own)
          else last.t += ' ';
        }
        if (last.key === q.key && !last.gap) { last.t += t; last.x1 = x1; } else f.runs.push({ t, st: q.st, key: q.key, x0, x1 });
        if (q.size > f.size) f.own = q.base; f.x1 = Math.max(f.x1, x1); f.size = Math.max(f.size, q.size);
      } else { f = { x0, x1, base: r.base, own: q.base, size: q.size, runs: [{ t, st: q.st, key: q.key, x0, x1 }] }; r.frags.push(f); }
    }
  }
  return rows.filter(r => r.frags.length).map(r => ({ base: r.base, size: r.size, frags: r.frags }));
}
function markLists(rows, bodyPt, boxed) { // a lone bullet or "3." standing before a line's text → a real list item (a tick inside a drawn box is a ticked box: it stays text)
  const CHK = /^[✓✔✗✘]$/;
  for (const r of rows) {
    for (let i = r.frags.length - 1; i >= 0; i--) { // ("3." and its text a wide gap apart in one fragment)
      const f = r.frags[i], t = f.runs[0].t.trim(); if (f.runs.length < 3 || !f.runs[1].gap || !(BUL.test(t) || NUMM.test(t))) continue;
      r.frags.splice(i, 1, { ...f, x1: f.runs[0].x1, runs: [f.runs[0]] }, { ...f, x0: f.runs[2].x0, runs: f.runs.slice(2) });
    }
    for (let i = r.frags.length - 2; i >= 0; i--) {
      const f0 = r.frags[i], f1 = r.frags[i + 1], t = textOf(f0).trim(); if (f1.x0 - f0.x1 > 40 || f0.runs.some(q => q.st.pos) || f1.marker) continue;
      const kind = BUL.test(t) ? 'bullet' : NUMM.test(t) && r.size < 1.12 * bodyPt ? 'number' : null; if (!kind || (CHK.test(t) && boxed && boxed(f0))) continue;
      f1.marker = { t, kind, x: f0.x0, st: f0.runs[0].st }; r.frags.splice(i, 1);
    }
  }
}

/* ───────── layout analysis: bands of single-column text, tables and side-by-side columns ───────── */
function isTableRows(rows, PW) { const cells = rows.flatMap(r => r.frags); if (!cells.length) return true; const wid = median(cells.map(f => f.x1 - f.x0)), ch = median(cells.map(f => textOf(f).length)); return !(wid > 0.26 * PW && ch > 28); }
function boxify(fr, cands) { // fills that wholly hold some text become shaded boxes (a one-cell table that grows with its text); their text goes inside
  const tTop = f => f.base - 0.8 * f.size, tBot = f => f.base + 0.25 * f.size;
  const inside = (f, F) => f.x0 >= F.x - 0.5 && f.x1 <= F.x + F.w + 0.5 && (f.box ? f.box.y >= F.y - 0.5 && f.box.y + f.box.h <= F.y + F.h + 0.5 : tTop(f) >= F.y - 0.5 && tBot(f) <= F.y + F.h + 0.5);
  const touches = (f, F) => f.x0 < F.x + F.w && f.x1 > F.x && (f.box ? f.box.y < F.y + F.h && f.box.y + f.box.h > F.y : tTop(f) < F.y + F.h && tBot(f) > F.y);
  let out = fr;
  for (const F of cands) {
    if (F.used || F.taken) continue;
    const ins = out.filter(f => inside(f, F)); if (!ins.length) continue;
    if (out.some(f => !ins.includes(f) && touches(f, F))) continue; // (text crossing its edge: not a box)
    const rs = new Set(ins.map(f => f.row)), rest = out.filter(f => !ins.includes(f) && !f.box);
    if (rs.size <= 2 && rest.some(f => rs.has(f.row))) continue; // (a cell of a grid: its line goes on outside)
    if ([...rs].some(r => ins.filter(f => f.row === r).length >= 2)) { // a shaded band of a table: the columns go on above or below it
      const a = ins.filter(f => !f.box), near = rest.filter(f => f.base > F.y - 4 * f.size && f.base < F.y + F.h + 4 * f.size);
      const xs = new Set(); for (const f of near) for (const g of a) { if (Math.abs(g.x0 - f.x0) < 2) xs.add(Math.round(g.x0)); if (Math.abs(g.x1 - f.x1) < 2) xs.add(Math.round(g.x1)); }
      if (xs.size >= 2) continue;
    }
    const size = F.h / 1.25, row = { base: F.y + 0.95 * size, size }, pf = { box: F, inner: ins, x0: F.x, x1: F.x + F.w, base: row.base, size, row, runs: [] };
    F.taken = true; out = out.filter(f => !ins.includes(f)).concat(pf);
  }
  return out;
}
function layBox(f, flow, o, blocks, maxPitch, G) {
  const F = f.box, top = tableStart(flow, F.y, blocks), cb = [], fl = new Flow(top);
  const inner = boxify(f.inner, G.boxes);
  layBands(segment(inner, G.PW), fl, { refL: F.x, areaR: F.x + F.w }, cb, maxPitch, G);
  if (!cb.length || cb[cb.length - 1].t !== 'p') cb.push(spacer(0.05));
  const Ht = Math.max(20, twip(F.y + F.h - top), Math.ceil((fl.bottom() - top) * 20));
  // drawing inside the box (badges, bullets, rules) would vanish under cell shading: then the fill stays a shape behind, in the page's drawing
  const holds = g => g !== F && g.x >= F.x - 0.5 && g.y >= F.y - 0.5 && g.x + g.w <= F.x + F.w + 0.5 && g.y + g.h <= F.y + F.h + 0.5;
  const part = g => (Math.max(0, Math.min(g.x + g.w, F.x + F.w) - Math.max(g.x, F.x)) * Math.max(0, Math.min(g.y + g.h, F.y + F.h) - Math.max(g.y, F.y))) / (g.w * g.h || 1); // (rounded corners are pictures half over the fill)
  const mine = g => g === F.src || g === F.pic;
  const deco = !!F.pic || G.rects.some(g => !g.used && !mine(g) && holds(g)) || G.pics.some(g => !mine(g) && (holds(g) || part(g) > 0.3));
  const T = { t: 'tbl', ind: F.x - o.refL, widths: [F.w], rows: [{ h: Ht, cells: [{ blocks: cb, shd: deco ? null : F.fill }], top, bot: top + Ht / 20, need: fl.bottom() - top }], box: true, cellL: F.x };
  blocks.push(T);
  if (!deco) F.used = true; // (the cell's shading is the fill; a band cut at the top keeps its shape above, under the shading)
  else if (F.src || F.pic) (F.src || F.pic).boxOf = T; // (the fill or picture hangs from the box, so it moves with it)
  flow.base = top + Ht / 20; flow.L = 0; flow.lastTable = true;
}
function rowsOfFrags(fr) { // fragments → their rows (only the fragments given), top to bottom
  const m = new Map(); for (const f of fr) { if (!m.has(f.row)) m.set(f.row, []); m.get(f.row).push(f); }
  return [...m].map(([r, fs]) => { const big = fs.reduce((a, b) => (b.size > a.size ? b : a)); return { base: big.own != null ? big.own : r.base, size: big.size, frags: fs.sort((a, b) => a.x0 - b.x0) }; }).sort((a, b) => a.base - b.base); // (each part keeps its own baseline: two columns' lines may be a little apart)
}
function segment(fr, PW) { // XY-cut: side-by-side parts (a table when they share their lines, else text columns), else split at the widest gap between rows
  const rows = rowsOfFrags(fr); if (rows.length <= 1) return [{ kind: 'flow', rows }];
  const s = [...fr].sort((a, b) => a.x0 - b.x0), groups = [[s[0]]]; let end = s[0].x1;
  for (const f of s.slice(1)) { if (f.x0 - end >= 9) groups.push([]); groups[groups.length - 1].push(f); end = Math.max(end, f.x1); }
  const join = parts => { const out = []; for (const band of parts.flat()) { const l = out[out.length - 1]; if (l && l.kind === 'flow' && band.kind === 'flow') l.rows.push(...band.rows); else out.push(band); } return out; };
  const vertical = () => {
    const sh = rows.map(r => groups.filter(g => g.some(f => f.row === r.frags[0].row)).length >= 2), n = sh.filter(Boolean).length;
    if (n >= 2 && n >= 0.5 * rows.length && isTableRows(rows.filter(r => r.frags.length > 1), PW) && !rows.slice(sh.indexOf(true), sh.lastIndexOf(true) + 1).some(r => r.frags.some(f => f.box))) { // a table; lone rows above or below it are not part of it
      let i0 = sh.indexOf(true), i1 = sh.lastIndexOf(true); const close = (a, b) => Math.abs(b.base - a.base) < 0.75 * Math.min(a.size, b.size); // (a cell's line set between two lines of its neighbour belongs to the table)
      while (i1 + 1 < rows.length && close(rows[i1], rows[i1 + 1])) i1++;
      while (i0 > 0 && close(rows[i0 - 1], rows[i0])) i0--;
      const part =(a, b) => { const fs = rows.slice(a, b).flatMap(r => r.frags); return fs.length ? segment(fs, PW) : []; };
      return join([part(0, i0), [{ kind: 'table', rows: rows.slice(i0, i1 + 1) }], part(i1 + 1, rows.length)]);
    }
    return [{ kind: 'cols', cols: groups.map(g => ({ x0: Math.min(...g.map(f => f.x0)), x1: Math.max(...g.map(f => f.x1)), bands: segment(g, PW) })) }];
  };
  const top = r => r.base - 0.95 * r.size, bot = r => r.base + 0.3 * r.size, byTop = [...rows].sort((a, b) => top(a) - top(b)); let best = -1e9, at = -1, b = bot(byTop[0]);
  for (let i = 1; i < byTop.length; i++) { const g = top(byTop[i]) - b; if (g > best) { best = g; at = i; } b = Math.max(b, bot(byTop[i])); }
  const across = () => { const cut = new Set(byTop.slice(0, at).flatMap(r => r.frags)); return join([segment(fr.filter(f => cut.has(f)), PW), segment(fr.filter(f => !cut.has(f)), PW)]); };
  if (at > 0 && best >= 6) return across(); // (a clear gap across the whole width comes first)
  if (groups.length > 1 && groups.every(g => new Set(g.map(f => f.row)).size >= 2)) return vertical();
  if (at > 0 && best > 0) return across();
  if (groups.length > 1) return vertical();
  return [{ kind: 'flow', rows }];
}
const linesOf = rows => rows.map(r => ({ base: r.base, size: r.size, frags: r.frags, x0: r.frags[0].x0, x1: r.frags[r.frags.length - 1].x1 }));
const MONO = /mono|courier|consol|cousine|menlo|lucida console/i, LEAD = /(\s*\.){3,}\s*\S{0,6}$/, REC = /^\S+(\s\S+)?\s*(->|=>|:=|=|:)\s/;
function paragraphsOf(lines, maxPitch, colR) {
  colR = colR || Math.max(...lines.map(l => l.x1)); const colL = Math.min(...lines.map(l => l.x0)), out = []; let cur = null;
  for (const l of lines) {
    const multi = l.frags.length > 1, mk = l.frags[0].marker;
    if (cur && !multi && !cur.multi && !mk) {
      const prev = cur.lines[cur.lines.length - 1], gap = l.base - prev.base, s0 = prev.frags[0].runs[0].st, s1 = l.frags[0].runs[0].st;
      const same = Math.abs(l.size - prev.size) < 0.6 && s0.name === s1.name;
      const pitchOk = gap > 0.9 * l.size && gap < maxPitch * l.size && (cur.pitch == null || Math.abs(gap - cur.pitch) < Math.max(0.6, 0.08 * l.size));
      const xOk = cur.marker ? Math.abs(l.x0 - cur.lines[0].x0) < 2 : cur.lines.length === 1 || Math.abs(l.x0 - prev.x0) < 2;
      const keep = MONO.test(s0.name) || MONO.test(s1.name) || (REC.test(textOf(prev.frags[0])) && REC.test(textOf(l.frags[0]))) || LEAD.test(textOf(prev.frags[0])) || LEAD.test(textOf(l.frags[0])) || !/\s|[-‐]$/.test(textOf(prev.frags[0]).trim()); // (code, "key -> value" records, contents lines and one-word lines such as a column of numbers keep their lines)
      if (same && pitchOk && xOk && !keep && prev.x1 > colR - Math.max(3 * l.size, 0.18 * (colR - colL))) { cur.lines.push(l); cur.gaps.push(gap); cur.pitch = cur.gaps.reduce((a, b) => a + b, 0) / cur.gaps.length; continue; }
    }
    cur = { lines: [l], pitch: null, multi, gaps: [], marker: mk || null }; out.push(cur);
  }
  return out;
}
function finishParas(paras, colL, colR) {
  for (const P of paras) {
    const L = P.lines; P.x0 = Math.min(...L.map(l => l.x0)); P.x1 = Math.max(...L.map(l => l.x1));
    if (L.length >= 2) {
      const body = L.slice(0, -1);
      const stretched = l => { const w = l.frags[0].runs.reduce((a, r) => a + measure(r.t, r.st), 0); return l.x1 - l.x0 > w * 1.008 + 0.5; }; // (a justified line is wider than its letters)
      const atEnd = l => Math.abs(l.x1 - colR) < 2, nearEnd = l => colR - l.x1 < Math.max(2, 0.04 * (colR - colL));
      if (Math.abs(L[1].x0 - P.x0) < 2 && body.every(nearEnd) && (L.length > 2 ? body.filter(atEnd).length >= 0.7 * body.length : body.every(atEnd) && body.some(stretched))) P.align = 'both';
      else if (L.every(l => Math.abs((l.x0 + l.x1) / 2 - (L[0].x0 + L[0].x1) / 2) < 2) && L[0].x0 > colL + 6) P.align = 'center';
      else if (L.every(l => Math.abs(l.x1 - L[0].x1) < 1.5) && L.some(l => Math.abs(l.x0 - L[0].x0) > 3)) P.align = 'right';
      else P.align = 'left';
    } else {
      const l = L[0], c = (l.x0 + l.x1) / 2;
      P.align = Math.abs(c - (colL + colR) / 2) < 2 && l.x0 > colL + 12 ? 'center' : Math.abs(l.x1 - colR) < 1.5 && l.x0 > colL + 0.4 * (colR - colL) ? 'right' : 'left';
    }
    if (P.multi || P.marker) P.align = P.align === 'both' && P.marker ? 'both' : 'left';
  }
  return paras;
}

/* ───────── vertical bookkeeping: where Word puts the last baseline (rounded twips carried forward) ───────── */
class Flow {
  constructor(top) { this.top = top; this.base = null; this.L = 0; this.lastTable = false; }
  bottom() { return this.base == null ? this.top : this.base + 0.2 * this.L; }
  spacing(base, L) { return base - this.bottom() - 0.8 * L; }
}

/* ───────── widths: each run is fitted to the room it has in the PDF ───────── */
function fitRun(t, st, target) { // → { sp: twentieths of a point after each letter, w: % width (stand-in fonts only), d: the unrounded sp }
  const n = [...t].length; if (!n || target == null) return { sp: 0, w: 100, d: 0 };
  let nat = measure(t, st), w = 100; const gap = /^\s+$/.test(t); // (a space of its own: a wide word gap)
  if (!st.exact && nat > 1 && !gap) { const r = target / nat; if (Math.abs(r - 1) > 0.08) { w = clamp(Math.round(r * 100), 50, 200); nat *= w / 100; } }
  const d = clamp((target - nat) / n * 20, -40, gap ? Math.max(80, 30 * st.size) : 80); return { sp: Math.round(d), w, d };
}
const runW = r => measure(r.t, r.st) * (r.w || 100) / 100 + [...r.t].length * (r.sp || 0) / 20;
function fitLine(runs, loose) { // targets: up to where the next run starts, the last run its own extent. loose (a justified line): only the half-point size is made up for, Word spreads the rest over the spaces
  const out = [];
  runs.forEach((r, j) => {
    const nx = runs[j + 1]; let target = nx ? nx.x0 - r.x0 : r.x1 - r.x0;
    if (!nx) { const tail = r.t.length - r.t.trimEnd().length; if (tail) target += measure(r.t.slice(r.t.length - tail), r.st); }
    if (loose && r.st.exact) target = Math.min(target, measure(r.t, r.st) * r.st.size / wsz(r.st.size));
    const f = fitRun(r.t, r.st, target), ch = [...r.t], n = ch.length;
    if (Math.abs(f.d - f.sp) * n >= 6 && n >= 4 && !loose) { // (a long run: part of it gets one twip more, so the rounding does not add up)
      const s0 = Math.floor(f.d), m = Math.round((f.d - s0) * n);
      if (m > 0 && m < n) { out.push({ ...r, t: ch.slice(0, n - m).join(''), sp: s0, w: f.w }, { ...r, t: ch.slice(n - m).join(''), sp: s0 + 1, w: f.w }); return; }
    }
    out.push(Object.assign(r, { sp: loose ? Math.min(f.sp, Math.floor(f.d)) : f.sp, w: f.w }));
  });
  return out;
}

/* ───────── one paragraph ───────── */
function leaderCut(rs) { // a line part ending in dot leaders (a table of contents): the dots go, a tab with a dot leader takes their place → where they ended, or null
  let k = rs.length - 1; while (k > 0 && /^[\s.]*$/.test(rs[k].t)) k--; // (runs after k hold dots only)
  if (!/(\s*\.){3,}\s*$/.test(rs.slice(k).map(r => r.t).join(''))) return null;
  const end = rs[rs.length - 1].x1, r = rs[k], old = r.t, keep = old.replace(/(\s*\.){2,}\s*$/, '').replace(/\s+$/, ''); rs.length = k + 1;
  if (keep) { r.x1 = r.x0 + measure(keep, r.st); r.t = keep; } else rs.pop(); // (spaced-out dots tell nothing of the text's own width: it keeps its natural width)
  return rs.length ? end : null;
}
let curWords = new Set(); // the words of the page being laid out (to tell a hyphenated word from a compound)
function layPara(P, flow, o, blocks) { // o: { refL: left edge of the text area, colR: right edge of the column, areaR: right edge of the text area }
  const L = P.lines, sz = Math.max(...L.map(l => l.size));
  let pitch = P.pitch || 1.2 * sz, before = flow.spacing(L[0].base, pitch);
  if (before < 0 && L.length === 1) { pitch = Math.max(0.9 * sz, pitch + before / 0.8); before = flow.spacing(L[0].base, pitch); }
  if (before < -0.05) { // a single-line paragraph just above: a smaller line height there makes room (its baseline stays put); a table just above: its last row's spare room
    const pb = blocks[blocks.length - 1];
    if (pb && pb.t === 'tbl' && !pb.cols && flow.lastTable) {
      const r = pb.rows[pb.rows.length - 1], spare = r.h / 20 - r.need, dh = Math.min(-before, spare - 0.05);
      if (dh > 0.05) { const nh = Math.max(20, twip(r.h / 20 - dh)), d = (r.h - nh) / 20; r.h = nh; r.bot -= d; flow.base -= d; before = flow.spacing(L[0].base, pitch); }
    }
    if (pb && pb.t === 'p' && pb.nLines === 1 && flow.base != null && Math.abs(flow.base - pb.base) < 0.01) {
      const Lp = pb.line / 20, dL = Math.min(-5 * before, Lp - 0.9 * pb.size);
      if (dL > 0.05) { const nl = Math.max(1, twip(Lp - dL)), nb = Math.max(0, twip(pb.base - pb.top0 - 0.8 * nl / 20)); pb.line = nl; pb.before = nb; flow.base = pb.base = pb.top0 + nb / 20 + 0.8 * nl / 20; flow.L = nl / 20; before = flow.spacing(L[0].base, pitch); }
    }
  }
  if (before < -0.05 && L.length > 1 && !P.multi) { // still no room above: the first line becomes a paragraph of its own (its line height can shrink), the rest follows
    const one = { ...P, lines: [L[0]], pitch: null, align: P.align === 'both' ? 'left' : P.align, x0: L[0].x0, x1: L[0].x1 };
    const rest = { ...P, lines: L.slice(1), gaps: P.gaps ? P.gaps.slice(1) : [], marker: null, x0: Math.min(...L.slice(1).map(l => l.x0)), x1: Math.max(...L.slice(1).map(l => l.x1)) };
    if (rest.lines.length < 2 && rest.align === 'both') rest.align = 'left';
    layPara(one, flow, o, blocks); return layPara(rest, flow, o, blocks);
  }
  if (before > Math.max(36, 3 * pitch)) { // a wide gap above: an empty paragraph of its own takes it (Enter in the text then copies no gap; a picture in the gap hangs from it)
    const s = spacer(before), b0 = flow.bottom(); s.gapSp = true; s.nLines = 1; s.top0 = b0; blocks.push(s);
    flow.base = b0 + 0.8 * s.line / 20; flow.L = s.line / 20; flow.lastTable = false; before = flow.spacing(L[0].base, pitch);
  }
  if (before < -1.5) P.neg = true;
  const first = L[0].x0 - o.refL, rest = (L.length > 1 ? L[1].x0 : L[0].x0) - o.refL;
  const B = { t: 'p', ind: { left: rest, first: first - rest, right: 0 }, jc: P.align, tabs: [], runs: [], num: null, size: sz };
  let right = o.areaR;
  if (P.marker) { const mk = P.marker; B.ind = { left: first, first: -(L[0].x0 - mk.x), right: 0 }; B.tabs.push(first); B.num = mk; }
  if (P.multi) {
    const l = L[0]; B.ind = { left: l.x0 - o.refL, first: 0, right: 0 }; B.tabs = []; B.jc = 'left';
    B.lead = [];
    l.frags.forEach((f, i) => {
      if (i) {
        const mk = f.marker; // (a bullet before a later part of the line stays as text before it)
        if (mk) { B.tabs.push(mk.x - o.refL); B.runs.push({ tab: true }, { t: mk.t, st: mk.st, key: styleKey(mk.st), x0: mk.x, x1: f.x0, sp: 0, w: 100 }); }
        B.tabs.push(f.x0 - o.refL); B.runs.push({ tab: true });
      }
      const rs = f.runs.map(r => ({ ...r }));
      const end = leaderCut(rs); if (end != null) { B.lead[B.tabs.length] = 'dot'; if (i === l.frags.length - 1) B.tabs.push(end - o.refL); }
      B.runs.push(...fitLine(rs)); if (end != null && i === l.frags.length - 1) B.runs.push({ tab: true });
    });
    right = Math.max(o.areaR, l.x1 + 20);
  } else {
    const raw = L.map(l => l.frags[0].runs.map(r => ({ ...r }))), joins = [], lead = L.length === 1 ? leaderCut(raw[0]) : null;
    for (let k = 0; k < raw.length - 1; k++) { // how line k meets line k+1: a space, nothing (after a dash) or a soft hyphen (a word split over the lines)
      const rs = raw[k], last = rs[rs.length - 1], nx = raw[k + 1][0].t, a = /(\p{L}+)[-‐]$/u.exec(last.t), b = /^(\p{Ll}+)/u.exec(nx);
      if (a && b && !/[-‐]\p{L}+[-‐]$/u.test(last.t) && curWords.has((a[1] + b[1]).toLowerCase())) joins.push('shy'); // (only when the page also has the word unbroken)
      else if (/\s$/.test(last.t)) joins.push(' ');
      else if (/[-‐–—]$/.test(last.t)) joins.push(''); // (not after '/': a line ending in a path is not glued to the next)
      else { last.t += ' '; joins.push(' '); }
    }
    const lines = raw.map((rs, k) => fitLine(rs, P.align === 'both' && k < raw.length - 1 && o.colR - L[k].x1 < 2));
    const brk = new Array(Math.max(0, lines.length - 1)).fill(false);
    if (lines.length > 1) {
      const ind = k => (k ? rest : first), nat = [], nxt = [], nsp = [];
      lines.forEach((rs, k) => {
        let w = 0; rs.forEach((r, j) => { w += j === rs.length - 1 ? runW({ ...r, t: r.t.replace(/\s+$/, '') }) : runW(r); });
        nat.push(w); nsp.push((rs.map(r => r.t).join('').replace(/\s+$/, '').match(/ /g) || []).length);
        if (k < lines.length - 1) {
          const last = rs[rs.length - 1], nr = lines[k + 1][0], whole = (nr.t.match(/^\S+/) || [''])[0], hy = /^[^\s\-‐]*?\p{L}[-‐](?=\p{L})/u.exec(whole);
          const word = hy ? hy[0] : whole, ww = runW({ ...nr, t: word }); // (Word may break after a hyphen: "date-" alone may come up onto the line)
          nxt.push(w + (joins[k] === ' ' ? measure(' ', last.st) + (last.sp || 0) / 20 : joins[k] === 'shy' ? -measure('-', last.st) : 0) + ww);
          if (joins[k] === ' ') nsp[k]++;
        }
      });
      if (P.align === 'both') { // Word fits one more word if squeezing the spaces makes room (each space by ~20%)
        const avail = o.colR - o.refL;
        for (let k = 0; k < nxt.length; k++) { const spW = measure(' ', lines[k][lines[k].length - 1].st); brk[k] = ind(k) + nxt[k] - SQUEEZE * spW * nsp[k] <= avail + 0.15; }
        right = o.colR;
      } else {
        const lo = Math.max(...nat.map((w, k) => (P.align === 'left' ? ind(k) : 0) + w)), his = nxt.map((w, k) => (P.align === 'left' ? ind(k) : 0) + w), hi = Math.min(...his);
        const target = hi - lo > 0.4 ? (lo + hi) / 2 : lo + 0.3;
        for (let k = 0; k < nxt.length; k++) brk[k] = his[k] <= target;
        if (P.align === 'left') right = o.refL + target;
        else if (P.align === 'center') { const c = (L[0].x0 + L[0].x1) / 2; B.ind = { left: c - target / 2 - o.refL, first: 0 }; right = c + target / 2; }
        else { B.ind = { left: P.x1 - target - o.refL, first: 0 }; right = P.x1; }
      }
    } else if (P.align === 'center') { const l = L[0], c = (l.x0 + l.x1) / 2, half = (l.x1 - l.x0) / 2 + 2 * sz; B.ind = { left: c - half - o.refL, first: 0 }; right = c + half; }
    else if (P.align === 'right') { B.ind = { left: 0, first: 0 }; right = P.x1; }
    else if (P.align === 'both') right = o.colR;
    lines.forEach((rs, k) => {
      if (k < brk.length && brk[k]) { const r = rs[rs.length - 1]; r.t = r.t.replace(/\s+$/, ''); r.br = true; }
      if (joins[k] === 'shy') rs[rs.length - 1].shy = true;
      B.runs.push(...rs);
    });
    if (lead != null) { B.runs.push({ tab: true }); B.tabs.push(lead - o.refL); B.lead = []; B.lead[B.tabs.length - 1] = 'dot'; right = Math.max(right, lead + 1); }
  }
  B.ind.right = o.areaR - right;
  const Lt = Math.max(1, twip(pitch)), Bt = Math.max(0, twip(before));
  B.before = Bt; B.line = Lt; B.top0 = flow.bottom(); B.nLines = L.length;
  const firstBase = B.top0 + Bt / 20 + 0.8 * Lt / 20;
  flow.base = B.base = firstBase + (L.length - 1) * Lt / 20; flow.L = Lt / 20; flow.lastTable = false;
  B.runs = mergeRuns(B.runs.filter(r => r.tab || r.t || r.br || r.shy));
  blocks.push(B); return B;
}
function mergeRuns(runs) { const out = []; for (const r of runs) { const l = out[out.length - 1]; if (l && !l.tab && !r.tab && !l.br && !l.shy && !l.fld && !r.fld && l.key === r.key && l.sp === r.sp && l.w === r.w) { l.t += r.t; l.br = r.br; l.shy = r.shy; } else out.push({ ...r }); } return out; }
const spacer = gap => ({ t: 'p', before: 0, line: Math.max(1, twip(gap)), ind: null, runs: [], tiny: true });

/* ───────── tables and side-by-side columns ───────── */
function layTable(band, flow, o, blocks, maxPitch, G) {
  const rows = band.rows, cols = [];
  for (const f of rows.flatMap(r => r.frags).sort((a, b) => a.x0 - b.x0)) { // cells whose extents overlap belong to one column
    const c = cols.find(c => f.x0 < c.x1 + 2 && f.x1 > c.x0 - 2);
    if (c) { c.fr.push(f); c.x0 = Math.min(c.x0, f.x0); c.x1 = Math.max(c.x1, f.x1); } else cols.push({ x0: f.x0, x1: f.x1, fr: [f] });
  }
  cols.sort((a, b) => a.x0 - b.x0);
  for (const c of cols) { const l0 = c.fr.map(f => f.x0), r1 = c.fr.map(f => f.x1); c.al = Math.max(...r1) - Math.min(...r1) < 1.5 && Math.max(...l0) - Math.min(...l0) > 2 ? 'right' : 'left'; }
  const edges = [cols[0].x0 - 4, ...cols.slice(1).map((c, k) => (cols[k].x1 + c.x0) / 2), cols[cols.length - 1].x1 + 6];
  // table rows: a cell's text that wraps onto more lines (or is centred between two lines of a neighbour) stays in one row
  const colOf = f => cols.findIndex(c => c.fr.includes(f)), rules = G ? G.rects.filter(g => g.h <= 3 && g.w > 8) : [];
  const LR = [];
  for (const r of rows) {
    const cur = LR[LR.length - 1], p = cur && cur.rows[cur.rows.length - 1];
    if (p) {
      const s = Math.min(r.size, p.size), d = r.base - p.base, ruled = rules.some(g => g.y > p.base && g.y + g.h < r.base - 0.6 * r.size && r.frags.some(f => g.x < f.x1 && g.x + g.w > f.x0));
      const above = c => { for (let i = cur.rows.length - 1; i >= 0; i--) { const f = cur.rows[i].frags.find(f => colOf(f) === c); if (f) return f; } return null; };
      const wraps = () => new Set(r.frags.map(colOf)).size < cols.length && r.frags.every(f => {
        const c = colOf(f), a = above(c); if (!a) return false; const C = cols[c];
        return (Math.abs(a.x0 - f.x0) < 2 || Math.abs(a.x1 - f.x1) < 2 || Math.abs(a.x0 + a.x1 - f.x0 - f.x1) < 4) && a.x1 > C.x1 - Math.max(3 * s, 0.25 * (C.x1 - C.x0)) && !/[.:;!?]$/.test(textOf(a).trim()) || false;
      });
      if (!ruled && (d < 0.75 * s || (d <= 1.35 * s && wraps()))) { cur.rows.push(r); continue; }
    }
    LR.push({ rows: [r] });
  }
  for (const g of LR) { g.top = Math.min(...g.rows.map(r => r.base - 0.95 * r.size)); g.bot = Math.max(...g.rows.map(r => r.base + 0.3 * r.size)); }
  const tops = [LR[0].top - 3, ...LR.slice(1).map((g, k) => (LR[k].bot + g.top) / 2)];
  const bots = [...tops.slice(1), LR[LR.length - 1].bot + 3];
  const bd = snapTable(edges, tops, bots, cols, LR, G);
  const top0 = tableStart(flow, tops[0], blocks), T = { t: 'tbl', ind: edges[0] - o.refL, widths: edges.slice(1).map((e, k) => e - edges[k]), rows: [] };
  let y = top0;
  LR.forEach((g, k) => {
    const bt = k === 0 ? bd.top : null, bb = bd.bot[k], wt = bt ? bt.w : 0, wb = bb ? bb.w : 0; // (Word adds a cell border's width to the row; a top border pushes the text down)
    const Ht = Math.max(20, twip(bots[k] - y - wt)), cells = [];
    cols.forEach((c, ci) => {
      const cb = [], lines = [];
      for (const r of g.rows) {
        const cf = r.frags.filter(f => c.fr.includes(f)); if (!cf.length) continue;
        const big = cf.reduce((a, b) => (b.size > a.size ? b : a)); lines.push({ base: big.own != null ? big.own : r.base, size: big.size, frags: cf, x0: cf[0].x0, x1: cf[cf.length - 1].x1 }); // (a cell keeps its own baseline)
      }
      const o2 = { refL: edges[ci], colR: edges[ci + 1], areaR: edges[ci + 1] }, fl = new Flow(y + wt);
      if (lines.length === 1 || lines.some(l => l.frags.length > 1)) for (const l of lines) {
        const cf = l.frags, P = { lines: [l], pitch: null, multi: cf.length > 1, x0: l.x0, x1: l.x1, marker: cf[0].marker || null, align: c.al === 'right' && cf.length === 1 ? 'right' : 'left' };
        layPara(P, fl, o2, cb);
      } else if (lines.length) {
        const cR = Math.max(...lines.map(l => l.x1)), cL = Math.min(...lines.map(l => l.x0));
        for (const P of finishParas(paragraphsOf(lines, maxPitch || 1.6, cR), cL, cR)) layPara(P, fl, P.align === 'both' ? { ...o2, colR: cR } : o2, cb);
      }
      const cell = { blocks: cb };
      if (bt && ci >= bt.c0 && ci < bt.c1) (cell.bd = cell.bd || {}).top = bt;
      if (bb && ci >= bb.c0 && ci < bb.c1) (cell.bd = cell.bd || {}).bottom = bb;
      cells.push(cell);
    });
    const need = Math.max(0, ...cells.map(c => c.blocks.length ? c.blocks[c.blocks.length - 1].base + 0.2 * c.blocks[c.blocks.length - 1].line / 20 - y - wt : 0));
    const Hf = Math.max(Ht, Math.ceil(need * 20 + 1)); // (Word grows a row its text does not fit in; the rows below follow from where it really ends)
    T.rows.push({ h: Hf, cells, top: y, bot: y + wt + Hf / 20 + wb, need, wb }); y += wt + Hf / 20 + wb;
  });
  if (G) for (const g of G.rects) { // vertical rules on column edges → left / right cell borders (they then move with the table)
    if (g.used || g.w > 3 || g.h < 6) continue; const xc = g.x + g.w / 2, n = cols.length, k = edges.findIndex(e => Math.abs(e - xc) < 2.5); if (k < 0) continue;
    const rs = T.rows.filter(r => Math.min(r.bot, g.y + g.h) - Math.max(r.top, g.y) > 0.5 * (r.bot - r.top)); if (!rs.length || Math.abs(rs.reduce((t, r) => t + r.bot - r.top, 0) - g.h) > 3) continue;
    const sz = BW.reduce((a, v) => (Math.abs(v - g.w * 8) < Math.abs(a - g.w * 8) ? v : a)), b = { sz, w: sz / 8, color: g.fill };
    for (const r of rs) { const c = r.cells[k < n ? k : n - 1]; (c.bd = c.bd || {})[k < n ? 'left' : 'right'] = b; }
    g.used = true;
  }
  shadeTable(T, edges, G);
  blocks.push(T); flow.base = y; flow.L = 0; flow.lastTable = true;
}
function tableStart(flow, top, blocks) { // a spacer paragraph takes the gap above a table (and keeps two tables apart)
  const prev = flow.bottom(), gap = top - prev;
  if (gap > 0.5 || flow.lastTable) { const s = spacer(Math.max(0.05, gap)); blocks.push(s); return prev + s.line / 20; }
  return prev;
}
const BW = [2, 4, 6, 8, 12, 18, 24, 36, 48]; // border widths Word draws (eighths of a point)
function snapTable(edges, tops, bots, cols, rows, G) { // fills and rules around a table say where its cells really end; rules on row edges become borders
  const bd = { top: null, bot: [] }; if (!G) return bd;
  const n = edges.length - 1, rTop = k => rows[k].top, rBot = k => rows[k].bot, inGap = (v, a, b) => v > a && v < b;
  const near = g => !(g.x + g.w < edges[0] - 14 || g.x > edges[n] + 14 || g.y + g.h < tops[0] - 14 || g.y > bots[bots.length - 1] + 14);
  // a rule drawn cell by cell (one piece per column) is one rule; where its pieces meet are column edges
  const hr = [], segs = G.rects.filter(g => !g.used && near(g) && g.h <= 3 && g.w > 8).sort((a, b) => a.x - b.x);
  for (const g of segs) { const m = hr.find(m => Math.abs(m.y - g.y) < 0.6 && Math.abs(m.h - g.h) < 0.4 && m.fill === g.fill && g.x - (m.x + m.w) < 3 && g.x - (m.x + m.w) > -1); if (m) { m.w = g.x + g.w - m.x; m.parts.push(g); } else hr.push({ x: g.x, y: g.y, w: g.w, h: g.h, fill: g.fill, parts: [g] }); }
  for (const m of hr) if (m.parts.length > 1) for (const g of m.parts) for (const v of [g.x, g.x + g.w]) for (let k = 1; k < n; k++) if (inGap(v, cols[k - 1].x1, cols[k].x0)) edges[k] = v;
  for (const g of G.rects) { // outer edges and column edges from fills and rules
    if (g.used || !near(g)) continue; const gx1 = g.x + g.w;
    if (g.x < edges[0] && g.x > edges[0] - 14) edges[0] = g.x; if (gx1 > edges[n] && gx1 < edges[n] + 14) edges[n] = gx1;
    if (g.h > 3) for (let k = 1; k < n; k++) { if (inGap(g.x, cols[k - 1].x1, cols[k].x0)) edges[k] = g.x; else if (inGap(gx1, cols[k - 1].x1, cols[k].x0)) edges[k] = gx1; }
  }
  for (const g of hr.concat(G.rects.filter(g => g.h > 3 || g.w <= 8))) {
    if (g.used || (!g.parts && !near(g))) continue; const gy1 = g.y + g.h;
    if (g.parts) Object.defineProperty(g, 'used', { get: () => g.parts.every(q => q.used), set: v => g.parts.forEach(q => { q.used = v; }) });
    if (g.h <= 3 && g.w > 8) { // a horizontal rule: a border when it runs between column edges, on a row edge
      const c0 = edges.findIndex(e => Math.abs(e - g.x) < 3), c1 = edges.findIndex(e => Math.abs(e - (g.x + g.w)) < 3); if (c0 < 0 || c1 <= c0) continue;
      const sz = BW.reduce((a, b) => (Math.abs(b - g.h * 8) < Math.abs(a - g.h * 8) ? b : a)), w = sz / 8, b = { sz, w, color: g.fill, c0, c1 };
      if (gy1 <= rTop(0) + 1 && g.y > tops[0] - 12) { if (!bd.top) { tops[0] = g.y; bd.top = b; g.used = true; } continue; }
      const last = rows.length - 1;
      if (g.y >= rBot(last) - 1 && g.y < bots[last] + 12) { if (!bd.bot[last]) { bots[last] = g.y; bd.bot[last] = b; g.used = true; } continue; }
      for (let k = 1; k < rows.length; k++) if (inGap(g.y + g.h / 2, rBot(k - 1) - 1, rTop(k) + 1) && !bd.bot[k - 1]) { bots[k - 1] = g.y; tops[k] = g.y + w; bd.bot[k - 1] = b; g.used = true; break; }
      continue;
    }
    if (g.y < tops[0] + 3 && g.y > tops[0] - 10 && g.y < rTop(0)) tops[0] = g.y;
    if (gy1 > bots[bots.length - 1] - 3 && gy1 < bots[bots.length - 1] + 10 && gy1 > rBot(rows.length - 1)) bots[bots.length - 1] = gy1;
    for (let k = 1; k < rows.length; k++) { if (bd.bot[k - 1]) continue; const a = rBot(k - 1), b = rTop(k); for (const v of [g.y, gy1]) if (inGap(v, a, b)) { tops[k] = bots[k - 1] = v; break; } }
  }
  return bd;
}
function shadeTable(T, edges, G) { // a fill matching whole cells → cell shading
  if (!G) return;
  const n = T.widths.length, R = T.rows;
  for (const g of G.rects) {
    if (g.used || g.h <= 3) continue; const gx1 = g.x + g.w, gy1 = g.y + g.h;
    if (g.y < R[0].top - 2 || gy1 > R[R.length - 1].bot + 2 || g.x < edges[0] - 2 || gx1 > edges[n] + 2) continue;
    const k0 = R.findIndex(r => Math.abs(r.top - g.y) < 2), k1 = R.findIndex(r => Math.abs(r.bot - gy1) < 2), c0 = edges.findIndex(e => Math.abs(e - g.x) < 2), c1 = edges.findIndex(e => Math.abs(e - gx1) < 2);
    if (k0 < 0 || k1 < k0 || c0 < 0 || c1 <= c0) continue;
    for (let k = k0; k <= k1; k++) for (let c = c0; c < c1; c++) R[k].cells[c].shd = g.fill;
    g.used = true;
  }
}
function layCols(band, flow, o, blocks, maxPitch, G) { // side-by-side text columns → a borderless table with one cell per column
  const cols = band.cols, allRows = cols.flatMap(c => c.bands.flatMap(b => b.kind === 'cols' ? [] : b.rows));
  const top = tableStart(flow, Math.min(...allRows.map(r => r.base - 0.8 * 1.9 * r.size)), blocks);
  const edges = [cols[0].x0, ...cols.slice(1).map((c, k) => (cols[k].x1 + c.x0) / 2), Math.max(o.areaR, cols[cols.length - 1].x1 + 1)];
  const T = { t: 'tbl', ind: edges[0] - o.refL, widths: edges.slice(1).map((e, k) => e - edges[k]), rows: [], cols: true }, cells = []; let bottom = top;
  cols.forEach((c, k) => {
    const fl = new Flow(top), cb = [];
    layBands(c.bands, fl, { refL: edges[k], areaR: edges[k + 1] }, cb, maxPitch, G);
    if (!cb.length || cb[cb.length - 1].t !== 'p') cb.push(spacer(0.05)); // (a cell ends with a paragraph)
    bottom = Math.max(bottom, fl.bottom()); cells.push({ blocks: cb });
  });
  const Ht = Math.max(20, Math.ceil((bottom - top) * 20));
  T.rows.push({ h: Ht, cells, top, bot: top + Ht / 20 });
  blocks.push(T); flow.base = top + Ht / 20; flow.L = 0; flow.lastTable = true;
}
function layBands(bands, flow, o, blocks, maxPitch, G) {
  for (const b of bands) {
    if (b.kind === 'table') { layTable(b, flow, o, blocks, maxPitch, G); continue; }
    if (b.kind === 'cols') { layCols(b, flow, o, blocks, maxPitch, G); continue; }
    let run = [];
    const text = () => { if (!run.length) return; const lines = linesOf(run), colR = Math.max(...lines.map(l => l.x1)), colL = Math.min(...lines.map(l => l.x0)); run = [];
      for (const P of finishParas(paragraphsOf(lines, maxPitch), colL, colR)) layPara(P, flow, { refL: o.refL, colR: P.align === 'both' ? colR : o.areaR, areaR: o.areaR }, blocks); };
    for (const r of b.rows) { if (r.frags[0].box) { text(); layBox(r.frags[0], flow, o, blocks, maxPitch, G); } else run.push(r); }
    text();
  }
}

/* ───────── graphics tied to text become formatting ───────── */
function underlines(rows, G) { // a rule under text → underline, through it → strike-through; a tight fill behind part of a line → highlight (run shading)
  for (const g of G.rects) {
    if (g.used || g.w < 3) continue; const yc = g.y + g.h / 2;
    if (g.h <= 2.5) {
      for (const r of rows) {
        const kind = yc > r.base - 0.1 * r.size && yc < r.base + 0.35 * r.size ? 'u' : yc > r.base - 0.5 * r.size && yc < r.base - 0.15 * r.size ? 's' : null; if (!kind) continue;
        const f = r.frags.find(f => g.x >= f.x0 - 2.5 && g.x + g.w <= f.x1 + 3); if (!f) continue;
        const hit = f.runs.filter(q => Math.min(q.x1, g.x + g.w) - Math.max(q.x0, g.x) > 0.5 * (q.x1 - q.x0)); if (!hit.length) continue;
        for (const q of hit) { q.st = { ...q.st, ...(kind === 'u' ? { u: 'single', ucol: g.fill } : { strike: true }) }; q.key = styleKey(q.st); }
        g.used = true; break;
      }
      continue;
    }
    if (g.geom !== 'rect') continue;
    for (const r of rows) {
      if (g.h < 0.8 * r.size || g.h > 1.7 * r.size || g.y > r.base - 0.6 * r.size || g.y + g.h < r.base + 0.05 * r.size) continue;
      const f = r.frags.find(f => g.x >= f.x0 - 2 && g.x + g.w <= f.x1 + 2); if (!f) continue;
      const hit = f.runs.filter(q => q.x0 >= g.x - 2 && q.x1 <= g.x + g.w + 2);
      if (!hit.length || Math.abs(hit[0].x0 - g.x) > 2 || Math.abs(hit[hit.length - 1].x1 - (g.x + g.w)) > 3) continue;
      for (const q of hit) { q.st = { ...q.st, shd: g.fill }; q.key = styleKey(q.st); }
      g.used = true; break;
    }
  }
}

/* ───────── drawing among the text moves with it ─────────
   Measured in Word: an anchor positioned relativeFrom="paragraph" is measured from the paragraph's top (before its space before);
   a bottom border sits right under the line box (w:space 0), its paragraph is before + line + border high, and the border runs ~1.45 pt past each indent. */
const RULE_OUT = 1.45;
function stamp(blocks, y) { // each block's top and bottom on the page, as Word stacks them
  for (const b of blocks) {
    b.gTop = y;
    if (b.t === 'tbl') { y = b.rows[b.rows.length - 1].bot; if (b.cols || b.box) b.rows[0].cells.forEach(c => stamp(c.blocks, b.rows[0].top)); }
    else y += b.before / 20 + (b.nLines || 1) * b.line / 20 + (b.rule ? b.rule.sz / 8 : 0);
    b.gBot = y;
  }
  return y;
}
const cellsOfT = (b, refL) => { let x = refL + b.ind; return b.widths.map((w, k) => { const c = { l: x, r: x + w, blocks: b.rows[0].cells[k].blocks, top: b.rows[0].top }; x += w; return c; }); };
function ruleInto(blocks, g, refL, areaR) { // a free horizontal rule → an empty paragraph with a bottom border, in the gap it lies in
  for (const b of blocks) if (b.t === 'tbl' && (b.cols || b.box) && g.y > b.gTop && g.y + g.h < b.gBot) {
    const c = cellsOfT(b, refL).find(c => g.x >= c.l - 1 && g.x + g.w <= c.r + 1); return c ? ruleInto(c.blocks, g, c.l, c.r) : false;
  }
  const sz = [2, 4, 6, 8, 12, 18, 24, 36, 48, 72, 96].reduce((a, v) => (Math.abs(v - g.h * 8) < Math.abs(a - g.h * 8) ? v : a)), bw = sz / 8;
  for (let i = 0; i < blocks.length; i++) {
    const B = blocks[i]; if (B.t !== 'p' || B.rule) continue;
    const spc = B.tiny && !B.anchor && !B.runs.length, free = spc ? B.gBot : B.gTop + B.before / 20;
    if (g.y < B.gTop - 0.01 || g.y + bw > free + 0.01) continue;
    const pr = blocks[i - 1]; if (pr && pr.rule && pr.rule.sz === sz && Math.abs(pr.rule.x - g.x) < 1 && Math.abs(pr.rule.w - g.w) < 1) return false; // (Word would merge two same borders)
    const bt = Math.max(0, twip(g.y - B.gTop - 0.05)), used = bt + 1 + 2.5 * sz;
    if (spc ? B.line - used < 1 : B.before - used < 0) return false;
    if (spc) B.line -= used; else B.before -= used;
    blocks.splice(i, 0, { t: 'p', before: bt, line: 1, ind: { left: g.x + RULE_OUT - refL, first: 0, right: areaR - (g.x + g.w - RULE_OUT) }, runs: [], tiny: true, rule: { sz, color: g.fill, x: g.x, w: g.w }, nLines: 1 });
    return true;
  }
  return false;
}
function hostIn(blocks, el, refL, cell) { // the paragraph a drawing hangs from: the last one starting above it (inside a column or box when it lies in one)
  for (const b of blocks) if (b.t === 'tbl' && (b.cols || b.box) && el.y >= b.gTop - 2 && el.y + el.h <= b.gBot + 2) { // (a frame's corner may stick out a little)
    const ov = c => Math.min(c.r, el.x + el.w) - Math.max(c.l, el.x), c = cellsOfT(b, refL).sort((p, q) => ov(q) - ov(p))[0]; // (the cell holding most of it; a box's frame may stand out of the text column)
    if (c && ov(c) >= 0.6 * el.w) { const h = hostIn(c.blocks, el, c.l, c); if (h) return h; }
  }
  let B = null; for (const b of blocks) if (b.t === 'p' && b.gTop <= el.y + 0.01) B = b;
  if (!B && cell) B = blocks.find(b => b.t === 'p') || null; // (just above a cell's first paragraph: hangs from it)
  if (B && el.y >= B.gBot - 0.01) { // below that paragraph's text and not within the table after it (above it, or reaching below it): it hangs from the paragraph after the table
    const i = blocks.indexOf(B), T = blocks[i + 1], nx = blocks.slice(i + 1).find(b => b.t === 'p');
    if (T && T.t === 'tbl' && nx && !(el.y >= T.gTop - 2 && el.y + el.h <= T.gBot + 2)) B = nx;
  }
  return B ? { B, list: blocks, dy: el.y - B.gTop, cellL: cell ? cell.l : null } : null;
}
function flowDrawing(blocks, G, o) {
  stamp(blocks, o.MT);
  for (const g of G.rects.filter(g => !g.used && g.h <= 3 && g.w >= 8 && g.w >= 4 * g.h && g.y > o.MT && g.y + g.h < o.bottom)) { if (ruleInto(blocks, g, o.refL, o.areaR)) { g.used = true; stamp(blocks, o.MT); } }
  const els = G.rects.filter(g => !g.used).concat(G.pics), area = g => g.w * g.h, PA = o.PW * o.PH;
  if (els.length > 600) return;
  const inFlow = new Set(els.filter(g => !g.boxOf && g.y >= o.MT - 0.5 && g.y + g.h <= o.bottom + 0.5 && area(g) < 0.4 * PA));
  const boxed = els.filter(g => g.boxOf), firstP = bs => bs.find(b => b.t === 'p');
  for (let ch = true; ch;) { // a shape that moves must carry what lies on it (else it would cover it)
    ch = false;
    for (const E of [...inFlow]) if (els.some(S => S !== E && !inFlow.has(S) && area(S) <= area(E) && S.x < E.x + E.w && S.x + S.w > E.x && S.y < E.y + E.h && S.y + S.h > E.y)) { inFlow.delete(E); ch = true; }
  }
  const own = new Map(); // a paragraph with text gets a 1-twip empty paragraph before it to hold its drawing: deleting the text then keeps the drawing
  for (const g of [...boxed, ...inFlow]) {
    let h = null;
    if (g.boxOf) { const cb = g.boxOf.rows[0].cells[0].blocks, B = firstP(cb); if (B) h = { B, list: cb, dy: g.y - B.gTop, cellL: g.boxOf.cellL }; } // (the box's own fill or picture: from its first paragraph)
    else h = hostIn(blocks, g, o.refL, null);
    if (!h) continue;
    if (h.B.runs.length && !h.B.tiny) {
      let A = own.get(h.B);
      if (!A) { A = { t: 'p', before: 0, line: 1, ind: null, runs: [], tiny: true, anchor: true, keep: true, nLines: 1, gTop: h.B.gTop }; h.list.splice(h.list.indexOf(h.B), 0, A); if (h.B.before >= 1) h.B.before -= 1; own.set(h.B, A); }
      h.B = A;
    } else if (h.B.gapSp) h.B.keep = true; // (a figure in a gap stays with the caption under it)
    g.host = h;
  }
}

/* ───────── one page → blocks for the body, its header and footer ───────── */
async function layoutPage(p, ctx, pi, opt = {}) { // opt.noFoot: the footer lines stay in the body
  const M = await pageModel(p), { PW, PH } = M;
  const straightAll = M.pieces.filter(q => !q.rot), scan = straightAll.length > 0 && straightAll.every(q => q.ghost);
  const straight = straightAll.filter(q => scan || !q.ghost), turned = M.pieces.filter(q => q.rot && !q.ghost);
  await loadFonts(straight.map(q => styleOfPiece(q, false)));
  const G = scan ? { rects: [], pics: [bboxPicOf(M)].filter(Boolean) } : drawing(M);
  for (const q of straight) if (q.fi.embed) ctx.embeds.set(q.fi.embed, ctx.embeds.get(q.fi.embed) || !!q.bold);
  let rows = rowsOf(straight, scan);
  const bodyPt = median(rows.flatMap(r => r.frags.flatMap(f => f.runs.flatMap(q => Array(Math.min(50, q.t.length)).fill(q.st.size))))) || 11;
  const boxed = f => { // ink of the drawing on all four sides of a mark, close by: a box around it
    const cx = (f.x0 + f.x1) / 2, cy = f.base - 0.35 * f.size, R = 1.3 * f.size, ink = (x, y) => { const X0 = Math.round(x * K), Y0 = Math.round(y * K); if (X0 < 0 || Y0 < 0 || X0 >= M.W || Y0 >= M.H) return false; const j = (Y0 * M.W + X0) * 4; return M.d[j] < 235 || M.d[j + 1] < 235 || M.d[j + 2] < 235; };
    const side = (dx, dy) => { for (let s = 0.25 * f.size; s <= R; s += 0.5 / K) if (ink(cx + dx * s, cy + dy * s)) return true; return false; };
    return side(1, 0) && side(-1, 0) && side(0, 1) && side(0, -1);
  };
  markLists(rows, bodyPt, boxed);
  for (const h of G.hidden || []) { // a button inside a picture (a pill on a gradient) holding its label → a shape of its own that moves with the label; the picture gets the background painted in under it
    const P = G.pics[h.pk]; if (!P || h.w < 20 || h.h < 8 || h.w * h.h > 0.25 * P.w * P.h) continue;
    if (!rows.some(r => r.frags.some(f => f.x0 >= h.x - 0.5 && f.x1 <= h.x + h.w + 0.5 && f.base - 0.8 * f.size >= h.y - 0.5 && f.base + 0.25 * f.size <= h.y + h.h + 0.5))) continue;
    h.over = true; G.rects.push(h); (P.inpaint = P.inpaint || []).push(h.px);
  }
  const xl = Math.min(...rows.map(r => r.frags[0].x0)), xr = Math.max(...rows.map(r => r.frags[r.frags.length - 1].x1));
  // a justified line with very wide word gaps (lines ending in a manual break, short lines in narrow columns) is one line, not a row of cells:
  // its words stand evenly apart, and it spans the whole page or the same width as a line just above or below it
  const spans = (o, a, b) => o.frags.some(f => Math.abs(f.x0 - a) < 2 && b - f.x1 > -2.5 && b - f.x1 < Math.max(2.5, 0.45 * o.size)); // (a neighbour may end a trimmed space short)
  const even = gaps => { const u = Math.min(...gaps); return u > 0 && !gaps.some(g => { const k = Math.round(g / u); return k > 2 || Math.abs(g - k * u) > 0.15 * k * u + 1; }); }; // (one gap may be double: two spaces after a full stop)
  const joinFr = (r, i, j) => { const f = r.frags[i]; for (const g of r.frags.slice(i + 1, j + 1)) { const l = f.runs[f.runs.length - 1]; f.runs.push({ t: ' ', st: l.st, key: l.key, x0: f.x1, x1: g.x0, gap: true }, ...g.runs); f.x1 = g.x1; f.size = Math.max(f.size, g.size); } r.frags.splice(i + 1, j - i); };
  for (let pass = 0; pass < 2; pass++) for (const r of rows) {
    const fs = r.frags; if (fs.length < 2) continue;
    if (fs.length >= 4 && !fs.some(f => f.marker || /\s/.test(textOf(f).trim())) && fs[0].x0 <= xl + 2 && fs[fs.length - 1].x1 >= xr - 2 && even(fs.slice(1).map((f, k) => f.x0 - fs[k].x1)) && Math.max(...fs.slice(1).map((f, k) => f.x0 - fs[k].x1)) <= 3 * r.size) { joinFr(r, 0, fs.length - 1); continue; }
    const nb = rows.filter(o => o !== r && Math.abs(o.base - r.base) < 2.6 * r.size);
    for (let i = 0; i < r.frags.length - 1; i++) for (let j = r.frags.length - 1; j > i; j--) { // (part of a row: the other part may be another column's text)
      const ps = r.frags.slice(i, j + 1), gaps = ps.slice(1).map((f, k) => f.x0 - ps[k].x1);
      if (ps.some(f => f.marker) || Math.max(...gaps) > 3 * r.size || !even(gaps) || !nb.some(o => spans(o, ps[0].x0, ps[ps.length - 1].x1))) continue;
      joinFr(r, i, j); break;
    }
  }
  underlines(rows, G);
  const rightmost = Math.max(...rows.map(r => r.frags[r.frags.length - 1].x1)), ratios = []; // (line pitch: from lines that wrap onto the next one)
  for (let i = 1; i < rows.length; i++) { const a = rows[i - 1], b = rows[i]; if (a.frags.length === 1 && b.frags.length === 1 && Math.abs(a.size - b.size) < 0.3 && a.frags[0].x1 > rightmost - 3 * a.size) { const r = (b.base - a.base) / b.size; if (r > 0.95 && r < 2.2) ratios.push(r); } }
  ctx.maxPitch = ratios.length >= 3 ? clamp(median(ratios) * 1.15, 1.3, 2.2) : 1.6; // (lines further apart than the page's usual pitch start a new paragraph)
  // header and footer lines (the same text on most pages) go to the page's header / footer
  const hfKey = r => r.frags.map(textOf).join('').replace(/\s+/g, '').replace(/\d+/g, '#');
  const hRows = rows.filter(r => r.base < PH * 0.09 && ctx.hf.head.has(hfKey(r))), fRows = opt.noFoot ? [] : rows.filter(r => r.base > PH * 0.9 && ctx.hf.foot.has(hfKey(r)));
  rows = rows.filter(r => !hRows.includes(r) && !fRows.includes(r));
  // this page's own foot: lines at the bottom of the page standing well apart from the text (a company's legal line, "Page 1 of 1", footnotes)
  // → the page's footer, so text added above them never pushes them onto a page of their own
  let zRows = [], zTop = PH;
  const markOf = r => { const lead = r.frags[0].runs[0]; return (lead.st.pos > 0 && /^\s*(\d{1,3}|[*†‡§¶])\s*$/.test(lead.t)) || (r.size < 0.95 * bodyPt && /^\s*(\d{1,3}|[*†‡§¶])\s/.test(textOf(r.frags[0]))); };
  if (!opt.noFoot && rows.length > 2) {
    const srt = [...rows].sort((a, b) => a.base - b.base), tp = r => r.base - 0.95 * r.size, bt = r => r.base + 0.3 * r.size;
    for (let s = srt.length - 1; s >= 1; s--) {
      const r = srt[s], tail = srt.slice(s); if (tp(r) < 0.72 * PH || tail.some(t => t.size > bodyPt + 0.3) || PH - tp(r) > 0.22 * PH) break;
      const above = Math.max(...srt.slice(0, s).map(bt)), gap = tp(r) - above; if (gap <= 0) continue;
      const mark = markOf(r), inGap = g => g.y > above - 0.5 && g.y + g.h < tp(r) + 0.5, sep = G.rects.some(g => g.h <= 2 && g.w >= 20 && inGap(g));
      const pic = G.pics.concat(G.rects.filter(g => g.h > 2)).some(g => g.y > above - 2 && g.y + g.h < tp(r) + 2 && g.h > 1.2 * bodyPt); // (a figure in the gap: the line is its caption)
      if (pic || srt[srt.length - 1].base < 0.85 * PH) continue;
      if (gap >= Math.max(3 * bodyPt, 24) || (mark && gap >= 0.8 * bodyPt && (sep || r.size < 0.95 * bodyPt || gap >= 1.6 * bodyPt))) { zRows = tail; zTop = Math.min(tp(r), ...G.rects.filter(g => g.h <= 2 && inGap(g)).map(g => g.y - 0.5)); }
    }
    rows = rows.filter(r => !zRows.includes(r));
    let prev = null; for (const r of zRows) { r.note = markOf(r) || !!(prev && prev.note && r.base - prev.base < 1.6 * r.size && !(r.frags.length > 1)); prev = r; } // (footnotes: this page only; a legal line goes on every page)
  }
  const xs = [...rows.flatMap(r => r.frags.map(f => f.marker ? f.marker.x : f.x0)), ...hRows.concat(fRows).map(r => r.frags[0].x0)], x1s = rows.concat(hRows, fRows).flatMap(r => r.frags.map(f => f.x1));
  const ML = xs.length ? Math.max(0, Math.min(...xs) - 0.5) : 36, MR = x1s.length ? clamp(PW - Math.max(...x1s) - 8, 4, 36) : 36, areaR = PW - MR;
  // header: an anchor paragraph for the drawing, then the running header lines
  const hb = [], hflow = new Flow(0); hb.push({ t: 'p', before: 0, line: 20, runs: [], tiny: true, anchor: true }); hflow.base = 0.8; hflow.L = 1;
  for (const r of hRows) layPara(finishParas([{ lines: linesOf([r]), multi: r.frags.length > 1, gaps: [], pitch: null }], ML, areaR)[0], hflow, { refL: ML, colR: areaR, areaR }, hb);
  const hdrBottom = hflow.bottom();
  let fb = null, fbRun = null, fd = 0, ftrTop = PH;
  const layFoot = rs => { // footer lines → paragraphs, from the first line's top
    const out = [], r0 = rs[0], fl = new Flow(r0.base - 0.96 * r0.size), lines = linesOf(rs), cR = Math.max(...lines.map(l => l.x1));
    for (const P of finishParas(paragraphsOf(lines, ctx.maxPitch), ML, areaR)) layPara(P, fl, { refL: ML, colR: P.align === 'both' ? cR : areaR, areaR }, out);
    return { out, fl };
  };
  const footRows = fRows.concat(zRows).sort((a, b) => a.base - b.base);
  if (footRows.length) {
    const F = layFoot(footRows); fb = F.out; ftrTop = F.fl.top; fd = Math.max(0, PH - F.fl.bottom());
    const run = footRows.filter(r => !r.note); fbRun = run.length === footRows.length ? fb : run.length ? layFoot(run).out : []; // (on a page the text runs onto: no footnotes)
  }
  for (const b of hb.concat(fb || [], fbRun || [])) if (b.runs) b.runs.forEach(r => { r.hf = true; });
  // body
  rows.forEach(r => r.frags.forEach(f => { f.row = r; }));
  curWords = new Set(rows.flatMap(r => r.frags.flatMap(f => textOf(f).toLowerCase().match(/\p{L}+/gu) || [])));
  // boxes: fills holding text. A band at the very top starts under the header; a big picture behind text (a gradient hero) is a box too, its picture moving with it
  const y0 = hdrBottom + 0.5, px = (g, x, y) => { const j = (Math.min(M.H - 1, Math.round(y * K)) * M.W + Math.min(M.W - 1, Math.round(x * K))) * 4; return [M.d[j], M.d[j + 1], M.d[j + 2]]; };
  const picBox = g => { // a picture that is all colour (no white page showing through) → { fill: the colour along its bottom }
    let n = 0, ink = 0; const cs = []; for (let y = g.y + 1; y < g.y + g.h - 1; y += Math.max(1, g.h / 24)) for (let x = g.x + 1; x < g.x + g.w - 1; x += Math.max(1, g.w / 24)) { const c = px(g, x, y); n++; if (c[0] < 245 || c[1] < 245 || c[2] < 245) ink++; }
    if (ink < 0.95 * n) return null;
    for (let x = g.x + 1; x < g.x + g.w - 1; x += Math.max(1, g.w / 40)) cs.push(px(g, x, g.y + g.h - 2));
    const m = k => median(cs.map(c => c[k])); return '#' + [m(0), m(1), m(2)].map(v => v.toString(16).padStart(2, '0')).join('');
  };
  G.PW = PW; G.boxes = [];
  for (const r of G.rects) if (!r.used && r.geom === 'rect' && r.w >= 30 && r.h >= 8 && r.y + r.h <= PH - 8) { if (r.y >= y0) G.boxes.push(r); else if (r.y + r.h > y0 + 12) G.boxes.push({ x: r.x, y: y0, w: r.w, h: r.y + r.h - y0, fill: r.fill, geom: 'rect', src: r }); }
  for (const g of G.pics) if (g.w >= 0.3 * PW && g.h >= 30 && g.y + g.h <= PH - 8 && g.y + g.h > y0 + 12) { const fill = picBox(g); if (fill) G.boxes.push({ x: g.x, y: Math.max(g.y, y0), w: g.w, h: g.y + g.h - Math.max(g.y, y0), fill, geom: 'rect', pic: g }); }
  G.boxes.sort((a, b) => b.w * b.h - a.w * a.h);
  const top = rows.length ? boxify(rows.flatMap(r => r.frags), G.boxes) : [], maxPitch = ctx.maxPitch, bands = top.length ? segment(top, PW) : [], blocks = [];
  const first = rows[0], tops = top.map(f => f.box ? f.box.y : f.base - 0.96 * f.size - 2), MT = Math.max(hdrBottom + 0.1, (tops.length ? Math.min(...tops) : 36) - (turned.length ? 1 : 0)), flow = new Flow(MT);
  if (turned.length) { blocks.push({ t: 'p', before: 0, line: 20, runs: [], tiny: true, anchor: true }); flow.base = MT + 0.8; flow.L = 1; } // (1 pt paragraph holding the turned text boxes)
  layBands(bands, flow, { refL: ML, areaR }, blocks, maxPitch, G);
  const endSp = !blocks.length || blocks[blocks.length - 1].t !== 'p'; if (endSp) blocks.push(spacer(1));
  const bodyBottom = flow.bottom(), bodyEnd = bodyBottom + (endSp ? 1 : 0), lastSize = first ? rows[rows.length - 1].size : 12;
  let MB = Math.max(0, Math.min(6, PH - bodyBottom - 1.2 * lastSize));
  if (fb) { const ft = Math.min(ftrTop, zTop); MB = Math.max(MB, PH - Math.min(ft - 0.5, Math.max(bodyBottom + 0.5, ft - 6))); } // (the text stops a little above the footer, so a table border never reaches it)
  flowDrawing(blocks, G, { refL: ML, areaR, MT, PW, PH, bottom: PH - MB });
  // confidence: paragraphs Word cannot place (negative space before) or text lines lying over one another → this page keeps text boxes
  const paras = []; const walk = bs => bs.forEach(b => { if (b.t === 'p' && b.runs.length) paras.push(b); else if (b.t === 'tbl') b.rows.forEach(r => r.cells.forEach(c => walk(c.blocks))); }); walk(blocks);
  let neg = 0; const negWalk = bs => bs.forEach(b => { if (b.neg) neg++; }); negWalk(paras);
  let over = 0; for (let i = 1; i < rows.length; i++) { const a = rows[i - 1], b = rows[i]; if (b.base - a.base < 0.6 * Math.min(a.size, b.size) && a.frags.some(f => b.frags.some(g => g.x0 < f.x1 && g.x1 > f.x0))) over++; }
  if (fb && bodyBottom > PH - MB + 1 && bodyBottom <= PH - Math.min(MB, 6) + 1) return layoutPage(p, ctx, pi, { noFoot: true }); // (the footer's room alone makes the page overflow)
  const lowConf = (neg >= Math.max(3, 0.25 * paras.length)) || over >= Math.max(3, 0.2 * rows.length) || M.type3 > 0.5 || bodyBottom > PH - MB + 1;
  const nT = turned.reduce((n, q) => n + q.str.length, 0), nS = straight.reduce((n, q) => n + q.str.length, 0);
  return { M, PW, PH, ML, MR, MT, MB, fd, blocks, hb, fb, fbRun, G, turned, bodyEnd, lowConf, turnedShare: nT / ((nT + nS) || 1), stats: { neg, over, paras: paras.length } };
}
function bboxPicOf(M) { const { W, H, d } = M; let x0 = W, y0 = H, x1 = -1, y1 = -1; for (let y = 0; y < H; y++) for (let x = 0, i = y * W * 4; x < W; x++, i += 4) if (d[i] < 250 || d[i + 1] < 250 || d[i + 2] < 250) { if (x < x0) x0 = x; if (x > x1) x1 = x; if (y < y0) y0 = y; y1 = y; } return x1 < 0 ? null : { x: x0 / K, y: y0 / K, w: (x1 - x0 + 1) / K, h: (y1 - y0 + 1) / K, px: { x0, y0, x1: x1 + 1, y1: y1 + 1 } }; }

/* ───────── header / footer text: the same lines at the top (bottom) of most pages ───────── */
async function findHF(pages) {
  const head = new Set(), foot = new Set(); if (pages.length < 2) return { head, foot };
  const hc = new Map(), fc = new Map(), add = (m, k) => m.set(k, (m.get(k) || 0) + 1);
  for (const p of pages) {
    if (p.src < 0) continue;
    try {
      const pg = await X.getPdfPage(p), vp = pg.getViewport({ scale: 1, rotation: (pg.rotate + (p.rot || 0)) % 360 }), tc = await pg.getTextContent(), rows = new Map();
      for (const it of tc.items) { if (!it.str || !it.str.trim()) continue; const m = pdfjsLib.Util.transform(vp.transform, it.transform); const b = Math.round(m[5]); const r = rows.get(b) || []; r.push([m[4], it.str]); rows.set(b, r); }
      const hs = new Set(), fs = new Set();
      for (const [b, r] of rows) { const k = r.sort((a, c) => a[0] - c[0]).map(v => v[1]).join('').replace(/\s+/g, '').replace(/\d+/g, '#'); if (!k) continue; if (b < vp.height * 0.09) hs.add(k); else if (b > vp.height * 0.9) fs.add(k); }
      hs.forEach(k => add(hc, k)); fs.forEach(k => add(fc, k));
    } catch (e) { /* page unreadable */ }
  }
  const need = Math.max(2, Math.ceil(pages.length * 0.6));
  for (const [k, n] of hc) if (n >= need) head.add(k); for (const [k, n] of fc) if (n >= need) foot.add(k);
  return { head, foot };
}

/* ───────── the .docx package ───────── */
const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" mc:Ignorable="w14"';
const XH = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>', RT = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/', CT = 'application/vnd.openxmlformats-officedocument.wordprocessingml';
function makePackage() {
  const zip = new JSZip(), media = new Map(), rels = {}, put = (n, d, o) => zip.file(n, d, { createFolders: false, ...o }); let nMedia = 0, nId = 0;
  const relsOf = part => rels[part] || (rels[part] = []);
  const image = (part, src) => {
    const url = typeof src === 'string', key = url ? src : src.key; let m = media.get(key);
    if (!m) { const ext = url ? (/^data:image\/jpe?g/.test(src) ? 'jpeg' : 'png') : src.ext; m = 'media/image' + (++nMedia) + '.' + ext; put('word/' + m, url ? src.split(',')[1] : src.bytes, url ? { base64: true } : {}); media.set(key, m); }
    const R = relsOf(part); let r = R.find(x => x.target === m); if (!r) { r = { id: 'rIdM' + (R.length + 1), type: 'image', target: m }; R.push(r); } return r.id;
  };
  const id = () => ++nId;
  // rel (a drawing that moves with the text): { dy: below its paragraph's top, cellL: the left edge of the table cell it is in (Word then measures from the cell) }
  const anchorOpen = (n, x, y, w, h, behind, z, rel) => `<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="${z}" behindDoc="${behind ? 1 : 0}" locked="1" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/>${rel && rel.cellL != null ? `<wp:positionH relativeFrom="column"><wp:posOffset>${emu(x - rel.cellL)}</wp:posOffset></wp:positionH>` : `<wp:positionH relativeFrom="page"><wp:posOffset>${emu(x)}</wp:posOffset></wp:positionH>`}${rel ? `<wp:positionV relativeFrom="paragraph"><wp:posOffset>${emu(rel.dy)}</wp:posOffset></wp:positionV>` : `<wp:positionV relativeFrom="page"><wp:posOffset>${emu(y)}</wp:posOffset></wp:positionV>`}<wp:extent cx="${emu(w)}" cy="${emu(h)}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>`;
  const picXml = (part, src, w, h, at, z) => {
    const rid = image(part, src), n = id(), ext = `cx="${emu(w)}" cy="${emu(h)}"`;
    const g = `<wp:docPr id="${n}" name="${at ? 'Page design' : 'Picture'} ${n}"/><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="${n}" name="Picture ${n}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext ${ext}/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic>`;
    if (!at) return `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent ${ext}/><wp:effectExtent l="0" t="0" r="0" b="0"/>${g}</wp:inline></w:drawing></w:r>`;
    return anchorOpen(n, at.x, at.y, w, h, true, z || n, at.rel) + g + '</wp:anchor></w:drawing></w:r>';
  };
  const shapeXml = (s, z, rel) => {
    const n = id(), w = Math.max(0.05, s.w), h = Math.max(0.05, s.h), ext = `cx="${emu(w)}" cy="${emu(h)}"`, geom = s.geom || 'rect';
    const av = geom === 'roundRect' ? `<a:avLst><a:gd name="adj" fmla="val ${s.adj || 16667}"/></a:avLst>` : '<a:avLst/>';
    return anchorOpen(n, s.x, s.y, w, h, true, z, rel) + `<wp:docPr id="${n}" name="Shape ${n}"/><wp:cNvGraphicFramePr/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext ${ext}/></a:xfrm><a:prstGeom prst="${geom}">${av}</a:prstGeom><a:solidFill><a:srgbClr val="${s.fill.slice(1).toUpperCase()}"/></a:solidFill><a:ln><a:noFill/></a:ln></wps:spPr><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>`;
  };
  const tboxXml = (x, y, w, h, rot, inner) => { // an invisible text box pinned to the page (turned around its centre)
    const n = id(), ext = `cx="${emu(w)}" cy="${emu(h)}"`, r = Math.round((((rot % 360) + 360) % 360) * 60000);
    return `<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="${5000 + n}" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="page"><wp:posOffset>${emu(x)}</wp:posOffset></wp:positionH><wp:positionV relativeFrom="page"><wp:posOffset>${emu(y)}</wp:posOffset></wp:positionV><wp:extent ${ext}/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/><wp:docPr id="${n}" name="Text ${n}"/><wp:cNvGraphicFramePr/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm${r ? ` rot="${r}"` : ''}><a:off x="0" y="0"/><a:ext ${ext}/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></wps:spPr><wps:txbx><w:txbxContent>${inner}</w:txbxContent></wps:txbx><wps:bodyPr rot="0" vert="horz" wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" anchor="t" anchorCtr="0"><a:spAutoFit/></wps:bodyPr></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>`;
  };
  return { zip, put, image, relsOf, picXml, shapeXml, tboxXml, id };
}
async function embedFonts(pk, fams) { // bundled TrueType families Word won't have → obfuscated .odttf parts (ECMA-376 font embedding)
  const out = [];
  for (const [fam, bold] of fams) {
    const B = X.BUNDLED[fam]; if (!B) continue; const faces = [['Regular', 'Regular'], bold && B[1] ? ['Bold', 'Bold'] : null].filter(Boolean); let x = '';
    for (const [file, tag] of faces) {
      let bytes; try { const r = await fetch(`lib/fonts/${B[0]}-${file}.ttf`); if (!r.ok) continue; bytes = new Uint8Array(await r.arrayBuffer()); } catch (e) { continue; }
      const sig = (bytes[0] << 24 | bytes[1] << 16 | bytes[2] << 8 | bytes[3]) >>> 0; if (sig !== 0x00010000 && sig !== 0x74727565) continue; // (TrueType outlines only)
      if (fsType(bytes) & 2) continue; // (licence forbids embedding)
      const g = crypto.getRandomValues(new Uint8Array(16)), hex = [...g].map(v => v.toString(16).padStart(2, '0').toUpperCase()).join(''), guid = `${hex.slice(0, 8)}-${hex.slice(8, 12)}-${hex.slice(12, 16)}-${hex.slice(16, 20)}-${hex.slice(20)}`;
      const key = []; for (let i = 0; i < 32; i += 2) key.push(parseInt(hex.slice(i, i + 2), 16)); key.reverse();
      const b = bytes.slice(); for (let i = 0; i < 32; i++) b[i] ^= key[i % 16];
      const n = out.length + 1; pk.put(`word/fonts/font${n}.odttf`, b); out.push({ id: 'rIdF' + n, target: `fonts/font${n}.odttf` });
      x += `<w:embed${tag} r:id="rIdF${n}" w:fontKey="{${guid}}"/>`;
    }
    if (x) out.xml = (out.xml || '') + `<w:font w:name="${xesc(fam)}"><w:charset w:val="00"/><w:family w:val="auto"/><w:pitch w:val="variable"/>${x}</w:font>`;
  }
  return out;
}
function fsType(b) { try { const n = b[4] << 8 | b[5]; for (let i = 0; i < n; i++) { const o = 12 + 16 * i; if (String.fromCharCode(b[o], b[o + 1], b[o + 2], b[o + 3]) === 'OS/2') { const t = (b[o + 8] << 24 | b[o + 9] << 16 | b[o + 10] << 8 | b[o + 11]) >>> 0; return b[t + 8] << 8 | b[t + 9]; } } } catch (e) { /* unreadable */ } return 0; }

/* ───────── writing XML ───────── */
function rPrXml(st, base, r) {
  if (!st) return '';
  let x = '';
  if (st.name !== base.name) x += `<w:rFonts w:ascii="${xesc(st.name)}" w:hAnsi="${xesc(st.name)}" w:eastAsia="${xesc(st.name)}" w:cs="${xesc(st.name)}"/>`;
  if (!!st.b !== !!base.b) x += st.b ? '<w:b/><w:bCs/>' : '<w:b w:val="0"/><w:bCs w:val="0"/>';
  if (!!st.i !== !!base.i) x += st.i ? '<w:i/><w:iCs/>' : '<w:i w:val="0"/><w:iCs w:val="0"/>';
  if (st.strike) x += '<w:strike/>';
  if (st.ghost) x += '<w:noProof/>';
  const c = st.ghost ? null : (st.color || '#000000').toLowerCase(); if (c && c !== (base.color || '#000000').toLowerCase()) x += `<w:color w:val="${c.slice(1).toUpperCase()}"/>`;
  if (r && r.sp) x += `<w:spacing w:val="${r.sp}"/>`;
  if (r && r.w && r.w !== 100) x += `<w:w w:val="${r.w}"/>`;
  if (st.pos) x += `<w:position w:val="${st.pos}"/>`;
  const sz = Math.max(2, Math.round(st.size * 2)); if (sz !== base.sz) x += `<w:sz w:val="${sz}"/><w:szCs w:val="${sz}"/>`;
  if (st.u) x += `<w:u w:val="${st.u}"${st.ucol && st.ucol.toLowerCase() !== (st.color || '#000000').toLowerCase() ? ` w:color="${st.ucol.slice(1).toUpperCase()}"` : ''}/>`;
  if (st.shd) x += `<w:shd w:val="clear" w:color="auto" w:fill="${st.shd.slice(1).toUpperCase()}"/>`;
  if (st.ghost) x += GHOST;
  return x ? `<w:rPr>${x}</w:rPr>` : '';
}
function runsXml(B, base, ctx) {
  let out = ''; const pageNo = ctx && (ctx.pages > 1 || /\bpage\s*\d/i.test(B.runs.map(r => r.t || '').join(''))); // ("Page 1 of 1" counts on when the text runs onto a second page)
  for (const r of B.runs) {
    if (r.tab) { out += '<w:r><w:tab/></w:r>'; continue; }
    const rp = rPrXml(r.st, base, r); let t = r.t;
    if (r.shy) t = t.replace(/[-\u2010]$/, '');
    if (r.hf && pageNo) { // page numbers in a running header / footer → Word fields
      const parts = t.split(/(\d+)/); let done = '';
      for (const s of parts) {
        if (!s) continue;
        const f = /^\d+$/.test(s) ? (+s === ctx.pageNo && !done.includes('P') ? 'PAGE' : +s === ctx.pages && !done.includes('N') ? 'NUMPAGES' : null) : null;
        if (f) { // (a complex field: each part keeps the run's look when Word updates it)
          done += f[0]; const fc = k => `<w:r>${rp}<w:fldChar w:fldCharType="${k}"/></w:r>`;
          out += fc('begin') + `<w:r>${rp}<w:instrText xml:space="preserve"> ${f} \\* MERGEFORMAT </w:instrText></w:r>` + fc('separate') + `<w:r>${rp}<w:t>${s}</w:t></w:r>` + fc('end');
        } else out += `<w:r>${rp}<w:t xml:space="preserve">${xesc(s)}</w:t></w:r>`;
      }
    } else if (t) out += `<w:r>${rp}<w:t xml:space="preserve">${xesc(t)}</w:t>${r.shy ? '<w:softHyphen/>' : ''}</w:r>`;
    else if (r.shy) out += `<w:r>${rp}<w:softHyphen/></w:r>`;
    if (r.br) out += '<w:r><w:br/></w:r>';
  }
  return out;
}
function pXml(B, ctx, sect, anchors) {
  let ppr = '';
  const base = ctx.styleOf(B);
  if (base.id !== 'Normal') ppr += `<w:pStyle w:val="${base.id}"/>`;
  if (B.keep) ppr += '<w:keepNext/>';
  if (B.num) ppr += `<w:numPr><w:ilvl w:val="0"/><w:numId w:val="${ctx.num.id(B, ctx.cont)}"/></w:numPr>`;
  if (B.rule) ppr += `<w:pBdr><w:bottom w:val="single" w:sz="${B.rule.sz}" w:space="0" w:color="${B.rule.color.slice(1).toUpperCase()}"/></w:pBdr>`;
  if (B.tabs && B.tabs.length) ppr += `<w:tabs>${B.tabs.map((x, i) => `<w:tab w:val="left"${B.lead && B.lead[i] ? ` w:leader="${B.lead[i]}"` : ''} w:pos="${twip(x)}"/>`).join('')}</w:tabs>`;
  ppr += `<w:spacing w:before="${B.before}" w:after="0" w:line="${B.line}" w:lineRule="exact"/>`;
  if (B.ind) { const i = B.ind; ppr += `<w:ind w:left="${twip(i.left)}" w:right="${twip(i.right || 0)}"${i.first > 0.3 ? ` w:firstLine="${twip(i.first)}"` : i.first < -0.3 ? ` w:hanging="${twip(-i.first)}"` : ''}/>`; }
  if (B.jc && B.jc !== 'left') ppr += `<w:jc w:val="${B.jc}"/>`;
  const lastSt = B.runs.length ? (B.runs.filter(r => r.st).pop() || {}).st : null;
  ppr += B.tiny || !lastSt ? '<w:rPr><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr>' : rPrXml(lastSt, base, null);
  return `<w:p><w:pPr>${ppr}${sect || ''}</w:pPr>${anchors || ''}${B.anch || ''}${runsXml(B, base, ctx)}</w:p>`;
}
function tblXml(T, ctx) {
  const tw = T.widths.map(twip), NIL = ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map(k => `<w:${k} w:val="nil"/>`).join('');
  const rows = T.rows.map(r => `<w:tr><w:trPr>${T.cols || T.box ? '' : '<w:cantSplit/>'}<w:trHeight w:val="${r.h}" w:hRule="atLeast"/></w:trPr>${r.cells.map((c, ci) => {
    const bd = c.bd ? `<w:tcBorders>${['top', 'left', 'bottom', 'right'].filter(s => c.bd[s]).map(s => `<w:${s} w:val="single" w:sz="${c.bd[s].sz}" w:space="0" w:color="${c.bd[s].color.slice(1).toUpperCase()}"/>`).join('')}</w:tcBorders>` : '';
    const inner = c.blocks.length ? blocksXml(c.blocks, ctx, '', '') :'<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="20" w:lineRule="exact"/><w:rPr><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr></w:pPr></w:p>';
    return `<w:tc><w:tcPr><w:tcW w:w="${tw[ci]}" w:type="dxa"/>${bd}${c.shd ? `<w:shd w:val="clear" w:color="auto" w:fill="${c.shd.slice(1).toUpperCase()}"/>` : ''}</w:tcPr>${inner}</w:tc>`;
  }).join('')}</w:tr>`).join('');
  return `<w:tbl><w:tblPr><w:tblW w:w="${tw.reduce((a, b) => a + b, 0)}" w:type="dxa"/><w:tblInd w:w="${twip(T.ind)}" w:type="dxa"/><w:tblBorders>${NIL}</w:tblBorders><w:tblLayout w:type="fixed"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar><w:tblLook w:val="0000" w:firstRow="0" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:noHBand="1" w:noVBand="1"/></w:tblPr><w:tblGrid>${tw.map(w => `<w:gridCol w:w="${w}"/>`).join('')}</w:tblGrid>${rows}</w:tbl>`;
}
function blocksXml(blocks, ctx, sect, anchors) {
  let out = '', anch = anchors || '', prev = null;
  blocks.forEach((b, i) => {
    const last = i === blocks.length - 1;
    ctx.cont = !!(prev && prev.t === 'p' && prev.num); // (a list goes on while list items follow each other)
    if (b.t === 'tbl') out += tblXml(b, ctx);
    else if (b.t === 'raw') out += b.xml;
    else { out += pXml(b, ctx, last ? sect : '', anch ? anch : ''); anch = ''; }
    if (!(b.t === 'p' && b.tiny)) prev = b;
  });
  return out;
}

/* ───────── styles: Normal and headings from the text's most common looks ───────── */
function stylesFrom(allBlocks) {
  const count = new Map(), add = (st, n) => { const k = [st.name, Math.max(2, Math.round(st.size * 2)), !!st.b, !!st.i, (st.color || '#000000').toLowerCase()].join('|'); count.set(k, (count.get(k) || 0) + n); };
  const dom = B => { const m = new Map(); for (const r of B.runs) if (r.st && r.t) { const k = [r.st.name, Math.max(2, Math.round(r.st.size * 2)), !!r.st.b, !!r.st.i, (r.st.color || '#000000').toLowerCase()].join('|'); m.set(k, (m.get(k) || 0) + r.t.length); } return [...m].sort((a, b) => b[1] - a[1])[0]; };
  const paras = []; const walk = bs => bs.forEach(b => { if (b.t === 'p' && b.runs.length) paras.push(b); else if (b.t === 'tbl') b.rows.forEach(r => r.cells.forEach(c => walk(c.blocks))); });
  allBlocks.forEach(walk);
  for (const B of paras) for (const r of B.runs) if (r.st && r.t && !r.st.ghost) add(r.st, r.t.length);
  if (!count.size) for (const B of paras) for (const r of B.runs) if (r.st && r.t) add(r.st, r.t.length); // (all see-through OCR text)
  const seen = B => B.runs.some(r => r.st && r.t && !r.st.ghost);
  const parse = k => { const [name, sz, b, i, color] = k.split('|'); return { name, sz: +sz, b: b === 'true', i: i === 'true', color }; };
  const top = [...count].sort((a, b) => b[1] - a[1])[0], body = top ? parse(top[0]) : { name: 'Calibri', sz: 22, b: false, i: false, color: '#000000' };
  const hs = new Map(); for (const B of paras) { if (B.hf || !seen(B)) continue; const d = dom(B); if (!d) continue; const s = parse(d[0]); if (s.sz >= body.sz * 1.15 && B.runs.filter(r => r.br).length < 3) { const e = hs.get(s.sz) || new Map(); e.set(d[0], (e.get(d[0]) || 0) + d[1]); hs.set(s.sz, e); } }
  const heads = [...hs.keys()].sort((a, b) => b - a).slice(0, 3).map((sz, i) => ({ id: 'Heading' + (i + 1), lvl: i, ...parse([...hs.get(sz)].sort((a, b) => b[1] - a[1])[0][0]) }));
  const normal = { id: 'Normal', ...body };
  const styleOf = B => { if (B.styleId) return B.styleId; const d = B.runs.length ? dom(B) : null; let s = normal; if (d && seen(B)) { const k = d[0], h = heads.find(h => [h.name, h.sz, h.b, h.i, h.color].join('|') === k); if (h) s = h; } B.styleId = s; return s; };
  const rpr = (s, def) => `<w:rPr>${def ? `<w:rFonts w:ascii="${xesc(s.name)}" w:hAnsi="${xesc(s.name)}" w:eastAsia="${xesc(s.name)}" w:cs="${xesc(s.name)}"/>` : s.name !== body.name ? `<w:rFonts w:ascii="${xesc(s.name)}" w:hAnsi="${xesc(s.name)}" w:eastAsia="${xesc(s.name)}" w:cs="${xesc(s.name)}"/>` : ''}${s.b ? '<w:b/><w:bCs/>' : ''}${s.i ? '<w:i/><w:iCs/>' : ''}${s.color !== '#000000' ? `<w:color w:val="${s.color.slice(1).toUpperCase()}"/>` : ''}<w:sz w:val="${s.sz}"/><w:szCs w:val="${s.sz}"/>${def ? '<w:lang w:val="en-US"/>' : ''}</w:rPr>`;
  const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
  const xml = `<w:styles ${W}><w:docDefaults><w:rPrDefault>${rpr({ ...body, b: false, i: false, color: '#000000' }, true)}</w:rPrDefault><w:pPrDefault><w:pPr><w:widowControl w:val="0"/><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>`
    + `<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/><w:pPr><w:widowControl w:val="0"/></w:pPr>${rpr({ ...body, name: body.name }, false)}</w:style>`
    + heads.map(h => `<w:style w:type="paragraph" w:styleId="${h.id}"><w:name w:val="heading ${h.lvl + 1}"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="9"/><w:qFormat/><w:pPr><w:keepNext w:val="0"/><w:keepLines w:val="0"/><w:widowControl w:val="0"/><w:spacing w:before="0" w:after="0"/><w:outlineLvl w:val="${h.lvl}"/></w:pPr>${rpr(h, false)}</w:style>`).join('')
    + '<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>'
    + '<w:style w:type="numbering" w:default="1" w:styleId="NoList"><w:name w:val="No List"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/></w:style></w:styles>';
  return { xml, styleOf, body };
}
function numbering() { // one Word list per run of list items; numbered lists start at the PDF's number
  const lists = [], abs = new Map(); let cur = null;
  const fmtOf = t => { const m = /^(\(?)(\d{1,3}|[a-zA-Z])([.)])$/.exec(t); if (!m) return null; const v = m[2]; return /\d/.test(v) ? ['decimal', +v, m[1] + '%1' + m[3]] : [/[a-z]/.test(v) ? 'lowerLetter' : 'upperLetter', v.toLowerCase().charCodeAt(0) - 96, m[1] + '%1' + m[3]]; };
  return {
    id(B, cont) {
      const mk = B.num, f = mk.kind === 'number' ? fmtOf(mk.t) : null, fmt = f ? f[0] : 'bullet', txt = f ? f[2] : mk.t; // (a private-use bullet keeps its Symbol / Wingdings font, as Word writes them)
      const ak = [fmt, txt, mk.st.name, Math.round(mk.st.size * 2), mk.st.color].join('|');
      if (!abs.has(ak)) abs.set(ak, { id: abs.size + 1, fmt, txt, st: mk.st });
      const a = abs.get(ak);
      if (!cur || cur.a !== a || !cont) { cur = { id: lists.length + 1, a, start: f ? f[1] : 1 }; lists.push(cur); }
      return cur.id;
    },
    xml() {
      if (!lists.length) return null;
      const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
      return `<w:numbering ${W}>${[...abs.values()].map(a => `<w:abstractNum w:abstractNumId="${a.id}"><w:multiLevelType w:val="singleLevel"/><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="${a.fmt}"/><w:lvlText w:val="${xesc(a.txt)}"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="360" w:hanging="360"/></w:pPr><w:rPr><w:rFonts w:ascii="${xesc(a.st.name)}" w:hAnsi="${xesc(a.st.name)}" w:hint="default"/>${a.st.b ? '<w:b/>' : ''}${a.st.color && a.st.color !== '#000000' ? `<w:color w:val="${a.st.color.slice(1).toUpperCase()}"/>` : ''}<w:sz w:val="${Math.max(2, Math.round(a.st.size * 2))}"/></w:rPr></w:lvl></w:abstractNum>`).join('')}${lists.map(l => `<w:num w:numId="${l.id}"><w:abstractNumId w:val="${l.a.id}"/>${l.a.fmt !== 'bullet' ? `<w:lvlOverride w:ilvl="0"><w:startOverride w:val="${l.start}"/></w:lvlOverride>` : ''}</w:num>`).join('')}</w:numbering>`;
    },
  };
}

/* ───────── the whole document ───────── */
async function build(pages) {
  await X.loadLib('jszip');
  const pk = makePackage(), out = [], fallback = [], why = [], t0 = performance.now();
  const ctx = { hf: await findHF(pages), maxPitch: 1.45, embeds: new Map() }, all = S().pages || [];
  for (let i = 0; i < pages.length; i++) { // pass 1: lay out each page; its drawing goes straight into its header part
    say(`Converting page ${i + 1} of ${pages.length}…`); await tick();
    const p = pages[i], [pw, ph] = X.rdims(p), no = all.indexOf(p) >= 0 ? all.indexOf(p) + 1 : i + 1; let L = null, reason = 'big';
    if (pw <= 1584 && ph <= 1584) {
      reason = 'failed';
      try {
        L = await layoutPage(p, ctx, i);
        if (p.rot % 360 && L.turnedShare > 0.6) { const L2 = await layoutPage({ ...p, rot: 0 }, ctx, i); if (!L2.lowConf) { L = L2; why.push({ page: no, why: 'upright' }); } } // (a page the user turned: its text stays upright, so it can flow)
        if (L.lowConf) reason = 'complex';
      } catch (e) { console.warn('Exact layout failed on page', no, e); L = null; }
    }
    if (!L || L.lowConf) {
      fallback.push(no);
      if (p.objs.some(o => o.cover && (o.redact || o.fill === '#000000'))) { // (blacked-out text must not come back: this page becomes a picture)
        const [w, h] = X.rdims(p), c = await X.renderPageImage(p, 150 / 72); why.push({ page: no, why: 'redacted' });
        out.push({ same: { w, h, margins: { t: 0, r: 0, b: 0, l: 0 }, paras: [{ ppr: '<w:spacing w:before="0" w:after="0" w:line="20" w:lineRule="exact"/>', body: pk.picXml('document', c.toDataURL('image/jpeg', 0.85), w, h, { x: 0, y: 0 }) }] } });
        continue;
      }
      why.push({ page: no, why: reason }); out.push({ same: await cv().samePage(p, { pic: (s, w, h, at) => pk.picXml('document', s, w, h, at), tbox: pk.tboxXml }) }); continue;
    }
    const rk = 'deco' + i, Z = 251658240; let deco = '', z = 0; // (Word's own drawing order numbers; small ones end up on top)
    const put = (g, x) => { if (g.host) g.host.B.anch = (g.host.B.anch || '') + x; else deco += x; }; // (drawing among the text hangs from its paragraph; the rest sits in the header)
    for (const r of L.G.rects.filter(r => !r.used && !r.over).sort((a, b) => b.w * b.h - a.w * a.h)) put(r, pk.shapeXml(r, Z + 1024 * z++, r.host)); // (bigger shapes behind smaller ones, pictures on top)
    for (const g of L.G.pics) put(g, await patchXml(pk, L.M.bg, g, g.host ? 'document' : rk, Z + 1024 * z++, g.host));
    for (const r of L.G.rects.filter(r => !r.used && r.over)) put(r, pk.shapeXml(r, Z + 1024 * z++, r.host)); // (a button taken out of a picture lies on it)
    let tb = ''; // turned text → small text boxes at their place
    if (L.turned.length) {
      const { linesOf: lo, blocksOf: bo, blockXml, hasFont } = cv(), groups = new Map();
      for (const q of L.turned) { const k = q.rot + '|' + q.pv.map(v => v.toFixed(1)).join(','); if (!groups.has(k)) groups.set(k, []); groups.get(k).push(q); }
      const place = (q, x, y) => { if (q.rot) { const a = q.rot * Math.PI / 180, dx = x - q.pv[0], dy = y - q.pv[1]; x = q.pv[0] + dx * Math.cos(a) - dy * Math.sin(a); y = q.pv[1] + dx * Math.sin(a) + dy * Math.cos(a); } return [x, y]; };
      const B = { pic: (s, w, h, at) => pk.picXml('document', s, w, h, at), tbox: pk.tboxXml };
      for (const g of groups.values()) for (const b of bo(lo(g.map(q => ({ ...q, exact: hasFont(q.font) }))))) tb += blockXml(b, B, place, 0, 1);
    }
    L.deco = deco; L.tb = tb; L.relKey = rk; L.M = null; L.G = null; L.turned = null;
    // a page that follows on from an exact page of the same size starts with a continuous break: text added on the page before then flows on
    // into this page (no near-empty page in between). The page before ends just under its last line (so nothing of this page fits there);
    // the space above this page's first line goes into its top margin, as Word drops space before at the top of a page.
    const P = out.length === i && out[i - 1] && out[i - 1].L;
    if (P && !deco && !tb && Math.abs(L.PW - P.PW) < 0.5 && Math.abs(L.PH - P.PH) < 0.5 && (!L.fb || L.fbRun === L.fb)) {
      let k = 0, h = 0; while (k < L.blocks.length && L.blocks[k].t === 'p' && L.blocks[k].tiny && !L.blocks[k].anch && !L.blocks[k].rule && !L.blocks[k].anchor) { h += L.blocks[k].before / 20 + (L.blocks[k].nLines || 1) * L.blocks[k].line / 20; k++; }
      const F = L.blocks[k], fh = b => (b ? (b.t === 'p' ? (b.anch || b.rule ? 0 : b.before / 20 + b.line / 20) : 0) : 99);
      const L1 = !F || (F.t === 'p' && (F.anch || F.tiny)) ? 0 : F.t === 'p' ? F.line / 20 : F.cols || F.box ? Math.min(...F.rows[0].cells.map(c => fh(c.blocks[0]))) : F.rows[0].h / 20;
      if (L1 >= 2 && k < L.blocks.length - 1) {
        L.MT += h + (F.t === 'p' ? F.before / 20 : 0); if (F.t === 'p') F.before = 0; L.blocks.splice(0, k); L.cont = true;
        const free = P.PH - P.MB - P.bodyEnd; if (free > 0) P.MB = P.PH - P.bodyEnd - Math.min(free, Math.max(0, L1 - 1.5));
      }
    }
    out.push({ L });
  }
  // pass 2: styles from the whole document, then the XML
  const st = stylesFrom(out.filter(o => o.L).map(o => o.L.blocks)), num = numbering();
  const xctx = { pages: pages.length, pageNo: 1, styleOf: st.styleOf, num };
  const parts = [], partOf = new Map(), anyFooter = out.some(o => o.L && o.L.fb);
  const part = (kind, xml, relKey) => { // header / footer parts; identical ones without pictures are shared
    const shared = !relKey || !pk.relsOf(relKey).length, k = kind + xml; if (shared && partOf.has(k)) return partOf.get(k);
    const n = parts.length + 1, rid = 'rIdH' + n; parts.push({ kind, name: `${kind}${n}.xml`, rid, xml, relKey }); if (shared) partOf.set(k, rid); return rid;
  };
  const EMPTY = '<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="20" w:lineRule="exact"/><w:rPr><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr></w:pPr></w:p>';
  let body = '';
  for (let i = 0; i < out.length; i++) {
    const o = out[i], last = i === out.length - 1; xctx.pageNo = i + 1;
    if (o.same) { // (every section names its own header, so a page never shows the previous page's drawing)
      const s = o.same, sect = sectXml(s.w, s.h, s.margins.t, s.margins.r, s.margins.b, s.margins.l, 0, 0, part('header', EMPTY), anyFooter ? part('footer', EMPTY) : null);
      s.paras.forEach((P, k) => { body += `<w:p><w:pPr>${P.ppr}${k === s.paras.length - 1 && !last ? sect : ''}</w:pPr>${P.body}</w:p>`; });
      if (last) body += sect; continue;
    }
    // the page's drawing sits in the section's first-page header; text that overflows onto a new page gets the running header only
    const L = o.L, hr1 = part('header', blocksXml(L.hb, xctx, '', L.deco), L.relKey), hr = L.deco ? part('header', blocksXml(L.hb, xctx, '', '')) : hr1;
    const fr1 = anyFooter ? part('footer', L.fb ? blocksXml(L.fb, xctx, '', '') : EMPTY) : null, fr = anyFooter && L.fbRun && L.fbRun !== L.fb ? part('footer', L.fbRun.length ? blocksXml(L.fbRun, xctx, '', '') : EMPTY) : fr1;
    const sect = sectXml(L.PW, L.PH, L.MT, L.MR, L.MB, L.ML, 0, L.fd, hr, fr, hr1, fr1, { type: L.cont ? 'continuous' : 'nextPage' });
    body += blocksXml(L.blocks, xctx, last ? '' : sect, L.tb);
    if (last) body += sect;
    out[i] = { done: true };
  }
  const fonts = await embedFonts(pk, [...ctx.embeds]), put = pk.put, numXml = num.xml();
  const docRels = [`<Relationship Id="rId1" Type="${RT}styles" Target="styles.xml"/>`, `<Relationship Id="rId2" Type="${RT}settings" Target="settings.xml"/>`, numXml ? `<Relationship Id="rId3" Type="${RT}numbering" Target="numbering.xml"/>` : '', fonts.length ? `<Relationship Id="rId4" Type="${RT}fontTable" Target="fontTable.xml"/>` : '']
    .concat(parts.map(p => `<Relationship Id="${p.rid}" Type="${RT}${p.kind}" Target="${p.name}"/>`), pk.relsOf('document').map(r => `<Relationship Id="${r.id}" Type="${RT}image" Target="${r.target}"/>`)).join('');
  const relsXml = R => `${XH}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${R}</Relationships>`;
  for (const p of parts) {
    const tag = p.kind === 'header' ? 'hdr' : 'ftr'; put(`word/${p.name}`, `${XH}<w:${tag} ${NS}>${p.xml}</w:${tag}>`);
    const R = p.relKey ? pk.relsOf(p.relKey) : []; if (R.length) put(`word/_rels/${p.name}.rels`, relsXml(R.map(r => `<Relationship Id="${r.id}" Type="${RT}image" Target="${r.target}"/>`).join('')));
  }
  const now = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
  put('[Content_Types].xml', `${XH}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Default Extension="jpeg" ContentType="image/jpeg"/>${fonts.length ? '<Default Extension="odttf" ContentType="application/vnd.openxmlformats-officedocument.obfuscatedFont"/>' : ''}<Override PartName="/word/document.xml" ContentType="${CT}.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="${CT}.styles+xml"/><Override PartName="/word/settings.xml" ContentType="${CT}.settings+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>${numXml ? `<Override PartName="/word/numbering.xml" ContentType="${CT}.numbering+xml"/>` : ''}${fonts.length ? `<Override PartName="/word/fontTable.xml" ContentType="${CT}.fontTable+xml"/>` : ''}${parts.map(p => `<Override PartName="/word/${p.name}" ContentType="${CT}.${p.kind}+xml"/>`).join('')}</Types>`);
  put('_rels/.rels', relsXml(`<Relationship Id="rId1" Type="${RT}officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>`));
  put('docProps/core.xml', `${XH}<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>${xesc(X.baseName())}</dc:title><dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified></cp:coreProperties>`);
  put('word/_rels/document.xml.rels', relsXml(docRels));
  put('word/styles.xml', XH + st.xml);
  if (numXml) put('word/numbering.xml', XH + numXml);
  if (fonts.length) { put('word/fontTable.xml', `${XH}<w:fonts ${NS}>${fonts.xml}</w:fonts>`); put('word/_rels/fontTable.xml.rels', relsXml(fonts.map(f => `<Relationship Id="${f.id}" Type="${RT}font" Target="${f.target}"/>`).join(''))); }
  put('word/settings.xml', `${XH}<w:settings ${NS}><w:zoom w:percent="100"/>${fonts.length ? '<w:embedTrueTypeFonts/><w:saveSubsetFonts w:val="0"/>' : ''}<w:defaultTabStop w:val="720"/><w:characterSpacingControl w:val="doNotCompress"/><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>`);
  put('word/document.xml', `${XH}<w:document ${NS}><w:body>${body}</w:body></w:document>`);
  say('Packing the Word file…'); await tick();
  const blob = await pk.zip.generateAsync({ type: 'blob', compression: 'DEFLATE', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
  blob.fallback = fallback; blob.fallbackWhy = why; blob.ms = Math.round(performance.now() - t0);
  return blob;
}
function sectXml(w, h, t, r, b, l, hd, fd, hr, fr, hr1, fr1, so = {}) { // hr1 / fr1: a first-page header / footer of its own (w:titlePg); so.type: how the section starts, so.cols: its columns
  const first = (hr1 && hr1 !== hr) || (fr1 && fr1 !== fr);
  return `<w:sectPr>${first && (hr1 || hr) ? `<w:headerReference w:type="first" r:id="${hr1 || hr}"/>` : ''}${hr ? `<w:headerReference w:type="default" r:id="${hr}"/>` : ''}${first && (fr1 || fr) ? `<w:footerReference w:type="first" r:id="${fr1 || fr}"/>` : ''}${fr ? `<w:footerReference w:type="default" r:id="${fr}"/>` : ''}<w:type w:val="${so.type || 'nextPage'}"/><w:pgSz w:w="${twip(w)}" w:h="${twip(h)}"${w > h ? ' w:orient="landscape"' : ''}/><w:pgMar w:top="${twip(t)}" w:right="${twip(r)}" w:bottom="${twip(b)}" w:left="${twip(l)}" w:header="${twip(hd)}" w:footer="${twip(fd)}" w:gutter="0"/>${so.cols || '<w:cols w:space="720"/>'}${first ? '<w:titlePg/>' : ''}</w:sectPr>`;
}
async function patchXml(pk, bg, g, part, z, rel) { // a piece of the page drawing as a picture behind the text
  const r = g.px, c = document.createElement('canvas'); c.width = Math.max(1, r.x1 - r.x0); c.height = Math.max(1, r.y1 - r.y0);
  const cx = c.getContext('2d'); cx.drawImage(bg, r.x0, r.y0, c.width, c.height, 0, 0, c.width, c.height);
  if (g.inpaint) { // shapes taken out of the picture: each row of their box blends from the pixel left of it to the one right of it
    const id = cx.getImageData(0, 0, c.width, c.height), D = id.data, W = c.width;
    for (const q of g.inpaint) {
      const a = clamp(q.x0 - r.x0 - 2, 0, W - 1), b = clamp(q.x1 - r.x0 + 1, 0, W - 1);
      for (let y = clamp(q.y0 - r.y0 - 1, 0, c.height - 1); y <= clamp(q.y1 - r.y0, 0, c.height - 1); y++) {
        const ia = (y * W + a) * 4, ib = (y * W + b) * 4;
        for (let x = a + 1; x < b; x++) { const t = (x - a) / (b - a || 1), i = (y * W + x) * 4; for (let k = 0; k < 4; k++) D[i + k] = Math.round(D[ia + k] * (1 - t) + D[ib + k] * t); }
      }
    }
    cx.putImageData(id, 0, 0);
  }
  let blob = await toBlob(c, 'image/png'), ext = 'png';
  if (blob.size > 200000) { const j = await toBlob(c, 'image/jpeg', 0.9); if (j && j.size < blob.size * 0.6) { blob = j; ext = 'jpeg'; } }
  const bytes = new Uint8Array(await blob.arrayBuffer());
  return pk.picXml(part, { bytes, ext, key: ext + fnv(bytes) }, g.w, g.h, { x: g.x, y: g.y, rel }, z);
}

X.pdf2word = { build, _fontOf: fontOf, _measure: measure, _model: pageModel, _rowsOf: rowsOf, _drawing: drawing,_layout: (p, ctx) => layoutPage(p, { hf: { head: new Set(), foot: new Set() }, maxPitch: 1.6, embeds: new Map(), ...ctx }, 0) };
})();
