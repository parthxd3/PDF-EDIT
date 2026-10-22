/* XD3 PDF Editor — native text engine. Reads a page's own drawing instructions (content stream) so that edited text can be
   removed from the page itself instead of being covered, and new text can be written with the document's own font resource:
   the saved page then has no cover boxes, no second copy of the font and no leftover text — it looks like it was never edited. */
(() => {
'use strict';
const X = window.XD3, { S } = X, L = PDFLib;
const N = n => L.PDFName.of(n);

/* ───────── tokenizer ───────── */
const WS = new Uint8Array(256); [0, 9, 10, 12, 13, 32].forEach(c => WS[c] = 1);
const DELIM = new Uint8Array(256); '()<>[]{}/%'.split('').forEach(c => DELIM[c.charCodeAt(0)] = 1);
function lex(b) { // → tokens { k: 'num'|'name'|'str'|'hex'|'op'|'['|']'|'<<'|'>>', v, s, e }
  const out = [], n = b.length; let i = 0;
  while (i < n) {
    const c = b[i];
    if (WS[c]) { i++; continue; }
    if (c === 37) { while (i < n && b[i] !== 10 && b[i] !== 13) i++; continue; } // % comment
    const s = i;
    if (c === 40) { // ( literal string )
      const bytes = []; let depth = 1; i++;
      while (i < n && depth) {
        let ch = b[i++];
        if (ch === 92) { // backslash escape
          const e2 = b[i++];
          if (e2 === 110) bytes.push(10); else if (e2 === 114) bytes.push(13); else if (e2 === 116) bytes.push(9); else if (e2 === 98) bytes.push(8); else if (e2 === 102) bytes.push(12);
          else if (e2 >= 48 && e2 <= 55) { let v = e2 - 48; for (let k = 0; k < 2 && b[i] >= 48 && b[i] <= 55; k++) v = v * 8 + b[i++] - 48; bytes.push(v & 255); }
          else if (e2 === 13) { if (b[i] === 10) i++; } else if (e2 === 10) { /* line continuation */ }
          else bytes.push(e2);
          continue;
        }
        if (ch === 40) depth++; else if (ch === 41 && !--depth) break;
        if (ch === 13) { ch = 10; if (b[i] === 10) i++; }
        bytes.push(ch);
      }
      out.push({ k: 'str', v: bytes, s, e: i }); continue;
    }
    if (c === 60 && b[i + 1] === 60) { out.push({ k: '<<', s, e: i += 2 }); continue; }
    if (c === 62 && b[i + 1] === 62) { out.push({ k: '>>', s, e: i += 2 }); continue; }
    if (c === 60) { // <hex>
      i++; let h = ''; while (i < n && b[i] !== 62) { if (!WS[b[i]]) h += String.fromCharCode(b[i]); i++; } i++;
      if (h.length % 2) h += '0'; const bytes = []; for (let k = 0; k < h.length; k += 2) bytes.push(parseInt(h.substr(k, 2), 16) || 0);
      out.push({ k: 'str', v: bytes, s, e: i, hex: 1 }); continue;
    }
    if (c === 91 || c === 93) { out.push({ k: c === 91 ? '[' : ']', s, e: ++i }); continue; }
    if (c === 123 || c === 125) { i++; continue; } // { } only in PostScript functions
    if (c === 47) { // /Name
      i++; let nm = ''; while (i < n && !WS[b[i]] && !DELIM[b[i]]) { if (b[i] === 35 && i + 2 < n) { nm += String.fromCharCode(parseInt(String.fromCharCode(b[i + 1], b[i + 2]), 16)); i += 3; } else nm += String.fromCharCode(b[i++]); }
      out.push({ k: 'name', v: nm, s, e: i }); continue;
    }
    let w = ''; while (i < n && !WS[b[i]] && !DELIM[b[i]]) w += String.fromCharCode(b[i++]);
    if (!w) { i++; continue; }
    if (/^[+-]?(\d+\.?\d*|\.\d+)$/.test(w)) out.push({ k: 'num', v: parseFloat(w), s, e: i });
    else if (w === 'true' || w === 'false' || w === 'null') out.push({ k: 'kw', v: w, s, e: i });
    else {
      out.push({ k: 'op', v: w, s, e: i });
      if (w === 'ID') { // inline image data: skip to EI
        i++; while (i < n) { if (b[i] === 69 && b[i + 1] === 73 && WS[b[i - 1]] && (i + 2 >= n || WS[b[i + 2]] || DELIM[b[i + 2]])) { i += 2; break; } i++; }
        out[out.length - 1].e = i;
      }
    }
  }
  return out;
}
function parseOps(b) { // → [{ op, args, s, e }] — args: numbers, { n: name }, byte arrays (strings, .hex), nested arrays, dict placeholders
  const tk = lex(b), ops = []; let args = [], start = -1;
  const stack = [];
  for (const t of tk) {
    if (start < 0) start = t.s;
    if (t.k === '[') { stack.push(args); args = []; continue; }
    if (t.k === ']') { const a = args; args = stack.pop() || []; args.push(a); continue; }
    if (t.k === '<<') { stack.push(args); args = []; continue; }
    if (t.k === '>>') { args = stack.pop() || []; args.push({ dict: 1 }); continue; }
    if (t.k === 'op' && !stack.length) { ops.push({ op: t.v, args, s: start, e: t.e }); args = []; start = -1; continue; }
    if (t.k === 'op') continue; // (an operator inside an array — malformed; ignored)
    args.push(t.k === 'num' ? t.v : t.k === 'name' ? { n: t.v } : t.k === 'str' ? Object.assign(t.v, { hex: t.hex }) : null);
  }
  return ops;
}

/* ───────── fonts: widths, code length, Unicode ⇄ code ───────── */
const AGL = { space: ' ', exclam: '!', quotedbl: '"', numbersign: '#', dollar: '$', percent: '%', ampersand: '&', quotesingle: "'", quoteright: '’', quoteleft: '‘', parenleft: '(', parenright: ')', asterisk: '*', plus: '+', comma: ',', hyphen: '-', minus: '−', period: '.', slash: '/', colon: ':', semicolon: ';', less: '<', equal: '=', greater: '>', question: '?', at: '@', bracketleft: '[', backslash: '\\', bracketright: ']', asciicircum: '^', underscore: '_', grave: '`', braceleft: '{', bar: '|', braceright: '}', asciitilde: '~', zero: '0', one: '1', two: '2', three: '3', four: '4', five: '5', six: '6', seven: '7', eight: '8', nine: '9', endash: '–', emdash: '—', bullet: '•', quotedblleft: '“', quotedblright: '”', ellipsis: '…', copyright: '©', registered: '®', degree: '°', trademark: '™', Euro: '€', sterling: '£', yen: '¥', section: '§', paragraph: '¶', periodcentered: '·', multiply: '×', divide: '÷', fi: 'ﬁ', fl: 'ﬂ', nbspace: ' ', rupee: '₹', rupeeindian: '₹' };
const glyphUni = nm => AGL[nm] || (nm.length === 1 ? nm : /^uni([0-9A-F]{4})$/i.test(nm) ? String.fromCharCode(parseInt(nm.slice(3), 16)) : /^u([0-9A-F]{4,6})$/i.test(nm) ? String.fromCodePoint(parseInt(nm.slice(1), 16)) : null);
const WINANSI_HI = { 128: '€', 130: '‚', 131: 'ƒ', 132: '„', 133: '…', 134: '†', 135: '‡', 136: 'ˆ', 137: '‰', 138: 'Š', 139: '‹', 140: 'Œ', 142: 'Ž', 145: '‘', 146: '’', 147: '“', 148: '”', 149: '•', 150: '–', 151: '—', 152: '˜', 153: '™', 154: 'š', 155: '›', 156: 'œ', 158: 'ž', 159: 'Ÿ' };
const num = o => (o instanceof L.PDFNumber ? o.asNumber() : null);
const arr = o => (o instanceof L.PDFArray ? o.asArray() : null);
function parseCMap(text) { // ToUnicode CMap → Map(code → string), plus the code length in bytes
  const map = new Map(); let bytes = 1; const hex2s = h => { let s = ''; for (let i = 0; i + 3 < h.length + 1; i += 4) s += String.fromCharCode(parseInt(h.substr(i, 4), 16)); return s; };
  for (const m of text.matchAll(/begincodespacerange([\s\S]*?)endcodespacerange/g)) { const h = /<([0-9a-fA-F]+)>/.exec(m[1]); if (h) bytes = Math.max(1, h[1].length >> 1); }
  for (const m of text.matchAll(/beginbfchar([\s\S]*?)endbfchar/g)) for (const p of m[1].matchAll(/<([0-9a-fA-F]+)>\s*<([0-9a-fA-F]*)>/g)) map.set(parseInt(p[1], 16), hex2s(p[2]));
  for (const m of text.matchAll(/beginbfrange([\s\S]*?)endbfrange/g)) {
    for (const p of m[1].matchAll(/<([0-9a-fA-F]+)>\s*<([0-9a-fA-F]+)>\s*(<([0-9a-fA-F]*)>|\[([^\]]*)\])/g)) {
      const lo = parseInt(p[1], 16), hi = Math.min(parseInt(p[2], 16), lo + 65535);
      if (p[4] != null) { const base = p[4], last = parseInt(base.slice(-4), 16); for (let c = lo; c <= hi; c++) map.set(c, hex2s(base.slice(0, -4) + (last + c - lo).toString(16).padStart(4, '0'))); }
      else { const list = [...p[5].matchAll(/<([0-9a-fA-F]*)>/g)].map(x => hex2s(x[1])); list.forEach((s, k) => map.set(lo + k, s)); }
    }
  }
  return { map, bytes };
}
const fontCache = new WeakMap();
function fontInfo(doc, fd) { // a font dictionary → { bytes, w0(code), uni(code), encode(ch), hex(codes), ok }
  if (!fd) return null; if (fontCache.has(fd)) return fontCache.get(fd);
  const ctx = doc.context, look = (d, k) => d && d.lookup(N(k)), sub = String(look(fd, 'Subtype') || '').replace('/', ''), info = { sub, ok: true, base: String(look(fd, 'BaseFont') || '').replace('/', '') };
  let toU = null; try { const tu = look(fd, 'ToUnicode'); if (tu instanceof L.PDFRawStream) toU = parseCMap(new TextDecoder('latin1').decode(L.decodePDFRawStream(tu).decode())); } catch (e) { /* no usable map */ }
  if (sub === 'Type0') {
    const enc = String(look(fd, 'Encoding') || ''), df = arr(look(fd, 'DescendantFonts')), cid = df && ctx.lookup(df[0]);
    info.bytes = 2; if (!/Identity-H|UCS2-H/.test(enc)) info.ok = false; // vertical or multi-byte CMaps are left alone
    const W = new Map(), dw = num(look(cid, 'DW')) ?? 1000, wa = arr(look(cid, 'W')) || [];
    for (let i = 0; i < wa.length;) {
      const a = num(ctx.lookup(wa[i])), b = ctx.lookup(wa[i + 1]);
      if (b instanceof L.PDFArray) { b.asArray().forEach((w, k) => W.set(a + k, num(ctx.lookup(w)))); i += 2; }
      else { const hi = num(b), w = num(ctx.lookup(wa[i + 2])); for (let c = a; c <= hi && c - a < 65536; c++) W.set(c, w); i += 3; }
    }
    info.w0 = c => (W.has(c) ? W.get(c) : dw) / 1000;
    info.uni = c => (toU && toU.map.get(c)) || null;
  } else {
    info.bytes = 1;
    const fc = num(look(fd, 'FirstChar')) ?? 0, wa = arr(look(fd, 'Widths')), fdesc = look(fd, 'FontDescriptor'), miss = num(look(fdesc, 'MissingWidth')) ?? 0;
    const fm = sub === 'Type3' ? (arr(look(fd, 'FontMatrix')) || []).map(v => num(ctx.lookup(v))) : null, scale = fm && fm[0] ? fm[0] : 0.001;
    let std = null; if (!wa) { try { const nm = Object.values(L.StandardFonts).find(v => v === info.base || v.replace(/-/g, '') === info.base.replace(/[-,]/g, '')); if (nm) std = L.StandardFontEmbedder.for(nm); } catch (e) { /* not a standard font */ } }
    info.w0 = c => { if (wa) { const w = c >= fc && c - fc < wa.length ? num(ctx.lookup(wa[c - fc])) : null; return (w ?? miss) * scale; } if (std) { try { return std.widthOfTextAtSize(String.fromCharCode(c), 1000) / 1000; } catch (e) { return 0.5; } } return 0.5; };
    const diff = {}, enc = look(fd, 'Encoding'); let baseEnc = String(enc || '');
    if (enc instanceof L.PDFDict) { baseEnc = String(enc.lookup(N('BaseEncoding')) || ''); const d = arr(enc.lookup(N('Differences'))) || []; let code = 0; for (const it of d) { const v = ctx.lookup(it); if (v instanceof L.PDFNumber) code = v.asNumber(); else if (v instanceof L.PDFName) diff[code++] = v.asString().replace('/', ''); } }
    info.uni = c => { if (toU && toU.map.has(c)) return toU.map.get(c); if (diff[c]) return glyphUni(diff[c]); if (c >= 32 && c < 127) return /Standard/.test(baseEnc) && c === 39 ? '’' : String.fromCharCode(c); if (/WinAnsi/.test(baseEnc) || !baseEnc) return WINANSI_HI[c] || (c >= 160 ? String.fromCharCode(c) : null); return null; };
    if (sub === 'Type3') info.ok = false; // Type 3 glyphs are drawings — removable, but new text is never written with them
  }
  const inv = new Map(), lig = new Map(); // Unicode → code, from every code the font can show; ligature glyphs (fi, fl, ffi …) by their letters
  if (info.ok) {
    const codes = toU ? [...toU.map.keys()] : Array.from({ length: 256 }, (_, i) => i);
    for (const c of codes) {
      const u = info.uni(c); if (!u || !(info.bytes === 2 || info.w0(c) > 0 || u === ' ')) continue; const n = [...u].length;
      if (n === 1) { if (!inv.has(u)) inv.set(u, c); } else if (n <= 3 && /^[a-zA-Z]+$/.test(u) && !lig.has(u)) lig.set(u, c);
      const dec = { 'ﬁ': 'fi', 'ﬂ': 'fl', 'ﬀ': 'ff', 'ﬃ': 'ffi', 'ﬄ': 'ffl' }[u]; if (dec && !lig.has(dec)) lig.set(dec, c);
    }
  }
  info.lig = lig;
  info.encode = ch => (inv.has(ch) ? inv.get(ch) : ch === ' ' && inv.has(' ') ? inv.get(' ') : null);
  info.hex = codes => L.PDFHexString.of(codes.map(c => c.toString(16).padStart(info.bytes * 2, '0')).join(''));
  fontCache.set(fd, info); return info;
}
function resources(node) { for (let n = node, k = 0; n && k < 32; n = n.lookup(N('Parent')), k++) { const r = n.lookup(N('Resources')); if (r instanceof L.PDFDict) return r; } return null; }
function fontByName(doc, node, name) { const r = resources(node), f = r && r.lookup(N('Font')), fd = f instanceof L.PDFDict ? f.lookup(N(name)) : null; return fd instanceof L.PDFDict ? fontInfo(doc, fd) : null; }

/* ───────── interpreter: where every glyph on the page is, and which instruction drew it ───────── */
const mul = (a, b) => [a[0] * b[0] + a[1] * b[2], a[0] * b[1] + a[1] * b[3], a[2] * b[0] + a[3] * b[2], a[2] * b[1] + a[3] * b[3], a[4] * b[0] + a[5] * b[2] + b[4], a[4] * b[1] + a[5] * b[3] + b[5]];
const ap = (m, x, y) => [x * m[0] + y * m[2] + m[4], x * m[1] + y * m[3] + m[5]];
function contentBytes(node) {
  const c = node.lookup(N('Contents')), list = c instanceof L.PDFArray ? c.asArray().map(r => node.context.lookup(r)) : c ? [c] : [], parts = [];
  for (const s of list) { if (!(s instanceof L.PDFStream)) continue; try { parts.push(s instanceof L.PDFRawStream ? L.decodePDFRawStream(s).decode() : s.getContents()); } catch (e) { return null; } } // a stream we cannot decode → the page is left as it is
  const total = parts.reduce((a, p) => a + p.length + 1, 0), out = new Uint8Array(total); let o = 0; for (const p of parts) { out.set(p, o); o += p.length; out[o++] = 10; }
  return out;
}
function interpret(doc, node, vt) { // → { bytes, ops, glyphs: [{ cx, cy, oi, gi, ... }] } with glyph centres in view space
  const bytes = contentBytes(node); if (!bytes) return null;
  const ops = parseOps(bytes), glyphs = [], stack = [], after = []; let inBT = false;
  let g = { ctm: [1, 0, 0, 1, 0, 0], font: null, res: '', Tfs: 0, Tc: 0, Tw: 0, Th: 1, TL: 0, Ts: 0, Tr: 0, fill: [{ op: 'g', args: [0] }], cs: 'DeviceGray' };
  let Tm = [1, 0, 0, 1, 0, 0], Tlm = Tm;
  const nl = () => { Tlm = mul([1, 0, 0, 1, 0, -g.TL], Tlm); Tm = Tlm; };
  const show = (oi, items) => {
    const F = g.font; let gi = 0;
    for (const it of items) {
      if (typeof it === 'number') { Tm = mul([1, 0, 0, 1, -it / 1000 * g.Tfs * g.Th, 0], Tm); continue; }
      if (!Array.isArray(it) || !F) continue;
      for (let k = 0; k + F.bytes <= it.length; k += F.bytes) {
        const code = F.bytes === 2 ? it[k] << 8 | it[k + 1] : it[k], w0 = F.w0(code), Trm = mul([g.Tfs * g.Th, 0, 0, g.Tfs, 0, g.Ts], mul(Tm, g.ctm)), V = mul(Trm, vt);
        const [cx, cy] = ap(V, w0 / 2, 0.3), q = [ap(V, 0, -0.2), ap(V, w0, -0.2), ap(V, 0, 0.8), ap(V, w0, 0.8)], bx = { x: Math.min(...q.map(v => v[0])), y: Math.min(...q.map(v => v[1])) }, adv = (w0 * g.Tfs + g.Tc + (F.bytes === 1 && code === 32 ? g.Tw : 0)) * g.Th;
        bx.w = Math.max(...q.map(v => v[0])) - bx.x; bx.h = Math.max(...q.map(v => v[1])) - bx.y;
        const [ox, oy] = ap(V, 0, 0), [ex] = ap(V, adv / ((g.Tfs * g.Th) || 1), 0);
        glyphs.push({ cx, cy, bx, ox, oy, ex, oi, gi: gi++, code, adv, w0, res: g.res, F, Tfs: g.Tfs, Tc: g.Tc, Tw: g.Tw, Th: g.Th, Tr: g.Tr, size: Math.hypot(V[2], V[3]), fill: g.fill, cs: g.cs, uni: F.uni(code) });
        Tm = mul([1, 0, 0, 1, adv, 0], Tm);
      }
    }
  };
  ops.forEach((o, oi) => {
    const a = o.args;
    switch (o.op) {
      case 'q': stack.push({ ...g }); break;
      case 'Q': if (stack.length) g = stack.pop(); break;
      case 'cm': if (a.length === 6) g.ctm = mul(a, g.ctm); break;
      case 'BT': Tm = Tlm = [1, 0, 0, 1, 0, 0]; inBT = true; break;
      case 'ET': inBT = false; break;
      case 'Tf': g.res = a[0] && a[0].n || ''; g.Tfs = a[1] || 0; g.font = fontByName(doc, node, g.res); break;
      case 'Tc': g.Tc = a[0] || 0; break; case 'Tw': g.Tw = a[0] || 0; break; case 'Tz': g.Th = (a[0] ?? 100) / 100; break;
      case 'TL': g.TL = a[0] || 0; break; case 'Ts': g.Ts = a[0] || 0; break; case 'Tr': g.Tr = a[0] || 0; break;
      case 'Td': Tlm = mul([1, 0, 0, 1, a[0] || 0, a[1] || 0], Tlm); Tm = Tlm; break;
      case 'TD': g.TL = -(a[1] || 0); Tlm = mul([1, 0, 0, 1, a[0] || 0, a[1] || 0], Tlm); Tm = Tlm; break;
      case 'Tm': if (a.length === 6) Tm = Tlm = a.slice(); break;
      case 'T*': nl(); break;
      case 'Tj': show(oi, [a[0]]); break;
      case 'TJ': show(oi, a[0] || []); break;
      case "'": nl(); show(oi, [a[0]]); break;
      case '"': g.Tw = a[0] || 0; g.Tc = a[1] || 0; nl(); show(oi, [a[2]]); break;
      case 'g': case 'rg': case 'k': g.fill = [{ op: o.op, args: a.slice() }]; g.cs = { g: 'DeviceGray', rg: 'DeviceRGB', k: 'DeviceCMYK' }[o.op]; break;
      case 'cs': g.cs = a[0] && a[0].n || ''; g.fill = [{ op: 'cs', args: a.slice() }]; break;
      case 'sc': case 'scn': g.fill = [...g.fill.filter(f => f.op === 'cs'), { op: o.op, args: a.slice() }]; break;
    }
    if (o.op === 'Tj' || o.op === 'TJ' || o.op === "'" || o.op === '"') after[oi] = { ctm: g.ctm, Tm, Tlm, Tfs: g.Tfs, Th: g.Th, inBT };
  });
  return { bytes, ops, glyphs, after };
}
const inside = (gl, r) => { // glyph centre inside a (possibly rotated) view-space box { x, y, w, h, rot }
  let { cx, cy } = gl; if (r.rot) { const c = [r.x + r.w / 2, r.y + r.h / 2], a = -r.rot * Math.PI / 180, dx = cx - c[0], dy = cy - c[1]; cx = c[0] + dx * Math.cos(a) - dy * Math.sin(a); cy = c[1] + dx * Math.sin(a) + dy * Math.cos(a); }
  return cx >= r.x && cx <= r.x + r.w && cy >= r.y && cy <= r.y + r.h;
};
function rgbOf(gl) { // the fill colour of a glyph as #rrggbb, when it is in a device colour space
  const last = gl.fill[gl.fill.length - 1]; if (!last) return null; const v = last.args.filter(x => typeof x === 'number'), h = x => Math.round(Math.max(0, Math.min(1, x)) * 255).toString(16).padStart(2, '0');
  const cs = last.op === 'g' ? 'DeviceGray' : last.op === 'rg' ? 'DeviceRGB' : last.op === 'k' ? 'DeviceCMYK' : gl.cs;
  if (cs === 'DeviceGray' && v.length === 1) return '#' + h(v[0]).repeat(3);
  if (cs === 'DeviceRGB' && v.length === 3) return '#' + v.map(h).join('');
  if (cs === 'DeviceCMYK' && v.length === 4) { const [c, m, y, k] = v; return '#' + [c, m, y].map(q => h((1 - q) * (1 - k))).join(''); }
  return null;
}

/* ───────── at edit time: what exactly is in an edited line ───────── */
const pageCache = new WeakMap(); // source → Map(page index → Promise<analysis>)
async function srcDoc(si) { const s = S.sources[si]; return s.lib || (s.lib = L.PDFDocument.load(s.bytes, { updateMetadata: false, ignoreEncryption: true })); }
function analyze(p) {
  if (p.src < 0) return Promise.resolve(null); const s = S.sources[p.src]; if (!pageCache.has(s)) pageCache.set(s, new Map()); const cache = pageCache.get(s), key = p.idx;
  if (!cache.has(key)) cache.set(key, (async () => {
    try { const doc = await srcDoc(p.src), pg = await X.getPdfPage(p); return interpret(doc, doc.getPage(p.idx).node, pg.getViewport({ scale: 1 }).transform); }
    catch (e) { console.warn('native analysis', e); return null; }
  })());
  return cache.get(key);
}
async function lineInfo(p, region, str) { // glyphs of the original page inside a view-space box → the dominant run's exact style, word gap and kerning
  const A = await analyze(p); if (!A) return null;
  const gs = A.glyphs.filter(gl => inside(gl, region)); if (!gs.length) return { count: 0 };
  const by = new Map(); for (const gl of gs) { const k = gl.res + '|' + gl.Tfs.toFixed(3); by.set(k, (by.get(k) || []).concat(gl)); }
  const top = [...by.values()].sort((a, b) => b.length - a.length)[0], g0 = top[0];
  const vis = gs.filter(gl => gl.Tr !== 3 && gl.Tr !== 7).sort((a, b) => a.ox - b.ox), kern = {}; let gap = null;
  if (vis.length > 1 && !region.rot) {
    const isSp = gl => gl.uni === ' ' || gl.uni === '\u00a0', gaps = []; let prev = null, spaced = false;
    for (const gl of vis) { // word gap = from the end of one word's last letter to the start of the next word's first letter
      if (isSp(gl)) { spaced = true; continue; }
      if (prev) { const d = gl.ox - prev.ex, jump = d > 0.15 * (prev.size || 1); if (spaced || jump) gaps.push(d); else if (prev.uni && gl.uni) { const kd = d / (prev.size || 1), key = prev.uni + gl.uni; if (Math.abs(kd) > 0.002 && Math.abs(kd) < 0.3 && !(key in kern)) kern[key] = +kd.toFixed(4); } }
      prev = gl; spaced = false;
    }
    if (gaps.length) gap = gaps.reduce((a, v) => a + v, 0) / gaps.length;
  }
  return { count: gs.length, visible: gs.some(gl => gl.Tr !== 3 && gl.Tr !== 7), gap, kern, res: g0.res, Tfs: g0.Tfs, size: +g0.size.toFixed(3), Tc: g0.Tc / (g0.Tfs || 1) * g0.size, Tw: g0.Tw / (g0.Tfs || 1) * g0.size, Tz: g0.Th * 100, color: rgbOf(g0), fill: g0.fill, canWrite: !!(g0.F && g0.F.ok && g0.Tr === 0) };
}

/* ───────── at save time: remove glyphs from a page, and write text with a page's own font ───────── */
const inv6 = m => { const det = m[0] * m[3] - m[1] * m[2] || 1; return [m[3] / det, -m[1] / det, -m[2] / det, m[0] / det, (m[2] * m[5] - m[3] * m[4]) / det, (m[1] * m[4] - m[0] * m[5]) / det]; };
const fmt = v => { const s = (Math.abs(v) < 1e-12 ? 0 : v).toFixed(10).replace(/\.?0+$/, ''); return s === '-0' || s === '' ? '0' : s; };
function strip(doc, page, vt, regions, opt = {}) { // regions: [{ id, r }] → Map(id → glyphs removed) (+ .placed); the page gets a rewritten content stream.
  // opt.inject: [{ id, regions: [region ids], ops: 'content operators' }] — replacement text written right after the original glyphs,
  // so it stays inside the same clip, drawing order and tag as the line it replaces; opt.inv maps view space to page space.
  const node = page.node, A = interpret(doc, node, vt), removed = new Map(regions.map(x => [x.id, 0])), lastOp = new Map(); removed.placed = new Set(); if (!A) return removed;
  const kill = new Map(); // op index → Set(glyph index)
  const touches = (gl, x) => x.touch ? gl.bx.x < x.r.x + x.r.w && gl.bx.x + gl.bx.w > x.r.x && gl.bx.y < x.r.y + x.r.h && gl.bx.y + gl.bx.h > x.r.y : inside(gl, x.r);
  for (const gl of A.glyphs) { const hits = regions.filter(x => touches(gl, x)); if (!hits.length) continue; for (const h of hits) { removed.set(h.id, removed.get(h.id) + 1); lastOp.set(h.id, Math.max(lastOp.get(h.id) ?? -1, gl.oi)); } if (!kill.has(gl.oi)) kill.set(gl.oi, new Set()); kill.get(gl.oi).add(gl.gi); }
  if (!kill.size) return removed;
  const byOp = new Map(); for (const gl of A.glyphs) if (kill.has(gl.oi)) (byOp.get(gl.oi) || byOp.set(gl.oi, []).get(gl.oi)).push(gl);
  const enc = new TextEncoder(), pieces = []; let pos = 0, injectAt = new Map();
  for (const inj of opt.inject || []) { const at = Math.max(-1, ...inj.regions.map(id => lastOp.get(id) ?? -1)); if (at >= 0 && kill.has(at)) (injectAt.get(at) || injectAt.set(at, []).get(at)).push(inj); }
  const hexOf = (codes, F) => '<' + codes.map(c => c.toString(16).padStart(F.bytes * 2, '0')).join('') + '>';
  for (const [oi, dead] of [...kill.entries()].sort((a, b) => a[0] - b[0])) {
    const o = A.ops[oi], gls = byOp.get(oi), F = gls[0].F, parts = []; let codes = [], nAdj = 0, gi = 0;
    const flushCodes = () => { if (codes.length) { if (nAdj) parts.push(+nAdj.toFixed(3)); nAdj = 0; parts.push(hexOf(codes, F)); codes = []; } };
    const items = o.op === 'TJ' ? o.args[0] || [] : [o.args[o.op === '"' ? 2 : 0]];
    for (const it of items) {
      if (typeof it === 'number') { flushCodes(); nAdj += it; continue; }
      if (!Array.isArray(it)) continue;
      for (let k = 0; k + F.bytes <= it.length; k += F.bytes) {
        const gl = gls[gi], code = F.bytes === 2 ? it[k] << 8 | it[k + 1] : it[k];
        if (dead.has(gi)) { flushCodes(); nAdj -= gl.adv * 1000 / ((gl.Tfs || 1) * (gl.Th || 1)); } else codes.push(code);
        gi++;
      }
    }
    flushCodes(); if (nAdj) parts.push(+nAdj.toFixed(3));
    const pre = o.op === "'" ? 'T* ' : o.op === '"' ? `${o.args[0] || 0} Tw ${o.args[1] || 0} Tc T* ` : '';
    let extra = '';
    for (const inj of injectAt.get(oi) || []) { // ET · q (view space → here) cm · new text · Q · BT with the text position restored for what follows
      const st = A.after[oi], X = mul(opt.inv, inv6(st.ctm)); let tail = '';
      if (st.inBT) { const T = st.Tm, Tl = st.Tlm, dx = Math.abs(Tl[0]) > 1e-9 ? (T[4] - Tl[4]) / Tl[0] : Math.abs(Tl[1]) > 1e-9 ? (T[5] - Tl[5]) / Tl[1] : 0; tail = ` BT ${Tl.map(fmt).join(' ')} Tm` + (Math.abs(dx) > 1e-6 && st.Tfs ? ` [${fmt(-dx * 1000 / (st.Tfs * st.Th))}] TJ` : ''); }
      extra += `\n${st.inBT ? 'ET ' : ''}q ${X.map(fmt).join(' ')} cm\n${inj.ops}\nQ${tail}\n`; removed.placed.add(inj.id);
    }
    pieces.push(A.bytes.subarray(pos, o.s), enc.encode(pre + '[' + parts.join(' ') + '] TJ' + extra)); pos = o.e;
  }
  pieces.push(A.bytes.subarray(pos));
  const total = pieces.reduce((a, b) => a + b.length, 0), out = new Uint8Array(total); let k = 0; for (const pc of pieces) { out.set(pc, k); k += pc.length; }
  node.set(N('Contents'), doc.context.obj([doc.context.register(doc.context.flateStream(out))])); // (an array, as pdf-lib expects when it adds drawing later)
  return removed;
}
function fontFor(doc, page, res) { return fontByName(doc, page.node, res); }
const fillOps = fill => fill.map(f => L.PDFOperator.of(f.op, f.args.map(a => typeof a === 'number' ? L.PDFNumber.of(a) : a && a.n ? N(a.n) : L.PDFNumber.of(0))));

/* ───────── tidy-up: nothing of a removed page may stay in the file ───────── */
function scrubDead(doc) { // links, named destinations, form fields and structure tags that still point at pages not in the document
  const ctx = doc.context, live = new Set(doc.getPages().map(p => p.ref.tag)), TYPE = N('Type'), PAGE = N('Page');
  const dead = v => { if (!(v instanceof L.PDFRef) || live.has(v.tag)) return false; const o = ctx.lookup(v); return o instanceof L.PDFDict && o.get(TYPE) === PAGE; };
  const destOf = a => { let d = a.lookup(N('Dest')); const act = a.lookup(N('A')); if (!d && act instanceof L.PDFDict) d = act.lookup(N('D')); return d; };
  for (const pg of doc.getPages()) { // a link to a removed page goes with it
    const an = pg.node.lookup(N('Annots')); if (!(an instanceof L.PDFArray)) continue;
    const keep = an.asArray().filter(r => { const a = ctx.lookup(r); if (!(a instanceof L.PDFDict) || a.get(N('Subtype')) !== N('Link')) return true; const d = destOf(a); return !(d instanceof L.PDFArray && d.size() && dead(d.get(0))); });
    if (keep.length !== an.size()) pg.node.set(N('Annots'), ctx.obj(keep));
  }
  const af = doc.catalog.lookup(N('AcroForm'));
  if (af instanceof L.PDFDict) { // form fields that only lived on removed pages
    const prune = list => { const out = []; for (const r of list.asArray()) { const f = ctx.lookup(r); if (!(f instanceof L.PDFDict)) continue; const kids = f.lookup(N('Kids')); if (kids instanceof L.PDFArray) { const k = prune(kids); if (!k.length) continue; f.set(N('Kids'), ctx.obj(k)); out.push(r); continue; } if (dead(f.get(N('P')))) continue; out.push(r); } return out; };
    const fs = af.lookup(N('Fields')); if (fs instanceof L.PDFArray) af.set(N('Fields'), ctx.obj(prune(fs)));
  }
  const seen = new Set(), todo = [doc.catalog]; // every other pointer (named destinations, tags, …) is cleared
  const push = v => { if (v instanceof L.PDFRef) { if (!seen.has(v.tag)) { seen.add(v.tag); todo.push(ctx.lookup(v)); } } else if (v instanceof L.PDFDict || v instanceof L.PDFArray || v instanceof L.PDFStream) todo.push(v); };
  while (todo.length) {
    const o = todo.pop();
    if (o instanceof L.PDFDict) { for (const [k, v] of o.entries()) { if (dead(v)) o.delete(k); else push(v); } }
    else if (o instanceof L.PDFArray) { for (let i = 0; i < o.size(); i++) { const v = o.get(i); if (dead(v)) o.set(i, L.PDFNull); else push(v); } }
    else if (o instanceof L.PDFStream) todo.push(o.dict);
  }
}
/* objects nothing points at any more (removed pages, replaced streams) are dropped */
function collectGarbage(doc) {
  const ctx = doc.context, seen = new Set(), todo = [ctx.trailerInfo.Root, ctx.trailerInfo.Info, ctx.trailerInfo.Encrypt].filter(Boolean);
  while (todo.length) {
    const o = todo.pop(); if (!o) continue;
    if (o instanceof L.PDFRef) { if (seen.has(o.tag)) continue; seen.add(o.tag); todo.push(ctx.lookup(o)); }
    else if (o instanceof L.PDFDict) for (const [, v] of o.entries()) todo.push(v);
    else if (o instanceof L.PDFArray) for (let i = 0; i < o.size(); i++) todo.push(o.get(i));
    else if (o instanceof L.PDFStream) todo.push(o.dict);
  }
  for (const [ref] of ctx.enumerateIndirectObjects()) if (!seen.has(ref.tag)) ctx.delete(ref);
  return seen;
}

X.native = { lex, parseOps, interpret, analyze, lineInfo, strip, fontFor, fillOps, scrubDead, collectGarbage, srcDoc, inside, rgbOf };
})();
