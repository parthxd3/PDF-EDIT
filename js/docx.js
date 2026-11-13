/* XD3 PDF Editor — Word (.docx) editor. Opens a Word file as an editable page in its own fonts, styles, lists, tables and
   pictures, and saves it back as .docx rewriting only what changed: headers and footers you do not touch, comments, footnotes,
   fields, sections, themes and settings are kept exactly. Auto Format and Fix Formatting tidy a document in one go; spelling,
   grammar (Harper), humanize and statistics come from the writing tools. Everything runs on the device. */
(() => {
'use strict';
const X = window.XD3, { $, $$, esc, toast, busy, dialog, download, icons, clamp } = X;
const store = { get: k => { try { return localStorage.getItem(k); } catch (e) { return null; } }, set: (k, v) => { try { localStorage.setItem(k, v); } catch (e) { /* not remembered */ } } };
const NS = { w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main', r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships', xml: 'http://www.w3.org/XML/1998/namespace',
  rel: 'http://schemas.openxmlformats.org/package/2006/relationships', ct: 'http://schemas.openxmlformats.org/package/2006/content-types' };
const RT = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/', DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';

/* ───────── XML helpers ───────── */
const kids = (el, n) => el ? [...el.children].filter(c => c.localName === n) : [];
const kid = (el, n) => { if (el) for (const c of el.children) if (c.localName === n) return c; return null; };
const wa = (el, n) => el ? (el.getAttributeNS(NS.w, n) ?? el.getAttribute('w:' + n)) : null;
const onoff = el => { const v = wa(el, 'val'); return !(v === '0' || v === 'false' || v === 'off'); };
const parseXml = s => { const d = new DOMParser().parseFromString(s, 'application/xml'); if (d.getElementsByTagName('parsererror').length) throw new Error('This Word file has a damaged part'); return d; };
const ser = d => { const s = new XMLSerializer().serializeToString(d); return /^<\?xml/.test(s) ? s : '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' + s; };
const wel = (doc, name, attrs) => { const e = doc.createElementNS(NS.w, 'w:' + name); if (attrs) for (const k in attrs) if (attrs[k] != null) e.setAttributeNS(NS.w, 'w:' + k, String(attrs[k])); return e; };
const canon = v => v && typeof v === 'object' ? JSON.stringify(Object.keys(v).sort().reduce((o, k) => { if (v[k] !== undefined) o[k] = v[k]; return o; }, {})) : JSON.stringify(v === undefined ? null : v);
const same = (a, b) => canon(a) === canon(b);
function insertOrdered(parent, e, order) { // schema order matters to Word
  const i = order.indexOf(e.localName);
  for (const c of parent.children) { const j = order.indexOf(c.localName); if (j > i || (j < 0 && c.namespaceURI !== NS.w)) { parent.insertBefore(e, c); return; } }
  parent.appendChild(e);
}
const NSDECL = `xmlns:w="${NS.w}" xmlns:r="${NS.r}" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"`;
const frag = (doc, xml) => doc.importNode(parseXml(`<x ${NSDECL}>${xml}</x>`).documentElement.firstElementChild, true); // an XML snippet → an element of `doc`
const XML_BAD = /[\x00-\x08\x0b\x0c\x0e-\x1f\ufffe\uffff\u200b]/g;

/* ───────── formatting model: the properties the editor understands; everything else in a run or paragraph is kept as it was ───────── */
const R_ORDER = ['rStyle', 'rFonts', 'b', 'bCs', 'i', 'iCs', 'caps', 'smallCaps', 'strike', 'dstrike', 'outline', 'shadow', 'emboss', 'imprint', 'noProof', 'snapToGrid', 'vanish', 'webHidden', 'color', 'spacing', 'w', 'kern', 'position', 'sz', 'szCs', 'highlight', 'u', 'effect', 'bdr', 'shd', 'fitText', 'vertAlign', 'rtl', 'cs', 'em', 'lang', 'eastAsianLayout', 'specVanish', 'oMath', 'rPrChange'];
const P_ORDER = ['pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr', 'widowControl', 'numPr', 'suppressLineNumbers', 'pBdr', 'shd', 'tabs', 'suppressAutoHyphens', 'kinsoku', 'wordWrap', 'overflowPunct', 'topLinePunct', 'autoSpaceDE', 'autoSpaceDN', 'bidi', 'adjustRightInd', 'snapToGrid', 'spacing', 'ind', 'contextualSpacing', 'mirrorIndents', 'suppressOverlap', 'jc', 'textDirection', 'textAlignment', 'textboxTightWrap', 'outlineLvl', 'divId', 'cnfStyle', 'rPr', 'sectPr', 'pPrChange'];
const TC_ORDER = ['cnfStyle', 'tcW', 'gridSpan', 'hMerge', 'vMerge', 'tcBorders', 'shd', 'noWrap', 'tcMar', 'textDirection', 'tcFitText', 'vAlign', 'hideMark'];
const R_EL = { rStyle: ['rStyle'], font: ['rFonts'], b: ['b', 'bCs'], i: ['i', 'iCs'], caps: ['caps'], smallCaps: ['smallCaps'], strike: ['strike'], dstrike: ['dstrike'], vanish: ['vanish'], color: ['color'], sz: ['sz', 'szCs'], hl: ['highlight'], u: ['u'], shd: ['shd'], va: ['vertAlign'], spc: ['spacing'], pos: ['position'], scale: ['w'] };
const P_EL = { style: ['pStyle'], keepNext: ['keepNext'], pbb: ['pageBreakBefore'], num: ['numPr'], bdr: ['pBdr'], shd: ['shd'], sp: ['spacing'], ind: ['ind'], ctx: ['contextualSpacing'], jc: ['jc'] };
const R_KEYS = Object.keys(R_EL), P_KEYS = Object.keys(P_EL);
function readR(rPr) {
  const o = {}; if (!rPr) return o; let e;
  if ((e = kid(rPr, 'rStyle'))) o.rStyle = wa(e, 'val');
  if ((e = kid(rPr, 'rFonts'))) { const f = wa(e, 'ascii') || wa(e, 'hAnsi'), t = wa(e, 'asciiTheme') || wa(e, 'hAnsiTheme'); if (f) o.font = f; else if (t) o.font = /^major/.test(t) ? '@major' : '@minor'; }
  for (const k of ['b', 'i', 'caps', 'smallCaps', 'strike', 'dstrike', 'vanish']) if ((e = kid(rPr, k))) o[k] = onoff(e);
  if ((e = kid(rPr, 'color'))) { const v = wa(e, 'val'); if (v) o.color = v.toLowerCase() === 'auto' ? 'auto' : v.toUpperCase(); }
  if ((e = kid(rPr, 'sz')) && +wa(e, 'val')) o.sz = +wa(e, 'val') / 2;
  if ((e = kid(rPr, 'highlight'))) o.hl = wa(e, 'val');
  if ((e = kid(rPr, 'u'))) o.u = wa(e, 'val') || 'single';
  if ((e = kid(rPr, 'shd'))) { const f = wa(e, 'fill'); if (f && f.toLowerCase() !== 'auto') o.shd = f.toUpperCase(); }
  if ((e = kid(rPr, 'vertAlign'))) o.va = wa(e, 'val');
  if ((e = kid(rPr, 'spacing')) && +wa(e, 'val')) o.spc = +wa(e, 'val'); // letter spacing, twentieths of a point
  if ((e = kid(rPr, 'position')) && +wa(e, 'val')) o.pos = +wa(e, 'val'); // raised / lowered, half-points
  if ((e = kid(rPr, 'w')) && +wa(e, 'val') && +wa(e, 'val') !== 100) o.scale = +wa(e, 'val');
  return o;
}
function writeRProp(rPr, k, v) {
  for (const n of R_EL[k]) for (let e; (e = kid(rPr, n));) e.remove();
  if (v == null) return; const d = rPr.ownerDocument, els = [];
  if (k === 'rStyle') els.push(wel(d, 'rStyle', { val: v }));
  else if (k === 'font') els.push(wel(d, 'rFonts', v[0] === '@' ? { asciiTheme: v === '@major' ? 'majorHAnsi' : 'minorHAnsi', hAnsiTheme: v === '@major' ? 'majorHAnsi' : 'minorHAnsi' } : { ascii: v, hAnsi: v, cs: v }));
  else if (k === 'b' || k === 'i') els.push(wel(d, k, v ? null : { val: '0' }), wel(d, k + 'Cs', v ? null : { val: '0' }));
  else if (k === 'sz') { const h = Math.round(v * 2); els.push(wel(d, 'sz', { val: h }), wel(d, 'szCs', { val: h })); }
  else if (k === 'color') els.push(wel(d, 'color', { val: v }));
  else if (k === 'hl') els.push(wel(d, 'highlight', { val: v }));
  else if (k === 'u') els.push(wel(d, 'u', { val: v }));
  else if (k === 'shd') els.push(wel(d, 'shd', { val: 'clear', color: 'auto', fill: v }));
  else if (k === 'va') els.push(wel(d, 'vertAlign', { val: v }));
  else if (k === 'spc') els.push(wel(d, 'spacing', { val: Math.round(v) }));
  else if (k === 'pos') els.push(wel(d, 'position', { val: Math.round(v) }));
  else if (k === 'scale') els.push(wel(d, 'w', { val: Math.round(v) }));
  else els.push(wel(d, R_EL[k][0], v ? null : { val: '0' }));
  for (const e of els) insertOrdered(rPr, e, R_ORDER);
}
function readP(pPr) {
  const o = {}; if (!pPr) return o; let e;
  if ((e = kid(pPr, 'pStyle'))) o.style = wa(e, 'val');
  if ((e = kid(pPr, 'keepNext'))) o.keepNext = onoff(e);
  if ((e = kid(pPr, 'pageBreakBefore'))) o.pbb = onoff(e);
  if ((e = kid(pPr, 'numPr'))) { const id = wa(kid(e, 'numId'), 'val'); if (id != null) o.num = { id, lvl: +(wa(kid(e, 'ilvl'), 'val') || 0) }; }
  if ((e = kid(pPr, 'spacing'))) { const s = {}; for (const a of ['before', 'after', 'line', 'lineRule', 'beforeAutospacing', 'afterAutospacing']) { const v = wa(e, a); if (v != null) s[a] = /Rule|Auto/.test(a) ? v : +v; } o.sp = s; }
  if ((e = kid(pPr, 'ind'))) { const s = {}; for (const [a, b] of [['left', 'left'], ['start', 'left'], ['right', 'right'], ['end', 'right'], ['firstLine', 'first'], ['hanging', 'hang']]) { const v = wa(e, a); if (v != null && s[b] === undefined) s[b] = +v; } o.ind = s; }
  if ((e = kid(pPr, 'contextualSpacing'))) o.ctx = onoff(e);
  if ((e = kid(pPr, 'jc'))) o.jc = wa(e, 'val');
  if ((e = kid(pPr, 'shd'))) { const f = wa(e, 'fill'); if (f && /^[0-9a-f]{6}$/i.test(f)) o.shd = f.toUpperCase(); }
  if ((e = kid(pPr, 'pBdr'))) { const b = {}; for (const side of ['top', 'left', 'bottom', 'right', 'between']) { const x = kid(e, side), v = wa(x, 'val'); if (x && v && v !== 'nil' && v !== 'none') b[side] = { val: v, sz: +(wa(x, 'sz') || 4), space: +(wa(x, 'space') || 1), color: wa(x, 'color') || 'auto' }; } if (Object.keys(b).length) o.bdr = b; }
  return o;
}
function writePProp(pPr, k, v) {
  for (const n of P_EL[k]) for (let e; (e = kid(pPr, n));) e.remove();
  if (v == null || ((k === 'sp' || k === 'ind' || k === 'bdr') && !Object.keys(v).length)) return; const d = pPr.ownerDocument; let e;
  if (k === 'style') e = wel(d, 'pStyle', { val: v });
  else if (k === 'num') { e = wel(d, 'numPr'); e.appendChild(wel(d, 'ilvl', { val: v.lvl || 0 })); e.appendChild(wel(d, 'numId', { val: v.id })); }
  else if (k === 'sp') e = wel(d, 'spacing', v);
  else if (k === 'ind') e = wel(d, 'ind', { left: v.left, right: v.right, firstLine: v.hang != null ? null : v.first, hanging: v.hang });
  else if (k === 'jc') e = wel(d, 'jc', { val: v });
  else if (k === 'shd') e = wel(d, 'shd', { val: 'clear', color: 'auto', fill: v });
  else if (k === 'bdr') { e = wel(d, 'pBdr'); for (const side of ['top', 'left', 'bottom', 'right', 'between']) if (v[side]) e.appendChild(wel(d, side, { val: v[side].val, sz: v[side].sz, space: v[side].space, color: v[side].color })); }
  else e = wel(d, P_EL[k][0], v ? null : { val: '0' });
  insertOrdered(pPr, e, P_ORDER);
}
function extrasP(pPr) { // shown, never changed: shading, borders, a section break
  const x = {}; if (!pPr) return x;
  const shd = wa(kid(pPr, 'shd'), 'fill'); if (shd && /^[0-9a-f]{6}$/i.test(shd) && shd.toUpperCase() !== 'FFFFFF') x.shd = shd;
  const b = kid(pPr, 'pBdr'); if (b) { const css = []; for (const side of ['top', 'bottom', 'left', 'right']) { const e = kid(b, side), v = wa(e, 'val'); if (e && v && v !== 'nil' && v !== 'none') css.push(`border-${side}:${Math.max(0.75, (+wa(e, 'sz') || 4) / 8)}pt ${/double|dotted|dashed/.test(v) ? v.replace(/ed$/, 'ed') : 'solid'} ${colorOf(wa(e, 'color'))}`); } if (css.length) x.bdr = css.join(';'); }
  if (kid(pPr, 'sectPr')) x.sect = true;
  return x;
}
const colorOf = v => v && /^[0-9a-f]{6}$/i.test(v) ? '#' + v : '#000';

/* ───────── state ───────── */
const D = { open: false, zip: null, name: 'Document', parts: {}, main: null, ST: null, NUM: null, theme: { major: 'Calibri Light', minor: 'Calibri' }, sect: null,
  PP: new Map(), RP: new Map(), RPK: new Map(), AT: new Map(), WR: new Map(), TB: new Map(), media: new Map(), cc: new Map(), seq: 0,
  dirty: false, undo: [], redo: [], mv: 0, model: null, zoom: 1, pending: null, lastSel: null, notes: 0, hdrPart: null, ftrPart: null, tab: 'proof', side: store.get('xd3.dxside') !== 'off' };
const nid = () => 'k' + (++D.seq).toString(36);
const RFC = new Map(), PFC = new Map(), RCSS = new Map();
const parseJ = (cache, s) => { let o = cache.get(s); if (!o) { try { o = JSON.parse(s || '{}'); } catch (e) { o = {}; } cache.set(s, o); } return o; };
const rfOf = el => parseJ(RFC, el.dataset.rf), pfOf = el => parseJ(PFC, el.dataset.pf);
const setRf = (el, o) => { el.dataset.rf = JSON.stringify(o); }, setPf = (el, o) => { el.dataset.pf = JSON.stringify(o); };

/* ───────── styles, numbering, theme ───────── */
function loadStyles(doc) {
  const S = { defR: {}, defP: {}, byId: {}, byName: {}, defPara: null, defChar: null, doc };
  if (!doc) return S;
  const root = doc.documentElement, dd = kid(root, 'docDefaults');
  S.defR = readR(kid(kid(dd, 'rPrDefault'), 'rPr')); S.defP = readP(kid(kid(dd, 'pPrDefault'), 'pPr'));
  for (const st of kids(root, 'style')) {
    const id = wa(st, 'styleId'), type = wa(st, 'type') || 'paragraph', name = wa(kid(st, 'name'), 'val') || id, pPr = kid(st, 'pPr'), ol = kid(pPr, 'outlineLvl');
    const o = { id, type, name, based: wa(kid(st, 'basedOn'), 'val'), next: wa(kid(st, 'next'), 'val'), r: readR(kid(st, 'rPr')), p: readP(pPr), outline: ol ? +wa(ol, 'val') : null,
      hidden: !!(kid(st, 'semiHidden') || kid(st, 'hidden')), q: !!kid(st, 'qFormat'), prio: +(wa(kid(st, 'uiPriority'), 'val') ?? 99), el: st, borders: type === 'table' ? hasBorders(kid(kid(st, 'tblPr'), 'tblBorders')) : null };
    S.byId[id] = o; S.byName[name.toLowerCase()] = o;
    if (onoffAttr(wa(st, 'default'))) { if (type === 'paragraph') S.defPara = id; else if (type === 'character') S.defChar = id; }
  }
  return S;
}
const onoffAttr = v => v === '1' || v === 'true' || v === 'on';
const hasBorders = el => el ? [...el.children].some(c => { const v = wa(c, 'val'); return v && v !== 'nil' && v !== 'none'; }) : null;
function chainOf(id, k, depth = 0) { // a style's properties with everything it is based on
  const s = id && D.ST.byId[id]; if (!s || depth > 15) return {};
  const key = id + '|' + k; if (D.cc.has(key)) return D.cc.get(key);
  const base = s.based ? chainOf(s.based, k, depth + 1) : {}, v = k === 'p' ? mergeP(base, s.p) : { ...base, ...s.r };
  D.cc.set(key, v); return v;
}
function mergeP(...ls) { const o = {}; for (const l of ls) for (const k in l) o[k] = (k === 'sp' || k === 'ind') && l[k] ? { ...(o[k] || {}), ...l[k] } : l[k]; return o; }
const styleLevel = id => { // 0 = title, 1–9 = heading level, null = body text
  for (let s = D.ST.byId[id], n = 0; s && n < 12; s = D.ST.byId[s.based], n++) { const m = /^heading (\d)$/i.exec(s.name); if (m) return +m[1]; if (/^title$/i.test(s.name)) return 0; if (s.outline != null && s.outline < 9) return s.outline + 1; }
  return null;
};
function loadNumbering(doc) {
  const N = { abs: {}, nums: {}, doc }; if (!doc) return N; const root = doc.documentElement;
  for (const a of kids(root, 'abstractNum')) {
    const lv = {};
    for (const l of kids(a, 'lvl')) { const pPr = kid(l, 'pPr'), rf = kid(kid(l, 'rPr'), 'rFonts'); lv[+(wa(l, 'ilvl') || 0)] = { fmt: wa(kid(l, 'numFmt'), 'val') || 'decimal', text: wa(kid(l, 'lvlText'), 'val') ?? '', start: +(wa(kid(l, 'start'), 'val') ?? 1), suff: wa(kid(l, 'suff'), 'val') || 'tab', font: rf ? wa(rf, 'ascii') || wa(rf, 'hAnsi') : null, ind: kid(pPr, 'ind') ? readP(pPr).ind : null }; }
    N.abs[wa(a, 'abstractNumId')] = { lvls: lv, link: wa(kid(a, 'numStyleLink'), 'val') };
  }
  for (const n of kids(root, 'num')) { const over = {}; for (const o of kids(n, 'lvlOverride')) { const s = kid(o, 'startOverride'); if (s) over[+wa(o, 'ilvl')] = +wa(s, 'val'); } N.nums[wa(n, 'numId')] = { abs: wa(kid(n, 'abstractNumId'), 'val'), over }; }
  for (const id in N.abs) { const a = N.abs[id]; if (a.link && !Object.keys(a.lvls).length) { const st = D.ST.byId[a.link], ni = st && st.p.num && st.p.num.id, t = ni && N.nums[ni] && N.abs[N.nums[ni].abs]; if (t) a.lvls = t.lvls; } } // list styles
  return N;
}
const lvlOf = num => { const n = num && D.NUM.nums[num.id], a = n && D.NUM.abs[n.abs]; return a ? a.lvls[num.lvl || 0] || a.lvls[0] || null : null; };
const SYM = { '\uf0b7': '•', '\uf0a7': '▪', '\uf0d8': '➢', '\uf076': '❖', '\uf0fc': '✓', '\uf06e': '■', '\uf071': '❑', '\uf0a8': '◆', '\uf0e0': '➔', '\uf0f0': '⇨', '\uf0d7': '▸', '\uf0b2': '◊', '\uf02d': '–', '\uf0a1': '○' };
const bulletText = (t, font) => { const c = t || '•'; if (SYM[c]) return SYM[c]; if (c.charCodeAt(0) >= 0xf000) return '•'; return /wingdings|symbol/i.test(font || '') && c === 'o' ? '○' : c; };
function fmtNum(n, fmt) {
  if (fmt === 'lowerLetter' || fmt === 'upperLetter') { let s = ''; for (let v = n; v > 0; v = Math.floor((v - 1) / 26)) s = String.fromCharCode(97 + (v - 1) % 26) + s; return fmt === 'upperLetter' ? s.toUpperCase() : s; }
  if (fmt === 'lowerRoman' || fmt === 'upperRoman') { let s = '', v = n; for (const [a, r] of [[1000, 'm'], [900, 'cm'], [500, 'd'], [400, 'cd'], [100, 'c'], [90, 'xc'], [50, 'l'], [40, 'xl'], [10, 'x'], [9, 'ix'], [5, 'v'], [4, 'iv'], [1, 'i']]) while (v >= a) { s += r; v -= a; } return fmt === 'upperRoman' ? s.toUpperCase() : s; }
  if (fmt === 'decimalZero') return String(n).padStart(2, '0');
  if (fmt === 'ordinal') return n + ([, 'st', 'nd', 'rd'][n % 100 > 10 && n % 100 < 14 ? 0 : n % 10] || 'th');
  return String(n);
}

/* ───────── package parts ───────── */
async function loadPart(name) {
  const f = D.zip.file(name); if (!f) return null;
  const relsName = name.replace(/([^/]+)$/, '_rels/$1.rels'), rf = D.zip.file(relsName);
  const doc = parseXml(await f.async('string'));
  return D.parts[name] = { name, doc, ser0: ser(doc), rels: rf ? parseXml(await rf.async('string')) : null, relsName, dir: name.replace(/[^/]+$/, ''), relsDirty: false };
}
const relsOf = part => { const m = {}; if (part && part.rels) for (const r of part.rels.documentElement.children) m[r.getAttribute('Id')] = { type: r.getAttribute('Type') || '', target: r.getAttribute('Target') || '', ext: r.getAttribute('TargetMode') === 'External' }; return m; };
const resolvePath = (part, t) => { if (t.startsWith('/')) return t.slice(1); const out = []; for (const s of (part.dir + t).split('/')) { if (s === '..') out.pop(); else if (s && s !== '.') out.push(s); } return out.join('/'); };
function addRel(part, type, target, external) {
  if (!part.rels) part.rels = parseXml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${NS.rel}"></Relationships>`);
  const root = part.rels.documentElement, ids = new Set([...root.children].map(r => r.getAttribute('Id'))); let n = root.children.length + 1; while (ids.has('rId' + n)) n++;
  const r = part.rels.createElementNS(NS.rel, 'Relationship'); r.setAttribute('Id', 'rId' + n); r.setAttribute('Type', type.startsWith('http') ? type : RT + type); r.setAttribute('Target', target); if (external) r.setAttribute('TargetMode', 'External');
  root.appendChild(r); part.relsDirty = true; return 'rId' + n;
}
function ensureType({ ext, part, type }) { // [Content_Types].xml
  if (!D.ct) return; const root = D.ct.documentElement;
  if (ext) { if ([...root.children].some(c => c.localName === 'Default' && (c.getAttribute('Extension') || '').toLowerCase() === ext)) return; const e = D.ct.createElementNS(NS.ct, 'Default'); e.setAttribute('Extension', ext); e.setAttribute('ContentType', type); root.insertBefore(e, root.firstChild); }
  else { if ([...root.children].some(c => c.localName === 'Override' && c.getAttribute('PartName') === '/' + part)) return; const e = D.ct.createElementNS(NS.ct, 'Override'); e.setAttribute('PartName', '/' + part); e.setAttribute('ContentType', type); root.appendChild(e); }
  D.ctDirty = true;
}
async function loadMedia(part) { // pictures of a part → blob URLs (formats a browser cannot show get a placeholder)
  for (const [id, r] of Object.entries(relsOf(part))) {
    if (!/\/image$/.test(r.type) || r.ext) continue; const path = resolvePath(part, r.target), f = D.zip.file(path), ext = (path.split('.').pop() || '').toLowerCase();
    const mime = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', webp: 'image/webp', svg: 'image/svg+xml' }[ext];
    if (f && mime) D.media.set(part.name + '#' + id, URL.createObjectURL(new Blob([await f.async('uint8array')], { type: mime })));
  }
}

/* ───────── XML → editable page ───────── */
const MARKS = new Set(['bookmarkStart', 'bookmarkEnd', 'commentRangeStart', 'commentRangeEnd', 'permStart', 'permEnd', 'moveFromRangeStart', 'moveFromRangeEnd', 'moveToRangeStart', 'moveToRangeEnd', 'customXmlInsRangeStart', 'customXmlInsRangeEnd', 'customXmlDelRangeStart', 'customXmlDelRangeEnd']);
const WRAPS = new Set(['hyperlink', 'ins', 'smartTag', 'customXml', 'fldSimple', 'dir', 'bdo', 'moveTo', 'sdt']);
const emptyP = () => { const p = document.createElement('p'); p.dataset.pf = '{}'; p.appendChild(document.createElement('br')); return p; };
function renderBlocks(xmlBox, part, out, pend) {
  for (const c of xmlBox.children) {
    const n = c.localName;
    if (n === 'p') { const p = renderP(c, part); if (pend.length) p.prepend(...pend.splice(0)); out.appendChild(p); }
    else if (n === 'tbl') out.appendChild(renderTbl(c, part, pend));
    else if (n === 'sdt' || n === 'customXml') { const w = wrapEl('div', 'dx-sdt', c, part); renderBlocks(n === 'sdt' ? kid(c, 'sdtContent') || c : c, part, w, pend); if (!w.querySelector('p')) w.appendChild(emptyP()); out.appendChild(w); }
    else if (/Pr$/.test(n) || n === 'proofErr') continue;
    else if (MARKS.has(n)) pend.push(atomEl(c, part));
    else out.appendChild(atomEl(c, part, true));
  }
  return out;
}
function renderRoot(xmlRoot, part) { // a whole body, header or footer; bookmarks left over at the end go into the last paragraph
  const pend = [], frag = renderBlocks(xmlRoot, part, document.createDocumentFragment(), pend);
  if (pend.length) { const ps = frag.querySelectorAll('p'); if (ps.length) ps[ps.length - 1].append(...pend); else { const p = emptyP(); p.prepend(...pend); frag.appendChild(p); } }
  return frag;
}
function renderP(wp, part) {
  const p = document.createElement('p'), pPr = kid(wp, 'pPr'), id = nid(), props = readP(pPr);
  D.PP.set(id, { el: pPr, props, attrs: [...wp.attributes].filter(a => !/^xmlns/.test(a.name)).map(a => [a.namespaceURI, a.name, a.value]), mark: readR(kid(pPr, 'rPr')), x: extrasP(pPr), part: part.name });
  p.dataset.p = id; setPf(p, props);
  renderInline(wp, part, p);
  const U = units(p); if (!U.length) p.appendChild(document.createElement('br')); else if (U[U.length - 1].t === 'b') p.appendChild(document.createElement('br')); // (a trailing line break needs the browser's placeholder after it)
  return p;
}
function wrapEl(tag, cls, x, part) {
  const id = nid(), e = document.createElement(tag); D.WR.set(id, { el: x, part: part.name }); e.dataset.w = id; e.className = cls;
  if (x.localName === 'hyperlink') { const rid = x.getAttributeNS(NS.r, 'id'), r = rid && relsOf(part)[rid], anchor = wa(x, 'anchor'); e.setAttribute('href', r ? r.target : anchor ? '#' + anchor : '#'); e.title = (r ? r.target : 'Bookmark ' + (anchor || '')) + ' — Ctrl+click to open'; }
  return e;
}
function renderInline(el, part, target) {
  for (const c of el.children) {
    const n = c.localName;
    if (/Pr$/.test(n) || n === 'proofErr') continue;
    if (n === 'r') renderRun(c, part, target);
    else if (WRAPS.has(n)) { const w = wrapEl(n === 'hyperlink' ? 'a' : 'span', 'dx-w dx-' + n, c, part); renderInline(n === 'sdt' ? kid(c, 'sdtContent') || c : c, part, w); target.appendChild(w); }
    else target.appendChild(atomEl(c, part));
  }
}
function rprId(rPr) { const key = rPr ? new XMLSerializer().serializeToString(rPr) : ''; let id = D.RPK.get(key); if (!id) { id = nid(); D.RPK.set(key, id); D.RP.set(id, { el: rPr, props: readR(rPr) }); } return id; }
function renderRun(r, part, target) {
  const rPr = kid(r, 'rPr'), rid = rprId(rPr), rf = JSON.stringify(D.RP.get(rid).props); let span = null;
  const run = () => { if (!span) { span = document.createElement('span'); span.dataset.r = rid; span.dataset.rf = rf; target.appendChild(span); } return span; };
  const text = s => { const sp = run(), l = sp.lastChild; if (l && l.nodeType === 3) l.appendData(s); else sp.appendChild(document.createTextNode(s)); };
  for (const c of r.children) {
    const n = c.localName;
    if (n === 'rPr' || n === 'lastRenderedPageBreak') continue;
    if (n === 't') text(c.textContent);
    else if (n === 'tab') text('\t');
    else if (n === 'noBreakHyphen') text('\u2011');
    else if (n === 'softHyphen') text('\u00ad');
    else if ((n === 'br' || n === 'cr') && !/page|column/.test(wa(c, 'type') || '')) run().appendChild(document.createElement('br'));
    else { span = null; const nr = r.cloneNode(false); if (rPr) nr.appendChild(rPr.cloneNode(true)); nr.appendChild(c.cloneNode(true)); target.appendChild(atomEl(nr, part)); }
  }
}
function atomEl(x, part, block) { // a piece of the document the editor shows but keeps as it is (pictures, fields, breaks, bookmarks…)
  const id = nid(); D.AT.set(id, { el: x, part: part.name });
  const e = document.createElement(block ? 'div' : 'span'); e.dataset.k = id; e.contentEditable = 'false'; e.className = 'dx-a';
  const inner = x.localName === 'r' ? [...x.children].find(c => c.localName !== 'rPr') : x, n = inner ? inner.localName : '';
  if (/^(drawing|pict|object|AlternateContent)$/.test(n)) picture(e, inner, part);
  else if (n === 'br') { e.className += ' dx-pb'; e.dataset.t = wa(inner, 'type') === 'column' ? 'Column break' : 'Page break'; }
  else if (n === 'fldChar') { e.className += ' dx-fc'; e.dataset.fc = wa(inner, 'fldCharType') || ''; }
  else if (n === 'instrText') { e.className += ' dx-instr'; e.textContent = inner.textContent; }
  else if (n === 'footnoteReference' || n === 'endnoteReference') { e.className += ' dx-note'; e.textContent = ++D.notes; e.title = n === 'footnoteReference' ? 'Footnote' : 'Endnote'; }
  else if (n === 'commentReference') { e.className += ' dx-cmt'; e.title = 'Comment (kept in the file)'; }
  else if (n === 'sym') { e.className += ' dx-sym'; const ch = parseInt(wa(inner, 'char') || '0', 16); e.textContent = SYM[String.fromCharCode(ch)] || String.fromCharCode(ch >= 0xf000 ? ch - 0xf000 : ch); }
  else if (n === 'del' || n === 'moveFrom') { e.className += ' dx-del'; e.textContent = [...inner.getElementsByTagNameNS(NS.w, 'delText')].map(t => t.textContent).join(''); e.title = 'Deleted text (tracked change)'; }
  else if (n === 'oMath' || n === 'oMathPara') { e.className += ' dx-math'; e.textContent = inner.textContent; e.title = 'Equation (kept as it is)'; }
  else if (n === 'ptab') { e.className += ' dx-ptab'; e.textContent = ' '; }
  else if (n === 'tbl') { e.className += ' dx-blk dx-frozen'; e.title = 'This table has a layout the editor keeps as it is'; e.appendChild(previewTable(inner)); }
  else if (block) { e.className += ' dx-blk'; e.textContent = inner && inner.textContent.trim() ? inner.textContent.trim().slice(0, 300) : '[embedded content]'; }
  else e.className += ' dx-z';
  return e;
}
function picture(e, inner, part) {
  if (D.paged && richDrawing(e, inner, part)) return;
  const blip = inner.getElementsByTagNameNS('*', 'blip')[0], vml = inner.getElementsByTagNameNS('*', 'imagedata')[0];
  const rid = blip ? blip.getAttributeNS(NS.r, 'embed') || blip.getAttributeNS(NS.r, 'link') : vml ? vml.getAttributeNS(NS.r, 'id') : null, url = rid && D.media.get(part.name + '#' + rid);
  const ext = inner.getElementsByTagNameNS('*', 'extent')[0]; let w = ext ? +ext.getAttribute('cx') / 12700 : 0, h = ext ? +ext.getAttribute('cy') / 12700 : 0;
  if (!w) { const shape = inner.getElementsByTagNameNS('*', 'shape')[0], st = shape && shape.getAttribute('style') || ''; const g = k => { const m = new RegExp(k + ':\\s*([\\d.]+)(pt|in|px)?').exec(st); return m ? +m[1] * (m[2] === 'in' ? 72 : m[2] === 'px' ? 0.75 : 1) : 0; }; w = g('width'); h = g('height'); }
  e.className += ' dx-img' + (inner.getElementsByTagNameNS('*', 'anchor').length ? ' dx-float' : ''); e.dataset.ow = w || 0; e.dataset.oh = h || 0;
  const box = inner.getElementsByTagNameNS('*', 'txbxContent')[0];
  if (url) { const im = document.createElement('img'); im.src = url; im.draggable = false; im.alt = ''; if (w) { im.style.width = w + 'pt'; im.style.height = h + 'pt'; } e.appendChild(im); }
  else if (box) { e.className += ' dx-txbx'; e.textContent = [...box.getElementsByTagNameNS(NS.w, 'p')].map(p => p.textContent).join('\n'); if (w) e.style.width = w + 'pt'; }
  else { e.className += ' dx-shape'; e.textContent = blip || vml ? 'Picture' : 'Shape'; if (w) { e.style.width = Math.min(w, 400) + 'pt'; e.style.height = Math.min(h || 40, 300) + 'pt'; } }
}
function previewTable(t) {
  const tb = document.createElement('table'); tb.className = 'dx-t dx-tb';
  for (const tr of t.getElementsByTagNameNS(NS.w, 'tr')) { if (tr.closest && tr.parentNode && tr.parentNode.closest && false) continue; const row = tb.insertRow(); for (const tc of kids(tr, 'tc')) row.insertCell().textContent = [...tc.getElementsByTagNameNS(NS.w, 'p')].map(p => p.textContent).join('\n'); }
  return tb;
}
function cellMar(el) { const o = {}; if (el) for (const [a, b] of [['top', 'top'], ['bottom', 'bottom'], ['left', 'left'], ['start', 'left'], ['right', 'right'], ['end', 'right']]) { const e = kid(el, a); if (e && wa(e, 'w') != null && o[b] === undefined && (wa(e, 'type') || 'dxa') === 'dxa') o[b] = +wa(e, 'w'); } return o; }
const plainTable = t => [...t.children].every(c => /^(tblPr|tblGrid|tr|bookmarkStart|bookmarkEnd|commentRangeStart|commentRangeEnd|proofErr)$/.test(c.localName))
  && kids(t, 'tr').every(tr => [...tr.children].every(c => /^(trPr|tblPrEx|tc|bookmarkStart|bookmarkEnd|commentRangeStart|commentRangeEnd|proofErr)$/.test(c.localName)));
function cellBorders(b) { // tcBorders → inset shadows (they do not change the cell's layout) and the top + bottom width in pt
  if (!b) return null; const sh = []; let tb = 0;
  for (const [n, a, x, y] of [['top', 'top', 0, 1], ['bottom', 'bottom', 0, -1], ['left', 'left', 1, 0], ['start', 'left', 1, 0], ['right', 'right', -1, 0], ['end', 'right', -1, 0]]) {
    const e = kid(b, n), v = e && wa(e, 'val'); if (!v || /^(nil|none)$/.test(v)) continue;
    const w = Math.max(0.25, (+wa(e, 'sz') || 4) / 8), c = wa(e, 'color'), col = c && /^[0-9a-f]{6}$/i.test(c) ? '#' + c : '#000';
    sh.push(`inset ${x * w}pt ${y * w}pt 0 ${col}`); if (a === 'top' || a === 'bottom') tb += w;
  }
  return sh.length ? { css: sh.join(','), tb } : null;
}
function renderTbl(t, part, pend) {
  if (!plainTable(t)) return atomEl(t, part, true);
  const id = nid(), tbl = document.createElement('table'), tblPr = kid(t, 'tblPr'), ts = wa(kid(tblPr, 'tblStyle'), 'val');
  D.TB.set(id, { el: t, part: part.name }); tbl.dataset.t = id; tbl.className = 'dx-t';
  let bord = hasBorders(kid(tblPr, 'tblBorders')); if (bord === null) { const st = ts && D.ST.byId[ts]; bord = !!(st && st.borders); } if (bord) tbl.classList.add('dx-tb');
  const grid = kids(kid(t, 'tblGrid'), 'gridCol').map(g => (+wa(g, 'w') || 0) / 20), cg = document.createElement('colgroup');
  for (const w of grid) { const c = document.createElement('col'); if (w) c.style.width = w + 'pt'; cg.appendChild(c); } tbl.appendChild(cg);
  const jc = wa(kid(tblPr, 'jc'), 'val'); if (jc === 'center') tbl.style.margin = '0 auto';
  const gw = grid.reduce((a, v) => a + v, 0), ind = +wa(kid(tblPr, 'tblInd'), 'w') || 0, stPr = ts && D.ST.byId[ts] ? kid(D.ST.byId[ts].el, 'tblPr') : null;
  if (wa(kid(tblPr, 'tblLayout'), 'type') === 'fixed' && gw) { tbl.style.tableLayout = 'fixed'; tbl.style.width = gw + 'pt'; }
  if (ind && jc !== 'center') { tbl.style.marginLeft = ind / 20 + 'pt'; tbl.dataset.ind = ind / 20; }
  const mar = cellMar(kid(tblPr, 'tblCellMar') || kid(stPr, 'tblCellMar'));
  for (const c of t.children) if (MARKS.has(c.localName)) pend.push(atomEl(c, part));
  const body = document.createElement('tbody'); tbl.appendChild(body);
  for (const tr of kids(t, 'tr')) {
    const row = document.createElement('tr'), rid = nid(); D.TB.set(rid, { el: tr, part: part.name }); row.dataset.tr = rid;
    const trh = kid(kid(tr, 'trPr'), 'trHeight'); let bw = 0; // (Word adds a cell's top and bottom border to the row height)
    for (const c of tr.children) if (MARKS.has(c.localName)) pend.push(atomEl(c, part));
    for (const tc of kids(tr, 'tc')) {
      const td = document.createElement('td'), cid = nid(), pr = kid(tc, 'tcPr'), vm = kid(pr, 'vMerge'), sh = wa(kid(pr, 'shd'), 'fill');
      const cb = cellBorders(kid(pr, 'tcBorders')); if (cb) { td.style.boxShadow = cb.css; bw = Math.max(bw, cb.tb); }
      D.TB.set(cid, { el: tc, part: part.name }); td.dataset.tc = cid; td.colSpan = +wa(kid(pr, 'gridSpan'), 'val') || 1;
      if (vm && wa(vm, 'val') !== 'restart') td.classList.add('dx-vm'); if (sh && /^[0-9a-f]{6}$/i.test(sh)) td.style.background = '#' + sh;
      const va = wa(kid(pr, 'vAlign'), 'val'); if (va === 'center' || va === 'bottom') td.style.verticalAlign = va === 'center' ? 'middle' : 'bottom';
      const m = { ...mar, ...cellMar(kid(pr, 'tcMar')) }; if (Object.keys(m).length) td.style.padding = ['top', 'right', 'bottom', 'left'].map(k => (m[k] ?? (k === 'left' || k === 'right' ? 108 : 0)) / 20 + 'pt').join(' ');
      renderBlocks(tc, part, td, pend); if (!td.querySelector('p')) td.appendChild(emptyP());
      row.appendChild(td);
    }
    if (trh && +wa(trh, 'val')) row.style.height = (+wa(trh, 'val') / 20 + bw) + 'pt';
    body.appendChild(row);
  }
  return tbl;
}

/* ───────── text positions: every paragraph is a line of characters (pictures, fields and line breaks count as one) ───────── */
const isRun = n => !!n && n.nodeType === 1 && n.tagName === 'SPAN' && n.dataset.rf != null;
const isAtom = n => !!n && n.nodeType === 1 && n.dataset.k != null;
const isWrap = n => !!n && n.nodeType === 1 && n.dataset.w != null && n.tagName !== 'DIV';
const isBlockWrap = n => !!n && n.nodeType === 1 && n.tagName === 'DIV' && n.dataset.w != null;
function units(p) {
  const out = []; let a = 0;
  const walk = el => { for (const c of el.childNodes) { if (c.nodeType === 3) { out.push({ n: c, a, len: c.length, t: 't' }); a += c.length; } else if (c.nodeType === 1) { if (c.tagName === 'BR') { out.push({ n: c, a, len: 1, t: 'b' }); a++; } else if (isAtom(c)) { out.push({ n: c, a, len: 1, t: 'a' }); a++; } else walk(c); } } };
  walk(p); out.total = a; return out;
}
const ptext = p => units(p).map(u => u.t === 't' ? u.n.data : u.t === 'b' ? '\n' : '\ufffc').join('');
function offsetOf(p, node, off) {
  const U = units(p);
  if (node.nodeType === 3) { const u = U.find(x => x.n === node); if (u) return u.a + Math.min(off, u.len); }
  const r = document.createRange(); try { r.setStart(node, off); } catch (e) { return 0; }
  for (const u of U) { try { if (r.comparePoint(u.n, 0) >= 0) return u.a; } catch (e) { /* */ } }
  return U.total;
}
function posOf(p, off) {
  const U = units(p);
  for (const u of U) if (u.t === 't' && off >= u.a && off <= u.a + u.len) return { node: u.n, offset: off - u.a };
  for (const u of U) { if (u.t !== 't' && off === u.a) return { node: u.n.parentNode, offset: [...u.n.parentNode.childNodes].indexOf(u.n) }; }
  for (const u of U) { if (u.t !== 't' && off === u.a + 1) return { node: u.n.parentNode, offset: [...u.n.parentNode.childNodes].indexOf(u.n) + 1 }; }
  const last = U[U.length - 1]; if (last && last.t === 'b') return { node: last.n.parentNode, offset: [...last.n.parentNode.childNodes].indexOf(last.n) };
  return { node: p, offset: p.childNodes.length };
}

/* ───────── the page ───────── */
const box = document.createElement('section'); box.id = 'dx'; box.hidden = true;
const tb = (c, icon, title, extra = '') => `<button class="dxb${extra}" data-c="${c}" title="${title}"><i data-icon="${icon}"></i></button>`;
const tl = (c, icon, label, title, extra = '') => `<button class="dxb dxl${extra}" data-c="${c}" title="${title}"><i data-icon="${icon}"></i><span>${label}</span></button>`; // a button with its name under the icon
const GALLERY = [['normal', 'Normal'], ['title', 'Title'], ['heading 1', 'Heading 1'], ['heading 2', 'Heading 2'], ['heading 3', 'Heading 3'], ['quote', 'Quote']];
box.innerHTML = `<header class="dx-top">
    <button class="btn dx-home" data-c="home" title="Home — the start screen; this document stays open (Continue editing)"><i data-icon="home"></i><span>Home</span></button>
    <span class="dx-logo"><i data-icon="word"></i></span>
    <input id="dxName" class="dx-name" spellcheck="false" title="File name">
    <span id="dxState" class="dx-state"></span>
    <div class="spacer"></div>
    <button class="btn ico" data-c="new" title="New document — blank or from a template (letter, résumé, invoice, report, meeting notes, memo)"><i data-icon="plus"></i></button>
    <button class="btn ico" data-c="open" title="Open a Word file (Ctrl+O)"><i data-icon="folder"></i></button>
    <button class="btn ico" data-c="undo" title="Undo (Ctrl+Z)"><i data-icon="undo"></i></button>
    <button class="btn ico" data-c="redo" title="Redo (Ctrl+Y)"><i data-icon="redo"></i></button>
    <span class="sep"></span>
    <button class="btn line" data-c="autofmt" title="Auto Format — headings, lists, spacing, fonts and punctuation tidied in one go"><i data-icon="wand"></i><span>Auto Format</span></button>
    <button class="btn line" data-c="fixfmt" title="Fix Formatting — find and fix inconsistent formatting"><i data-icon="brush"></i><span>Fix Formatting</span></button>
    <button class="btn ico" data-c="print" title="Print (Ctrl+P)"><i data-icon="print"></i></button>
    <button class="btn line" data-c="pdf" title="Save as PDF"><i data-icon="pages"></i><span>PDF</span></button>
    <button class="btn primary" data-c="save" title="Save the Word file (Ctrl+S)"><i data-icon="download"></i><span>Save .docx</span></button>
  </header>
  <nav class="dx-rtabs" id="dxRtabs"><button data-rt="home" class="on">Home</button><button data-rt="insert">Insert</button><button data-rt="layout">Layout</button><button data-rt="review">Review</button><button data-rt="view">View</button></nav>
  <div class="dx-ribbon" id="dxRib">
    <div class="dxp on" data-rp="home">
      ${tb('paint', 'brush', 'Format painter — copies the look of the text at the cursor; then select the text that should get it (double-click: keeps painting until Esc)')}
      <span class="sep"></span>
      <select id="dxStyle" title="Paragraph style"></select>
      <input id="dxFont" list="dxFonts" spellcheck="false" title="Font"><datalist id="dxFonts"></datalist>
      <input id="dxSize" list="dxSizes" spellcheck="false" title="Font size" class="dx-size"><datalist id="dxSizes">${[8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72].map(s => `<option value="${s}">`).join('')}</datalist>
      <button class="dxb dxt" data-c="grow" title="Bigger text (Ctrl+])">A<sup>+</sup></button><button class="dxb dxt" data-c="shrink" title="Smaller text (Ctrl+[)">A<sup>−</sup></button>
      <span class="sep"></span>
      ${tb('b', 'bold', 'Bold (Ctrl+B)')}${tb('i', 'italic', 'Italic (Ctrl+I)')}${tb('u', 'underline', 'Underline (Ctrl+U)')}${tb('strike', 'strike', 'Strikethrough')}${tb('sup', 'sup', 'Superscript')}${tb('sub', 'sub', 'Subscript')}
      <button class="dxb dxt" data-c="case" title="Change case — Sentence case, lowercase, UPPERCASE, Capitalize Each Word">Aa<b class="dxdd2">▾</b></button>
      <span class="dxsplit"><button class="dxb" data-c="color" title="Text colour"><i data-icon="fcolor"></i><b class="dxbar" id="dxColBar"></b></button><button class="dxb dxdd" data-c="colormenu" title="Choose a text colour">▾</button></span>
      <span class="dxsplit"><button class="dxb" data-c="hl" title="Highlight"><i data-icon="hlpen"></i><b class="dxbar" id="dxHlBar"></b></button><button class="dxb dxdd" data-c="hlmenu" title="Choose a highlight colour">▾</button></span>
      ${tb('clear', 'clearfmt', 'Clear formatting (Ctrl+Space)')}
      <span class="sep"></span>
      ${tb('al', 'alignleft', 'Align left (Ctrl+L)')}${tb('ac', 'aligncenter', 'Centre (Ctrl+E)')}${tb('ar', 'alignright', 'Align right (Ctrl+R)')}${tb('aj', 'alignjust', 'Justify (Ctrl+J)')}
      <span class="sep"></span>
      ${tb('bullets', 'listbul', 'Bullets (Ctrl+Shift+L)')}${tb('numbers', 'listnum', 'Numbering')}${tb('outdent', 'outdent', 'Decrease indent (Shift+Tab in a list)')}${tb('indent', 'indent', 'Increase indent (Tab in a list)')}${tb('spacing', 'linesp', 'Line and paragraph spacing')}
      ${tb('shade', 'drop', 'Paragraph shading')}${tb('border', 'rect', 'Paragraph borders')}
      <span class="sep"></span>
      <div class="dx-gallery">${GALLERY.map(([k, n]) => `<button class="dxg" data-c="sty" data-s="${k}" data-g="${k.replace(' ', '')}" title="${n} style">${n}</button>`).join('')}</div>
      <span class="sep"></span>
      ${tb('find', 'search', 'Find and replace (Ctrl+F / Ctrl+H)')}
      <button class="dxb dxai" data-c="aiwrite" title="AI writing assistant — humanize, fix grammar, change the tone, shorten, summarize (runs offline on this device)"><i data-icon="sparkle"></i><span>AI</span></button>
    </div>
    <div class="dxp" data-rp="insert">
      ${tl('table', 'table', 'Table', 'Insert a table')}${tl('picture', 'image', 'Picture', 'Insert a picture')}${tl('link', 'link', 'Link', 'Insert a link (Ctrl+K)')}${tl('symbol', 'omega', 'Symbol', 'Symbols')}
      <span class="sep"></span>
      ${tl('pbreak', 'pbreak', 'Page break', 'Page break (Ctrl+Enter)')}${tl('hr', 'minus', 'Line', 'Horizontal line under this paragraph')}${tl('check', 'check', 'Checklist', 'Tick boxes ☐ in front of the selected paragraphs — click a box to tick it')}${tl('date', 'cal', 'Date', 'Insert the date or time')}
      <span class="sep"></span>
      ${tl('hf', 'hf', 'Header / footer', 'Header, footer and page numbers')}${tl('pagenum', 'hash', 'Page numbers', 'Page numbers in the footer')}${tl('letterhead', 'stamp', 'Letterhead', 'Put a saved letterhead at the top of every page, or design a new one')}
      <span class="sep"></span>
      ${tl('toc', 'toc', 'Contents', 'Table of contents from the headings (Word fills in the page numbers)')}
    </div>
    <div class="dxp" data-rp="layout">
      ${tl('margins', 'pagesetup', 'Margins', 'Margins: Normal, Narrow, Moderate, Wide')}${tl('orient', 'rotate', 'Orientation', 'Portrait or landscape')}${tl('size', 'pages', 'Size', 'Paper size: A4, Letter, Legal…')}${tl('cols', 'spread', 'Columns', 'One, two or three columns')}
      <span class="sep"></span>
      ${tl('spacing', 'linesp', 'Spacing', 'Line and paragraph spacing')}${tl('indent', 'indent', 'Indent', 'Increase indent')}${tl('outdent', 'outdent', 'Outdent', 'Decrease indent')}
      <span class="sep"></span>
      ${tl('page', 'pagesetup', 'Page setup', 'Page setup — size, orientation, margins in cm')}
      <span class="sep"></span>
      ${tl('flow', 'wand', 'Make flowing', 'Make fully flowing — a document that keeps the exact look of a PDF becomes normal text that reflows as you edit (one Undo brings the exact look back)')}
    </div>
    <div class="dxp" data-rp="review">
      ${tl('proof', 'spell', 'Spelling', 'Spelling and grammar (side panel)')}${tl('human', 'sparkle', 'Humanize', 'Make the wording natural (side panel)')}${tl('aiwrite', 'sparkle', 'AI Write', 'Offline AI: humanize, fix grammar, formal / friendly tone, shorter, longer, summary, bullet points')}
      <span class="sep"></span>
      ${tl('autofmt', 'wand', 'Auto Format', 'Auto Format — headings, lists, spacing, fonts and punctuation tidied in one go')}${tl('fixfmt', 'brush', 'Fix Formatting', 'Find and fix inconsistent formatting')}
      <span class="sep"></span>
      ${tl('count', 'hash', 'Word count', 'Words, characters, paragraphs, pages and reading time')}${tl('speak', 'speaker', 'Read aloud', 'Read the selection (or from the cursor) aloud — click again to stop')}
    </div>
    <div class="dxp" data-rp="view">
      ${tl('nav', 'outline', 'Navigation', 'Headings of the document — click one to go there')}${tl('panel', 'panel', 'Side panel', 'Show / hide the side panel')}${tl('focus', 'expand', 'Focus', 'Hide the ribbon and side panel while you write (Esc to return)')}
      <span class="sep"></span>
      ${tl('zout', 'zoomout', 'Zoom out', 'Zoom out')}${tl('z100', 'search', '100%', 'Actual size')}${tl('zfit', 'fit', 'Page width', 'Fit the page width')}${tl('zin', 'zoomin', 'Zoom in', 'Zoom in')}
    </div>
    <div class="spacer"></div>
    ${tb('panel', 'panel', 'Show / hide the side panel')}
  </div>
  <div class="dx-find" id="dxFind" hidden>
    <input id="dxQ" placeholder="Find" spellcheck="false"><span id="dxQn" class="mut"></span>
    <button class="dxb" data-c="fprev" title="Previous (Shift+Enter)"><i data-icon="up"></i></button><button class="dxb" data-c="fnext" title="Next (Enter)"><i data-icon="down"></i></button>
    <input id="dxR" placeholder="Replace with" spellcheck="false"><button class="btn line" data-c="rone">Replace</button><button class="btn line" data-c="rall">All</button>
    <label class="inl"><input type="checkbox" id="dxQc">Match case</label><label class="inl"><input type="checkbox" id="dxQw">Whole words</label>
    <button class="dxb" data-c="fclose" title="Close (Esc)"><i data-icon="x"></i></button>
  </div>
  <div class="dx-main">
    <div class="dx-scroll" id="dxScroll"><div class="dx-paper" id="dxPaper">
      <div class="dx-hf dx-hdr" id="dxHdr" contenteditable="true" spellcheck="false" data-label="Header"></div>
      <div class="dx-doc" id="dxDoc" contenteditable="true" spellcheck="false"></div>
      <div class="dx-hf dx-ftr" id="dxFtr" contenteditable="true" spellcheck="false" data-label="Footer"></div>
      <div class="dx-pages" id="dxPages"></div>
    </div></div>
    <aside class="dx-side" id="dxSide">
      <nav class="wr-tabs dx-tabs"><button class="dx-sideclose" data-c="panel" title="Close the panel">✕</button><button data-t="proof" title="Spelling and grammar"><i data-icon="spell"></i><span>Spelling</span><b id="dxCnt">0</b></button><button data-t="format" title="Formatting that is not consistent"><i data-icon="wand"></i><span>Format</span><b id="dxFcnt"></b></button><button data-t="human" title="Humanize and the offline AI writing assistant"><i data-icon="sparkle"></i><span>Humanize</span><em>AI</em></button><button data-t="outline" title="Headings — click one to go there"><i data-icon="outline"></i><span>Outline</span></button><button data-t="stats" title="Words, reading time, readability"><i data-icon="stats"></i><span>Stats</span></button></nav>
      <div class="wr-pane" id="dxPane"></div>
    </aside>
  </div>
  <footer class="dx-status"><span id="dxPg"></span><span id="dxWords"></span>
    <select id="dxDialect" title="Variety of English">${['American', 'British', 'Australian', 'Canadian', 'Indian'].map(d => `<option value="${d}">${d} English</option>`).join('')}</select>
    <div class="spacer"></div>
    <button class="dxb" data-c="zout" title="Zoom out"><i data-icon="zoomout"></i></button><input type="range" id="dxZoom" min="40" max="200" step="10" value="100" title="Zoom"><button class="dxb" data-c="zin" title="Zoom in"><i data-icon="zoomin"></i></button><span id="dxZl">100%</span>
  </footer>
  <div id="dxPop" class="dx-pop" hidden></div>
  <input type="file" id="dxFile" accept=".docx" hidden><input type="file" id="dxPic" accept="image/png,image/jpeg,image/gif,image/bmp,image/webp" hidden>`;
document.body.appendChild(box); icons(box);
const docEl = $('#dxDoc'), hdrEl = $('#dxHdr'), ftrEl = $('#dxFtr'), paper = $('#dxPaper'), scroller = $('#dxScroll'), pane = $('#dxPane'), pop = $('#dxPop');
const frames = document.createElement('div'); frames.className = 'dx-frames'; paper.prepend(frames); // (the pages of the page view, behind the text)
const roots = () => [hdrEl, docEl, ftrEl].filter(r => !r.hidden); // (a document without a header or footer has no such area)
const allParas = () => roots().flatMap(r => [...r.querySelectorAll('p')]);
const rootOf = n => { const e = n && (n.nodeType === 3 ? n.parentNode : n); return e && e.closest ? e.closest('#dxDoc, #dxHdr, #dxFtr') : null; };
const paraOf = n => { const e = n && (n.nodeType === 3 ? n.parentNode : n); return e && e.closest ? e.closest('#dx p') : null; };
const partOfRoot = r => r === hdrEl ? D.hdrPart : r === ftrEl ? D.ftrPart : D.main;
try { document.execCommand('defaultParagraphSeparator', false, 'p'); } catch (e) { /* */ }

/* ───────── fonts and looks ───────── */
const LOOK = { calibri: 'Carlito', 'calibri light': 'Carlito', cambria: 'Caladea', georgia: 'Gelasio', verdana: 'DejaVu Sans', tahoma: 'DejaVu Sans' };
const LH = { calibri: 1.22, 'calibri light': 1.22, cambria: 1.17, arial: 1.15, 'times new roman': 1.15, georgia: 1.14, verdana: 1.22, 'segoe ui': 1.33, tahoma: 1.21, 'courier new': 1.13, garamond: 1.12, aptos: 1.2, 'aptos display': 1.2, carlito: 1.22, consolas: 1.17 };
const HL = { yellow: '#ffff00', green: '#00ff00', cyan: '#00ffff', magenta: '#ff00ff', blue: '#0000ff', red: '#ff0000', darkBlue: '#000080', darkCyan: '#008080', darkGreen: '#008000', darkMagenta: '#800080', darkRed: '#800000', darkYellow: '#808000', darkGray: '#808080', lightGray: '#c0c0c0', black: '#000000', white: '#ffffff' };
const fontName = f => !f ? D.ST.defFont || D.theme.minor : f === '@major' ? D.theme.major : f === '@minor' ? D.theme.minor : f;
const stacks = new Map();
function fontStack(f) {
  let s = stacks.get(f); if (s) return s;
  if (D.embNames && D.embNames.has(f)) { stacks.set(f, s = `"${f.replace(/["\\]/g, '')}", ` + (X.BUNDLED[f] ? (X.ensureFont(f), X.FONTS[f]) : 'Arial, sans-serif')); return s; }
  if (X.BUNDLED[f]) { X.ensureFont(f); s = X.FONTS[f]; }
  else {
    const lk = LOOK[f.toLowerCase()]; if (lk) X.ensureFont(lk);
    const g = /mono|courier|consol/i.test(f) ? 'monospace' : /times|georgia|cambria|garamond|serif|book|palatino|minion|baskerville|century|roman|constantia/i.test(f) && !/sans/i.test(f) ? 'serif' : 'sans-serif';
    s = `"${f.replace(/["\\]/g, '')}"${lk ? `, "XD3 ${lk}"` : ''}, ${g === 'serif' ? '"Times New Roman"' : g === 'monospace' ? '"Courier New"' : 'Arial'}, ${g}`;
  }
  stacks.set(f, s); return s;
}
function effP(p) {
  const pf = pfOf(p), sid = pf.style || D.ST.defPara || '', base = mergeP(D.ST.defP, chainOf(sid, 'p')), e = mergeP(base, pf); e.style = sid;
  if (e.num && e.num.id && e.num.id !== '0') { const L = lvlOf(e.num); if (L && L.ind) e.ind = { ...(base.ind || {}), ...L.ind, ...(pf.ind || {}) }; }
  return e;
}
const effR = (rf, pStyle) => ({ ...D.ST.defR, ...chainOf(pStyle, 'r'), ...(rf.rStyle ? chainOf(rf.rStyle, 'r') : {}), ...rf });
function cssR(e) {
  const sz = e.sz || 10, small = e.va === 'superscript' || e.va === 'subscript', s = [`font-family:${fontStack(fontName(e.font))}`, `font-size:${small ? sz * 0.65 : sz}pt`];
  s.push(`font-weight:${e.b ? 700 : 400}`, `font-style:${e.i ? 'italic' : 'normal'}`);
  const deco = []; if (e.u && e.u !== 'none') deco.push('underline'); if (e.strike || e.dstrike) deco.push('line-through');
  s.push('text-decoration:' + (deco.length ? deco.join(' ') + (e.u === 'double' || e.dstrike ? ' double' : e.u === 'dotted' ? ' dotted' : /dash/.test(e.u || '') ? ' dashed' : /wave/.test(e.u || '') ? ' wavy' : '') : 'none'));
  s.push('color:' + (e.color && e.color !== 'auto' ? '#' + e.color : 'inherit'));
  if (e.hl && e.hl !== 'none' && HL[e.hl]) s.push('background:' + HL[e.hl]); else if (e.shd) s.push('background:#' + e.shd);
  if (small) s.push('vertical-align:' + (e.va === 'superscript' ? 'super' : 'sub'));
  if (e.caps) s.push('text-transform:uppercase'); if (e.smallCaps) s.push('font-variant:small-caps');
  if (e.vanish) s.push('opacity:.45', 'text-decoration:underline dotted');
  if (e.spc) s.push(`letter-spacing:${e.spc / 20}pt`);
  if (e.pos) s.push('position:relative', `top:${-e.pos / 2}pt`);
  return s.join(';');
}
function cssP(e, m, up, down, x) {
  const sp = e.sp || {}, ind = e.ind || {}, t = v => (+v || 0) / 20, auto = v => v === '1' || v === 'true' || v === 'on';
  const before = up ? 0 : auto(sp.beforeAutospacing) ? 14 : t(sp.before), after = down ? 0 : auto(sp.afterAutospacing) ? 14 : t(sp.after);
  const left = t(ind.left), right = t(ind.right), first = ind.hang != null && +ind.hang ? -t(ind.hang) : t(ind.first);
  const font = fontName(m.font), size = m.sz || 10, f = LH[font.toLowerCase()] || 1.17, line = +sp.line || 240, rule = sp.lineRule || 'auto';
  const lh = rule === 'auto' ? (line / 240 * f).toFixed(3) : rule === 'exact' ? t(line) + 'pt' : t(line) > size * f ? t(line) + 'pt' : f.toFixed(3);
  const jc = { both: 'justify', distribute: 'justify', center: 'center', right: 'right', end: 'right' }[e.jc] || 'left';
  let s = `margin:0 ${right}pt 0 ${left}pt;padding:${before}pt 0 ${after}pt;text-indent:${first}pt;text-align:${jc};line-height:${lh};font-family:${fontStack(font)};font-size:${size}pt;--hang:${Math.max(0, -first)}pt`;
  const shd = e.shd || x.shd, bdr = e.bdr ? Object.entries(e.bdr).filter(([k]) => k !== 'between').map(([side, b]) => `border-${side}:${Math.max(0.75, b.sz / 8)}pt ${/double/.test(b.val) ? 'double' : /dot/.test(b.val) ? 'dotted' : /dash/.test(b.val) ? 'dashed' : 'solid'} ${b.color === 'auto' ? '#000' : colorOf(b.color)}` + (side === 'left' || side === 'right' ? `;padding-${side}:${b.space}pt` : '')).join(';') : x.bdr;
  if (shd) s += `;background:#${shd}`; if (bdr) s += ';' + bdr;
  return s;
}
const exactCss = (e, mk) => `;font-size:0;--fs:${mk.sz || 10}pt;--xs:${(0.3 * (+e.sp.line || 240) / 20).toFixed(2)}pt`; // exact line height L: the text sits on a baseline 0.8·L down, as in Word
function applyP(p, prev, next, lastOfId) { // paragraph and run looks from the styles (p._pe of the neighbours must be fresh)
  const e = p._pe, rec = p.dataset.p && D.PP.get(p.dataset.p), same = q => q && q.parentNode === p.parentNode && q._pe && q._pe.style === e.style;
  const x = rec ? rec.x : {}, mk = effR(rec ? rec.mark : {}, e.style), ind = e.ind || {};
  const sb = q => q && q.parentNode === p.parentNode && q._pe && q._pe.bdr && canon(q._pe.bdr) === canon(e.bdr); let eb = e; // (neighbours with the same border share one box, as in Word)
  if (e.bdr && (sb(prev) || sb(next))) { const b = { ...e.bdr }; if (sb(prev)) delete b.top; if (sb(next)) delete b.bottom; eb = { ...e, bdr: b }; }
  const css0 = cssP(eb, mk, e.ctx && same(prev), e.ctx && same(next), x);
  p._il = (+ind.left || 0) / 20; p._ir = (+ind.right || 0) / 20; p._tol = !!(D.paged && e.sp && e.sp.lineRule === 'exact'); // (a line Word fits exactly gets a hair more room: browsers measure a little wider)
  p._css0 = p._tol ? css0 + `;margin-right:calc(${p._ir}pt - 0.3% - 1px)` + exactCss(e, mk) : css0; const css = p._css0 + layCss(p); p.classList.toggle('dx-ex', p._tol);
  if (p._css !== css || p.getAttribute('style') !== p._attr) { p.style.cssText = css; p._css = css; p._attr = p.getAttribute('style'); }
  p.classList.toggle('dx-pbb', !!e.pbb); p.classList.toggle('dx-sect', !!x.sect && lastOfId);
  const lv = styleLevel(e.style); if (lv != null) { if (p.dataset.h !== String(lv)) p.dataset.h = lv; } else if (p.dataset.h != null) delete p.dataset.h;
  for (const s of p.querySelectorAll('span[data-rf]')) {
    const key = e.style + '|' + s.dataset.rf; let c = RCSS.get(key); if (!c) { const er = effR(rfOf(s), e.style); c = { css: cssR(er), e: er }; RCSS.set(key, c); }
    if (s._css !== c.css || s.getAttribute('style') !== s._attr) { s.style.cssText = c.css; s._css = c.css; s._attr = s.getAttribute('style'); } s._e = c.e;
  }
}
function restyle(list = roots()) {
  for (const root of list) {
    const ps = [...root.querySelectorAll('p')], lastOf = new Map(); ps.forEach(p => { p._pe = effP(p); if (p.dataset.p) lastOf.set(p.dataset.p, p); });
    ps.forEach((p, i) => applyP(p, ps[i - 1], ps[i + 1], lastOf.get(p.dataset.p) === p));
    renumber(ps); root._pc = ps.length;
  }
}
const prevP = p => { const q = p.previousElementSibling; return q && q.tagName === 'P' ? q : null; }, nextP = p => { const q = p.nextElementSibling; return q && q.tagName === 'P' ? q : null; };
function restyleQuick(ps) { // typing: only the paragraphs around the cursor (lists are renumbered when paragraphs come or go)
  ps = ps.filter(p => p && p.isConnected); const rs = new Set();
  for (const p of ps) { p._pe = effP(p); for (const q of [prevP(p), nextP(p)]) if (q && !q._pe) q._pe = effP(q); rs.add(rootOf(p)); }
  for (const p of ps) { const n = nextP(p); applyP(p, prevP(p), n, !(n && n.dataset.p && n.dataset.p === p.dataset.p)); }
  for (const root of rs) { if (!root) continue; const all = root.querySelectorAll('p'); if (all.length !== root._pc || ps.some(p => p._pe.num)) { renumber([...all].map(q => { if (!q._pe) q._pe = effP(q); return q; })); root._pc = all.length; } }
}
function renumber(ps) {
  const cnt = {}, last = {};
  for (const p of ps) {
    const e = p._pe, num = e && e.num, n = num && num.id !== '0' && D.NUM.nums[num.id], a = n && D.NUM.abs[n.abs], lv = num ? num.lvl || 0 : 0, L = a && (a.lvls[lv] || a.lvls[0]);
    if (!L) { if (p.dataset.label != null) delete p.dataset.label; continue; }
    const c = cnt[n.abs] || (cnt[n.abs] = []);
    if (last[n.abs] !== num.id) { last[n.abs] = num.id; for (const k in n.over) c[k] = n.over[k] - 1; }
    c[lv] = (c[lv] === undefined ? L.start - 1 : c[lv]) + 1; for (let k = lv + 1; k < 10; k++) c[k] = undefined;
    const text = L.fmt === 'bullet' ? bulletText(L.text, L.font) : L.fmt === 'none' ? '' : L.text.replace(/%(\d)/g, (m, d) => { const k = +d - 1, v = c[k] === undefined ? (a.lvls[k] ? a.lvls[k].start : 1) : c[k]; return fmtNum(v, (a.lvls[k] || L).fmt); });
    const lab = text + (L.suff === 'space' ? ' ' : ''); if (p.dataset.label !== lab) p.dataset.label = lab;
  }
}

/* ───────── keeping the editable page well-formed after the browser edits it ───────── */
function runLike(ref, p) { // a new run formatted like its neighbour (or like the paragraph mark)
  const s = document.createElement('span');
  if (ref) { if (ref.dataset.r) s.dataset.r = ref.dataset.r; s.dataset.rf = ref.dataset.rf; }
  else { const m = p && p._mr; if (m) { if (m.r) s.dataset.r = m.r; s.dataset.rf = m.rf; } else { const rec = p && p.dataset.p && D.PP.get(p.dataset.p); setRf(s, rec ? rec.mark : {}); } }
  return s;
}
function cleanInline(el) {
  for (const c of [...el.childNodes]) {
    if (c.nodeType === 3) continue;
    if (c.nodeType !== 1) { c.remove(); continue; }
    if (c.tagName === 'BR') { for (const a of [...c.attributes]) c.removeAttribute(a.name); continue; }
    if (isAtom(c)) { if (!D.AT.has(c.dataset.k)) c.remove(); continue; }
    if (c.classList.contains('dx-tab') && c.textContent === '	') continue; // (a tab laid out to its stop)
    if (isRun(c) || (isWrap(c) && D.WR.has(c.dataset.w))) { cleanInline(c); continue; }
    cleanInline(c); c.replaceWith(...c.childNodes); // anything else the browser added: its content stays, the element goes
  }
}
function fixPara(p) {
  cleanInline(p);
  for (let inner; (inner = p.querySelector('span[data-rf] span[data-rf]'));) { // a run inside a run → side by side
    const outer = inner.parentNode; if (!isRun(outer)) { outer.replaceWith(...outer.childNodes); continue; }
    const after = outer.cloneNode(false); while (inner.nextSibling) after.appendChild(inner.nextSibling);
    outer.after(inner); inner.after(after); if (!after.hasChildNodes()) after.remove(); if (!outer.hasChildNodes()) outer.remove();
  }
  const wrapStray = el => {
    for (const c of [...el.childNodes]) {
      if (c.nodeType === 3 && c.data) { let ref = c.previousSibling; while (ref && !isRun(ref)) ref = ref.previousSibling; if (!ref) { ref = c.nextSibling; while (ref && !isRun(ref)) ref = ref.nextSibling; } const s = runLike(ref, p); c.replaceWith(s); s.appendChild(c); }
      else if (c.nodeType === 3) c.remove();
      else if (isWrap(c)) wrapStray(c);
    }
  };
  wrapStray(p);
  const merge = el => {
    for (let c = el.firstChild; c; c = c.nextSibling) {
      if (isWrap(c)) { merge(c); continue; }
      if (!isRun(c)) continue;
      while (isRun(c.nextSibling) && c.nextSibling.dataset.rf === c.dataset.rf && (c.nextSibling.dataset.r || '') === (c.dataset.r || '')) { const n = c.nextSibling; while (n.firstChild) c.appendChild(n.firstChild); n.remove(); }
      c.normalize();
    }
    for (const c of [...el.childNodes]) if (isRun(c) && !c.hasChildNodes()) c.remove();
  };
  merge(p);
  if (!units(p).length) p.appendChild(document.createElement('br'));
  for (const a of [...p.attributes]) if (!/^(data-p|data-pf|data-label|data-h|style|class)$/.test(a.name)) p.removeAttribute(a.name);
}
function fixBlocks(boxEl) {
  let inl = null;
  for (const c of [...boxEl.childNodes]) {
    if (c.nodeType === 1 && c.tagName === 'P') { inl = null; if (!c.dataset.pf) c.dataset.pf = '{}'; fixPara(c); continue; }
    if (c.nodeType === 1 && c.tagName === 'TABLE' && c.dataset.t) { inl = null; for (const td of c.querySelectorAll(':scope > tbody > tr > td')) fixBlocks(td); continue; }
    if (isBlockWrap(c) && D.WR.has(c.dataset.w)) { inl = null; fixBlocks(c); continue; }
    if (c.nodeType === 1 && c.tagName === 'DIV' && isAtom(c) && D.AT.has(c.dataset.k)) { inl = null; continue; }
    if (c.nodeType === 3 && !c.data.trim() && !inl) { c.remove(); continue; }
    if (c.nodeType === 1 && /^(DIV|P|H[1-6]|LI|UL|OL|BLOCKQUOTE|SECTION|ARTICLE|PRE|TABLE|TBODY|TR|TD)$/.test(c.tagName)) { // a block the browser made → a paragraph
      inl = null; const p = document.createElement('p'); p.dataset.pf = '{}'; c.replaceWith(p); p.append(...c.childNodes); fixPara(p); continue;
    }
    if (!inl) { inl = document.createElement('p'); inl.dataset.pf = '{}'; boxEl.insertBefore(inl, c); }
    inl.appendChild(c); fixPara(inl);
  }
  if (!boxEl.querySelector('p')) boxEl.appendChild(emptyP());
}
function tidy(list = roots()) { for (const r of list) fixBlocks(r); restyle(list); }

/* ───────── selection ───────── */
function getSel() {
  const s = getSelection(); if (!s.rangeCount) return null; const r = s.getRangeAt(0); if (!rootOf(r.startContainer) || !rootOf(r.endContainer)) return null;
  let pa = paraOf(r.startContainer), pb = paraOf(r.endContainer); const P = allParas();
  if (!pa) { const n = r.startContainer.childNodes[r.startOffset]; pa = n && n.querySelector ? (n.tagName === 'P' ? n : n.querySelector('p')) : null; }
  if (!pb) { const n = r.endContainer.childNodes[Math.max(0, r.endOffset - 1)]; pb = n && n.querySelectorAll ? (n.tagName === 'P' ? n : [...n.querySelectorAll('p')].pop()) : null; }
  if (!pa || !pb) return null;
  const a = [P.indexOf(pa), paraOf(r.startContainer) ? offsetOf(pa, r.startContainer, r.startOffset) : 0], b = [P.indexOf(pb), paraOf(r.endContainer) ? offsetOf(pb, r.endContainer, r.endOffset) : units(pb).total];
  return { a, b, c: r.collapsed };
}
function setSel(x, scroll) {
  if (!x) return; const P = allParas(), pa = P[x.a[0]], pb = P[x.b[0]]; if (!pa || !pb) return;
  const A = posOf(pa, x.a[1]), B = posOf(pb, x.b[1]), r = document.createRange();
  try { r.setStart(A.node, A.offset); r.setEnd(B.node, B.offset); } catch (e) { return; }
  const s = getSelection(); s.removeAllRanges(); s.addRange(r);
  const root = rootOf(pa); if (root && document.activeElement !== root) root.focus({ preventScroll: true });
  if (scroll) { const rc = r.getBoundingClientRect(), sc = scroller.getBoundingClientRect(); if (rc.top < sc.top + 40 || rc.bottom > sc.bottom - 40) scroller.scrollTop += rc.top - sc.top - sc.height / 3; }
}
const inEditor = () => { const s = getSelection(); return s.rangeCount && !!rootOf(s.getRangeAt(0).startContainer); };
const useSel = () => { if (!inEditor() && D.lastSel) setSel(D.lastSel); return getSel(); }; // a ribbon field may have taken the focus
function selParas() { const s = useSel(); if (!s) return []; return allParas().slice(s.a[0], s.b[0] + 1); }
function splitAt(p, off) { for (const u of units(p)) if (u.t === 't' && off > u.a && off < u.a + u.len) { u.n.splitText(off - u.a); return; } }
function isolate(t, p) { // the text node alone in its own run
  const s = t.parentNode;
  if (!isRun(s)) { let ref = t.previousSibling; while (ref && !isRun(ref)) ref = ref.previousSibling; const r = runLike(ref, p); t.replaceWith(r); r.appendChild(t); return r; }
  if (s.childNodes.length === 1) return s;
  const before = s.cloneNode(false), after = s.cloneNode(false);
  while (s.firstChild !== t) before.appendChild(s.firstChild); while (t.nextSibling) after.appendChild(t.nextSibling);
  if (before.hasChildNodes()) s.before(before); if (after.hasChildNodes()) s.after(after); return s;
}
function selRuns() { // the runs inside the selection, split at its edges
  const sel = useSel(); if (!sel) return null; const P = allParas(), out = [];
  for (let i = sel.a[0]; i <= sel.b[0]; i++) {
    const p = P[i], a = i === sel.a[0] ? sel.a[1] : 0, b = i === sel.b[0] ? sel.b[1] : Infinity;
    splitAt(p, a); if (b !== Infinity) splitAt(p, b);
    for (const u of units(p)) if (u.t === 't' && u.len && u.a >= a && u.a + u.len <= b) out.push(isolate(u.n, p));
  }
  return { sel, runs: [...new Set(out)], paras: P.slice(sel.a[0], sel.b[0] + 1) };
}
function wordSel() { // caret inside a word → that word (as Word does)
  const sel = useSel(); if (!sel || !sel.c) return null; const p = allParas()[sel.a[0]], t = ptext(p), o = sel.a[1], L = /[\p{L}\p{N}'’_-]/u; let a = o, b = o;
  while (a > 0 && L.test(t[a - 1])) a--; while (b < t.length && L.test(t[b])) b++;
  if (!(a < o && o < b)) return null; setSel({ a: [sel.a[0], a], b: [sel.a[0], b] }); return sel;
}

/* ───────── undo ───────── */
const modelStr = () => { if (!D.model || D.model.mv !== D.mv) D.model = { mv: D.mv, styles: D.ST.doc ? ser(D.ST.doc) : '', num: D.NUM.doc ? ser(D.NUM.doc) : '', sect: D.sect ? new XMLSerializer().serializeToString(D.sect) : '' }; return D.model; };
const state = () => ({ h: hdrEl.innerHTML, d: docEl.innerHTML, f: ftrEl.innerHTML, sel: getSel() || D.lastSel, m: modelStr() });
function snapshot() { D.undo.push(state()); if (D.undo.length > 80) D.undo.shift(); D.redo.length = 0; dirty(); }
function restore(s) {
  hdrEl.innerHTML = s.h; docEl.innerHTML = s.d; ftrEl.innerHTML = s.f;
  if (s.m.mv !== D.mv) { // styles, lists or page setup changed in between
    if (D.ST.doc && s.m.styles) { D.stylesPart.doc = parseXml(s.m.styles); D.ST = loadStyles(D.stylesPart.doc); }
    if (D.NUM.doc && s.m.num) { D.numPart.doc = parseXml(s.m.num); D.NUM = loadNumbering(D.numPart.doc); }
    if (D.sect && s.m.sect) { const n = D.main.doc.importNode(parseXml(s.m.sect).documentElement, true); D.sect.replaceWith(n); if (D.geo === D.sect) D.geo = n; D.sect = n; }
    D.mv = s.m.mv; D.model = s.m; changedModel(true); pageGeometry();
  }
  restyle(); setSel(s.sel); afterEdit(true);
}
function undo() { if (!D.undo.length) return; D.redo.push(state()); restore(D.undo.pop()); dirty(); }
function redo() { if (!D.redo.length) return; D.undo.push(state()); restore(D.redo.pop()); dirty(); }
function dirty() { D.dirty = true; clearTimeout(D.asT); D.asT = setTimeout(autosave, 15000); $('#dxState').textContent = 'Edited'; $('[data-c=undo]', box).disabled = !D.undo.length; $('[data-c=redo]', box).disabled = !D.redo.length; }
function changedModel(keepMv) { if (!keepMv) D.mv++; D.cc.clear(); RCSS.clear(); stacks.clear(); D.ST.defFont = D.ST.defR.font ? fontName(D.ST.defR.font) : null; fillStyles(); }

/* ───────── formatting commands ───────── */
function cmd(fn, opts = {}) { // one undo step: remember, change, tidy, put the selection back
  const sel = useSel(); if (!sel && !opts.any) return toast('Click in the document first', 'info');
  snapshot(); let r; try { r = fn(sel); } catch (e) { console.error(e); toast(e.message || 'That did not work', true); }
  tidy(); if (!opts.keepSel) setSel(sel); afterEdit(); syncBar(); return r;
}
function setRunProp(k, v, toggle) {
  const s = useSel(); if (!s) return;
  if (s.c && !wordSel()) { const c = curRun(); D.pending = { ...(D.pending || {}), [k]: !toggle ? v : k === 'u' ? (c.u && c.u !== 'none' ? 'none' : 'single') : !c[k] }; D.pendingAt = JSON.stringify(getSel()); return syncBar(); } // the next typed letters get it
  cmd(() => {
    const R = selRuns(); if (!R || !R.runs.length) return;
    const on = toggle && R.runs.every(r => { const e = r._e || effR(rfOf(r), effP(paraOf(r)).style); return k === 'u' ? e.u && e.u !== 'none' : !!e[k]; });
    for (const r of R.runs) {
      const rf = { ...rfOf(r) }, st = effR({ ...rf, [k]: undefined }, effP(paraOf(r)).style);
      let val = toggle ? (k === 'u' ? (on ? 'none' : 'single') : !on) : v;
      if (val === undefined || (toggle && (k === 'u' ? (val === 'none') === (!st.u || st.u === 'none') : !!val === !!st[k]))) delete rf[k]; else rf[k] = val; // (a value the style already gives is not repeated)
      if (k === 'va' && val === rf.va && !toggle) { /* */ }
      setRf(r, rf);
    }
  });
  if (s.c) setSel(s);
}
function curRun() { // the formatting at the cursor (what the ribbon shows)
  const s = getSelection(); let e = {};
  if (s.rangeCount && inEditor()) {
    const r0 = s.getRangeAt(0); let n = r0.startContainer; if (n.nodeType === 1) n = n.childNodes[r0.startOffset] || n.childNodes[r0.startOffset - 1] || n;
    const el = n.nodeType === 3 ? n.parentNode : n, run = el && el.closest ? el.closest('span[data-rf]') : null, p = paraOf(n);
    if (run) e = run._e || effR(rfOf(run), p ? effP(p).style : ''); else if (p) { const rec = p.dataset.p && D.PP.get(p.dataset.p); e = effR(rec ? rec.mark : {}, effP(p).style); }
  }
  return { ...e, ...(D.pending || {}) };
}
function sizeStep(d) {
  const STEPS = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72];
  const s = useSel(); if (!s) return; if (s.c && !wordSel()) { const cur = curRun().sz || 10; D.pending = { ...(D.pending || {}), sz: d > 0 ? STEPS.find(v => v > cur) || cur + 10 : [...STEPS].reverse().find(v => v < cur) || Math.max(1, cur - 1) }; D.pendingAt = JSON.stringify(getSel()); return syncBar(); }
  cmd(() => { const R = selRuns(); for (const r of R.runs) { const cur = (r._e || {}).sz || 10; setRf(r, { ...rfOf(r), sz: d > 0 ? STEPS.find(v => v > cur) || cur + 10 : [...STEPS].reverse().find(v => v < cur) || Math.max(1, cur - 1) }); } });
  if (s.c) setSel(s);
}
function clearFormatting(all) {
  cmd(() => {
    const R = selRuns(); if (R) for (const r of R.runs) { const rf = rfOf(r); setRf(r, rf.rStyle && /hyperlink/i.test(rf.rStyle) ? { rStyle: rf.rStyle } : {}); }
    if (all) for (const p of selParas()) { const pf = pfOf(p); setPf(p, pf.num ? { num: pf.num } : {}); }
  });
}
function paraCmd(fn) { cmd(() => { for (const p of selParas()) { const pf = { ...pfOf(p) }; fn(pf, p); for (const k in pf) if (pf[k] === undefined) delete pf[k]; setPf(p, pf); } }); }
const align = v => paraCmd(pf => { pf.jc = v === 'left' ? undefined : v; });
function setStyle(sid) {
  sid = ensureStyle(sid); if (!sid) return;
  const lvl = styleLevel(sid);
  paraCmd((pf, p) => {
    pf.style = sid === D.ST.defPara ? undefined : sid;
    if (lvl != null) { // a heading: formatting that covers the whole paragraph gives way to the heading's look
      const rs = [...p.querySelectorAll('span[data-rf]')].filter(r => r.textContent.trim()); if (!rs.length) return;
      for (const k of ['font', 'sz', 'b', 'color', 'i', 'caps']) { const v0 = canon(rfOf(rs[0])[k]); if (rfOf(rs[0])[k] !== undefined && rs.every(r => canon(rfOf(r)[k]) === v0)) rs.forEach(r => { const rf = { ...rfOf(r) }; delete rf[k]; setRf(r, rf); }); }
      if (pf.num && !D.ST.byId[sid].p.num) pf.num = undefined;
    }
  });
}
function listToggle(kind) {
  const ps = selParas(); if (!ps.length) return toast('Click in the document first', 'info');
  const has = p => { const e = effP(p); if (!e.num || e.num.id === '0') return false; const L = lvlOf(e.num); return L && (kind === 'bullet' ? L.fmt === 'bullet' : L.fmt !== 'bullet'); };
  const off = ps.every(has);
  cmd(() => {
    let id = null;
    if (!off) { const prev = ps[0].previousElementSibling; const pe = prev && prev.tagName === 'P' && effP(prev); id = pe && pe.num && pe.num.id !== '0' && has(prev) ? pe.num.id : addList(kind); }
    for (const p of ps) {
      const pf = { ...pfOf(p) }, sn = chainOf(pf.style || D.ST.defPara, 'p').num;
      if (off) { if (sn) pf.num = { id: '0', lvl: 0 }; else delete pf.num; }
      else { pf.num = { id, lvl: pf.num && pf.num.id !== '0' ? pf.num.lvl || 0 : 0 }; if (pf.ind) delete pf.ind; }
      setPf(p, pf);
    }
  });
}
function indent(d) {
  paraCmd((pf, p) => {
    const e = effP(p);
    if (e.num && e.num.id !== '0') { pf.num = { id: e.num.id, lvl: clamp((e.num.lvl || 0) + d, 0, 8) }; return; }
    const ind = { ...(pf.ind || {}) }, cur = (e.ind && +e.ind.left) || 0; ind.left = Math.max(0, cur + d * 720); pf.ind = ind;
  });
}
function spacingMenu(at) {
  const it = (l, a, d) => X.mi(l, '', a, { d, keep: 1 });
  X.showCtx(at, [it('Line spacing 1.0', 'dx_ls', { v: 240 }), it('Line spacing 1.15', 'dx_ls', { v: 276 }), it('Line spacing 1.5', 'dx_ls', { v: 360 }), it('Line spacing 2.0', 'dx_ls', { v: 480 }), '-',
    it('Add space before paragraph', 'dx_sp', { k: 'before', v: 240 }), it('Remove space before', 'dx_sp', { k: 'before', v: 0 }), it('Add space after paragraph', 'dx_sp', { k: 'after', v: 160 }), it('Remove space after', 'dx_sp', { k: 'after', v: 0 }), '-',
    it('Keep with next paragraph', 'dx_keep', {}), it('Page break before', 'dx_pbb', {})]);
}

/* ───────── styles and lists that may need adding to the file ───────── */
const BUILTIN = {
  'heading 1': ['Heading1', '<w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="240" w:after="0"/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:rFonts w:ascii="Calibri Light" w:hAnsi="Calibri Light" w:cs="Calibri Light"/><w:color w:val="2F5496"/><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr>'],
  'heading 2': ['Heading2', '<w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="40" w:after="0"/><w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:rFonts w:ascii="Calibri Light" w:hAnsi="Calibri Light" w:cs="Calibri Light"/><w:color w:val="2F5496"/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr>'],
  'heading 3': ['Heading3', '<w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="40" w:after="0"/><w:outlineLvl w:val="2"/></w:pPr><w:rPr><w:rFonts w:ascii="Calibri Light" w:hAnsi="Calibri Light" w:cs="Calibri Light"/><w:color w:val="1F3763"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr>'],
  'heading 4': ['Heading4', '<w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="40" w:after="0"/><w:outlineLvl w:val="3"/></w:pPr><w:rPr><w:rFonts w:ascii="Calibri Light" w:hAnsi="Calibri Light" w:cs="Calibri Light"/><w:i/><w:iCs/><w:color w:val="2F5496"/></w:rPr>'],
  title: ['Title', '<w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/><w:contextualSpacing/></w:pPr><w:rPr><w:rFonts w:ascii="Calibri Light" w:hAnsi="Calibri Light" w:cs="Calibri Light"/><w:spacing w:val="-10"/><w:kern w:val="28"/><w:sz w:val="56"/><w:szCs w:val="56"/></w:rPr>'],
  subtitle: ['Subtitle', '<w:pPr><w:spacing w:after="160"/></w:pPr><w:rPr><w:color w:val="5A5A5A"/><w:spacing w:val="15"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr>'],
  quote: ['Quote', '<w:pPr><w:spacing w:before="200" w:after="160"/><w:ind w:left="864" w:right="864"/><w:jc w:val="center"/></w:pPr><w:rPr><w:i/><w:iCs/><w:color w:val="404040"/></w:rPr>'],
  'list paragraph': ['ListParagraph', '<w:pPr><w:ind w:left="720"/><w:contextualSpacing/></w:pPr>'],
  'no spacing': ['NoSpacing', '<w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'],
  hyperlink: ['Hyperlink', '<w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr>', 'character'],
  'table grid': ['TableGrid', '<w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders></w:tblPr>', 'table'],
};
const SHOWN = ['normal', 'title', 'subtitle', 'heading 1', 'heading 2', 'heading 3', 'heading 4', 'quote', 'list paragraph', 'no spacing'];
function ensureStyle(sidOrName) { // an id, or the English name of a built-in style ("heading 1") that is added when the file lacks it
  if (D.ST.byId[sidOrName]) return sidOrName;
  const key = String(sidOrName).toLowerCase(), have = D.ST.byName[key]; if (have) return have.id;
  const b = BUILTIN[key]; if (!b || !D.ST.doc) return null;
  const [id, inner, type = 'paragraph'] = b, root = D.ST.doc.documentElement;
  const st = frag(D.ST.doc, `<w:style w:type="${type}" w:styleId="${id}"><w:name w:val="${key}"/>${type === 'paragraph' && D.ST.defPara ? `<w:basedOn w:val="${D.ST.defPara}"/><w:next w:val="${D.ST.defPara}"/>` : ''}<w:uiPriority w:val="${/heading|title/.test(key) ? 9 : 34}"/><w:qFormat/>${inner}</w:style>`);
  root.appendChild(st); D.ST = loadStyles(D.ST.doc); changedModel(); return id;
}
function ensureNumbering() {
  if (D.NUM.doc) return;
  const name = 'word/numbering.xml', doc = parseXml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering xmlns:w="${NS.w}"></w:numbering>`);
  D.numPart = D.parts[name] = { name, doc, rels: null, relsName: 'word/_rels/numbering.xml.rels', dir: 'word/', relsDirty: false };
  addRel(D.main, 'numbering', resolveRel(D.main, name)); ensureType({ part: name, type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml' });
  D.NUM = loadNumbering(doc);
}
const resolveRel = (part, path) => path.startsWith(part.dir) ? path.slice(part.dir.length) : '/' + path;
function addList(kind) { // a new list definition → its numId
  ensureNumbering(); const doc = D.NUM.doc, root = doc.documentElement;
  const absId = Math.max(-1, ...kids(root, 'abstractNum').map(a => +wa(a, 'abstractNumId') || 0)) + 1, numId = Math.max(0, ...kids(root, 'num').map(n => +wa(n, 'numId') || 0)) + 1;
  const BUL = ['•', 'o', '▪'], FMT = kind === 'bullet' ? null : [kind === 'decimal' ? 'decimal' : kind, 'lowerLetter', 'lowerRoman'];
  let lv = ''; for (let l = 0; l < 9; l++) {
    const f = kind === 'bullet' ? 'bullet' : FMT[l % 3], text = kind === 'bullet' ? BUL[l % 3] : `%${l + 1}.`;
    lv += `<w:lvl w:ilvl="${l}"><w:start w:val="1"/><w:numFmt w:val="${f}"/><w:lvlText w:val="${text}"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="${720 * (l + 1)}" w:hanging="360"/></w:pPr>${kind === 'bullet' && l % 3 === 1 ? '<w:rPr><w:rFonts w:ascii="Courier New" w:hAnsi="Courier New" w:cs="Courier New" w:hint="default"/></w:rPr>' : ''}</w:lvl>`;
  }
  const a = frag(doc, `<w:abstractNum w:abstractNumId="${absId}"><w:multiLevelType w:val="hybridMultilevel"/>${lv}</w:abstractNum>`);
  const firstNum = kid(root, 'num'); if (firstNum) root.insertBefore(a, firstNum); else root.appendChild(a); // (every abstractNum comes before the nums)
  const n = frag(doc, `<w:num w:numId="${numId}"><w:abstractNumId w:val="${absId}"/></w:num>`);
  const after = kid(root, 'numIdMacAtCleanup'); if (after) root.insertBefore(n, after); else root.appendChild(n);
  D.NUM = loadNumbering(doc); changedModel(); return String(numId);
}

/* ───────── typing ───────── */
let typeAt = 0, typeKind = '';
for (const r of [hdrEl, docEl, ftrEl]) {
  r.addEventListener('beforeinput', e => {
    const t = e.inputType || '';
    if (/^format/.test(t)) { e.preventDefault(); const m = { formatBold: 'b', formatItalic: 'i', formatUnderline: 'u', formatStrikeThrough: 'strike', formatSuperscript: 'sup', formatSubscript: 'sub' }[t]; if (m) run(m); return; }
    if (t === 'historyUndo') { e.preventDefault(); undo(); return; } if (t === 'historyRedo') { e.preventDefault(); redo(); return; }
    const kind = /^delete/.test(t) ? 'del' : 'ins';
    if (t === 'insertParagraph' || t === 'insertLineBreak' || /Paste|Drop/.test(t) || kind !== typeKind || Date.now() - typeAt > 1200) snapshot(); else dirty();
    typeAt = Date.now(); typeKind = kind;
    if (kind === 'del') { const c = curRunEl(); if (c) { const p = paraOf(c); if (p) p._mr = { r: c.dataset.r, rf: c.dataset.rf }; } } // a paragraph emptied by deleting keeps the look of its text
    D.wide = !getSelection().isCollapsed || /Paste|Drop|Yank|Cut|Replacement/.test(t); // a change across paragraphs or tables gets a full check
    if (t === 'insertText' && D.pending && e.data) { e.preventDefault(); typePending(e.data); }
  });
  r.addEventListener('input', e => {
    const s = getSel(), P = allParas(), p = s && P[s.a[0]];
    if (e.inputType === 'insertParagraph' && p) newParagraph(p);
    if (!p || D.wide) { fixBlocks(rootOf(p) || r); restyle([rootOf(p) || r]); }
    else { const touched = [prevP(p), p, nextP(p)].filter(Boolean); touched.forEach(fixPara); if (p.parentNode && p.parentNode !== r && !p.parentNode.closest('td, .dx-sdt')) fixBlocks(r); restyleQuick(touched); }
    D.wide = false; if (s) setSel(s); afterEdit();
  });
}
const curRunEl = () => { const s = getSelection(); if (!s.rangeCount) return null; const n = s.getRangeAt(0).startContainer, e = n.nodeType === 3 ? n.parentNode : n.childNodes[s.getRangeAt(0).startOffset - 1] || n; return e && e.closest ? e.closest('span[data-rf]') : null; };
function typePending(text) { // letters typed after B / I / a size was chosen with nothing selected
  const s = getSel(); if (!s) return; const p = allParas()[s.a[0]], span = runLike(curRunEl(), p);
  const rf = { ...rfOf(span), ...D.pending }, sid = effP(p).style;
  for (const k in D.pending) { const base = effR({ ...rf, [k]: undefined }, sid)[k]; if (k === 'u' ? (rf.u || 'none') === (base || 'none') : rf[k] === base || (!rf[k] && !base)) delete rf[k]; } // (only what differs from the style is written)
  setRf(span, rf); span.textContent = text; D.pending = null;
  insertNodeAtCaret(span, s); fixPara(p); restyleQuick([p]); setSel({ a: [s.a[0], s.a[1] + text.length], b: [s.a[0], s.a[1] + text.length] }); afterEdit();
}
function newParagraph(np) { // Enter: a heading is followed by its "next" style; an empty new paragraph starts plain
  const prev = np.previousElementSibling; if (!prev || prev.tagName !== 'P') return;
  const pe = effP(prev), st = D.ST.byId[pe.style], txt = ptext(np).replace(/\n$/, '');
  if (!txt && st && st.next && st.next !== pe.style) { const pf = { ...pfOf(np) }; pf.style = st.next === D.ST.defPara ? undefined : st.next; delete pf.num; if (pf.style === undefined) delete pf.style; setPf(np, pf); np.replaceChildren(document.createElement('br')); np._mr = null; }
}

/* ───────── keys ───────── */
box.addEventListener('keydown', e => {
  if ($('#dlg').open) return;
  const k = e.key.toLowerCase(), mod = e.ctrlKey || e.metaKey, inDoc = !!rootOf(e.target);
  if (e.key === 'Escape') { if (!pop.hidden) { pop.hidden = true; return; } if (!$('#dxFind').hidden) { findBar(false); return; } }
  if (mod && !e.altKey) {
    const map = { s: () => save(), p: () => printDoc(), o: () => $('#dxFile').click(), f: () => findBar(true), h: () => { findBar(true); $('#dxR').focus(); }, z: () => (e.shiftKey ? redo() : undo()), y: redo };
    if (map[k] && (k !== 'z' && k !== 'y' || inDoc || !/INPUT|SELECT|TEXTAREA/.test(e.target.tagName))) { e.preventDefault(); e.stopPropagation(); map[k](); return; }
    if (!inDoc) return;
    const fm = { b: 'b', i: 'i', u: 'u', e: 'ac', l: e.shiftKey ? 'bullets' : 'al', r: 'ar', j: 'aj', k: 'link', ']': 'grow', '[': 'shrink', ' ': 'clearchar', '=': 'sub', '+': 'sup' };
    if (k === 'enter') { e.preventDefault(); pageBreak(); return; }
    if (fm[k] || fm[e.key]) { e.preventDefault(); e.stopPropagation(); run(fm[k] || fm[e.key]); return; }
    return;
  }
  if (mod && e.altKey && /^[1-3]$/.test(e.key) && inDoc) { e.preventDefault(); setStyle('heading ' + e.key); return; }
  if (!inDoc) return;
  if (e.key === 'Tab') {
    e.preventDefault(); const p = paraOf(getSelection().anchorNode), td = p && p.closest('td');
    const s = getSel(), pe = p && effP(p);
    if (pe && pe.num && pe.num.id !== '0' && s && s.a[1] === 0) return indent(e.shiftKey ? -1 : 1);
    if (td && td.closest('#dx table')) return cellMove(td, e.shiftKey ? -1 : 1);
    if (e.shiftKey) return; document.execCommand('insertText', false, '\t'); return;
  }
  if (e.key === 'Enter' && !e.shiftKey) {
    const s = getSel(), p = s && allParas()[s.a[0]], pe = p && effP(p);
    if (s && s.c && pe && pe.num && pe.num.id !== '0' && !ptext(p).replace(/\n$/, '').trim()) { e.preventDefault(); cmd(() => { const pf = { ...pfOf(p) }; if (chainOf(pf.style || D.ST.defPara, 'p').num) pf.num = { id: '0', lvl: 0 }; else delete pf.num; setPf(p, pf); }); } // Enter on an empty list item ends the list
  }
  if (e.key === 'Backspace') {
    const s = getSel(), p = s && allParas()[s.a[0]], pe = p && effP(p);
    if (s && s.c && s.a[1] === 0 && pe && pe.num && pe.num.id !== '0') { e.preventDefault(); cmd(() => { const pf = { ...pfOf(p) }; if (chainOf(pf.style || D.ST.defPara, 'p').num) pf.num = { id: '0', lvl: 0 }; else delete pf.num; setPf(p, pf); }); } // the bullet goes first
  }
}, true);
function cellMove(td, d) {
  const cells = [...td.closest('table').querySelectorAll(':scope > tbody > tr > td')], i = cells.indexOf(td);
  if (d > 0 && i === cells.length - 1) { tableOp('rowbelow'); return; }
  const n = cells[i + d]; if (!n) return; const p = n.querySelector('p'), P = allParas(), U = units(p), end = U.total - (U.length && U[U.length - 1].t === 'b' ? 1 : 0);
  setSel({ a: [P.indexOf(p), 0], b: [P.indexOf(p), Math.max(0, end)] }); // (the next cell's text is selected, as in Word)
}

/* ───────── clipboard ───────── */
function fragHtml(range) {
  const div = document.createElement('div'); div.appendChild(range.cloneContents());
  for (const el of div.querySelectorAll('*')) { el.removeAttribute('style'); if (isAtom(el)) { const a = D.AT.get(el.dataset.k), n = a && (a.el.localName === 'r' ? [...a.el.children].pop() : a.el).localName; if (!a || MARKS.has(n) || /Reference$/.test(n)) el.remove(); } } // (bookmarks, comments and notes are not duplicated)
  return div;
}
for (const r of [hdrEl, docEl, ftrEl]) {
  r.addEventListener('copy', e => onCopy(e, false)); r.addEventListener('cut', e => onCopy(e, true));
  r.addEventListener('paste', onPaste);
  r.addEventListener('drop', e => { const f = [...(e.dataTransfer && e.dataTransfer.files || [])]; if (f.length) { e.preventDefault(); e.stopPropagation(); const im = f.find(x => x.type.startsWith('image/')); if (im) insertPicture(im); else if (f[0] && /\.docx$/i.test(f[0].name)) openFile(f[0]); } });
}
function onCopy(e, cut) {
  const s = getSelection(); if (!s.rangeCount || s.isCollapsed) return; const r = s.getRangeAt(0), div = fragHtml(r);
  e.preventDefault(); e.clipboardData.setData('text/html', `<div data-xd3="1">${div.innerHTML}</div>`);
  const ps = div.querySelectorAll('p'); e.clipboardData.setData('text/plain', ps.length ? [...ps].map(p => ptext(p).replace(/\ufffc/g, '').replace(/\n$/, '')).join('\r\n') : s.toString());
  if (cut) { snapshot(); document.execCommand('delete'); }
}
function onPaste(e) {
  e.preventDefault(); const cd = e.clipboardData; if (!cd) return;
  const img = [...cd.files].find(f => f.type.startsWith('image/')); if (img) return insertPicture(img);
  const html = cd.getData('text/html'), text = cd.getData('text/plain');
  snapshot();
  if (html && /data-xd3="1"/.test(html)) { const t = document.createElement('template'); t.innerHTML = html; const div = t.content.querySelector('[data-xd3]'); for (const el of div.querySelectorAll('[data-k]')) if (!D.AT.has(el.dataset.k)) el.remove(); document.execCommand('insertHTML', false, div.innerHTML); }
  else if (html && /<(p|div|h\d|li|b|strong|i|em|u|table|span)\b/i.test(html)) document.execCommand('insertHTML', false, foreignHtml(html));
  else { const lines = (text || '').replace(/\r\n?/g, '\n').split('\n'); lines.forEach((l, i) => { if (i) document.execCommand('insertParagraph'); if (l) document.execCommand('insertText', false, l); }); }
  tidy(); afterEdit();
}
function foreignHtml(html) { // a web page or another program → paragraphs, headings, lists and bold / italic / underline
  const doc = new DOMParser().parseFromString(html, 'text/html'), out = []; let cur = null;
  const para = (pf = {}) => { cur = { pf, runs: [] }; out.push(cur); };
  const walk = (n, f) => {
    if (n.nodeType === 3) { const t = n.data.replace(/\s+/g, ' '); if (!t.trim() && (!cur || !cur.runs.length)) return; if (!cur) para(); cur.runs.push([t, f]); return; }
    if (n.nodeType !== 1 || /^(SCRIPT|STYLE|HEAD|META|TITLE|NOSCRIPT)$/.test(n.tagName)) return;
    const st = n.style || {}, g = { ...f };
    if (/^(B|STRONG)$/.test(n.tagName) || +st.fontWeight >= 600 || st.fontWeight === 'bold') g.b = true;
    if (/^(I|EM)$/.test(n.tagName) || st.fontStyle === 'italic') g.i = true;
    if (n.tagName === 'U' || /underline/.test(st.textDecoration || '')) g.u = 'single';
    if (/^(S|STRIKE|DEL)$/.test(n.tagName) || /line-through/.test(st.textDecoration || '')) g.strike = true;
    if (n.tagName === 'SUP') g.va = 'superscript'; if (n.tagName === 'SUB') g.va = 'subscript';
    if (n.tagName === 'BR') { para(cur ? cur.pf : {}); return; }
    const h = /^H([1-6])$/.exec(n.tagName), block = /^(P|DIV|H[1-6]|LI|BLOCKQUOTE|PRE|TR|SECTION|ARTICLE|HEADER|FOOTER)$/.test(n.tagName);
    if (block) para(h ? { style: ensureStyle('heading ' + Math.min(4, +h[1])) || undefined } : n.tagName === 'LI' ? { list: n.parentNode && n.parentNode.tagName === 'OL' ? 'decimal' : 'bullet' } : {});
    for (const c of n.childNodes) walk(c, g);
    if (n.tagName === 'TD' || n.tagName === 'TH') { if (cur) cur.runs.push(['\t', f]); }
    if (block) cur = null;
  };
  walk(doc.body, {});
  const lists = {}; return out.filter(q => q.runs.some(([t]) => t.trim())).map(q => {
    const pf = { ...q.pf }; if (pf.list) { pf.num = { id: lists[pf.list] || (lists[pf.list] = addList(pf.list)), lvl: 0 }; delete pf.list; }
    return `<p data-pf="${esc(JSON.stringify(pf))}">${q.runs.map(([t, f]) => `<span data-rf="${esc(JSON.stringify(f))}">${esc(t)}</span>`).join('')}</p>`;
  }).join('');
}

/* ───────── inserting things ───────── */
function insertNodeAtCaret(node, s) {
  const p = allParas()[s.a[0]]; splitAt(p, s.a[1]); const pos = posOf(p, s.a[1]);
  if (pos.node.nodeType === 3) { const t = pos.node, host = t.parentNode; if (pos.offset > 0 && pos.offset < t.length) t.splitText(pos.offset); if (isRun(host)) { const after = host.cloneNode(false); let n = pos.offset === 0 ? t : t.nextSibling; while (n) { const nx = n.nextSibling; after.appendChild(n); n = nx; } host.after(after); after.before(node); if (!after.hasChildNodes()) after.remove(); } else (pos.offset === 0 ? t.before(node) : t.after(node)); }
  else pos.node.insertBefore(node, pos.node.childNodes[pos.offset] || null);
  return p;
}
function newAtom(xmlString, part, attrs) { // XML for a new piece (picture, page break, field) → an atom in the page
  const x = frag(part.doc, `<w:r>${xmlString}</w:r>`);
  const el = atomEl(x, part); Object.assign(el.dataset, attrs || {}); return el;
}
/* ───────── more of what Word does: format painter, change case, shading, borders, lines, checklists, contents, page layout, reading aloud ───────── */
const SHADES = [['F2F2F2', 'Light grey'], ['D9D9D9', 'Grey'], ['DEEAF6', 'Light blue'], ['E2EFD9', 'Light green'], ['FFF2CC', 'Light yellow'], ['FBE4D5', 'Light orange'], ['EDE2F6', 'Lavender'], ['1F3864', 'Navy (use white text)'], ['404040', 'Dark grey (use white text)']];
const PAPER = [['A4', 11906, 16838], ['Letter', 12240, 15840], ['Legal', 12240, 20160], ['A5', 8391, 11906], ['A3', 16838, 23811], ['Executive', 10440, 15120]];
function painter(btn) {
  const s = useSel(); if (!s) return toast('Click in the text whose look you want to copy', 'info');
  if (D.paint) { D.paint = null; btn.classList.remove('on'); return; }
  const n = getSelection().anchorNode, el = n && (n.nodeType === 3 ? n.parentNode : n), run = el && el.closest && el.closest('span[data-rf]'), p = allParas()[s.a[0]], pf = pfOf(p), keep = {};
  for (const k of ['style', 'jc', 'sp', 'ind', 'shd', 'bdr']) if (pf[k] !== undefined) keep[k] = pf[k];
  D.paint = { rf: run ? { ...rfOf(run) } : {}, pf: keep, sticky: Date.now() - (D.paintClick || 0) < 400 }; D.paintClick = Date.now(); btn.classList.add('on');
  toast(D.paint.sticky ? 'Format painter stays on — select text to paint it, Esc to stop' : 'Now select the text that should get this look', 'info');
}
function paintSelection() {
  const s = getSelection(); if (!D.paint || !s.rangeCount || s.isCollapsed || !inEditor()) return;
  const P = D.paint; cmd(() => { const R = selRuns(); if (!R) return; for (const r of R.runs) setRf(r, { ...P.rf }); for (const p of R.paras) setPf(p, { ...pfOf(p), ...P.pf }); });
  if (!P.sticky) { D.paint = null; $('[data-c=paint]', box).classList.remove('on'); }
}
function changeCase(kind) {
  const s = useSel(); if (!s) return; if (s.c && !wordSel()) return toast('Select the text to change first', 'info');
  cmd(() => {
    const R = selRuns(); if (!R) return; let capNext = true, prev = ' ';
    for (const r of R.runs) for (const t of [...r.childNodes].filter(n => n.nodeType === 3)) {
      let out = '';
      for (const ch of t.data) {
        let c = ch; const letter = /\p{L}/u.test(ch);
        if (kind === 'upper') c = ch.toUpperCase(); else if (kind === 'lower') c = ch.toLowerCase(); else if (kind === 'toggle') c = ch === ch.toUpperCase() ? ch.toLowerCase() : ch.toUpperCase();
        else if (kind === 'title') c = letter && !/[\p{L}'’]/u.test(prev) ? ch.toUpperCase() : ch.toLowerCase();
        else if (kind === 'sentence') { c = letter ? (capNext ? ch.toUpperCase() : ch.toLowerCase()) : ch; if (letter) capNext = false; if (/[.!?]/.test(ch)) capNext = true; }
        out += c; prev = ch;
      }
      t.data = out;
    }
  });
}
const BDR = { val: 'single', sz: 6, space: 1, color: 'auto' };
function horizontalLine() {
  cmd(s => { const p = allParas()[s.a[0]], q = emptyP(); setPf(q, { bdr: { bottom: { ...BDR, sz: 8, color: '7F7F7F' } } }); p.after(q); }, { keepSel: true });
}
function checklist() {
  const ps = selParas(); if (!ps.length) return toast('Click in the document first', 'info');
  const on = ps.every(p => /^[☐☑]\s?/.test(ptext(p)));
  cmd(() => { for (const p of ps) { const t = ptext(p); if (on) { const m = /^[☐☑]\s?/.exec(t); replaceText(p, 0, m[0].length, ''); } else replaceText(p, 0, 0, '☐ '); } });
}
function toggleBox(e) { // a click on ☐ / ☑ at the start of a paragraph ticks it
  const p = e.target.closest && e.target.closest('#dx p'); if (!p || !/^[☐☑]/.test(ptext(p))) return;
  const r = document.createRange(), first = posOf(p, 0), second = posOf(p, 1); try { r.setStart(first.node, first.offset); r.setEnd(second.node, second.offset); } catch (err) { return; }
  const b = r.getBoundingClientRect(); if (e.clientX < b.left - 2 || e.clientX > b.right + 2 || e.clientY < b.top - 2 || e.clientY > b.bottom + 2) return;
  snapshot(); replaceText(p, 0, 1, ptext(p)[0] === '☐' ? '☑' : '☐'); tidy(); afterEdit();
}
function headingsForToc() { return [...docEl.querySelectorAll('p')].filter(p => !p.closest('td') && ptext(p).trim() && styleLevel(effP(p).style) != null && styleLevel(effP(p).style) <= 2); }
function insertToc() {
  const hs = headingsForToc(); if (!hs.length) return toast('Give your headings a Heading style first (Home ▸ Heading 1, 2 or 3) — the table of contents lists them', 'info');
  const s = useSel(); if (!s) return toast('Click where the table of contents should go', 'info');
  const H = tw2pt(D.pg.h - Math.abs(D.pg.t) - Math.abs(D.pg.b)) * 96 / 72, pageOf = p => Math.max(1, Math.floor((p.offsetTop - docEl.offsetTop) / H) + 1), right = Math.round(D.pg.w - D.pg.l - D.pg.r);
  const esc2 = t => t.replace(/&/g, '&amp;').replace(/</g, '&lt;');
  const rows = hs.map((p, i) => `<w:p><w:pPr><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="${right}"/></w:tabs><w:spacing w:before="0" w:after="60"/><w:ind w:left="${styleLevel(effP(p).style) * 360}"/></w:pPr>${i === 0 ? '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>' : ''}<w:r><w:t xml:space="preserve">${esc2(ptext(p).trim())}</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>${pageOf(p)}</w:t></w:r>${i === hs.length - 1 ? '<w:r><w:fldChar w:fldCharType="end"/></w:r>' : ''}</w:p>`).join('');
  const xml = `<w:sdt><w:sdtPr><w:docPartObj><w:docPartGallery w:val="Table of Contents"/><w:docPartUnique/></w:docPartObj></w:sdtPr><w:sdtContent><w:p><w:pPr><w:spacing w:before="0" w:after="120"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="28"/></w:rPr><w:t>Contents</w:t></w:r></w:p>${rows}</w:sdtContent></w:sdt>`;
  cmd(sel => { const p = allParas()[sel.a[0]]; if (!p.closest('#dxDoc')) throw new Error('Put the cursor in the body of the document'); const el = atomEl(frag(D.main.doc, xml), D.main, true); el.classList.add('dx-toc'); el.title = 'Table of contents — Word updates the page numbers when the file is opened'; el.textContent = hs.map(h => ' '.repeat(styleLevel(effP(h).style)) + ptext(h).replace(/￼/g, '').trim() + '  …  ' + pageOf(h)).join('\n'); p.before(el); }, { keepSel: true });
  D.updateFields = true; toast('Table of contents added — Word fills in the exact page numbers when it opens the file', 'success');
}
function applyPage(fn) { // every section of the document gets the same change (as Page setup does)
  snapshot(); const all = [D.sect, ...[...D.PP.values()].map(x => x.el && kid(x.el, 'sectPr')).filter(Boolean)];
  for (const sp of all) fn(sp);
  D.mv++; pageGeometry(); fitZoom(); afterEdit();
}
const pgEl = (sp, n, attrs) => { let e = kid(sp, n); if (!e) { e = wel(sp.ownerDocument, n, attrs); insertOrdered(sp, e, SECT_ORDER); } return e; };
function setPaper(name) { const z = PAPER.find(x => x[0] === name); if (!z) return; const land = D.pg.w > D.pg.h; applyPage(sp => { const e = pgEl(sp, 'pgSz'); const [w, h] = land ? [z[2], z[1]] : [z[1], z[2]]; e.setAttributeNS(NS.w, 'w:w', w); e.setAttributeNS(NS.w, 'w:h', h); if (land) e.setAttributeNS(NS.w, 'w:orient', 'landscape'); else e.removeAttributeNS(NS.w, 'orient'); }); }
function setOrient(v) { applyPage(sp => { const e = pgEl(sp, 'pgSz'), w = +wa(e, 'w') || 12240, h = +wa(e, 'h') || 15840, land = v === 'l'; if ((w > h) !== land) { e.setAttributeNS(NS.w, 'w:w', h); e.setAttributeNS(NS.w, 'w:h', w); } if (land) e.setAttributeNS(NS.w, 'w:orient', 'landscape'); else e.removeAttributeNS(NS.w, 'orient'); }); }
function setMargins(v) { if (v === 'c') return pageSetup(); const [tb2, lr] = v.split(',').map(Number); applyPage(sp => { const m = pgEl(sp, 'pgMar', { header: 720, footer: 720, gutter: 0 }); for (const [k, x] of [['top', tb2], ['bottom', tb2], ['left', lr], ['right', lr]]) m.setAttributeNS(NS.w, 'w:' + k, x); }); }
function setColumns(n) { applyPage(sp => { for (let e; (e = kid(sp, 'cols'));) e.remove(); if (n > 1) insertOrdered(sp, wel(sp.ownerDocument, 'cols', { num: n, space: 720 }), SECT_ORDER); }); }
function readAloud(btn) {
  const sy = window.speechSynthesis; if (!sy) return toast('Reading aloud is not available here', true);
  if (sy.speaking) { sy.cancel(); btn.classList.remove('on'); return; }
  const sel = getSelection().toString().trim(), s = useSel(), P = allParas(), text = sel || (s ? P.slice(s.a[0]).map(ptext).join('\n') : docEl.innerText);
  if (!text.trim()) return toast('Nothing to read', 'info');
  const u = new SpeechSynthesisUtterance(text.slice(0, 32000)); u.lang = navigator.language || 'en-US'; u.onend = u.onerror = () => btn.classList.remove('on');
  btn.classList.add('on'); sy.speak(u);
}
function pageBreak() { cmd(s => { const root = rootOf(allParas()[s.a[0]]); insertNodeAtCaret(newAtom('<w:br w:type="page"/>', partOfRoot(root)), s); }); }
async function insertPicture(file) {
  const s = useSel() || (D.lastSel); if (!s) return toast('Click in the document where the picture should go', 'info');
  try {
    const im = await X.readImage(file), bytes = new Uint8Array(await file.arrayBuffer()), ext = (/png|gif|bmp|webp/.exec(file.type) || ['jpeg'])[0].replace('jpeg', 'jpeg'), root = rootOf(allParas()[s.a[0]]), part = partOfRoot(root);
    let n = 1; while (D.zip.file(`word/media/xd3image${n}.${ext}`)) n++; const path = `word/media/xd3image${n}.${ext}`; D.zip.file(path, bytes, { createFolders: false });
    ensureType({ ext, type: file.type || 'image/' + ext }); const rid = addRel(part, 'image', resolveRel(part, path)); D.media.set(part.name + '#' + rid, URL.createObjectURL(new Blob([bytes], { type: file.type })));
    const maxW = contentWidthPt(); let w = im.w * 0.75, h = im.h * 0.75; if (w > maxW) { h *= maxW / w; w = maxW; }
    const id = 1 + Math.max(0, ...[...docPrIds()]), cx = Math.round(w * 12700), cy = Math.round(h * 12700);
    const xml = `<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="${id}" name="Picture ${id}"/><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="${id}" name="${esc(file.name || 'picture')}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>`;
    setSel(s); cmd(sel => insertNodeAtCaret(newAtom(xml, part), sel));
  } catch (e) { toast('This picture could not be added: ' + (e.message || e), true); }
}
const docPrIds = () => { const ids = new Set(); for (const r of roots()) for (const a of r.querySelectorAll('[data-k]')) { const x = D.AT.get(a.dataset.k); if (x) for (const d of x.el.getElementsByTagNameNS('*', 'docPr')) ids.add(+d.getAttribute('id') || 0); } return ids; };
async function insertTable() {
  const r = await dialog('Insert table', `<div class="row"><label>Columns<input type="number" name="c" value="3" min="1" max="20"></label><label>Rows<input type="number" name="r" value="3" min="1" max="200"></label></div><label class="inl"><input type="checkbox" name="hdr" checked>First row is a header (bold)</label>`, { ok: 'Insert' });
  if (!r) return; const C = clamp(+r.c || 3, 1, 20), R = clamp(+r.r || 3, 1, 200), sid = ensureStyle('table grid');
  cmd(s => {
    const P = allParas(), p = P[s.a[0]], root = rootOf(p), part = partOfRoot(root), tw = Math.round(contentWidthPt() * 20 / C);
    const x = frag(part.doc, `<w:tbl><w:tblPr>${sid ? `<w:tblStyle w:val="${sid}"/>` : ''}<w:tblW w:w="0" w:type="auto"/>${sid ? '' : '<w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders>'}<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/></w:tblPr><w:tblGrid>${`<w:gridCol w:w="${tw}"/>`.repeat(C)}</w:tblGrid>${`<w:tr>${`<w:tc><w:tcPr><w:tcW w:w="${tw}" w:type="dxa"/></w:tcPr><w:p/></w:tc>`.repeat(C)}</w:tr>`.repeat(R)}</w:tbl>`);
    const t = renderTbl(x, part, []);
    if (r.hdr) for (const td of t.rows[0].cells) { const q = td.querySelector('p'); q.replaceChildren(); const sp = document.createElement('span'); setRf(sp, { b: true }); sp.appendChild(document.createElement('br')); q.appendChild(sp); }
    if (!ptext(p).replace(/\n$/, '').trim() && !p.querySelector('.dx-a')) p.before(t); else p.after(t); // (an empty line takes the table; the line stays after it)
    if (!t.nextElementSibling || t.nextElementSibling.tagName !== 'P') t.after(emptyP());
    const q = t.querySelector('p'); setTimeout(() => { const P2 = allParas(); setSel({ a: [P2.indexOf(q), 0], b: [P2.indexOf(q), 0] }); }, 0);
  }, { keepSel: true });
}
function tableOp(op) {
  const s = useSel(), p = s && allParas()[s.a[0]], td = p && p.closest('#dx td'); if (!td) return toast('Click in a table first', 'info');
  const tr = td.parentNode, table = tr.closest('table'), ci = [...tr.cells].indexOf(td);
  const blankCell = c => { const n = c.cloneNode(false); n.removeAttribute('style'); n.classList.remove('dx-vm'); const q = c.querySelector('p'), np = q ? q.cloneNode(false) : emptyP(); np.replaceChildren(document.createElement('br')); n.appendChild(np); return n; };
  cmd(() => {
    for (const g of [...table.querySelectorAll('tr.dx-gaprow')]) g.remove(); // (page spacers: laid out again afterwards)
    if (op === 'rowabove' || op === 'rowbelow') { const n = tr.cloneNode(false); for (const c of tr.cells) n.appendChild(blankCell(c)); op === 'rowabove' ? tr.before(n) : tr.after(n); }
    else if (op === 'colleft' || op === 'colright') { for (const row of table.rows) { const c = row.cells[Math.min(ci, row.cells.length - 1)]; if (c) (op === 'colleft' ? c.before(blankCell(c)) : c.after(blankCell(c))); } const cg = table.querySelector('colgroup'); if (cg) cg.appendChild(document.createElement('col')); }
    else if (op === 'delrow') { if (table.rows.length <= 1) { table.replaceWith(emptyP()); } else tr.remove(); }
    else if (op === 'delcol') { const n = Math.max(...[...table.rows].map(r => r.cells.length)); if (n <= 1) table.replaceWith(emptyP()); else { for (const row of table.rows) { const c = row.cells[Math.min(ci, row.cells.length - 1)]; if (c && c.colSpan > 1) c.colSpan--; else if (c) c.remove(); } const cg = table.querySelector('colgroup'); if (cg && cg.lastChild) cg.lastChild.remove(); } }
    else if (op === 'deltable') table.replaceWith(emptyP());
  });
}
async function insertLink() {
  const s = useSel(); if (!s) return toast('Click in the document first', 'info');
  const p = allParas()[s.a[0]], sel = getSelection().toString(), cur = p && getSelection().anchorNode && (getSelection().anchorNode.parentElement || {}).closest && getSelection().anchorNode.parentElement.closest('a[data-w]');
  const r = await dialog('Link', `<label>Text${s.c && !cur ? '<input type="text" name="t" autocomplete="off">' : `<input type="text" name="t" value="${esc(sel || (cur ? cur.textContent : ''))}" disabled>`}</label><label>Web address or e-mail<input type="text" name="u" value="${esc(cur ? cur.getAttribute('href') : '')}" placeholder="https://" autocomplete="off"></label>${cur ? '<label class="inl"><input type="checkbox" name="rm">Remove the link</label>' : ''}`, { ok: 'OK' });
  if (!r) return; setSel(s);
  let url = (r.u || '').trim(); if (url && !/^(https?:|mailto:|#)/i.test(url)) url = /@/.test(url) && !/\//.test(url) ? 'mailto:' + url : 'https://' + url;
  const part = partOfRoot(rootOf(p)), hl = ensureStyle('hyperlink');
  if (cur) return cmd(() => { if (r.rm || !url) { for (const sp of cur.querySelectorAll('span[data-rf]')) { const rf = { ...rfOf(sp) }; if (rf.rStyle === hl) delete rf.rStyle; setRf(sp, rf); } cur.replaceWith(...cur.childNodes); return; } const rec = D.WR.get(cur.dataset.w), x = rec.el.cloneNode(false); const rid = addRel(part, 'hyperlink', url, true); x.setAttributeNS(NS.r, 'r:id', rid); x.removeAttributeNS(NS.w, 'anchor'); const id = nid(); D.WR.set(id, { el: x, part: part.name }); cur.dataset.w = id; cur.setAttribute('href', url); });
  if (!url) return;
  cmd(sel => {
    const rid = url.startsWith('#') ? null : addRel(part, 'hyperlink', url, true), x = part.doc.createElementNS(NS.w, 'w:hyperlink'); if (rid) x.setAttributeNS(NS.r, 'r:id', rid); else x.setAttributeNS(NS.w, 'w:anchor', url.slice(1)); x.setAttributeNS(NS.w, 'w:history', '1');
    const a = document.createElement('a'), id = nid(); D.WR.set(id, { el: x, part: part.name }); a.dataset.w = id; a.className = 'dx-w dx-hyperlink'; a.setAttribute('href', url);
    if (sel.c) { const sp = document.createElement('span'); setRf(sp, hl ? { rStyle: hl } : { color: '0563C1', u: 'single' }); sp.textContent = r.t || url; a.appendChild(sp); insertNodeAtCaret(a, sel); }
    else { const R = selRuns(); if (!R.runs.length) return; R.runs[0].before(a); for (const sp of R.runs) { setRf(sp, { ...rfOf(sp), ...(hl ? { rStyle: hl } : { color: '0563C1', u: 'single' }) }); a.appendChild(sp); } }
  });
}
const SYMBOLS = '© ® ™ § ¶ ° ± × ÷ ≈ ≠ ≤ ≥ ∞ √ € £ ¥ ₹ ¢ — – … • · ← → ↑ ↓ ✓ ✗ ½ ¼ ¾ ² ³ µ α β π Ω Σ “ ” ‘ ’ « »'.split(' ');
function symbolMenu(at) { X.showCtx(at, SYMBOLS.map(c => X.mi(c, '', 'dx_sym', { d: { v: c }, keep: 1 }))); }
async function hfMenu(at) {
  X.showCtx(at, [X.mi(D.hdrPart ? 'Edit the header' : 'Add a header', 'hf', 'dx_hf', { d: { v: 'header' }, keep: 1 }), X.mi(D.ftrPart ? 'Edit the footer' : 'Add a footer', 'hf', 'dx_hf', { d: { v: 'footer' }, keep: 1 }), '-', X.mi('Page numbers in the footer', 'hash', 'dx_hf', { d: { v: 'pagenum' }, keep: 1 })]);
}
async function headerFooter(kind) { // makes the part when the document has none (for the whole document's last section)
  const isH = kind === 'header', has = isH ? D.hdrPart : D.ftrPart, el = isH ? hdrEl : ftrEl;
  if (!has) {
    snapshot(); let n = 1; const base = isH ? 'header' : 'footer'; while (D.zip.file(`word/xd3${base}${n}.xml`)) n++;
    const name = `word/xd3${base}${n}.xml`, doc = parseXml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:${isH ? 'hdr' : 'ftr'} xmlns:w="${NS.w}" xmlns:r="${NS.r}"><w:p/></w:${isH ? 'hdr' : 'ftr'}>`);
    const part = D.parts[name] = { name, doc, rels: null, relsName: name.replace(/([^/]+)$/, '_rels/$1.rels'), dir: 'word/', relsDirty: false };
    ensureType({ part: name, type: `application/vnd.openxmlformats-officedocument.wordprocessingml.${base}+xml` });
    const rid = addRel(D.main, base, resolveRel(D.main, name)), ref = wel(D.main.doc, isH ? 'headerReference' : 'footerReference', { type: 'default' }); ref.setAttributeNS(NS.r, 'r:id', rid);
    for (const old of kids(D.sect, ref.localName).filter(x => (wa(x, 'type') || 'default') === 'default')) old.remove();
    insertOrdered(D.sect, ref, SECT_ORDER); D.mv++;
    if (isH) D.hdrPart = part; else D.ftrPart = part;
    el.replaceChildren(renderRoot(doc.documentElement, part)); el.hidden = false; tidy([el]); HFV.delete(name);
  }
  if (kind === 'pagenum') {
    const p = [...el.querySelectorAll('p')].pop(); setSel({ a: [allParas().indexOf(p), 0], b: [allParas().indexOf(p), 0] });
    cmd(() => { const q = emptyP(); q.replaceChildren(); setPf(q, { jc: 'center' }); el.appendChild(q); const part = D.ftrPart, x = part.doc.createElementNS(NS.w, 'w:fldSimple'); x.setAttributeNS(NS.w, 'w:instr', ' PAGE   \\* MERGEFORMAT '); const w = document.createElement('span'), id = nid(); D.WR.set(id, { el: x, part: part.name }); w.dataset.w = id; w.className = 'dx-w dx-fldSimple'; const sp = document.createElement('span'); setRf(sp, {}); sp.textContent = '1'; w.appendChild(sp); q.appendChild(w); if (!units(el.querySelector('p')).some(u => u.t !== 'b') && el.querySelectorAll('p').length > 1) el.querySelector('p').remove(); }, { keepSel: true });
    toast('Page numbers added to the footer — Word shows the right number on every page', 'success'); return;
  }
  el.hidden = false; if (D.paged) { frames._sig = ''; D.layDirty = true; paginate(); } const p = el.querySelector('p'), P = allParas(); setSel({ a: [P.indexOf(p), 0], b: [P.indexOf(p), 0] }, true);
}

/* ───────── page geometry ───────── */
const SECT_ORDER = ['headerReference', 'footerReference', 'footnotePr', 'endnotePr', 'type', 'pgSz', 'pgMar', 'paperSrc', 'pgBorders', 'lnNumType', 'pgNumType', 'cols', 'formProt', 'vAlign', 'noEndnote', 'titlePg', 'textDirection', 'bidi', 'rtlGutter', 'docGrid', 'printerSettings', 'sectPrChange'];
const tw2pt = v => (+v || 0) / 20;
function pageGeometry() { // the page as the first section has it (a document can change size or turn landscape later on)
  const s = D.geo || D.sect, sz = kid(s, 'pgSz'), m = kid(s, 'pgMar'), g = k => wa(m, k) == null ? 1440 : +wa(m, k);
  D.pg = { w: +wa(sz, 'w') || 12240, h: +wa(sz, 'h') || 15840, t: g('top'), b: g('bottom'), l: g('left'), r: g('right'), hd: +wa(m, 'header') || 720, ft: +wa(m, 'footer') || 720 };
  const P = D.pg; paper.style.width = tw2pt(P.w) + 'pt'; paper.style.minHeight = tw2pt(P.h) + 'pt';
  const cols = +(wa(kid(s, 'cols'), 'num') || 1); docEl.style.columnCount = cols > 1 ? cols : ''; docEl.style.columnGap = cols > 1 ? tw2pt(+(wa(kid(s, 'cols'), 'space') || 720)) + 'pt' : '';
  paper.style.padding = `${tw2pt(Math.max(P.hd, 0))}pt ${tw2pt(P.r)}pt ${tw2pt(P.ft)}pt ${tw2pt(P.l)}pt`;
  docEl.style.marginTop = Math.max(0, tw2pt(Math.abs(P.t) - P.hd)) + 'pt'; docEl.style.marginBottom = Math.max(0, tw2pt(Math.abs(P.b) - P.ft)) + 'pt';
  let st = $('#dxPageCss'); if (!st) { st = document.createElement('style'); st.id = 'dxPageCss'; document.head.appendChild(st); }
  st.textContent = `@page { size: ${tw2pt(P.w)}pt ${tw2pt(P.h)}pt; margin: ${tw2pt(Math.abs(P.t))}pt ${tw2pt(P.r)}pt ${tw2pt(Math.abs(P.b))}pt ${tw2pt(P.l)}pt; }`;
  if (D.paged) { paper.style.padding = '0'; paper.style.minHeight = ''; docEl.style.margin = '0'; docEl.style.columnCount = ''; paper.style.width = tw2pt(D.paperW || P.w) + 'pt'; D.layDirty = true; }
}
const contentWidthPt = () => tw2pt(D.pg.w - D.pg.l - D.pg.r);
async function pageSetup() {
  const P = D.pg, cm = v => (v / 567).toFixed(2), sizes = [['A4', 11906, 16838], ['Letter', 12240, 15840], ['Legal', 12240, 20160], ['A5', 8391, 11906], ['A3', 16838, 23811], ['Executive', 10440, 15120]];
  const land = P.w > P.h, cur = sizes.find(([, w, h]) => Math.abs(Math.min(P.w, P.h) - w) < 30 && Math.abs(Math.max(P.w, P.h) - h) < 30);
  const r = await dialog('Page setup', `<div class="row"><label>Paper<select name="size">${sizes.map(([n]) => `<option${cur && cur[0] === n ? ' selected' : ''}>${n}</option>`).join('')}${cur ? '' : '<option value="keep" selected>Current size</option>'}</select></label>
    <label>Orientation<select name="o"><option value="p"${land ? '' : ' selected'}>Portrait</option><option value="l"${land ? ' selected' : ''}>Landscape</option></select></label></div>
    <label>Margins<select name="m"><option value="">Keep the current margins</option><option value="1440,1440">Normal — 2.54 cm</option><option value="720,720">Narrow — 1.27 cm</option><option value="1440,1080">Moderate — 2.54 / 1.91 cm</option><option value="1440,2880">Wide — 2.54 / 5.08 cm</option><option value="c">Custom…</option></select></label>
    <div class="row"><label>Top (cm)<input type="number" step="0.1" name="t" value="${cm(Math.abs(P.t))}"></label><label>Bottom (cm)<input type="number" step="0.1" name="b" value="${cm(Math.abs(P.b))}"></label><label>Left (cm)<input type="number" step="0.1" name="l" value="${cm(P.l)}"></label><label>Right (cm)<input type="number" step="0.1" name="rr" value="${cm(P.r)}"></label></div>
    <p class="mut">Applies to the whole document.</p>`, { ok: 'Apply' });
  if (!r) return; snapshot();
  const s = sizes.find(([n]) => n === r.size); let w = s ? s[1] : Math.min(P.w, P.h), h = s ? s[2] : Math.max(P.w, P.h); if (r.o === 'l') [w, h] = [h, w];
  let mt = Math.abs(P.t), mb = Math.abs(P.b), ml = P.l, mr = P.r;
  if (r.m === 'c') { mt = Math.round(+r.t * 567); mb = Math.round(+r.b * 567); ml = Math.round(+r.l * 567); mr = Math.round(+r.rr * 567); } else if (r.m) { const [tb2, lr] = r.m.split(',').map(Number); mt = mb = tb2; ml = mr = lr; }
  const all = [D.sect, ...[...D.PP.values()].map(x => x.el && kid(x.el, 'sectPr')).filter(Boolean)];
  for (const sp of all) {
    let sz = kid(sp, 'pgSz'); if (!sz) { sz = wel(sp.ownerDocument, 'pgSz'); insertOrdered(sp, sz, SECT_ORDER); }
    sz.setAttributeNS(NS.w, 'w:w', w); sz.setAttributeNS(NS.w, 'w:h', h); if (w > h) sz.setAttributeNS(NS.w, 'w:orient', 'landscape'); else sz.removeAttributeNS(NS.w, 'orient');
    let mg = kid(sp, 'pgMar'); if (!mg) { mg = wel(sp.ownerDocument, 'pgMar', { header: 720, footer: 720, gutter: 0 }); insertOrdered(sp, mg, SECT_ORDER); }
    for (const [k, v] of [['top', mt], ['bottom', mb], ['left', ml], ['right', mr]]) mg.setAttributeNS(NS.w, 'w:' + k, v);
  }
  D.mv++; pageGeometry(); fitZoom(); afterEdit(); toast('Page setup changed', 'success');
}

/* ───────── page breaks drawn over the page (an estimate — Word lays out the final pages) ───────── */
function paginate() {
  const lay = $('#dxPages'); if (!D.pg || box.hidden) return; layTabs();
  if (D.paged) { if (lay._h) { lay.innerHTML = ''; lay._h = ''; } return layoutPaged(); } const k = 96 / 72, H = tw2pt(D.pg.h - Math.abs(D.pg.t) - Math.abs(D.pg.b)) * k, top0 = docEl.offsetTop;
  const breaks = []; let bottom = top0 + H;
  for (const b of docEl.children) {
    const t = b.offsetTop, h = b.offsetHeight;
    if (b.classList.contains('dx-pbb') && t > top0 + 2) { breaks.push(t); bottom = t + H; }
    if (t + h > bottom + 1) { if (h <= H || t > bottom) { breaks.push(t > bottom ? bottom : t); bottom = (t > bottom ? bottom : t) + H; while (t + h > bottom + 1 && h > H) { breaks.push(bottom); bottom += H; } } else { while (t + h > bottom + 1) { breaks.push(bottom); bottom += H; } } }
    for (const a of b.querySelectorAll('.dx-pb')) { const y = a.offsetTop + a.offsetHeight; if (y > top0) { breaks.push(y); bottom = y + H; } }
    if (b.classList.contains('dx-sect')) { const y = t + h; breaks.push(y); bottom = y + H; }
  }
  const html = breaks.map((y, i) => `<div class="dx-brk" style="top:${y}px"><span>Page ${i + 2}</span></div>`).join('');
  if (lay._h !== html) { lay.innerHTML = html; lay._h = html; }
  D.pages = breaks.length + 1;
  const st = scroller.scrollTop + scroller.clientHeight / 3 - paper.offsetTop * D.zoom; let cur = 1; for (const y of breaks) if (y * D.zoom < st) cur++;
  $('#dxPg').textContent = `Page ${cur} of ${D.pages}`;
}
scroller.addEventListener('scroll', () => { clearTimeout(paginate.s); paginate.s = setTimeout(paginate, 80); }, { passive: true });
if (document.fonts) document.fonts.addEventListener('loadingdone', () => { for (const p of docEl.querySelectorAll('p')) p._tsig = null; if (D.open) afterEdit(); }); // (text widths change when a font arrives)

/* ───────── page view for fixed-layout documents (PDF → Word "exact"): each section on its own pages, drawings where Word puts them ───────── */
const WPNS = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing', EMU = 12700, PXPT = 96 / 72, GAP = 18;
const ANC = new Map(), HFV = new Map(); // atom id → anchor; header part → its look for each page
const first = (el, n) => el ? el.getElementsByTagNameNS('*', n)[0] || null : null;
function exactLook(body) { // most paragraphs with exact line spacing → the file copies a fixed page (2: every one of them)
  const ps = [...body.getElementsByTagNameNS(NS.w, 'p')]; let ex = 0;
  for (const p of ps) { const sp = kid(kid(p, 'pPr'), 'spacing'); if (sp && wa(sp, 'lineRule') === 'exact') ex++; }
  return ex >= 2 && ex >= ps.length * 0.4 ? (ex === ps.length ? 2 : 1) : ex && ex === ps.length ? 0.5 : 0;
}
async function pageAnchored(P) { // drawings placed on the page (in the body or a header) also need real pages
  const re = /<wp:positionV[^>]*relativeFrom="page"/;
  if (re.test(P.ser0)) return true;
  for (const r of Object.values(relsOf(P))) if (/\/(header|footer)$/.test(r.type)) { const f = D.zip.file(resolvePath(P, r.target)); if (f && re.test(await f.async('string'))) return true; }
  return false;
}
function pagedCss() {
  if ($('#dxPagedCss')) return; const st = document.createElement('style'); st.id = 'dxPagedCss';
  st.textContent = `#dx.dx-paged .dx-paper { background: transparent; box-shadow: none; z-index: 0; overflow-x: clip; }
.dx-frames { position: absolute; left: 0; top: 0; right: 0; bottom: 0; pointer-events: none; z-index: -1; }
.dx-frame { position: absolute; background: #fff; box-shadow: 0 4px 22px rgba(0, 0, 0, .3), 0 0 0 1px rgba(0, 0, 0, .05); overflow: hidden; }
.dx-hfv { position: absolute; white-space: pre-wrap; color: #000; }
.dx-rp { display: block; margin: 0; } .dx-rt { display: table; border-collapse: collapse; } .dx-rr { display: table-row; } .dx-rc { display: table-cell; vertical-align: top; }
#dx.dx-paged .dx-doc { display: flow-root; }
.dx-ex span[data-rf], .dx-rp.dx-ex > span { line-height: 0; vertical-align: calc(-1 * var(--xs)); }
.dx-ex[data-label]::before { font-size: var(--fs); line-height: 0; vertical-align: calc(-1 * var(--xs)); }
#dx.dx-paged .dx-sect::after { display: none; }
#dx.dx-paged .dx-pbb { box-shadow: none; }
#dx.dx-paged .dx-doc .dx-anc { position: absolute; outline: 0; pointer-events: none; }
.dx-frame .dx-anc { position: absolute; }
#dx.dx-paged .dx-t:not(.dx-tb) td { border: 0; }
#dx .dx-t tr.dx-gaprow > th { padding: 0; border: 0; background: transparent; box-shadow: none; }
#dx.dx-paged .dx-t:not(.dx-tb):hover td { box-shadow: inset 0 0 0 1px #d4d9e6; }
.dx-tab { display: inline-block; height: 0; overflow: hidden; vertical-align: baseline; text-indent: 0; }
.dx-rich { position: relative; line-height: normal; vertical-align: bottom; }
.dx-rich > svg { position: absolute; left: 0; top: 0; overflow: visible; }
.dx-rich > img { width: 100%; height: 100%; }
.dx-txc { position: absolute; left: 0; top: 0; right: 0; bottom: 0; display: flex; flex-direction: column; white-space: pre-wrap; color: #000; text-align: left; }
#dx.dx-paged .dx-hf .dx-anc { display: none; }
#dx.dx-paged .dx-hf { background: rgba(255, 255, 255, .55); border-color: #c3cadb; }
.dx-ppage { position: relative; overflow: hidden; z-index: 0; break-after: page; page-break-after: always; margin: 0 auto; }
.dx-ppage:last-child { break-after: auto; page-break-after: auto; }
.dx-ppage > .dx-frame { left: 0; top: 0; z-index: -2; box-shadow: none; }
.dx-ppage > .dx-doc { position: absolute; }
.dx-ppage > .dx-doc > * { position: absolute; margin: 0; }
@media print { html.dxprint #dx.dx-paged #dxPaper { display: none !important; } html.dxprint .dx-print { display: block !important; } }`;
  document.head.appendChild(st);
}
function ancInfo(x) { // a run with a wp:anchor → where it is placed, its size and whether it is behind the text
  const a = first(x, 'anchor'); if (!a || a.namespaceURI !== WPNS) return null;
  const val = (el, n) => { const c = el && [...el.children].find(q => q.localName === n); return c ? c.textContent.trim() : null; };
  const ph = kid(a, 'positionH'), pv = kid(a, 'positionV'), ext = kid(a, 'extent'), sp = a.getAttribute('simplePos') === '1' && kid(a, 'simplePos');
  const o = { hRel: ph ? ph.getAttribute('relativeFrom') : 'column', vRel: pv ? pv.getAttribute('relativeFrom') : 'paragraph', hOff: +(val(ph, 'posOffset') || 0) / EMU, vOff: +(val(pv, 'posOffset') || 0) / EMU, hAl: val(ph, 'align'), vAl: val(pv, 'align'),
    w: ext ? +ext.getAttribute('cx') / EMU : 0, h: ext ? +ext.getAttribute('cy') / EMU : 0, behind: a.getAttribute('behindDoc') === '1', pic: !first(a, 'wsp') && !!first(a, 'blip') };
  if (sp) { o.hRel = o.vRel = 'page'; o.hOff = +sp.getAttribute('x') / EMU; o.vOff = +sp.getAttribute('y') / EMU; o.hAl = o.vAl = null; }
  return o;
}
const geoPt = g => ({ w: g.w / 20, h: g.h / 20, l: g.l / 20, r: g.r / 20, t: g.t / 20, b: g.b / 20 });
function ancPlace(o, g, paraTop) { // → top-left on its page, in pt (g in pt; paraTop = the paragraph's top on the page)
  const { w: W, h: H, l, r, t, b } = g;
  const hA = { page: [0, W], leftMargin: [0, l], insideMargin: [0, l], rightMargin: [W - r, r], outsideMargin: [W - r, r] }[o.hRel] || [l, W - l - r];
  const vA = { page: [0, H], margin: [t, H - t - b], topMargin: [0, t], insideMargin: [0, t], bottomMargin: [H - b, b], outsideMargin: [H - b, b] }[o.vRel] || [paraTop, 0];
  const at = (A, al, size, off) => al === 'center' ? A[0] + (A[1] - size) / 2 : /right|bottom|outside/.test(al || '') ? A[0] + A[1] - size : al ? A[0] : A[0] + off;
  return { x: at(hA, o.hAl, o.w, o.hOff), y: at(vA, o.vAl, o.h, o.vOff) };
}
function fillOf(spPr, style, kind) { // a DrawingML fill or outline → [css colour, opacity] (null: none)
  const el = kind === 'ln' ? kid(spPr, 'ln') : spPr;
  if (el && kid(el, 'noFill')) return null;
  const f = el && kid(el, 'solidFill'), c = f && f.firstElementChild;
  if (!c) return style && first(style, kind === 'ln' ? 'lnRef' : 'fillRef') && (kind !== 'ln' || !el || !kid(el, 'noFill')) ? [kind === 'ln' ? '#2F528F' : '#4472C4', 1] : null;
  const v = c.localName === 'srgbClr' ? '#' + c.getAttribute('val') : c.localName === 'sysClr' ? '#' + (c.getAttribute('lastClr') || '000000') : c.localName === 'prstClr' ? c.getAttribute('val') : '#7f7f7f', al = kid(c, 'alpha');
  return [v, al ? +al.getAttribute('val') / 100000 : 1];
}
function richDrawing(e, inner, part) { // pictures, shapes and text boxes as they look (false: left to the placeholder)
  const anc = first(inner, 'anchor'), ext = first(inner, 'extent'), wsp = first(inner, 'wsp'), blip = first(inner, 'blip');
  if (!ext || (!wsp && !blip)) return false;
  const w = +ext.getAttribute('cx') / EMU, h = +ext.getAttribute('cy') / EMU, spPr = first(wsp || inner, 'spPr'), xf = kid(spPr, 'xfrm');
  const rot = xf ? (+xf.getAttribute('rot') || 0) / 60000 : 0, fh = !!xf && xf.getAttribute('flipH') === '1', fv = !!xf && xf.getAttribute('flipV') === '1';
  e.className += ' dx-img dx-rich' + (anc ? ' dx-anc' : ''); e.dataset.ow = w; e.dataset.oh = h; e.style.width = w + 'pt'; e.style.height = h + 'pt';
  if (anc) e.style.zIndex = anc.getAttribute('behindDoc') === '1' ? '-1' : '2';
  const flip = [fh ? 'scaleX(-1)' : '', fv ? 'scaleY(-1)' : ''].join(' ').trim();
  if (!wsp) {
    const rid = blip.getAttributeNS(NS.r, 'embed') || blip.getAttributeNS(NS.r, 'link'), url = rid && D.media.get(part.name + '#' + rid);
    if (url) { const im = document.createElement('img'); im.src = url; im.draggable = false; im.alt = ''; im.style.width = w + 'pt'; im.style.height = h + 'pt'; e.appendChild(im); }
    if (rot || flip) e.style.transform = `${rot ? `rotate(${rot}deg) ` : ''}${flip}`;
    return true;
  }
  const geom = (kid(spPr, 'prstGeom') || kid(spPr, 'custGeom') || { getAttribute: () => 'rect' }).getAttribute('prst') || 'rect', style = kid(wsp, 'style');
  const fill = fillOf(spPr, style, 'fill'), ln = fillOf(spPr, style, 'ln'), lw = ln ? Math.max(0.25, (+(kid(spPr, 'ln') || { getAttribute: () => 9525 }).getAttribute('w') || 9525) / EMU) : 0;
  const W = w * PXPT, H = h * PXPT, S = lw * PXPT, ns = 'http://www.w3.org/2000/svg', svg = document.createElementNS(ns, 'svg');
  svg.setAttribute('width', Math.max(W, 1)); svg.setAttribute('height', Math.max(H, 1));
  const line = /line|Connector/i.test(geom), shp = document.createElementNS(ns, line ? 'line' : /ellipse/.test(geom) ? 'ellipse' : 'rect');
  if (line) { shp.setAttribute('x1', fh ? W : 0); shp.setAttribute('x2', fh ? 0 : W); shp.setAttribute('y1', fv ? H : 0); shp.setAttribute('y2', fv ? 0 : H); }
  else if (/ellipse/.test(geom)) { shp.setAttribute('cx', W / 2); shp.setAttribute('cy', H / 2); shp.setAttribute('rx', W / 2); shp.setAttribute('ry', H / 2); }
  else { shp.setAttribute('width', W); shp.setAttribute('height', H); if (/round/i.test(geom)) shp.setAttribute('rx', Math.min(W, H) / 6); }
  shp.setAttribute('fill', fill && !line ? fill[0] : 'none'); if (fill && fill[1] < 1) shp.setAttribute('fill-opacity', fill[1]);
  if (ln) { shp.setAttribute('stroke', ln[0]); shp.setAttribute('stroke-width', S); if (ln[1] < 1) shp.setAttribute('stroke-opacity', ln[1]); const dash = wa(kid(kid(spPr, 'ln'), 'prstDash'), 'val') || (kid(kid(spPr, 'ln'), 'prstDash') || { getAttribute: () => '' }).getAttribute('val'); if (dash && dash !== 'solid') shp.setAttribute('stroke-dasharray', /dot/i.test(dash) ? `${S} ${S * 2}` : `${S * 4} ${S * 3}`); }
  if (fill || ln) svg.appendChild(shp); e.appendChild(svg);
  const tx = first(wsp, 'txbxContent');
  if (tx) { // the text box's text, read-only, with the box's insets and vertical alignment
    const bp = kid(wsp, 'bodyPr'), ins = (n, d) => (bp && bp.getAttribute(n) != null ? +bp.getAttribute(n) : d) / EMU, c = document.createElement('span');
    c.className = 'dx-txc'; c.style.padding = `${ins('tIns', 45720)}pt ${ins('rIns', 91440)}pt ${ins('bIns', 45720)}pt ${ins('lIns', 91440)}pt`;
    const an = bp && bp.getAttribute('anchor'); c.style.justifyContent = an === 'ctr' ? 'center' : an === 'b' ? 'flex-end' : 'flex-start'; if (bp && bp.getAttribute('wrap') === 'none') c.style.whiteSpace = 'pre';
    roBlocks(tx, part, c, null); e.appendChild(c);
  }
  if (rot) e.style.transform = `rotate(${rot}deg)`; if (flip && !tx) svg.style.transform = flip;
  return true;
}
function roBlocks(xmlBox, part, out, anchors) { // a read-only copy (headers on every page, text boxes)
  for (const c of xmlBox.children) {
    const n = c.localName;
    if (n === 'p') out.appendChild(roP(c, part, anchors));
    else if (n === 'tbl') {
      const t = document.createElement('span'); t.className = 'dx-rt'; // (spans only: this can sit inside a paragraph of the page)
      for (const tr of kids(c, 'tr')) { const row = document.createElement('span'); row.className = 'dx-rr'; t.appendChild(row); for (const tc of kids(tr, 'tc')) { const td = document.createElement('span'); td.className = 'dx-rc'; row.appendChild(td); roBlocks(tc, part, td, anchors); } }
      out.appendChild(t);
    } else if (n === 'sdt') roBlocks(kid(c, 'sdtContent') || c, part, out, anchors);
    else if (n === 'customXml') roBlocks(c, part, out, anchors);
  }
  return out;
}
function roP(wp, part, anchors) {
  const pPr = kid(wp, 'pPr'), pf = readP(pPr), sid = pf.style || D.ST.defPara || '', e = mergeP(D.ST.defP, chainOf(sid, 'p'), pf), d = document.createElement('span');
  const mk = effR(readR(kid(pPr, 'rPr')), sid), ex = D.paged && e.sp && e.sp.lineRule === 'exact';
  d.className = 'dx-rp' + (ex ? ' dx-ex' : ''); d.style.cssText = cssP(e, mk, false, false, extrasP(pPr)) + (ex ? exactCss(e, mk) : '');
  const walk = el => {
    for (const c of el.children) {
      const n = c.localName;
      if (n === 'r') {
        const sp = document.createElement('span'); sp.style.cssText = cssR(effR(readR(kid(c, 'rPr')), sid));
        for (const q of c.children) {
          const m = q.localName;
          if (m === 't') sp.append(q.textContent); else if (m === 'tab') sp.append('\t'); else if ((m === 'br' || m === 'cr') && !/page|column/.test(wa(q, 'type') || '')) sp.appendChild(document.createElement('br'));
          else if (/^(drawing|pict|AlternateContent)$/.test(m)) { const a = document.createElement('span'); a.className = 'dx-a'; picture(a, q, part); if (anchors && first(q, 'anchor')) anchors.push({ el: a, o: ancInfo(q) }); else sp.appendChild(a); }
        }
        d.appendChild(sp);
      } else if (WRAPS.has(n)) walk(n === 'sdt' ? kid(c, 'sdtContent') || c : c);
    }
  };
  walk(wp); if (!d.textContent && !d.querySelector('.dx-a')) d.appendChild(document.createElement('br'));
  return d;
}
function hfLook(name) { // a header or footer part, drawn once and copied onto each page
  if (HFV.has(name)) return HFV.get(name);
  const part = D.parts[name]; let v = null;
  if (part) { const anchors = [], blocks = roBlocks(part.doc.documentElement, part, document.createElement('div'), anchors); v = { blocks, anchors, text: !!blocks.textContent.trim() }; }
  HFV.set(name, v); return v;
}
function sectGeo(sp) {
  const sz = kid(sp, 'pgSz'), m = kid(sp, 'pgMar'), g = (k, d) => wa(m, k) == null ? d : +wa(m, k), tp = kid(sp, 'titlePg');
  return { sp, w: +wa(sz, 'w') || 12240, h: +wa(sz, 'h') || 15840, t: Math.abs(g('top', 1440)), b: Math.abs(g('bottom', 1440)), l: g('left', 1440), r: g('right', 1440), hd: g('header', 720), ft: g('footer', 720), type: wa(kid(sp, 'type'), 'val') || 'nextPage', titlePg: !!tp && onoff(tp) };
}
function sectionsNow() { // the sections in order, each with its top-level blocks and its headers (a section without its own uses the one before)
  const out = []; let cur = [];
  for (const b of docEl.children) {
    cur.push(b);
    const p = b.tagName === 'P' ? b : b.tagName === 'DIV' && !isAtom(b) ? [...b.querySelectorAll('p')].pop() : null;
    const rec = p && p.classList.contains('dx-sect') && p.dataset.p && D.PP.get(p.dataset.p), sp = rec && rec.el && kid(rec.el, 'sectPr');
    if (sp) { out.push({ g: sectGeo(sp), blocks: cur }); cur = []; }
  }
  if (cur.length || !out.length) out.push({ g: sectGeo(D.sect), blocks: cur });
  const rels = relsOf(D.main), hf = {};
  for (const s of out) {
    for (const ref of [...kids(s.g.sp, 'headerReference'), ...kids(s.g.sp, 'footerReference')]) { const rel = rels[ref.getAttributeNS(NS.r, 'id')]; if (rel) hf[ref.localName[0] + (wa(ref, 'type') || 'default')] = resolvePath(D.main, rel.target); }
    s.hf = { ...hf };
  }
  return out;
}
const layCss = p => p._lx == null ? '' : `;margin-left:calc(${p._il || 0}pt + ${p._lx}px);margin-right:calc(${p._ir || 0}pt + ${p._ly}px${p._tol ? ' - 0.3% - 1px' : ''})` + (p._gap ? `;margin-top:${p._gap}px` : '');
function restyleLay(b) {
  if (b.tagName === 'P') { if (b._css0 == null) return; const css = b._css0 + layCss(b); b.style.cssText = css; b._css = css; b._attr = b.getAttribute('style'); return; }
  const ind = b.tagName === 'TABLE' ? +(b.dataset.ind || 0) * PXPT : 0;
  b.style.marginLeft = b._lx == null ? '' : (b._lx + ind) + 'px'; if (b.tagName !== 'TABLE') b.style.marginRight = b._ly == null ? '' : b._ly + 'px';
  b.style.marginTop = b._gap ? b._gap + 'px' : '';
}
function setLay(b, x, y) { if (b._lx != null && Math.abs(b._lx - x) < 0.3 && Math.abs(b._ly - y) < 0.3) return; b._lx = x; b._ly = y; restyleLay(b); }
function setGap(b, v) { v = Math.round(v * 100) / 100; if (Math.abs((b._gap || 0) - v) < 0.5) return; b._gap = v; restyleLay(b); }
function clearLay() { for (const b of docEl.children) { b._lx = b._ly = undefined; b._gap = 0; restyleLay(b); } }
const offIn = (el, root) => { let x = 0, y = 0; for (let n = el; n && n !== root; n = n.offsetParent) { x += n.offsetLeft; y += n.offsetTop; } return { x, y }; };
function layoutPaged(force) { // real pages: each section starts a page of its own size, a block that does not fit moves to the next page
  const key = docEl.childElementCount + ':' + docEl.offsetHeight + ':' + docEl.offsetWidth;
  if (!force && !D.layDirty && D.layKey === key && D.layPages) return pageStatus();
  D.layDirty = false; for (const g of [...docEl.querySelectorAll('tr.dx-gaprow')]) g.remove();
  const secs = sectionsNow(), Wpt = Math.max(...secs.map(s => s.g.w)) / 20, Wpx = Wpt * PXPT, T = v => v / 20 * PXPT, pages = [];
  if (D.paperW !== Wpt * 20) { D.paperW = Wpt * 20; paper.style.width = Wpt + 'pt'; }
  for (const s of secs) { s.left = (Wpt - s.g.w / 20) / 2 * PXPT; for (const b of s.blocks) setLay(b, s.left + T(s.g.l), Wpx - s.left - T(s.g.w) + T(s.g.r)); }
  for (const b of [...docEl.querySelectorAll('p')]) if (b.parentNode !== docEl && b._lx != null) { b._lx = b._ly = undefined; b._gap = 0; restyleLay(b); } // (moved into a table or box)
  const newPage = (s, prev) => {
    if (prev) prev.hpx = Math.max(prev.hpx, prev.bottom - prev.top + T(prev.g.b));
    const pg = { s, g: s.g, left: s.left, top: prev ? prev.top + prev.hpx + GAP : 0, hpx: T(s.g.h), wpx: T(s.g.w), firstOfSec: !prev || prev.s !== s };
    pg.bottom = pg.top; pages.push(pg); return pg;
  };
  let pg = null;
  for (const s of secs) {
    const g = s.g; let fresh = false;
    if (!pg || g.type !== 'continuous') { pg = newPage(s, pg); fresh = true; }
    let brk = false;
    for (const b of s.blocks) {
      const top = b.offsetTop - (b._gap || 0), h = b.offsetHeight, cTop = pg.top + T(pg.g.t), cBot = pg.top + T(pg.g.h) - T(pg.g.b);
      const rows = b.tagName === 'TABLE' ? [...b.rows] : [], split = rows.length > 1, r0 = split ? rows[0].offsetTop + rows[0].offsetHeight : h; // (a table breaks between its rows, as in Word)
      if (!fresh && (brk || b.classList.contains('dx-pbb') || (top + h > cBot + 1 && (split ? top + r0 > cBot + 1 && r0 <= cBot - cTop : h <= cBot - cTop)))) { pg = newPage(s, pg); fresh = true; }
      const want = fresh ? Math.max(0, pg.top + T(pg.g.t) - top) : 0;
      setGap(b, want); fresh = false; brk = !!b.querySelector('.dx-pb');
      if (split && top + want + h > pg.top + T(pg.g.h) - T(pg.g.b) + 1) pg = splitRows(b, rows, s, pg, newPage, T);
      pg.bottom = Math.max(pg.bottom, b.offsetTop + b.offsetHeight);
    }
  }
  if (pg) pg.hpx = Math.max(pg.hpx, pg.bottom - pg.top + T(pg.g.b));
  // the pages, their headers and footers (drawn behind the text)
  const hfName = (pg, k) => { const s = pg.s.hf; return (pg.firstOfSec && pg.g.titlePg ? s[k + 'first'] : null) || s[k + 'default'] || null; };
  const liveHf = (k, name) => { const el = k === 'h' ? hdrEl : ftrEl, part = k === 'h' ? D.hdrPart : D.ftrPart; if (el.hidden || !part || part.name !== name) return null; return { el, page: pages.find(q => hfName(q, k) === name) }; };
  const sig = pages.map(p => [p.left, p.top, p.wpx, p.hpx, hfName(p, 'h'), hfName(p, 'f'), p.g.hd, p.g.ft, p.g.l, p.g.r].join(',')).join('|') + D.mv + (hdrEl.hidden ? '' : '|h' + hdrEl.innerHTML.length + hdrEl.textContent) + (ftrEl.hidden ? '' : '|f' + ftrEl.innerHTML.length + ftrEl.textContent);
  if (frames._sig !== sig) {
    frames._sig = sig; frames.replaceChildren();
    for (const p of pages) {
      const f = document.createElement('div'), G = geoPt(p.g); f.className = 'dx-frame'; Object.assign(f.style, { left: p.left + 'px', top: p.top + 'px', width: p.wpx + 'px', height: p.hpx + 'px' });
      for (const k of ['h', 'f']) {
        const name = hfName(p, k), v = name && hfLook(name); if (!v) continue;
        const live = liveHf(k, name), c = live ? (live.page === p ? null : hfClone(live.el)) : v.blocks.cloneNode(true); // (a header being edited: its live text, except on the page where it is edited)
        if (c) { c.className = 'dx-hfv'; c.style.left = G.l + 'pt'; c.style.width = Math.max(10, G.w - G.l - G.r) + 'pt'; if (k === 'h') c.style.top = (p.g.hd / 20) + 'pt'; else c.style.bottom = (p.g.ft / 20) + 'pt'; f.appendChild(c); }
        for (const a of [...v.anchors].sort((x, y) => (y.o && y.o.behind ? 1 : 0) - (x.o && x.o.behind ? 1 : 0))) { if (!a.o) continue; const q = ancPlace(a.o, G, k === 'h' ? p.g.hd / 20 : G.h - p.g.ft / 20), el = a.el.cloneNode(true); el.style.left = q.x + 'pt'; el.style.top = q.y + 'pt'; el.style.zIndex = ''; f.appendChild(el); }
      }
      frames.appendChild(f);
    }
  }
  // drawings anchored in the body: on the page of their paragraph
  for (const a of docEl.querySelectorAll('.dx-anc')) {
    let o = ANC.get(a.dataset.k); if (o === undefined) { const rec = D.AT.get(a.dataset.k); o = rec ? ancInfo(rec.el) : null; ANC.set(a.dataset.k, o); } if (!o) continue;
    const para = a.closest('p') || a.parentNode, py = offIn(para, paper).y; let i = 0; while (i + 1 < pages.length && pages[i + 1].top <= py + 0.5) i++;
    const P = pages[i]; if (!P) continue; const q = ancPlace(o, geoPt(P.g), (py - P.top) / PXPT), op = a.offsetParent, base = op && op !== paper ? offIn(op, paper) : { x: 0, y: 0 };
    a.style.left = (P.left + q.x * PXPT - base.x) + 'px'; a.style.top = (P.top + q.y * PXPT - base.y) + 'px';
  }
  for (const [el, k] of [[hdrEl, 'h'], [ftrEl, 'f']]) { // a header or footer being edited sits on the first page that shows it
    if (el.hidden) continue; const part = k === 'h' ? D.hdrPart : D.ftrPart, p = pages.find(q => part && hfName(q, k) === part.name) || pages[0]; if (!p) continue; const G = geoPt(p.g);
    Object.assign(el.style, { position: 'absolute', zIndex: '3', left: (p.left + G.l * PXPT) + 'px', width: Math.max(10, G.w - G.l - G.r) * PXPT + 'px', top: (k === 'h' ? p.top + p.g.hd / 20 * PXPT : p.top + p.hpx - p.g.ft / 20 * PXPT - el.offsetHeight) + 'px' });
  }
  const last = pages[pages.length - 1]; paper.style.minHeight = last ? (last.top + last.hpx) + 'px' : '';
  D.layPages = pages; D.pages = pages.length; D.layKey = docEl.childElementCount + ':' + docEl.offsetHeight + ':' + docEl.offsetWidth;
  pageStatus();
}
function splitRows(b, rows, s, pg, newPage, T) { // the rows that do not fit go to the next page: a spacer row (never saved) fills the rest of the page
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i], cTop = pg.top + T(pg.g.t), cBot = pg.top + T(pg.g.h) - T(pg.g.b), rt = b.offsetTop + r.offsetTop;
    if (rt + r.offsetHeight <= cBot + 1 || rt <= cTop + 1) continue;
    pg.bottom = Math.max(pg.bottom, rt); pg = newPage(s, pg);
    const gap = pg.top + T(pg.g.t) - rt, gr = document.createElement('tr'), th = document.createElement('th');
    gr.className = 'dx-gaprow'; gr.contentEditable = 'false'; gr.setAttribute('aria-hidden', 'true'); th.colSpan = Math.max(1, ...rows.map(q => q.cells.length));
    th.style.height = Math.max(0, gap) + 'px'; gr.appendChild(th); r.before(gr);
    const d = pg.top + T(pg.g.t) - (b.offsetTop + r.offsetTop); if (Math.abs(d) > 0.5) th.style.height = Math.max(0, gap + d) + 'px';
  }
  return pg;
}
function hfClone(el) { // the live header as a read-only copy (spans only)
  const c = document.createElement('span');
  for (const b of el.children) { const q = b.cloneNode(true); for (const a of [...q.querySelectorAll('.dx-anc')]) a.remove(); const s = q.tagName === 'P' ? document.createElement('span') : q; if (s !== q) { s.className = 'dx-rp ' + q.className; s.setAttribute('style', q.getAttribute('style') || ''); s.append(...q.childNodes); } c.appendChild(s); }
  for (const x of c.querySelectorAll('[contenteditable]')) x.removeAttribute('contenteditable'); return c;
}
function pageStatus() {
  const P = D.layPages || [], st = (scroller.scrollTop + scroller.clientHeight / 3 - paper.offsetTop * D.zoom) / D.zoom; let cur = 1;
  for (let i = 1; i < P.length; i++) if (P[i].top < st) cur = i + 1;
  $('#dxPg').textContent = `Page ${cur} of ${P.length || 1}`;
}
async function layoutNow() { if (document.fonts && document.fonts.ready) await document.fonts.ready; D.layDirty = true; paginate(); }

/* ───────── tab stops: each tab gets the width that takes it to its stop (the paragraph's and its style's stops, then the default grid) ───────── */
function tabStops(p) { // → stops in pt from the text column's left edge (null: the paragraph has none of its own)
  const rec = p.dataset.p && D.PP.get(p.dataset.p), sid = (p._pe && p._pe.style) || pfOf(p).style || D.ST.defPara || '', key = 'tabs|' + (p.dataset.p || '') + '|' + sid;
  if (D.cc.has(key)) return D.cc.get(key);
  const m = new Map(), add = tabs => { for (const t of kids(tabs, 'tab')) { const pos = +wa(t, 'pos'), v = wa(t, 'val') || 'left'; if (isNaN(pos)) continue; if (v === 'clear') m.delete(pos); else m.set(pos, { pos: pos / 20, val: v, leader: wa(t, 'leader') }); } };
  const chain = []; for (let id = sid, n = 0; id && n < 15; n++) { const st = D.ST.byId[id]; if (!st) break; chain.unshift(st); id = st.based; }
  for (const st of chain) add(kid(kid(st.el, 'pPr'), 'tabs'));
  if (rec && rec.el) add(kid(rec.el, 'tabs'));
  const v = m.size ? [...m.values()].sort((a, b) => a.pos - b.pos) : null; D.cc.set(key, v); return v;
}
function wrapTabs(p) { // every tab character alone in a span.dx-tab (only for the look: saving and editing see the tab character)
  for (const t of [...p.querySelectorAll('span.dx-tab')]) if (t.textContent !== '\t' || t.childNodes.length !== 1) t.replaceWith(...t.childNodes);
  const w = document.createTreeWalker(p, NodeFilter.SHOW_TEXT), list = []; for (let n; (n = w.nextNode());) if (n.data.indexOf('\t') >= 0 && !(n.parentNode.classList && n.parentNode.classList.contains('dx-tab'))) list.push(n);
  for (let n of list) {
    for (let i; (i = n.data.indexOf('\t')) >= 0;) {
      if (i > 0) n = n.splitText(i); const rest = n.length > 1 ? n.splitText(1) : null;
      const sp = document.createElement('span'); sp.className = 'dx-tab'; sp.contentEditable = 'false'; n.replaceWith(sp); sp.appendChild(n);
      if (!rest) break; n = rest;
    }
  }
}
function layTabs(force) {
  const P = [...docEl.querySelectorAll('p'), ...(hdrEl.hidden ? [] : hdrEl.querySelectorAll('p')), ...(ftrEl.hidden ? [] : ftrEl.querySelectorAll('p'))];
  let keep = null;
  for (const p of P) {
    const txt = p.textContent; if (txt.indexOf('\t') < 0) continue;
    const stops = tabStops(p); if (!stops && !D.paged && !p.querySelector('.dx-tab')) continue;
    const sig = txt + '|' + p.innerHTML.length + '|' + p.offsetWidth + '|' + (p._css || '').length; if (!force && p._tsig === sig) continue;
    if ([...p.querySelectorAll('span.dx-tab')].length !== (txt.match(/\t/g) || []).length || p.querySelector('span.dx-tab span, span.dx-tab:empty')) {
      if (keep === null) { const s = getSelection(); keep = s.rangeCount && rootOf(s.getRangeAt(0).startContainer) ? getSel() : false; }
      wrapTabs(p);
    }
    const tabs = [...p.querySelectorAll('span.dx-tab')]; for (const t of tabs) { t.style.width = '0'; t.style.borderBottom = ''; }
    const pr = p.getBoundingClientRect(), k = pr.width / (p.offsetWidth || 1) || 1, cs = getComputedStyle(p), e = p._pe || effP(p), ind = e.ind || {};
    const il = (+ind.left || 0) / 20, x0 = pr.left + (p.clientLeft + parseFloat(cs.paddingLeft || 0)) * k, wPt = (p.clientWidth - parseFloat(cs.paddingLeft || 0) - parseFloat(cs.paddingRight || 0)) * 0.75 + il, dt = D.defTab || 36;
    const all = (stops || []).slice(); if (+ind.hang && !all.some(s => Math.abs(s.pos - il) < 0.5)) { all.push({ pos: il, val: 'left' }); all.sort((a, b) => a.pos - b.pos); }
    const pt = px => (px - x0) / k * 0.75 + il;
    tabs.forEach((t, i) => {
      const r = t.getBoundingClientRect(), x = pt(r.left); let st = all.find(s => s.pos > x + 0.3);
      if (!st) st = { pos: (Math.floor((x + 0.3) / dt) + 1) * dt, val: 'left' };
      const stop = st.val === 'left' || st.val === 'start' || st.val === 'bar' ? st.pos : Math.min(st.pos, wPt - 0.5); let w = stop - x; // (a right tab past the margin aligns at the margin)
      if (st.val !== 'left' && st.val !== 'start' && st.val !== 'bar') { // the text up to the next tab sits before (right), around (center) or at its decimal point
        const rg = document.createRange(); rg.setStartAfter(t); if (tabs[i + 1]) rg.setEndBefore(tabs[i + 1]); else rg.setEnd(p, p.childNodes.length);
        if (st.val === 'decimal') { const tx = rg.toString(), d = tx.search(/[.,]\d/); if (d >= 0) { const u = document.createRange(), w2 = document.createTreeWalker(p, NodeFilter.SHOW_TEXT); let left = d; u.setStart(rg.startContainer, rg.startOffset); for (let n; (n = w2.nextNode());) { if (!rg.intersectsNode(n) || n.parentNode === t) continue; const a = n === rg.startContainer ? rg.startOffset : 0; if (n.length - a > left) { u.setEnd(n, a + left); break; } left -= n.length - a; } rg.setEnd(u.endContainer, u.endOffset); } }
        let right = r.right; for (const q of rg.getClientRects()) if (q.width && q.top - 2 <= r.top && q.bottom + 2 >= r.top) right = Math.max(right, q.right); // (the same line only)
        const tw = (right - r.right) / k * 0.75; w -= st.val === 'center' ? tw / 2 : tw;
      }
      w = Math.max(0, Math.min(w, wPt - x)); t.style.width = w.toFixed(2) + 'pt';
      if (st.leader && st.leader !== 'none' && w > 2) t.style.borderBottom = `${st.leader === 'heavy' ? 2 : 1}px ${st.leader === 'dot' || st.leader === 'middleDot' ? 'dotted' : st.leader === 'hyphen' ? 'dashed' : 'solid'} currentColor`;
    });
    p._tsig = txt + '|' + p.innerHTML.length + '|' + p.offsetWidth + '|' + (p._css || '').length;
  }
  if (keep) setSel(keep);
}

/* ───────── fonts embedded in the file (word/fonts/*.odttf) ───────── */
function deobfuscate(u, key) { // the first 32 bytes are XOR-ed with the font key (the GUID's bytes, last first)
  const magic = b => { const t = (b[0] << 24 | b[1] << 16 | b[2] << 8 | b[3]) >>> 0; return t === 0x00010000 || t === 0x4f54544f || t === 0x74727565 || t === 0x74746366; };
  if (magic(u)) return u.buffer.slice(u.byteOffset, u.byteOffset + u.byteLength);
  const hex = String(key || '').replace(/[^0-9a-f]/gi, ''); if (hex.length !== 32) return null;
  const kb = []; for (let i = 0; i < 16; i++) kb.push(parseInt(hex.substr(i * 2, 2), 16));
  for (const k of [kb.slice().reverse(), kb]) { const b = u.slice(); for (let i = 0; i < 32 && i < b.length; i++) b[i] ^= k[i % 16]; if (magic(b)) return b.buffer; }
  return null;
}
async function loadEmbeddedFonts() {
  for (const f of D.embFonts || []) try { document.fonts.delete(f); } catch (e) { /* */ }
  D.embFonts = []; D.embNames = new Set();
  const rel = Object.values(relsOf(D.main)).find(r => /\/fontTable$/.test(r.type)); if (!rel || typeof FontFace === 'undefined') return;
  try {
    const path = resolvePath(D.main, rel.target), fz = D.zip.file(path); if (!fz) return;
    const doc = parseXml(await fz.async('string')), rf = D.zip.file(path.replace(/([^/]+)$/, '_rels/$1.rels')), rels = rf ? relsOf({ rels: parseXml(await rf.async('string')) }) : {}, dir = path.replace(/[^/]+$/, '');
    for (const font of kids(doc.documentElement, 'font')) {
      const name = wa(font, 'name'); if (!name) continue;
      for (const [tag, weight, style] of [['embedRegular', '400', 'normal'], ['embedBold', '700', 'normal'], ['embedItalic', '400', 'italic'], ['embedBoldItalic', '700', 'italic']]) {
        const el = kid(font, tag), r = el && rels[el.getAttributeNS(NS.r, 'id')], file = r && D.zip.file(resolvePath({ dir }, r.target)); if (!file) continue;
        const buf = deobfuscate(await file.async('uint8array'), wa(el, 'fontKey')); if (!buf) continue;
        try { const ff = new FontFace(name, buf, { weight, style }); await ff.load(); document.fonts.add(ff); D.embFonts.push(ff); D.embNames.add(name); } catch (e) { /* a font the browser cannot read: the look-alike is used */ }
      }
    }
  } catch (e) { /* no embedded fonts then */ }
  stacks.clear(); RCSS.clear();
}

/* ───────── Make flowing: an exact-layout document → normal text that reflows (one undo step) ───────── */
function inlineXml(x) { // a run with an anchored picture → the same picture inline
  const r = x.cloneNode(true), a = first(r, 'anchor'); if (!a) return r;
  const inl = r.ownerDocument.createElementNS(WPNS, 'wp:inline'); for (const k of ['distT', 'distB', 'distL', 'distR']) inl.setAttribute(k, '0');
  for (const n of ['extent', 'effectExtent', 'docPr', 'cNvGraphicFramePr', 'graphic']) { const c = [...a.children].find(q => q.localName === n); if (c) inl.appendChild(c); }
  a.replaceWith(inl); return r;
}
function retarget(x, xPt, dyPt, onPage, hRel) { // an anchored drawing → x from the page (or column) edge, and y below the top of its paragraph (it then moves with the text)
  const r = x.cloneNode(true), a = first(r, 'anchor'); if (!a) return r; a.setAttribute('simplePos', '0');
  for (const [n, rel, v] of [['positionH', hRel || 'page', xPt], ['positionV', onPage ? 'page' : 'paragraph', dyPt]]) {
    let e = kid(a, n); if (!e) continue; e.setAttribute('relativeFrom', rel); e.replaceChildren();
    const o = r.ownerDocument.createElementNS(WPNS, 'wp:posOffset'); o.textContent = String(Math.round(v * EMU)); e.appendChild(o);
  }
  return r;
}
async function makeFlowing() { // (a drawing-heavy page can take a few seconds)
  if (!D.zip) return; let r; await busy(async () => { await new Promise(res => setTimeout(res, 30)); r = await flowNow(); }, 'Making the text flow…'); return r;
}
const luma = h => { const n = parseInt(h, 16); return ((n >> 16 & 255) * 0.299 + (n >> 8 & 255) * 0.587 + (n & 255) * 0.114) / 255; };
function hexOf(c, a) { // a css colour (and opacity) → RRGGBB as it looks on white paper
  let h = /^#?[0-9a-f]{6}$/i.test(c || '') ? c.replace('#', '') : null;
  if (!h) { const g = document.createElement('canvas').getContext('2d'); g.fillStyle = '#7f7f7f'; g.fillStyle = c || ''; h = /^#[0-9a-f]{6}$/i.test(g.fillStyle) ? g.fillStyle.slice(1) : '7F7F7F'; }
  if (a != null && a < 1) { const n = parseInt(h, 16); h = [n >> 16 & 255, n >> 8 & 255, n & 255].map(x => Math.round(x * a + 255 * (1 - a)).toString(16).padStart(2, '0')).join(''); }
  return h.toUpperCase();
}
function drawKind(x) { // what an anchored drawing is: a picture, a line, a box, a text box…
  const wsp = first(x, 'wsp'), blip = first(x, 'blip');
  if (!wsp) return { kind: blip && !first(x, 'wgp') && !first(x, 'chart') ? 'pic' : 'other' };
  const spPr = first(wsp, 'spPr'), style = kid(wsp, 'style'), g = kid(spPr, 'prstGeom'), geom = g ? g.getAttribute('prst') || 'rect' : kid(spPr, 'custGeom') ? 'cust' : 'rect';
  const fill = fillOf(spPr, style, 'fill'), ln = fillOf(spPr, style, 'ln'), lw = ln ? (+(kid(spPr, 'ln') || { getAttribute: () => 9525 }).getAttribute('w') || 9525) / EMU : 0;
  return { kind: first(wsp, 'txbxContent') ? 'text' : /line|Connector/i.test(geom) ? 'line' : /rect|Rect|flowChartProcess/.test(geom) ? 'rect' : 'shape', fill, ln, lw };
}
function rasterize(list, bb) { // drawings (pictures, boxes, circles, lines) → one PNG (null: something cannot be drawn)
  const S = Math.min(2, 4000 / Math.max(bb.w, bb.h, 1)), cv = document.createElement('canvas'); cv.width = Math.max(1, Math.ceil(bb.w * S)); cv.height = Math.max(1, Math.ceil(bb.h * S));
  const g = cv.getContext('2d'); g.scale(S, S);
  for (const it of list) {
    const v = it.view, w = it.r.w, h = it.r.h; if (!v || v.querySelector('.dx-txc')) return null;
    g.save(); g.translate(it.r.x - bb.x + w / 2, it.r.y - bb.y + h / 2); const rot = /rotate\(([-\d.]+)deg\)/.exec(v.style.transform || ''); if (rot) g.rotate(+rot[1] * Math.PI / 180); g.translate(-w / 2, -h / 2);
    const im = v.querySelector('img'), shp = v.querySelector('svg > *');
    if (im) { if (!im.complete || !im.naturalWidth) { g.restore(); return null; } if (/scaleX\(-1\)/.test(v.style.transform)) { g.translate(w, 0); g.scale(-1, 1); } g.drawImage(im, 0, 0, w, h); }
    else if (shp) {
      const A = n => shp.getAttribute(n), N = n => +A(n) || 0; g.beginPath();
      if (shp.tagName === 'line') { g.moveTo(N('x1'), N('y1')); g.lineTo(N('x2'), N('y2')); }
      else if (shp.tagName === 'ellipse') g.ellipse(N('cx'), N('cy'), Math.max(0, N('rx')), Math.max(0, N('ry')), 0, 0, Math.PI * 2);
      else { const rx = Math.min(N('rx'), N('width') / 2, N('height') / 2); if (rx > 0 && g.roundRect) g.roundRect(0, 0, N('width'), N('height'), rx); else g.rect(0, 0, N('width'), N('height')); }
      if (A('fill') && A('fill') !== 'none') { g.globalAlpha = A('fill-opacity') ? +A('fill-opacity') : 1; g.fillStyle = A('fill'); g.fill(); }
      if (A('stroke')) { g.globalAlpha = A('stroke-opacity') ? +A('stroke-opacity') : 1; g.strokeStyle = A('stroke'); g.lineWidth = N('stroke-width') || 1; const d = A('stroke-dasharray'); if (d) g.setLineDash(d.split(/[ ,]+/).map(Number)); g.stroke(); }
    }
    g.restore();
  }
  return cv.toDataURL('image/png');
}
function pngRun(url, wPt, hPt, ids, at) { // a PNG → a new picture run in the main part: inline, or (at: {x, y} in pt from the page corner) behind the text
  const bin = atob(url.split(',')[1]), u = new Uint8Array(bin.length); for (let i = 0; i < bin.length; i++) u[i] = bin.charCodeAt(i);
  let n = 1; while (D.zip.file(`word/media/xd3flow${n}.png`)) n++; const path = `word/media/xd3flow${n}.png`; D.zip.file(path, u, { createFolders: false });
  ensureType({ ext: 'png', type: 'image/png' }); const rid = addRel(D.main, 'image', resolveRel(D.main, path)); D.media.set(D.main.name + '#' + rid, URL.createObjectURL(new Blob([u], { type: 'image/png' })));
  const id = ids.next++, cx = Math.round(wPt * EMU), cy = Math.round(hPt * EMU), E = v => Math.round(v * EMU);
  const pic = `<wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>${at ? '<wp:wrapNone/>' : ''}<wp:docPr id="${id}" name="Drawing ${id}"/><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="${id}" name="Drawing ${id}.png"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic>`;
  const dr = at ? `<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="${id}" behindDoc="1" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="page"><wp:posOffset>${E(at.x)}</wp:posOffset></wp:positionH><wp:positionV relativeFrom="page"><wp:posOffset>${E(at.y)}</wp:posOffset></wp:positionV>${pic}</wp:anchor>` : `<wp:inline distT="0" distB="0" distL="0" distR="0">${pic}</wp:inline>`;
  return frag(D.main.doc, `<w:r><w:drawing>${dr}</w:drawing></w:r>`);
}
const inlinePng = (url, wPt, hPt, ids) => atomEl(pngRun(url, wPt, hPt, ids), D.main);
function packGroups(list, pg, ids, decode) { // a header drawn in hundreds of pieces (strips of a gradient, boxes, circles) → one picture for each group of pieces that touch
  const par = list.map((_, i) => i), f = i => par[i] === i ? i : (par[i] = f(par[i])), R = list.map(it => it.r);
  const ord = list.map((_, i) => i).sort((a, b) => R[a].x - R[b].x); // (a sweep along x: only neighbours are compared)
  for (let a = 0; a < ord.length; a++) { const i = ord[a], ri = R[i]; for (let b = a + 1; b < ord.length; b++) { const j = ord[b], rj = R[j]; if (rj.x > ri.x + ri.w + 2) break; if (rj.y <= ri.y + ri.h + 2 && ri.y <= rj.y + rj.h + 2) par[f(i)] = f(j); } }
  const gs = new Map(); list.forEach((it, i) => { const g = f(i); if (!gs.has(g)) gs.set(g, []); gs.get(g).push(it); });
  const out = [];
  for (const g of gs.values()) {
    if (g.length < 2) { out.push(g[0]); continue; }
    const x0 = Math.min(...g.map(i => i.r.x)), y0 = Math.min(...g.map(i => i.r.y)), bb = { x: x0, y: y0, w: Math.max(...g.map(i => i.r.x + i.r.w)) - x0, h: Math.max(...g.map(i => i.r.y + i.r.h)) - y0 };
    if (g.reduce((a, i) => a + i.r.w * i.r.h, 0) < 0.8 * bb.w * bb.h) { out.push(...g); continue; } // (only pieces that fill their area: a frame of lines stays lines)
    const url = rasterize(g, bb); if (!url) { out.push(...g); continue; }
    const view = document.createElement('span'), im = document.createElement('img'); view.className = 'dx-a'; im.src = url; im.style.width = bb.w + 'px'; im.style.height = bb.h + 'px'; view.appendChild(im); decode.push(im.decode().catch(() => null));
    const wPt = bb.w / PXPT, hPt = bb.h / PXPT, at = { x: (bb.x - pg.left) / PXPT, y: (bb.y - pg.top) / PXPT }; let xml = null;
    const it = { part: D.main, o: { hRel: 'page', vRel: 'page', hOff: at.x, vOff: at.y, hAl: null, vAl: null, w: wPt, h: hPt, behind: true, pic: true }, r: bb, el: null, view, kind: 'pic', vec: !g.some(i => i.kind === 'pic') };
    Object.defineProperty(it, 'x', { get: () => xml || (xml = pngRun(url, wPt, hPt, ids, at)) }); // (the picture is added to the file only when it is kept)
    out.push(it);
  }
  return out;
}
function picTone(it, mode) { // a picture's average colour, and whether it is flat (one colour or a soft gradient)
  try {
    const im = it.view && it.view.querySelector('img'); if (!im || !im.naturalWidth) return null; const cv = document.createElement('canvas'); cv.width = cv.height = 16; const g = cv.getContext('2d'); g.drawImage(im, 0, 0, 16, 16);
    const d = g.getImageData(0, 0, 16, 16).data, m = [0, 0, 0], q = [0, 0, 0]; let n = 0;
    for (let i = 0; i < d.length; i += 4) { if (d[i + 3] < 128) continue; n++; for (let c = 0; c < 3; c++) { m[c] += d[i + c]; q[c] += d[i + c] * d[i + c]; } }
    if (n < 40) return null; const mean = m.map(v => v / n), sd = Math.sqrt(q.reduce((a, v, c) => a + v / n - mean[c] * mean[c], 0) / 3), hx = v => v.map(c => Math.round(c).toString(16).padStart(2, '0')).join('').toUpperCase();
    if (sd >= 45 && mode) { // (a box drawn with a few small marks on it: the colour of most of the box)
      const B = new Map(); for (let i = 0; i < d.length; i += 4) if (d[i + 3] >= 128) { const k = (d[i] >> 4) << 8 | (d[i + 1] >> 4) << 4 | d[i + 2] >> 4, b = B.get(k) || [0, 0, 0, 0]; b[0]++; b[1] += d[i]; b[2] += d[i + 1]; b[3] += d[i + 2]; B.set(k, b); }
      const top = [...B.values()].sort((a, b) => b[0] - a[0])[0]; if (top && top[0] >= n * 0.6) return { hex: hx([top[1] / top[0], top[2] / top[0], top[3] / top[0]]), flat: true };
    }
    return { hex: hx(mean), flat: sd < 45 };
  } catch (e) { return null; }
}
function shadeCell(td, hex) { // cell shading (a new record: Undo keeps the old one)
  const rec = D.TB.get(td.dataset.tc); if (!rec) return; const tc = rec.el.cloneNode(true), doc = tc.ownerDocument;
  let pr = kid(tc, 'tcPr'); if (!pr) { pr = wel(doc, 'tcPr'); tc.insertBefore(pr, tc.firstChild); } for (let e; (e = kid(pr, 'shd'));) e.remove();
  insertOrdered(pr, wel(doc, 'shd', { val: 'clear', color: 'auto', fill: hex }), TC_ORDER);
  const id = nid(); D.TB.set(id, { ...rec, el: tc }); td.dataset.tc = id; td.style.background = '#' + hex;
}
const cellShd = td => { const rec = td && D.TB.get(td.dataset.tc), f = rec && wa(kid(kid(rec.el, 'tcPr'), 'shd'), 'fill'); return f && /^[0-9a-f]{6}$/i.test(f) && f.toUpperCase() !== 'FFFFFF' ? f.toUpperCase() : null; };
function colText(t) { // a borderless table whose columns are running text (lines that continue down a column), not a table of data
  const rows = [...t.rows].filter(r => !r.classList.contains('dx-gaprow')); if (t.classList.contains('dx-tb') || t.closest('.dx-a') || !rows.length) return false;
  const n = rows[0].cells.length; if (n < 2 || rows.some(r => r.cells.length !== n || [...r.cells].some(c => c.colSpan > 1 || c.classList.contains('dx-vm')))) return false;
  if (rows.length === 1) return true;
  let pairs = 0, cont = 0;
  for (let c = 0; c < n; c++) for (let i = 0; i + 1 < rows.length; i++) {
    const a = rows[i].cells[c].innerText.trim(), b = rows[i + 1].cells[c].innerText.trim(); if (!a || !b) continue; pairs++;
    if (!/[.!?:;]["'”’)]?$/.test(a) && /^\p{Ll}/u.test(b)) cont++;
  }
  return pairs >= 2 && cont >= pairs * 0.4;
}
async function flowNow() {
  if (D.paged) layoutPaged(true);
  const secs = sectionsNow(), s0 = secs[0].g, W0 = s0.w / 20, P = D.layPages || [];
  const need = secs.length > 1 || [...docEl.querySelectorAll('p')].some(p => { const pf = pfOf(p), sp = pf.sp || {}; return sp.lineRule === 'exact' || sp.lineRule === 'atLeast' || (pf.ind && pf.ind.right); }) || [...docEl.querySelectorAll('span[data-rf]')].some(r => rfOf(r).spc != null || rfOf(r).scale != null)
    || [...docEl.querySelectorAll('table.dx-t')].some(colText) || docEl.querySelector('.dx-anc');
  if (!need) return toast('This document already flows normally', 'info');
  snapshot(); D.mv++;
  const pr = paper.getBoundingClientRect(), k = pr.width / (paper.offsetWidth || 1) || 1, boxOf = q => ({ x: (q.left - pr.left) / k, y: (q.top - pr.top) / k, w: q.width / k, h: q.height / k });
  const ov = (a, b) => Math.max(0, Math.min(a.x + a.w, b.x + b.w) - Math.max(a.x, b.x)) * Math.max(0, Math.min(a.y + a.h, b.y + b.h) - Math.max(a.y, b.y));
  const frac = (t, r) => ov(t, r) / Math.max(1, t.w * t.h), pageAt = y => { let i = 0; for (let n = 1; n < P.length; n++) if (P[n].top <= y + 0.5) i = n; return P[i]; };
  // 1 where the text is: each paragraph's ink box (as laid out now)
  const T = [], rg = document.createRange();
  for (const p of docEl.querySelectorAll('p')) {
    let L = 1e9, U = 1e9, R = -1e9, B = -1e9; const tw = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
    for (let n; (n = tw.nextNode());) { if (!n.data.trim() || n.parentNode.closest('.dx-a')) continue; rg.selectNodeContents(n); for (const q of rg.getClientRects()) if (q.width) { L = Math.min(L, q.left); U = Math.min(U, q.top); R = Math.max(R, q.right); B = Math.max(B, q.bottom); } }
    if (R > L) { const b = boxOf({ left: L, top: U, width: R - L, height: B - U }); const o = offIn(p, paper), ind = +((pfOf(p).ind || {}).left || 0); T.push({ p, ...b, r: b.x + b.w, b: b.y + b.h, py: o.y, px: o.x + (parseFloat(getComputedStyle(p).paddingLeft) || 0), pw: p.clientWidth, nind: ind >= 720 ? ind / 20 : 0 }); }
  }
  // 2 every anchored drawing: in the body, or in a header that holds only drawings (a page made of thousands of them stays a header)
  const items = [], moved = new Set(), mrels0 = relsOf(D.main), decode = [], ids = { next: 1 + Math.max(0, ...[...docPrIds()]) };
  for (const a of docEl.querySelectorAll('.dx-anc')) { const rec = D.AT.get(a.dataset.k), o = rec && ancInfo(rec.el); if (!o) continue; const b = offIn(a, paper); items.push({ x: rec.el, part: D.parts[rec.part] || D.main, o, r: { x: b.x, y: b.y, w: o.w * PXPT, h: o.h * PXPT }, el: a, view: a, ...drawKind(rec.el) }); }
  for (const sc of secs) {
    const pg = P.find(p => p.s && p.s.blocks[0] === sc.blocks[0]) || P[0]; if (!pg) continue;
    for (const ref of kids(sc.g.sp, 'headerReference')) {
      const rel = mrels0[ref.getAttributeNS(NS.r, 'id')], name = rel && resolvePath(D.main, rel.target), part = name && D.parts[name], v = part && hfLook(name); if (!v || v.text || moved.has(name)) continue;
      const ds = [...part.doc.getElementsByTagNameNS('*', 'drawing'), ...part.doc.getElementsByTagNameNS('*', 'AlternateContent')]; if (ds.length > 8000) continue; moved.add(name);
      const mine = [];
      for (const d of ds) {
        const r = d.parentNode; if (!r || r.localName !== 'r' || (d.localName === 'drawing' && d.closest('AlternateContent')) || d.closest('txbxContent')) continue;
        const nr = r.cloneNode(false), rp = kid(r, 'rPr'); if (rp) nr.appendChild(rp.cloneNode(true)); nr.appendChild(d.cloneNode(true));
        const o = ancInfo(nr); if (!o) continue; const q = ancPlace(o, geoPt(sc.g), sc.g.hd / 20), view = document.createElement('span'); view.className = 'dx-a'; picture(view, d, part);
        for (const im of view.querySelectorAll('img')) decode.push(im.decode().catch(() => null));
        mine.push({ x: nr, part, o: { ...o, behind: true }, r: { x: pg.left + q.x * PXPT, y: pg.top + q.y * PXPT, w: o.w * PXPT, h: o.h * PXPT }, el: null, view, ...drawKind(nr) }); // (a header is always behind the text)
      }
      if (mine.length > 120) await Promise.all(decode); items.push(...(mine.length > 120 ? packGroups(mine, pg, ids, decode) : mine));
    }
  }
  await Promise.all(decode);
  // 3 a box behind text becomes the shading (and border) of that text; a rule just under or over a line becomes its border
  const runShd = new Map(), pfx = new Map(), setX = (p, f) => { const o = pfx.get(p) || { ...pfOf(p) }; f(o); pfx.set(p, o); };
  const bdOf = it => it.ln && it.lw > 0 ? { val: 'single', sz: Math.max(2, Math.min(96, Math.round(it.lw * 8))), space: 1, color: hexOf(it.ln[0], it.ln[1]) } : null;
  { // a frame drawn in pieces (edges and corners) around text → a border box round that text
    const bit = it => !it.done && (((it.kind === 'line' || it.kind === 'rect' || it.kind === 'pic') && Math.min(it.r.w, it.r.h) <= 4) || (it.kind === 'pic' && it.r.w <= 12 && it.r.h <= 12));
    const fs = items.filter(bit), fg = fs.map((_, i) => i), fr = i => fg[i] === i ? i : (fg[i] = fr(fg[i]));
    for (let i = 0; i < fs.length; i++) for (let j = i + 1; j < fs.length; j++) { const a2 = fs[i].r, b2 = fs[j].r; if (ov({ x: a2.x - 3, y: a2.y - 3, w: a2.w + 6, h: a2.h + 6 }, b2) > 0) fg[fr(i)] = fr(j); }
    const gs = new Map(); fs.forEach((it, i) => { const g = fr(i); if (!gs.has(g)) gs.set(g, []); gs.get(g).push(it); });
    for (const g of gs.values()) {
      const x0 = Math.min(...g.map(i => i.r.x)), y0 = Math.min(...g.map(i => i.r.y)), bb = { x: x0, y: y0, w: Math.max(...g.map(i => i.r.x + i.r.w)) - x0, h: Math.max(...g.map(i => i.r.y + i.r.h)) - y0 };
      if (bb.w < 30 || bb.h < 14) continue; const near = (v, w) => Math.abs(v - w) <= 4;
      const sides = [g.some(i => near(i.r.y, bb.y) && i.r.w > bb.w / 3), g.some(i => near(i.r.y + i.r.h, bb.y + bb.h) && i.r.w > bb.w / 3), g.some(i => near(i.r.x, bb.x) && i.r.h > bb.h / 3), g.some(i => near(i.r.x + i.r.w, bb.x + bb.w) && i.r.h > bb.h / 3)].filter(Boolean).length;
      const cov = T.filter(t => frac(t, bb) >= 0.8); if (sides < 3 || !cov.length) continue;
      const e = g.find(i => i.kind !== 'pic' && Math.min(i.r.w, i.r.h) <= 4), c = e && (e.fill || e.ln), bd = { val: 'single', sz: 4, space: 4, color: c ? hexOf(c[0], c[1]) : 'BFBFBF' };
      for (const t of cov) setX(t.p, o => { o.bdr = { top: bd, left: bd, bottom: bd, right: bd }; }); for (const i of g) i.done = true;
    }
  }
  for (const it of items.slice().sort((a, b) => b.r.w * b.r.h - a.r.w * a.r.h)) {
    if (it.done) continue;
    const thin = (it.kind === 'line' || ((it.kind === 'rect' || it.kind === 'pic') && (it.fill || it.kind === 'pic'))) && it.r.h <= 6 && it.r.w > 24;
    if (thin) {
      const y = it.r.y + it.r.h / 2, picC = () => { try { const im = it.view.querySelector('img'), cv = document.createElement('canvas'); cv.width = cv.height = 1; const g = cv.getContext('2d'); g.drawImage(im, 0, 0, 1, 1); const d = g.getImageData(0, 0, 1, 1).data; return hexOf('#' + [d[0], d[1], d[2]].map(v => v.toString(16).padStart(2, '0')).join(''), d[3] / 255); } catch (e) { return 'BFBFBF'; } };
      const bd = it.kind === 'line' ? bdOf(it) : { val: 'single', sz: Math.max(2, Math.min(96, Math.round(it.r.h * 0.75 * 8))), space: 1, color: it.fill ? hexOf(it.fill[0], it.fill[1]) : picC() }; if (!bd) continue;
      const hov = t => Math.max(0, Math.min(t.r, it.r.x + it.r.w) - Math.max(t.x, it.r.x)) >= 0.5 * Math.min(t.w, it.r.w);
      const up = T.filter(t => y - t.b >= -2 && y - t.b <= 16 && hov(t)).sort((a, b) => (y - a.b) - (y - b.b))[0], dn = !up && T.filter(t => t.y - y >= -2 && t.y - y <= 12 && hov(t)).sort((a, b) => a.y - b.y)[0];
      if (up || dn) { setX((up || dn).p, o => { o.bdr = { ...(o.bdr || {}), [up ? 'bottom' : 'top']: bd }; }); it.done = true; }
      continue;
    }
    const pa = (q => q ? q.wpx * q.hpx : 1e9)(pageAt(it.r.y)), tone = it.kind === 'pic' && it.o.behind ? picTone(it, it.vec || it.r.w * it.r.h < 0.4 * pa) : null; // (a flat picture behind text: a colour band or a button)
    if (!(it.kind === 'rect' || it.kind === 'shape' || (tone && tone.flat)) || !(it.o.behind || !it.fill)) continue;
    const cov = T.filter(t => frac(t, it.r) >= 0.5), light = tone && luma(tone.hex) >= 0.97;
    if (!cov.length) { if (light && T.some(t => ov(t, it.r) > 0)) it.done = true; continue; } // (a white card behind text: the text alone reads the same)
    const fh = tone ? tone.hex : it.fill ? hexOf(it.fill[0], it.fill[1]) : null, fill = fh && luma(fh) < 0.97 ? fh : null, bd = tone ? null : bdOf(it);
    const tds = new Set(cov.map(t => t.p.closest('td'))), cells = fill && !tds.has(null) && [...tds].every(td => frac(boxOf(td.getBoundingClientRect()), it.r) >= 0.6);
    const narrow = !cells && it.r.w < 0.6 * Math.max(...cov.map(t => t.p.offsetWidth)); // (a box much narrower than its paragraph: shade the words, not the line)
    if (cells) { for (const td of tds) for (const c of td.parentNode.cells) if (frac(boxOf(c.getBoundingClientRect()), it.r) >= 0.6) shadeCell(c, fill); }
    else for (const t of cov) { if (narrow) { if (fill) runShd.set(t.p, fill); } else setX(t.p, o => { if (fill) o.shd = fill; if (bd) o.bdr = { top: bd, left: bd, bottom: bd, right: bd }; }); }
    if ((cells || narrow) && bd) for (const t of cov) setX(t.p, o => { o.bdr = { top: bd, left: bd, bottom: bd, right: bd }; });
    it.done = true; const near = { x: it.r.x - 3, y: it.r.y - 3, w: it.r.w + 6, h: it.r.h + 6 };
    for (const o2 of items) if (!o2.done && Math.max(o2.r.w, o2.r.h) <= 16 && ov(near, o2.r) > 0) o2.done = true; // (its rounded corners go with it)
  }
  { // tiny pieces (rounded corners) of a background that became shading go with it
    const bgs = [...docEl.querySelectorAll('td')].filter(td => td.style.background).map(td => boxOf(td.getBoundingClientRect())).concat(T.filter(t => (pfx.get(t.p) || {}).shd).map(t => ({ x: t.x - 12, y: t.y - 12, w: t.w + 24, h: t.h + 24 })));
    for (const it of items) if (!it.done && Math.max(it.r.w, it.r.h) <= 16 && bgs.some(b => ov({ x: b.x - 4, y: b.y - 4, w: b.w + 8, h: b.h + 8 }, it.r) > 0)) it.done = true;
  }
  // 4 the rest in groups of drawings that touch: a group behind text (or a text box, a page background) stays a decoration that moves
  //   with its text; a group beside a line of text goes into that line; any other group becomes one picture in the reading order
  const rest = items.filter(it => !it.done), grp = rest.map((_, i) => i), root = i => grp[i] === i ? i : (grp[i] = root(grp[i]));
  const grow = r => ({ x: r.x - 2, y: r.y - 2, w: r.w + 4, h: r.h + 4 });
  for (let i = 0; i < rest.length; i++) for (let j = i + 1; j < rest.length; j++) if (ov(grow(rest[i].r), rest[j].r) > 0) grp[root(i)] = root(j);
  const groups = new Map(); rest.forEach((it, i) => { const g = root(i); if (!groups.has(g)) groups.set(g, []); groups.get(g).push(it); });
  const keepOver = [], inLine = [], blocks = [], anchored = []; let kept = 0;
  for (const g of groups.values()) {
    const bb = g.reduce((a, it) => { const x = Math.min(a.x, it.r.x), y = Math.min(a.y, it.r.y); return { x, y, w: Math.max(a.x + a.w, it.r.x + it.r.w) - x, h: Math.max(a.y + a.h, it.r.y + it.r.h) - y }; }, { ...g[0].r });
    const pg = pageAt(bb.y), backs = T.filter(t => g.some(it => frac(t, it.r) >= 0.5)), big = pg && bb.w * bb.h > 0.5 * pg.wpx * pg.hpx;
    if (backs.length || big || g.some(it => it.kind === 'text' || it.kind === 'other')) { anchored.push({ g, bb, host: backs.sort((a, b) => a.y - b.y)[0] || null, pg }); keepOver.push(bb); continue; }
    if (bb.w < 4 && bb.h > 8) continue; // (a lone upright rule: no meaning in flowing text)
    const one = g.length === 1 && g[0].kind === 'pic', url = one ? null : rasterize(g.slice().sort((a, b) => (b.o.behind ? 1 : 0) - (a.o.behind ? 1 : 0)), bb);
    if (!one && !url) { anchored.push({ g, bb, host: null, pg }); continue; }
    const make = () => one ? atomEl(inlineXml(g[0].x), g[0].part) : inlinePng(url, bb.w * 0.75, bb.h * 0.75, ids);
    const cy = bb.y + bb.h / 2, line = bb.h <= 40 && T.filter(t => cy >= t.y && cy <= t.b && bb.x >= t.x - 60 && bb.x <= t.r + 10).sort((a, b) => Math.abs(a.y + a.h / 2 - cy) - Math.abs(b.y + b.h / 2 - cy))[0];
    if (line) { inLine.push({ t: line, x: bb.x, cy, make }); continue; }
    const hx = t => Math.min(t.r, bb.x + bb.w) > Math.max(t.x, bb.x) - 30;
    const below = T.filter(t => t.y >= bb.y - 2 && hx(t) && pageAt(t.y) === pg).sort((a, b) => a.y - b.y)[0], above = !below && T.filter(t => t.b <= bb.y + bb.h && hx(t)).sort((a, b) => b.b - a.b)[0];
    blocks.push({ t: below || above, before: !!below, bb, make });
  }
  D._fdbg = { items: items.map(it => ({ kind: it.kind, beh: it.o.behind, fill: it.fill, ln: it.ln, done: !!it.done, r: it.r, hdr: !it.el })), T: T.map(t => ({ x: t.x, y: t.y, w: t.w, h: t.h, s: t.p.textContent.slice(0, 20) })) };
  for (const it of items) if (it.el) it.el.remove();
  // in a line: at the character under the drawing's left edge
  for (const n of inLine.sort((a, b) => b.cy - a.cy || b.x - a.x)) { // (bottom lines first, right to left: earlier places do not move)
    const p = n.t.p; let off = null, last = 0;
    for (const u of units(p)) {
      if (u.t !== 't') continue;
      for (let i = 0; i < u.len && off == null; i++) { rg.setStart(u.n, i); rg.setEnd(u.n, i + 1); const q = rg.getClientRects()[0]; if (!q) continue; const cy = ((q.top + q.bottom) / 2 - pr.top) / k, cx = ((q.left + q.right) / 2 - pr.left) / k; if (Math.abs(cy - n.cy) > Math.max(6, q.height / k * 0.6)) continue; last = u.a + i + 1; if (cx >= n.x) off = u.a + i; }
      if (off != null) break;
    }
    const at = posOf(p, off == null ? last : off), a = n.make(), sp = document.createTextNode(' ');
    if (at.node.nodeType === 3) { const t2 = at.offset ? at.node.splitText(at.offset) : at.node; t2.before(a, sp); } else { at.node.insertBefore(a, at.node.childNodes[at.offset] || null); a.after(sp); }
  }
  for (const n of blocks) {
    const q = emptyP(); q.replaceChildren(n.make()); const host = n.t && n.t.p;
    if (host) { const left = offIn(host, paper).x + (parseFloat(getComputedStyle(host).paddingLeft) || 0), ind = Math.round((n.bb.x - left) * 0.75 * 20); if (ind > 40 && n.bb.x + n.bb.w - left < host.clientWidth + 2) setPf(q, { ind: { left: ind } }); if (n.before) host.before(q); else host.after(q); }
    else docEl.appendChild(q);
  }
  // a decoration that stays: anchored to the paragraph it backs (or the one above it on its page), all of a group together
  const paras = T.slice().sort((a, b) => a.y - b.y);
  const home = host => { let hp = host && host.p; if (!hp || !hp.isConnected) { hp = emptyP(); docEl.prepend(hp); return { hp, hy: 0 }; } return { hp, hy: host.p === hp ? host.py : offIn(hp, paper).y }; };
  for (const n of anchored) {
    let host = n.host; if (!host) { const onPg = paras.filter(t => pageAt(t.y) === n.pg); host = onPg.filter(t => t.y <= n.bb.y + 1).pop() || onPg[0] || paras.filter(t => t.y <= n.bb.y).pop() || paras[0]; }
    const pgl = n.pg ? n.pg.left : 0;
    for (const it of n.g) { // (a piece behind text goes with the first paragraph it backs)
      const own = T.filter(t => frac(t, it.r) >= 0.5).sort((a, b) => a.y - b.y)[0], h = own || host, { hp, hy } = home(h);
      const cx = it.r.x + it.r.w / 2, inCol = h && h.p === hp && cx >= h.px - 8 && cx <= h.px + h.pw + 8 && it.r.w < 0.9 * h.pw; // (inside its paragraph's old column: it follows the paragraph sideways too)
      const xPt = (it.r.x - pgl) / PXPT, x = inCol ? Math.max(0, (it.r.x - h.px) / PXPT + h.nind) : xPt + it.o.w > W0 && it.o.w <= W0 ? W0 - it.o.w : Math.max(0, xPt);
      hp.prepend(atomEl(retarget(it.x, x, (it.r.y - hy) / PXPT, false, inCol ? 'column' : 'page'), it.part)); kept++; (D._fdbg.anch = D._fdbg.anch || []).push({ r: it.r, own: !!own, h: h && { s: h.p.textContent.slice(0, 30), px: h.px, pw: h.pw, x: h.x, y: h.y, w: h.w, hh: h.h }, inCol, x });
    }
  }
  // 5 tables used as columns → their paragraphs, column after column (a cell's shading goes with its paragraphs)
  for (let t; (t = [...docEl.querySelectorAll('table.dx-t')].find(colText));) {
    const rs = [...t.rows].filter(r => !r.classList.contains('dx-gaprow')), out = [], n = rs[0].cells.length;
    for (let c = 0; c < n; c++) for (const row of rs) { const td = row.cells[c], sh = cellShd(td); for (const q of [...td.childNodes].filter(x => x.nodeType === 1)) { if (q.tagName === 'P' && sh) setX(q, o => { if (!o.shd) o.shd = sh; }); out.push(q); } }
    for (const q of out) if (q.tagName === 'P') setX(q, o => { if (o.ind && +o.ind.left > 0 && +o.ind.left < 720) { o.ind = { ...o.ind }; delete o.ind.left; } }); // (a column's small inset is not an indent)
    // the last column starts mid-sentence: it continues the text that runs on below the table on the same page (a newspaper column)
    const lead = [...rs[0].cells[n - 1].querySelectorAll('p')].find(q => ptext(q).trim());
    let last = null;
    if (lead && /^\p{Ll}/u.test(ptext(lead).trim())) {
      const pgT = pageAt(offIn(t, paper).y);
      for (let x = t.nextElementSibling; x && x.tagName === 'P'; x = x.nextElementSibling) { const tt = T.find(q => q.p === x); if (!tt) continue; if (pageAt(tt.py) !== pgT) break; last = x; }
      if (last && /[.!?:;]["'”’)]?$/.test(ptext(last).trim())) last = null;
    }
    t.replaceWith(...out.filter(q => !last || !rs.some(r => r.cells[n - 1] === q.parentNode)));
    if (last) last.after(...out.filter(q => !q.isConnected));
  }
  for (const [p, o] of pfx) setPf(p, o);
  // 6 the lines of the old page: line breaks and word gaps made of tabs become spaces, a paragraph cut by a wrap window is joined again
  for (const p of docEl.querySelectorAll('p')) {
    for (const br of [...p.querySelectorAll('br')]) { const U = units(p), u = U.find(x => x.n === br); if (!u || u.a + 1 >= U.total) continue; const t = ptext(p), prev = t[u.a - 1] || ' ', next = t[u.a + 1] || ' '; if (/\s/.test(prev) || /\s/.test(next)) br.remove(); else br.replaceWith(document.createTextNode(' ')); }
    const t = ptext(p), seg = t.split('\t'); if (seg.length >= 5 && seg.filter(s => /^\p{Ll}[\p{L}'’-]*[,.;:!?]?$/u.test(s.trim())).length >= seg.length * 0.6) { /* (tabs between lower-case words: a justified line, not columns) */ const w = document.createTreeWalker(p, NodeFilter.SHOW_TEXT); for (let n; (n = w.nextNode());) if (n.data.indexOf('\t') >= 0) n.data = n.data.replace(/\t/g, ' '); for (const s of [...p.querySelectorAll('span.dx-tab')]) s.replaceWith(...s.childNodes); }
  }
  const lastRun = p => [...p.querySelectorAll('span[data-rf]')].filter(r => r.textContent.trim()).pop(), firstRun = p => [...p.querySelectorAll('span[data-rf]')].find(r => r.textContent.trim());
  const szOf = (r, p) => r ? effR(rfOf(r), effP(p).style).sz || 10 : 0;
  for (const p of [...docEl.querySelectorAll('p')]) {
    for (let n = p.nextElementSibling; p.isConnected && n && n.tagName === 'P';) {
      const ea = effP(p), eb = effP(n), ta = ptext(p).replace(/\s+$/, ''), tb = ptext(n).replace(/^\s+/, '');
      if (!ta || !tb || ea.style !== eb.style || ea.num || eb.num || /[.!?:;]["'”’)]?$/.test(ta) || !/^\p{Ll}/u.test(tb) || Math.abs(szOf(lastRun(p), p) - szOf(firstRun(n), n)) > 0.5) break;
      const lr = lastRun(p); if (!/\s$/.test(ptext(p))) lr.appendChild(document.createTextNode(' '));
      for (const c of [...n.childNodes]) if (!(c.tagName === 'BR' && c === n.lastChild && !units(n).some(u => u.t !== 'b'))) p.appendChild(c);
      const nx = n.nextElementSibling; n.remove(); n = nx;
    }
  }
  // 7 paragraphs and runs: no exact line heights, wrap windows, letter spacing, page-placing gaps or section breaks
  for (const p of docEl.querySelectorAll('p')) {
    const pf = { ...pfOf(p) }, sz = szOf(firstRun(p), p) || 11;
    if (pf.sp) { const sp = { ...pf.sp }; if (sp.lineRule === 'exact' || sp.lineRule === 'atLeast') { delete sp.line; delete sp.lineRule; } if (+sp.before > Math.max(360, sz * 30)) sp.before = Math.round(Math.max(360, sz * 30)); pf.sp = sp; }
    if (pf.ind && pf.ind.right) { const ind = { ...pf.ind }; delete ind.right; pf.ind = ind; }
    setPf(p, pf);
    const rec = p.dataset.p && D.PP.get(p.dataset.p);
    if (rec && rec.x.sect && rec.el) { const el = rec.el.cloneNode(true), sp = kid(el, 'sectPr'); if (sp) sp.remove(); const id = nid(); D.PP.set(id, { ...rec, el, x: { ...rec.x, sect: false } }); p.dataset.p = id; p.classList.remove('dx-sect'); }
    // white text that has lost the colour behind it turns automatic (black)
    const shaded = pf.shd || (rec && rec.x.shd) || cellShd(p.closest('td')), tb = T.find(t => t.p === p), over = tb && keepOver.some(b => frac(tb, b) >= 0.3);
    for (const r of p.querySelectorAll('span[data-rf]')) {
      const rf = rfOf(r), o = { ...rf }; let ch = false;
      if (rf.spc != null || rf.scale != null) { delete o.spc; delete o.scale; ch = true; }
      if (runShd.has(p) && r.textContent.trim()) { o.shd = runShd.get(p); ch = true; }
      const c = effR(rf, effP(p).style).color; if (!shaded && !runShd.has(p) && !over && c && c !== 'auto' && luma(c) > 0.85) { o.color = 'auto'; ch = true; }
      if (ch) setRf(r, o);
    }
  }
  // 8 one section: the first page's size, margins that leave room on the right, headers that hold text
  const sp = D.sect, g = s0, fsp = s0.sp, mrels = relsOf(D.main);
  const refs = [...kids(fsp, 'headerReference'), ...kids(fsp, 'footerReference')].filter(ref => { const rel = mrels[ref.getAttributeNS(NS.r, 'id')]; return rel && !moved.has(resolvePath(D.main, rel.target)); }).map(r => r.cloneNode(true)), tp = kid(fsp, 'titlePg') && kid(fsp, 'titlePg').cloneNode(true);
  for (const n of ['headerReference', 'footerReference', 'type', 'titlePg']) for (const e of kids(sp, n)) e.remove();
  for (const ref of refs) insertOrdered(sp, ref, SECT_ORDER);
  if (tp && refs.some(r => wa(r, 'type') === 'first')) insertOrdered(sp, tp, SECT_ORDER);
  const sz = pgEl(sp, 'pgSz'); sz.setAttributeNS(NS.w, 'w:w', g.w); sz.setAttributeNS(NS.w, 'w:h', g.h); if (g.w > g.h) sz.setAttributeNS(NS.w, 'w:orient', 'landscape'); else sz.removeAttributeNS(NS.w, 'orient');
  const mg = pgEl(sp, 'pgMar', { gutter: 0 }), side = Math.max(720, g.l);
  for (const [k2, v] of [['top', Math.max(720, g.t)], ['bottom', Math.max(720, g.b)], ['left', side], ['right', Math.max(side, g.r)], ['header', g.hd || 720], ['footer', g.ft || 720]]) mg.setAttributeNS(NS.w, 'w:' + k2, v);
  D.geo = sp; ANC.clear(); HFV.clear(); frames._sig = '';
  changedModel(true); pageGeometry(); tidy(); fitZoom(); afterEdit(true); syncBar();
  toast(kept ? `The text now flows like a normal Word file. ${kept} drawing${kept > 1 ? 's' : ''} behind the text move with ${kept > 1 ? 'their' : 'its'} paragraph — Undo brings the exact layout back` : 'The document now flows like a normal Word file — Undo brings the exact layout back', 'success');
  return { kept };
}

/* ───────── after every change ───────── */
let editTimer = 0;
function afterEdit(now) {
  clearTimeout(editTimer);
  D.layDirty = true; editTimer = setTimeout(() => { paginate(); countWords(); lintSoon(); if (D.tab === 'format' || D.tab === 'outline' || D.tab === 'stats') renderSide(); if (findState.q) runFind(false); }, now ? 0 : 350);
}
function countWords() { const n = docEl.innerText.match(/[\p{L}\p{N}'’-]+/gu); $('#dxWords').textContent = `${(n ? n.length : 0).toLocaleString()} words`; }

/* ───────── spelling and grammar (Harper, in the writing tools) ───────── */
const LINT = { cache: new Map(), ignored: new Set(), busy: false, ok: true };
const cpIndex = text => { if (!/[\ud800-\udfff]/.test(text)) return i => i; const map = []; let u = 0; for (const ch of text) { map.push(u); u += ch.length; } map.push(u); return i => map[Math.min(i, map.length - 1)]; };
const CAT = k => k === 'Spelling' || k === 'Typo' ? 'spell' : k === 'Clarity' ? 'style' : /Style|Readability|Redundancy|Repetition|WordChoice|Enhancement/.test(k) ? 'improve' : 'grammar';
async function lintText(text) {
  const key = (store.get('xd3.style') !== 'off' ? 's' : 'n') + text; if (LINT.cache.has(key)) return LINT.cache.get(key);
  const out = []; const t = text.replace(/[\ufffc\n]/g, ' ');
  if (t.trim()) {
    const L = await X.writing.engine(), lints = await L.lint(t, { language: 'plaintext' }), at = cpIndex(t);
    for (const l of lints) { const s = l.span(), a = at(s.start), b = at(s.end), word = t.slice(a, b); const sug = [...new Set(l.suggestions().map(x => { const k = x.kind(), r = x.get_replacement_text(); x.free(); return k === 1 ? '' : k === 2 ? word + r : r; }))].slice(0, 5); out.push({ kind: l.lint_kind(), pretty: l.lint_kind_pretty(), msg: l.message(), a, b, word, sug }); l.free(); }
    if (store.get('xd3.style') !== 'off') for (const is of X.writing.styleIssues(t)) { is.word = t.slice(is.a, is.b); if (!out.some(o => o.a < is.b && is.a < o.b)) out.push(is); }
  }
  out.sort((x, y) => x.a - y.a); LINT.cache.set(key, out); if (LINT.cache.size > 5000) LINT.cache.clear(); return out;
}
let lintTimer = 0;
function lintSoon() { clearTimeout(lintTimer); lintTimer = setTimeout(lintAll, 700); }
async function lintAll() {
  if (LINT.busy || box.hidden || !LINT.ok) return; LINT.busy = true;
  try {
    const P = allParas(), t0 = performance.now();
    for (const p of P) { const t = ptext(p); if (p._lt === t && p._issues) continue; p._issues = await lintText(t); p._lt = t; if (performance.now() - t0 > 30) { await new Promise(r => setTimeout(r)); } }
    paintIssues(); if (D.tab === 'proof') renderSide();
  } catch (e) { LINT.ok = false; console.warn(e); $('#dxCnt').textContent = '!'; }
  finally { LINT.busy = false; }
}
const issuesOf = p => (p._issues || []).filter(is => !LINT.ignored.has(is.kind + ':' + is.word.toLowerCase()));
function rangeOf(p, a, b) { const A = posOf(p, a), B = posOf(p, b), r = new Range(); try { r.setStart(A.node, A.offset); r.setEnd(B.node, B.offset); } catch (e) { return null; } return r; }
function paintIssues() {
  let n = 0; const H = { spell: [], grammar: [], improve: [], style: [] };
  for (const p of allParas()) { if (p._lt !== ptext(p)) continue; for (const is of issuesOf(p)) { n++; const r = rangeOf(p, is.a, is.b); if (r) H[CAT(is.kind)].push(r); } }
  $('#dxCnt').textContent = n;
  if (window.Highlight && CSS.highlights) for (const k in H) CSS.highlights.set('dx-' + k, new Highlight(...H[k]));
}
function issueAtCaret() { const s = getSel(); if (!s || !s.c) return null; const p = allParas()[s.a[0]]; if (!p || p._lt !== ptext(p)) return null; const is = issuesOf(p).find(x => x.a <= s.a[1] && s.a[1] <= x.b && x.b > x.a); return is ? { p, is } : null; }
function replaceText(p, a, b, text) { // keeps the formatting of the first character replaced
  splitAt(p, a); splitAt(p, b); const U = units(p).filter(u => u.a >= a && u.a + u.len <= b && u.len), host = U.find(u => u.t === 't');
  const t = document.createTextNode(text);
  if (host) { host.n.before(t); } else { const pos = posOf(p, a); if (pos.node.nodeType === 3) { pos.node.splitText(pos.offset); pos.node.after(t); } else pos.node.insertBefore(t, pos.node.childNodes[pos.offset] || null); }
  for (const u of U) u.n.remove();
  fixPara(p);
}
function showIssuePop(p, is) {
  const r = rangeOf(p, is.a, is.b); if (!r) return; const rc = r.getBoundingClientRect(), c = CAT(is.kind);
  pop.innerHTML = `<div class="wpk wm-${c}">${esc(is.pretty)}</div><div class="wpm">${esc(is.msg)}</div><div class="wps">${is.sug.map((s, k) => `<button data-s="${k}">${s ? esc(s) : '<i>remove</i>'}</button>`).join('') || '<span class="mut">No automatic fix — edit the text directly.</span>'}</div><div class="wpa">${c === 'spell' ? '<button data-a="dict">Add to dictionary</button>' : ''}<button data-a="ignore">Ignore</button></div>`;
  pop.hidden = false; pop.style.left = clamp(rc.left, 8, innerWidth - pop.offsetWidth - 8) + 'px'; pop.style.top = (rc.bottom + 6 + pop.offsetHeight > innerHeight ? rc.top - pop.offsetHeight - 6 : rc.bottom + 6) + 'px';
  pop.onclick = async e => { const s = e.target.closest('[data-s]'), a = e.target.closest('[data-a]'); if (!s && !a) return; pop.hidden = true; if (s) return fixIssue(p, is, is.sug[+s.dataset.s]); if (a.dataset.a === 'dict') { await X.writing.dictAdd(is.word); LINT.cache.clear(); allParas().forEach(q => { q._lt = null; }); lintAll(); } else { LINT.ignored.add(is.kind + ':' + is.word.toLowerCase()); paintIssues(); renderSide(); } };
}
function fixIssue(p, is, rep) { if (ptext(p).slice(is.a, is.b) !== is.word) return toast('The text has changed — check again', 'info'); const P = allParas(), i = P.indexOf(p); snapshot(); replaceText(p, is.a, is.b, rep); restyle([rootOf(p)]); setSel({ a: [i, is.a + rep.length], b: [i, is.a + rep.length] }); afterEdit(); lintSoon(); }
docEl.addEventListener('click', e => {
  if ((e.ctrlKey || e.metaKey) && e.target.closest('a[href]')) { const h = e.target.closest('a[href]').getAttribute('href'); if (/^(https?|mailto):/i.test(h)) window.open(h, '_blank', 'noopener'); return; }
  const img = e.target.closest('.dx-img'); if (img) return selectImage(img);
  const x = issueAtCaret(); if (x) showIssuePop(x.p, x.is); else pop.hidden = true;
});
document.addEventListener('pointerdown', e => { if (!pop.hidden && !e.target.closest('#dxPop')) pop.hidden = true; }, true);

/* ───────── pictures: click to select, drag the corner to resize ───────── */
let grip = null;
function selectImage(a) {
  const r = document.createRange(); r.selectNode(a); const s = getSelection(); s.removeAllRanges(); s.addRange(r);
  $$('.dx-img.on', box).forEach(x => x.classList.remove('on')); a.classList.add('on');
  if (!grip) { grip = document.createElement('div'); grip.className = 'dx-grip'; grip.title = 'Drag to resize'; paper.appendChild(grip); grip.addEventListener('pointerdown', resizeImg); }
  grip.hidden = false; grip._a = a; placeGrip();
}
function placeGrip() { const a = grip && grip._a; if (!a || !a.isConnected || !a.classList.contains('on')) { if (grip) grip.hidden = true; return; } const im = a.querySelector('img') || a, pr = paper.getBoundingClientRect(), r = im.getBoundingClientRect(); grip.style.left = (r.right - pr.left) / D.zoom - 6 + 'px'; grip.style.top = (r.bottom - pr.top) / D.zoom - 6 + 'px'; }
function resizeImg(e) {
  e.preventDefault(); const a = grip._a, im = a && a.querySelector('img'); if (!im) return; snapshot();
  const w0 = im.offsetWidth, h0 = im.offsetHeight, x0 = e.clientX, ratio = h0 / w0 || 1;
  const move = ev => { const w = clamp(w0 + (ev.clientX - x0) / D.zoom, 16, contentWidthPt() * 4 / 3); im.style.width = w + 'px'; im.style.height = w * ratio + 'px'; placeGrip(); };
  const up = () => { removeEventListener('pointermove', move); removeEventListener('pointerup', up); a.dataset.sw = (im.offsetWidth * 0.75).toFixed(2); a.dataset.sh = (im.offsetHeight * 0.75).toFixed(2); dirty(); afterEdit(); };
  addEventListener('pointermove', move); addEventListener('pointerup', up);
}
document.addEventListener('selectionchange', () => {
  if (box.hidden) return;
  if (inEditor()) { D.lastSel = getSel(); if (D.pending && D.pendingAt !== JSON.stringify(D.lastSel)) D.pending = null; }
  if (grip && !grip.hidden) { const s = getSelection(); if (!s.rangeCount || !s.getRangeAt(0).intersectsNode(grip._a)) { grip._a.classList.remove('on'); grip.hidden = true; } }
  clearTimeout(syncBar.t); syncBar.t = setTimeout(syncBar, 60);
});

/* ───────── the ribbon follows the cursor ───────── */
function fillStyles() {
  const sel = $('#dxStyle'), list = Object.values(D.ST.byId).filter(s => s.type === 'paragraph' && !s.hidden && (s.q || s.prio < 50 || SHOWN.includes(s.name.toLowerCase())));
  const pretty = n => n.replace(/^heading (\d)$/i, 'Heading $1').replace(/^\w/, c => c.toUpperCase());
  const rank = s => { const i = SHOWN.indexOf(s.name.toLowerCase()); return i < 0 ? 100 + s.prio : i; };
  const have = new Set(list.map(s => s.name.toLowerCase())), extra = ['title', 'heading 1', 'heading 2', 'heading 3', 'quote'].filter(n => !have.has(n));
  sel.innerHTML = list.sort((a, b) => rank(a) - rank(b) || a.name.localeCompare(b.name)).map(s => `<option value="${esc(s.id)}">${esc(pretty(s.name))}</option>`).join('') + extra.map(n => `<option value="${esc(n)}">${esc(pretty(n))}</option>`).join('');
  const fonts = new Set(['Calibri', 'Calibri Light', 'Aptos', 'Arial', 'Arial Narrow', 'Cambria', 'Candara', 'Century Gothic', 'Comic Sans MS', 'Consolas', 'Constantia', 'Corbel', 'Courier New', 'Franklin Gothic Book', 'Garamond', 'Georgia', 'Palatino Linotype', 'Segoe UI', 'Tahoma', 'Times New Roman', 'Trebuchet MS', 'Verdana', ...Object.keys(X.BUNDLED)]);
  for (const s of Object.values(D.ST.byId)) if (s.r.font && s.r.font[0] !== '@') fonts.add(s.r.font); fonts.add(D.theme.minor); fonts.add(D.theme.major);
  $('#dxFonts').innerHTML = [...fonts].sort().map(f => `<option value="${esc(f)}">`).join('');
}
function syncBar() {
  if (box.hidden) return;
  const e = curRun(), s = getSelection(), p = s.rangeCount && paraOf(s.getRangeAt(0).startContainer), pe = p ? effP(p) : null;
  const on = (c, v) => { const b = $(`[data-c=${c}]`, box); if (b) b.classList.toggle('on', !!v); };
  on('b', e.b); on('i', e.i); on('u', e.u && e.u !== 'none'); on('strike', e.strike); on('sup', e.va === 'superscript'); on('sub', e.va === 'subscript');
  if (document.activeElement !== $('#dxFont')) $('#dxFont').value = fontName(e.font); if (document.activeElement !== $('#dxSize')) $('#dxSize').value = e.sz || 10;
  if (pe) { const jc = { both: 'aj', distribute: 'aj', center: 'ac', right: 'ar', end: 'ar' }[pe.jc] || 'al'; for (const c of ['al', 'ac', 'ar', 'aj']) on(c, c === jc); if (document.activeElement !== $('#dxStyle')) { const el = $('#dxStyle'), want = pe.style || D.ST.defPara || ''; el.value = want; if (el.selectedIndex < 0) { const nm = ((D.ST.byId[want] || {}).name || 'Normal'); const o = [...el.options].find(x => x.textContent.toLowerCase() === nm.toLowerCase()) || [...el.options].find(x => /^normal$/i.test(x.textContent)); el.value = o ? o.value : (el.options[0] || {}).value; } } const nm = ((D.ST.byId[pe.style] || {}).name || 'normal').toLowerCase(); $$('.dxg', box).forEach(g => g.classList.toggle('on', g.dataset.s === nm)); const L = pe.num && pe.num.id !== '0' && lvlOf(pe.num); on('bullets', L && L.fmt === 'bullet'); on('numbers', L && L.fmt !== 'bullet'); }
  $('#dxColBar').style.background = D.color || '#c00000'; $('#dxHlBar').style.background = HL[D.hl || 'yellow'];
}
$('#dxStyle').addEventListener('change', e => setStyle(e.target.value));
$('#dxFont').addEventListener('change', e => { const f = e.target.value.trim(); if (f) setRunProp('font', f); });
$('#dxFont').addEventListener('keydown', e => { e.stopPropagation(); if (e.key === 'Enter') { e.preventDefault(); e.target.blur(); } });
$('#dxSize').addEventListener('change', e => { const v = parseFloat(e.target.value); if (v >= 1 && v <= 1638) setRunProp('sz', Math.round(v * 2) / 2); });
$('#dxSize').addEventListener('keydown', e => { e.stopPropagation(); if (e.key === 'Enter') { e.preventDefault(); e.target.blur(); } });
for (const id of ['#dxStyle', '#dxFont', '#dxSize']) $(id).addEventListener('mousedown', () => { if (inEditor()) D.lastSel = getSel(); });
$('#dxRib').addEventListener('mousedown', e => { if (e.target.closest('button')) e.preventDefault(); }); // ribbon buttons never take the cursor out of the text

/* ───────── commands ───────── */
const COLORS = [['C00000', 'Dark red'], ['FF0000', 'Red'], ['ED7D31', 'Orange'], ['FFC000', 'Gold'], ['70AD47', 'Green'], ['00B050', 'Bright green'], ['00B0F0', 'Light blue'], ['0070C0', 'Blue'], ['4472C4', 'Blue-grey'], ['002060', 'Navy'], ['7030A0', 'Purple'], ['404040', 'Dark grey'], ['7F7F7F', 'Grey']];
function run(c, btn) {
  const at = btn ? (() => { const r = btn.getBoundingClientRect(); return { clientX: r.left, clientY: r.bottom + 4, target: btn }; })() : null;
  switch (c) {
    case 'b': case 'i': case 'strike': return setRunProp(c, true, true);
    case 'u': return setRunProp('u', 'single', true);
    case 'sup': case 'sub': { const v = c === 'sup' ? 'superscript' : 'subscript'; return setRunProp('va', curRun().va === v ? undefined : v); }
    case 'grow': return sizeStep(1); case 'shrink': return sizeStep(-1);
    case 'color': return setRunProp('color', (D.color || '#C00000').slice(1).toUpperCase());
    case 'colormenu': return X.showCtx(at, [X.mi('Automatic (black)', '', 'dx_color', { d: { v: 'auto' }, keep: 1, checked: !curRun().color || curRun().color === 'auto' }), '-', ...COLORS.map(([v, n]) => X.mi(n, '', 'dx_color', { d: { v }, keep: 1, checked: curRun().color === v }))]);
    case 'hl': return setRunProp('hl', D.hl || 'yellow');
    case 'hlmenu': return X.showCtx(at, [X.mi('No highlight', '', 'dx_hl', { d: { v: 'none' }, keep: 1 }), ...['yellow', 'green', 'cyan', 'magenta', 'red', 'blue', 'lightGray', 'darkYellow'].map(v => X.mi(v.replace(/([A-Z])/g, ' $1').toLowerCase(), '', 'dx_hl', { d: { v }, keep: 1 }))]);
    case 'clear': return clearFormatting(true); case 'clearchar': return clearFormatting(false);
    case 'al': return align('left'); case 'ac': return align('center'); case 'ar': return align('right'); case 'aj': return align('both');
    case 'bullets': return listToggle('bullet'); case 'numbers': return listToggle('decimal');
    case 'indent': return indent(1); case 'outdent': return indent(-1);
    case 'spacing': return spacingMenu(at);
    case 'table': return insertTable(); case 'picture': if (inEditor()) D.lastSel = getSel(); return $('#dxPic').click();
    case 'link': return insertLink(); case 'pbreak': return pageBreak(); case 'symbol': if (inEditor()) D.lastSel = getSel(); return symbolMenu(at); case 'hf': return hfMenu(at);
    case 'find': return findBar(true); case 'page': return pageSetup(); case 'flow': return makeFlowing();
    case 'panel': D.side = !D.side; store.set('xd3.dxside', D.side ? 'on' : 'off'); box.classList.toggle('noside', !D.side); renderSide(); return setTimeout(() => { fitZoom(); paginate(); }, 30);
    case 'close': return close(); case 'home': return X.ACT.home(); case 'new': return newFromTemplate(); case 'open': return $('#dxFile').click();
    case 'undo': return undo(); case 'redo': return redo();
    case 'save': return save(); case 'pdf': return savePdf(); case 'print': return printDoc();
    case 'autofmt': return autoFormat(); case 'fixfmt': D.side = true; box.classList.remove('noside'); return setTab('format');
    case 'fprev': return stepFind(-1); case 'fnext': return stepFind(1); case 'rone': return replaceOne(); case 'rall': return replaceAll(); case 'fclose': return findBar(false);
    case 'zin': return setZoom(D.zoom + 0.1); case 'zout': return setZoom(D.zoom - 0.1); case 'z100': return setZoom(1); case 'zfit': return fitZoom();
    case 'paint': return painter(btn); case 'sty': return setStyle(btn.dataset.s);
    case 'case': return X.showCtx(at, [['sentence', 'Sentence case.'], ['lower', 'lowercase'], ['upper', 'UPPERCASE'], ['title', 'Capitalize Each Word'], ['toggle', 'tOGGLE cASE']].map(([v, l]) => X.mi(l, '', 'dx_case', { d: { v }, keep: 1 })));
    case 'shade': return X.showCtx(at, [X.mi('No shading', '', 'dx_pshd', { d: { v: '' }, keep: 1 }), '-', ...SHADES.map(([v, n]) => X.mi(n, '', 'dx_pshd', { d: { v }, keep: 1 }))]);
    case 'border': return X.showCtx(at, [['bottom', 'Bottom border'], ['top', 'Top border'], ['topbottom', 'Top and bottom'], ['box', 'Box around'], ['left', 'Left bar'], ['', 'No border']].map(([v, l]) => X.mi(l, v ? 'rect' : '', 'dx_pbdr', { d: { v }, keep: 1 })));
    case 'pagenum': return headerFooter('pagenum'); case 'hr': return horizontalLine(); case 'check': return checklist(); case 'toc': return insertToc();
    case 'date': { const d = new Date(), fm = o => d.toLocaleDateString(undefined, o), t = d.toLocaleTimeString(undefined, { hour: '2-digit', minute: '2-digit' }); if (inEditor()) D.lastSel = getSel(); return X.showCtx(at, [fm({ day: 'numeric', month: 'long', year: 'numeric' }), fm({ weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' }), fm(), d.toISOString().slice(0, 10), t, fm() + ' ' + t].map(v => X.mi(v, '', 'dx_text', { d: { v }, keep: 1 }))); }
    case 'letterhead': return X.letterheads ? X.letterheads.menu(at, 'docx') : null;
    case 'margins': return X.showCtx(at, [['1440,1440', 'Normal — 2.54 cm'], ['720,720', 'Narrow — 1.27 cm'], ['1440,1080', 'Moderate — 2.54 / 1.91 cm'], ['1440,2880', 'Wide — 2.54 / 5.08 cm'], ['c', 'Custom margins…']].map(([v, l]) => X.mi(l, '', 'dx_margin', { d: { v }, keep: 1 })));
    case 'orient': return X.showCtx(at, [X.mi('Portrait', '', 'dx_orient', { d: { v: 'p' }, keep: 1, checked: D.pg.w <= D.pg.h }), X.mi('Landscape', '', 'dx_orient', { d: { v: 'l' }, keep: 1, checked: D.pg.w > D.pg.h })]);
    case 'size': return X.showCtx(at, PAPER.map(([n, w, h]) => X.mi(`${n} — ${(w / 567).toFixed(1)} × ${(h / 567).toFixed(1)} cm`, '', 'dx_size', { d: { v: n }, keep: 1, checked: Math.abs(Math.min(D.pg.w, D.pg.h) - w) < 30 && Math.abs(Math.max(D.pg.w, D.pg.h) - h) < 30 })));
    case 'cols': { const n = +(wa(kid(D.geo || D.sect, 'cols'), 'num') || 1); return X.showCtx(at, [1, 2, 3].map(v => X.mi(['One column', 'Two columns', 'Three columns'][v - 1], '', 'dx_cols', { d: { v }, keep: 1, checked: n === v }))); }
    case 'aiwrite': c = 'human'; // falls through: the AI lives on the Humanize tab
    case 'proof': case 'human': case 'nav': case 'count': D.side = true; store.set('xd3.dxside', 'on'); box.classList.remove('noside'); setTab(c === 'nav' ? 'outline' : c === 'count' ? 'stats' : c); return setTimeout(() => { fitZoom(); paginate(); }, 30);
    case 'speak': return readAloud(btn);
    case 'focus': box.classList.toggle('dxfocus'); return setTimeout(() => { fitZoom(); paginate(); }, 30);
  }
}
box.addEventListener('click', e => { const b = e.target.closest('[data-c]'); if (b && box.contains(b) && !b.disabled) run(b.dataset.c, b); });
Object.assign(X.ACT, {
  dx: () => openEditor(), dxnew: () => newFromTemplate(), dx_ai: b => { if (inEditor()) D.lastSel = getSel(); if (D.aiScope === 'doc') { D.aiScope = 'auto'; } run('aiwrite'); setTimeout(() => aiGo(b.dataset.v), 60); }, dx_color: b => { if (b.dataset.v !== 'auto') D.color = '#' + b.dataset.v; setRunProp('color', b.dataset.v); }, dx_hl: b => { if (b.dataset.v !== 'none') D.hl = b.dataset.v; setRunProp('hl', b.dataset.v === 'none' ? undefined : b.dataset.v); if (b.dataset.v === 'none') setRunProp('shd', undefined); },
  dx_ls: b => paraCmd(pf => { pf.sp = { ...(pf.sp || {}), line: +b.dataset.v, lineRule: 'auto' }; }), dx_sp: b => paraCmd(pf => { pf.sp = { ...(pf.sp || {}), [b.dataset.k]: +b.dataset.v }; }),
  dx_keep: () => paraCmd((pf, p) => { pf.keepNext = !effP(p).keepNext || undefined; }), dx_pbb: () => paraCmd((pf, p) => { pf.pbb = !effP(p).pbb || undefined; }),
  dx_sym: b => { useSel(); snapshot(); document.execCommand('insertText', false, b.dataset.v); }, dx_hf: b => headerFooter(b.dataset.v),
  dx_case: b => changeCase(b.dataset.v), dx_pshd: b => paraCmd(pf => { pf.shd = b.dataset.v || undefined; }),
  dx_pbdr: b => paraCmd(pf => { const v = b.dataset.v; pf.bdr = !v ? undefined : v === 'box' ? { top: BDR, left: BDR, bottom: BDR, right: BDR } : v === 'topbottom' ? { top: BDR, bottom: BDR } : v === 'left' ? { left: { ...BDR, sz: 24, space: 4, color: '4472C4' } } : { [v]: BDR }; }),
  dx_text: b => { useSel(); snapshot(); document.execCommand('insertText', false, b.dataset.v); }, dx_margin: b => setMargins(b.dataset.v), dx_orient: b => setOrient(b.dataset.v), dx_size: b => setPaper(b.dataset.v), dx_cols: b => setColumns(+b.dataset.v),
  dx_tbl: b => tableOp(b.dataset.v), dx_cut: () => { useSel(); document.execCommand('cut'); }, dx_copy: () => { useSel(); document.execCommand('copy'); },
  dx_paste: async () => { useSel(); try { const t = await navigator.clipboard.readText(); snapshot(); document.execCommand('insertText', false, t); tidy(); afterEdit(); } catch (e) { toast('Use Ctrl+V to paste', 'info'); } },
  dx_selall: () => { const r = document.createRange(), root = rootOf(getSelection().anchorNode) || docEl; r.selectNodeContents(root); const s = getSelection(); s.removeAllRanges(); s.addRange(r); },
  dx_fix: b => { const x = issueAtCaret(); if (x) fixIssue(x.p, x.is, x.is.sug[+b.dataset.v]); }, dx_link: () => insertLink(), dx_unlink: () => insertLink(),
});
['dx', 'dxnew', 'dx_ai', 'dx_color', 'dx_hl', 'dx_ls', 'dx_sp', 'dx_keep', 'dx_pbb', 'dx_sym', 'dx_hf', 'dx_tbl', 'dx_cut', 'dx_copy', 'dx_paste', 'dx_selall', 'dx_fix', 'dx_link', 'dx_unlink', 'dx_case', 'dx_pshd', 'dx_pbdr', 'dx_text', 'dx_margin', 'dx_orient', 'dx_size', 'dx_cols'].forEach(a => X.NO_DOC.add(a));
$('#dxRtabs').addEventListener('click', e => { const t = e.target.closest('[data-rt]'); if (!t) return; $$('#dxRtabs button', box).forEach(b => b.classList.toggle('on', b === t)); $$('.dxp', box).forEach(p => p.classList.toggle('on', p.dataset.rp === t.dataset.rt)); store.set('xd3.dxrt', t.dataset.rt); });
$('#dxRtabs').addEventListener('mousedown', e => e.preventDefault());
{ const t = store.get('xd3.dxrt'); if (t && t !== 'home') { const b = $(`#dxRtabs [data-rt=${t}]`); if (b) b.click(); } }
for (const r of [docEl, hdrEl, ftrEl]) { r.addEventListener('mouseup', () => setTimeout(paintSelection, 0)); r.addEventListener('click', toggleBox); }
window.addEventListener('keydown', e => { if (e.key === 'Escape' && D.open && (D.paint || box.classList.contains('dxfocus'))) { D.paint = null; const pb = $('[data-c=paint]', box); if (pb) pb.classList.remove('on'); box.classList.remove('dxfocus'); setTimeout(fitZoom, 30); } });
$('#dxPic').addEventListener('change', e => { const f = e.target.files[0]; e.target.value = ''; if (f) insertPicture(f); });
$('#dxFile').addEventListener('change', e => { const f = e.target.files[0]; e.target.value = ''; if (f) openFile(f); });
for (const r of [hdrEl, docEl, ftrEl]) r.addEventListener('contextmenu', e => {
  e.preventDefault(); const s0 = getSelection(), inside = s0.rangeCount && !s0.isCollapsed && [...s0.getRangeAt(0).getClientRects()].some(q => e.clientX >= q.left && e.clientX <= q.right && e.clientY >= q.top && e.clientY <= q.bottom);
  if (!inside && document.caretRangeFromPoint) { const pos = document.caretRangeFromPoint(e.clientX, e.clientY); if (pos && rootOf(pos.startContainer)) { s0.removeAllRanges(); s0.addRange(pos); } } // (a right-click moves the cursor, unless it is on the selected text)
  const x = issueAtCaret();
  const td = e.target.closest('td'), link = e.target.closest('a[data-w]'), M = X.mi;
  X.showCtx(e, [...(x ? [...x.is.sug.slice(0, 4).map((s, k) => M(s || '(remove)', 'check', 'dx_fix', { d: { v: k }, keep: 1 })), '-'] : []),
    M('Cut', 'cut', 'dx_cut', { key: 'Ctrl+X', keep: 1 }), M('Copy', 'copy', 'dx_copy', { key: 'Ctrl+C', keep: 1 }), M('Paste', 'pages', 'dx_paste', { key: 'Ctrl+V', keep: 1 }), M('Select all', 'selall', 'dx_selall', { key: 'Ctrl+A', keep: 1 }), '-',
    { label: 'AI writing', icon: 'sparkle', sub: ['humanize', 'fix', 'formal', 'friendly', 'shorter', 'longer', 'summary', 'bullets', 'continue'].map(k => M(X.ai.TASKS[k].label, X.ai.TASKS[k].icon, 'dx_ai', { d: { v: k }, keep: 1 })) }, '-',
    link ? M('Edit or remove link…', 'link', 'dx_link', { keep: 1 }) : M('Link…', 'link', 'dx_link', { key: 'Ctrl+K', keep: 1 }),
    td && { label: 'Table', icon: 'table', sub: [M('Insert row above', 'up', 'dx_tbl', { d: { v: 'rowabove' }, keep: 1 }), M('Insert row below', 'down', 'dx_tbl', { d: { v: 'rowbelow' }, keep: 1 }), M('Insert column left', 'left', 'dx_tbl', { d: { v: 'colleft' }, keep: 1 }), M('Insert column right', 'right', 'dx_tbl', { d: { v: 'colright' }, keep: 1 }), '-', M('Delete row', 'trash', 'dx_tbl', { d: { v: 'delrow' }, keep: 1 }), M('Delete column', 'trash', 'dx_tbl', { d: { v: 'delcol' }, keep: 1 }), M('Delete table', 'trash', 'dx_tbl', { d: { v: 'deltable' }, keep: 1, danger: 1 })] },
  ]);
});

/* ───────── find & replace ───────── */
const findState = { q: '', hits: [], cur: -1 };
function findBar(show) { $('#dxFind').hidden = !show; if (show) { const t = getSelection().toString(); if (t && t.length < 80 && !/\n/.test(t)) $('#dxQ').value = t; $('#dxQ').focus(); $('#dxQ').select(); runFind(true); } else { findState.q = ''; findState.hits = []; paintFind(); } }
function runFind(jump) {
  const q = $('#dxQ').value; findState.q = q; findState.hits = []; if (!q) { paintFind(); $('#dxQn').textContent = ''; return; }
  const rx = new RegExp(($('#dxQw').checked ? '\\b' : '') + q.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + ($('#dxQw').checked ? '\\b' : ''), $('#dxQc').checked ? 'g' : 'gi');
  for (const p of allParas()) { const t = ptext(p); for (const m of t.matchAll(rx)) if (m[0]) findState.hits.push({ p, a: m.index, b: m.index + m[0].length }); }
  if (jump || findState.cur >= findState.hits.length) findState.cur = findState.hits.length ? 0 : -1;
  $('#dxQn').textContent = findState.hits.length ? `${findState.cur + 1} / ${findState.hits.length}` : 'No matches'; paintFind();
}
function paintFind() {
  if (!window.Highlight || !CSS.highlights) return; const all = [], cur = [];
  findState.hits.forEach((h, i) => { const r = rangeOf(h.p, h.a, h.b); if (r) (i === findState.cur ? cur : all).push(r); });
  CSS.highlights.set('dx-find', new Highlight(...all)); CSS.highlights.set('dx-findcur', new Highlight(...cur));
}
function stepFind(d) { if (!findState.hits.length) return runFind(true); findState.cur = (findState.cur + d + findState.hits.length) % findState.hits.length; const h = findState.hits[findState.cur], P = allParas(); setSel({ a: [P.indexOf(h.p), h.a], b: [P.indexOf(h.p), h.b] }, true); $('#dxQ').focus(); $('#dxQn').textContent = `${findState.cur + 1} / ${findState.hits.length}`; paintFind(); }
function replaceOne() { const h = findState.hits[findState.cur]; if (!h) return runFind(true); snapshot(); replaceText(h.p, h.a, h.b, $('#dxR').value); restyle(); runFind(false); if (findState.hits.length) stepFind(0); afterEdit(); }
function replaceAll() { runFind(false); const n = findState.hits.length; if (!n) return toast('Nothing to replace'); snapshot(); for (const h of [...findState.hits].reverse()) replaceText(h.p, h.a, h.b, $('#dxR').value); restyle(); runFind(true); afterEdit(); toast(`${n} replacement(s)`, 'success'); }
$('#dxQ').addEventListener('input', () => runFind(true));
for (const id of ['#dxQ', '#dxR']) $(id).addEventListener('keydown', e => { e.stopPropagation(); if (e.key === 'Enter') { e.preventDefault(); if (id === '#dxR') replaceOne(); else stepFind(e.shiftKey ? -1 : 1); } if (e.key === 'Escape') findBar(false); });
for (const id of ['#dxQc', '#dxQw']) $(id).addEventListener('change', () => runFind(true));

/* ───────── zoom ───────── */
function setZoom(z) { D.zoom = clamp(Math.round(z * 100) / 100, 0.4, 2); paper.style.zoom = D.zoom; $('#dxZoom').value = Math.round(D.zoom * 100); $('#dxZl').textContent = Math.round(D.zoom * 100) + '%'; placeGrip(); clearTimeout(paginate.s); paginate.s = setTimeout(paginate, 60); }
function fitZoom() { const avail = scroller.clientWidth - 32, w = tw2pt(D.paged && D.paperW ? D.paperW : D.pg.w) * 96 / 72; setZoom(avail > 0 && w > avail ? avail / w : 1); }
$('#dxZoom').addEventListener('input', e => setZoom(+e.target.value / 100));
scroller.addEventListener('wheel', e => { if (!e.ctrlKey) return; e.preventDefault(); setZoom(D.zoom * (e.deltaY < 0 ? 1.1 : 1 / 1.1)); }, { passive: false });
$('#dxDialect').addEventListener('change', e => { X.writing.W.dialect = e.target.value; store.set('xd3.dialect', e.target.value); LINT.cache.clear(); allParas().forEach(p => { p._lt = null; }); lintAll(); });

/* ───────── side panel ───────── */
function setTab(t) { D.tab = t; $$('.dx-tabs button[data-t]', box).forEach(b => b.classList.toggle('on', b.dataset.t === t)); renderSide(); }
$$('.dx-tabs button[data-t]', box).forEach(b => b.onclick = () => setTab(b.dataset.t));
const CATNAME = { spell: 'Spelling', grammar: 'Grammar', improve: 'Word choice', style: 'Clarity' };
const tidyMsg = m => String(m || '').replace(/`([^`]*)`/g, '“$1”').replace(/\s+/g, ' ').trim();
const ctxOf = (p, is) => { const t = ptext(p).replace(/￼/g, ' '), a = Math.max(0, is.a - 36), b = Math.min(t.length, is.b + 36); return `${a ? '…' : ''}${esc(t.slice(a, is.a))}<mark>${esc(t.slice(is.a, is.b)) || '&nbsp;'}</mark>${esc(t.slice(is.b, b))}${b < t.length ? '…' : ''}`; };
function renderSide() {
  if (box.hidden || !D.side) return;
  if (D.tab === 'proof') renderProof(); else if (D.tab === 'format') renderFormat(); else if (D.tab === 'human') renderHuman(); else if (D.tab === 'outline') renderOutline(); else renderStats();
  icons(pane);
}
const paraTexts = () => [...docEl.querySelectorAll('p')].map(p => ({ text: ptext(p).replace(/￼/g, '').replace(/\n$/, '') }));

/* ───────── Spelling & grammar: cards with the sentence around each mistake, filters, Fix all, Ignore, Add to dictionary ───────── */
D.pf = 'all';
const sugLabel = s => !s ? '<i>remove</i>' : /^\s+$/.test(s) ? '<i>one space</i>' : esc(s);
function proofList() { const list = []; allParas().forEach(p => { if (p._lt === ptext(p) && !p.closest('.dx-toc')) issuesOf(p).forEach(is => list.push({ p, is, c: CAT(is.kind) })); }); return list; }
function renderProof() {
  const all = proofList(), by = {}; for (const x of all) by[x.c] = (by[x.c] || 0) + 1;
  if (D.pf !== 'all' && !by[D.pf]) D.pf = 'all';
  const list = D.pf === 'all' ? all : all.filter(x => x.c === D.pf); D.plist = list;
  const fixable = c => all.filter(x => (c === 'all' || x.c === c) && x.is.sug.length && x.is.sug[0] != null).length, fs = fixable('spell'), fa = fixable('all');
  pane.innerHTML = `${!LINT.ok ? '<div class="ai-load err">The spelling checker could not start here.</div>' : ''}
    <div class="wsum"><div class="wsn"><b>${all.length}</b><span>${all.length === 1 ? 'suggestion' : 'suggestions'}<small>${LINT.busy ? 'Checking…' : all.length ? 'Click a card to see it in the text' : 'Checked on this device'}</small></span></div>
      <div class="wsa">${fs && fs !== fa ? `<button class="btn line sm" data-pfix="spell" title="Use the first suggestion for every spelling mistake — Ctrl+Z undoes it">Fix spelling (${fs})</button>` : ''}${fa ? `<button class="btn primary sm" data-pfix="all" title="Use the first suggestion everywhere — Ctrl+Z undoes it">Fix all (${fa})</button>` : ''}</div></div>
    ${all.length ? `<div class="wchips">${['all', ...Object.keys(CATNAME)].filter(c => c === 'all' || by[c]).map(c => `<button class="wchip wm-${c}${D.pf === c ? ' on' : ''}" data-pf="${c}">${c === 'all' ? 'All' : CATNAME[c]} <b>${c === 'all' ? all.length : by[c]}</b></button>`).join('')}</div>` : ''}
    <div class="wcards">${list.length ? list.slice(0, 300).map(({ p, is, c }, k) => `<div class="wcard wm-${c}" data-k="${k}"><div class="wck wm-${c}">${CATNAME[c]}${is.pretty && !/^(spelling|typo|miscellaneous|clarity)$/i.test(is.pretty) && is.pretty.toLowerCase() !== CATNAME[c].toLowerCase() ? ` · <span>${esc(is.pretty)}</span>` : ''}</div>
      <div class="wctx">${ctxOf(p, is)}</div><div class="wcm">${esc(tidyMsg(is.msg))}</div>
      <div class="wps">${is.sug.slice(0, 3).map((s, j) => `<button data-s="${j}">${sugLabel(s)}</button>`).join('')}<span class="wsp"></span>${c === 'spell' ? '<button class="lnk" data-dict title="Never mark this word again (kept on this device)">Add to dictionary</button>' : ''}<button class="lnk" data-ign title="Hide this suggestion">Ignore</button></div></div>`).join('')
      : `<div class="empty">${LINT.busy ? 'Checking…' : '<b>No issues found</b><br>Spelling and grammar look good. 🎉'}</div>`}</div>
    <details class="wopts"><summary>Checking options</summary><label class="inl"><input type="checkbox" id="dxStyleChk"${store.get('xd3.style') !== 'off' ? ' checked' : ''}>Clarity suggestions (long sentences, passive voice, weak words)</label>${LINT.ignored.size ? `<button class="lnk" data-unign>Show ${LINT.ignored.size} ignored suggestion(s) again</button>` : ''}</details>`;
  const chk = $('#dxStyleChk'); chk.onchange = () => { store.set('xd3.style', chk.checked ? 'on' : 'off'); allParas().forEach(p => { p._lt = null; }); lintAll(); };
}
function fixAll(cat) { // the first suggestion for every issue of a kind (last first in each paragraph) — one undo step
  const list = proofList().filter(x => (cat === 'all' || x.c === cat) && x.is.sug.length && x.is.sug[0] != null); if (!list.length) return;
  snapshot(); let n = 0; const by = new Map(); for (const x of list) { if (!by.has(x.p)) by.set(x.p, []); by.get(x.p).push(x.is); }
  for (const [p, iss] of by) { let last = Infinity; for (const is of iss.sort((a, b) => b.a - a.a)) { if (is.b > last || ptext(p).slice(is.a, is.b) !== is.word) continue; replaceText(p, is.a, is.b, is.sug[0]); last = is.a; n++; } }
  restyle(); afterEdit(); allParas().forEach(p => { p._lt = null; }); lintAll(); toast(`${n} correction${n === 1 ? '' : 's'} made — Ctrl+Z to undo`, 'success');
}
function onProof(e) {
  const pf = e.target.closest('[data-pf]'); if (pf) { D.pf = pf.dataset.pf; return renderSide(); }
  const fx = e.target.closest('[data-pfix]'); if (fx) return fixAll(fx.dataset.pfix);
  if (e.target.closest('[data-unign]')) { LINT.ignored.clear(); paintIssues(); return renderSide(); }
  const card = e.target.closest('.wcard'); if (!card) return; const x = D.plist[+card.dataset.k]; if (!x) return;
  const s = e.target.closest('[data-s]'); if (s) return fixIssue(x.p, x.is, x.is.sug[+s.dataset.s]);
  if (e.target.closest('[data-ign]')) { LINT.ignored.add(x.is.kind + ':' + x.is.word.toLowerCase()); paintIssues(); return renderSide(); }
  if (e.target.closest('[data-dict]')) { X.writing.dictAdd(x.is.word).then(() => { LINT.cache.clear(); allParas().forEach(q => { q._lt = null; }); lintAll(); toast(`“${x.is.word}” added to your dictionary`, 'success'); }); return; }
  const P = allParas(); setSel({ a: [P.indexOf(x.p), x.is.a], b: [P.indexOf(x.p), x.is.b] }, true); setTimeout(() => showIssuePop(x.p, x.is), 50);
}

/* ───────── Outline: the headings, where the cursor is, one click to go there ───────── */
function headOfCursor(hs) { const s = getSel() || D.lastSel; if (!s) return null; const p = allParas()[s.a[0]]; let cur = null; for (const h of hs) { if (p && (h === p || (h.compareDocumentPosition(p) & Node.DOCUMENT_POSITION_FOLLOWING))) cur = h; } return cur; }
function renderOutline() {
  const hs = [...docEl.querySelectorAll('p[data-h]')].filter(p => ptext(p).trim()); D.olist = hs;
  if (!hs.length) { pane.innerHTML = '<div class="empty"><b>No headings yet</b><br>Give titles a Heading style (Ctrl+Alt+1, 2, 3), or let Auto Format find them.</div><button class="btn primary wfull" data-c="autofmt"><i data-icon="wand"></i><span>Auto Format…</span></button>'; return; }
  const cur = headOfCursor(hs), H = tw2pt(D.pg.h - Math.abs(D.pg.t) - Math.abs(D.pg.b)) * 96 / 72, pg = p => Math.max(1, Math.floor((p.offsetTop - docEl.offsetTop) / Math.max(1, H)) + 1);
  pane.innerHTML = `<div class="wsum"><div class="wsn"><b>${hs.length}</b><span>heading${hs.length > 1 ? 's' : ''}<small>Click one to go there</small></span></div><div class="wsa"><button class="btn line sm" data-c="toc" title="Insert a table of contents at the cursor">Table of contents</button></div></div>
    <div class="dx-outline">${hs.map((p, i) => `<button data-o="${i}" class="lv${p.dataset.h}${p === cur ? ' on' : ''}" style="padding-left:${10 + Math.max(0, +p.dataset.h - 1) * 14}px"><span>${esc(ptext(p).replace(/\n|￼/g, ' ').trim().slice(0, 90))}</span><small>${pg(p)}</small></button>`).join('')}</div>`;
}

/* ───────── Stats: counts, readability, and the sentences that are hard to read ───────── */
function renderStats() {
  const long = []; allParas().forEach((p, pi) => { if (!p.closest('#dxDoc')) return; const t = ptext(p); let m; const re = /[^.!?\n]+[.!?]*/g; while ((m = re.exec(t))) { const w = (m[0].match(/[\p{L}\p{N}'’]+/gu) || []).length; if (w > 25) long.push({ pi, a: m.index, b: m.index + m[0].length, w, s: m[0].trim() }); } });
  long.sort((a, b) => b.w - a.w); D.slong = long;
  pane.innerHTML = X.writing.statsHtml(paraTexts()) + (long.length ? `<h4 class="wh">Hard-to-read sentences <small>${long.length} over 25 words</small></h4><div class="wcards">${long.slice(0, 8).map((x, k) => `<div class="wcard wm-style" data-ls="${k}"><div class="wck wm-style">${x.w} words</div><div class="wctx">${esc(x.s.length > 160 ? x.s.slice(0, 160) + '…' : x.s)}</div><div class="wps"><button data-lsai="${k}" title="Let the offline AI shorten this sentence">Shorten with AI</button></div></div>`).join('')}</div>` : '');
}

pane.addEventListener('click', e => {
  if (D.tab === 'proof') return onProof(e);
  if (D.tab === 'outline') { const b = e.target.closest('[data-o]'); if (b) { const p = D.olist[+b.dataset.o], P = allParas(); setSel({ a: [P.indexOf(p), 0], b: [P.indexOf(p), 0] }); scroller.scrollTop += p.getBoundingClientRect().top - scroller.getBoundingClientRect().top - 30; X.flash(p, 'wflash'); $$('.dx-outline button', pane).forEach(x => x.classList.toggle('on', x === b)); } return; }
  if (D.tab === 'stats') { const ai = e.target.closest('[data-lsai]'), c = e.target.closest('[data-ls]'), x = c && D.slong[+c.dataset.ls]; if (!x) return; setSel({ a: [x.pi, x.a], b: [x.pi, x.b] }, true); if (ai) { setTab('human'); aiGo('shorter'); } return; }
  if (D.tab === 'human') return onHuman(e);
  if (D.tab === 'format') return onFormat(e);
});

/* ───────── Humanize: the offline AI (rewrite, tone, length, summaries) and instant rule-based fixes ───────── */
D.tone = 'natural'; D.ai = null; // D.ai: { task, segs: [{ p, a, b, from, to, refused, done }], i, running, insert }
D.aiScope = store.get('xd3.aiscope') || 'auto';
function aiSegments(task) { // what the AI works on: the selection, else the paragraph with the cursor; or every paragraph of the document
  const P = allParas(), body = P.filter(p => p.closest('#dxDoc') && !p.closest('.dx-sdt'));
  if (D.aiScope === 'doc') {
    if (task === 'summary' || task === 'bullets' || task === 'continue') { const ps = body.filter(p => ptext(p).trim()); if (!ps.length) return []; const last = ps[ps.length - 1]; return [{ p: last, a: 0, b: ptext(last).length, from: ps.map(p => ptext(p).replace(/￼/g, '').trim()).join('\n'), whole: true }]; }
    return body.filter(p => !p.dataset.h && (ptext(p).match(/[\p{L}\p{N}]+/gu) || []).length >= 4).map(p => ({ p, a: 0, b: ptext(p).length, from: ptext(p) }));
  }
  const s = getSel() || D.lastSel; if (!s) return [];
  if (s.a[0] === s.b[0] && s.a[1] === s.b[1] || s.c) { const p = P[s.a[0]]; return p ? [{ p, a: 0, b: ptext(p).length, from: ptext(p) }] : []; }
  const out = []; for (let i = s.a[0]; i <= s.b[0]; i++) { const p = P[i]; if (!p) continue; const t = ptext(p), a = i === s.a[0] ? s.a[1] : 0, b = i === s.b[0] ? s.b[1] : t.length; if (t.slice(a, b).trim()) out.push({ p, a, b, from: t.slice(a, b) }); }
  return out;
}
const insertsNew = t => t === 'summary' || t === 'bullets' || t === 'continue';
async function aiGo(task) {
  if (D.ai && D.ai.running) return toast('The AI is still working — Stop it first', 'info');
  const segs = aiSegments(task).map(s => ({ ...s, from: s.from.replace(/￼/g, '') })).filter(s => s.from.trim());
  if (!segs.length) return toast('Select some text, or put the cursor in a paragraph, first', 'info');
  D.ai = { task, segs, i: 0, running: true }; renderHuman();
  try {
    for (let i = 0; i < segs.length && D.ai.running; i++) {
      const s = segs[i]; D.ai.i = i; s.to = ''; s.live = true; drawAi();
      const r = await X.ai.run(task, s.from, { onToken: t => { s.to += t; drawAi(true); } });
      s.live = false; s.to = r.text; s.refused = r.refused; s.done = true; drawAi();
    }
  } catch (e) { toast(e.message, true); } finally { if (D.ai) D.ai.running = false; drawAi(); }
}
function aiApply(s) { // the new wording replaces the old word by word, so formatting of unchanged words stays; summaries / lists / more text go below
  const task = D.ai.task;
  if (insertsNew(task)) {
    const lines = task === 'bullets' ? s.to.split('\n').filter(Boolean) : task === 'continue' ? null : s.to.split(/\n+/).filter(Boolean);
    if (task === 'continue') { const t = ptext(s.p), at = s.whole ? t.length : s.b; replaceText(s.p, at, at, (/\s$/.test(t.slice(0, at)) ? '' : ' ') + s.to.trim()); return; }
    let after = s.p; for (const l of lines) { const q = emptyP(); q.dataset.pf = s.p.dataset.pf || '{}'; q.replaceChildren(document.createTextNode(l)); after.after(q); after = q; }
    return;
  }
  const t = ptext(s.p); if (t.slice(s.a, s.b).replace(/￼/g, '') !== s.from) throw new Error('changed');
  for (const c of X.writing.changes(s.from, s.to).reverse()) replaceText(s.p, s.a + c.a, s.a + c.b, c.rep);
}
function aiAccept(list) {
  snapshot(); let n = 0;
  for (const s of [...list].sort((x, y) => (x.p === y.p ? y.a - x.a : 0))) { if (!s.done || s.refused || s.applied) continue; try { aiApply(s); s.applied = true; n++; } catch (e) { s.refused = 'the text has changed since'; } }
  tidy(); afterEdit(true); drawAi(); if (n) toast(`${n} change${n > 1 ? 's' : ''} made — Ctrl+Z to undo`, 'success');
}
const AI_WHY = { 'no change': 'The AI found nothing to change.', 'changed too much': 'The AI changed too much here, so its answer was not used.', 'too long': 'The answer was too long, so it was not used.', 'off topic': 'The answer drifted off topic, so it was not used.', 'no answer': 'No answer this time — try again.' };
function aiCard(s, k) {
  const task = D.ai.task, ins = insertsNew(task), diff = !ins && s.done && !s.refused ? X.writing.changes(s.from, s.to) : null;
  let shown = esc(s.to || '');
  const churn = diff ? diff.reduce((n, c) => n + (c.b - c.a) + c.rep.length, 0) / Math.max(1, s.from.length + s.to.length) : 1;
  if (diff && diff.length && churn < 0.4) { let o = '', at = 0; for (const c of diff) { o += esc(s.from.slice(at, c.a)) + (c.b > c.a ? `<del>${esc(s.from.slice(c.a, c.b))}</del>` : '') + (c.rep ? `<ins>${esc(c.rep)}</ins>` : ''); at = c.b; } shown = o + esc(s.from.slice(at)); }
  return `<div class="wcard airesult${s.applied ? ' done' : ''}" data-ai-k="${k}">
    ${ins ? '' : `<div class="aiq">${esc(s.from.length > 220 ? s.from.slice(0, 220) + '…' : s.from)}</div>`}
    <div class="aia${s.live ? ' live' : ''}">${shown || (s.live ? '<span class="mut">Thinking…</span>' : '')}${s.live ? '<i class="caret"></i>' : ''}</div>
    ${s.refused && s.done ? `<div class="wcm">${esc(AI_WHY[s.refused] || 'Not used: ' + s.refused)}</div>` : ''}
    ${s.done ? `<div class="wps">${s.applied ? '<span class="mut">✓ Applied</span>' : s.refused ? '' : `<button data-aiacc="${k}">${ins ? (task === 'continue' ? 'Add to the text' : 'Insert below') : 'Replace'}</button>`}<button data-aiagain="${k}">Try again</button>${s.to ? `<button data-aicopy="${k}">Copy</button>` : ''}${s.applied ? '' : `<button data-aiskip="${k}">Discard</button>`}</div>` : ''}
  </div>`;
}
function drawAi(live) {
  const host = $('#dxAiOut'); if (!host || !D.ai) return;
  if (live) { const s = D.ai.segs[D.ai.i], c = host.querySelector(`[data-ai-k="${D.ai.i}"] .aia`); if (c && s) { c.textContent = s.to; const i = document.createElement('i'); i.className = 'caret'; c.appendChild(i); return; } }
  const { segs, running, task } = D.ai, open = segs.filter(s => s.done && !s.refused && !s.applied).length, T = X.ai.TASKS[task];
  host.innerHTML = `<div class="aibar"><b>${esc(T.label)}</b><span class="mut">${running ? (segs.length > 1 ? `paragraph ${D.ai.i + 1} of ${segs.length}…` : 'writing…') : segs.length > 1 ? `${segs.filter(s => s.done).length} of ${segs.length} done` : ''}</span>
    ${running ? '<button class="btn line sm" data-aistop>Stop</button>' : open > 1 ? `<button class="btn primary sm" data-aiall>Accept all ${open}</button>` : ''}</div>${X.ai.statusHtml()}
    <div class="wcards">${segs.map((s, k) => s.done || s.live ? aiCard(s, k) : '').join('')}</div>`;
}
function renderHuman() {
  const list = X.writing.rewrites(D.tone, [...docEl.querySelectorAll('p')].map(p => ({ text: ptext(p) }))), P = [...docEl.querySelectorAll('p')]; D.hlist = list.map(r => ({ ...r, p: P[r.i] }));
  const sel = getSel() || D.lastSel, hasSel = sel && !sel.c && !(sel.a[0] === sel.b[0] && sel.a[1] === sel.b[1]);
  pane.innerHTML = `<section class="aibox"><div class="aihd"><i data-icon="sparkle"></i><div><b>AI writing assistant</b><small>Runs offline on this device — nothing is uploaded</small></div></div>
      <label class="aiscope">Work on <select id="dxAiScope"><option value="auto"${D.aiScope === 'auto' ? ' selected' : ''}>${hasSel ? 'the selected text' : 'the paragraph with the cursor'}</option><option value="doc"${D.aiScope === 'doc' ? ' selected' : ''}>the whole document</option></select></label>
      <div class="aigrid">${Object.entries(X.ai.TASKS).map(([k, t]) => `<button data-ai="${k}" title="${esc(t.hint)}"><i data-icon="${t.icon}"></i><span>${esc(t.label)}</span></button>`).join('')}</div>
      <div id="dxAiOut"></div></section>
    <h4 class="wh">Quick fixes <small>instant, no AI</small></h4>
    <div class="wtone">${[['natural', 'Natural'], ['casual', 'Casual'], ['formal', 'Formal']].map(([v, l]) => `<button data-tone="${v}" class="${D.tone === v ? 'on' : ''}">${l}</button>`).join('')}</div>
    ${list.length ? `<button class="btn line wfull" data-hall>Accept all ${list.length} quick fix${list.length > 1 ? 'es' : ''}</button>` : ''}
    <div class="wcards">${list.length ? D.hlist.slice(0, 300).map((r, k) => `<div class="wcard wrw" data-k="${k}"><div class="wcw"><s>${esc(r.from)}</s> → <b>${r.to ? esc(r.to) : '<i>remove</i>'}</b></div><div class="wcm">${esc(r.why)}</div><div class="wps"><button data-acc="${k}">Accept</button><button data-skip="${k}">Skip</button></div></div>`).join('') : '<div class="empty">No quick fixes in this tone.</div>'}</div>`;
  $('#dxAiScope').onchange = e => { D.aiScope = e.target.value; store.set('xd3.aiscope', D.aiScope); };
  drawAi(); icons(pane);
}
document.addEventListener('xd3:ai', () => { if (D.open && D.tab === 'human') drawAi(); });
function onHuman(e) {
  const ai = e.target.closest('[data-ai]'); if (ai) return aiGo(ai.dataset.ai);
  if (e.target.closest('[data-aistop]')) { X.ai.stop(); if (D.ai) D.ai.running = false; return; }
  if (e.target.closest('[data-aiall]')) return aiAccept(D.ai.segs);
  const k = el => D.ai && D.ai.segs[+el.dataset[Object.keys(el.dataset).find(x => x.startsWith('ai'))]];
  const acc = e.target.closest('[data-aiacc]'); if (acc) return aiAccept([k(acc)]);
  const sk = e.target.closest('[data-aiskip]'); if (sk) { const s = k(sk); s.refused = 'discarded'; s.applied = false; sk.closest('.wcard').remove(); return; }
  const cp = e.target.closest('[data-aicopy]'); if (cp) { navigator.clipboard.writeText(k(cp).to).then(() => toast('Copied'), () => toast('Copy is not available here', true)); return; }
  const ag = e.target.closest('[data-aiagain]'); if (ag) { const s = k(ag); if (D.ai.running) return; s.done = false; s.live = true; s.to = ''; s.refused = ''; s.applied = false; D.ai.i = D.ai.segs.indexOf(s); D.ai.running = true; drawAi();
    X.ai.run(D.ai.task, s.from, { retry: true, onToken: t => { s.to += t; drawAi(true); } }).then(r => { s.to = r.text; s.refused = r.refused; }, err => toast(err.message, true)).finally(() => { s.live = false; s.done = true; D.ai.running = false; drawAi(); }); return; }
  const tn = e.target.closest('[data-tone]'); if (tn) { D.tone = tn.dataset.tone; return renderHuman(); }
  const acc2 = e.target.closest('[data-acc]'), skip = e.target.closest('[data-skip]');
  if (acc2) { const r = D.hlist[+acc2.dataset.acc]; if (r && ptext(r.p).slice(r.a, r.b) === r.from) { snapshot(); replaceText(r.p, r.a, r.b, r.to); restyle([rootOf(r.p)]); afterEdit(); } return renderHuman(); }
  if (skip) return skip.closest('.wcard').remove();
  if (e.target.closest('[data-hall]')) { snapshot(); let n = 0; for (const r of [...D.hlist].sort((x, y) => x.i - y.i || y.a - x.a)) if (ptext(r.p).slice(r.a, r.b) === r.from) { replaceText(r.p, r.a, r.b, r.to); n++; } restyle(); afterEdit(); renderHuman(); toast(`${n} change(s) made — Ctrl+Z to undo`, 'success'); return; }
  const c = e.target.closest('.wrw'); if (c) { const r = D.hlist[+c.dataset.k], P = allParas(); setSel({ a: [P.indexOf(r.p), r.a], b: [P.indexOf(r.p), r.b] }, true); }
}

/* ───────── Fix Formatting: what is inconsistent, and one click to fix it ───────── */
const NUMRX = /^\s*(\(?(\d{1,3}|[a-zA-Z]|[ivxIVX]{1,5})[.)])\s+/, BULRX = /^\s*([•●▪■◦○◆♦✓✔➢►*·]|-|–|—)\s+/;
const jcKey = v => v === 'both' || v === 'distribute' ? 'both' : v === 'center' ? 'center' : v === 'right' || v === 'end' ? 'right' : 'left';
function bodyParas() { return [...docEl.querySelectorAll('p')].filter(p => !p.closest('table') && !p.closest('.dx-sdt')); }
function stats() { // the document's usual body font, size, alignment and spacing
  const font = {}, size = {}, jc = {}, sp = {};
  for (const p of bodyParas()) { const e = p._pe || effP(p); if (styleLevel(e.style) != null) continue; const t = ptext(p).trim(); if (t.length < 2) continue;
    for (const r of p.querySelectorAll('span[data-rf]')) { const x = r._e || effR(rfOf(r), e.style), n = r.textContent.length, f = fontName(x.font); font[f] = (font[f] || 0) + n; if (!x.va) size[x.sz || 10] = (size[x.sz || 10] || 0) + n; }
    if (t.length > 120) { const k = jcKey(e.jc); jc[k] = (jc[k] || 0) + 1; }
    const k = canon(e.sp || {}); sp[k] = (sp[k] || 0) + 1; }
  const top = o => { const e = Object.entries(o).sort((a, b) => b[1] - a[1]); return e.length ? { v: e[0][0], n: e[0][1], all: e.reduce((a, b) => a + b[1], 0) } : null; };
  return { font: top(font), size: top(size), jc: top(jc), sp: top(sp) };
}
const runsOf = (ps, ok = () => true) => ps.flatMap(p => [...p.querySelectorAll('span[data-rf]')].filter(r => r.textContent.trim() && ok(r, p)));
const isLinkRun = r => !!r.closest('a[data-w]') || /hyperlink/i.test(rfOf(r).rStyle || '');
const monoFont = f => /courier|consol|mono|symbol|wingdings|webdings/i.test(f);
function headingGuess(st) { // short bold or large paragraphs, followed by ordinary text, that are not styled as headings
  const ps = bodyParas(), dom = st.size ? +st.size.v : 11, cand = new Map();
  const info = p => { const t = ptext(p).replace(/\ufffc/g, '').trim(), rs = [...p.querySelectorAll('span[data-rf]')].filter(r => r.textContent.trim()), e = p._pe || effP(p); return { t, rs, e, words: t ? t.split(/\s+/).length : 0, max: Math.max(0, ...rs.map(r => (r._e || {}).sz || 10)), bold: rs.length && rs.every(r => (r._e || {}).b), caps: t.length > 3 && t === t.toUpperCase() && /\p{L}/u.test(t) }; };
  const I = ps.map(info), nextText = i => { for (let j = i + 1; j < I.length; j++) if (I[j].t) return j; return -1; };
  I.forEach((x, i) => {
    if (!x.t || x.words > 12 || x.t.length > 90 || styleLevel(x.e.style) != null || (x.e.num && x.e.num.id !== '0') || /[.,;!?]$/.test(x.t) || NUMRX.test(x.t) && !/^\d+(\.\d+)*\.?\s+\p{Lu}/u.test(x.t)) return;
    if (!(x.bold || x.max >= dom + 1.5 || (x.caps && x.words <= 8))) return;
    const j = nextText(i); if (j < 0) return; cand.set(i, x);
  });
  const listy = y => (y.e.num && y.e.num.id !== '0') || BULRX.test(y.t) || NUMRX.test(y.t); // (a heading may be followed by a list)
  const real = new Map(); for (const [i, x] of cand) { const j = nextText(i); if (j >= 0 && (I[j].words >= 6 || cand.has(j) || listy(I[j]))) real.set(i, x); }
  const firstI = I.findIndex(y => y.t), top = Math.max(0, ...[...real.values()].map(x => x.max)); // the first line, alone in the biggest size, is the title
  const titleI = real.size > 1 && real.has(firstI) && real.get(firstI).max === top && [...real.values()].filter(x => x.max === top).length === 1 ? firstI : -1;
  const sizes = [...new Set([...real].filter(([i]) => i !== titleI).map(([, x]) => x.max))].sort((a, b) => b - a);
  return [...real].map(([i, x]) => ({ p: ps[i], title: i === titleI, lvl: Math.min(3, sizes.indexOf(x.max) + 1) || 1 }));
}
function manualLists() { // typed "• " or "1." markers on two or more paragraphs in a row
  const ps = bodyParas(), out = []; let group = [];
  const flush = () => { if (group.length >= 2) out.push(group); group = []; };
  for (const p of ps) { const t = ptext(p), e = p._pe || effP(p); const m = !(e.num && e.num.id !== '0') && (BULRX.exec(t) || NUMRX.exec(t)); if (!m) { if (t.trim()) flush(); continue; } const kind = BULRX.test(t) ? 'bullet' : /^\s*\(?\d/.test(t) ? 'decimal' : /^\s*\(?[ivx]+[.)]/.test(t) && !/^\s*\(?[a-h][.)]/.test(t) ? 'lowerRoman' : /^\s*\(?[A-Z]/.test(t) ? 'upperLetter' : 'lowerLetter'; if (group.length && group[0].kind !== kind) flush(); group.push({ p, kind, len: m[0].length }); }
  flush(); return out;
}
function checks() {
  const st = stats(), ps = bodyParas(), out = [], domF = st.font && st.font.v, domS = st.size && +st.size.v;
  const add = (key, title, detail, targets, fix) => { if (targets.length) out.push({ key, title, detail, targets, fix }); };
  const heads = headingGuess(st);
  add('headings', 'Headings typed as plain text', `${heads.length} short bold or large line(s) look like headings but use the body style — real heading styles give the document an outline, a table of contents and a consistent look.`, heads.map(h => h.p), () => {
    for (const h of heads) { const sid = ensureStyle(h.title ? 'title' : 'heading ' + h.lvl); if (!sid) continue; const pf = { ...pfOf(h.p), style: sid }; delete pf.jc; setPf(h.p, pf); for (const r of h.p.querySelectorAll('span[data-rf]')) { const rf = { ...rfOf(r) }; for (const k of ['font', 'sz', 'b', 'color', 'caps']) delete rf[k]; setRf(r, rf); } }
  });
  const hdf = runsOf([...docEl.querySelectorAll('p[data-h]')], r => { const rf = rfOf(r); return rf.font !== undefined || rf.sz !== undefined || rf.color !== undefined; });
  add('headdirect', 'Headings with their own fonts or sizes', `${hdf.length} part(s) of headings override the heading style, so headings look different from each other.`, hdf, () => hdf.forEach(r => { const rf = { ...rfOf(r) }; delete rf.font; delete rf.sz; delete rf.color; setRf(r, rf); }));
  const lists = manualLists();
  add('lists', 'Lists typed by hand', `${lists.reduce((a, g) => a + g.length, 0)} paragraph(s) start with typed bullets or numbers — real lists keep their indents and renumber themselves.`, lists.flat().map(x => x.p), () => {
    for (const g of lists) { const id = addList(g[0].kind); for (const x of g) { replaceText(x.p, 0, x.len, ''); const pf = { ...pfOf(x.p), num: { id, lvl: 0 } }; delete pf.ind; setPf(x.p, pf); } }
  });
  const headP = new Set(heads.map(h => h.p)), bodyRuns = runsOf(ps, (r, p) => styleLevel((p._pe || effP(p)).style) == null && !headP.has(p) && !isLinkRun(r));
  const badF = domF ? bodyRuns.filter(r => { const f = fontName((r._e || {}).font); return f !== domF && !monoFont(f); }) : [];
  add('fonts', 'Mixed fonts in the text', `${badF.length} part(s) of the body text are not in ${domF} like the rest.`, badF.map(r => paraOf(r)), () => badF.forEach(r => { const rf = { ...rfOf(r) }; delete rf.font; setRf(r, rf); if (fontName(effR(rf, effP(paraOf(r)).style).font) !== domF) setRf(r, { ...rf, font: domF }); }));
  const badS = domS ? bodyRuns.filter(r => { const e = r._e || {}; return !e.va && Math.abs((e.sz || 10) - domS) >= 0.5; }) : [];
  add('sizes', 'Mixed text sizes', `${badS.length} part(s) of the body text are not ${domS} pt like the rest.`, badS.map(r => paraOf(r)), () => badS.forEach(r => { const rf = { ...rfOf(r) }; delete rf.sz; setRf(r, rf); if ((effR(rf, effP(paraOf(r)).style).sz || 10) !== domS) setRf(r, { ...rf, sz: domS }); }));
  const coloured = bodyRuns.filter(r => { const c = (r._e || {}).color; return c && c !== 'auto' && c !== '000000' && rfOf(r).color; }), colN = coloured.reduce((a, r) => a + r.textContent.length, 0), allN = bodyRuns.reduce((a, r) => a + r.textContent.length, 0) || 1;
  if (colN / allN < 0.15) add('colors', 'Stray text colours', `${coloured.length} part(s) of the body text have a colour of their own (often left over from copying).`, coloured.map(r => paraOf(r)), () => coloured.forEach(r => { const rf = { ...rfOf(r) }; delete rf.color; setRf(r, rf); }));
  const hl = runsOf(ps, r => { const e = r._e || {}; return (e.hl && e.hl !== 'none') || rfOf(r).shd; });
  add('highlight', 'Highlighted text', `${hl.length} part(s) are highlighted — remove the highlights if they were only notes.`, hl.map(r => paraOf(r)), () => hl.forEach(r => { const rf = { ...rfOf(r) }; delete rf.hl; delete rf.shd; setRf(r, rf); }));
  const sp = []; for (const p of [...docEl.querySelectorAll('p')]) { const t = ptext(p); for (const m of t.matchAll(/(?<=\S) {2,}(?=\S)| +(?=[,.;:!?](\s|$))|^[ \u00a0]+(?=\S)| +$/gm)) sp.push({ p, a: m.index, b: m.index + m[0].length, rep: /^ {2,}$/.test(m[0]) && m.index > 0 && /\S/.test(t[m.index - 1]) && /\S/.test(t[m.index + m[0].length] || '') ? ' ' : '' }); }
  add('spaces', 'Extra spaces', `${sp.length} double space(s), space(s) before punctuation or at the start or end of a line.`, sp.map(x => x.p), () => { for (const x of [...sp].reverse()) replaceText(x.p, x.a, x.b, x.rep); });
  const tabs = ps.filter(p => /^\t+|^ {2,}/.test(ptext(p)) && !(p._pe && p._pe.num && p._pe.num.id !== '0') && ptext(p).trim());
  add('indent', 'Indents made with tabs or spaces', `${tabs.length} paragraph(s) start with tabs or spaces instead of a first-line indent.`, tabs, () => tabs.forEach(p => { const n = /^[\t ]+/.exec(ptext(p))[0].length; replaceText(p, 0, n, ''); const pf = { ...pfOf(p) }; pf.ind = { ...(pf.ind || {}), first: 720 }; delete pf.ind.hang; setPf(p, pf); }));
  const empties = []; let run = [];
  const endRun = () => { if (run.length >= 2 || (run.length === 1 && run[0].dataset.mid)) empties.push(...(run.length >= 2 ? run.slice(1) : run)); run = []; };
  docEl.querySelectorAll(':scope > p').forEach(p => { const isEmpty = !ptext(p).replace(/[\n\s]/g, '').length && !p.classList.contains('dx-sect') && !p.querySelector('.dx-a:not(.dx-z)'); if (isEmpty) { const prev = p.previousElementSibling; if (!prev || prev.tagName !== 'TABLE') run.push(p); } else endRun(); }); endRun();
  add('empty', 'Empty lines used for spacing', `${empties.length} empty paragraph(s) — paragraph spacing keeps gaps even and pages tidy.`, empties, () => empties.forEach(p => { if (p.parentNode.children.length > 1) p.remove(); }));
  if (st.jc && st.jc.n / st.jc.all >= 0.7 && st.jc.all >= 3) { const want = st.jc.v, odd = ps.filter(p => { const e = p._pe || effP(p), t = ptext(p).trim(); return t.length > 120 && styleLevel(e.style) == null && jcKey(e.jc) !== want && jcKey(e.jc) !== 'center'; });
    add('align', 'Mixed alignment', `${odd.length} paragraph(s) are ${want === 'both' ? 'not justified like the rest' : 'aligned differently from the rest'}.`, odd, () => odd.forEach(p => setPf(p, { ...pfOf(p), jc: want === 'left' ? undefined : want }))); }
  const spOdd = ps.filter(p => { const pf = pfOf(p), e = p._pe || effP(p); return pf.sp && styleLevel(e.style) == null && ptext(p).trim() && st.sp && canon(e.sp || {}) !== st.sp.v; });
  add('spacing', 'Uneven paragraph spacing', `${spOdd.length} paragraph(s) have their own line or paragraph spacing.`, spOdd, () => spOdd.forEach(p => { const pf = { ...pfOf(p) }; delete pf.sp; setPf(p, pf); }));
  const brk = ps.filter(p => { const t = ptext(p); return (t.match(/\n/g) || []).length >= 2 && t.length > 80 && t.split('\n').slice(0, -1).filter(l => l.trim() && !/[.!?:;]$/.test(l.trim())).length >= 1; });
  add('breaks', 'Lines broken by hand', `${brk.length} paragraph(s) are broken into lines with line breaks (often pasted from a PDF or e-mail).`, brk, () => brk.forEach(p => { for (const b of [...p.querySelectorAll('br')].slice(0, -1)) b.replaceWith(document.createTextNode(' ')); fixPara(p); const t = ptext(p); for (const m of [...t.matchAll(/ {2,}/g)].reverse()) replaceText(p, m.index, m.index + m[0].length, ' '); }));
  const quotes = []; let curly = 0, straight = 0; for (const p of [...docEl.querySelectorAll('p')]) { const t = ptext(p); curly += (t.match(/[“”‘’]/g) || []).length; straight += (t.match(/["']/g) || []).length; }
  if (curly && straight) { for (const p of [...docEl.querySelectorAll('p')]) if (/["']/.test(ptext(p))) quotes.push(p); add('quotes', 'Mixed straight and curly quotes', `${straight} straight quote mark(s) next to ${curly} curly ones.`, quotes, () => quotes.forEach(p => smartQuotes(p))); }
  return out;
}
function smartQuotes(p) { const t = ptext(p), out = []; for (let i = 0; i < t.length; i++) { const c = t[i]; if (c !== '"' && c !== "'") continue; const prev = t[i - 1] || ' ', open = /[\s(\[{\u2014\u2013-]/.test(prev); out.push([i, c === '"' ? (open ? '“' : '”') : (open ? '‘' : '’')]); } for (const [i, c] of out.reverse()) replaceText(p, i, i + 1, c); }
function punctuation(p) { const t = ptext(p), rep = []; for (const m of t.matchAll(/ -- | - (?=\p{L})|\.\.\.|(?<=\p{L})--(?=\p{L})/gu)) rep.push([m.index, m.index + m[0].length, m[0] === '...' ? '…' : m[0].trim() === '--' && m[0][0] !== ' ' ? '—' : ' – ']); for (const [a, b, r] of rep.reverse()) replaceText(p, a, b, r); return rep.length; }
function renderFormat() {
  const list = checks(); D.fchecks = list; $('#dxFcnt').textContent = list.length || '';
  pane.innerHTML = `${list.length ? `<div class="wsum"><div class="wsn"><b>${list.length}</b><span>kind${list.length > 1 ? 's' : ''} of inconsistent formatting<small>Fix one kind at a time, or all at once — Ctrl+Z undoes it</small></span></div></div>` : ''}
    ${list.length ? `<button class="btn primary wfull" data-fall>Fix all ${list.length} kind(s)</button>` : ''}
    <div class="wcards">${list.length ? list.map((c, k) => `<div class="wcard dx-fcard" data-f="${k}"><div class="wck">${esc(c.title)} <b class="dx-fn">${c.targets.length}</b></div><div class="wcm">${esc(c.detail)}</div><div class="wps"><button data-fix="${k}">Fix</button><button data-show="${k}">Show</button></div></div>`).join('') : '<div class="empty">The formatting is consistent. ✓<br><br>Auto Format can still restyle the whole document.</div>'}</div>
    <button class="btn line wfull" data-c="autofmt"><i data-icon="wand"></i><span>Auto Format…</span></button>`;
  icons(pane);
}
function onFormat(e) {
  const fx = e.target.closest('[data-fix]'), sh = e.target.closest('[data-show]');
  if (e.target.closest('[data-fall]')) { applyFixes(D.fchecks); return; }
  if (fx) { applyFixes([D.fchecks[+fx.dataset.fix]]); return; }
  if (sh) { const c = D.fchecks[+sh.dataset.show]; c._i = ((c._i ?? -1) + 1) % c.targets.length; const p = c.targets[c._i]; if (p && p.isConnected) { const P = allParas(); setSel({ a: [P.indexOf(p), 0], b: [P.indexOf(p), 0] }); scroller.scrollTop += p.getBoundingClientRect().top - scroller.getBoundingClientRect().top - 60; X.flash(p, 'wflash'); } }
}
function applyFixes(list) { // one kind after the other, each looked at again after the one before (new headings are no longer body text)
  if (!list.length) return; snapshot(); let n = 0;
  for (const key of list.map(c => c.key)) { const c = checks().find(x => x.key === key); if (!c) continue; try { c.fix(); n += c.targets.length; } catch (err) { console.warn(key, err); } tidy(); }
  afterEdit(true); renderFormat(); toast(`Fixed ${n} item(s) — Ctrl+Z to undo`, 'success');
}

/* ───────── Auto Format ───────── */
const STYLESETS = {
  keep: null,
  office: { body: 'Calibri', head: 'Calibri Light', colors: ['2F5496', '2F5496', '1F3763'], title: '000000' },
  classic: { body: 'Times New Roman', head: 'Georgia', colors: ['000000', '000000', '000000'], title: '000000' },
  modern: { body: 'Segoe UI', head: 'Segoe UI Semibold', colors: ['1F2937', '1F2937', '374151'], title: '111827' },
  elegant: { body: 'Garamond', head: 'Garamond', colors: ['7F1D1D', '7F1D1D', '44403C'], title: '1C1917' },
  report: { body: 'Arial', head: 'Arial', colors: ['0F4C81', '0F4C81', '0F4C81'], title: '0F4C81' },
};
function styleSet(id, opt) { // rewrites the styles themselves, so every paragraph using them follows
  const set = (sid, fn) => { const s = sid && D.ST.byId[sid]; if (!s) return; let rPr = kid(s.el, 'rPr'); if (!rPr) { rPr = wel(D.ST.doc, 'rPr'); s.el.appendChild(rPr); } let pPr = kid(s.el, 'pPr'); fn(rPr, () => { if (!pPr) { pPr = wel(D.ST.doc, 'pPr'); const r = kid(s.el, 'rPr'); s.el.insertBefore(pPr, r); } return pPr; }); };
  const normal = D.ST.defPara || ensureStyle('normal'), ss = STYLESETS[id];
  set(normal, (rPr, pp) => {
    const font = opt.font !== 'keep' ? opt.font : ss ? ss.body : null; if (font) writeRProp(rPr, 'font', font);
    if (opt.size !== 'keep') writeRProp(rPr, 'sz', +opt.size);
    const pPr = pp(); if (opt.line !== 'keep' || opt.after !== 'keep') { const cur = readP(pPr).sp || {}, sp = { ...cur }; if (opt.line !== 'keep') { sp.line = +opt.line; sp.lineRule = 'auto'; } if (opt.after !== 'keep') sp.after = +opt.after; writePProp(pPr, 'sp', sp); }
    if (opt.jc !== 'keep') writePProp(pPr, 'jc', opt.jc === 'left' ? undefined : opt.jc);
  });
  if (ss) {
    [1, 2, 3].forEach(l => { const sid = ensureStyle('heading ' + l); set(sid, rPr => { writeRProp(rPr, 'font', ss.head); writeRProp(rPr, 'color', ss.colors[l - 1]); }); });
    const t = ensureStyle('title'); set(t, rPr => { writeRProp(rPr, 'font', ss.head); writeRProp(rPr, 'color', ss.title); });
  }
  D.ST = loadStyles(D.ST.doc); changedModel();
}
async function autoFormat() {
  if (!D.zip) return;
  const st = stats();
  const r = await dialog('Auto Format', `<p>Tidies the whole document in one go. Everything can be undone with Ctrl+Z.</p>
    <div class="dx-afgrid">
      <label class="inl"><input type="checkbox" name="headings" checked>Find headings and give them heading styles</label>
      <label class="inl"><input type="checkbox" name="lists" checked>Turn typed bullets and numbers into real lists</label>
      <label class="inl"><input type="checkbox" name="spaces" checked>Remove extra spaces</label>
      <label class="inl"><input type="checkbox" name="punct" checked>Smart quotes, dashes and ellipses (“ ” – … )</label>
      <label class="inl"><input type="checkbox" name="empty" checked>Remove empty lines used for spacing</label>
      <label class="inl"><input type="checkbox" name="consistent" checked>Make fonts, sizes, spacing and alignment consistent</label>
      <label class="inl"><input type="checkbox" name="indent" checked>Replace tab and space indents</label>
      <label class="inl"><input type="checkbox" name="breaks" checked>Join lines broken by hand</label>
      <label class="inl"><input type="checkbox" name="colors">Remove stray colours and highlights</label>
    </div>
    <h4 style="margin:12px 0 6px">Look</h4>
    <div class="row"><label>Style set<select name="set"><option value="keep">Keep the current look</option><option value="office">Office — Calibri, blue headings</option><option value="classic">Classic — Times New Roman, Georgia</option><option value="modern">Modern — Segoe UI</option><option value="elegant">Elegant — Garamond</option><option value="report">Report — Arial, navy headings</option></select></label>
    <label>Body font<select name="font"><option value="keep">Keep (${esc(st.font ? st.font.v : 'current')})</option>${['Calibri', 'Aptos', 'Arial', 'Cambria', 'Georgia', 'Times New Roman', 'Segoe UI', 'Verdana', 'Garamond'].map(f => `<option>${f}</option>`).join('')}</select></label></div>
    <div class="row"><label>Size<select name="size"><option value="keep">Keep${st.size ? ` (${st.size.v} pt)` : ''}</option>${[10, 10.5, 11, 12, 13, 14].map(s => `<option value="${s}">${s} pt</option>`).join('')}</select></label>
    <label>Line spacing<select name="line"><option value="keep">Keep</option><option value="240">1.0</option><option value="259">1.08</option><option value="276">1.15</option><option value="360">1.5</option><option value="480">2.0</option></select></label>
    <label>Space after<select name="after"><option value="keep">Keep</option><option value="0">0 pt</option><option value="120">6 pt</option><option value="160">8 pt</option><option value="200">10 pt</option><option value="240">12 pt</option></select></label>
    <label>Alignment<select name="jc"><option value="keep">Keep</option><option value="left">Left</option><option value="both">Justified</option></select></label></div>`, { ok: 'Auto Format', wide: true });
  if (!r) return;
  await busy(async () => {
    snapshot(); let n = 0; const bodyPs = () => [...docEl.querySelectorAll('p')];
    const runCheck = key => { const c = checks().find(x => x.key === key); if (c) { try { c.fix(); n += c.targets.length; } catch (e) { console.warn(key, e); } tidy(); } };
    if (r.spaces) runCheck('spaces');
    if (r.punct) { for (const p of bodyPs()) { const t = ptext(p); if (/["']/.test(t)) { smartQuotes(p); n++; } n += punctuation(p); } tidy(); }
    if (r.breaks) runCheck('breaks');
    if (r.headings) { runCheck('headings'); runCheck('headdirect'); } // (before lists: a heading is recognised by the text that follows it)
    if (r.lists) runCheck('lists');
    if (r.indent) runCheck('indent');
    if (r.empty) runCheck('empty');
    if (r.colors) { runCheck('colors'); runCheck('highlight'); }
    const custom = r.set !== 'keep' || r.font !== 'keep' || r.size !== 'keep' || r.line !== 'keep' || r.after !== 'keep' || r.jc !== 'keep';
    if (r.consistent || custom) {
      if (custom) { // the chosen look goes into the styles; direct formatting that fights it is cleared from the body text
        styleSet(r.set, { font: r.font, size: r.size, line: r.line, after: r.after, jc: r.jc });
        for (const p of bodyParas()) { if (styleLevel(effP(p).style) != null) continue; const pf = { ...pfOf(p) }; if (r.line !== 'keep' || r.after !== 'keep') delete pf.sp; if (r.jc !== 'keep') delete pf.jc; setPf(p, pf); for (const sp of p.querySelectorAll('span[data-rf]')) { if (isLinkRun(sp)) continue; const rf = { ...rfOf(sp) }; if (r.font !== 'keep' || r.set !== 'keep') delete rf.font; if (r.size !== 'keep') delete rf.sz; setRf(sp, rf); } n++; }
        tidy();
      }
      if (r.consistent) { for (const key of ['fonts', 'sizes', 'spacing', 'align']) runCheck(key); }
    }
    tidy(); afterEdit(true); setTab('format');
    toast(n ? `Auto Format made ${n} change(s) — Ctrl+Z to undo` : 'The document already looks tidy', 'success');
  }, 'Formatting…');
}

/* ───────── saving: the page → Word XML, only what changed is rewritten ───────── */
function fixRelIds(x, from, to) { // a piece moved between parts brings its pictures and links along
  if (!from || from === to) return; const fr = relsOf(from);
  for (const el of [x, ...x.getElementsByTagNameNS('*', '*')]) for (const a of [...el.attributes]) if (a.namespaceURI === NS.r) { const r = fr[a.value]; if (r) el.setAttributeNS(NS.r, a.name, addRel(to, r.type, r.ext ? r.target : resolveRel(to, resolvePath(from, r.target)), r.ext)); }
}
function rPrFor(span, part, doc, mark) {
  const rec = span && span.dataset.r && D.RP.get(span.dataset.r), rf = span ? rfOf(span) : mark || {};
  const rPr = rec && rec.el ? doc.importNode(rec.el, true) : wel(doc, 'rPr'), orig = rec ? rec.props : {};
  for (const k of R_KEYS) if (!same(orig[k], rf[k])) writeRProp(rPr, k, rf[k]);
  return rPr.childNodes.length ? rPr : null;
}
function emitInline(el, part, doc, target, ctx) {
  let cur = null; const close = () => { cur = null; };
  const runFor = span => { if (cur && cur.span === span) return cur.r; const r = wel(doc, 'r'), rp = rPrFor(span, part, doc, ctx.mark); if (rp) r.appendChild(rp); target.appendChild(r); cur = { span, r }; return r; };
  const text = (s, span) => {
    const r = runFor(span); let buf = '';
    const flush = () => { if (buf) { const t = wel(doc, 't'); t.setAttributeNS(NS.xml, 'xml:space', 'preserve'); t.textContent = buf; r.appendChild(t); buf = ''; } };
    for (const ch of s.replace(XML_BAD, '')) { if (ch === '\t') { flush(); r.appendChild(wel(doc, 'tab')); } else if (ch === '\u2011') { flush(); r.appendChild(wel(doc, 'noBreakHyphen')); } else if (ch === '\u00ad') { flush(); r.appendChild(wel(doc, 'softHyphen')); } else buf += ch; }
    flush();
  };
  const walk = (node, span) => {
    for (const c of node.childNodes) {
      if (c.nodeType === 3) { if (c.data) text(c.data, span); continue; }
      if (c.nodeType !== 1) continue;
      if (c.tagName === 'BR') { if (c !== ctx.placeholder) runFor(span).appendChild(wel(doc, 'br')); continue; }
      if (isAtom(c)) { const rec = D.AT.get(c.dataset.k); if (!rec) continue; close(); const x = doc.importNode(rec.el, true); fixRelIds(x, D.parts[rec.part], part); if (c.dataset.sw) resizeXml(x, +c.dataset.sw, +c.dataset.sh); target.appendChild(x); continue; }
      if (isRun(c)) { walk(c, c); continue; }
      if (isWrap(c)) {
        const rec = D.WR.get(c.dataset.w); if (!rec) { walk(c, span); continue; } close();
        const w = doc.importNode(rec.el, false); for (const ch of rec.el.children) if (/Pr$/.test(ch.localName)) w.appendChild(doc.importNode(ch, true));
        fixRelIds(w, D.parts[rec.part], part); let inner = w; if (rec.el.localName === 'sdt') { inner = wel(doc, 'sdtContent'); w.appendChild(inner); }
        emitInline(c, part, doc, inner, ctx); target.appendChild(w); continue;
      }
      walk(c, span);
    }
  };
  walk(el, null);
}
function resizeXml(x, w, h) { const cx = Math.round(w * 12700), cy = Math.round(h * 12700); for (const e of x.getElementsByTagNameNS('*', 'extent')) { e.setAttribute('cx', cx); e.setAttribute('cy', cy); } for (const e of x.getElementsByTagNameNS('*', 'ext')) if (e.parentNode && e.parentNode.localName === 'xfrm') { e.setAttribute('cx', cx); e.setAttribute('cy', cy); } }
function pToXml(p, part, doc, F) {
  const id = p.dataset.p, rec0 = id && D.PP.get(id), rec = rec0 && rec0.part === part.name ? rec0 : null, wp = wel(doc, 'p');
  if (rec && F.first.get(id) === p) for (const [ns, n, v] of rec.attrs) { try { wp.setAttributeNS(ns, n, v); } catch (e) { /* */ } }
  const pf = pfOf(p), pPr = rec && rec.el ? doc.importNode(rec.el, true) : wel(doc, 'pPr'), orig = rec ? rec.props : {};
  for (const k of P_KEYS) if (!same(orig[k], pf[k])) writePProp(pPr, k, pf[k]);
  if (rec && rec.x.sect && F.last.get(id) !== p) { const s = kid(pPr, 'sectPr'); if (s) s.remove(); } // (a split paragraph keeps its section break on the last half)
  if (pPr.childNodes.length) wp.appendChild(pPr);
  const U = units(p), last = U[U.length - 1];
  emitInline(p, part, doc, wp, { placeholder: last && last.t === 'b' ? last.n : null, mark: rec ? rec.mark : (p._mr ? rfOf({ dataset: { rf: p._mr.rf } }) : {}) });
  return wp;
}
function tblToXml(t, part, doc, F) {
  const rec = D.TB.get(t.dataset.t), orig = rec && rec.el, x = wel(doc, 'tbl');
  const rows = [...t.rows].filter(r => !r.classList.contains('dx-gaprow')), cols = Math.max(1, ...rows.map(r => [...r.cells].reduce((a, c) => a + (c.colSpan || 1), 0)));
  const tblPr = orig && kid(orig, 'tblPr'); x.appendChild(tblPr ? doc.importNode(tblPr, true) : parseInto(doc, `<w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>`));
  const g = orig && kid(orig, 'tblGrid'), gc = kids(g, 'gridCol'), keep = g && gc.length === cols;
  if (keep) x.appendChild(doc.importNode(g, true)); else { const total = gc.reduce((a, c) => a + (+wa(c, 'w') || 0), 0) || Math.round(contentWidthPt() * 20), ng = wel(doc, 'tblGrid'); for (let i = 0; i < cols; i++) ng.appendChild(wel(doc, 'gridCol', { w: Math.round(total / cols) })); x.appendChild(ng); }
  for (const tr of rows) {
    const rr = D.TB.get(tr.dataset.tr), xr = wel(doc, 'tr'); if (rr) for (const ch of rr.el.children) if (ch.localName === 'trPr' || ch.localName === 'tblPrEx') xr.appendChild(doc.importNode(ch, true));
    for (const td of tr.cells) {
      const cr = D.TB.get(td.dataset.tc), xc = wel(doc, 'tc'), op = cr && kid(cr.el, 'tcPr'), tcPr = op ? doc.importNode(op, true) : wel(doc, 'tcPr');
      for (let e; (e = kid(tcPr, 'gridSpan'));) e.remove(); if (td.colSpan > 1) insertOrdered(tcPr, wel(doc, 'gridSpan', { val: td.colSpan }), TC_ORDER);
      if (!keep) for (let e; (e = kid(tcPr, 'tcW'));) e.remove();
      if (tcPr.childNodes.length) xc.appendChild(tcPr);
      const blocks = blocksToXml(td, part, doc, F); if (!blocks.length || blocks[blocks.length - 1].localName !== 'p') blocks.push(wel(doc, 'p'));
      for (const b of blocks) xc.appendChild(b); xr.appendChild(xc);
    }
    if (kids(xr, 'tc').length) x.appendChild(xr);
  }
  return kids(x, 'tr').length ? x : null;
}
const parseInto = frag;
function blocksToXml(boxEl, part, doc, F) {
  const out = [];
  for (const c of boxEl.children) {
    if (c.tagName === 'P') out.push(pToXml(c, part, doc, F));
    else if (c.tagName === 'TABLE' && c.dataset.t) { const x = tblToXml(c, part, doc, F); if (x) out.push(x); }
    else if (isBlockWrap(c)) { const rec = D.WR.get(c.dataset.w); if (!rec) { out.push(...blocksToXml(c, part, doc, F)); continue; } const w = doc.importNode(rec.el, false); for (const ch of rec.el.children) if (/Pr$/.test(ch.localName)) w.appendChild(doc.importNode(ch, true)); const inner = rec.el.localName === 'sdt' ? wel(doc, 'sdtContent') : w; if (inner !== w) w.appendChild(inner); const bl = blocksToXml(c, part, doc, F); if (!bl.length) bl.push(wel(doc, 'p')); bl.forEach(b => inner.appendChild(b)); out.push(w); }
    else if (isAtom(c)) { const rec = D.AT.get(c.dataset.k); if (rec) { const x = doc.importNode(rec.el, true); fixRelIds(x, D.parts[rec.part], part); out.push(x); } }
  }
  return out;
}
function validate(root) { // fields that lost a piece and bookmarks without their other half are taken out, so Word opens the file cleanly
  const stack = [], drop = new Set();
  for (const el of [...root.getElementsByTagNameNS(NS.w, '*')]) {
    if (el.localName === 'fldChar') { const t = wa(el, 'fldCharType'), r = el.parentNode; if (t === 'begin') stack.push({ parts: [r] }); else if (!stack.length) drop.add(r); else if (t === 'separate') stack[stack.length - 1].parts.push(r); else if (t === 'end') stack.pop(); }
    else if (el.localName === 'instrText' && stack.length) stack[stack.length - 1].parts.push(el.parentNode);
  }
  for (const f of stack) f.parts.forEach(r => drop.add(r));
  for (const [s, e] of [['bookmarkStart', 'bookmarkEnd'], ['commentRangeStart', 'commentRangeEnd']]) {
    const S = [...root.getElementsByTagNameNS(NS.w, s)], E = [...root.getElementsByTagNameNS(NS.w, e)], si = new Set(S.map(x => wa(x, 'id'))), ei = new Set(E.map(x => wa(x, 'id'))), seen = new Set();
    for (const x of S) { const id = wa(x, 'id'); if (!ei.has(id) || seen.has(id)) drop.add(x); seen.add(id); }
    for (const x of E) if (!si.has(wa(x, 'id'))) drop.add(x);
  }
  for (const r of drop) if (r && r.parentNode) r.remove();
}
function uniqueDocPr(docs) { const used = new Set(); let max = 0; const all = docs.flatMap(d => [...d.getElementsByTagNameNS('*', 'docPr')]); for (const e of all) max = Math.max(max, +e.getAttribute('id') || 0); for (const e of all) { const id = +e.getAttribute('id') || 0; if (!id || used.has(id)) e.setAttribute('id', ++max); used.add(+e.getAttribute('id')); } }
function writePart(part, el, kind) {
  if (!part) return null; const doc = part.doc, F = { first: new Map(), last: new Map() };
  for (const p of el.querySelectorAll('p')) if (p.dataset.p) { if (!F.first.has(p.dataset.p)) F.first.set(p.dataset.p, p); F.last.set(p.dataset.p, p); }
  const blocks = blocksToXml(el, part, doc, F); if (!blocks.length || blocks[blocks.length - 1].localName !== 'p') blocks.push(wel(doc, 'p'));
  const root = kind === 'body' ? kid(doc.documentElement, 'body') : doc.documentElement;
  root.replaceChildren(...blocks); if (kind === 'body' && D.sect) root.appendChild(D.sect);
  validate(root); return doc;
}
async function buildDocx() {
  tidy(); const written = [];
  written.push(writePart(D.main, docEl, 'body')); if (D.hdrPart && !hdrEl.hidden) written.push(writePart(D.hdrPart, hdrEl)); if (D.ftrPart && !ftrEl.hidden) written.push(writePart(D.ftrPart, ftrEl));
  uniqueDocPr(written.filter(Boolean));
  for (const part of Object.values(D.parts)) { if ([D.main, D.hdrPart, D.ftrPart, D.stylesPart, D.numPart].includes(part)) { const x = ser(part.doc); if (x !== part.ser0) D.zip.file(part.name, x, { createFolders: false }); } /* (a part that did not change keeps its original bytes) */ if (part.relsDirty || (part.rels && !D.zip.file(part.relsName))) D.zip.file(part.relsName, ser(part.rels), { createFolders: false }); }
  if (D.ctDirty) D.zip.file('[Content_Types].xml', ser(D.ct), { createFolders: false });
  if (D.updateFields) { const st = D.zip.file('word/settings.xml'); if (st) { try { const x = parseXml(await st.async('string')), root = x.documentElement; if (!root.getElementsByTagNameNS(NS.w, 'updateFields').length) { const e = x.createElementNS(NS.w, 'w:updateFields'); e.setAttributeNS(NS.w, 'w:val', 'true'); root.insertBefore(e, root.firstChild); D.zip.file('word/settings.xml', ser(x), { createFolders: false }); } } catch (e) { /* Word then updates the contents when asked */ } } } // (a new table of contents: Word fills in the page numbers on opening)
  const core = D.zip.file('docProps/core.xml'); if (core) { try { const c = parseXml(await core.async('string')), m = c.getElementsByTagNameNS('http://purl.org/dc/terms/', 'modified')[0]; if (m) { m.textContent = new Date().toISOString().replace(/\.\d+Z$/, 'Z'); D.zip.file('docProps/core.xml', ser(c), { createFolders: false }); } } catch (e) { /* left as it is */ } }
  for (const n of Object.keys(D.zip.files)) if (D.zip.files[n].dir) delete D.zip.files[n]; // (only the folder entry — remove() would take its files too)
  return D.zip.generateAsync({ type: 'blob', compression: 'DEFLATE', compressionOptions: { level: 6 }, mimeType: DOCX_MIME });
}
const fileName = () => ($('#dxName').value.trim().replace(/\.docx$/i, '').replace(/[\\/:*?"<>|]+/g, '-') || 'Document') + '.docx';
async function save() { if (!D.zip) return; await busy(async () => { download(await buildDocx(), fileName()); D.dirty = false; clearDraft(); $('#dxState').textContent = 'Saved'; toast('Word document saved to your downloads', 'success'); }, 'Saving…'); }
function printPages() { // the page view as printed pages: each page shows its frame (header, footer, drawings) and the text that lies on it
  if (D.layDirty || !D.layPages) layoutPaged(true);
  const pages = D.layPages || [], out = document.createElement('div'), css = [], names = new Map(); out.className = 'dx-paper dx-print';
  const blocks = [...docEl.children].map(b => ({ b, ...offIn(b, paper), h: b.offsetHeight, w: b.offsetWidth }));
  const ancs = [...docEl.querySelectorAll('.dx-anc')].map(a => ({ a, ...offIn(a, paper), h: a.offsetHeight }));
  const hf = [hdrEl, ftrEl].filter(el => !el.hidden).map(el => ({ el, ...offIn(el, paper), h: el.offsetHeight }));
  pages.forEach((pg, i) => {
    const wPt = pg.g.w / 20, hPt = +(pg.hpx * 0.75).toFixed(2), key = wPt + 'x' + hPt; if (!names.has(key)) { names.set(key, 'xd3p' + names.size); css.push(`@page ${names.get(key)} { size: ${wPt}pt ${hPt}pt; margin: 0; }`); }
    const d = document.createElement('div'); d.className = 'dx-ppage'; d.style.cssText = `width:${pg.wpx}px;height:${pg.hpx}px;page:${names.get(key)}`;
    const fr = frames.children[i] && frames.children[i].cloneNode(true); if (fr) { fr.style.left = fr.style.top = '0'; d.appendChild(fr); }
    const body = document.createElement('div'); body.className = 'dx-doc'; body.style.cssText = `left:${-pg.left}px;top:${-pg.top}px;width:${paper.offsetWidth}px;height:0`; d.appendChild(body);
    const on = x => x.y < pg.top + pg.hpx && x.y + x.h > pg.top;
    for (const x of blocks) if (on(x)) { const c = x.b.cloneNode(true); for (const a of [...c.querySelectorAll('.dx-anc')]) a.remove(); c.style.left = x.x + 'px'; c.style.top = x.y + 'px'; c.style.width = x.w + 'px'; c.style.margin = '0'; body.appendChild(c); }
    for (const x of ancs) if (on(x)) { const c = x.a.cloneNode(true); c.style.left = x.x + 'px'; c.style.top = x.y + 'px'; body.appendChild(c); }
    for (const x of hf) if (on(x)) { const c = hfClone(x.el); c.className = 'dx-hfv'; c.style.cssText = `left:${x.x}px;top:${x.y}px;width:${x.el.offsetWidth}px`; body.appendChild(c); }
    out.appendChild(d);
  });
  const first = pages[0]; css.unshift(`@page { margin: 0; ${first ? `size: ${first.g.w / 20}pt ${+(first.hpx * 0.75).toFixed(2)}pt;` : ''} }`);
  const st = document.createElement('style'); st.id = 'dxPrintCss'; st.textContent = css.join('\n'); document.head.appendChild(st);
  out.hidden = true; scroller.appendChild(out);
  return { first, done: () => { out.remove(); st.remove(); } };
}
function printDoc() {
  const pp = D.paged ? printPages() : null; document.documentElement.classList.add('dxprint');
  const done = () => { document.documentElement.classList.remove('dxprint'); removeEventListener('afterprint', done); if (pp) pp.done(); }; addEventListener('afterprint', done); setTimeout(() => { window.print(); setTimeout(done, 1000); }, 50);
}
async function savePdf() {
  if (X.desk && X.desk.printToPdf) { // the desktop app prints the page itself to PDF — exactly what is on screen
    await busy(async () => {
      const pp = D.paged ? printPages() : null; document.documentElement.classList.add('dxprint'); await new Promise(r => setTimeout(r, 120));
      try {
        const P = D.pg, inch = v => Math.abs(v) / 1440, f = pp && pp.first;
        const o = f ? { width: f.g.w / 1440, height: f.hpx * 0.75 / 72, margins: { top: 0, bottom: 0, left: 0, right: 0 }, css: true } : { width: inch(P.w), height: inch(P.h), margins: { top: inch(P.t), bottom: inch(P.b), left: inch(P.l), right: inch(P.r) } };
        const bytes = await X.desk.printToPdf(o); download(new Blob([bytes], { type: 'application/pdf' }), fileName().replace(/\.docx$/, '.pdf')); toast('PDF saved to your downloads', 'success');
      } finally { document.documentElement.classList.remove('dxprint'); if (pp) pp.done(); }
    }, 'Making the PDF…');
    return;
  }
  const r = await dialog('Save as PDF', `<p>Two ways to make a PDF of this document:</p><label class="inl"><input type="radio" name="m" value="print" checked><span><b>Print to PDF</b> — looks exactly like the page here. Choose “Save as PDF” as the printer.</span></label><label class="inl"><input type="radio" name="m" value="edit"><span><b>Open as PDF in the editor</b> — converted on this device, then edit or save it as a PDF.</span></label>`, { ok: 'Continue' });
  if (!r) return; if (r.m === 'print') return printDoc();
  const blob = await buildDocx(); close(true); await X.openFiles([new File([blob], fileName())], false, { convert: true });
}

/* ───────── opening ───────── */
function resetState() {
  for (const u of D.media.values()) URL.revokeObjectURL(u);
  Object.assign(D, { parts: {}, PP: new Map(), RP: new Map(), RPK: new Map(), AT: new Map(), WR: new Map(), TB: new Map(), media: new Map(), cc: new Map(), undo: [], redo: [], mv: 0, model: null, pending: null, lastSel: null, notes: 0, hdrPart: null, ftrPart: null, stylesPart: null, numPart: null, ct: null, ctDirty: false, dirty: false });
  RFC.clear(); PFC.clear(); RCSS.clear(); stacks.clear(); LINT.cache.clear(); LINT.ignored.clear(); ANC.clear(); HFV.clear(); D.paged = D.exact = false; D.paperW = 0; D.layPages = null;
}
async function loadZip(zip, name) {
  resetState(); D.zip = zip; D.name = name;
  const ct = zip.file('[Content_Types].xml'); D.ct = ct ? parseXml(await ct.async('string')) : null;
  let main = 'word/document.xml'; try { const rr = parseXml(await zip.file('_rels/.rels').async('string')); for (const r of rr.documentElement.children) if (/\/officeDocument$/.test(r.getAttribute('Type'))) main = r.getAttribute('Target').replace(/^\//, ''); } catch (e) { /* the usual place */ }
  const P = await loadPart(main); if (!P || !kid(P.doc.documentElement, 'body')) throw new Error('This is not a Word document (no document body)'); D.main = P;
  for (const r of Object.values(relsOf(P))) {
    const path = resolvePath(P, r.target);
    if (/\/styles$/.test(r.type)) D.stylesPart = await loadPart(path); else if (/\/numbering$/.test(r.type)) D.numPart = await loadPart(path);
    else if (/\/theme$/.test(r.type)) { try { const t = parseXml(await zip.file(path).async('string')), lat = n => { const f = t.getElementsByTagNameNS('*', n)[0], l = f && [...f.children].find(c => c.localName === 'latin'); return l && l.getAttribute('typeface'); }; D.theme = { major: lat('majorFont') || 'Calibri Light', minor: lat('minorFont') || 'Calibri' }; } catch (e) { /* default theme */ } }
  }
  if (!D.stylesPart) { // a document without styles gets an empty styles part so styles can be added
    const name = 'word/styles.xml'; D.stylesPart = D.parts[name] = { name, doc: parseXml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="${NS.w}"><w:docDefaults><w:rPrDefault><w:rPr/></w:rPrDefault><w:pPrDefault><w:pPr/></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style></w:styles>`), rels: null, relsName: 'word/_rels/styles.xml.rels', dir: 'word/', relsDirty: false };
    addRel(P, 'styles', 'styles.xml'); ensureType({ part: name, type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml' });
  }
  D.ST = loadStyles(D.stylesPart.doc); D.NUM = loadNumbering(D.numPart && D.numPart.doc); changedModel(true);
  const body = kid(P.doc.documentElement, 'body'), sects = [...body.children].filter(c => c.localName === 'sectPr');
  D.sect = sects[sects.length - 1] || null; if (!D.sect) { D.sect = parseInto(P.doc, '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>'); }
  const allSect = [...P.doc.getElementsByTagNameNS(NS.w, 'sectPr')].reverse(), rels = relsOf(P); D.geo = allSect[allSect.length - 1] || D.sect; // (the first section sets the page shown)
  const hf = async ref => { for (const s of allSect) { const r = kids(s, ref).find(x => (wa(x, 'type') || 'default') === 'default'); if (r) { const rel = rels[r.getAttributeNS(NS.r, 'id')]; if (rel) return loadPart(resolvePath(P, rel.target)); } } return null; };
  D.hdrPart = await hf('headerReference'); D.ftrPart = await hf('footerReference');
  await loadMedia(P); if (D.hdrPart) await loadMedia(D.hdrPart); if (D.ftrPart) await loadMedia(D.ftrPart);
  D.defTab = 36; try { const st = zip.file('word/settings.xml'), m = st && /<w:defaultTabStop w:val="(\d+)"/.exec(await st.async('string')); if (m && +m[1] > 0) D.defTab = +m[1] / 20; } catch (e) { /* 36pt then */ }
  await loadEmbeddedFonts(); const xl = exactLook(body); D.exact = xl >= 1; D.paged = D.exact || (xl > 0 && await pageAnchored(P)); // (a normal file with a page-anchored logo keeps the normal view)
  if (D.paged) D.paperW = Math.max(...allSect.map(s => +wa(kid(s, 'pgSz'), 'w') || 12240), +wa(kid(D.sect, 'pgSz'), 'w') || 0);
  if (D.paged) for (const r of Object.values(rels)) if (/\/(header|footer)$/.test(r.type)) { const part = D.parts[resolvePath(P, r.target)] || await loadPart(resolvePath(P, r.target)); if (part) await loadMedia(part); } // (every page shows its own header)
  box.classList.toggle('dx-paged', !!D.paged); pagedCss();
  docEl.replaceChildren(renderRoot(body, P));
  hdrEl.replaceChildren(D.hdrPart ? renderRoot(D.hdrPart.doc.documentElement, D.hdrPart) : ''); hdrEl.hidden = !D.hdrPart;
  ftrEl.replaceChildren(D.ftrPart ? renderRoot(D.ftrPart.doc.documentElement, D.ftrPart) : ''); ftrEl.hidden = !D.ftrPart;
  frames.replaceChildren(); frames._sig = ''; D.layPages = null; paper.style.minHeight = ''; hdrEl.removeAttribute('style'); ftrEl.removeAttribute('style');
  if (D.paged) hdrEl.hidden = ftrEl.hidden = true; // (in the page view headers are drawn on each page)
  $('#dxName').value = name; $('#dxState').textContent = ''; document.title = name + ' — XD3 Word editor';
  pageGeometry(); tidy(roots().filter(r => !r.hidden)); fitZoom(); dirty(); D.dirty = false; $('#dxState').textContent = '';
  scroller.scrollTop = 0; afterEdit(true);
  const fl = $('[data-c=flow]', box); if (fl) { if (!fl._sep) fl._sep = fl.previousElementSibling; if (D.exact) fl.parentNode.prepend(fl, fl._sep); else fl.parentNode.append(fl._sep, fl); } // (first on the Layout tab for such a file: in view on a phone)
  if (D.exact) setTimeout(() => toast('This file keeps the exact look of its PDF. Layout ▸ Make flowing (the first button) turns it into normal text that reflows as you edit.', 'info'), 400);
  const first = docEl.querySelector('p'); if (first) { const P2 = allParas(); setSel({ a: [P2.indexOf(first), 0], b: [P2.indexOf(first), 0] }); }
}
async function openFile(file) {
  if (!file) return; if (D.open && D.dirty && !confirm('Open another document? Unsaved changes to this one will be lost.')) return;
  show();
  await busy(async () => {
    const JSZip = await X.loadLib('jszip'); let zip;
    try { zip = await JSZip.loadAsync(await file.arrayBuffer()); } catch (e) { throw new Error('This is not a Word (.docx) file, or it is damaged'); }
    await loadZip(zip, file.name.replace(/\.docx$/i, ''));
    if (X.noteRecent) try { X.noteRecent([file]); } catch (e) { /* */ }
  }, 'Opening the Word document…');
  if (!D.zip) close(true);
}
async function blankZip() {
  const JSZip = await X.loadLib('jszip'), z0 = new JSZip(), z = { file: (n, d) => z0.file(n, d, { createFolders: false }) }, H = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>', us = /^en-(US|CA)|^es-(MX|US)/.test(navigator.language || '');
  const [w, h] = us ? [12240, 15840] : [11906, 16838], now = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
  z.file('[Content_Types].xml', `${H}<Types xmlns="${NS.ct}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Default Extension="jpeg" ContentType="image/jpeg"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>`);
  z.file('_rels/.rels', `${H}<Relationships xmlns="${NS.rel}"><Relationship Id="rId1" Type="${RT}officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="${RT}extended-properties" Target="docProps/app.xml"/></Relationships>`);
  z.file('word/_rels/document.xml.rels', `${H}<Relationships xmlns="${NS.rel}"><Relationship Id="rId1" Type="${RT}styles" Target="styles.xml"/><Relationship Id="rId2" Type="${RT}settings" Target="settings.xml"/></Relationships>`);
  z.file('word/document.xml', `${H}<w:document xmlns:w="${NS.w}" xmlns:r="${NS.r}" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><w:body><w:p/><w:sectPr><w:pgSz w:w="${w}" w:h="${h}"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/><w:cols w:space="708"/></w:sectPr></w:body></w:document>`);
  const st = (id, name, inner, x = '') => `<w:style w:type="paragraph" w:styleId="${id}"${x}><w:name w:val="${name}"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="9"/><w:qFormat/>${inner}</w:style>`;
  z.file('word/styles.xml', `${H}<w:styles xmlns:w="${NS.w}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="${esc(navigator.language || 'en-US')}"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>`
    + '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style><w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/><w:uiPriority w:val="1"/><w:semiHidden/><w:unhideWhenUsed/></w:style><w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>'
    + ['title', 'heading 1', 'heading 2', 'heading 3'].map(k => st(BUILTIN[k][0], k, BUILTIN[k][1])).join('') + `<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:basedOn w:val="DefaultParagraphFont"/><w:uiPriority w:val="99"/><w:unhideWhenUsed/>${BUILTIN.hyperlink[1]}</w:style></w:styles>`);
  z.file('word/settings.xml', `${H}<w:settings xmlns:w="${NS.w}"><w:zoom w:percent="100"/><w:defaultTabStop w:val="720"/><w:characterSpacingControl w:val="doNotCompress"/><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>`);
  z.file('docProps/core.xml', `${H}<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title></dc:title><dc:creator></dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified></cp:coreProperties>`);
  z.file('docProps/app.xml', `${H}<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>XD3 PDF Editor</Application></Properties>`);
  return z0;
}

/* ───────── templates and letterheads ───────── */
const TX = t => String(t).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
const tP = (text, o = {}) => { // one paragraph of a template: style, alignment, spacing, bold / size / colour, a bottom line
  const ppr = (o.style ? `<w:pStyle w:val="${o.style}"/>` : '') + (o.line ? '<w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="A6A6A6"/></w:pBdr>' : '') + (o.tab ? `<w:tabs><w:tab w:val="right" w:pos="${o.tab}"/></w:tabs>` : '') + (o.after != null || o.before != null ? `<w:spacing${o.before != null ? ` w:before="${o.before}"` : ''}${o.after != null ? ` w:after="${o.after}"` : ''}/>` : '') + (o.jc ? `<w:jc w:val="${o.jc}"/>` : '');
  const rpr = (o.b ? '<w:b/>' : '') + (o.i ? '<w:i/>' : '') + (o.color ? `<w:color w:val="${o.color}"/>` : '') + (o.sz ? `<w:sz w:val="${o.sz * 2}"/><w:szCs w:val="${o.sz * 2}"/>` : '');
  const runs = String(text).split('\t').map((t, i) => (i ? '<w:r><w:tab/></w:r>' : '') + (t ? `<w:r>${rpr ? `<w:rPr>${rpr}</w:rPr>` : ''}<w:t xml:space="preserve">${TX(t)}</w:t></w:r>` : '')).join('');
  return `<w:p>${ppr ? `<w:pPr>${ppr}</w:pPr>` : ''}${runs}</w:p>`;
};
const tTable = (rows, widths, o = {}) => { const bd = o.borders === false ? 'nil' : 'single', hdr = o.header !== false;
  return `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders>${['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map(k => `<w:${k} w:val="${bd}" w:sz="4" w:space="0" w:color="BFBFBF"/>`).join('')}</w:tblBorders><w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/></w:tblPr><w:tblGrid>${widths.map(w => `<w:gridCol w:w="${w}"/>`).join('')}</w:tblGrid>`
    + rows.map((r, ri) => `<w:tr>${r.map((c, ci) => `<w:tc><w:tcPr><w:tcW w:w="${widths[ci]}" w:type="dxa"/>${hdr && ri === 0 && o.borders !== false ? '<w:shd w:val="clear" w:color="auto" w:fill="DEEAF6"/>' : ''}</w:tcPr>${tP(c, { b: hdr && ri === 0 && o.borders !== false, after: 0, jc: ci > 0 && o.right && o.right.includes(ci) ? 'right' : undefined })}</w:tc>`).join('')}</w:tr>`).join('') + '</w:tbl>'; };
const TODAY = () => new Date().toLocaleDateString(undefined, { day: 'numeric', month: 'long', year: 'numeric' });
const TEMPLATES = {
  blank: { name: 'Blank document', desc: 'An empty page', body: () => tP('') },
  letter: { name: 'Letter', desc: 'Date, address, greeting and signature', body: () => [tP(TODAY(), { jc: 'right', after: 240 }), tP('[Recipient name]', { after: 0 }), tP('[Company]', { after: 0 }), tP('[Street address]', { after: 0 }), tP('[City, postcode]', { after: 360 }),
    tP('Subject: [What this letter is about]', { b: true, after: 240 }), tP('Dear [Name],'), tP('[Write the first paragraph here. Say why you are writing.]'), tP('[Add the details in a second paragraph.]'), tP('[Close with what you would like to happen next.]', { after: 360 }),
    tP('Yours sincerely,', { after: 720 }), tP('[Your name]', { b: true, after: 0 }), tP('[Your title]', { after: 0 })].join('') },
  resume: { name: 'Résumé / CV', desc: 'Name, profile, experience, education, skills', body: () => [tP('[Your Name]', { style: 'Title' }), tP('[Email]  ·  [Phone]  ·  [City]  ·  [LinkedIn or website]', { color: '595959', after: 240, line: true }),
    tP('Profile', { style: 'Heading1' }), tP('[Two or three sentences about who you are and what you are good at.]'),
    tP('Experience', { style: 'Heading1' }), tP('[Job title] — [Company]\t[2022 – present]', { b: true, after: 40, tab: 9026 }), tP('•  [What you did and what came of it — with numbers where you can]', { after: 40 }), tP('•  [Another achievement]'),
    tP('[Job title] — [Company]\t[2019 – 2022]', { b: true, after: 40, tab: 9026 }), tP('•  [What you did and what came of it]'),
    tP('Education', { style: 'Heading1' }), tP('[Degree] — [University]\t[Year]', { b: true, after: 40, tab: 9026 }), tP('[Grade, focus or award]'),
    tP('Skills', { style: 'Heading1' }), tP('[Skill]  ·  [Skill]  ·  [Skill]  ·  [Language]')].join('') },
  invoice: { name: 'Invoice', desc: 'From, bill to, items, total', body: () => [tP('INVOICE', { style: 'Title' }), tP(`Invoice no. [001]\tDate: ${TODAY()}`, { tab: 9026, after: 240, line: true }),
    tTable([['From', 'Bill to'], ['[Your company]', '[Client name]'], ['[Address]', '[Address]'], ['[Email / phone]', '[Email / phone]']], [4513, 4513], { borders: false, header: false }), tP('', { after: 120 }),
    tTable([['Description', 'Qty', 'Price', 'Amount'], ['[Item or service]', '1', '0.00', '0.00'], ['[Item or service]', '1', '0.00', '0.00'], ['', '', 'Total', '0.00']], [4826, 1100, 1550, 1550], { right: [1, 2, 3] }),
    tP('', { after: 120 }), tP('Payment terms: [e.g. within 14 days]', { after: 40 }), tP('Bank details: [account name, number, IFSC / IBAN]', { after: 240 }), tP('Thank you for your business!', { i: true, color: '595959' })].join('') },
  report: { name: 'Report', desc: 'Title page, headings and sections', body: () => [tP('[Report Title]', { style: 'Title' }), tP(`[Subtitle]  ·  ${TODAY()}`, { color: '595959', after: 360 }),
    tP('Summary', { style: 'Heading1' }), tP('[The main points in a few sentences.]'), tP('Background', { style: 'Heading1' }), tP('[Why this report was written.]'),
    tP('Findings', { style: 'Heading1' }), tP('[First finding]', { style: 'Heading2' }), tP('[Details]'), tP('[Second finding]', { style: 'Heading2' }), tP('[Details]'),
    tP('Recommendations', { style: 'Heading1' }), tP('[What should happen next.]')].join('') },
  meeting: { name: 'Meeting notes', desc: 'Attendees, agenda, notes, action items', body: () => [tP('[Meeting name]', { style: 'Title' }), tP(`${TODAY()}  ·  [Time]  ·  [Place / call link]`, { color: '595959', after: 240, line: true }),
    tP('Attendees', { style: 'Heading2' }), tP('[Names]'), tP('Agenda', { style: 'Heading2' }), tP('1.  [Topic]', { after: 40 }), tP('2.  [Topic]'),
    tP('Notes', { style: 'Heading2' }), tP('[What was discussed and decided]'), tP('Action items', { style: 'Heading2' }), tP('☐ [Task] — [Owner] — [Due date]', { after: 40 }), tP('☐ [Task] — [Owner] — [Due date]')].join('') },
  memo: { name: 'Memo', desc: 'To, from, date, subject', body: () => [tP('MEMO', { style: 'Title' }), tP('To:\t[Name]', { tab: 1440, after: 40 }), tP('From:\t[Name]', { tab: 1440, after: 40 }), tP(`Date:\t${TODAY()}`, { tab: 1440, after: 40 }), tP('Subject:\t[Subject]', { tab: 1440, b: true, line: true, after: 240 }),
    tP('[Write the memo here — keep it short and say clearly what you need.]')].join('') },
};
async function newFromTemplate() {
  if (D.open && D.dirty && !confirm('Start a new document? Unsaved changes will be lost.')) return;
  const lhs = X.letterheads ? X.letterheads.list() : [];
  const html = `<div class="tpl-grid">${Object.entries(TEMPLATES).map(([k, t], i) => `<label class="tpl"><input type="radio" name="t" value="${k}"${i === 0 ? ' checked' : ''}><b>${esc(t.name)}</b><span>${esc(t.desc)}</span></label>`).join('')}</div>
    <label>Letterhead<select name="lh"><option value="">None</option>${lhs.map((l, i) => `<option value="${i}">${esc(l.label)}</option>`).join('')}<option value="new">Design a new letterhead…</option></select></label>`;
  const r = await dialog('New document', html, { ok: 'Create' }); if (!r) return;
  const t = TEMPLATES[r.t] || TEMPLATES.blank, z = await blankZip(), xml = await z.file('word/document.xml').async('string');
  z.file('word/document.xml', xml.replace('<w:body><w:p/>', '<w:body>' + t.body()), { createFolders: false });
  show(); await loadZip(z, r.t === 'blank' ? 'Document' : t.name); D.dirty = false;
  let lh = r.lh === 'new' && X.letterheads ? await X.letterheads.design() : r.lh !== '' && r.lh != null ? lhs[+r.lh] : null;
  if (lh) await applyLetterhead(lh);
  const first = allParas().find(p => p.closest('#dxDoc')); if (first) setSel({ a: [allParas().indexOf(first), 0], b: [allParas().indexOf(first), 0] }, true);
}
async function applyLetterhead(lh) { // the letterhead becomes the header (and footer) of every page
  if (!D.zip || !lh) return;
  const toFile = async (url, n) => new File([await (await fetch(url)).blob()], n, { type: 'image/png' });
  for (const [kind, url] of [['header', lh.header], ['footer', lh.footer]]) {
    if (!url) continue;
    await headerFooter(kind); const el = kind === 'header' ? hdrEl : ftrEl;
    snapshot(); el.replaceChildren(emptyP()); tidy([el]); const p = el.querySelector('p'), P = allParas(); setSel({ a: [P.indexOf(p), 0], b: [P.indexOf(p), 0] });
    await insertPicture(await toFile(url, kind === 'header' ? 'letterhead.png' : 'letterhead-footer.png'));
  }
  const first = allParas().find(q => q.closest('#dxDoc')); if (first) setSel({ a: [allParas().indexOf(first), 0], b: [allParas().indexOf(first), 0] });
  toast(`Letterhead “${lh.label}” is now the ${lh.footer ? 'header and footer' : 'header'} of every page`, 'success');
}
async function newDocument() { if (D.open && D.dirty && !confirm('Start a new document? Unsaved changes will be lost.')) return; show(); await loadZip(await blankZip(), 'Document'); }
function show() { document.querySelectorAll('.menu.open').forEach(m => m.classList.remove('open')); if (X.hideCtx) X.hideCtx(); if (!D.open) { D.open = true; if (innerWidth < 860 && store.get('xd3.dxside') == null) D.side = false; // (on a phone the page comes first)
  box.hidden = false; document.body.classList.add('modal', 'dxopen'); box.classList.toggle('noside', !D.side); setTab(D.tab); $('#dxDialect').value = X.writing ? X.writing.W.dialect : 'American'; X.writing && X.writing.engine().catch(() => { LINT.ok = false; }); } }
/* ───────── autosave: a copy of the document while you work, offered back if the window closed before you saved ───────── */
const DRAFT = 'dxdraft';
const clearDraft = () => X.DB && X.DB.put(DRAFT, null).catch(() => {});
function autosave(now) {
  if ((!D.open && !now) || !D.dirty || !D.zip) return;
  const go = async () => {
    if ((!D.open && !now) || !D.dirty) return; const sel = getSel();
    try { const blob = await buildDocx(); await X.DB.put(DRAFT, { name: fileName(), time: Date.now(), blob }); $('#dxState').textContent = 'Edited · draft kept'; }
    catch (e) { /* the next change tries again */ }
    if (sel && inEditor()) setSel(sel);
  };
  if (now) return go(); window.requestIdleCallback ? requestIdleCallback(go, { timeout: 4000 }) : setTimeout(go, 300);
}
async function offerDraft() { // → true when a draft was opened
  let d = null; try { d = X.DB && await X.DB.get(DRAFT); } catch (e) { return false; }
  if (!d || !d.blob || (D.open && D.dirty)) return false;
  const when = new Date(d.time).toLocaleString(undefined, { day: 'numeric', month: 'short', hour: '2-digit', minute: '2-digit' });
  const r = await dialog('Recover your document?', `<p><b>${esc(d.name)}</b> was not saved before the editor closed (${esc(when)}). A copy was kept on this device.</p><label class="inl"><input type="radio" name="a" value="open" checked>Open the recovered copy</label><label class="inl"><input type="radio" name="a" value="drop">Throw it away</label>`, { ok: 'OK' });
  if (!r) return false; if (r.a !== 'open') { clearDraft(); return false; }
  await openFile(new File([d.blob], d.name, { type: DOCX_MIME })); D.dirty = true; $('#dxState').textContent = 'Recovered — save it to keep it'; return true;
}
async function openEditor() { show(); if (!D.zip && !(await offerDraft())) await newDocument(); }
function close(force) {
  if (!force && D.dirty && !confirm('Close the Word editor? Unsaved changes will be lost.')) return;
  if (!force && D.dirty) clearDraft(); clearTimeout(D.asT);
  D.open = false; box.hidden = true; pop.hidden = true; document.body.classList.remove('modal', 'dxopen'); X.updateUI();
  if (window.CSS && CSS.highlights) for (const k of ['dx-spell', 'dx-grammar', 'dx-improve', 'dx-style', 'dx-find', 'dx-findcur']) CSS.highlights.delete(k);
}
function leave() { // to the home screen: the document stays open in the background (and a recovery copy is kept), nothing is asked or lost
  if (box.hidden) return; if (D.dirty) autosave(true); close(true);
}
window.addEventListener('beforeunload', e => { if (D.zip && D.dirty) { e.preventDefault(); e.returnValue = ''; } });
window.addEventListener('resize', () => { if (!box.hidden) { D.layDirty = true; clearTimeout(paginate.s); paginate.s = setTimeout(() => { paginate(); placeGrip(); }, 120); } });
$('#dxName').addEventListener('keydown', e => { e.stopPropagation(); if (e.key === 'Enter') e.target.blur(); });
$('#dxName').addEventListener('change', () => { document.title = $('#dxName').value + ' — XD3 Word editor'; dirty(); });

X.docx = { D, leave, name: () => fileName(), openFile, newDocument, newFromTemplate, applyLetterhead, autosave, offerDraft, open: openEditor, close, isOpen: () => D.open, buildDocx, checks, autoFormat, applyFixes, setStyle, listToggle, setRunProp, getSel, setSel, allParas, ptext, undo, redo, tidy, loadZip, blankZip, replaceText, lintAll, insertTable, tableOp, align, makeFlowing, layoutNow };
})();
