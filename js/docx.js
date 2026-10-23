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
const R_EL = { rStyle: ['rStyle'], font: ['rFonts'], b: ['b', 'bCs'], i: ['i', 'iCs'], caps: ['caps'], smallCaps: ['smallCaps'], strike: ['strike'], dstrike: ['dstrike'], vanish: ['vanish'], color: ['color'], sz: ['sz', 'szCs'], hl: ['highlight'], u: ['u'], shd: ['shd'], va: ['vertAlign'] };
const P_EL = { style: ['pStyle'], keepNext: ['keepNext'], pbb: ['pageBreakBefore'], num: ['numPr'], sp: ['spacing'], ind: ['ind'], ctx: ['contextualSpacing'], jc: ['jc'] };
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
  return o;
}
function writePProp(pPr, k, v) {
  for (const n of P_EL[k]) for (let e; (e = kid(pPr, n));) e.remove();
  if (v == null || ((k === 'sp' || k === 'ind') && !Object.keys(v).length)) return; const d = pPr.ownerDocument; let e;
  if (k === 'style') e = wel(d, 'pStyle', { val: v });
  else if (k === 'num') { e = wel(d, 'numPr'); e.appendChild(wel(d, 'ilvl', { val: v.lvl || 0 })); e.appendChild(wel(d, 'numId', { val: v.id })); }
  else if (k === 'sp') e = wel(d, 'spacing', v);
  else if (k === 'ind') e = wel(d, 'ind', { left: v.left, right: v.right, firstLine: v.hang != null ? null : v.first, hanging: v.hang });
  else if (k === 'jc') e = wel(d, 'jc', { val: v });
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
const plainTable = t => [...t.children].every(c => /^(tblPr|tblGrid|tr|bookmarkStart|bookmarkEnd|commentRangeStart|commentRangeEnd|proofErr)$/.test(c.localName))
  && kids(t, 'tr').every(tr => [...tr.children].every(c => /^(trPr|tblPrEx|tc|bookmarkStart|bookmarkEnd|commentRangeStart|commentRangeEnd|proofErr)$/.test(c.localName)));
function renderTbl(t, part, pend) {
  if (!plainTable(t)) return atomEl(t, part, true);
  const id = nid(), tbl = document.createElement('table'), tblPr = kid(t, 'tblPr'), ts = wa(kid(tblPr, 'tblStyle'), 'val');
  D.TB.set(id, { el: t, part: part.name }); tbl.dataset.t = id; tbl.className = 'dx-t';
  let bord = hasBorders(kid(tblPr, 'tblBorders')); if (bord === null) { const st = ts && D.ST.byId[ts]; bord = !!(st && st.borders); } if (bord) tbl.classList.add('dx-tb');
  const grid = kids(kid(t, 'tblGrid'), 'gridCol').map(g => (+wa(g, 'w') || 0) / 20), cg = document.createElement('colgroup');
  for (const w of grid) { const c = document.createElement('col'); if (w) c.style.width = w + 'pt'; cg.appendChild(c); } tbl.appendChild(cg);
  const jc = wa(kid(tblPr, 'jc'), 'val'); if (jc === 'center') tbl.style.margin = '0 auto';
  for (const c of t.children) if (MARKS.has(c.localName)) pend.push(atomEl(c, part));
  const body = document.createElement('tbody'); tbl.appendChild(body);
  for (const tr of kids(t, 'tr')) {
    const row = document.createElement('tr'), rid = nid(); D.TB.set(rid, { el: tr, part: part.name }); row.dataset.tr = rid;
    for (const c of tr.children) if (MARKS.has(c.localName)) pend.push(atomEl(c, part));
    for (const tc of kids(tr, 'tc')) {
      const td = document.createElement('td'), cid = nid(), pr = kid(tc, 'tcPr'), vm = kid(pr, 'vMerge'), sh = wa(kid(pr, 'shd'), 'fill');
      D.TB.set(cid, { el: tc, part: part.name }); td.dataset.tc = cid; td.colSpan = +wa(kid(pr, 'gridSpan'), 'val') || 1;
      if (vm && wa(vm, 'val') !== 'restart') td.classList.add('dx-vm'); if (sh && /^[0-9a-f]{6}$/i.test(sh)) td.style.background = '#' + sh;
      const va = wa(kid(pr, 'vAlign'), 'val'); if (va === 'center' || va === 'bottom') td.style.verticalAlign = va === 'center' ? 'middle' : 'bottom';
      renderBlocks(tc, part, td, pend); if (!td.querySelector('p')) td.appendChild(emptyP());
      row.appendChild(td);
    }
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
box.innerHTML = `<header class="dx-top">
    <button class="btn ico" data-c="close" title="Close the Word editor"><i data-icon="left"></i></button>
    <span class="dx-logo"><i data-icon="word"></i></span>
    <input id="dxName" class="dx-name" spellcheck="false" title="File name">
    <span id="dxState" class="dx-state"></span>
    <div class="spacer"></div>
    <button class="btn ico" data-c="new" title="New Word document"><i data-icon="plus"></i></button>
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
  <div class="dx-ribbon" id="dxRib">
    <select id="dxStyle" title="Paragraph style"></select>
    <input id="dxFont" list="dxFonts" spellcheck="false" title="Font"><datalist id="dxFonts"></datalist>
    <input id="dxSize" list="dxSizes" spellcheck="false" title="Font size" class="dx-size"><datalist id="dxSizes">${[8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72].map(s => `<option value="${s}">`).join('')}</datalist>
    <button class="dxb dxt" data-c="grow" title="Bigger text (Ctrl+])">A<sup>+</sup></button><button class="dxb dxt" data-c="shrink" title="Smaller text (Ctrl+[)">A<sup>−</sup></button>
    <span class="sep"></span>
    ${tb('b', 'bold', 'Bold (Ctrl+B)')}${tb('i', 'italic', 'Italic (Ctrl+I)')}${tb('u', 'underline', 'Underline (Ctrl+U)')}${tb('strike', 'strike', 'Strikethrough')}${tb('sup', 'sup', 'Superscript')}${tb('sub', 'sub', 'Subscript')}
    <span class="dxsplit"><button class="dxb" data-c="color" title="Text colour"><i data-icon="fcolor"></i><b class="dxbar" id="dxColBar"></b></button><button class="dxb dxdd" data-c="colormenu" title="Choose a text colour">▾</button></span>
    <span class="dxsplit"><button class="dxb" data-c="hl" title="Highlight"><i data-icon="hlpen"></i><b class="dxbar" id="dxHlBar"></b></button><button class="dxb dxdd" data-c="hlmenu" title="Choose a highlight colour">▾</button></span>
    ${tb('clear', 'clearfmt', 'Clear formatting (Ctrl+Space)')}
    <span class="sep"></span>
    ${tb('al', 'alignleft', 'Align left (Ctrl+L)')}${tb('ac', 'aligncenter', 'Centre (Ctrl+E)')}${tb('ar', 'alignright', 'Align right (Ctrl+R)')}${tb('aj', 'alignjust', 'Justify (Ctrl+J)')}
    <span class="sep"></span>
    ${tb('bullets', 'listbul', 'Bullets (Ctrl+Shift+L)')}${tb('numbers', 'listnum', 'Numbering')}${tb('outdent', 'outdent', 'Decrease indent (Shift+Tab in a list)')}${tb('indent', 'indent', 'Increase indent (Tab in a list)')}${tb('spacing', 'linesp', 'Line and paragraph spacing')}
    <span class="sep"></span>
    ${tb('table', 'table', 'Insert a table')}${tb('picture', 'image', 'Insert a picture')}${tb('link', 'link', 'Insert a link (Ctrl+K)')}${tb('pbreak', 'pbreak', 'Page break (Ctrl+Enter)')}${tb('symbol', 'omega', 'Symbols')}${tb('hf', 'hf', 'Header, footer and page numbers')}
    <span class="sep"></span>
    ${tb('find', 'search', 'Find and replace (Ctrl+F / Ctrl+H)')}${tb('page', 'pagesetup', 'Page setup — size, orientation, margins')}
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
      <nav class="wr-tabs dx-tabs"><button class="dx-sideclose" data-c="panel" title="Close the panel">✕</button><button data-t="proof">Spelling <b id="dxCnt">0</b></button><button data-t="format">Format <b id="dxFcnt">·</b></button><button data-t="human">Humanize</button><button data-t="outline">Outline</button><button data-t="stats">Stats</button></nav>
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
  if (x.shd) s += `;background:#${x.shd}`; if (x.bdr) s += ';' + x.bdr;
  return s;
}
function applyP(p, prev, next, lastOfId) { // paragraph and run looks from the styles (p._pe of the neighbours must be fresh)
  const e = p._pe, rec = p.dataset.p && D.PP.get(p.dataset.p), same = q => q && q.parentNode === p.parentNode && q._pe && q._pe.style === e.style;
  const x = rec ? rec.x : {}, css = cssP(e, effR(rec ? rec.mark : {}, e.style), e.ctx && same(prev), e.ctx && same(next), x);
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
function dirty() { D.dirty = true; $('#dxState').textContent = 'Edited'; $('[data-c=undo]', box).disabled = !D.undo.length; $('[data-c=redo]', box).disabled = !D.redo.length; }
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
    el.replaceChildren(renderRoot(doc.documentElement, part)); el.hidden = false; tidy([el]);
  }
  if (kind === 'pagenum') {
    const p = [...el.querySelectorAll('p')].pop(); setSel({ a: [allParas().indexOf(p), 0], b: [allParas().indexOf(p), 0] });
    cmd(() => { const q = emptyP(); q.replaceChildren(); setPf(q, { jc: 'center' }); el.appendChild(q); const part = D.ftrPart, x = part.doc.createElementNS(NS.w, 'w:fldSimple'); x.setAttributeNS(NS.w, 'w:instr', ' PAGE   \\* MERGEFORMAT '); const w = document.createElement('span'), id = nid(); D.WR.set(id, { el: x, part: part.name }); w.dataset.w = id; w.className = 'dx-w dx-fldSimple'; const sp = document.createElement('span'); setRf(sp, {}); sp.textContent = '1'; w.appendChild(sp); q.appendChild(w); if (!units(el.querySelector('p')).some(u => u.t !== 'b') && el.querySelectorAll('p').length > 1) el.querySelector('p').remove(); }, { keepSel: true });
    toast('Page numbers added to the footer — Word shows the right number on every page', 'success'); return;
  }
  el.hidden = false; const p = el.querySelector('p'), P = allParas(); setSel({ a: [P.indexOf(p), 0], b: [P.indexOf(p), 0] }, true);
}

/* ───────── page geometry ───────── */
const SECT_ORDER = ['headerReference', 'footerReference', 'footnotePr', 'endnotePr', 'type', 'pgSz', 'pgMar', 'paperSrc', 'pgBorders', 'lnNumType', 'pgNumType', 'cols', 'formProt', 'vAlign', 'noEndnote', 'titlePg', 'textDirection', 'bidi', 'rtlGutter', 'docGrid', 'printerSettings', 'sectPrChange'];
const tw2pt = v => (+v || 0) / 20;
function pageGeometry() { // the page as the first section has it (a document can change size or turn landscape later on)
  const s = D.geo || D.sect, sz = kid(s, 'pgSz'), m = kid(s, 'pgMar'), g = k => wa(m, k) == null ? 1440 : +wa(m, k);
  D.pg = { w: +wa(sz, 'w') || 12240, h: +wa(sz, 'h') || 15840, t: g('top'), b: g('bottom'), l: g('left'), r: g('right'), hd: +wa(m, 'header') || 720, ft: +wa(m, 'footer') || 720 };
  const P = D.pg; paper.style.width = tw2pt(P.w) + 'pt'; paper.style.minHeight = tw2pt(P.h) + 'pt';
  paper.style.padding = `${tw2pt(Math.max(P.hd, 0))}pt ${tw2pt(P.r)}pt ${tw2pt(P.ft)}pt ${tw2pt(P.l)}pt`;
  docEl.style.marginTop = Math.max(0, tw2pt(Math.abs(P.t) - P.hd)) + 'pt'; docEl.style.marginBottom = Math.max(0, tw2pt(Math.abs(P.b) - P.ft)) + 'pt';
  let st = $('#dxPageCss'); if (!st) { st = document.createElement('style'); st.id = 'dxPageCss'; document.head.appendChild(st); }
  st.textContent = `@page { size: ${tw2pt(P.w)}pt ${tw2pt(P.h)}pt; margin: ${tw2pt(Math.abs(P.t))}pt ${tw2pt(P.r)}pt ${tw2pt(Math.abs(P.b))}pt ${tw2pt(P.l)}pt; }`;
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
  const lay = $('#dxPages'); if (!D.pg || box.hidden) return; const k = 96 / 72, H = tw2pt(D.pg.h - Math.abs(D.pg.t) - Math.abs(D.pg.b)) * k, top0 = docEl.offsetTop;
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

/* ───────── after every change ───────── */
let editTimer = 0;
function afterEdit(now) {
  clearTimeout(editTimer);
  editTimer = setTimeout(() => { paginate(); countWords(); lintSoon(); if (D.tab === 'format' || D.tab === 'outline' || D.tab === 'stats') renderSide(); if (findState.q) runFind(false); }, now ? 0 : 350);
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
  if (pe) { const jc = { both: 'aj', distribute: 'aj', center: 'ac', right: 'ar', end: 'ar' }[pe.jc] || 'al'; for (const c of ['al', 'ac', 'ar', 'aj']) on(c, c === jc); if (document.activeElement !== $('#dxStyle')) $('#dxStyle').value = pe.style; const L = pe.num && pe.num.id !== '0' && lvlOf(pe.num); on('bullets', L && L.fmt === 'bullet'); on('numbers', L && L.fmt !== 'bullet'); }
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
    case 'find': return findBar(true); case 'page': return pageSetup();
    case 'panel': D.side = !D.side; store.set('xd3.dxside', D.side ? 'on' : 'off'); box.classList.toggle('noside', !D.side); renderSide(); return setTimeout(() => { fitZoom(); paginate(); }, 30);
    case 'close': return close(); case 'new': return newDocument(); case 'open': return $('#dxFile').click();
    case 'undo': return undo(); case 'redo': return redo();
    case 'save': return save(); case 'pdf': return savePdf(); case 'print': return printDoc();
    case 'autofmt': return autoFormat(); case 'fixfmt': D.side = true; box.classList.remove('noside'); return setTab('format');
    case 'fprev': return stepFind(-1); case 'fnext': return stepFind(1); case 'rone': return replaceOne(); case 'rall': return replaceAll(); case 'fclose': return findBar(false);
    case 'zin': return setZoom(D.zoom + 0.1); case 'zout': return setZoom(D.zoom - 0.1);
  }
}
box.addEventListener('click', e => { const b = e.target.closest('[data-c]'); if (b && box.contains(b) && !b.disabled) run(b.dataset.c, b); });
Object.assign(X.ACT, {
  dx: () => openEditor(), dx_color: b => { if (b.dataset.v !== 'auto') D.color = '#' + b.dataset.v; setRunProp('color', b.dataset.v); }, dx_hl: b => { if (b.dataset.v !== 'none') D.hl = b.dataset.v; setRunProp('hl', b.dataset.v === 'none' ? undefined : b.dataset.v); if (b.dataset.v === 'none') setRunProp('shd', undefined); },
  dx_ls: b => paraCmd(pf => { pf.sp = { ...(pf.sp || {}), line: +b.dataset.v, lineRule: 'auto' }; }), dx_sp: b => paraCmd(pf => { pf.sp = { ...(pf.sp || {}), [b.dataset.k]: +b.dataset.v }; }),
  dx_keep: () => paraCmd((pf, p) => { pf.keepNext = !effP(p).keepNext || undefined; }), dx_pbb: () => paraCmd((pf, p) => { pf.pbb = !effP(p).pbb || undefined; }),
  dx_sym: b => { useSel(); snapshot(); document.execCommand('insertText', false, b.dataset.v); }, dx_hf: b => headerFooter(b.dataset.v),
  dx_tbl: b => tableOp(b.dataset.v), dx_cut: () => { useSel(); document.execCommand('cut'); }, dx_copy: () => { useSel(); document.execCommand('copy'); },
  dx_paste: async () => { useSel(); try { const t = await navigator.clipboard.readText(); snapshot(); document.execCommand('insertText', false, t); tidy(); afterEdit(); } catch (e) { toast('Use Ctrl+V to paste', 'info'); } },
  dx_selall: () => { const r = document.createRange(), root = rootOf(getSelection().anchorNode) || docEl; r.selectNodeContents(root); const s = getSelection(); s.removeAllRanges(); s.addRange(r); },
  dx_fix: b => { const x = issueAtCaret(); if (x) fixIssue(x.p, x.is, x.is.sug[+b.dataset.v]); }, dx_link: () => insertLink(), dx_unlink: () => insertLink(),
});
['dx', 'dx_color', 'dx_hl', 'dx_ls', 'dx_sp', 'dx_keep', 'dx_pbb', 'dx_sym', 'dx_hf', 'dx_tbl', 'dx_cut', 'dx_copy', 'dx_paste', 'dx_selall', 'dx_fix', 'dx_link', 'dx_unlink'].forEach(a => X.NO_DOC.add(a));
$('#dxPic').addEventListener('change', e => { const f = e.target.files[0]; e.target.value = ''; if (f) insertPicture(f); });
$('#dxFile').addEventListener('change', e => { const f = e.target.files[0]; e.target.value = ''; if (f) openFile(f); });
for (const r of [hdrEl, docEl, ftrEl]) r.addEventListener('contextmenu', e => {
  e.preventDefault(); const s0 = getSelection(), inside = s0.rangeCount && !s0.isCollapsed && [...s0.getRangeAt(0).getClientRects()].some(q => e.clientX >= q.left && e.clientX <= q.right && e.clientY >= q.top && e.clientY <= q.bottom);
  if (!inside && document.caretRangeFromPoint) { const pos = document.caretRangeFromPoint(e.clientX, e.clientY); if (pos && rootOf(pos.startContainer)) { s0.removeAllRanges(); s0.addRange(pos); } } // (a right-click moves the cursor, unless it is on the selected text)
  const x = issueAtCaret();
  const td = e.target.closest('td'), link = e.target.closest('a[data-w]'), M = X.mi;
  X.showCtx(e, [...(x ? [...x.is.sug.slice(0, 4).map((s, k) => M(s || '(remove)', 'check', 'dx_fix', { d: { v: k }, keep: 1 })), '-'] : []),
    M('Cut', 'cut', 'dx_cut', { key: 'Ctrl+X', keep: 1 }), M('Copy', 'copy', 'dx_copy', { key: 'Ctrl+C', keep: 1 }), M('Paste', 'pages', 'dx_paste', { key: 'Ctrl+V', keep: 1 }), M('Select all', 'selall', 'dx_selall', { key: 'Ctrl+A', keep: 1 }), '-',
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
function fitZoom() { const avail = scroller.clientWidth - 32, w = tw2pt(D.pg.w) * 96 / 72; setZoom(avail > 0 && w > avail ? avail / w : 1); }
$('#dxZoom').addEventListener('input', e => setZoom(+e.target.value / 100));
scroller.addEventListener('wheel', e => { if (!e.ctrlKey) return; e.preventDefault(); setZoom(D.zoom * (e.deltaY < 0 ? 1.1 : 1 / 1.1)); }, { passive: false });
$('#dxDialect').addEventListener('change', e => { X.writing.W.dialect = e.target.value; store.set('xd3.dialect', e.target.value); LINT.cache.clear(); allParas().forEach(p => { p._lt = null; }); lintAll(); });

/* ───────── side panel ───────── */
function setTab(t) { D.tab = t; $$('.dx-tabs button[data-t]', box).forEach(b => b.classList.toggle('on', b.dataset.t === t)); renderSide(); }
$$('.dx-tabs button[data-t]', box).forEach(b => b.onclick = () => setTab(b.dataset.t));
const CATNAME = { spell: 'Spelling', grammar: 'Grammar', improve: 'Word choice', style: 'Clarity' };
function renderSide() {
  if (box.hidden || !D.side) return;
  if (D.tab === 'proof') {
    const list = []; allParas().forEach(p => { if (p._lt === ptext(p)) issuesOf(p).forEach(is => list.push({ p, is })); });
    const by = {}; for (const x of list) (by[CAT(x.is.kind)] = by[CAT(x.is.kind)] || []).push(x); D.plist = list;
    pane.innerHTML = `<div class="wchips">${Object.keys(CATNAME).map(c => `<span class="wchip wm-${c}">${CATNAME[c]} <b>${(by[c] || []).length}</b></span>`).join('')}</div>
      <label class="inl wopt"><input type="checkbox" id="dxStyleChk"${store.get('xd3.style') !== 'off' ? ' checked' : ''}>Clarity suggestions (long sentences, passive voice, weak words)</label>
      ${!LINT.ok ? '<p class="mut">The spelling checker could not start here.</p>' : ''}
      <div class="wcards">${list.length ? list.slice(0, 300).map(({ is }, k) => `<div class="wcard wm-${CAT(is.kind)}" data-k="${k}"><div class="wck wm-${CAT(is.kind)}">${esc(is.pretty)}</div><div class="wcw">${esc(is.word.length > 60 ? is.word.slice(0, 60) + '…' : is.word)}</div><div class="wcm">${esc(is.msg)}</div>${is.sug.length ? `<div class="wps">${is.sug.slice(0, 3).map((s, j) => `<button data-s="${j}">${s ? esc(s) : '<i>remove</i>'}</button>`).join('')}</div>` : ''}</div>`).join('') : `<div class="empty">${LINT.busy ? 'Checking…' : 'No spelling or grammar issues found. 🎉'}</div>`}</div>`;
    const chk = $('#dxStyleChk'); chk.onchange = () => { store.set('xd3.style', chk.checked ? 'on' : 'off'); allParas().forEach(p => { p._lt = null; }); lintAll(); };
  } else if (D.tab === 'format') renderFormat();
  else if (D.tab === 'human') renderHuman();
  else if (D.tab === 'outline') {
    const hs = [...docEl.querySelectorAll('p[data-h]')].filter(p => ptext(p).trim());
    pane.innerHTML = hs.length ? `<div class="dx-outline">${hs.map((p, i) => `<button data-o="${i}" style="padding-left:${8 + Math.max(0, +p.dataset.h - 1) * 14}px" class="${p.dataset.h === '0' ? 'ttl' : ''}">${esc(ptext(p).replace(/\n|\ufffc/g, ' ').trim().slice(0, 90))}</button>`).join('')}</div>` : '<div class="empty">No headings yet. Give titles a Heading style (Ctrl+Alt+1, 2, 3) — or let Auto Format find them.</div>';
    D.olist = hs;
  } else pane.innerHTML = X.writing.statsHtml(paraTexts());
}
const paraTexts = () => [...docEl.querySelectorAll('p')].map(p => ({ text: ptext(p).replace(/\ufffc/g, '').replace(/\n$/, '') }));
pane.addEventListener('click', e => {
  if (D.tab === 'proof') { const card = e.target.closest('.wcard'); if (!card) return; const x = D.plist[+card.dataset.k], s = e.target.closest('[data-s]'); if (!x) return; if (s) return fixIssue(x.p, x.is, x.is.sug[+s.dataset.s]); const P = allParas(); setSel({ a: [P.indexOf(x.p), x.is.a], b: [P.indexOf(x.p), x.is.b] }, true); setTimeout(() => showIssuePop(x.p, x.is), 50); }
  else if (D.tab === 'outline') { const b = e.target.closest('[data-o]'); if (b) { const p = D.olist[+b.dataset.o], P = allParas(); setSel({ a: [P.indexOf(p), 0], b: [P.indexOf(p), 0] }); scroller.scrollTop += p.getBoundingClientRect().top - scroller.getBoundingClientRect().top - 30; X.flash(p, 'wflash'); } }
  else if (D.tab === 'human') onHuman(e);
  else if (D.tab === 'format') onFormat(e);
});

/* ───────── humanize ───────── */
D.tone = 'natural';
function renderHuman() {
  const P = [...docEl.querySelectorAll('p')], list = X.writing.rewrites(D.tone, P.map(p => ({ text: ptext(p) }))); D.hlist = list.map(r => ({ ...r, p: P[r.i] }));
  pane.innerHTML = `<p class="mut">Plainer words, fewer stock phrases and contractions make writing sound like a person${D.tone === 'formal' ? ' — Formal does the opposite' : ''}. Nothing changes until you accept it; the formatting stays.</p>
    <div class="wtone">${[['natural', 'Natural'], ['casual', 'Casual'], ['formal', 'Formal']].map(([v, l]) => `<button data-tone="${v}" class="${D.tone === v ? 'on' : ''}">${l}</button>`).join('')}</div>
    ${list.length ? `<button class="btn primary wfull" data-hall>Accept all ${list.length} change(s)</button>` : ''}
    <div class="wcards">${list.length ? D.hlist.slice(0, 300).map((r, k) => `<div class="wcard wrw" data-k="${k}"><div class="wcw"><s>${esc(r.from)}</s> → <b>${r.to ? esc(r.to) : '<i>remove</i>'}</b></div><div class="wcm">${esc(r.why)}</div><div class="wps"><button data-acc="${k}">Accept</button><button data-skip="${k}">Skip</button></div></div>`).join('') : '<div class="empty">Nothing to change in this tone.</div>'}</div>`;
}
function onHuman(e) {
  const tn = e.target.closest('[data-tone]'); if (tn) { D.tone = tn.dataset.tone; return renderHuman(); }
  const acc = e.target.closest('[data-acc]'), skip = e.target.closest('[data-skip]');
  if (acc) { const r = D.hlist[+acc.dataset.acc]; if (r && ptext(r.p).slice(r.a, r.b) === r.from) { snapshot(); replaceText(r.p, r.a, r.b, r.to); restyle([rootOf(r.p)]); afterEdit(); } return renderHuman(); }
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
  const list = checks(); D.fchecks = list; $('#dxFcnt').textContent = list.length || '✓';
  pane.innerHTML = `<p class="mut">Inconsistent formatting found in this document. Fix one kind at a time, or everything at once — Ctrl+Z undoes it.</p>
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
  const cols = Math.max(1, ...[...t.rows].map(r => [...r.cells].reduce((a, c) => a + (c.colSpan || 1), 0)));
  const tblPr = orig && kid(orig, 'tblPr'); x.appendChild(tblPr ? doc.importNode(tblPr, true) : parseInto(doc, `<w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>`));
  const g = orig && kid(orig, 'tblGrid'), gc = kids(g, 'gridCol'), keep = g && gc.length === cols;
  if (keep) x.appendChild(doc.importNode(g, true)); else { const total = gc.reduce((a, c) => a + (+wa(c, 'w') || 0), 0) || Math.round(contentWidthPt() * 20), ng = wel(doc, 'tblGrid'); for (let i = 0; i < cols; i++) ng.appendChild(wel(doc, 'gridCol', { w: Math.round(total / cols) })); x.appendChild(ng); }
  for (const tr of t.rows) {
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
  written.push(writePart(D.main, docEl, 'body')); if (D.hdrPart) written.push(writePart(D.hdrPart, hdrEl)); if (D.ftrPart) written.push(writePart(D.ftrPart, ftrEl));
  uniqueDocPr(written.filter(Boolean));
  for (const part of Object.values(D.parts)) { if ([D.main, D.hdrPart, D.ftrPart, D.stylesPart, D.numPart].includes(part)) { const x = ser(part.doc); if (x !== part.ser0) D.zip.file(part.name, x, { createFolders: false }); } /* (a part that did not change keeps its original bytes) */ if (part.relsDirty || (part.rels && !D.zip.file(part.relsName))) D.zip.file(part.relsName, ser(part.rels), { createFolders: false }); }
  if (D.ctDirty) D.zip.file('[Content_Types].xml', ser(D.ct), { createFolders: false });
  const core = D.zip.file('docProps/core.xml'); if (core) { try { const c = parseXml(await core.async('string')), m = c.getElementsByTagNameNS('http://purl.org/dc/terms/', 'modified')[0]; if (m) { m.textContent = new Date().toISOString().replace(/\.\d+Z$/, 'Z'); D.zip.file('docProps/core.xml', ser(c), { createFolders: false }); } } catch (e) { /* left as it is */ } }
  for (const n of Object.keys(D.zip.files)) if (D.zip.files[n].dir) delete D.zip.files[n]; // (only the folder entry — remove() would take its files too)
  return D.zip.generateAsync({ type: 'blob', compression: 'DEFLATE', compressionOptions: { level: 6 }, mimeType: DOCX_MIME });
}
const fileName = () => ($('#dxName').value.trim().replace(/\.docx$/i, '').replace(/[\\/:*?"<>|]+/g, '-') || 'Document') + '.docx';
async function save() { if (!D.zip) return; await busy(async () => { download(await buildDocx(), fileName()); D.dirty = false; $('#dxState').textContent = 'Saved'; toast('Word document saved to your downloads', 'success'); }, 'Saving…'); }
function printDoc() { document.documentElement.classList.add('dxprint'); const done = () => { document.documentElement.classList.remove('dxprint'); removeEventListener('afterprint', done); }; addEventListener('afterprint', done); setTimeout(() => { window.print(); setTimeout(done, 1000); }, 50); }
async function savePdf() {
  if (X.desk && X.desk.printToPdf) { // the desktop app prints the page itself to PDF — exactly what is on screen
    await busy(async () => {
      document.documentElement.classList.add('dxprint'); await new Promise(r => setTimeout(r, 120));
      try { const P = D.pg, inch = v => Math.abs(v) / 1440, bytes = await X.desk.printToPdf({ width: inch(P.w), height: inch(P.h), margins: { top: inch(P.t), bottom: inch(P.b), left: inch(P.l), right: inch(P.r) } }); download(new Blob([bytes], { type: 'application/pdf' }), fileName().replace(/\.docx$/, '.pdf')); toast('PDF saved to your downloads', 'success'); }
      finally { document.documentElement.classList.remove('dxprint'); }
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
  RFC.clear(); PFC.clear(); RCSS.clear(); stacks.clear(); LINT.cache.clear(); LINT.ignored.clear();
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
  docEl.replaceChildren(renderRoot(body, P));
  hdrEl.replaceChildren(D.hdrPart ? renderRoot(D.hdrPart.doc.documentElement, D.hdrPart) : ''); hdrEl.hidden = !D.hdrPart;
  ftrEl.replaceChildren(D.ftrPart ? renderRoot(D.ftrPart.doc.documentElement, D.ftrPart) : ''); ftrEl.hidden = !D.ftrPart;
  $('#dxName').value = name; $('#dxState').textContent = ''; document.title = name + ' — XD3 Word editor';
  pageGeometry(); tidy(roots().filter(r => !r.hidden)); fitZoom(); dirty(); D.dirty = false; $('#dxState').textContent = '';
  scroller.scrollTop = 0; afterEdit(true);
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
async function newDocument() { if (D.open && D.dirty && !confirm('Start a new document? Unsaved changes will be lost.')) return; show(); await loadZip(await blankZip(), 'Document'); }
function show() { if (!D.open) { D.open = true; if (innerWidth < 860 && store.get('xd3.dxside') == null) D.side = false; // (on a phone the page comes first)
  box.hidden = false; document.body.classList.add('modal', 'dxopen'); box.classList.toggle('noside', !D.side); setTab(D.tab); $('#dxDialect').value = X.writing ? X.writing.W.dialect : 'American'; X.writing && X.writing.engine().catch(() => { LINT.ok = false; }); } }
async function openEditor() { show(); if (!D.zip) await newDocument(); }
function close(force) {
  if (!force && D.dirty && !confirm('Close the Word editor? Unsaved changes will be lost.')) return;
  D.open = false; box.hidden = true; pop.hidden = true; document.body.classList.remove('modal', 'dxopen'); X.updateUI();
  if (window.CSS && CSS.highlights) for (const k of ['dx-spell', 'dx-grammar', 'dx-improve', 'dx-style', 'dx-find', 'dx-findcur']) CSS.highlights.delete(k);
}
window.addEventListener('beforeunload', e => { if (D.open && D.dirty) { e.preventDefault(); e.returnValue = ''; } });
window.addEventListener('resize', () => { if (!box.hidden) { clearTimeout(paginate.s); paginate.s = setTimeout(() => { paginate(); placeGrip(); }, 120); } });
$('#dxName').addEventListener('keydown', e => { e.stopPropagation(); if (e.key === 'Enter') e.target.blur(); });
$('#dxName').addEventListener('change', () => { document.title = $('#dxName').value + ' — XD3 Word editor'; dirty(); });

X.docx = { D, openFile, newDocument, open: openEditor, close, isOpen: () => D.open, buildDocx, checks, autoFormat, applyFixes, setStyle, listToggle, setRunProp, getSel, setSel, allParas, ptext, undo, redo, tidy, loadZip, blankZip, replaceText, lintAll, insertTable, tableOp, align };
})();
