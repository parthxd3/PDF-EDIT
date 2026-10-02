/* XD3 PDF Editor — © XD3Labs / XDCybertech Pvt Ltd. Runs fully offline, client-side (pdf.js + pdf-lib, bundled in /lib). */
(() => {
'use strict';

pdfjsLib.GlobalWorkerOptions.workerSrc = 'lib/pdf.worker.min.js';
// isEvalSupported: false closes a known pdf.js hole where a crafted font inside a PDF could run script (CVE-2024-4367)
const PDFJS_OPTS = { cMapUrl: 'lib/cmaps/', cMapPacked: true, standardFontDataUrl: 'lib/standard_fonts/', isEvalSupported: false, fontExtraProperties: true }; // (extra font data: edited text reuses the document's own fonts)

const CSS_UNITS = 96 / 72, LH = 1.2, BASE = 0.93, A4 = [595.28, 841.89];
const SVGNS = 'http://www.w3.org/2000/svg';
const COARSE = matchMedia('(pointer: coarse)').matches;
const FONTS = { Helvetica: 'Arial, Helvetica, sans-serif', Times: '"Times New Roman", Times, serif', Courier: '"Courier New", Courier, monospace' };
const PDF_FONTS = {
  Helvetica: ['Helvetica', 'HelveticaBold', 'HelveticaOblique', 'HelveticaBoldOblique'],
  Times: ['TimesRoman', 'TimesRomanBold', 'TimesRomanItalic', 'TimesRomanBoldItalic'],
  Courier: ['Courier', 'CourierBold', 'CourierOblique', 'CourierBoldOblique'],
};
// fonts shipped in lib/fonts and embedded into saved PDFs: family → [file stem, has a bold file, CSS fallback, group]
const BUNDLED = {
  'Roboto': ['Roboto', 1, 'sans-serif', 'Sans serif'], 'Open Sans': ['OpenSans', 1, 'sans-serif', 'Sans serif'], 'Lato': ['Lato', 1, 'sans-serif', 'Sans serif'],
  'Montserrat': ['Montserrat', 1, 'sans-serif', 'Sans serif'], 'Poppins': ['Poppins', 1, 'sans-serif', 'Sans serif'], 'Raleway': ['Raleway', 1, 'sans-serif', 'Sans serif'],
  'Nunito': ['Nunito', 1, 'sans-serif', 'Sans serif'], 'Ubuntu': ['Ubuntu', 1, 'sans-serif', 'Sans serif'], 'Hind': ['Hind', 1, 'sans-serif', 'Sans serif'],
  'Playfair Display': ['PlayfairDisplay', 1, 'serif', 'Serif'], 'Lora': ['Lora', 1, 'serif', 'Serif'], 'Libre Baskerville': ['LibreBaskerville', 1, 'serif', 'Serif'],
  'PT Serif': ['PTSerif', 1, 'serif', 'Serif'], 'Crimson Text': ['CrimsonText', 1, 'serif', 'Serif'], 'Roboto Slab': ['RobotoSlab', 1, 'serif', 'Serif'],
  'Roboto Mono': ['RobotoMono', 1, 'monospace', 'Monospace'], 'Source Code Pro': ['SourceCodePro', 1, 'monospace', 'Monospace'],
  'Oswald': ['Oswald', 1, 'sans-serif', 'Display'], 'Anton': ['Anton', 0, 'sans-serif', 'Display'], 'Lobster': ['Lobster', 0, 'cursive', 'Display'], 'Comic Neue': ['ComicNeue', 1, 'cursive', 'Display'], 'Permanent Marker': ['PermanentMarker', 0, 'cursive', 'Display'],
  'Caveat': ['Caveat', 1, 'cursive', 'Handwriting'], 'Pacifico': ['Pacifico', 0, 'cursive', 'Handwriting'], 'Satisfy': ['Satisfy', 0, 'cursive', 'Handwriting'], 'Courgette': ['Courgette', 0, 'cursive', 'Handwriting'],
  'Kaushan Script': ['KaushanScript', 0, 'cursive', 'Handwriting'], 'Sacramento': ['Sacramento', 0, 'cursive', 'Handwriting'], 'Allura': ['Allura', 0, 'cursive', 'Handwriting'], 'Great Vibes': ['GreatVibes', 0, 'cursive', 'Handwriting'],
};
for (const n in BUNDLED) FONTS[n] = `"XD3 ${n}", ${BUNDLED[n][2]}`;
// Ligatures, contextual alternates and kerning are switched off for bundled fonts — on screen (FEAT_CSS) and in the saved PDF (FEAT_OFF) —
// so every character stays one searchable glyph with a known width, and the page looks the same in both places.
const FEAT_OFF = { ccmp: false, locl: false, rlig: false, calt: false, clig: false, liga: false, rclt: false, dlig: false, frac: false, numr: false, dnom: false, kern: false, curs: false };
const FEAT_CSS = '"liga" 0, "clig" 0, "calt" 0, "rlig" 0, "rclt" 0, "ccmp" 0, "locl" 0, "kern" 0';
const fontReq = {}, fontData = {};
function ensureFont(name) { // start loading a bundled (or document) font; resolves once it can be measured and drawn
  if (S.xfonts[name]) return useXFont(name);
  const B = BUNDLED[name]; if (!B) return Promise.resolve();
  return fontReq[name] || (fontReq[name] = Promise.all([['Regular', '400'], B[1] && ['Bold', '700']].filter(Boolean).map(([s, w]) => {
    const ff = new FontFace('XD3 ' + name, `url(lib/fonts/${B[0]}-${s}.ttf)`, name === 'Hind' ? { weight: w } : { weight: w, featureSettings: FEAT_CSS }); document.fonts.add(ff); return ff.load(); // (Hind keeps its default shaping, which Devanagari needs)
  })).then(fontsChanged, () => { /* file missing — the fallback family is used */ }));
}
function fontsChanged() { clearTimeout(fontsChanged.t); fontsChanged.t = setTimeout(() => { S.pages.forEach(drawObjs); drawSel(); }, 30); } // text widths change once the real font arrives
const fontBytes = stem => fontData[stem] || (fontData[stem] = fetch(`lib/fonts/${stem}.ttf`).then(r => { if (!r.ok) throw new Error('font file missing'); return r.arrayBuffer(); }));
const fontFamilyFor = pdfName => { const k = String(pdfName || '').toLowerCase().replace(/[^a-z]/g, ''); return Object.keys(BUNDLED).sort((a, b) => b.length - a.length).find(n => k.includes(n.toLowerCase().replace(/ /g, ''))) || null; };
// Fonts lifted out of the opened PDFs (S.xfonts: id → { name, ps, bytes, bold, italic, fb }). Text edited in place keeps the document's own typeface,
// on screen and in the saved file; characters the embedded (usually subsetted) font does not contain fall back to the nearest family, `fb`.
const xfReady = {};
const cleanFontName = n => {
  let s = String(n || '').replace(/^[A-Z]{6}\+/, '').split(/[-,]/)[0].replace(/(PS)?MT$|PS$/, ''), t;
  do { t = s; s = s.replace(/(Bold|Italic|Oblique|Regular|Medium|Semibold|Light|Black|Heavy)$/, ''); } while (s !== t && s.length > 3);
  return s.replace(/([a-z])([A-Z])/g, '$1 $2').replace(/[^\w .-]/g, '').trim() || 'Document font';
};
function xfId(name, bytes) { let h = 2166136261 ^ bytes.length; const st = Math.max(1, bytes.length >> 11); for (let i = 0; i < bytes.length; i += st) h = Math.imul(h ^ bytes[i], 16777619); for (const ch of name) h = Math.imul(h ^ ch.charCodeAt(0), 16777619); return 'pf' + (h >>> 0).toString(36); }
function useXFont(id) { // register a document font with the browser so it can be measured and drawn
  const x = S.xfonts[id]; if (!x) return Promise.resolve(); if (xfReady[id]) return xfReady[id];
  FONTS[id] = `"XD3F ${id}", ${FONTS[x.fb] || FONTS.Helvetica}`;
  let load; try { const ff = new FontFace('XD3F ' + id, x.bytes, { weight: x.bold ? '700' : '400', style: x.italic ? 'italic' : 'normal', featureSettings: FEAT_CSS }); document.fonts.add(ff); load = ff.load().catch(() => { /* unusable font data — the fallback family shows */ }); } catch (e) { load = Promise.resolve(); }
  return xfReady[id] = Promise.all([load, ensureFont(x.fb)]).then(fontsChanged);
}
const fontLabel = f => { const x = S.xfonts[f]; return x ? x.name + (x.bold ? ' Bold' : '') + (x.italic ? ' Italic' : '') : f; };
const AL = { left: 0, center: 0.5, right: 1 };
const ICONS = {
  folder: 'M3 7a2 2 0 0 1 2-2h4l2 2h8a2 2 0 0 1 2 2v8a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z',
  download: 'M12 3v12m0 0l-4-4m4 4l4-4M4 19h16', upload: 'M12 16V4m0 0L8 8m4-4l4 4M4 20h16',
  print: 'M6 9V3h12v6M6 18H4a2 2 0 0 1-2-2v-5a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2v5a2 2 0 0 1-2 2h-2M6 14h12v7H6z',
  undo: 'M9 14L4 9l5-5M4 9h10a6 6 0 0 1 0 12h-3', redo: 'M15 14l5-5-5-5M20 9H10a6 6 0 0 0 0 12h3',
  zoomin: 'M11 4a7 7 0 1 0 0 14 7 7 0 0 0 0-14zM21 21l-5-5M11 8v6M8 11h6', zoomout: 'M11 4a7 7 0 1 0 0 14 7 7 0 0 0 0-14zM21 21l-5-5M8 11h6',
  search: 'M11 4a7 7 0 1 0 0 14 7 7 0 0 0 0-14zM21 21l-5-5', fit: 'M4 9V4h5M20 9V4h-5M4 15v5h5M20 15v5h-5',
  up: 'M6 15l6-6 6 6', down: 'M6 9l6 6 6-6', pages: 'M7 3h8l4 4v14H7zM15 3v4h4', plus: 'M12 5v14M5 12h14',
  copy: 'M8 8h12v12H8zM4 16V4h12', rotate: 'M20 11a8 8 0 1 0-2.3 5.7M20 4v7h-7', rotatel: 'M4 11a8 8 0 1 1 2.3 5.7M4 4v7h7',
  trash: 'M4 7h16M9 7V4h6v3M6 7l1 13h10l1-13', merge: 'M6 4v5a6 6 0 0 0 6 6v5M18 4v5a6 6 0 0 1-6 6',
  split: 'M12 4v6M12 10l-6 6v4M12 10l6 6v4', image: 'M4 5h16v14H4zM4 16l5-5 4 4 3-3 4 4M9 9h.01',
  sign: 'M3 17c2-6 4-10 6-10s-1 9 2 9 3-6 5-6 1 4 5 4M3 21h18', drop: 'M12 3s6 6 6 11a6 6 0 0 1-12 0c0-5 6-11 6-11z',
  hash: 'M5 9h14M5 15h14M10 4L8 20M16 4l-2 16', stamp: 'M5 21h14M6 17h12v-3H6zM10 14V9a2 2 0 1 1 4 0v5',
  check: 'M5 12l5 5L20 7', x: 'M6 6l12 12M18 6L6 18', more: 'M5 12h.01M12 12h.01M19 12h.01',
  info: 'M12 3a9 9 0 1 0 0 18 9 9 0 0 0 0-18zM12 11v6M12 7h.01', key: 'M3 7h18v10H3zM7 11h.01M11 11h.01M15 11h.01M8 14h8',
  panel: 'M3 5h18v14H3zM9 5v14', moon: 'M20 14A8 8 0 1 1 10 4a6 6 0 0 0 10 10z',
  cursor: 'M5 3l14 8-6 2-2 6z', text: 'M5 6V4h14v2M12 4v16M9 20h6', edittext: 'M4 20h4L19 9l-4-4L4 16zM13 7l4 4',
  editobj: 'M4 4h4M4 4v4M20 4h-4M20 4v4M4 20h4M4 20v-4M20 20h-4M20 20v-4M9 9h6v6H9z',
  hand: 'M8 13V5a1.5 1.5 0 0 1 3 0v6V4a1.5 1.5 0 0 1 3 0v7V6a1.5 1.5 0 0 1 3 0v9c0 4-2 6-6 6s-5-2-7-6l-1-2a1.5 1.5 0 0 1 2.6-1.5z',
  note: 'M4 4h16v11l-5 5H4zM15 20v-5h5M8 9h8M8 13h4', cal: 'M4 6h16v14H4zM4 10h16M8 3v4M16 3v4',
  pen: 'M3 17c3-6 5-6 7-2s4 4 6-2 3-6 5-4', marker: 'M9 11l-5 5v4h4l5-5M14 4l6 6-8 8-6-6z', hl: 'M4 7h16v10H4zM8 12h8',
  rect: 'M4 6h16v12H4z', ellipse: 'M12 5c5 0 9 3 9 7s-4 7-9 7-9-3-9-7 4-7 9-7z', line: 'M5 19L19 5', arrow: 'M5 19L19 5M10 5h9v9',
  eraser: 'M7 21h10M5 13l8-8 6 6-8 8H8z', redact: 'M3 8h18v8H3zM3 12h18M7 8v8M12 8v8M17 8v8', stroke: 'M4 6h16M4 11h16M4 17h16',
  front: 'M12 19V5M5 12l7-7 7 7', back: 'M12 5v14M5 12l7 7 7-7', lock: 'M6 11h12v9H6zM8 11V8a4 4 0 0 1 8 0v3',
  alignleft: 'M4 6h16M4 12h10M4 18h13', aligncenter: 'M4 6h16M7 12h10M5 18h14', alignright: 'M4 6h16M10 12h10M7 18h13',
  grid: 'M4 4h7v7H4zM13 4h7v7h-7zM4 13h7v7H4zM13 13h7v7h-7z', swap: 'M7 4L3 8l4 4M3 8h14M17 12l4 4-4 4M21 16H7',
  tools: 'M15 4a5 5 0 0 0-4.6 6.9L4 17.3V20h2.7l6.4-6.4A5 5 0 0 0 20 9l-3 3-3-1-1-3z', scan: 'M4 8V4h4M20 8V4h-4M4 16v4h4M20 16v4h-4M3 12h18',
  word: 'M6 3h9l4 4v14H6zM15 3v4h4M9 12l1.5 5 1.5-4 1.5 4 1.5-5', layers: 'M12 4l9 5-9 5-9-5zM3 14l9 5 9-5',
  sliders: 'M4 6h10M18 6h2M4 12h4M12 12h8M4 18h12M14 4v4M8 10v4M16 16v4', wand: 'M4 20L15 9M14 3v3M12.5 4.5h3M19 9v3M17.5 10.5h3M19 3l1 1',
  corner: 'M4 20V10a6 6 0 0 1 6-6h10', left: 'M15 6l-6 6 6 6', right: 'M9 6l6 6-6 6',
  al_l: 'M4 3v18M8 7h10v4H8zM8 14h6v4H8z', al_c: 'M12 3v18M6 7h12v4H6zM8 14h8v4H8z', al_r: 'M20 3v18M6 7h10v4H6zM10 14h6v4h-6z',
  al_t: 'M3 4h18M7 8h4v10H7zM14 8h4v6h-4z', al_m: 'M3 12h18M7 6h4v12H7zM14 8h4v8h-4z', al_b: 'M3 20h18M7 6h4v10H7zM14 10h4v6h-4z',
  dist_h: 'M4 4v16M20 4v16M9 8h6v8H9z', dist_v: 'M4 4h16M4 20h16M8 9h8v6H8z', brush: 'M4 20c3 0 4-2 4-4l8-9 3 3-9 8c-2 0-3 1-6 2zM14 5l5 5',
  selall: 'M4 4h4M4 4v4M20 4h-4M20 4v4M4 20h4M4 20v-4M20 20h-4M20 20v-4M9 12l2 2 4-4', compress: 'M12 3v6M9 6l3 3 3-3M12 21v-6M9 18l3-3 3 3M4 12h16', minus: 'M5 12h14',
  bell: 'M6 16V11a6 6 0 0 1 12 0v5l2 2H4zM10 20a2 2 0 0 0 4 0', bolt: 'M13 3L5 13h6l-1 8 8-10h-6z', sun: 'M12 8a4 4 0 1 0 0 8 4 4 0 0 0 0-8zM12 2v2M12 20v2M2 12h2M20 12h2M5 5l1.5 1.5M17.5 17.5L19 19M5 19l1.5-1.5M17.5 6.5L19 5',
  cut: 'M6 6a2 2 0 1 0 0 4 2 2 0 0 0 0-4zM6 14a2 2 0 1 0 0 4 2 2 0 0 0 0-4zM8 9l12 8M8 15L20 7',
};
const HINTS = {
  select: 'Click to select · drag to move · corner handles resize · top handle rotates · Shift+click or drag a box to select several · double‑click text to edit',
  hand: 'Drag to pan around the document',
  edittext: 'Click any outlined text to rewrite it — the original font, size and colour are kept · text inside scans and pictures is recognised automatically (OCR) · ¶ Paragraphs edits whole blocks',
  editobj: 'Click an outlined image, or drag a box around anything on the page (it snaps to the content), to pick it up — then move, resize, rotate or delete it',
  text: 'Click anywhere to add text · change font, size and colour in the bar above while typing',
  note: 'Click to place a sticky note',
  pen: 'Draw freehand', marker: 'Drag over content to highlight freehand', hl: 'Drag a box to highlight an area',
  rect: 'Drag to draw a rectangle (Shift = square)', ellipse: 'Drag to draw an ellipse (Shift = circle)',
  line: 'Drag to draw a line (Shift = 45° steps)', arrow: 'Drag to draw an arrow (Shift = 45° steps)',
  whiteout: 'Drag a box to cover content with white', redact: 'Drag a box to black out content (visual cover only)',
};

const $ = (s, r = document) => r.querySelector(s);
const $$ = (s, r = document) => [...r.querySelectorAll(s)];
const svg = (n, a = {}, parent) => { const e = document.createElementNS(SVGNS, n); for (const k in a) e.setAttribute(k, a[k]); if (parent) parent.appendChild(e); return e; };
const uid = () => Math.random().toString(36).slice(2, 10);
const clone = o => JSON.parse(JSON.stringify(o));
const clamp = (v, a, b) => Math.min(b, Math.max(a, v));
const esc = s => String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
const r2 = v => Math.round(v * 100) / 100;

const viewer = $('#viewer'), thumbs = $('#thumbs'), dlg = $('#dlg'), ctxMenu = $('#ctx');
const S = {
  sources: [], pages: [], assets: {}, zoom: 1, tool: 'select', sel: null, cur: 0, srcVer: 0,
  undo: [], redo: [], fileName: 'document.pdf', dirty: false, meta: {}, editing: null, clip: null,
  style: { color: '#111111', fill: 'none', width: 2, dash: 'solid', rx: 0, size: 16, font: 'Helvetica', bold: false, italic: false, underline: false, strike: false, align: 'left', lh: 1.2, opacity: 1 },
  para: false, xfonts: {}, styleClip: null, ctxPt: null, notes: [], unread: 0,
  find: { q: '', hits: [], cur: -1 },
};
const PE = new Map(), visible = new Set();
let io = null, tio = null;

/* ───────── helpers ───────── */
const NOTE_ICON = { success: 'check', error: 'x', warn: 'info', info: 'info' };
const SUCCESS_RE = /saved|added|copied|restored|merged|replaced|compressed|removed|resized|deleted|picked up|converted|protected|inserted|extracted|unlocked/i;
function toast(msg, kind) { // kind: true or 'error' | 'success' | 'warn' | 'info'; left out, it is guessed from the wording
  const type = kind === true ? 'error' : typeof kind === 'string' ? kind : SUCCESS_RE.test(msg) ? 'success' : 'info';
  S.notes.unshift({ msg, type, t: Date.now() }); if (S.notes.length > 40) S.notes.length = 40;
  if (!$('#mNotes').classList.contains('open')) { S.unread++; updateBell(true); }
  const host = $('#toasts'); if (dlg.open && host.parentNode !== dlg) dlg.appendChild(host); // a modal dialog covers everything outside itself
  const t = document.createElement('div'); t.className = 'toast t-' + type;
  t.innerHTML = `<i data-icon="${NOTE_ICON[type]}"></i><div class="tmsg"></div><button class="tx" title="Dismiss">&times;</button><div class="tbar"></div>`;
  $('.tmsg', t).textContent = msg; icons(t); host.appendChild(t);
  while (host.children.length > 4) host.firstChild.remove();
  const life = type === 'error' || type === 'warn' ? 6500 : 3600; $('.tbar', t).style.animationDuration = life + 'ms';
  const close = () => { if (t.classList.contains('out')) return; t.classList.add('out'); setTimeout(() => t.remove(), 300); };
  let timer = setTimeout(close, life);
  t.onmouseenter = () => clearTimeout(timer); t.onmouseleave = () => { timer = setTimeout(close, 1500); }; t.onclick = close; // hovering keeps it on screen
}
function updateBell(ring) {
  const b = $('#bBell'), n = $('.badge', b); n.hidden = !S.unread; n.textContent = S.unread > 9 ? '9+' : S.unread;
  if (ring) { b.classList.remove('ringing'); void b.offsetWidth; b.classList.add('ringing'); }
}
function flash(el, cls) { if (!el) return; el.classList.remove(cls); void el.offsetWidth; el.classList.add(cls); } // replay a one-shot CSS animation
async function busy(fn, txt = 'Working…') {
  $('#busyTxt').textContent = txt; $('#busy').hidden = false;
  try { return await fn(); } catch (e) { console.error(e); toast(e.message || String(e), true); }
  finally { $('#busy').hidden = true; }
}
function dialog(title, html, { ok = 'OK', onOpen, onClose, wide, cancel = true } = {}) {
  return new Promise(res => {
    document.body.appendChild($('#toasts')); // (it may have been parked inside the previous dialog)
    dlg.className = wide ? 'wide' : '';
    dlg.innerHTML = `<form method="dialog"><h3>${title}</h3><div class="dbody">${html}</div><div class="dfoot"><button value="ok" class="btn primary">${ok}</button>${cancel ? '<button value="cancel" class="btn" formnovalidate>Cancel</button>' : ''}</div></form>`;
    if (onOpen) onOpen(dlg);
    dlg.onclose = () => { const r = dlg.returnValue === 'ok' ? Object.fromEntries(new FormData(dlg.firstChild)) : null; document.body.appendChild($('#toasts')); if (onClose) onClose(r); res(r); };
    dlg.returnValue = ''; dlg.showModal();
  });
}
function download(blob, name) {
  const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = name;
  document.body.appendChild(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(a.href), 30000);
}
const pageById = id => S.pages.find(p => p.id === id);
const rdims = p => p.rot % 180 ? [p.h, p.w] : [p.w, p.h];
const scaleCss = () => S.zoom * CSS_UNITS;
const selObjs = () => { if (!S.sel) return null; const p = pageById(S.sel.pid); if (!p) return null; const os = p.objs.filter(o => S.sel.ids.includes(o.id)); return os.length ? { p, os } : null; };
const selOne = () => { const s = selObjs(); return s && s.os.length === 1 ? { p: s.p, o: s.os[0] } : null; };
const editingObj = () => { if (!S.editing) return null; const p = pageById(S.editCtx.pid); return p && p.objs.find(o => o.id === S.editing) || null; };
const baseName = () => S.fileName.replace(/\.pdf$/i, '');
const getPdfPage = p => S.sources[p.src].pdf.getPage(p.idx + 1);

/* ───────── session autosave (IndexedDB) ───────── */
const DB = {
  open() { return this.db || (this.db = new Promise((res, rej) => { const r = indexedDB.open('xd3pdf', 1); r.onupgradeneeded = () => r.result.createObjectStore('kv'); r.onsuccess = () => res(r.result); r.onerror = () => rej(r.error); })); },
  async tx(mode, fn) { const db = await this.open(); return new Promise((res, rej) => { const t = db.transaction('kv', mode), rq = fn(t.objectStore('kv')); t.oncomplete = () => res(rq.result); t.onerror = () => rej(t.error); }); },
  get: k => DB.tx('readonly', s => s.get(k)), put: (k, v) => DB.tx('readwrite', s => s.put(v, k)),
};
let savedSrcVer = -1;
function scheduleSave() { clearTimeout(scheduleSave.t); scheduleSave.t = setTimeout(saveSession, 1500); }
async function saveSession() {
  try {
    if (!S.pages.length || S.sources.reduce((a, s) => a + s.bytes.length, 0) > 2e8) return;
    if (savedSrcVer !== S.srcVer) { await DB.put('src', S.sources.map(s => s.bytes)); savedSrcVer = S.srcVer; }
    const used = new Set(), xf = {}; for (const p of S.pages) for (const o of p.objs) if (o.type === 'text') used.add(o.font);
    for (const id in S.xfonts) if (used.has(id)) xf[id] = S.xfonts[id];
    await DB.put('session', { fileName: S.fileName, meta: S.meta, pages: S.pages, assets: S.assets, xfonts: xf, n: S.sources.length });
  } catch (e) { /* storage unavailable — autosave is best-effort */ }
}
async function restoreSession() {
  await busy(async () => {
    const ses = await DB.get('session'), srcs = await DB.get('src') || [];
    if (!ses || !ses.pages || srcs.length < ses.n) throw new Error('No saved session found');
    resetDoc();
    for (const b of srcs.slice(0, ses.n)) await addSource(b);
    savedSrcVer = S.srcVer;
    Object.assign(S, { pages: ses.pages, assets: ses.assets || {}, xfonts: ses.xfonts || {}, fileName: ses.fileName || 'document.pdf', meta: ses.meta || {}, dirty: true });
    buildPages(); S.zoom = clamp(fitZoom(), 0.25, 1.25); layout(); syncControls(); updateUI(); toast('Session restored');
  }, 'Restoring…');
}

/* ───────── history ───────── */
const snap = () => JSON.stringify(S.pages);
const structSig = () => S.pages.map(p => p.id + ':' + p.rot).join(',');
function pushUndo(s) { S.undo.push(s || snap()); if (S.undo.length > 80) S.undo.shift(); S.redo.length = 0; S.dirty = true; updateUI(); scheduleSave(); }
function restore(s) {
  commitEdit(); const sig = structSig(); S.pages = JSON.parse(s); S.sel = null;
  if (structSig() !== sig) { const y = viewer.scrollTop; buildPages(); viewer.scrollTop = y; } else S.pages.forEach(drawObjs);
  drawSel(); syncControls(); updateUI(); scheduleSave();
}
function undo() { if (!S.undo.length) return; const s = snap(); const u = S.undo.pop(); S.redo.push(s); restore(u); }
function redo() { if (!S.redo.length) return; const s = snap(); const r = S.redo.pop(); S.undo.push(s); restore(r); }

/* ───────── loading ───────── */
function resetDoc() {
  commitEdit();
  Object.assign(S, { sources: [], pages: [], assets: {}, xfonts: {}, undo: [], redo: [], sel: null, cur: 0, meta: {}, dirty: false, find: { q: '', hits: [], cur: -1 } });
  S.srcVer++;
}
async function addSource(bytes) {
  const task = pdfjsLib.getDocument({ data: bytes.slice(), ...PDFJS_OPTS });
  task.onPassword = (cb, reason) => { // (window.prompt is unavailable in the desktop build, so use our own dialog)
    dialog('Password required', `<p>${reason === 1 ? 'This PDF is password protected.' : 'That password was not correct — try again.'}</p><label>Password<input type="password" name="pw" autocomplete="off" style="height:34px;width:100%;background:var(--panel2);border:1px solid var(--line);border-radius:7px;padding:0 7px;color:var(--txt)"></label>`,
      { ok: 'Unlock', onOpen: d => setTimeout(() => $('[name=pw]', d).focus(), 50) }).then(r => r ? cb(r.pw) : task.destroy());
  };
  const pdf = await task.promise; S.srcVer++;
  return S.sources.push({ bytes, pdf, text: {}, scan: {}, blocks: {}, fonts: {} }) - 1;
}
async function loadSource(bytes) {
  const si = await addSource(bytes), pdf = S.sources[si].pdf, pages = [];
  for (let i = 0; i < pdf.numPages; i++) {
    const vp = (await pdf.getPage(i + 1)).getViewport({ scale: 1 });
    pages.push({ id: uid(), src: si, idx: i, w: vp.width, h: vp.height, rot: 0, objs: [] });
  }
  return pages;
}
function readImage(file) {
  return new Promise((res, rej) => {
    const fr = new FileReader();
    fr.onerror = () => rej(new Error('Could not read image'));
    fr.onload = () => {
      const im = new Image();
      im.onerror = () => rej(new Error('Unsupported image: ' + file.name));
      im.onload = () => {
        let url = fr.result, w = im.naturalWidth || 300, h = im.naturalHeight || 150; const jpeg = /^data:image\/jpeg/.test(url), big = Math.max(w, h) > 3600;
        if (big || !/^data:image\/(png|jpeg)/.test(url)) { // normalise webp/gif/svg/… for pdf-lib, and shrink oversized photos so editing and saving stay fast
          const f = big ? 3600 / Math.max(w, h) : 1, c = document.createElement('canvas'); c.width = Math.round(w * f); c.height = Math.round(h * f);
          c.getContext('2d').drawImage(im, 0, 0, c.width, c.height); url = jpeg ? c.toDataURL('image/jpeg', 0.9) : c.toDataURL('image/png'); w = c.width; h = c.height;
        }
        res({ url, w, h });
      };
      im.src = fr.result;
    };
    fr.readAsDataURL(file);
  });
}
function addAsset(url) { const id = 'a' + uid(); S.assets[id] = url; return id; }
function imagePage(im) {
  let w = im.w * 0.75, h = im.h * 0.75; const f = Math.min(1, A4[0] / w, A4[1] / h); w *= f; h *= f;
  return { id: uid(), src: -1, w, h, rot: 0, objs: [{ id: uid(), type: 'image', asset: addAsset(im.url), x: 0, y: 0, w, h, opacity: 1 }] };
}
const isPdf = f => f.type === 'application/pdf' || /\.pdf$/i.test(f.name);
const isDocx = f => /\.docx$/i.test(f.name);
async function openFiles(files, append) {
  if ([...files].some(f => /\.doc$/i.test(f.name))) toast('Old .doc files are not supported — save as .docx in Word first', true);
  files = [...files].filter(f => isPdf(f) || isDocx(f) || f.type.startsWith('image/'));
  if (!files.length) return toast('Please choose PDF, Word (.docx) or image files', true);
  if (!append && S.dirty && S.pages.length && !confirm('Discard unsaved changes and open a new document?')) return;
  await busy(async () => {
    if (append) pushUndo(); else resetDoc();
    let first = !append;
    try {
      for (const f of files) {
        if (isPdf(f)) S.pages.push(...await loadSource(new Uint8Array(await f.arrayBuffer())));
        else if (isDocx(f)) S.pages.push(...await window.XD3.docxToPages(f));
        else S.pages.push(imagePage(await readImage(f)));
        if (first) { S.fileName = f.name.replace(/\.\w+$/, '') + '.pdf'; first = false; }
      }
    } finally {
      buildPages();
      if (!append && S.pages.length) { S.zoom = clamp(fitZoom(), 0.25, 1.25); layout(); viewer.scrollTop = 0; }
      setTool('select'); syncControls(); updateUI(); scheduleSave();
    }
    if (append) toast(`Added ${files.length} file(s)`);
  }, 'Opening…');
}
function newDoc() {
  if (S.dirty && S.pages.length && !confirm('Discard unsaved changes and start a new document?')) return;
  resetDoc(); S.fileName = 'untitled.pdf';
  S.pages.push({ id: uid(), src: -1, w: A4[0], h: A4[1], rot: 0, objs: [] });
  buildPages(); S.zoom = clamp(fitZoom(), 0.25, 1.25); layout(); setTool('select'); syncControls(); updateUI();
}

/* ───────── page DOM ───────── */
const ROT_TF = p => p.rot === 90 ? `matrix(0,1,-1,0,${p.h},0)` : p.rot === 180 ? `matrix(-1,0,0,-1,${p.w},${p.h})` : p.rot === 270 ? `matrix(0,-1,1,0,0,${p.w})` : '';
function buildPages() {
  commitEdit();
  if (io) io.disconnect(); if (tio) tio.disconnect();
  viewer.textContent = ''; thumbs.textContent = ''; PE.clear(); visible.clear();
  io = new IntersectionObserver(ens => ens.forEach(en => {
    const id = en.target.dataset.id, p = pageById(id);
    if (en.isIntersecting) { visible.add(id); if (p) { renderCanvas(p); drawTI(p); } } else visible.delete(id);
  }), { root: viewer, rootMargin: '600px 0px' });
  tio = new IntersectionObserver(ens => ens.forEach(en => {
    if (!en.isIntersecting) return; tio.unobserve(en.target); const p = pageById(en.target.dataset.id); if (p) renderThumb(p, en.target);
  }), { root: thumbs, rootMargin: '300px 0px' });

  S.pages.forEach((p, i) => {
    const wrap = document.createElement('div'); wrap.className = 'page'; wrap.dataset.id = p.id;
    const canvas = document.createElement('canvas');
    const s = svg('svg', { class: 'overlay' }), g = svg('g', {}, s);
    const pe = { wrap, canvas, svg: s, g, gFind: svg('g', {}, g), gTI: svg('g', {}, g), gObjs: svg('g', {}, g), gUI: svg('g', {}, g), key: '' };
    wrap.append(canvas, s); viewer.appendChild(wrap); PE.set(p.id, pe);
    s.addEventListener('pointerdown', e => onDown(e, p.id));
    s.addEventListener('dblclick', e => {
      const oe = e.target.closest('[data-id]'); if (!oe || S.editing) return;
      const pg = pageById(p.id), o = pg.objs.find(x => x.id === oe.dataset.id);
      if (o && o.type === 'text') startEdit(pg, o, snap());
      else if (o && o.type === 'image' && S.tool === 'select') { select(pg.id, o.id); ACT.imgedit(); }
    });
    io.observe(wrap);

    const [rw, rh] = rdims(p);
    const t = document.createElement('div'); t.className = 'thumb'; t.draggable = true; t.dataset.id = p.id;
    t.innerHTML = `<div class="tcanvas" style="aspect-ratio:${rw}/${rh}"></div><div class="tnum">${i + 1}</div>
      <div class="tact"><button data-a="rot" title="Rotate"><i data-icon="rotate"></i></button><button data-a="dup" title="Duplicate"><i data-icon="copy"></i></button><button data-a="del" title="Delete" class="danger"><i data-icon="trash"></i></button></div>`;
    thumbs.appendChild(t); tio.observe(t);
  });
  icons(thumbs);
  S.cur = clamp(S.cur, 0, Math.max(0, S.pages.length - 1));
  layout(); S.pages.forEach(drawObjs); drawFind(); updateUI();
}
function layout() {
  const sc = scaleCss();
  for (const p of S.pages) {
    const pe = PE.get(p.id); if (!pe) continue; const [rw, rh] = rdims(p);
    pe.wrap.style.width = rw * sc + 'px'; pe.wrap.style.height = rh * sc + 'px';
    pe.svg.setAttribute('viewBox', `0 0 ${rw} ${rh}`); pe.g.setAttribute('transform', ROT_TF(p));
  }
  drawSel();
}
async function renderCanvas(p) {
  const pe = PE.get(p.id); if (!pe) return;
  const [rw, rh] = rdims(p);
  let scale = scaleCss() * Math.min(window.devicePixelRatio || 1, 2);
  if (rw * rh * scale * scale > 2.4e7) scale = Math.sqrt(2.4e7 / (rw * rh));
  const key = scale + '|' + p.rot; if (pe.key === key) return; pe.key = key;
  if (pe.task) pe.task.cancel();
  const c = document.createElement('canvas'); c.width = Math.round(rw * scale); c.height = Math.round(rh * scale);
  const ctx = c.getContext('2d'); ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, c.width, c.height);
  if (p.src >= 0) {
    const pg = await getPdfPage(p);
    const task = pg.render({ canvasContext: ctx, viewport: pg.getViewport({ scale, rotation: (pg.rotate + p.rot) % 360 }) });
    pe.task = task;
    try { await task.promise; } catch (e) { return; }
    if (pe.task === task) pe.task = null;
  }
  if (PE.get(p.id) !== pe || pe.key !== key) return;
  pe.canvas.replaceWith(c); pe.canvas = c; pe.done = key;
}
async function renderThumb(p, el) {
  const [rw, rh] = rdims(p), scale = 150 / rw * Math.min(window.devicePixelRatio || 1, 2);
  const c = document.createElement('canvas'); c.width = Math.round(rw * scale); c.height = Math.round(rh * scale);
  const ctx = c.getContext('2d'); ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, c.width, c.height);
  if (p.src >= 0) {
    const pg = await getPdfPage(p);
    try { await pg.render({ canvasContext: ctx, viewport: pg.getViewport({ scale, rotation: (pg.rotate + p.rot) % 360 }) }).promise; } catch (e) { return; }
  }
  $('.tcanvas', el).replaceChildren(c);
}
function setZoom(z) {
  z = clamp(z, 0.1, 6); if (!S.pages.length) return;
  const ratio = (viewer.scrollTop + viewer.clientHeight / 2) / viewer.scrollHeight;
  S.zoom = z; layout();
  viewer.scrollTop = ratio * viewer.scrollHeight - viewer.clientHeight / 2;
  clearTimeout(setZoom.t); setZoom.t = setTimeout(() => visible.forEach(id => { const p = pageById(id); if (p) renderCanvas(p); }), 140);
  updateUI();
}
const viewPad = () => innerWidth <= 980 ? 20 : 64;
function fitZoom(page) {
  const p = S.pages[S.cur]; if (!p) return 1; const [rw, rh] = rdims(p), w = (viewer.clientWidth - viewPad()) / (rw * CSS_UNITS);
  return page ? Math.min(w, (viewer.clientHeight - 40) / (rh * CSS_UNITS)) : w;
}
function goToPage(i) {
  i = clamp(i, 0, S.pages.length - 1); const pe = PE.get(S.pages[i] && S.pages[i].id); if (!pe) return;
  viewer.scrollTop = pe.wrap.offsetTop - 12; S.cur = i; updateUI();
}
viewer.addEventListener('scroll', () => { hideCtx(); if (!onScroll.r) onScroll.r = requestAnimationFrame(onScroll); });
function onScroll() {
  onScroll.r = 0; const y = viewer.scrollTop + viewer.clientHeight / 3;
  for (let i = 0; i < S.pages.length; i++) {
    const w = PE.get(S.pages[i].id).wrap;
    if (w.offsetTop + w.offsetHeight > y) { if (S.cur !== i) { S.cur = i; updateUI(); } break; }
  }
}

/* ───────── object geometry ───────── */
const mctx = document.createElement('canvas').getContext('2d'); mctx.fontKerning = 'none'; // PDFs are written without kerning, so measure without it too
const mfont = o => { mctx.font = cssFont(o); mctx.wordSpacing = (o.ws || 0) * o.size + 'px'; }; // o.ws = extra space between words (in em), kept from the original line
const cssFont = o => `${o.italic ? 'italic ' : ''}${o.bold ? 'bold ' : ''}${o.size}px ${FONTS[o.font] || FONTS.Helvetica}`;
const isLine = o => o.type === 'line' || o.type === 'arrow';
const center = b => [b.x + b.w / 2, b.y + b.h / 2];
function rotP([x, y], [cx, cy], deg) { const a = deg * Math.PI / 180, c = Math.cos(a), s = Math.sin(a); return [cx + (x - cx) * c - (y - cy) * s, cy + (x - cx) * s + (y - cy) * c]; }
function bbox(o) {
  switch (o.type) {
    case 'line': case 'arrow': return { x: Math.min(o.x1, o.x2), y: Math.min(o.y1, o.y2), w: Math.abs(o.x2 - o.x1), h: Math.abs(o.y2 - o.y1) };
    case 'path': { let a = 1e9, b = 1e9, c = -1e9, d = -1e9; for (const [x, y] of o.pts) { a = Math.min(a, x); b = Math.min(b, y); c = Math.max(c, x); d = Math.max(d, y); } return { x: a, y: b, w: c - a, h: d - b }; }
    case 'text': { mfont(o); const ls = o.text.split('\n'); return { x: o.x, y: o.y, w: Math.max(4, ...ls.map(l => mctx.measureText(l).width)), h: ls.length * (o.lh || LH) * o.size }; }
    default: return { x: o.x, y: o.y, w: o.w, h: o.h };
  }
}
const rotTf = o => { if (!o.rot || isLine(o)) return ''; const c = center(bbox(o)); return `rotate(${o.rot} ${r2(c[0])} ${r2(c[1])})`; };
function shiftObj(o, src, dx, dy) {
  if (isLine(o)) { o.x1 = src.x1 + dx; o.y1 = src.y1 + dy; o.x2 = src.x2 + dx; o.y2 = src.y2 + dy; }
  else if (o.type === 'path') o.pts = src.pts.map(([x, y]) => [x + dx, y + dy]);
  else { o.x = src.x + dx; o.y = src.y + dy; }
}
function applyBBox(o, src, ob, nb) {
  if (o.type === 'path') o.pts = src.pts.map(([x, y]) => [nb.x + (ob.w ? (x - ob.x) / ob.w * nb.w : 0), nb.y + (ob.h ? (y - ob.y) / ob.h * nb.h : 0)]);
  else if (o.type === 'text') { o.size = Math.max(4, +(src.size * nb.h / ob.h).toFixed(2)); o.x = nb.x; o.y = nb.y; }
  else { o.x = nb.x; o.y = nb.y; o.w = nb.w; o.h = nb.h; }
}
function pathSegs(pts) { // → list of ['M',x,y] | ['Q',cx,cy,x,y] | ['L',x,y]
  const s = [['M', ...pts[0]]];
  if (pts.length < 3) { s.push(['L', ...pts[pts.length - 1]]); return s; }
  for (let i = 1; i < pts.length - 1; i++) s.push(['Q', pts[i][0], pts[i][1], (pts[i][0] + pts[i + 1][0]) / 2, (pts[i][1] + pts[i + 1][1]) / 2]);
  s.push(['L', ...pts[pts.length - 1]]); return s;
}
const pathD = pts => pathSegs(pts).map(s => s[0] + s.slice(1).map(r2).join(' ')).join('');
function arrowGeo(o) {
  const L = Math.max(9, o.width * 4.5), a = Math.atan2(o.y2 - o.y1, o.x2 - o.x1), len = Math.hypot(o.x2 - o.x1, o.y2 - o.y1);
  const k = Math.min(L, len), bx = o.x2 - k * Math.cos(a), by = o.y2 - k * Math.sin(a), wx = -Math.sin(a) * k * 0.42, wy = Math.cos(a) * k * 0.42;
  return { bx, by, head: [[o.x2, o.y2], [bx + wx, by + wy], [bx - wx, by - wy]] };
}
const dashArr = o => o.dash === 'dash' ? [4 * o.width, 3 * o.width] : o.dash === 'dot' ? [0.01, 2 * o.width] : null;
const baseOff = o => (BASE + ((o.lh || LH) - LH) / 2) * o.size; // first baseline below the box top (matches the CSS line box)
const decoY = o => [o.underline ? 0.12 : null, o.strike ? -0.3 : null].filter(v => v !== null); // underline / strike offsets in em
const measure = (text, o) => { mfont(o); return mctx.measureText(text).width; };
function textLines(o) { // per-line layout shared by screen, raster and PDF output
  const b = bbox(o), k = AL[o.align] || 0; mfont(o);
  return o.text.split('\n').map((ln, i) => { const w = mctx.measureText(ln).width; return { ln, w, x: o.x + (b.w - w) * k, y: o.y + baseOff(o) + i * (o.lh || LH) * o.size }; });
}

/* ───────── rendering objects ───────── */
function objEl(o) {
  const g = svg('g', { 'data-id': o.id, opacity: o.opacity ?? 1 }); if (o.blend) g.style.mixBlendMode = 'multiply';
  const tf = rotTf(o); if (tf) g.setAttribute('transform', tf);
  const hasS = o.color && o.color !== 'none' && o.width > 0, hasF = o.fill && o.fill !== 'none';
  const st = { stroke: hasS ? o.color : 'none', 'stroke-width': o.width || 0, 'stroke-linecap': 'round', 'stroke-linejoin': 'round' };
  const da = hasS && dashArr(o); if (da) st['stroke-dasharray'] = da.join(' ');
  const hitW = Math.max((o.width || 0) + 8, COARSE ? 22 : 12);
  const hit = (n, a) => svg(n, { ...a, fill: hasF ? 'transparent' : 'none', stroke: 'transparent', 'stroke-width': hitW, 'pointer-events': hasF ? 'all' : 'stroke' }, g);
  switch (o.type) {
    case 'path': { const d = pathD(o.pts); svg('path', { d, fill: 'none', ...st }, g); hit('path', { d }); break; }
    case 'rect': { const a = { x: o.x, y: o.y, width: o.w, height: o.h }; if (o.rx) a.rx = Math.min(o.rx, o.w / 2, o.h / 2); svg('rect', { ...a, fill: hasF ? o.fill : 'none', ...st, 'stroke-linejoin': 'miter' }, g); hit('rect', a); break; }
    case 'ellipse': { const a = { cx: o.x + o.w / 2, cy: o.y + o.h / 2, rx: o.w / 2, ry: o.h / 2 }; svg('ellipse', { ...a, fill: hasF ? o.fill : 'none', ...st }, g); hit('ellipse', a); break; }
    case 'line': { const a = { x1: o.x1, y1: o.y1, x2: o.x2, y2: o.y2 }; svg('line', { ...a, ...st }, g); hit('line', a); break; }
    case 'arrow': {
      const ar = arrowGeo(o);
      svg('line', { x1: o.x1, y1: o.y1, x2: ar.bx, y2: ar.by, ...st, 'stroke-linecap': 'butt' }, g);
      svg('polygon', { points: ar.head.map(p => p.map(r2).join(',')).join(' '), fill: o.color }, g);
      hit('line', { x1: o.x1, y1: o.y1, x2: o.x2, y2: o.y2 }); break;
    }
    case 'image': svg('image', { href: S.assets[o.asset] || '', x: o.x, y: o.y, width: o.w, height: o.h, preserveAspectRatio: 'none' }, g); break;
    case 'text': {
      ensureFont(o.font);
      const b = bbox(o), bg = o.bg && o.bg !== 'none';
      if (bg) svg('rect', { x: b.x - 3, y: b.y - 2, width: b.w + 6, height: b.h + 4, fill: o.bg, rx: 1.5 }, g);
      const t = svg('text', { fill: o.color, 'font-size': o.size, 'font-family': FONTS[o.font] || FONTS.Helvetica, 'font-weight': o.bold ? 'bold' : 'normal', 'font-style': o.italic ? 'italic' : 'normal', ...(o.ws ? { 'word-spacing': r2(o.ws * o.size) } : {}) }, g);
      t.style.whiteSpace = 'pre';
      for (const l of textLines(o)) {
        svg('tspan', { x: l.x, y: l.y }, t).textContent = l.ln;
        if (l.ln) for (const dy of decoY(o)) svg('line', { x1: l.x, y1: l.y + o.size * dy, x2: l.x + l.w, y2: l.y + o.size * dy, stroke: o.color, 'stroke-width': Math.max(0.5, o.size / 16) }, g);
      }
      svg('rect', { x: b.x - 3, y: b.y - 2, width: b.w + 6, height: b.h + 4, fill: 'transparent' }, g); break;
    }
  }
  return g;
}
function drawObjs(p) {
  const pe = PE.get(p.id); if (!pe) return; pe.gObjs.textContent = '';
  for (const o of p.objs) if (o.id !== S.editing) pe.gObjs.appendChild(objEl(o));
  if (S.sel && S.sel.pid === p.id) drawSel();
}
function drawSel() {
  $$('.selg', viewer).forEach(e => e.remove());
  const s = selObjs(); if (!s) return; const pe = PE.get(s.p.id); if (!pe) return;
  const sc = scaleCss(), r = (COARSE ? 9 : 4.5) / sc, m = 3 / sc, single = s.os.length === 1;
  for (const o of s.os) {
    const g = svg('g', { class: 'selg' }, pe.gUI), tf = rotTf(o); if (tf) g.setAttribute('transform', tf);
    const h = (x, y, n, cur) => svg('circle', { cx: x, cy: y, r, class: 'handle' + (n === 'rot' ? ' rot' : ''), 'data-h': n, style: 'touch-action:none;cursor:' + cur }, g);
    if (isLine(o) && single) { h(o.x1, o.y1, 'p1', 'crosshair'); h(o.x2, o.y2, 'p2', 'crosshair'); continue; }
    const b = bbox(o);
    svg('rect', { x: b.x - m, y: b.y - m, width: b.w + 2 * m, height: b.h + 2 * m, class: 'selbox' }, g);
    if (!single) continue;
    const top = b.y - m - (COARSE ? 34 : 24) / sc;
    svg('line', { x1: b.x + b.w / 2, y1: b.y - m, x2: b.x + b.w / 2, y2: top, class: 'rotline' }, g);
    h(b.x + b.w / 2, top, 'rot', 'grab');
    h(b.x - m, b.y - m, 'nw', 'nwse-resize'); h(b.x + b.w + m, b.y - m, 'ne', 'nesw-resize');
    h(b.x + b.w + m, b.y + b.h + m, 'se', 'nwse-resize'); h(b.x - m, b.y + b.h + m, 'sw', 'nesw-resize');
  }
}
function select(pid, ids) {
  ids = ids ? [].concat(ids) : [];
  S.sel = ids.length ? { pid, ids } : null; drawSel(); syncControls(); updateUI();
}

/* ───────── pointer interaction ───────── */
function toPt(e, pe) { const p = new DOMPoint(e.clientX, e.clientY).matrixTransform(pe.g.getScreenCTM().inverse()); return [p.x, p.y]; }
const shape = (type, P, st) => ({ type, x: P[0], y: P[1], w: 0, h: 0, color: st.color, fill: st.fill, width: st.width, dash: st.dash, opacity: st.opacity, ...(type === 'rect' && st.rx ? { rx: st.rx } : {}) });
const cover = (P, fill, extra) => ({ type: 'rect', x: P[0], y: P[1], w: 0, h: 0, color: 'none', fill, width: 0, opacity: 1, ...extra });
const NEW = {
  pen: (P, st) => ({ type: 'path', pts: [P], color: st.color, width: st.width, opacity: st.opacity }),
  marker: P => ({ type: 'path', pts: [P], color: '#facc15', width: 14, opacity: 0.4, blend: true }),
  hl: P => cover(P, '#facc15', { opacity: 0.4, blend: true }),
  whiteout: P => cover(P, '#ffffff', { cover: true }), redact: P => cover(P, '#000000', { cover: true }),
  rect: (P, st) => shape('rect', P, st), ellipse: (P, st) => shape('ellipse', P, st),
  line: (P, st) => ({ type: 'line', x1: P[0], y1: P[1], x2: P[0], y2: P[1], color: st.color, width: st.width, dash: st.dash, opacity: st.opacity }),
  arrow: (P, st) => ({ type: 'arrow', x1: P[0], y1: P[1], x2: P[0], y2: P[1], color: st.color, width: st.width, dash: st.dash, opacity: st.opacity }),
};
const STICKY = new Set(['pen', 'marker', 'hl', 'whiteout', 'redact']);
let drag = null;
function onDown(e, pid) {
  hideCtx();
  if (e.button !== 0) return;
  if (e.target.closest('.txed')) return;
  if (S.editing) { commitEdit(); e.preventDefault(); return; }
  const p = pageById(pid), pe = PE.get(pid), P = toPt(e, pe), tool = S.tool;
  const oe = e.target.closest('[data-id]'), hit = oe && p.objs.find(x => x.id === oe.dataset.id);
  if (tool === 'hand') drag = { kind: 'pan', cx: e.clientX, cy: e.clientY, st: viewer.scrollTop, sl: viewer.scrollLeft };
  else if (tool === 'select') {
    const h = e.target.closest('[data-h]'), one = selOne();
    if (h && one) drag = { kind: h.dataset.h === 'rot' ? 'rotate' : 'resize', h: h.dataset.h, p: one.p, o: one.o, src: clone(one.o), ob: bbox(one.o) };
    else if (hit) {
      const cur = S.sel && S.sel.pid === pid ? S.sel.ids : [];
      if (e.shiftKey) { select(pid, cur.includes(hit.id) ? cur.filter(i => i !== hit.id) : [...cur, hit.id]); e.preventDefault(); return; }
      if (!cur.includes(hit.id)) select(pid, hit.id);
      const s = selObjs(); drag = { kind: 'move', p, os: s.os, srcs: s.os.map(clone) };
    } else { select(null); if (e.pointerType === 'touch') return; drag = { kind: 'marq', p }; }
  } else if (tool === 'text' || tool === 'note') {
    e.preventDefault();
    if (hit && hit.type === 'text') return startEdit(p, hit, snap());
    const st = S.style, before = snap(), note = tool === 'note', size = note ? 12 : st.size;
    const o = { id: uid(), type: 'text', x: P[0], y: P[1] - size * 0.6, text: '', size, color: note ? '#1f2937' : st.color, font: st.font, bold: !note && st.bold, italic: !note && st.italic,
      underline: !note && st.underline, strike: !note && st.strike, align: note ? 'left' : st.align, lh: note ? LH : st.lh, bg: note ? '#fef08a' : 'none', opacity: note ? 1 : st.opacity };
    p.objs.push(o); return startEdit(p, o, before);
  } else if (tool === 'edittext') {
    if (hit && hit.type === 'text') { e.preventDefault(); return startEdit(p, hit, snap()); }
    const ti = e.target.closest('[data-ti]'); if (ti) { e.preventDefault(); convertTextItem(p, +ti.dataset.ti); }
    return;
  } else if (tool === 'editobj') {
    if (hit) { e.preventDefault(); setTool('select'); select(pid, hit.id); return; }
    const im = e.target.closest('[data-im]'); drag = { kind: 'region', p, im: im ? +im.dataset.im : -1 };
  } else if (NEW[tool]) {
    pushUndo(); const o = { id: uid(), ...NEW[tool](P, S.style) }; p.objs.push(o);
    drag = { kind: 'draw', p, o };
  } else return;
  e.preventDefault();
  drag.pe = pe; drag.start = P;
  window.addEventListener('pointermove', onMove); window.addEventListener('pointerup', onUp); window.addEventListener('pointercancel', onUp);
}
function onMove(e) {
  if (!drag) return;
  if (drag.kind === 'pan') { viewer.scrollTop = drag.st - (e.clientY - drag.cy); viewer.scrollLeft = drag.sl - (e.clientX - drag.cx); return; }
  const { o, src, start } = drag, P = toPt(e, drag.pe), dx = P[0] - start[0], dy = P[1] - start[1];
  drag.last = P;
  if (drag.kind === 'marq' || drag.kind === 'region') {
    if (!drag.el) drag.el = svg('rect', { class: 'marq' }, drag.pe.gUI);
    const a = { x: Math.min(start[0], P[0]), y: Math.min(start[1], P[1]), width: Math.abs(dx), height: Math.abs(dy) };
    for (const k in a) drag.el.setAttribute(k, a[k]); return;
  }
  if (drag.kind !== 'draw' && !drag.moved) { if (Math.hypot(dx, dy) * scaleCss() < 3) return; drag.moved = true; pushUndo(); }
  if (drag.kind === 'move') drag.os.forEach((ob, i) => shiftObj(ob, drag.srcs[i], dx, dy));
  else if (drag.kind === 'rotate') {
    const c = center(drag.ob); let a = Math.atan2(P[1] - c[1], P[0] - c[0]) * 180 / Math.PI + 90;
    if (e.shiftKey) a = Math.round(a / 15) * 15; else if (Math.abs(a % 90) < 3 || Math.abs(a % 90) > 87) a = Math.round(a / 90) * 90; // snap to right angles
    o.rot = Math.round(((a % 360) + 360) % 360 * 10) / 10; $('#pRot').value = Math.round(o.rot);
  } else if (drag.kind === 'resize') {
    const h = drag.h, ob = drag.ob;
    if (h === 'p1') { o.x1 = P[0]; o.y1 = P[1]; } else if (h === 'p2') { o.x2 = P[0]; o.y2 = P[1]; }
    else {
      const rot = src.rot || 0, c0 = center(ob), Q = rot ? rotP(P, c0, -rot) : P; // pointer in the object's own frame
      const W = h.includes('w'), N = h.includes('n'), ax = W ? ob.x + ob.w : ob.x, ay = N ? ob.y + ob.h : ob.y;
      let w = Math.max(4, W ? ax - Q[0] : Q[0] - ax), hh = Math.max(4, N ? ay - Q[1] : Q[1] - ay);
      const keep = o.type === 'text' || (o.type === 'image') !== e.shiftKey;
      if (keep && ob.w && ob.h) { const k = Math.max(w / ob.w, hh / ob.h); w = ob.w * k; hh = ob.h * k; }
      const nb = { x: W ? ax - w : ax, y: N ? ay - hh : ay, w, h: hh };
      if (rot) { const A = [ax, ay], w1 = rotP(A, c0, rot), w2 = rotP(A, center(nb), rot); nb.x += w1[0] - w2[0]; nb.y += w1[1] - w2[1]; } // keep the opposite corner pinned
      applyBBox(o, src, ob, nb);
    }
  } else { // draw
    drag.made = true;
    if (o.type === 'path') { const l = o.pts[o.pts.length - 1]; if (Math.hypot(P[0] - l[0], P[1] - l[1]) > 0.6) o.pts.push([r2(P[0]), r2(P[1])]); }
    else if (isLine(o)) {
      o.x2 = P[0]; o.y2 = P[1];
      if (e.shiftKey) { const a = Math.round(Math.atan2(dy, dx) / (Math.PI / 4)) * Math.PI / 4, l = Math.hypot(dx, dy); o.x2 = start[0] + l * Math.cos(a); o.y2 = start[1] + l * Math.sin(a); }
    } else {
      let w = Math.abs(dx), h = Math.abs(dy); if (e.shiftKey) w = h = Math.max(w, h);
      o.x = dx < 0 ? start[0] - w : start[0]; o.y = dy < 0 ? start[1] - h : start[1]; o.w = w; o.h = h;
    }
  }
  if (!drag.raf) { const d = drag; d.raf = requestAnimationFrame(() => { d.raf = 0; drawObjs(d.p); }); } // at most one redraw per frame while dragging
}
const hitsRect = (b, r) => b.x < r.x + r.w && b.x + b.w > r.x && b.y < r.y + r.h && b.y + b.h > r.y;
function onUp() {
  window.removeEventListener('pointermove', onMove); window.removeEventListener('pointerup', onUp); window.removeEventListener('pointercancel', onUp);
  const d = drag; drag = null; if (!d) return;
  if (d.raf) { cancelAnimationFrame(d.raf); drawObjs(d.p); }
  if (d.el) d.el.remove();
  if (d.kind === 'marq' || d.kind === 'region') {
    const P = d.last || d.start, r = { x: Math.min(d.start[0], P[0]), y: Math.min(d.start[1], P[1]), w: Math.abs(P[0] - d.start[0]), h: Math.abs(P[1] - d.start[1]) };
    const big = Math.max(r.w, r.h) * scaleCss() > 6;
    if (d.kind === 'marq') { if (big) select(d.p.id, d.p.objs.filter(o => hitsRect(bbox(o), r)).map(o => o.id)); }
    else if (big) convertRegion(d.p, r);
    else if (d.im >= 0) getImages(d.p).then(a => a[d.im] && convertRegion(d.p, a[d.im], a[d.im]));
    else toast('Drag a box around the content you want to pick up');
  } else if (d.kind === 'draw') {
    const b = bbox(d.o), tiny = d.o.type === 'path' ? false : Math.max(b.w, b.h) < 3;
    if (tiny) { d.p.objs.pop(); S.undo.pop(); drawObjs(d.p); updateUI(); return; }
    if (!STICKY.has(S.tool)) { setTool('select'); select(d.p.id, d.o.id); }
  } else if (d.moved) syncControls();
}

/* ───────── text editing ───────── */
function styleEditor(o) {
  const d = S.editEl.d;
  const deco = [o.underline && 'underline', o.strike && 'line-through'].filter(Boolean).join(' ') || 'none';
  d.style.cssText = `font:${cssFont(o)};line-height:${o.lh || LH};color:${o.color};opacity:${o.opacity ?? 1};text-align:${o.align || 'left'};text-decoration:${deco};word-spacing:${(o.ws || 0) * o.size}px;background:${o.bg && o.bg !== 'none' ? o.bg : 'transparent'}`;
  S.editEl.box.setAttribute('y', o.y + baseOff(o) - cssBase(o)); // typing happens exactly where the text will be drawn
}
function cssBase(o) { // where the browser puts the first baseline inside the editing box, for this font
  mctx.font = cssFont(o); const m = mctx.measureText('Hg'), a = m.fontBoundingBoxAscent, d = m.fontBoundingBoxDescent;
  return a > 0 ? ((o.lh || LH) * o.size - (a + d)) / 2 + a : baseOff(o);
}
function startEdit(p, o, before, selectAll) {
  commitEdit(); select(null);
  S.editing = o.id; S.editCtx = { pid: p.id, before, tl: o.rot ? rotP([o.x, o.y], center(bbox(o)), o.rot) : null }; drawObjs(p);
  const pe = PE.get(p.id), big = Math.max(p.w, p.h) * 3;
  const wrap = svg('g', { transform: rotTf(o) }, pe.gUI);
  const fo = svg('foreignObject', { x: o.x, y: o.y, width: big, height: big, class: 'txfo' }, wrap);
  const d = document.createElement('div'); d.className = 'txed';
  d.contentEditable = 'plaintext-only'; if (d.contentEditable !== 'plaintext-only') d.contentEditable = 'true';
  d.textContent = o.text; fo.appendChild(d); S.editEl = { fo: wrap, d, box: fo }; styleEditor(o);
  d.addEventListener('blur', e => { if (e.relatedTarget && e.relatedTarget.closest('#props, .menu')) return; setTimeout(commitEdit, 0); });
  d.addEventListener('keydown', e => {
    e.stopPropagation(); const k = e.key.toLowerCase();
    if (e.key === 'Escape') commitEdit();
    else if ((e.ctrlKey || e.metaKey) && { b: 'bold', i: 'italic', u: 'underline' }[k]) { e.preventDefault(); const pk = { b: 'bold', i: 'italic', u: 'underline' }[k]; setProp(pk, !o[pk]); syncControls(); }
  });
  d.focus({ preventScroll: true });
  const sel = getSelection(); sel.selectAllChildren(d); if (!selectAll) sel.collapseToEnd();
  syncControls();
}
function commitEdit() {
  if (!S.editing) return;
  const id = S.editing, { pid, before, tl } = S.editCtx, { fo, d } = S.editEl; S.editing = null;
  const text = d.innerText.replace(/\r/g, '').replace(/\n$/, ''); fo.remove();
  const p = pageById(pid); if (!p) return;
  const o = p.objs.find(x => x.id === id);
  if (o) {
    if (text.trim() === '') p.objs = p.objs.filter(x => x !== o);
    else { o.text = text; if (o.rot && tl) { const w = rotP([o.x, o.y], center(bbox(o)), o.rot); o.x += tl[0] - w[0]; o.y += tl[1] - w[1]; } }
  }
  if (snap() !== before) pushUndo(before);
  drawObjs(p); syncControls(); if (S.tool === 'edittext') drawTI(p);
}

/* ───────── page content: text lines + images ───────── */
const T6 = (a, b) => pdfjsLib.Util.transform(a, b);
function scanPage(p) { // one pass over the page's drawing commands: where its images sit, and which fonts only ever draw invisible text (the hidden OCR layer of a scanned PDF)
  const s = S.sources[p.src]; if (s.scan[p.idx]) return s.scan[p.idx];
  return s.scan[p.idx] = (async () => {
    const imgs = [], all = [], vis = {};
    try {
      const pg = await getPdfPage(p), vt = pg.getViewport({ scale: 1 }).transform, ol = await pg.getOperatorList(), O = pdfjsLib.OPS;
      let ctm = [1, 0, 0, 1, 0, 0], mode = 0, font = ''; const st = [];
      for (let i = 0; i < ol.fnArray.length; i++) {
        const fn = ol.fnArray[i], a = ol.argsArray[i];
        if (fn === O.save) st.push([ctm, mode, font]); else if (fn === O.restore) { if (st.length) [ctm, mode, font] = st.pop(); }
        else if (fn === O.transform) ctm = T6(ctm, a);
        else if (fn === O.paintFormXObjectBegin) { st.push([ctm, mode, font]); if (a && a[0]) ctm = T6(ctm, a[0]); }
        else if (fn === O.paintFormXObjectEnd) { if (st.length) [ctm, mode, font] = st.pop(); }
        else if (fn === O.setFont) font = a[0]; else if (fn === O.setTextRenderingMode) mode = a[0];
        else if (fn === O.showText) vis[font] = vis[font] || !(mode === 3 || mode === 7);
        else if (fn === O.paintImageXObject || fn === O.paintInlineImageXObject || fn === O.paintImageMaskXObject) {
          const m = T6(vt, ctm), xs = [m[4], m[4] + m[0], m[4] + m[2], m[4] + m[0] + m[2]], ys = [m[5], m[5] + m[1], m[5] + m[3], m[5] + m[1] + m[3]];
          const b = { x: Math.max(0, Math.min(...xs)), y: Math.max(0, Math.min(...ys)) }; b.w = Math.min(p.w, Math.max(...xs)) - b.x; b.h = Math.min(p.h, Math.max(...ys)) - b.y;
          if (fn === O.paintImageXObject && typeof a[0] === 'string') { b.id = a[0]; b.m = m; }
          if (b.w > 8 && b.h > 8) { all.push(b); if (b.w * b.h < 0.9 * p.w * p.h) imgs.push(b); }
        }
      }
    } catch (e) { console.warn('page scan failed', e); }
    return { imgs, all, hidden: new Set(Object.keys(vis).filter(k => !vis[k])) };
  })();
}
const getImages = async p => p.src < 0 ? [] : (await scanPage(p)).imgs; // bounding boxes of the images painted on the original page (view space)
function readText(p) { // → { segs: runs set in one font (what line mode edits), lines: whole lines (paragraphs, find, export) }
  const s = S.sources[p.src]; if (s.text[p.idx]) return s.text[p.idx];
  return s.text[p.idx] = (async () => {
    const pg = await getPdfPage(p), vt = pg.getViewport({ scale: 1 }).transform, [tc, sc] = await Promise.all([pg.getTextContent(), scanPage(p)]), rows = [], segs = [], lines = [];
    const box = it => { it.y = it.base - it.size * 0.85; it.h = it.size * 1.12; return it; };
    for (const it of tc.items) {
      if (!it.str || !it.str.trim() || !it.width) continue;
      const m = T6(vt, it.transform), size = Math.hypot(m[2], m[3]); if (size < 2) continue;
      const stl = tc.styles[it.fontName] || {}, fam = stl.fontFamily || ''; if (stl.vertical) continue;
      const f = { str: it.str, x: m[4], w: it.width, base: m[5], size, fontName: it.fontName, hidden: sc.hidden.has(it.fontName), font: /mono/i.test(fam) ? 'Courier' : /sans/i.test(fam) ? 'Helvetica' : /serif/i.test(fam) ? 'Times' : 'Helvetica' };
      const ang = Math.atan2(m[1], m[0]);
      if (Math.abs(ang) > 0.02) { // text set at an angle is edited piece by piece, as a box turned around its own centre
        const cs = Math.cos(ang), sn = Math.sin(ang), vx = f.w / 2, vy = -0.29 * size, cx = f.x + vx * cs - vy * sn, cy = f.base + vx * sn + vy * cs;
        const r = { ...f, o: [f.x, f.base], rot: Math.round(ang * 1800 / Math.PI) / 10, x: cx - f.w / 2, y: cy - 0.56 * size, h: 1.12 * size }; r.base = r.y + 0.85 * size;
        segs.push(r); lines.push(r); continue;
      }
      let row = null; // group fragments that share a baseline into one row
      for (let i = rows.length - 1; i >= 0 && i >= rows.length - 40; i--) { const r = rows[i]; if (Math.abs(r.base - f.base) < 0.3 * Math.min(r.size, size) && r.size / size > 0.8 && r.size / size < 1.25) { row = r; break; } }
      if (row) row.frags.push(f); else rows.push({ base: f.base, size, frags: [f] });
    }
    const join = fr => { // neighbouring fragments → one piece of text
      const o = { ...fr[0] };
      for (const f of fr.slice(1)) { if (f.x - (o.x + o.w) > f.size * 0.15 && !/\s$/.test(o.str) && !/^\s/.test(f.str)) o.str += ' '; o.str += f.str; o.w = f.x + f.w - o.x; }
      return o;
    };
    for (const row of rows) { // merge neighbours into editable lines; big gaps (columns, tables) stay separate
      row.frags.sort((a, b) => a.x - b.x); const groups = []; let cur = null, end = 0;
      for (const f of row.frags) { const gap = f.x - end; if (cur && gap < f.size * 1.2 && gap > -f.size * 0.6) cur.push(f); else groups.push(cur = [f]); end = f.x + f.w; }
      for (const g of groups) {
        const line = join(g), n = {}; let top = g[0];
        for (const f of g) { n[f.fontName] = (n[f.fontName] || 0) + f.str.length; if (n[f.fontName] > n[top.fontName]) top = f; }
        Object.assign(line, { size: top.size, fontName: top.fontName, font: top.font, base: top.base, hidden: g.every(f => f.hidden) }); lines.push(box(line));
        let run = [g[0]]; // a change of typeface inside the line (a bold word, say) starts a new piece, so each piece keeps its own font
        for (const f of g.slice(1)) { if (f.fontName === run[0].fontName && Math.abs(f.size - run[0].size) < 0.5) run.push(f); else { segs.push(box(join(run))); run = [f]; } }
        segs.push(run.length === g.length ? line : box(join(run)));
      }
    }
    return { segs, lines };
  })().catch(e => { delete s.text[p.idx]; throw e; });
}
const getText = async p => p.src < 0 ? [] : (await readText(p)).segs;
const getLines = async p => p.src < 0 ? [] : (await readText(p)).lines;
const ocrOf = (p, force) => window.XD3.ocrItems ? window.XD3.ocrItems(p, force) : Promise.resolve([]);
let tiSeq = 0;
async function drawTI(p) {
  const pe = PE.get(p.id); if (!pe) return; const tool = S.tool, para = S.para, tok = pe.tiTok = ++tiSeq;
  pe.g.insertBefore(pe.gTI, tool === 'edittext' ? pe.gUI : pe.gObjs); // text outlines sit above placed pictures (an opened image, a scan), so text inside them can be seen and clicked
  if (tool !== 'edittext' && tool !== 'editobj') { pe.gTI.textContent = ''; pe.items = null; return; }
  const stale = () => S.tool !== tool || S.para !== para || pe.tiTok !== tok || PE.get(p.id) !== pe;
  if (tool === 'editobj') {
    const items = await getImages(p); if (stale()) return; pe.gTI.textContent = ''; pe.items = null;
    items.forEach((b, i) => svg('rect', { x: b.x, y: b.y, width: b.w, height: b.h, class: 'ti img', 'data-im': i }, pe.gTI)); return;
  }
  const paint = items => {
    pe.items = items = items.filter(it => !covered(p, it)); pe.gTI.textContent = '';
    items.forEach((it, i) => { const a = { x: it.x - 1, y: it.y, width: it.w + 2, height: it.h, class: it.ocr ? 'ti ocr' : 'ti', 'data-ti': i }; if (it.rot) a.transform = `rotate(${it.rot} ${r2(it.x + it.w / 2)} ${r2(it.y + it.h / 2)})`; svg('rect', a, pe.gTI); });
  };
  const nat = para ? await getBlocks(p) : await getText(p); if (stale()) return; paint(nat);
  const oc = await ocrOf(p).catch(() => []); if (stale() || !oc.length) return; // text inside pictures and scans arrives a little later
  paint([...nat, ...(para ? toBlocks(oc) : oc)]);
}
function refreshTI() { PE.forEach(pe => { pe.gTI.textContent = ''; pe.items = null; }); visible.forEach(id => { const p = pageById(id); if (p) drawTI(p); }); }
const hex = (r, g, b) => '#' + [r, g, b].map(v => v.toString(16).padStart(2, '0')).join('');
function modeColor(ctx, strips) { // most common colour across pixel strips [x,y,w,h]
  const cnt = {}; let best = 0, col = [255, 255, 255];
  for (const [x, y, w, h] of strips) {
    const X = clamp(Math.round(x), 0, ctx.canvas.width - 1), Y = clamp(Math.round(y), 0, ctx.canvas.height - 1);
    const W = clamp(Math.round(w), 1, ctx.canvas.width - X), H = clamp(Math.round(h), 1, ctx.canvas.height - Y), d = ctx.getImageData(X, Y, W, H).data;
    for (let i = 0; i < d.length; i += 4) { const k = d[i] << 16 | d[i + 1] << 8 | d[i + 2], n = cnt[k] = (cnt[k] || 0) + 1; if (n > best) { best = n; col = [d[i], d[i + 1], d[i + 2]]; } }
  }
  return col;
}
// A render of what lies under the edits — the original page plus any pictures placed on it — to sample colours from, rebuild backgrounds and read text (OCR).
const isBaseImg = o => o.type === 'image' && !o.cover && (o.opacity ?? 1) > 0.6;
const baseSig = p => p.id + '|' + p.objs.filter(isBaseImg).map(o => [o.asset, r2(o.x), r2(o.y), r2(o.w), r2(o.h), o.rot || 0].join(',')).join(';');
const cvCache = [];
async function baseCanvas(p, minK = 1.2) { // → { c: canvas (unrotated page), k: pixels per point }; the last few are kept, so repeated edits on a page are instant
  const imgs = p.objs.filter(isBaseImg), pe = PE.get(p.id);
  if (!imgs.length && p.src >= 0 && pe && pe.key && pe.done === pe.key && !p.rot && pe.canvas.width / p.w >= minK) return { c: pe.canvas, k: pe.canvas.width / p.w };
  const sig = baseSig(p); let e = cvCache.find(x => x.sig === sig && x.k >= minK - 0.01);
  if (!e) {
    const k = Math.min(Math.max(minK, 2), Math.sqrt(1.6e7 / (p.w * p.h)));
    e = { sig, k, c: await renderPageImage({ ...p, rot: 0, objs: imgs }, k) }; cvCache.unshift(e); cvCache.length = Math.min(cvCache.length, 3);
  }
  return e;
}
const aabb = it => { // upright bounds of a (possibly rotated) box
  if (!it.rot) return it; const c = [it.x + it.w / 2, it.y + it.h / 2], P = [[it.x, it.y], [it.x + it.w, it.y], [it.x, it.y + it.h], [it.x + it.w, it.y + it.h]].map(q => rotP(q, c, it.rot));
  const x = Math.min(...P.map(q => q[0])), y = Math.min(...P.map(q => q[1])); return { x, y, w: Math.max(...P.map(q => q[0])) - x, h: Math.max(...P.map(q => q[1])) - y };
};
function sampleColors(p, it, c) { // background + ink colour from a rendered canvas
  const out = { bg: '#ffffff', fg: '#000000' };
  try {
    if (!c) return out; const f = c.width / p.w, ctx = c.getContext('2d', { willReadFrequently: true }); it = aabb(it);
    const x = Math.max(0, Math.floor((it.x - 1) * f)), y = Math.max(0, Math.floor(it.y * f));
    const w = Math.min(c.width - x, Math.ceil((it.w + 2) * f)), h = Math.min(c.height - y, Math.ceil(it.h * f));
    if (w < 2 || h < 2) return out;
    const [br, bg, bb] = modeColor(ctx, [[x, y, w, 1], [x, y + h - 1, w, 1]]), d = ctx.getImageData(x, y, w, h).data;
    let md = 0, fr = 0, fg = 0, fb = 0;
    for (let i = 0; i < d.length; i += 4) { const dd = Math.abs(d[i] - br) + Math.abs(d[i + 1] - bg) + Math.abs(d[i + 2] - bb); if (dd > md) { md = dd; fr = d[i]; fg = d[i + 1]; fb = d[i + 2]; } }
    out.bg = hex(br, bg, bb); if (md > 90) out.fg = hex(fr, fg, fb);
  } catch (e) { /* keep defaults */ }
  return out;
}
function coverFor(p, r, cv, whole) { // hide part of the original page: a flat box where the background is plain, otherwise a patch rebuilt from the pixels around it (paper grain, gradients and photos survive)
  const flat = fill => ({ id: uid(), type: 'rect', cover: true, x: r.x, y: r.y, w: r.w, h: r.h, color: 'none', fill, width: 0, opacity: 1, ...(r.rot ? { rot: r.rot } : {}) });
  try {
    if (r.rot) return flat(sampleColors(p, r, cv.c).bg);
    const k = cv.k, c = cv.c, g = Math.max(2, Math.round(k)), X0 = clamp(Math.floor(r.x * k), g, c.width - g - 2), Y0 = clamp(Math.floor(r.y * k), g, c.height - g - 2);
    const W = clamp(Math.ceil(r.w * k), 2, c.width - g - X0), H = clamp(Math.ceil(r.h * k), 2, c.height - g - Y0), w2 = W + 2 * g, h2 = H + 2 * g;
    const d = c.getContext('2d', { willReadFrequently: true }).getImageData(X0 - g, Y0 - g, w2, h2).data, at = (x, y) => (y * w2 + x) * 4;
    const ring = []; for (let x = 0; x < w2; x++) ring.push(at(x, 0), at(x, h2 - 1)); for (let y = 1; y < h2 - 1; y++) ring.push(at(0, y), at(w2 - 1, y));
    const cnt = {}; let best = 0, bk = 0; for (const i of ring) { const q = (d[i] >> 3) << 10 | (d[i + 1] >> 3) << 5 | d[i + 2] >> 3, n = cnt[q] = (cnt[q] || 0) + 1; if (n > best) { best = n; bk = q; } }
    let m = [0, 0, 0], mn = 0, near = 0; const B = [(bk >> 10) * 8 + 4, (bk >> 5 & 31) * 8 + 4, (bk & 31) * 8 + 4];
    for (const i of ring) { const dd = Math.abs(d[i] - B[0]) + Math.abs(d[i + 1] - B[1]) + Math.abs(d[i + 2] - B[2]); if (dd < 22) near++; if (dd < 14) { m[0] += d[i]; m[1] += d[i + 1]; m[2] += d[i + 2]; mn++; } }
    m = mn ? m.map(v => Math.round(v / mn)) : B;
    if (near / ring.length > 0.9) return flat(hex(...m)); // plain background → a crisp vector box
    if (!whole && k < 1.9) return null; // a patch needs a sharper render than this one — the caller fetches it
    const med = (get, n) => { // edge colours, with stray ink (a neighbouring line's descender, a rule) filtered out
      const out = [], rad = Math.max(3, Math.round(k * 3));
      for (let i = 0; i < n; i++) { const px = []; for (let j = Math.max(0, i - rad); j <= Math.min(n - 1, i + rad); j++) px.push(get(j)); px.sort((a, b) => (a[0] + a[1] + a[2]) - (b[0] + b[1] + b[2])); out.push(px[px.length >> 1]); }
      return out;
    };
    const px = i => [d[i], d[i + 1], d[i + 2]], Tp = med(x => px(at(x + g, 0)), W), Bt = med(x => px(at(x + g, h2 - 1)), W), Lf = whole && med(y => px(at(0, y + g)), H), Rt = whole && med(y => px(at(w2 - 1, y + g)), H);
    const o = document.createElement('canvas'); o.width = W; o.height = H; const ox = o.getContext('2d'), od = ox.createImageData(W, H), e = od.data, ink = new Uint8Array(W * H), fillc = new Float32Array(W * H * 3);
    for (let y = 0; y < H; y++) for (let x = 0; x < W; x++) {
      const ty = (y + 0.5) / H, tx = (x + 0.5) / W, j = (y * W + x) * 3, i = at(x + g, y + g); let dd = 0;
      for (let q = 0; q < 3; q++) { let v = Tp[x][q] * (1 - ty) + Bt[x][q] * ty; if (whole) v = (v + Lf[y][q] * (1 - tx) + Rt[y][q] * tx) / 2; fillc[j + q] = v; dd += Math.abs(d[i + q] - v); }
      if (whole || dd > 42) ink[y * W + x] = 1;
    }
    const sp = whole ? 0 : Math.max(1, Math.round(k * 0.6)); // also repaint the soft halo around each glyph
    for (let y = 0; y < H; y++) for (let x = 0; x < W; x++) {
      let on = ink[y * W + x]; for (let yy = Math.max(0, y - sp); !on && yy <= Math.min(H - 1, y + sp); yy++) for (let xx = Math.max(0, x - sp); xx <= Math.min(W - 1, x + sp); xx++) if (ink[yy * W + xx]) { on = 1; break; }
      const i = (y * W + x) * 4, s = at(x + g, y + g), j = (y * W + x) * 3;
      for (let q = 0; q < 3; q++) e[i + q] = on ? fillc[j + q] : d[s + q]; e[i + 3] = 255;
    }
    ox.putImageData(od, 0, 0);
    return { id: uid(), type: 'image', cover: true, asset: addAsset(o.toDataURL('image/jpeg', 0.92)), x: X0 / k, y: Y0 / k, w: W / k, h: H / k, opacity: 1 };
  } catch (e) { console.warn('cover', e); return flat('#ffffff'); }
}
const pdfFont = (pg, name) => new Promise(res => { // pdf.js' record of a font, once it has finished loading it
  const t = setTimeout(() => res(null), 2500); try { pg.commonObjs.get(name, f => { clearTimeout(t); res(f); }); } catch (e) { clearTimeout(t); res(null); }
});
async function origFont(p, it, light) { // the typeface an original line is set in → { font, bold, italic, real, family }; `light` only describes it, without taking a copy of the font
  if (!it.fontName) return { font: it.font || 'Helvetica', bold: false, italic: false, real: '', family: null };
  const s = S.sources[p.src], key = it.fontName + (light ? '|l' : ''); if (s.fonts[key]) return s.fonts[key];
  let f = null; try { const pg = await getPdfPage(p); if (!pg.commonObjs.has(it.fontName)) await scanPage(p); f = await pdfFont(pg, it.fontName); } catch (e) { /* described by its fallback below */ }
  const name = (f && f.name) || '', fam = fontFamilyFor(name), bold = !!f && (!!f.black || /bold|black|heavy|semibold|demi/i.test(name)), italic = /ital|obli/i.test(name) || !!(f && f.italic);
  const fb = fam || (/courier|mono|consol|typewriter/i.test(name) ? 'Courier' : /times|georgia|cambria|garamond|palatino|minion|bookman|century|serif|nimbusrom|cmr\d|cmbx|cmti/i.test(name) && !/sans/i.test(name) ? 'Times' : /arial|helvet|verdana|tahoma|calibri|segoe|sans/i.test(name) ? 'Helvetica' : it.font || 'Helvetica');
  const out = { font: fb, bold, italic, real: name ? cleanFontName(name) : '', family: fam, type3: !!(f && f.isType3Font) };
  if (!light && f && f.data && f.data.length > 200 && !f.isType3Font) { // embedded in the PDF → keep using that very font
    const bytes = new Uint8Array(f.data), id = xfId(name, bytes);
    if (!S.xfonts[id]) S.xfonts[id] = { name: out.real, ps: name, bytes, bold, italic, fb };
    await useXFont(id); out.font = id;
  }
  return s.fonts[key] = out;
}
async function fontStyle(p, fontName) { const f = await origFont(p, { fontName }, true); return { bold: f.bold, italic: f.italic, family: f.family, real: f.real }; } // bold / italic / family of an embedded font
async function editOriginal(p, it, str) { // cover an original line / paragraph and put an editable copy on top, in the same font
  const L = it.lines || [it], X = window.XD3, fx = it.ocr || it.hidden ? null : await origFont(p, it), visual = (!fx || fx.type3) && X.fitLine && !it.rot;
  const cv = await baseCanvas(p, visual ? 3 : 1.2), text = str == null ? it.str : str; let o = null, hi = null;
  if (visual) { // printed / scanned text has no font to reuse: measure the ink and pick the closest typeface, size and colour
    const long = L.reduce((a, b) => b.str.length > a.str.length ? b : a), f = await X.fitLine(cv, long), f1 = f && (L[0] === long ? f : await X.fitLine(cv, L[0], f));
    if (f && f1) {
      o = textObj(text, f.size, { color: f.color, font: f.font, bold: f.bold, italic: f.italic, lh: L.length > 1 ? clamp((L[L.length - 1].base - L[0].base) / (L.length - 1) / f.size, 0.9, 3) : LH });
      await ensureFont(o.font); o.x = f1.x; o.y = f1.base - baseOff(o);
    }
  }
  if (!o) {
    const col = sampleColors(p, L[0], cv.c), F = fx || { font: it.font || 'Helvetica', bold: false, italic: false };
    o = textObj(text, +it.size.toFixed(2), { color: col.fg, font: F.font, bold: F.bold, italic: F.italic, lh: it.lh || LH, align: it.align || 'left' });
    await ensureFont(o.font);
    { // keep the original spacing between words (justified lines; fonts whose space is not a glyph)
      const ws = L.map(l => { const t = l.str.trim(), n = t.split(' ').length - 1; return n ? (l.w - measure(t, o)) / n / o.size : null; }).filter(v => v != null);
      if (ws.length) { const v = clamp(Math.min(...ws), -0.12, 0.8); if (Math.abs(v) > 0.003) o.ws = +v.toFixed(4); }
    }
    if (it.rot) { // keep the start of the baseline where it was
      const b = bbox(o), a = it.rot * Math.PI / 180, cs = Math.cos(a), sn = Math.sin(a), vx = -b.w / 2, vy = baseOff(o) - b.h / 2, cx = it.o[0] - (vx * cs - vy * sn), cy = it.o[1] - (vx * sn + vy * cs);
      o.rot = it.rot; o.x = cx - b.w / 2; o.y = cy - b.h / 2;
    } else { o.x = it.align === 'center' ? it.x + (it.w - bbox(o).w) / 2 : it.x; o.y = it.base - baseOff(o); }
  }
  for (const l of L) { const r = { x: l.x - 1.5, y: l.y - 0.5, w: l.w + 3, h: l.h + 1, rot: l.rot }; p.objs.push(coverFor(p, r, cv) || coverFor(p, r, hi || (hi = await baseCanvas(p, 2)))); }
  p.objs.push(o); return o;
}
async function convertTextItem(p, i) {
  const pe = PE.get(p.id), it = (pe && pe.items ? pe.items : S.para ? await getBlocks(p) : await getText(p))[i]; if (!it) return;
  const before = snap(), o = await editOriginal(p, it); startEdit(p, o, before, true);
}
function toBlocks(lines) { // lines grouped into paragraphs for paragraph-mode editing
  const items = lines.filter(l => !l.rot).sort((a, b) => a.base - b.base || a.x - b.x), blocks = lines.filter(l => l.rot);
  for (const it of items) {
    let hit = null;
    for (let i = blocks.length - 1; i >= 0 && i >= blocks.length - 12; i--) {
      if (!blocks[i].lines) continue;
      const L = blocks[i].lines, l = L[L.length - 1], gap = it.base - l.base, sz = Math.min(l.size, it.size);
      const sameCol = Math.abs(it.x - L[0].x) < sz * 2.5 || Math.abs(it.x + it.w / 2 - (l.x + l.w / 2)) < sz * 1.5;
      if (Math.abs(l.size - it.size) < (it.ocr ? sz * 0.22 : 0.6) && gap > sz * 0.9 && gap < sz * 1.75 && sameCol && it.x < l.x + l.w && l.x < it.x + it.w && l.fontName === it.fontName) { hit = blocks[i]; break; }
    }
    if (hit) hit.lines.push(it); else blocks.push({ lines: [it] });
  }
  for (const b of blocks) {
    if (!b.lines) continue;
    const L = b.lines, f = L[0], last = L[L.length - 1], x = Math.min(...L.map(l => l.x)), r = Math.max(...L.map(l => l.x + l.w));
    Object.assign(b, { x, w: r - x, y: f.y, h: last.y + last.h - f.y, base: f.base, size: f.size, font: f.font, fontName: f.fontName, ocr: f.ocr, hidden: f.hidden, str: L.map(l => l.str).join('\n') });
    if (L.length > 1) {
      b.lh = clamp((last.base - f.base) / (L.length - 1) / f.size, 0.9, 3);
      if (L.some(l => Math.abs(l.x - x) > f.size * 0.5) && L.every(l => Math.abs(l.x + l.w / 2 - (x + b.w / 2)) < f.size)) b.align = 'center';
    }
  }
  return blocks;
}
async function getBlocks(p) {
  if (p.src < 0) return [];
  const s = S.sources[p.src]; return s.blocks[p.idx] || (s.blocks[p.idx] = toBlocks(await getLines(p)));
}
const covered = (p, it) => { const cx = it.x + it.w / 2, cy = it.y + it.h / 2; return p.objs.some(o => { if (!o.cover) return false; const b = bbox(o); return cx > b.x && cx < b.x + b.w && cy > b.y && cy < b.y + b.h; }); };
async function visibleText(p) { // every line a reader would see: uncovered original text (and text already recognised in pictures) + text added in the editor
  const X = window.XD3, orig = [...await getLines(p), ...(X.ocrCached ? X.ocrCached(p) : [])];
  const out = orig.map((it, i) => ({ ...it, src: 'orig', i })).filter(it => !covered(p, it));
  for (const o of p.objs) if (o.type === 'text') textLines(o).forEach((l, li) => {
    if (l.ln.trim()) out.push({ str: l.ln, x: l.x, y: l.y - 0.85 * o.size, w: l.w, h: o.size * 1.12, base: l.y, size: o.size, font: o.font, bold: o.bold, italic: o.italic, underline: o.underline, strike: o.strike, color: o.color, src: 'obj', id: o.id, li });
  });
  return out;
}
async function nativeImage(p, b) { // the picture itself, at its own resolution and with its transparency — when it sits upright and nothing is drawn over it
  try {
    if (!b.id || !b.m || Math.abs(b.m[1]) > 1e-3 * Math.abs(b.m[0]) || Math.abs(b.m[2]) > 1e-3 * Math.abs(b.m[3]) || b.m[0] <= 0 || b.m[3] >= 0) return null;
    const sc = await scanPage(p), over = q => q !== b && hitsRect(q, { x: b.x + 1, y: b.y + 1, w: b.w - 2, h: b.h - 2 });
    if (sc.all.slice(sc.all.indexOf(b) + 1).some(over) || (await getLines(p)).some(l => hitsRect(aabb(l), b))) return null;
    const pg = await getPdfPage(p), objs = b.id.startsWith('g_') ? pg.commonObjs : pg.objs; if (!objs.has(b.id)) return null;
    const im = objs.get(b.id); if (!im) return null; let src = im.bitmap;
    if (!src) {
      if (!im.data || !im.width) return null; const n = im.width * im.height, rgba = new Uint8ClampedArray(n * 4), s = im.data;
      if (s.length === n * 4) rgba.set(s); else if (s.length === n * 3) for (let i = 0, j = 0; i < n; i++, j += 3) { rgba[i * 4] = s[j]; rgba[i * 4 + 1] = s[j + 1]; rgba[i * 4 + 2] = s[j + 2]; rgba[i * 4 + 3] = 255; } else return null;
      src = document.createElement('canvas'); src.width = im.width; src.height = im.height; src.getContext('2d').putImageData(new ImageData(rgba, im.width, im.height), 0, 0);
    }
    const f = Math.min(1, 3200 / Math.max(src.width, src.height)), c = document.createElement('canvas'); c.width = Math.max(1, Math.round(src.width * f)); c.height = Math.max(1, Math.round(src.height * f));
    c.getContext('2d').drawImage(src, 0, 0, c.width, c.height);
    const t = document.createElement('canvas'); t.width = t.height = 48; const tx = t.getContext('2d', { willReadFrequently: true }); tx.drawImage(c, 0, 0, 48, 48);
    const td = tx.getImageData(0, 0, 48, 48).data; let alpha = false; for (let i = 3; i < td.length; i += 4) if (td[i] < 250) { alpha = true; break; }
    return alpha || c.width * c.height < 3e5 ? c.toDataURL('image/png') : c.toDataURL('image/jpeg', 0.92);
  } catch (e) { return null; }
}
function snapRegion(r, cv) { // shrink a dragged box onto the content inside it, when it sits on a plain background
  try {
    const k = cv.k, c = cv.c, X = clamp(Math.round(r.x * k), 0, c.width - 2), Y = clamp(Math.round(r.y * k), 0, c.height - 2), W = clamp(Math.round(r.w * k), 2, c.width - X), H = clamp(Math.round(r.h * k), 2, c.height - Y);
    const d = c.getContext('2d', { willReadFrequently: true }).getImageData(X, Y, W, H).data, B = [d[0], d[1], d[2]], dev = i => Math.abs(d[i] - B[0]) + Math.abs(d[i + 1] - B[1]) + Math.abs(d[i + 2] - B[2]);
    let edge = 0, n = 0;
    for (let x = 0; x < W; x++) { n += 2; if (dev(x * 4) < 36) edge++; if (dev(((H - 1) * W + x) * 4) < 36) edge++; }
    for (let y = 0; y < H; y++) { n += 2; if (dev(y * W * 4) < 36) edge++; if (dev((y * W + W - 1) * 4) < 36) edge++; }
    if (edge / n < 0.97) return r; // the box cuts through content — take it exactly as drawn
    let x0 = W, y0 = H, x1 = -1, y1 = -1;
    for (let y = 0; y < H; y++) for (let x = 0; x < W; x++) if (dev((y * W + x) * 4) > 48) { if (x < x0) x0 = x; if (x > x1) x1 = x; if (y < y0) y0 = y; if (y > y1) y1 = y; }
    if (x1 < 0) return null; const m = 1.5 * k;
    x0 = Math.max(0, x0 - m); y0 = Math.max(0, y0 - m); x1 = Math.min(W, x1 + 1 + m); y1 = Math.min(H, y1 + 1 + m);
    return { x: (X + x0) / k, y: (Y + y0) / k, w: (x1 - x0) / k, h: (y1 - y0) / k };
  } catch (e) { return r; }
}
async function convertRegion(p, r, im) { // lift a region of the original page (or one of its images, `im`) into a movable image object
  if (p.src < 0) return toast('This page has no original content — use the Select tool for your own objects');
  const x = clamp(r.x, 0, p.w), y = clamp(r.y, 0, p.h); r = { x, y, w: Math.min(r.x + r.w, p.w) - x, h: Math.min(r.y + r.h, p.h) - y };
  if (r.w < 3 || r.h < 3) return;
  await busy(async () => {
    const cv = await baseCanvas({ ...p, objs: [] }, 3), k = cv.k; let url = im ? await nativeImage(p, im) : null;
    if (!im) { r = snapRegion(r, cv); if (!r) return toast('Nothing to pick up there — that area is empty'); }
    if (!url) {
      const o = document.createElement('canvas'); o.width = Math.max(1, Math.round(r.w * k)); o.height = Math.max(1, Math.round(r.h * k));
      o.getContext('2d').drawImage(cv.c, Math.round(r.x * k), Math.round(r.y * k), o.width, o.height, 0, 0, o.width, o.height); url = o.toDataURL('image/png');
    }
    pushUndo();
    const img = { id: uid(), type: 'image', asset: addAsset(url), x: r.x, y: r.y, w: r.w, h: r.h, opacity: 1 };
    p.objs.push(coverFor(p, { x: r.x - 0.8, y: r.y - 0.8, w: r.w + 1.6, h: r.h + 1.6 }, cv, true), img);
    drawObjs(p); setTool('select'); select(p.id, img.id);
    toast('Picked up — drag to move, handles to resize or rotate, Del to remove');
  }, 'Picking up object…');
}

/* ───────── find ───────── */
async function runFind(q, keep) {
  S.find = { q, hits: [], cur: -1 };
  if (q) {
    const lq = q.toLowerCase();
    for (let i = 0; i < S.pages.length; i++) {
      const p = S.pages[i], items = await visibleText(p); if (S.find.q !== q) return;
      for (const it of items) {
        const ls = it.str.toLowerCase(); let k = -1;
        while ((k = ls.indexOf(lq, k + 1)) >= 0) S.find.hits.push({ pi: i, pid: p.id, x: it.x + it.w * k / ls.length, y: it.y, w: it.w * lq.length / ls.length, h: it.h, it, k });
      }
    }
    if (S.find.hits.length) S.find.cur = clamp(keep || 0, 0, S.find.hits.length - 1);
  }
  drawFind(); if (keep == null) showHit();
}
const reEsc = s => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
async function replaceIn(p, it, q, rep, k) { // k = char index of one match, or null for every match in the line
  const ns = k == null ? it.str.replace(new RegExp(reEsc(q), 'gi'), () => rep) : it.str.slice(0, k) + rep + it.str.slice(k + q.length);
  if (it.src === 'obj') { const o = p.objs.find(x => x.id === it.id); if (!o) return; const ls = o.text.split('\n'); ls[it.li] = ns; o.text = ls.join('\n'); }
  else await editOriginal(p, it, ns);
}
async function replaceOne() {
  const h = S.find.hits[S.find.cur], q = S.find.q; if (!h) return toast('Nothing to replace — type a search term first');
  const p = pageById(h.pid); pushUndo(); await replaceIn(p, h.it, q, $('#replInput').value, h.k);
  drawObjs(p); await runFind(q, S.find.cur); showHit();
}
async function replaceAll() {
  const q = S.find.q, rep = $('#replInput').value; if (!S.find.hits.length) return toast('Nothing to replace — type a search term first');
  await busy(async () => {
    pushUndo(); let n = 0;
    for (const p of S.pages) {
      const items = (await visibleText(p)).filter(it => it.str.toLowerCase().includes(q.toLowerCase())); if (!items.length) continue;
      for (const it of items) { n += it.str.toLowerCase().split(q.toLowerCase()).length - 1; await replaceIn(p, it, q, rep, null); }
      drawObjs(p);
    }
    await runFind(q, 0); toast(`Replaced ${n} match(es)`);
  }, 'Replacing…');
}
function drawFind() {
  PE.forEach(pe => pe.gFind.textContent = '');
  S.find.hits.forEach((h, i) => { const pe = PE.get(h.pid); if (pe) svg('rect', { x: h.x, y: h.y, width: h.w, height: h.h, class: 'hit' + (i === S.find.cur ? ' cur' : '') }, pe.gFind); });
  $('#findCnt').textContent = `${S.find.cur + 1}/${S.find.hits.length}`;
}
function showHit() {
  const h = S.find.hits[S.find.cur]; if (!h) return; const pe = PE.get(h.pid); if (!pe) return;
  viewer.scrollTop = pe.wrap.offsetTop + h.y * scaleCss() - viewer.clientHeight / 3;
}
function stepFind(d) { const n = S.find.hits.length; if (!n) return; S.find.cur = (S.find.cur + d + n) % n; drawFind(); showHit(); }
function toggleFind(show) {
  const fb = $('#findbar'); fb.hidden = show === true ? false : !fb.hidden;
  if (fb.hidden) runFind(''); else { $('#findInput').focus(); $('#findInput').select(); if ($('#findInput').value) runFind($('#findInput').value); }
}
$('#replInput').addEventListener('keydown', e => { e.stopPropagation(); if (e.key === 'Enter') replaceOne(); if (e.key === 'Escape') toggleFind(); });
$('#findInput').addEventListener('input', e => { clearTimeout(runFind.t); runFind.t = setTimeout(() => runFind(e.target.value.trim()), 220); });
$('#findInput').addEventListener('keydown', e => { e.stopPropagation(); if (e.key === 'Enter') stepFind(e.shiftKey ? -1 : 1); if (e.key === 'Escape') toggleFind(); });

/* ───────── tools & properties ───────── */
function setTool(t) {
  commitEdit(); S.tool = t; document.body.dataset.mode = t;
  $$('button[data-tool]').forEach(b => b.classList.toggle('active', b.dataset.tool === t));
  const pane = ($(`button[data-tool="${t}"]`) || document.body).closest('.pane'); if (pane && !pane.classList.contains('active')) setTab(pane.dataset.pane); // show the tab that holds this tool
  if (t !== 'select') select(null);
  $('#hint').textContent = HINTS[t] || ''; flash($('#hint'), 'swap'); refreshTI(); syncControls();
}
function applyProp(o, k, v) {
  if (k === 'color') { if (o.type === 'image') return; if (o.color === 'none' && o.fill && o.fill !== 'none') o.fill = v; else o.color = v; }
  else if (k === 'fill') { if (o.type === 'text') o.bg = v; else if (o.type === 'rect' || o.type === 'ellipse') o.fill = v; }
  else if (k === 'width') { if (o.type !== 'text' && o.type !== 'image' && o.color !== 'none') o.width = v; }
  else if (k === 'dash') { if (o.type !== 'text' && o.type !== 'image') o.dash = v; }
  else if (k === 'rx') { if (o.type === 'rect') o.rx = v; }
  else if (k === 'opacity') o.opacity = v;
  else if (k === 'rot') { if (!isLine(o)) o.rot = v; }
  else if (o.type === 'text') o[k] = v; // size, font, bold, italic, underline, align
}
let armed = false;
function setProp(k, v) {
  const ed = editingObj();
  if (ed) { applyProp(ed, k, v); styleEditor(ed); return; } // live restyle while typing; undo is recorded on commit
  const s = selObjs(); if (!s) { if (k !== 'rot') S.style[k] = v; return; }
  if (!armed) { pushUndo(); armed = true; }
  s.os.forEach(o => applyProp(o, k, v)); drawObjs(s.p);
}
function syncControls() {
  const ed = editingObj(), s = selObjs(), one = s && s.os.length === 1 ? s.os[0] : null, o = ed || one || (s ? s.os[0] : S.style);
  const c = o.color && o.color !== 'none' ? o.color : (o.fill && o.fill !== 'none' ? o.fill : S.style.color);
  $('#pColor').value = c;
  const fv = o.type === 'text' ? o.bg : o.fill, f = fv && fv !== 'none'; $('#pFillOn').checked = !!f; if (f) $('#pFill').value = fv;
  if (o.width) $('#pWidth').value = o.width;
  $('#pDash').value = o.dash || 'solid'; $('#pRx').value = o.rx || 0;
  if (o.size) $('#pSize').value = Math.round(o.size * 10) / 10;
  if (o.font) { const fb = $('#pFontBtn span'); fb.textContent = fontLabel(o.font); fb.style.fontFamily = FONTS[o.font] || ''; }
  $('#pBold').classList.toggle('active', !!o.bold); $('#pItalic').classList.toggle('active', !!o.italic); $('#pUnderline').classList.toggle('active', !!o.underline);
  $('#pStrike').classList.toggle('active', !!o.strike); $('#pPara').classList.toggle('active', S.para);
  $('#pLh').value = String([1, 1.2, 1.5, 2].reduce((a, b) => Math.abs(b - (o.lh || LH)) < Math.abs(a - (o.lh || LH)) ? b : a));
  const al = $('#pAlign'); al.dataset.v = o.align || 'left'; al.innerHTML = `<i data-icon="align${al.dataset.v}"></i>`; icons(al);
  $('#pOpacity').value = o.opacity ?? 1;
  $('#pRot').value = Math.round((ed || one || {}).rot || 0); $('.selonly-wrap').classList.toggle('off', !one || isLine(one));
  const src = ed || one;
  $('#props').dataset.ctx = src ? (src.type === 'text' ? 'text' : src.type === 'image' ? 'image' : 'shape') : s ? 'all' : ['text', 'edittext', 'note'].includes(S.tool) ? 'text' : 'shape';
}
const bindProp = (sel, k, get) => { const el = $(sel); el.addEventListener('input', () => setProp(k, get(el))); el.addEventListener('change', () => { armed = false; }); el.addEventListener('keydown', e => e.stopPropagation()); };
bindProp('#pColor', 'color', el => el.value);
bindProp('#pFill', 'fill', el => { $('#pFillOn').checked = true; return el.value; });
bindProp('#pFillOn', 'fill', el => el.checked ? $('#pFill').value : 'none');
bindProp('#pWidth', 'width', el => clamp(+el.value || 1, 0.25, 60));
bindProp('#pDash', 'dash', el => el.value);
bindProp('#pRx', 'rx', el => clamp(+el.value || 0, 0, 500));
bindProp('#pLh', 'lh', el => +el.value);
$('#pPara').addEventListener('click', () => { S.para = !S.para; $('#pPara').classList.toggle('active', S.para); $('#hint').textContent = S.para ? 'Paragraph mode — click a block of text to rewrite the whole paragraph' : HINTS.edittext; refreshTI(); });
bindProp('#pSize', 'size', el => clamp(+el.value || 12, 4, 300));
bindProp('#pOpacity', 'opacity', el => +el.value);
bindProp('#pRot', 'rot', el => ((+el.value || 0) % 360 + 360) % 360);
const refocus = () => { if (S.editing) S.editEl.d.focus({ preventScroll: true }); };
for (const [sel, k] of [['#pBold', 'bold'], ['#pItalic', 'italic'], ['#pUnderline', 'underline'], ['#pStrike', 'strike']]) $(sel).addEventListener('click', e => {
  const on = !e.currentTarget.classList.contains('active'); e.currentTarget.classList.toggle('active', on); setProp(k, on); armed = false; refocus();
});
$('#pAlign').addEventListener('click', e => {
  const v = { left: 'center', center: 'right', right: 'left' }[e.currentTarget.dataset.v || 'left']; setProp('align', v); armed = false; syncControls(); refocus();
});

function delSel() { const s = selObjs(); if (!s) return; pushUndo(); s.p.objs = s.p.objs.filter(x => !s.os.includes(x)); S.sel = null; drawObjs(s.p); drawSel(); syncControls(); updateUI(); }
function addCopies(os, p) { const ns = os.map(o => { const n = clone(o); n.id = uid(); shiftObj(n, o, 12, 12); return n; }); p.objs.push(...ns); drawObjs(p); select(p.id, ns.map(n => n.id)); return ns; }
function dupSel() { const s = selObjs(); if (!s) return; pushUndo(); addCopies(s.os, s.p); }
function copySel() { const s = selObjs(); if (s) { S.clip = s.os.map(clone); toast(`Copied ${s.os.length} object(s)`); } }
function pasteClip(at) { // at = { pid, P } pastes with the top-left corner at that point; otherwise cascades from the copied position
  const p = at ? pageById(at.pid) : S.pages[S.cur]; if (!S.clip || !p) return; pushUndo(); setTool('select');
  const ns = addCopies(S.clip, p);
  if (!at) { S.clip = ns.map(clone); return; }
  const bs = ns.map(bbox), x = Math.min(...bs.map(b => b.x)), y = Math.min(...bs.map(b => b.y));
  ns.forEach(o => shiftObj(o, clone(o), at.P[0] - x, at.P[1] - y)); drawObjs(p);
}
function reorder(front) { const s = selObjs(); if (!s) return; pushUndo(); const a = s.p.objs.filter(x => !s.os.includes(x)); s.p.objs = front ? [...a, ...s.os] : [...s.os, ...a]; drawObjs(s.p); }
function alignSel(m) { // l c r t m b — to the selection's bounds, or to the page when one object is selected
  const s = selObjs(); if (!s) return; pushUndo();
  const bs = s.os.map(bbox); let R = { x: 0, y: 0, w: s.p.w, h: s.p.h };
  if (s.os.length > 1) { const x1 = Math.min(...bs.map(b => b.x)), y1 = Math.min(...bs.map(b => b.y)); R = { x: x1, y: y1, w: Math.max(...bs.map(b => b.x + b.w)) - x1, h: Math.max(...bs.map(b => b.y + b.h)) - y1 }; }
  s.os.forEach((o, i) => {
    const b = bs[i], dx = m === 'l' ? R.x - b.x : m === 'c' ? R.x + R.w / 2 - b.x - b.w / 2 : m === 'r' ? R.x + R.w - b.x - b.w : 0;
    const dy = m === 't' ? R.y - b.y : m === 'm' ? R.y + R.h / 2 - b.y - b.h / 2 : m === 'b' ? R.y + R.h - b.y - b.h : 0;
    shiftObj(o, clone(o), dx, dy);
  });
  drawObjs(s.p);
}
function distributeSel(axis) {
  const s = selObjs(); if (!s || s.os.length < 3) return toast('Select three or more objects to distribute');
  pushUndo(); const k = axis === 'h' ? 'x' : 'y', d = axis === 'h' ? 'w' : 'h';
  const L = s.os.map(o => ({ o, b: bbox(o) })).sort((a, b) => a.b[k] - b.b[k]), first = L[0].b, last = L[L.length - 1].b;
  const gap = (last[k] + last[d] - first[k] - L.reduce((a, e) => a + e.b[d], 0)) / (L.length - 1); let pos = first[k];
  for (const e of L) { const dv = pos - e.b[k]; shiftObj(e.o, clone(e.o), axis === 'h' ? dv : 0, axis === 'h' ? 0 : dv); pos += e.b[d] + gap; }
  drawObjs(s.p);
}
const STYLE_KEYS = ['color', 'fill', 'width', 'dash', 'rx', 'opacity', 'size', 'font', 'bold', 'italic', 'underline', 'strike', 'align', 'lh', 'bg'];
function copyStyle() { const s = selOne(); if (!s) return toast('Select one object to copy its style'); S.styleClip = {}; for (const k of STYLE_KEYS) if (k in s.o) S.styleClip[k] = s.o[k]; toast('Style copied — select other objects and choose Paste style'); }
function pasteStyle() {
  const s = selObjs(); if (!s || !S.styleClip) return toast('Copy a style first'); pushUndo();
  const fits = (o, k) => o.type === 'image' ? k === 'opacity' : o.type === 'text' ? !['fill', 'width', 'dash', 'rx'].includes(k)
    : k === 'rx' ? o.type === 'rect' : k === 'fill' ? o.type === 'rect' || o.type === 'ellipse' : ['color', 'width', 'dash', 'opacity'].includes(k);
  for (const o of s.os) for (const k in S.styleClip) if (fits(o, k)) o[k] = S.styleClip[k];
  drawObjs(s.p); syncControls();
}
function nudge(dx, dy) { const s = selObjs(); if (!s) return; if (!armed) { pushUndo(); armed = true; } clearTimeout(nudge.t); nudge.t = setTimeout(() => armed = false, 600); s.os.forEach(o => shiftObj(o, clone(o), dx, dy)); drawObjs(s.p); }

/* ───────── right-click menus ───────── */
// A menu item is '-' (separator) or { label, icon, act, d: {data-*}, key, danger, disabled, checked, keep, sub: [items] }.
// `keep` = the action must not end an in-progress text edit (clipboard and style commands).
const mi = (label, icon, act, o) => ({ label, icon, act, ...o });
const ctxSubs = []; let subStore = [], subTimer = 0;
function menuHtml(items) {
  const L = items.filter(Boolean).filter((it, i, a) => it !== '-' || (i > 0 && a[i - 1] !== '-'));
  while (L[L.length - 1] === '-') L.pop();
  return L.map(it => {
    if (it === '-') return '<hr>';
    const cls = [it.sub && 'hassub', it.danger && 'danger', it.checked && 'on'].filter(Boolean).join(' ');
    const data = it.sub ? ` data-sub="${subStore.push(it.sub) - 1}"` : ` data-act="${it.act}"` + Object.entries(it.d || {}).map(([k, v]) => ` data-${k}="${esc(v)}"`).join('') + (it.keep ? ' data-keep' : '');
    return `<button${cls ? ` class="${cls}"` : ''}${data}${it.disabled ? ' disabled' : ''}><i data-icon="${it.checked && !it.icon ? 'check' : it.icon || 'none'}"></i><span>${esc(it.label)}</span>${it.key ? `<kbd>${it.key}</kbd>` : ''}${it.sub ? '<b class="arr">›</b>' : ''}</button>`;
  }).join('');
}
function placeMenu(m, x, y, flipX) { // keep the menu on screen; a submenu flips to the left of its parent when there is no room
  m.style.left = Math.max(4, x + m.offsetWidth > innerWidth - 6 ? (flipX != null ? flipX - m.offsetWidth : innerWidth - m.offsetWidth - 6) : x) + 'px';
  m.style.top = Math.max(4, Math.min(y, innerHeight - m.offsetHeight - 6)) + 'px';
}
function closeSubs(depth) { while (ctxSubs.length > depth) ctxSubs.pop().remove(); }
function hideCtx() { clearTimeout(subTimer); closeSubs(0); ctxMenu.classList.remove('open'); if (ctxMenu.parentNode !== document.body) document.body.appendChild(ctxMenu); }
function showCtx(e, items) {
  hideCtx(); $$('.menu.open').forEach(m => m.classList.remove('open'));
  const host = (e.target.closest && e.target.closest('dialog[open]')) || document.body; host.appendChild(ctxMenu); // inside an open dialog the menu must live in it to appear on top
  subStore = []; ctxMenu.innerHTML = menuHtml(items); ctxMenu.dataset.depth = 0; icons(ctxMenu); ctxMenu.classList.add('open'); placeMenu(ctxMenu, e.clientX, e.clientY);
}
function openSub(btn) {
  const parent = btn.parentNode, depth = +parent.dataset.depth || 0;
  closeSubs(depth); $$('.subopen', parent).forEach(b => b.classList.remove('subopen')); btn.classList.add('subopen');
  const m = document.createElement('div'); m.className = 'menu fx open ctxsub'; m.dataset.depth = depth + 1; m.innerHTML = menuHtml(subStore[btn.dataset.sub]); icons(m);
  ctxMenu.parentNode.appendChild(m); const r = btn.getBoundingClientRect(); placeMenu(m, r.right - 3, r.top - 6, r.left + 3); ctxSubs.push(m); return m;
}
document.addEventListener('mouseover', e => { // submenus open on hover
  const b = e.target.closest && e.target.closest('#ctx button, .ctxsub button'); if (!b) return;
  clearTimeout(subTimer); subTimer = setTimeout(() => {
    if (!b.isConnected) return; const depth = +b.parentNode.dataset.depth || 0;
    if (b.classList.contains('hassub')) { if (!b.classList.contains('subopen')) openSub(b); }
    else { closeSubs(depth); $$('.subopen', b.parentNode).forEach(x => x.classList.remove('subopen')); }
  }, 130);
});
document.addEventListener('mousedown', e => { if (e.target.closest && e.target.closest('.menu.fx button')) e.preventDefault(); }, true); // menus never steal focus from the text being edited
window.addEventListener('keydown', e => { // arrow keys / Enter / Esc inside an open right-click menu
  if (!ctxMenu.classList.contains('open')) return;
  const m = ctxSubs[ctxSubs.length - 1] || ctxMenu, btns = $$('button:not([disabled])', m), cur = $('.kb', m), i = btns.indexOf(cur), k = e.key;
  const mark = (mm, b) => { $$('.kb', mm).forEach(x => x.classList.remove('kb')); if (b) { b.classList.add('kb'); b.scrollIntoView({ block: 'nearest' }); } };
  if (k === 'Escape') { if (ctxSubs.length) closeSubs(ctxSubs.length - 1); else hideCtx(); }
  else if (k === 'ArrowDown') mark(m, btns[(i + 1) % btns.length]);
  else if (k === 'ArrowUp') mark(m, btns[(i - 1 + btns.length) % btns.length]);
  else if (k === 'ArrowRight') { if (cur && cur.classList.contains('hassub')) { const s = openSub(cur); mark(s, $('button:not([disabled])', s)); } }
  else if (k === 'ArrowLeft') { if (ctxSubs.length) closeSubs(ctxSubs.length - 1); }
  else if (k === 'Enter') { if (cur) cur.click(); }
  else { hideCtx(); return; }
  e.preventDefault(); e.stopImmediatePropagation();
}, true);

const STAMPS = [['APPROVED', '#16a34a'], ['REJECTED', '#dc2626'], ['CONFIDENTIAL', '#dc2626'], ['DRAFT', '#64748b'], ['PAID', '#2563eb'], ['COPY', '#64748b'], ['SIGN HERE', '#d97706']];
const TOOL_LIST = [['select', 'Select / move', 'cursor', 'V'], ['hand', 'Pan', 'hand'], ['edittext', 'Edit text', 'edittext', 'E'], ['editobj', 'Edit object', 'editobj', 'B'], ['text', 'Add text', 'text', 'T'], ['note', 'Sticky note', 'note', 'N'],
  ['pen', 'Pen', 'pen', 'P'], ['marker', 'Highlighter', 'marker', 'M'], ['hl', 'Highlight area', 'hl', 'H'], ['rect', 'Rectangle', 'rect', 'R'], ['ellipse', 'Ellipse', 'ellipse', 'O'], ['line', 'Line', 'line', 'L'], ['arrow', 'Arrow', 'arrow', 'A'],
  ['whiteout', 'Whiteout', 'eraser', 'W'], ['redact', 'Blackout', 'redact']];
const toolsSub = () => TOOL_LIST.map(([t, l, ic, key]) => mi(l, ic, 'tool', { d: { t }, key, checked: S.tool === t }));
const insertSub = () => [mi('Image…', 'image', 'image'), mi('Signature…', 'sign', 'sign'), mi('Today’s date', 'cal', 'datestamp'),
  { label: 'Stamp', icon: 'stamp', sub: [...STAMPS.map(([t, c]) => mi(t[0] + t.slice(1).toLowerCase(), 'stamp', 'stamp', { d: { t, c } })), mi('Check mark', 'check', 'stamp', { d: { t: '✓', c: '#16a34a', plain: 1 } }), mi('Cross mark', 'x', 'stamp', { d: { t: '✗', c: '#dc2626', plain: 1 } })] },
  '-', mi('Watermark…', 'drop', 'watermark'), mi('Page numbers…', 'hash', 'pagenum'), mi('Header / footer…', 'stroke', 'headfoot'), '-',
  mi('Scan document…', 'scan', 'scan'), mi('Add files (PDF, Word, images)…', 'merge', 'merge'), mi('Blank page', 'plus', 'blank')];
const pageSub = () => { const n = S.pages.length, p = S.pages[S.cur]; return [
  mi('Rotate right', 'rotate', 'rotr'), mi('Rotate left', 'rotatel', 'rotl'), mi('Duplicate', 'copy', 'dup'), mi('Insert blank page after', 'plus', 'blank'), mi('Replace with another file…', 'swap', 'replacepage'), '-',
  mi('Move up', 'up', 'pgup', { disabled: S.cur === 0 }), mi('Move down', 'down', 'pgdown', { disabled: S.cur >= n - 1 }), mi('Move to start', 'front', 'pgtop', { disabled: S.cur === 0 }), mi('Move to end', 'back', 'pgend', { disabled: S.cur >= n - 1 }), '-',
  mi('Extract…', 'split', 'extract'), mi('Change size…', 'fit', 'resize'), mi('Add page numbers…', 'hash', 'pagenum'), mi('Export as PNG', 'image', 'png'), '-',
  mi('Remove all edits on page', 'eraser', 'clearpage', { disabled: !p || !p.objs.length }), mi('Delete page', 'trash', 'delpage', { danger: 1, disabled: n < 2 })]; };
const zoomSub = () => [mi('Zoom in', 'zoomin', 'zoomin', { key: 'Ctrl +' }), mi('Zoom out', 'zoomout', 'zoomout', { key: 'Ctrl −' }), mi('Actual size (100%)', '', 'zoomreset', { key: 'Ctrl+0' }), mi('Fit width', 'fit', 'fit'), mi('Fit whole page', 'pages', 'fitpage'), '-',
  mi('Show / hide page panel', 'panel', 'sidebar'), mi('Light / dark theme', 'moon', 'theme')];
const convertSub = () => [mi('PDF → Word (.docx)…', 'word', 'pdf2docx'), mi('PDF → Images…', 'image', 'pdf2img'), mi('Export text (.txt)', 'text', 'exporttxt'), '-',
  mi('Compress…', 'compress', 'compress'), mi('Split into single pages', 'split', 'splitzip'), '-', mi('Password protect…', 'lock', 'protect'), mi('Flatten (permanent redaction)…', 'layers', 'flatten')];
const editItems = () => [mi('Undo', 'undo', 'undo', { key: 'Ctrl+Z', disabled: !S.undo.length }), mi('Redo', 'redo', 'redo', { key: 'Ctrl+Y', disabled: !S.redo.length })];
const clipMenu = () => [mi('Cut', 'cut', 'ed_cut', { key: 'Ctrl+X', keep: 1 }), mi('Copy', 'copy', 'ed_copy', { key: 'Ctrl+C', keep: 1 }), mi('Paste', 'pages', 'ed_paste', { key: 'Ctrl+V', keep: 1 }), mi('Select all', 'selall', 'ed_selall', { key: 'Ctrl+A', keep: 1 })];
const textStyleSub = (o, keep) => [
  mi('Bold', '', 'tstyle', { d: { k: 'bold' }, key: 'Ctrl+B', checked: !!o.bold, keep }), mi('Italic', '', 'tstyle', { d: { k: 'italic' }, key: 'Ctrl+I', checked: !!o.italic, keep }),
  mi('Underline', '', 'tstyle', { d: { k: 'underline' }, key: 'Ctrl+U', checked: !!o.underline, keep }), mi('Strikethrough', '', 'tstyle', { d: { k: 'strike' }, checked: !!o.strike, keep }), '-',
  ...['left', 'center', 'right'].map(v => mi('Align ' + (v === 'center' ? 'centre' : v), 'align' + v, 'talign', { d: { v }, checked: (o.align || 'left') === v, keep })), '-',
  mi('Bigger', 'plus', 'tsize', { d: { v: 2 }, keep }), mi('Smaller', 'minus', 'tsize', { d: { v: -2 }, keep })];
function pageMenu(here) {
  return [mi('Paste', 'pages', here ? 'pastehere' : 'paste', { key: 'Ctrl+V', disabled: !S.clip }),
    mi('Add text' + (here ? ' here' : ''), 'text', 'texthere'), mi('Add sticky note' + (here ? ' here' : ''), 'note', 'notehere'), { label: 'Insert', icon: 'plus', sub: insertSub() }, '-',
    { label: 'Tools', icon: 'cursor', sub: toolsSub() }, mi('Select all on page', 'selall', 'selall', { key: 'Ctrl+A' }), mi('Find & replace…', 'search', 'findrep', { key: 'Ctrl+H' }), '-',
    ...editItems(), '-',
    { label: 'This page', icon: 'pages', sub: pageSub() }, mi('Organize pages…', 'grid', 'organize', { key: 'G' }), { label: 'Zoom & view', icon: 'zoomin', sub: zoomSub() }, { label: 'Convert & protect', icon: 'word', sub: convertSub() }, '-',
    mi('Print…', 'print', 'print', { key: 'Ctrl+P' }), mi('Save PDF', 'download', 'save', { key: 'Ctrl+S' })];
}
function objectMenu(s) {
  const os = s.os, one = os.length === 1 ? os[0] : null, t = one && one.type, tx = os.find(o => o.type === 'text'), sh = os.find(o => o.type !== 'text' && o.type !== 'image');
  const op = Math.round(((one || os[0]).opacity ?? 1) * 100), A = (l, ic, m) => mi(l, ic, 'align', { d: { m } });
  return [
    t === 'text' && mi('Edit text', 'edittext', 'edittxt', { key: 'Enter' }),
    t === 'image' && mi('Edit image (filters, crop)…', 'sliders', 'imgedit'), t === 'image' && mi('Remove background…', 'wand', 'bgremove'),
    t === 'image' && mi('Replace image…', 'swap', 'imgreplace'), t === 'image' && mi('Save image as…', 'download', 'imgsave'), (t === 'text' || t === 'image') && '-',
    mi('Cut', 'cut', 'cut', { key: 'Ctrl+X' }), mi('Copy', 'copy', 'copy', { key: 'Ctrl+C' }), mi('Paste', 'pages', 'paste', { key: 'Ctrl+V', disabled: !S.clip }), mi('Duplicate', 'copy', 'osdup', { key: 'Ctrl+D' }),
    mi('Delete', 'trash', 'del', { key: 'Del', danger: 1 }), '-',
    tx && { label: 'Text style', icon: 'text', sub: textStyleSub(tx) },
    sh && { label: 'Line style', icon: 'stroke', sub: [['solid', 'Solid'], ['dash', 'Dashed'], ['dot', 'Dotted']].map(([v, l]) => mi(l, '', 'sdash', { d: { v }, checked: (sh.dash || 'solid') === v })) },
    { label: 'Opacity', icon: 'drop', sub: [100, 75, 50, 25].map(v => mi(v + '%', '', 'opac', { d: { v: v / 100 }, checked: op === v })) },
    os.some(o => !isLine(o)) && { label: 'Rotate', icon: 'rotate', sub: [mi('90° right', 'rotate', 'orot', { d: { v: 90 } }), mi('90° left', 'rotatel', 'orot', { d: { v: -90 } }), mi('180°', '', 'orot', { d: { v: 180 } }), mi('Reset rotation', '', 'orot', { d: { v: 0 } })] },
    { label: 'Arrange', icon: 'layers', sub: [mi('Bring to front', 'front', 'front'), mi('Bring forward', 'up', 'fwd'), mi('Send backward', 'down', 'bwd'), mi('Send to back', 'back', 'back')] },
    { label: one ? 'Align on page' : 'Align', icon: 'al_c', sub: [A('Left', 'al_l', 'l'), A('Centre', 'al_c', 'c'), A('Right', 'al_r', 'r'), A('Top', 'al_t', 't'), A('Middle', 'al_m', 'm'), A('Bottom', 'al_b', 'b'), '-',
      mi('Distribute horizontally', 'dist_h', 'distribute', { d: { m: 'h' }, disabled: os.length < 3 }), mi('Distribute vertically', 'dist_v', 'distribute', { d: { m: 'v' }, disabled: os.length < 3 })] }, '-',
    mi('Copy style', 'brush', 'copystyle', { disabled: !one }), mi('Paste style', 'brush', 'pastestyle', { disabled: !S.styleClip }),
    S.pages.length > 1 && mi('Copy to all pages', 'pages', 'toallpages'), '-',
    mi('Select all on page', 'selall', 'selall', { key: 'Ctrl+A' }),
  ];
}
const appMenu = () => S.pages.length ? [
  mi('Open…', 'folder', 'open', { key: 'Ctrl+O' }), mi('Add files…', 'merge', 'merge'), mi('Save PDF', 'download', 'save', { key: 'Ctrl+S' }), mi('Print…', 'print', 'print', { key: 'Ctrl+P' }), '-', ...editItems(), '-',
  { label: 'Insert', icon: 'plus', sub: insertSub() }, { label: 'Tools', icon: 'cursor', sub: toolsSub() }, { label: 'This page', icon: 'pages', sub: pageSub() }, mi('Organize pages…', 'grid', 'organize', { key: 'G' }),
  { label: 'Zoom & view', icon: 'zoomin', sub: zoomSub() }, { label: 'Convert & protect', icon: 'word', sub: convertSub() }, '-',
  mi('Document properties…', 'info', 'props'), mi('Keyboard shortcuts', 'key', 'shortcuts'), mi('About', 'info', 'about'),
] : [
  mi('Open file…', 'folder', 'open', { key: 'Ctrl+O' }), mi('Merge files…', 'merge', 'mergedlg'), mi('Scan document…', 'scan', 'scan'), mi('Word → PDF…', 'word', 'docx2pdf'), mi('New blank PDF', 'plus', 'newdoc'),
  !$('#bRestore').hidden && mi('Restore last session', 'undo', 'restore'), '-', mi('Light / dark theme', 'moon', 'theme'), mi('Keyboard shortcuts', 'key', 'shortcuts'), mi('About', 'info', 'about'),
];
function onContext(e, pid) { // right-click on a page: the object under the pointer, or the page itself
  commitEdit();
  const p = pageById(pid), oe = e.target.closest('[data-id]'), hit = oe && p.objs.find(x => x.id === oe.dataset.id);
  if (hit) { if (S.tool !== 'select') setTool('select'); if (!(S.sel && S.sel.pid === pid && S.sel.ids.includes(hit.id))) select(pid, hit.id); }
  else select(null);
  S.cur = S.pages.indexOf(p); S.ctxPt = { pid, P: toPt(e, PE.get(pid)) };
  const s = selObjs(); showCtx(e, s ? objectMenu(s) : pageMenu(true)); updateUI();
}
document.addEventListener('contextmenu', e => { // our own menu everywhere — the browser's menu is never shown
  e.preventDefault(); const t = e.target; if (!t.closest || t.closest('#ctx, .ctxsub')) return;
  const field = t.closest('input, textarea');
  if (field) { if (/^(text|search|number|password|email|url|)$/.test(field.getAttribute('type') || '')) { field.focus(); showCtx(e, clipMenu()); } return; }
  if (t.closest('.txed')) return showCtx(e, [...clipMenu(), '-', { label: 'Text style', icon: 'text', sub: textStyleSub(editingObj() || {}, 1) }, '-', mi('Done editing', 'check', 'ed_done', { key: 'Esc' })]);
  if (t.closest('dialog')) return;
  if (t.closest('#org')) { if (window.XD3.orgCtx) window.XD3.orgCtx(e); return; }
  const pg = t.closest('.page'); if (pg) return onContext(e, pg.dataset.id);
  commitEdit(); S.ctxPt = null;
  const th = t.closest('.thumb');
  if (th) { goToPage(S.pages.findIndex(p => p.id === th.dataset.id)); return showCtx(e, [...pageSub(), '-', mi('Organize pages…', 'grid', 'organize', { key: 'G' })]); }
  showCtx(e, t.closest('#viewer') && S.pages.length ? pageMenu(false) : appMenu());
});

// quick actions used by the menus
function stepOrder(dir) { // one layer up / down
  const s = selObjs(); if (!s) return; pushUndo(); const a = s.p.objs, ids = new Set(s.os.map(o => o.id));
  if (dir > 0) { for (let i = a.length - 2; i >= 0; i--) if (ids.has(a[i].id) && !ids.has(a[i + 1].id)) [a[i], a[i + 1]] = [a[i + 1], a[i]]; }
  else for (let i = 1; i < a.length; i++) if (ids.has(a[i].id) && !ids.has(a[i - 1].id)) [a[i], a[i - 1]] = [a[i - 1], a[i]];
  drawObjs(s.p);
}
function quickProp(k, v) { setProp(k, v); armed = false; syncControls(); refocus(); } // a one-shot property change (its own undo step)
const textTarget = () => { const ed = editingObj(); if (ed) return ed; const s = selObjs(); return (s && s.os.find(o => o.type === 'text')) || null; };
function rotateSel(v) { const s = selObjs(); if (!s) return; pushUndo(); s.os.forEach(o => { if (!isLine(o)) o.rot = v ? (((o.rot || 0) + v) % 360 + 360) % 360 : 0; }); drawObjs(s.p); syncControls(); }
function toAllPages() {
  const s = selObjs(); if (!s) return toast('Select the object(s) to copy first'); if (S.pages.length < 2) return toast('This document has only one page');
  pushUndo(); for (const p of S.pages) { if (p === s.p) continue; for (const o of s.os) { const c = clone(o); c.id = uid(); p.objs.push(c); } drawObjs(p); }
  toast(`Copied to ${S.pages.length - 1} other page(s)`);
}
function imgReplace() {
  const s = selOne(); if (!s || s.o.type !== 'image') return; const i = document.createElement('input'); i.type = 'file'; i.accept = 'image/*';
  i.onchange = () => i.files[0] && busy(async () => { const im = await readImage(i.files[0]), o = s.o, cy = o.y + o.h / 2; pushUndo(); o.asset = addAsset(im.url); o.h = o.w * im.h / im.w; o.y = cy - o.h / 2; drawObjs(s.p); });
  i.click();
}
function imgSave() { const s = selOne(); if (!s || s.o.type !== 'image') return; const url = S.assets[s.o.asset]; fetch(url).then(r => r.blob()).then(b => download(b, `${baseName()}-image.${/^data:image\/jpe?g/.test(url) ? 'jpg' : 'png'}`)); }
function edPaste() { // paste into the focused text field / text box
  const fail = () => toast('Press Ctrl+V to paste');
  if (navigator.clipboard && navigator.clipboard.readText) navigator.clipboard.readText().then(t => { if (t) document.execCommand('insertText', false, t); }).catch(fail); else fail();
}
function setTab(t) { $$('.tab').forEach(b => b.classList.toggle('active', b.dataset.tab === t)); $$('.pane').forEach(p => p.classList.toggle('active', p.dataset.pane === t)); }
$('#tabs').addEventListener('click', e => { const b = e.target.closest('.tab'); if (b) setTab(b.dataset.tab); });
function buildFontMenu() {
  const doc = Object.keys(S.xfonts), groups = { ...(doc.length ? { 'From this document': doc } : {}), Standard: ['Helvetica', 'Times', 'Courier'] }; for (const n in BUNDLED) (groups[BUNDLED[n][3]] = groups[BUNDLED[n][3]] || []).push(n);
  doc.forEach(useXFont);
  $('#mFonts').innerHTML = Object.entries(groups).map(([g, list]) => `<div class="mhead">${g}</div>` + list.map(n => `<button data-act="setfont" data-f="${n}" data-keep style='font-family:${FONTS[n]}'>${esc(fontLabel(n))}</button>`).join('')).join('');
}
function newTextAt(note) { // from the right-click menu: add text / a note where the user clicked (or mid-view)
  const p = S.ctxPt ? pageById(S.ctxPt.pid) : S.pages[S.cur]; if (!p) return;
  const P = S.ctxPt ? S.ctxPt.P : viewCenter(p), st = S.style, before = snap(), size = note ? 12 : st.size;
  setTool(note ? 'note' : 'text');
  const o = textObj('', size, { x: P[0], y: P[1] - size * 0.6, color: note ? '#1f2937' : st.color, font: st.font, bg: note ? '#fef08a' : 'none', lh: note ? LH : st.lh });
  p.objs.push(o); startEdit(p, o, before);
}

/* ───────── inserting things ───────── */
function viewCenter(p) {
  const pe = PE.get(p.id), r = viewer.getBoundingClientRect();
  const c = new DOMPoint(r.left + r.width / 2, r.top + r.height / 2).matrixTransform(pe.g.getScreenCTM().inverse());
  return [clamp(c.x, 0, p.w), clamp(c.y, 0, p.h)];
}
function placeObj(o, w, h) { // centre a new object in the visible part of the current page and select it
  const p = S.pages[S.cur]; if (!p) return; pushUndo(); const c = viewCenter(p);
  o.x = clamp(c[0] - w / 2, 0, Math.max(0, p.w - w)); o.y = clamp(c[1] - h / 2, 0, Math.max(0, p.h - h));
  p.objs.push(o); drawObjs(p); setTool('select'); select(p.id, o.id);
}
function insertImage(url, w, h, opacity = 1) {
  const p = S.pages[S.cur]; if (!p) return;
  const f = Math.min(1, p.w * 0.6 / w, p.h * 0.6 / h); w *= f; h *= f;
  placeObj({ id: uid(), type: 'image', asset: addAsset(url), w, h, opacity }, w, h);
}
const textObj = (text, size, extra) => ({ id: uid(), type: 'text', x: 0, y: 0, text, size, color: '#333333', font: 'Helvetica', bold: false, italic: false, underline: false, strike: false, align: 'left', lh: LH, bg: 'none', opacity: 1, ...extra });
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

/* ───────── page operations ───────── */
function pageOp(fn) { if (!S.pages.length) return; pushUndo(); fn(); const y = viewer.scrollTop; buildPages(); viewer.scrollTop = y; }
const rotatePage = (i, d) => pageOp(() => { S.pages[i].rot = (S.pages[i].rot + d + 360) % 360; });
const dupPage = i => pageOp(() => { const n = clone(S.pages[i]); n.id = uid(); n.objs.forEach(o => o.id = uid()); S.pages.splice(i + 1, 0, n); });
function deletePage(i) { if (S.pages.length < 2) return toast('A document needs at least one page', true); pageOp(() => S.pages.splice(i, 1)); }
function addBlank() { const c = S.pages[S.cur], [w, h] = c ? rdims(c) : A4; pageOp(() => S.pages.splice(S.cur + 1, 0, { id: uid(), src: -1, w, h, rot: 0, objs: [] })); goToPage(S.cur + 1); }
function movePage(from, to) { if (from === to || from < 0) return; pageOp(() => { const [p] = S.pages.splice(from, 1); S.pages.splice(to, 0, p); }); }
function clearPage() { const p = S.pages[S.cur]; if (!p || !p.objs.length) return toast('No edits on this page'); pushUndo(); p.objs = []; S.sel = null; drawObjs(p); drawSel(); updateUI(); toast('Edits removed (Ctrl+Z to undo)'); }
function parseRange(str, n) {
  const out = [];
  for (const part of str.split(',').map(s => s.trim()).filter(Boolean)) {
    const m = part.match(/^(\d+)(?:\s*-\s*(\d+))?$/); if (!m) return null;
    const a = +m[1], b = m[2] ? +m[2] : a; if (a < 1 || b > n || a > b) return null;
    for (let i = a; i <= b; i++) out.push(i - 1);
  }
  return out.length ? out : null;
}
async function extractDlg() {
  const n = S.pages.length;
  const r = await dialog('Extract / split pages', `<p>Save selected pages as a new PDF (with your edits).</p>
    <label>Pages (e.g. 1-3, 5)<input type="text" name="range" value="${S.cur + 1}" required></label>`, { ok: 'Extract' });
  if (!r) return; const idx = parseRange(r.range, n); if (!idx) return toast(`Invalid range. Use numbers 1–${n}, e.g. 1-3, 5`, true);
  await busy(async () => download(new Blob([await buildPdf(idx.map(i => S.pages[i]))], { type: 'application/pdf' }), `${baseName()}-pages.pdf`), 'Extracting…');
}
const loadImage = url => new Promise((res, rej) => { const im = new Image(); im.onload = () => res(im); im.onerror = () => rej(new Error('image')); im.src = url; });
function paintObj(ctx, o, imgs) { // canvas twin of objEl — image exports use it so bundled fonts and blend modes come out exactly as on screen
  ctx.save(); ctx.globalAlpha = o.opacity ?? 1; ctx.fontKerning = 'none'; if (o.blend) ctx.globalCompositeOperation = 'multiply';
  if (o.rot && !isLine(o)) { const c = center(bbox(o)); ctx.translate(c[0], c[1]); ctx.rotate(o.rot * Math.PI / 180); ctx.translate(-c[0], -c[1]); }
  const hasS = o.color && o.color !== 'none' && o.width > 0, hasF = o.fill && o.fill !== 'none';
  ctx.lineCap = ctx.lineJoin = 'round'; if (hasS) { ctx.strokeStyle = o.color; ctx.lineWidth = o.width; ctx.setLineDash(dashArr(o) || []); } if (hasF) ctx.fillStyle = o.fill;
  const paint = () => { if (hasF) ctx.fill(); if (hasS) ctx.stroke(); };
  if (o.type === 'path') { ctx.beginPath(); for (const s of pathSegs(o.pts)) { if (s[0] === 'M') ctx.moveTo(s[1], s[2]); else if (s[0] === 'L') ctx.lineTo(s[1], s[2]); else ctx.quadraticCurveTo(s[1], s[2], s[3], s[4]); } if (hasS) ctx.stroke(); }
  else if (o.type === 'rect') { ctx.lineJoin = 'miter'; ctx.beginPath(); const r = o.rx ? Math.min(o.rx, o.w / 2, o.h / 2) : 0; if (r && ctx.roundRect) ctx.roundRect(o.x, o.y, o.w, o.h, r); else ctx.rect(o.x, o.y, o.w, o.h); paint(); }
  else if (o.type === 'ellipse') { ctx.beginPath(); ctx.ellipse(o.x + o.w / 2, o.y + o.h / 2, o.w / 2, o.h / 2, 0, 0, Math.PI * 2); paint(); }
  else if (o.type === 'line') { if (hasS) { ctx.beginPath(); ctx.moveTo(o.x1, o.y1); ctx.lineTo(o.x2, o.y2); ctx.stroke(); } }
  else if (o.type === 'arrow') {
    const ar = arrowGeo(o); ctx.lineCap = 'butt'; ctx.beginPath(); ctx.moveTo(o.x1, o.y1); ctx.lineTo(ar.bx, ar.by); if (hasS) ctx.stroke();
    ctx.fillStyle = o.color; ctx.beginPath(); ar.head.forEach((h, i) => i ? ctx.lineTo(h[0], h[1]) : ctx.moveTo(h[0], h[1])); ctx.fill();
  } else if (o.type === 'image') { const im = imgs[o.asset]; if (im) ctx.drawImage(im, o.x, o.y, o.w, o.h); }
  else if (o.type === 'text') {
    const b = bbox(o); if (o.bg && o.bg !== 'none') { ctx.fillStyle = o.bg; ctx.fillRect(b.x - 3, b.y - 2, b.w + 6, b.h + 4); }
    ctx.fillStyle = o.color; ctx.wordSpacing = (o.ws || 0) * o.size + 'px';
    for (const l of textLines(o)) { ctx.font = cssFont(o); ctx.fillText(l.ln, l.x, l.y); if (l.ln) for (const dy of decoY(o)) ctx.fillRect(l.x, l.y + o.size * dy - o.size / 32, l.w, Math.max(0.5, o.size / 16)); }
  }
  ctx.restore();
}
async function renderPageImage(p, k) { // the page with all edits, as a canvas at k pixels per point
  const [rw, rh] = rdims(p), c = document.createElement('canvas'); c.width = Math.max(1, Math.round(rw * k)); c.height = Math.max(1, Math.round(rh * k));
  const ctx = c.getContext('2d', { willReadFrequently: true }); ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, c.width, c.height);
  // intent 'print' keeps pdf.js from waiting on animation frames, so exports also finish in a background tab
  if (p.src >= 0) { const pg = await getPdfPage(p); await pg.render({ canvasContext: ctx, intent: 'print', viewport: pg.getViewport({ scale: k, rotation: (pg.rotate + p.rot) % 360 }) }).promise; }
  if (p.objs.length) {
    const imgs = {};
    await Promise.all([...new Set(p.objs.filter(o => o.type === 'text').map(o => o.font))].map(ensureFont));
    await Promise.all([...new Set(p.objs.filter(o => o.type === 'image' && S.assets[o.asset]).map(o => o.asset))].map(async a => { try { imgs[a] = await loadImage(S.assets[a]); } catch (e) { /* unreadable picture is skipped */ } }));
    ctx.setTransform(k, 0, 0, k, 0, 0);
    if (p.rot === 90) ctx.transform(0, 1, -1, 0, p.h, 0); else if (p.rot === 180) ctx.transform(-1, 0, 0, -1, p.w, p.h); else if (p.rot === 270) ctx.transform(0, -1, 1, 0, 0, p.w);
    for (const o of p.objs) paintObj(ctx, o, imgs);
    ctx.setTransform(1, 0, 0, 1, 0, 0);
  }
  return c;
}
async function exportPng() {
  commitEdit(); const p = S.pages[S.cur]; if (!p) return;
  await busy(async () => {
    const c = await renderPageImage(p, 2 * CSS_UNITS);
    download(await new Promise(res => c.toBlob(res, 'image/png')), `${baseName()}-page${S.cur + 1}.png`);
  }, 'Rendering…');
}

/* ───────── export (pdf-lib) ───────── */
const rgb01 = h => { const n = parseInt(h.slice(1), 16); return [(n >> 16 & 255) / 255, (n >> 8 & 255) / 255, (n & 255) / 255]; };
function rasterText(o) {
  const b = bbox(o), k = 4, c = document.createElement('canvas'); c.width = Math.ceil((b.w + 4) * k); c.height = Math.ceil((b.h + 4) * k);
  const x = c.getContext('2d'); x.fontKerning = 'none'; x.scale(k, k); x.translate(-o.x, -o.y); x.font = cssFont(o); x.fillStyle = o.color; x.wordSpacing = (o.ws || 0) * o.size + 'px';
  for (const l of textLines(o)) { x.font = cssFont(o); x.fillText(l.ln, l.x, l.y); if (l.ln) for (const dy of decoY(o)) x.fillRect(l.x, l.y + o.size * dy - o.size / 32, l.w, Math.max(0.5, o.size / 16)); }
  return { url: c.toDataURL('image/png'), w: c.width / k, h: c.height / k };
}
async function buildPdf(pages) {
  commitEdit();
  const L = PDFLib, out = await L.PDFDocument.create(), docs = {}, copies = new Map(), occ = {}, batches = {};
  for (const p of pages) { if (p.src < 0) continue; const kk = p.src + '|' + p.idx, k = occ[kk] = (occ[kk] || 0) + 1; (batches[p.src + '|' + k] = batches[p.src + '|' + k] || []).push(p); }
  for (const key in batches) {
    const src = +key.split('|')[0];
    if (!docs[src]) {
      try { docs[src] = await L.PDFDocument.load(S.sources[src].bytes); }
      catch (e) { throw new Error(/encrypt/i.test(e.message) ? 'This PDF is encrypted / password-protected and cannot be re-saved. Remove the password first.' : 'Could not process this PDF: ' + e.message); }
    }
    const list = batches[key], cp = await out.copyPages(docs[src], list.map(p => p.idx));
    list.forEach((p, i) => copies.set(p, cp[i]));
  }
  const fonts = {}, imgs = {};
  let kitReady = false;
  const simple = cp => cp < 0x0530 || (cp >= 0x1e00 && cp <= 0x22ff); // scripts that need no shaping: Latin, Greek, Cyrillic, punctuation, symbols
  const kitFont = async (key, bytes) => { // an embedded font file → { f, set }, or null when it cannot be embedded (whole file: pdf-lib's subsetter corrupts several faces)
    if (!(key in fonts)) {
      try { if (!kitReady) { out.registerFontkit(window.fontkit); kitReady = true; } const f = await out.embedFont(await bytes(), { subset: false, features: FEAT_OFF }); fonts[key] = { f, set: new Set(f.getCharacterSet()) }; }
      catch (e) { console.warn('font embed failed', key, e); fonts[key] = null; }
    }
    return fonts[key];
  };
  const famFont = async (fam, o) => { // one of the pickable families → { font, has(codepoint), slant, heavy }
    const B = BUNDLED[fam];
    if (!B) {
      const n = (PDF_FONTS[fam] || PDF_FONTS.Helvetica)[(o.bold ? 1 : 0) + (o.italic ? 2 : 0)], font = fonts[n] || (fonts[n] = await out.embedFont(L.StandardFonts[n])), ok = {};
      return { font, has: cp => cp in ok ? ok[cp] : (ok[cp] = (() => { try { font.encodeText(String.fromCodePoint(cp)); return true; } catch (e) { return false; } })()) };
    }
    const stem = B[0] + (o.bold && B[1] ? '-Bold' : '-Regular'), F = await kitFont(stem, () => fontBytes(stem));
    return F && { font: F.f, has: cp => simple(cp) && F.set.has(cp), slant: !!o.italic, heavy: !!o.bold && !B[1] }; // no italic / bold file → slant or thicken the regular face, as the browser does on screen
  };
  const getRuns = async o => { // → text → [{ text, font, slant, heavy }] for one line, or null when this object has to be drawn as an image instead
    const X = S.xfonts[o.font], D = X && await kitFont(o.font, () => X.bytes); if (X && !D) return null;
    const P = D && { font: D.f, has: cp => simple(cp) && D.set.has(cp), slant: !!o.italic && !X.italic, heavy: !!o.bold && !X.bold }; // the document's own font …
    const fb = await famFont(X ? X.fb : o.font, o), pick = cp => P && P.has(cp) ? P : fb && fb.has(cp) ? fb : null; // … and the nearest family for characters it lacks
    for (const ch of o.text) if (ch !== '\n' && !pick(ch.codePointAt(0))) return null;
    return ln => { const runs = []; for (const ch of ln) { const F = pick(ch.codePointAt(0)), r = runs[runs.length - 1]; if (r && r.F === F && !(o.ws && r.text.endsWith(' '))) r.text += ch; else runs.push({ F, text: ch }); } return runs; };
  };
  const embed = url => /^data:image\/jpe?g/.test(url) ? out.embedJpg(url) : out.embedPng(url);
  const gs = (page, op, blend) => { if (op >= 1 && !blend) return []; const d = { Type: 'ExtGState', ca: op, CA: op }; if (blend) d.BM = 'Multiply'; return [L.setGraphicsState(page.node.newExtGState('XGS', out.context.obj(d)))]; };
  const drawImg = (page, img, o, x, y, w, h) => {
    page.pushOperators(L.pushGraphicsState(), L.concatTransformationMatrix(w, 0, 0, -h, x, y + h));
    page.drawImage(img, { x: 0, y: 0, width: 1, height: 1, opacity: o.opacity ?? 1 });
    page.pushOperators(L.popGraphicsState());
  };
  const rectOps = (x, y, w, h) => [L.moveTo(x, y), L.lineTo(x + w, y), L.lineTo(x + w, y + h), L.lineTo(x, y + h), L.closePath()];
  const roundRectOps = (x, y, w, h, r) => { const k = r * 0.4477; return [L.moveTo(x + r, y), L.lineTo(x + w - r, y), L.appendBezierCurve(x + w - k, y, x + w, y + k, x + w, y + r),
    L.lineTo(x + w, y + h - r), L.appendBezierCurve(x + w, y + h - k, x + w - k, y + h, x + w - r, y + h), L.lineTo(x + r, y + h), L.appendBezierCurve(x + k, y + h, x, y + h - k, x, y + h - r),
    L.lineTo(x, y + r), L.appendBezierCurve(x, y + k, x + k, y, x + r, y), L.closePath()]; };

  async function drawOne(page, o) { // coordinates here are page view space (y down)
    const op = o.opacity ?? 1;
    if (o.type === 'image') { const url = S.assets[o.asset]; if (!url) return; drawImg(page, imgs[o.asset] || (imgs[o.asset] = await embed(url)), o, o.x, o.y, o.w, o.h); return; }
    if (o.type === 'text') {
      await ensureFont(o.font);
      const runsOf = await getRuns(o), b = bbox(o), k = AL[o.align] || 0, col = rgb01(o.color);
      if (o.bg && o.bg !== 'none') page.pushOperators(L.pushGraphicsState(), ...gs(page, op), L.setFillingRgbColor(...rgb01(o.bg)), ...rectOps(b.x - 3, b.y - 2, b.w + 6, b.h + 4), L.fill(), L.popGraphicsState());
      if (!runsOf) { const r = rasterText(o); drawImg(page, await out.embedPng(r.url), o, o.x, o.y, r.w, r.h); return; } // text no font can encode (e.g. Hindi, Arabic, CJK) → crisp image
      textLines(o).forEach(l => {
        if (!l.ln) return; const runs = runsOf(l.ln); runs.forEach(r => r.w = r.F.font.widthOfTextAtSize(r.text, o.size) + (r.text.endsWith(' ') ? (o.ws || 0) * o.size : 0));
        const lw = runs.reduce((a, r) => a + r.w, 0), x = o.x + (b.w - lw) * k; let rx = x;
        for (const { F, text, w } of runs) {
          page.pushOperators(L.pushGraphicsState(), L.concatTransformationMatrix(1, 0, 0, -1, rx, l.y));
          if (F.heavy) page.pushOperators(L.setTextRenderingMode(L.TextRenderingMode.FillAndOutline), L.setStrokingRgbColor(...col), L.setLineWidth(o.size * 0.035));
          page.drawText(text, { x: 0, y: 0, size: o.size, font: F.font, color: L.rgb(...col), opacity: op, ...(F.slant ? { ySkew: L.degrees(12) } : {}) });
          page.pushOperators(L.popGraphicsState()); rx += w;
        }
        for (const dy of decoY(o)) page.pushOperators(L.pushGraphicsState(), ...gs(page, op), L.setStrokingRgbColor(...rgb01(o.color)), L.setLineWidth(Math.max(0.5, o.size / 16)),
          L.moveTo(x, l.y + o.size * dy), L.lineTo(x + lw, l.y + o.size * dy), L.stroke(), L.popGraphicsState());
      });
      return;
    }
    const hasS = o.color && o.color !== 'none' && o.width > 0, hasF = o.fill && o.fill !== 'none', ops = [L.pushGraphicsState(), ...gs(page, op, o.blend)];
    if (hasS) { ops.push(L.setStrokingRgbColor(...rgb01(o.color)), L.setLineWidth(o.width), L.setLineCap(L.LineCapStyle.Round), L.setLineJoin(L.LineJoinStyle.Round)); const da = dashArr(o); if (da) ops.push(L.setDashPattern(da, 0)); }
    if (hasF) ops.push(L.setFillingRgbColor(...rgb01(o.fill)));
    const paint = () => hasF && hasS ? L.fillAndStroke() : hasF ? L.fill() : hasS ? L.stroke() : L.endPath();
    if (o.type === 'path') {
      let cx = 0, cy = 0;
      for (const s of pathSegs(o.pts)) {
        if (s[0] === 'M') ops.push(L.moveTo(s[1], s[2])); else if (s[0] === 'L') ops.push(L.lineTo(s[1], s[2]));
        else ops.push(L.appendBezierCurve(cx + 2 / 3 * (s[1] - cx), cy + 2 / 3 * (s[2] - cy), s[3] + 2 / 3 * (s[1] - s[3]), s[4] + 2 / 3 * (s[2] - s[4]), s[3], s[4]));
        cx = s[s.length - 2]; cy = s[s.length - 1];
      }
      ops.push(L.stroke());
    } else if (o.type === 'rect') ops.push(L.setLineJoin(L.LineJoinStyle.Miter), ...(o.rx ? roundRectOps(o.x, o.y, o.w, o.h, Math.min(o.rx, o.w / 2, o.h / 2)) : rectOps(o.x, o.y, o.w, o.h)), paint());
    else if (o.type === 'ellipse') {
      const rx = o.w / 2, ry = o.h / 2, cx = o.x + rx, cy = o.y + ry, kx = rx * 0.5523, ky = ry * 0.5523;
      ops.push(L.moveTo(cx + rx, cy), L.appendBezierCurve(cx + rx, cy + ky, cx + kx, cy + ry, cx, cy + ry), L.appendBezierCurve(cx - kx, cy + ry, cx - rx, cy + ky, cx - rx, cy),
        L.appendBezierCurve(cx - rx, cy - ky, cx - kx, cy - ry, cx, cy - ry), L.appendBezierCurve(cx + kx, cy - ry, cx + rx, cy - ky, cx + rx, cy), L.closePath(), paint());
    } else if (o.type === 'line') ops.push(L.moveTo(o.x1, o.y1), L.lineTo(o.x2, o.y2), L.stroke());
    else if (o.type === 'arrow') {
      const ar = arrowGeo(o), h = ar.head;
      ops.push(L.setLineCap(L.LineCapStyle.Butt), L.moveTo(o.x1, o.y1), L.lineTo(ar.bx, ar.by), L.stroke(),
        L.setFillingRgbColor(...rgb01(o.color)), L.moveTo(...h[0]), L.lineTo(...h[1]), L.lineTo(...h[2]), L.closePath(), L.fill());
    }
    ops.push(L.popGraphicsState()); page.pushOperators(...ops);
  }

  for (const p of pages) {
    let page, inv, total = p.rot;
    if (p.src < 0) { page = out.addPage([p.w, p.h]); inv = [1, 0, 0, -1, 0, p.h]; }
    else {
      page = out.addPage(copies.get(p));
      const pg = await getPdfPage(p), [a, b, c, d, e, f] = pg.getViewport({ scale: 1 }).transform, det = a * d - b * c;
      inv = [d / det, -b / det, -c / det, a / det, (c * f - d * e) / det, (b * e - a * f) / det]; total = (pg.rotate + p.rot) % 360;
    }
    page.setRotation(L.degrees(total));
    if (!p.objs.length) continue;
    page.pushOperators(L.pushGraphicsState(), L.concatTransformationMatrix(...inv));
    for (const o of p.objs) {
      const rot = !isLine(o) && o.rot;
      if (rot) { const c = center(bbox(o)), a = rot * Math.PI / 180, cs = Math.cos(a), sn = Math.sin(a); page.pushOperators(L.pushGraphicsState(), L.concatTransformationMatrix(cs, sn, -sn, cs, c[0] - c[0] * cs + c[1] * sn, c[1] - c[0] * sn - c[1] * cs)); }
      await drawOne(page, o);
      if (rot) page.pushOperators(L.popGraphicsState());
    }
    page.pushOperators(L.popGraphicsState());
  }
  const m = S.meta;
  if (m.title) out.setTitle(m.title); if (m.author) out.setAuthor(m.author); if (m.subject) out.setSubject(m.subject);
  if (m.keywords) out.setKeywords(m.keywords.split(',').map(s => s.trim()).filter(Boolean));
  out.setProducer('XD3 PDF Editor (XD3Labs / XDCybertech Pvt Ltd)'); out.setModificationDate(new Date());
  return out.save();
}
async function save() {
  await busy(async () => {
    const bytes = await buildPdf(S.pages);
    download(new Blob([bytes], { type: 'application/pdf' }), /-edited$/.test(baseName()) || S.sources.length === 0 ? S.fileName : baseName() + '-edited.pdf');
    S.dirty = false; updateUI(); flash($('#topbar [data-act=save]'), 'flash'); toast('PDF saved to your downloads');
  }, 'Building PDF…');
}
async function printDoc() {
  await busy(async () => {
    const url = URL.createObjectURL(new Blob([await buildPdf(S.pages)], { type: 'application/pdf' }));
    let f = $('#printframe'); if (f) f.remove();
    f = document.createElement('iframe'); f.id = 'printframe'; f.style.cssText = 'position:fixed;right:0;bottom:0;width:0;height:0;border:0';
    f.onload = () => setTimeout(() => { try { f.contentWindow.focus(); f.contentWindow.print(); } catch (e) { window.open(url, '_blank'); } }, 300);
    f.src = url; document.body.appendChild(f);
  }, 'Preparing print…');
}
async function propsDlg() {
  const m = S.meta;
  const r = await dialog('Document properties', `
    <label>File name<input type="text" name="file" value="${esc(S.fileName)}"></label>
    <label>Title<input type="text" name="title" value="${esc(m.title || '')}"></label>
    <label>Author<input type="text" name="author" value="${esc(m.author || '')}"></label>
    <label>Subject<input type="text" name="subject" value="${esc(m.subject || '')}"></label>
    <label>Keywords (comma separated)<input type="text" name="keywords" value="${esc(m.keywords || '')}"></label>
    <p style="color:var(--mut)">${S.pages.length} page(s) · applied when you save.</p>`, { ok: 'Apply' });
  if (!r) return;
  S.meta = { title: r.title, author: r.author, subject: r.subject, keywords: r.keywords };
  if (r.file.trim()) S.fileName = r.file.trim().replace(/\.pdf$/i, '') + '.pdf';
  S.dirty = true; updateUI(); scheduleSave();
}

/* ───────── UI wiring ───────── */
function icons(root = document) {
  $$('i[data-icon]', root).forEach(i => { if (!i.firstChild) i.innerHTML = `<svg viewBox="0 0 24 24"><path d="${ICONS[i.dataset.icon] || ''}"/></svg>`; });
}
function updateUI() {
  const has = S.pages.length > 0; document.body.classList.toggle('hasdoc', has);
  $('#bUndo').disabled = !S.undo.length; $('#bRedo').disabled = !S.redo.length;
  $('#zoomLbl').textContent = Math.round(S.zoom * 100) + '%';
  if (document.activeElement !== $('#pageNo')) $('#pageNo').value = has ? S.cur + 1 : 0;
  $('#pageNo').max = S.pages.length; $('#pageCnt').textContent = '/ ' + S.pages.length;
  $('#fileName').textContent = has ? S.fileName + (S.dirty ? ' •' : '') : '';
  const sel = !!selObjs(); $$('.selonly').forEach(b => b.disabled = !sel);
  const cur = S.pages[S.cur];
  $$('.thumb', thumbs).forEach(t => { const on = cur && t.dataset.id === cur.id; if (on && !t.classList.contains('active')) t.scrollIntoView({ block: 'nearest' }); t.classList.toggle('active', !!on); });
}
const NO_DOC = new Set(['palette', 'accent', 'notesclear', 'merge', 'mergedlg', 'ed_cut', 'ed_copy', 'ed_paste', 'ed_selall', 'ed_done', 'open', 'newdoc', 'theme', 'about', 'shortcuts', 'sidebar', 'restore', 'scan', 'docx2pdf']);
const ACT = {
  open: () => $('#fOpen').click(), save, print: printDoc, undo, redo, newdoc: newDoc, restore: restoreSession,
  zoomin: () => setZoom(S.zoom * 1.2), zoomout: () => setZoom(S.zoom / 1.2), zoomreset: () => setZoom(1), fit: () => setZoom(fitZoom()), fitpage: () => setZoom(fitZoom(true)),
  prev: () => goToPage(S.cur - 1), next: () => goToPage(S.cur + 1),
  blank: addBlank, dup: () => dupPage(S.cur), rotr: () => rotatePage(S.cur, 90), rotl: () => rotatePage(S.cur, -90),
  rotall: () => pageOp(() => S.pages.forEach(p => p.rot = (p.rot + 90) % 360)), delpage: () => deletePage(S.cur),
  fwd: () => stepOrder(1), bwd: () => stepOrder(-1), orot: b => rotateSel(+b.dataset.v), opac: b => quickProp('opacity', +b.dataset.v), sdash: b => quickProp('dash', b.dataset.v),
  tstyle: b => { const o = textTarget(); if (o) quickProp(b.dataset.k, !o[b.dataset.k]); }, talign: b => quickProp('align', b.dataset.v),
  tsize: b => { const o = textTarget(); if (o) quickProp('size', clamp(Math.round(o.size) + +b.dataset.v, 4, 300)); }, setfont: b => { ensureFont(b.dataset.f); quickProp('font', b.dataset.f); },
  toallpages: toAllPages, imgreplace: imgReplace, imgsave: imgSave,
  ed_cut: () => document.execCommand('cut'), ed_copy: () => document.execCommand('copy'), ed_paste: edPaste, ed_selall: () => document.execCommand('selectAll'), ed_done: () => commitEdit(),
  pgtop: () => { if (S.cur > 0) { movePage(S.cur, 0); goToPage(0); } }, pgend: () => { const n = S.pages.length - 1; if (S.cur < n) { movePage(S.cur, n); goToPage(n); } },
  merge: () => $('#fMerge').click(), imgpages: () => $('#fImgPages').click(), extract: extractDlg, png: exportPng,
  image: () => $('#fImg').click(), sign: signDlg, watermark: watermarkDlg, pagenum: pageNumDlg, headfoot: headFootDlg, datestamp: insertDate,
  stamp: b => { const s = makeStamp(b.dataset.t, b.dataset.c, !!b.dataset.plain); insertImage(s.url, s.w, s.h); },
  props: propsDlg, find: toggleFind, findprev: () => stepFind(-1), findnext: () => stepFind(1),
  selall: () => { const p = S.pages[S.cur]; setTool('select'); select(p.id, p.objs.map(o => o.id)); if (!p.objs.length) toast('Nothing to select on this page yet'); }, clearpage: clearPage,
  sidebar: () => { document.body.classList.toggle('nosb'); document.body.classList.toggle('sbopen'); },
  theme: () => { const t = document.documentElement.dataset.theme === 'dark' ? 'light' : 'dark'; document.documentElement.dataset.theme = t; try { localStorage.setItem('xd3.theme', t); } catch (e) { /* ignore */ } },
  del: delSel, osdup: dupSel, copy: copySel, cut: () => { copySel(); delSel(); }, paste: () => pasteClip(), pastehere: () => pasteClip(S.ctxPt), front: () => reorder(true), back: () => reorder(false),
  align: b => alignSel(b.dataset.m), distribute: b => distributeSel(b.dataset.m), copystyle: copyStyle, pastestyle: pasteStyle,
  texthere: () => newTextAt(false), notehere: () => newTextAt(true), tool: b => setTool(b.dataset.t),
  pgup: () => { if (S.cur > 0) { movePage(S.cur, S.cur - 1); goToPage(S.cur - 1); } }, pgdown: () => { if (S.cur < S.pages.length - 1) { movePage(S.cur, S.cur + 1); goToPage(S.cur + 1); } },
  findrep: () => { toggleFind(true); $('#replInput').focus(); }, repl1: replaceOne, replall: replaceAll, docx2pdf: () => $('#fDocx').click(),
  edittxt: () => { const s = selOne(); if (s && s.o.type === 'text') startEdit(s.p, s.o, snap()); },
  shortcuts: () => dialog('Keyboard shortcuts', `<div class="keys">${[
    ['Ctrl+O / Ctrl+S / Ctrl+P', 'Open · Save · Print'], ['Ctrl+Z / Ctrl+Y', 'Undo / redo'], ['Ctrl+F / Ctrl+H', 'Find · Find & replace'], ['G', 'Organize pages'], ['Ctrl+X', 'Cut selection'],
    ['Ctrl + / −  ·  Ctrl+wheel', 'Zoom'], ['V E B T', 'Select · Edit text (also in scans and pictures) · Edit object · Add text'], ['P M H N', 'Pen · Highlighter · Highlight area · Note'],
    ['R O L A W', 'Rectangle · Ellipse · Line · Arrow · Whiteout'], ['Del', 'Delete selection'], ['Ctrl+A', 'Select everything on the page'], ['Ctrl+D', 'Duplicate selection'],
    ['Ctrl+C / Ctrl+V', 'Copy / paste objects (or paste an image)'], ['Ctrl+B / I / U', 'Bold · italic · underline'], ['Arrows', 'Nudge selection (Shift = ×10)'],
    ['Shift+drag', 'Square / 45° lines / 15° rotation / free image resize'], ['Shift+click', 'Add to selection'], ['Double‑click', 'Edit a text box, or open the image editor'], ['Right‑click', 'Menu for the object, page or thumbnail'],
  ].map(([k, v]) => `<kbd>${k}</kbd><span>${v}</span>`).join('')}</div>`),
  about: () => dialog('XD3 PDF Editor', `<p><b>A free, private, offline PDF editor.</b></p>
    <p>Built by <b>XD3Labs</b> — a product of <b>XDCybertech Pvt Ltd</b>.</p>
    <p style="color:var(--mut)">Everything runs on your device — no internet connection is needed and your documents are never uploaded. The app is locked down so it can only load its own bundled files.</p>
    <p style="color:var(--mut)">Text in scans and pictures is read on your device by the bundled Tesseract OCR engine (Apache 2.0 licence).</p>
    <p style="color:var(--mut)">Bundled fonts come from Google Fonts under the SIL Open Font, Apache 2.0 and Ubuntu Font licences (see lib/fonts/README.txt).</p>
    <p style="color:var(--mut)">Note: edited text, picked-up objects, whiteout and blackout cover the original content visually; the original data underneath remains in the file.</p>`),
};
document.addEventListener('click', e => {
  const sb = e.target.closest('.hassub'); if (sb) { clearTimeout(subTimer); if (!sb.classList.contains('subopen')) openSub(sb); return; } // a submenu parent keeps the menu open
  const mb = e.target.closest('[data-menu]'), ab = e.target.closest('[data-act]');
  hideCtx(); $$('.menu.open').forEach(m => { if (!mb || m.id !== mb.dataset.menu) m.classList.remove('open'); });
  if (mb) {
    const m = $('#' + mb.dataset.menu); m.classList.toggle('open');
    if (m.id === 'mFonts') { buildFontMenu(); Object.keys(BUNDLED).forEach(ensureFont); } // fonts found in the document are listed too, and every name previews in its own face
    if (m.classList.contains('fx')) { const r = mb.getBoundingClientRect(); placeMenu(m, r.left, r.bottom + 6); }
    return;
  }
  const tb = e.target.closest('button[data-tool]'); if (tb) { if (S.pages.length) setTool(tb.dataset.tool); return; }
  if (ab) { const a = ab.dataset.act; if (!S.pages.length && !NO_DOC.has(a)) return toast('Open a PDF first'); if (!('keep' in ab.dataset)) commitEdit(); if (ACT[a]) ACT[a](ab); }
});
thumbs.addEventListener('click', e => {
  const t = e.target.closest('.thumb'); if (!t) return; const i = S.pages.findIndex(p => p.id === t.dataset.id), a = e.target.closest('[data-a]');
  if (!a) { goToPage(i); document.body.classList.remove('sbopen'); return; }
  if (a.dataset.a === 'rot') rotatePage(i, 90); else if (a.dataset.a === 'dup') dupPage(i); else deletePage(i);
});
let dragPid = null;
thumbs.addEventListener('dragstart', e => { const t = e.target.closest('.thumb'); if (!t) return; dragPid = t.dataset.id; e.dataTransfer.effectAllowed = 'move'; e.dataTransfer.setData('text/plain', dragPid); });
thumbs.addEventListener('dragover', e => { if (!dragPid) return; e.preventDefault(); $$('.dropt', thumbs).forEach(x => x.classList.remove('dropt')); const t = e.target.closest('.thumb'); if (t) t.classList.add('dropt'); });
thumbs.addEventListener('drop', e => { if (!dragPid) return; e.preventDefault(); e.stopPropagation(); const t = e.target.closest('.thumb'), from = S.pages.findIndex(p => p.id === dragPid); dragPid = null; if (t) movePage(from, S.pages.findIndex(p => p.id === t.dataset.id)); });
thumbs.addEventListener('dragend', () => { dragPid = null; $$('.dropt', thumbs).forEach(x => x.classList.remove('dropt')); });

$('#pageNo').addEventListener('change', e => goToPage((+e.target.value || 1) - 1));
$('#pageNo').addEventListener('keydown', e => e.stopPropagation());
const onFiles = (sel, fn) => $(sel).addEventListener('change', e => { const f = [...e.target.files]; e.target.value = ''; if (f.length) fn(f); });
onFiles('#fOpen', f => openFiles(f, false));
onFiles('#fMerge', f => openFiles(f, S.pages.length > 0));
onFiles('#fImgPages', f => openFiles(f, S.pages.length > 0));
onFiles('#fDocx', f => openFiles(f, false).then(() => { if (S.pages.length) toast('Word document converted — review it, then click Save PDF'); }));
onFiles('#fImg', f => busy(async () => { const im = await readImage(f[0]); insertImage(im.url, im.w * 0.75, im.h * 0.75); }));

// drag & drop files
let dragDepth = 0; const veil = $('#dropveil'), hasFiles = e => e.dataTransfer && [...e.dataTransfer.types].includes('Files');
window.addEventListener('dragenter', e => { if (hasFiles(e)) { dragDepth++; veil.classList.add('on'); } });
window.addEventListener('dragleave', e => { if (hasFiles(e) && --dragDepth <= 0) { dragDepth = 0; veil.classList.remove('on'); } });
window.addEventListener('dragover', e => { if (hasFiles(e)) e.preventDefault(); });
window.addEventListener('drop', e => {
  if (!hasFiles(e)) return; e.preventDefault(); dragDepth = 0; veil.classList.remove('on');
  const files = [...e.dataTransfer.files]; if (!files.length) return;
  if (!S.pages.length) return openFiles(files, false);
  const imgs = files.filter(f => f.type.startsWith('image/')), pdfs = files.filter(f => isPdf(f) || isDocx(f));
  if (pdfs.length) openFiles(pdfs, true);
  else if (imgs.length) busy(async () => { for (const f of imgs) { const im = await readImage(f); insertImage(im.url, im.w * 0.75, im.h * 0.75); } });
});
window.addEventListener('paste', e => {
  if (!S.pages.length || dlg.open || S.editing || /INPUT|TEXTAREA|SELECT/.test(document.activeElement.tagName)) return;
  const f = [...(e.clipboardData ? e.clipboardData.files : [])].find(x => x.type.startsWith('image/'));
  if (f) { e.preventDefault(); busy(async () => { const im = await readImage(f); insertImage(im.url, im.w * 0.75, im.h * 0.75); }); }
  else if (S.clip) pasteClip();
});

viewer.addEventListener('wheel', e => { if (!e.ctrlKey && !e.metaKey) return; e.preventDefault(); setZoom(S.zoom * (e.deltaY < 0 ? 1.1 : 1 / 1.1)); }, { passive: false });
// two-finger pinch zoom
let pinch = null; const tdist = t => Math.hypot(t[0].clientX - t[1].clientX, t[0].clientY - t[1].clientY);
viewer.addEventListener('touchstart', e => { if (e.touches.length === 2) pinch = { d: tdist(e.touches), z: S.zoom }; }, { passive: true });
viewer.addEventListener('touchmove', e => { if (pinch && e.touches.length === 2) { e.preventDefault(); setZoom(pinch.z * tdist(e.touches) / pinch.d); } }, { passive: false });
viewer.addEventListener('touchend', e => { if (e.touches.length < 2) pinch = null; });

const TOOL_KEYS = { v: 'select', t: 'text', e: 'edittext', b: 'editobj', n: 'note', p: 'pen', m: 'marker', h: 'hl', r: 'rect', o: 'ellipse', l: 'line', a: 'arrow', w: 'whiteout' };
window.addEventListener('keydown', e => {
  if (dlg.open || document.body.classList.contains('modal')) return;
  const typing = /INPUT|TEXTAREA|SELECT/.test(e.target.tagName) || e.target.isContentEditable, k = e.key.toLowerCase(), mod = e.ctrlKey || e.metaKey;
  if (mod) {
    if (k === 'o') { e.preventDefault(); ACT.open(); }
    else if (k === 'h') { e.preventDefault(); if (S.pages.length) ACT.findrep(); }
    else if (k === 's') { e.preventDefault(); if (S.pages.length) save(); }
    else if (k === 'p') { e.preventDefault(); if (S.pages.length) printDoc(); }
    else if (k === 'f') { e.preventDefault(); if (S.pages.length) toggleFind(); }
    else if (typing || !S.pages.length) return;
    else if (k === 'z') { e.preventDefault(); e.shiftKey ? redo() : undo(); }
    else if (k === 'y') { e.preventDefault(); redo(); }
    else if (k === 'd') { e.preventDefault(); dupSel(); }
    else if (k === 'a') { e.preventDefault(); ACT.selall(); }
    else if (k === 'c') copySel();
    else if (k === 'x') { if (selObjs()) { e.preventDefault(); ACT.cut(); } }
    else if ({ b: 1, i: 1, u: 1 }[k] && selObjs()) { e.preventDefault(); const pk = { b: 'bold', i: 'italic', u: 'underline' }[k], t = selObjs().os.find(o => o.type === 'text'); if (t) { setProp(pk, !t[pk]); armed = false; syncControls(); } }
    else if (k === '=' || k === '+') { e.preventDefault(); setZoom(S.zoom * 1.2); }
    else if (k === '-') { e.preventDefault(); setZoom(S.zoom / 1.2); }
    else if (k === '0') { e.preventDefault(); setZoom(1); }
    return;
  }
  if (typing || !S.pages.length) return;
  if (k === 'delete' || k === 'backspace') { if (selObjs()) { e.preventDefault(); delSel(); } }
  else if (k === 'escape') { hideCtx(); select(null); setTool('select'); }
  else if (k === 'enter') { const s = selOne(); if (s && s.o.type === 'text') { e.preventDefault(); startEdit(s.p, s.o, snap()); } }
  else if (k.startsWith('arrow') && selObjs()) { e.preventDefault(); const d = e.shiftKey ? 10 : 1; nudge(k === 'arrowleft' ? -d : k === 'arrowright' ? d : 0, k === 'arrowup' ? -d : k === 'arrowdown' ? d : 0); }
  else if (k === 'g' && !e.altKey) ACT.organize();
  else if (TOOL_KEYS[k] && !e.altKey) setTool(TOOL_KEYS[k]);
});
window.addEventListener('beforeunload', e => { if (S.dirty && S.pages.length) { e.preventDefault(); e.returnValue = ''; } });
window.addEventListener('resize', () => { hideCtx(); drawSel(); });

try { const t = localStorage.getItem('xd3.theme'); if (t) document.documentElement.dataset.theme = t; } catch (e) { /* ignore */ }
DB.get('session').then(s => { if (s && s.pages && s.pages.length) { const b = $('#bRestore'); b.hidden = false; b.title = s.fileName || ''; } }).catch(() => {});
$('#yr').textContent = new Date().getFullYear();
$('#hint').textContent = HINTS.select;
buildFontMenu(); icons(); syncControls(); updateUI();
const openArg = new URLSearchParams(location.search).get('open'); // desktop app: a file passed on the command line ("Open with…")
if (openArg) fetch('__open__/' + encodeURIComponent(openArg)).then(r => r.blob())
  .then(b => openFiles([new File([b], openArg, { type: /\.pdf$/i.test(openArg) ? 'application/pdf' : b.type })], false)).catch(() => toast('Could not open ' + openArg, true));
// shared API for the feature modules (imaging.js, convert.js, organize.js) and for automation
window.XD3 = { S, PE, ACT, NO_DOC, HINTS, A4, CSS_UNITS, LH, BASE, viewer, $, $$, svg, uid, clone, clamp, esc, toast, busy, dialog, download, icons,
  pushUndo, snap, resetDoc, scheduleSave, buildPages, layout, drawObjs, drawSel, select, selObjs, selOne, setTool, syncControls, updateUI, goToPage, fitZoom, setZoom,
  bbox, center, rotP, shiftObj, isLine, textLines, measure, baseOff, cssFont, rdims, baseName, parseRange, textObj, placeObj, insertImage, addAsset, readImage, imagePage,
  visible, getPdfPage, addSource, loadSource, openFiles, isPdf, isDocx, getText, getLines, getBlocks, toBlocks, getImages, scanPage, visibleText, covered, fontStyle, origFont, fontLabel, baseCanvas, baseSig, isBaseImg, aabb, hitsRect, hex, editOriginal, coverFor, convertRegion, renderPageImage, buildPdf, pageOp, movePage,
  rotatePage, dupPage, deletePage, pageNumDlg, runFind, convertTextItem, drawTI, showCtx, mi, pageSub, ensureFont, BUNDLED, FONTS, commitEdit, viewCenter, updateBell, flash };
})();
