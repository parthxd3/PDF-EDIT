/* XD3 PDF Editor — the home screen and the header tool menus: every tool as a card (from the tool list in js/flow.js),
   category chips and search, recent files and tools, Continue cards, and the writing assistant. */
(() => {
'use strict';
const X = window.XD3, { S, $, $$, toast, mi, esc } = X;

function pickFiles(accept) { // → the chosen files ([] when the picker is closed)
  return new Promise(res => { const i = document.createElement('input'); i.type = 'file'; i.accept = accept; i.onchange = () => res([...i.files]); i.addEventListener('cancel', () => res([])); i.click(); });
}
async function withDoc(run) { // a tool that works on an open document: with none open, the file is asked for first, then the tool starts
  if (!S.pages.length) {
    const files = await pickFiles('.pdf,application/pdf,.docx,image/*'); if (!files.length) return;
    await X.openFiles(files, false); if (!S.pages.length) return;
  }
  if (document.body.classList.contains('athome')) resume();
  await run();
}
const MENUS = {
  convert: () => [mi('PDF → Word (.docx)…', 'word', 'h_need', { d: { a: 'pdf2docx' } }), mi('Word → PDF…', 'word', 'docx2pdf'), '-',
    mi('PDF → PowerPoint (.pptx)…', 'slides', 'h_need', { d: { a: 'topptx' } }), mi('PDF → Excel (.xlsx)…', 'table', 'h_need', { d: { a: 'toxlsx' } }), '-', mi('PDF → images…', 'image', 'h_need', { d: { a: 'pdf2img' } }), mi('Images → PDF…', 'image', 'imgpages'), '-', mi('Export text (.txt)', 'text', 'h_need', { d: { a: 'exporttxt' } })],
  protect: () => [mi('Password protect…', 'lock', 'h_need', { d: { a: 'protect' } }), mi('Black out sensitive text', 'redact', 'h_tool', { d: { t: 'redact' } }), mi('Flatten (permanent redaction)…', 'layers', 'h_need', { d: { a: 'flatten' } }), mi('Remove hidden data…', 'shield', 'h_need', { d: { a: 'sanitize' } })],
  scan: () => [mi('Scan with the camera or photos', 'scan', 'scan'), mi('Make a PDF searchable (OCR)…', 'scan', 'h_need', { d: { a: 'ocrall' } }), '-',
    mi('Compare two PDFs…', 'swap', 'h_need', { d: { a: 'compare' } }), mi('Batch process many PDFs…', 'layers', 'batch')],
  office: () => [mi('PDF → Word (.docx)…', 'word', 'h_need', { d: { a: 'pdf2docx' } }), mi('PDF → PowerPoint (.pptx)…', 'slides', 'h_need', { d: { a: 'topptx' } }), mi('PDF → Excel (.xlsx)…', 'table', 'h_need', { d: { a: 'toxlsx' } }), '-',
    mi('Save every picture (ZIP)', 'image', 'h_need', { d: { a: 'extractimg' } }), mi('Export text (.txt)', 'text', 'h_need', { d: { a: 'exporttxt' } })],
  pages: () => [mi('Rearrange pages', 'grid', 'h_need', { d: { a: 'organize' } }), mi('Rotate all pages', 'rotate', 'h_need', { d: { a: 'pgrotall' } }), mi('Reverse the page order', 'reverse', 'h_need', { d: { a: 'pgreverse' } }), '-',
    mi('Remove blank pages…', 'trash', 'h_need', { d: { a: 'pgblank' } }), mi('Split every N pages…', 'split', 'h_need', { d: { a: 'pgsplit' } }), mi('Extract pages…', 'split', 'h_need', { d: { a: 'extract' } })],
  sign: () => [mi('Sign…', 'sign', 'h_need', { d: { a: 'sign' } }), mi("Stamp today's date", 'stamp', 'h_need', { d: { a: 'datestamp' } }), '-',
    mi('Highlight', 'marker', 'h_tool', { d: { t: 'marker' } }), mi('Underline text', 'underline', 'h_tool', { d: { t: 'uline' } }), mi('Strike through text', 'strike', 'h_tool', { d: { t: 'sline' } }), mi('Sticky note', 'note', 'h_tool', { d: { t: 'note' } }), mi('Arrows and shapes', 'arrow', 'h_tool', { d: { t: 'arrow' } })],
};

/* ───────── writing assistant ───────── */
const DRAFT = 'xd3.wadraft', ta = $('#waText');
try { ta.value = localStorage.getItem(DRAFT) || ''; } catch (e) { /* not remembered */ }
ta.addEventListener('input', () => { try { localStorage.setItem(DRAFT, ta.value); } catch (e) { /* not remembered */ } });
ta.addEventListener('keydown', e => { e.stopPropagation(); if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) X.ACT.h_wcheck(); }); // (typing here never triggers the editor's shortcuts)
async function toWriter(tab, fix) {
  const t = ta.value.replace(/\r\n?/g, '\n'); if (!t.trim()) { toast('Paste or type some text first — or open a file', 'info'); ta.focus(); return; }
  const Wr = X.writing; Wr.show(); await Wr.openWith(t.split('\n').map(text => ({ text, orig: text, issues: [] })), 'text', 'text'); Wr.setTab(tab);
  if (fix) await X.busy(async () => { const n = await Wr.fixIssues('all'); toast(n ? `${n} mistake${n > 1 ? 's' : ''} corrected — Ctrl+Z to undo, or check what is left` : 'No clear mistakes found', n ? '' : 'info'); }, 'Correcting…');
}

/* ───────── the tool list: js/flow.js owns it (X.flow.TOOLS / CATS); this short copy is used until it is there ───────── */
const need = a => b => withDoc(() => X.ACT[a](b)), useTool = t => () => withDoc(() => X.setTool(t)), act = a => b => X.ACT[a] && X.ACT[a](b);
const OWN_CATS = [{ id: 'organize', label: 'Organize', color: 'var(--c-organize)' }, { id: 'optimize', label: 'Optimize', color: 'var(--c-optimize)' }, { id: 'convert', label: 'Convert', color: 'var(--c-word)' },
  { id: 'edit', label: 'Edit', color: 'var(--c-edit)' }, { id: 'security', label: 'Sign & Security', color: 'var(--c-security)' }, { id: 'write', label: 'Write & Word', color: 'var(--c-write)' }];
const T = (id, title, desc, cat, icon, run, keywords) => ({ id, title, desc, cat, icon, run, keywords });
const OWN_TOOLS = [
  T('merge-pdf', 'Merge PDF', 'Combine PDFs, Word files and images into one PDF in the order you want.', 'organize', 'merge', act('mergedlg'), 'combine join'),
  T('split-pdf', 'Split PDF', 'Split by page ranges, every N pages, or into single pages (ZIP).', 'organize', 'split', need('pgsplit'), 'separate'),
  T('remove-pages', 'Remove pages', "Pick the pages you don't need and delete them.", 'organize', 'trash', need('organize'), 'delete'),
  T('extract-pages', 'Extract pages', 'Save only the pages you choose as a new PDF.', 'organize', 'split', need('extract')),
  T('organize-pdf', 'Organize PDF', 'Drag pages into order, rotate, duplicate, replace or insert blank pages.', 'organize', 'grid', need('organize'), 'rearrange order sort'),
  T('rotate-pdf', 'Rotate PDF', 'Rotate one page, chosen pages or the whole document.', 'organize', 'rotate', need('pgrotall')),
  T('remove-blank-pages', 'Remove blank pages', 'Find empty scanner pages and remove them in one click.', 'organize', 'trash', need('pgblank')),
  T('compress-pdf', 'Compress PDF', 'Make the PDF smaller while text stays sharp and selectable.', 'optimize', 'compress', need('compress'), 'reduce size shrink'),
  T('ocr-pdf', 'OCR PDF', 'Turn scans into searchable, selectable text — offline, 14 languages.', 'optimize', 'scan', need('ocrall'), 'searchable recognise'),
  T('batch', 'Batch process', 'Compress, OCR, watermark or number many PDFs at once.', 'optimize', 'layers', act('batch')),
  T('word-to-pdf', 'Word to PDF', 'Turn .docx files into PDFs that keep their fonts and layout.', 'convert', 'word', act('docx2pdf'), 'docx'),
  T('jpg-to-pdf', 'JPG to PDF', 'Turn photos and images into PDF pages.', 'convert', 'image', act('imgpages'), 'image png photo'),
  T('scan-to-pdf', 'Scan to PDF', 'Use your camera or photos to make a clean, cropped PDF.', 'convert', 'scan', act('scan'), 'camera'),
  T('create-pdf', 'Create PDF', 'Start a blank PDF and add text, images and pages.', 'convert', 'plus', act('newdoc'), 'new blank'),
  T('pdf-to-word', 'PDF to Word', 'Get an editable .docx that looks exactly like the PDF.', 'convert', 'word', need('pdf2docx'), 'docx'),
  T('pdf-to-powerpoint', 'PDF to PowerPoint', 'One editable slide for each page.', 'convert', 'slides', need('topptx'), 'pptx slides'),
  T('pdf-to-excel', 'PDF to Excel', 'Tables become rows and columns; numbers stay numbers.', 'convert', 'table', need('toxlsx'), 'xlsx spreadsheet'),
  T('pdf-to-jpg', 'PDF to JPG', 'Save every page as a PNG or JPG image.', 'convert', 'image', need('pdf2img'), 'png image'),
  T('extract-images', 'Extract images', 'Save every picture inside the PDF (ZIP).', 'convert', 'image', need('extractimg'), 'pictures'),
  T('pdf-to-text', 'PDF to Text', 'Export all the text as a plain .txt file.', 'convert', 'text', need('exporttxt'), 'txt'),
  T('edit-pdf', 'Edit PDF', 'Change existing text in its own font, add text, images and shapes.', 'edit', 'edittext', useTool('edittext')),
  T('fill-forms', 'Fill PDF forms', 'Type into form fields and tick boxes, then save.', 'edit', 'text', useTool('select'), 'form'),
  T('annotate-pdf', 'Annotate PDF', 'Highlight, underline, strike out, draw and add sticky notes.', 'edit', 'marker', useTool('marker'), 'highlight comment'),
  T('add-image', 'Add image', 'Place a picture or logo anywhere; remove its background.', 'edit', 'image', need('image'), 'logo picture'),
  T('add-watermark', 'Add watermark', 'Stamp text or an image over your pages, at any angle and opacity.', 'edit', 'drop'),
  T('page-numbers', 'Add page numbers', 'Number pages with your choice of position, style and start number.', 'edit', 'hash', need('pagenum')),
  T('header-footer', 'Header & footer', 'Add a title, date or page text to the top and bottom of every page.', 'edit', 'stroke', need('headfoot')),
  T('letterhead', 'Letterhead', 'Put your letterhead on every page — design one or use your own.', 'edit', 'stamp', () => withDoc(() => X.ACT.letterhead())),
  T('crop-pdf', 'Crop & resize', 'Trim margins or change the page size (A4, Letter…).', 'edit', 'crop', useTool('crop'), 'page size'),
  T('sign-pdf', 'Sign PDF', "Draw, type or upload your signature, plus stamps and today's date.", 'security', 'sign', need('sign'), 'signature'),
  T('protect-pdf', 'Protect PDF', 'Add a password with AES encryption.', 'security', 'lock', need('protect'), 'password encrypt'),
  T('unlock-pdf', 'Unlock PDF', 'Open a password-protected PDF and save a copy without the password.', 'security', 'key', act('open'), 'password remove'),
  T('redact-pdf', 'Redact PDF', 'Black out sensitive text and areas — removed from the file for good.', 'security', 'redact', useTool('redact'), 'black out'),
  T('flatten-pdf', 'Flatten PDF', 'Make edits, forms and annotations permanent and uneditable.', 'security', 'layers', need('flatten')),
  T('remove-hidden-data', 'Remove hidden data', 'Strip metadata, scripts, attachments and comments.', 'security', 'shield', need('sanitize'), 'privacy metadata clean'),
  T('compare-pdf', 'Compare PDF', 'See what changed between two versions, side by side.', 'security', 'swap', need('compare'), 'diff'),
  T('word-editor', 'Word editor', 'Edit .docx files with their own fonts and styles — Auto Format and Fix Formatting.', 'write', 'word', act('dx'), 'docx'),
  T('word-templates', 'Word templates', 'Letter, résumé, invoice, report and memo templates with your letterhead.', 'write', 'stamp', act('dxnew'), 'resume cv invoice'),
  T('spelling-grammar', 'Spelling & grammar', 'Find and fix mistakes in PDFs, Word and text — offline.', 'write', 'spell', act('writer'), 'proofread'),
  T('ai-humanizer', 'AI Humanizer', 'Make stiff or AI-sounding text read naturally.', 'write', 'sparkle', () => { X.writing.show(); X.writing.setTab('human'); }, 'rewrite'),
];
OWN_TOOLS.find(t => t.id === 'add-watermark').run = need('watermark');
const live = () => !!(X.flow && Array.isArray(X.flow.TOOLS) && X.flow.TOOLS.length);
const tools = () => live() ? X.flow.TOOLS : OWN_TOOLS;
const cats = () => (live() && Array.isArray(X.flow.CATS) && X.flow.CATS.length) ? X.flow.CATS : OWN_CATS;
const byId = id => tools().find(t => t.id === id);
const RECENT = 'xd3.recentTools';
const ownRecent = () => { try { return JSON.parse(localStorage.getItem(RECENT) || '[]'); } catch (e) { return []; } };
function noteTool(id) { try { localStorage.setItem(RECENT, JSON.stringify([id, ...ownRecent().filter(x => x !== id)].slice(0, 12))); } catch (e) { /* not remembered */ } }
const recentTools = () => { let ids = []; try { ids = live() && typeof X.flow.recent === 'function' ? X.flow.recent() || [] : ownRecent(); } catch (e) { ids = ownRecent(); } return ids.map(x => typeof x === 'string' ? byId(x) : x).filter(t => t && t.id).slice(0, 6); };
function openTool(id, b) {
  closeMega(); closeSheet();
  if (live() && typeof X.flow.open === 'function') return X.flow.open(id); // (the tool screen keeps its own recent list)
  const t = OWN_TOOLS.find(x => x.id === id); if (!t) return;
  noteTool(id); renderRecentTools(); t.run(b);
}

/* colour and tag of a tool's icon tile: convert tools take the other format's colour */
const cssCol = c => !c ? '' : /^--/.test(c) ? `var(${c})` : c;
const FMT = [[/word|docx/i, 'var(--c-word)', 'DOC'], [/powerpoint|pptx|slide/i, 'var(--c-ppt)', 'PPT'], [/excel|xlsx|sheet/i, 'var(--c-excel)', 'XLS'], [/jpg|jpeg|png|image|picture|photo/i, 'var(--c-image)', 'JPG'], [/text|txt/i, 'var(--c-security)', 'TXT']];
function tile(t, sm) {
  const c = cats().find(x => x.id === t.cat) || {}; let col = cssCol(t.color) || cssCol(c.color) || 'var(--c-edit)', tag = t.tag || '';
  if (!t.color && /conv/i.test(t.cat || '')) {
    const s = t.title || '', f = FMT.find(([re]) => re.test(s));
    if (f) col = f[1];
    if (!tag) tag = /to pdf$/i.test(s) ? 'PDF' : /^pdf to/i.test(s) && f ? f[2] : '';
  }
  return `<span class="tico${sm ? ' sm' : ''}" style="--cat:${esc(col)}"><i data-icon="${esc(t.icon || 'pdf')}"></i>${tag ? `<b class="ttag">${esc(tag)}</b>` : ''}</span>`;
}
const card = (t, i) => `<button class="tcard" data-act="h_flow" data-t="${esc(t.id)}" data-cat="${esc(t.cat || '')}" style="animation-delay:${Math.min(i, 14) * 18}ms">${tile(t)}<span class="tt">${esc(t.title)}</span><span class="td">${esc(t.desc || '')}</span>${t.badge ? `<b class="tbadge">${esc(t.badge)}</b>` : ''}</button>`;
const item = t => `<button class="mi" data-act="h_flow" data-t="${esc(t.id)}" title="${esc(t.desc || '')}">${tile(t, true)}<span>${esc(t.title)}</span></button>`;
const isToPdf = t => t.group ? t.group === 'to' : /to pdf$|^scan|^create|^jpg|^image/i.test(t.title || '');

/* ───────── header: Merge · Split · Compress · Convert ▾ · All PDF tools ▾ ───────── */
const nav = $('#hnav'), megaC = $('#megaConvert'), megaA = $('#megaAll'), sheet = $('#navSheet'), hamb = $('#bHamb');
const find = re => tools().find(t => re.test(t.title || ''));
function renderNav() {
  const quick = [find(/^merge/i), find(/^split/i), find(/^compress/i)].filter(Boolean);
  nav.innerHTML = quick.map(t => `<button class="hn" data-act="h_flow" data-t="${esc(t.id)}" title="${esc(t.desc || '')}">${esc(t.title)}</button>`).join('')
    + '<button class="hn" data-mega="megaConvert" aria-haspopup="true" aria-expanded="false">Convert PDF<b class="chev"></b></button><button class="hn" data-mega="megaAll" aria-haspopup="true" aria-expanded="false">All PDF tools<b class="chev"></b></button>';
  const conv = tools().filter(t => /conv/i.test(t.cat || ''));
  megaC.innerHTML = `<div class="mcol"><div class="mcolh">Convert to PDF</div>${conv.filter(isToPdf).map(item).join('')}</div><div class="mcol"><div class="mcolh">Convert from PDF</div>${conv.filter(t => !isToPdf(t)).map(item).join('')}</div>`;
  const cs = cats().filter(c => tools().some(t => t.cat === c.id));
  megaA.style.setProperty('--cols', Math.max(1, cs.length));
  megaA.innerHTML = cs.map(c => `<div class="mcol"><div class="mcolh">${esc(c.label)}</div>${tools().filter(t => t.cat === c.id).map(item).join('')}</div>`).join('') + `<div class="mfoot">${extras().join('')}</div>`;
  sheet.innerHTML = cs.map((c, i) => `<details${i === 0 ? ' open' : ''}><summary><span class="cdot" style="--cat:${esc(cssCol(c.color) || 'var(--c-edit)')}"></span>${esc(c.label)}</summary>${tools().filter(t => t.cat === c.id).map(item).join('')}</details>`).join('') + `<div class="nsfoot">${extras().join('')}</div>`;
  X.icons(nav); X.icons(megaC); X.icons(megaA); X.icons(sheet);
}
// the Windows download: not inside the Windows app, and not on a site without the files (download.js hides the home button then)
const heroDl = $('#welcome [data-act="dlwin"]');
const hasAct = a => !!X.ACT[a] && (a !== 'dlwin' || (!document.documentElement.classList.contains('desk') && !(heroDl && heroDl.hidden)));
if (heroDl) new MutationObserver(() => { renderNav(); renderFooter(); }).observe(heroDl, { attributes: true, attributeFilter: ['hidden'] });
const extras = () => [['batch', 'layers', 'Batch process'], ['dlwin', 'windows', 'Windows app'], ['shortcuts', 'key', 'Keyboard shortcuts'], ['palette', 'search', 'Search every command']]
  .filter(([a]) => hasAct(a)).map(([a, ic, l]) => `<button class="mi" data-act="${a}"><i data-icon="${ic}"></i><span>${l}</span></button>`);
let megaT = 0, openId = null, openedAt = 0;
const trig = id => $(`.hn[data-mega="${id}"]`, nav);
function openMega(id) {
  const t = trig(id), m = $('#' + id); if (!t || !m) return; if (openId && openId !== id) closeMega();
  m.hidden = false; t.classList.add('open'); t.setAttribute('aria-expanded', 'true'); openId = id; openedAt = Date.now();
  if (id === 'megaConvert') m.style.left = Math.max(12, Math.min(t.offsetLeft - 16, innerWidth - m.offsetWidth - 12)) + 'px';
}
function closeMega() { clearTimeout(megaT); if (!openId) return; const t = trig(openId), m = $('#' + openId); if (m) m.hidden = true; if (t) { t.classList.remove('open'); t.setAttribute('aria-expanded', 'false'); } openId = null; }
const canHover = () => !(window.matchMedia && matchMedia('(hover: none)').matches);
document.addEventListener('mouseover', e => {
  if (!canHover() || !e.target.closest) return;
  const t = e.target.closest('.hn[data-mega]'), inMega = e.target.closest('.mega'); clearTimeout(megaT);
  if (t) megaT = setTimeout(() => openMega(t.dataset.mega), openId ? 60 : 150);
  else if (!inMega && openId) megaT = setTimeout(closeMega, 220);
});
document.addEventListener('click', e => {
  const t = e.target.closest && e.target.closest('.hn[data-mega]');
  if (t) { const id = t.dataset.mega; if (openId === id && Date.now() - openedAt > 500) closeMega(); else openMega(id); return; }
  if (openId && !(e.target.closest && e.target.closest('.mega .mcolh'))) closeMega();
});
nav.addEventListener('keydown', e => {
  const t = e.target.closest('.hn[data-mega]'); if (!t || e.key !== 'ArrowDown') return;
  e.preventDefault(); openMega(t.dataset.mega); const f = $('#' + t.dataset.mega + ' button'); if (f) f.focus();
});
for (const m of [megaC, megaA]) m.addEventListener('keydown', e => { // arrows move through the items, Esc goes back to the menu name
  const bs = $$('button', m), i = bs.indexOf(document.activeElement);
  if (e.key === 'ArrowDown' || e.key === 'ArrowUp') { e.preventDefault(); const n = bs[(i + (e.key === 'ArrowDown' ? 1 : -1) + bs.length) % bs.length]; if (n) n.focus(); }
});
window.addEventListener('keydown', e => {
  if (e.key !== 'Escape') return;
  if (openId) { const t = trig(openId); closeMega(); if (t) t.focus(); e.stopImmediatePropagation(); }
  else if (!sheet.hidden) { closeSheet(); e.stopImmediatePropagation(); }
}, true);
/* phones and small windows: the menu becomes a full-screen sheet */
function closeSheet() { if (sheet.hidden) return; sheet.hidden = true; hamb.setAttribute('aria-expanded', 'false'); hamb.classList.remove('active'); }
hamb.addEventListener('click', () => { const open = sheet.hidden; sheet.hidden = !open; hamb.setAttribute('aria-expanded', String(open)); hamb.classList.toggle('active', open); if (open) sheet.scrollTop = 0; });
sheet.addEventListener('click', e => { if (e.target.closest('[data-act]')) closeSheet(); });
window.addEventListener('resize', () => { if (innerWidth > 992) closeSheet(); if (openId) closeMega(); });
const bodyCl = document.body.classList, leftHome = () => (bodyCl.contains('hasdoc') && !bodyCl.contains('athome')) || bodyCl.contains('dxopen');
new MutationObserver(() => { if (leftHome()) { closeSheet(); closeMega(); } }).observe(document.body, { attributes: true, attributeFilter: ['class'] });
document.addEventListener('xd3:opened', () => { closeSheet(); closeMega(); });

/* ───────── the tool grid: category chips + search ───────── */
const grid = $('#toolGrid'), chipsEl = $('#toolChips'), tf = $('#toolFind'), none = $('#toolNone');
let cat = 'all';
function renderChips() {
  const cs = cats().filter(c => tools().some(t => t.cat === c.id));
  if (cat !== 'all' && !cs.some(c => c.id === cat)) cat = 'all';
  chipsEl.innerHTML = [{ id: 'all', label: 'All' }, ...cs].map(c => `<button class="chip${c.id === cat ? ' active' : ''}" role="tab" aria-selected="${c.id === cat}" data-cat="${esc(c.id)}">${esc(c.label)}</button>`).join('');
}
function renderGrid() {
  const q = tf.value.trim().toLowerCase().split(/\s+/).filter(Boolean), cl = id => (cats().find(c => c.id === id) || {}).label || '';
  let list = tools().filter(t => {
    if (q.length) { const hay = [t.title, t.desc, cl(t.cat), ...[].concat(t.keywords || [])].join(' ').toLowerCase(); return q.every(w => hay.includes(w)); }
    return cat === 'all' || t.cat === cat;
  });
  if (q.length) { // best match first (Enter opens the first card): the tool's name, then its keywords, then its description
    const score = t => {
      const title = t.title.toLowerCase(), tw = title.split(/[^a-z0-9]+/), kw = [].concat(t.keywords || []).join(' ').toLowerCase(), kws = kw.split(/[^a-z0-9]+/);
      let s = title.startsWith(q.join(' ')) ? 200 : 0;
      q.forEach(w => { s += tw.includes(w) ? 60 : tw.some(x => x.startsWith(w)) ? 45 : title.includes(w) ? 30 : kws.includes(w) ? 20 : kws.some(x => x.startsWith(w)) ? 12 : kw.includes(w) ? 8 : 0; });
      return s;
    };
    list = list.map((t, i) => ({ t, i, s: score(t) })).sort((a, b) => b.s - a.s || a.i - b.i).map(o => o.t);
  }
  grid.innerHTML = list.map(card).join(''); X.icons(grid);
  none.hidden = !!list.length;
  if (!list.length) none.innerHTML = `No tool matches <b>“${esc(tf.value.trim())}”</b> — try “compress”, “sign” or “excel”, or press <kbd>Ctrl K</kbd> to search every command.`;
}
chipsEl.addEventListener('click', e => {
  const b = e.target.closest('.chip'); if (!b) return; cat = b.dataset.cat;
  if (tf.value) tf.value = ''; $$('.chip', chipsEl).forEach(c => { const on = c === b; c.classList.toggle('active', on); c.setAttribute('aria-selected', String(on)); }); renderGrid();
});
tf.addEventListener('keydown', e => { e.stopPropagation(); if (e.key === 'Enter') { const first = $('.tcard', grid); if (first) first.click(); } if (e.key === 'Escape') { tf.value = ''; renderGrid(); } });
tf.addEventListener('input', () => {
  if (tf.value.trim() && cat !== 'all') { cat = 'all'; $$('.chip', chipsEl).forEach(c => { c.classList.toggle('active', c.dataset.cat === 'all'); c.setAttribute('aria-selected', String(c.dataset.cat === 'all')); }); }
  renderGrid();
});
function renderRecentTools() {
  const host = $('#recentTools'); if (!host) return; const list = recentTools();
  host.hidden = !list.length; host.innerHTML = list.length ? `<h3><i data-icon="clock"></i>Recently used tools</h3><div class="tgrid mini">${list.map(card).join('')}</div>` : '';
  X.icons(host);
}
/* first visit: three steps, dismissed for good */
const TIP = 'xd3.hometip';
function renderTip() {
  const h = $('#homeTip'); let seen = false; try { seen = localStorage.getItem(TIP) === '1'; } catch (e) { seen = true; }
  h.hidden = seen; if (seen) return;
  h.innerHTML = '<ol><li><b>1</b>Pick a tool</li><li><b>2</b>Add your files</li><li><b>3</b>Download the result — it never leaves this device</li></ol><button class="btn line sm" type="button">Got it</button>';
  $('button', h).onclick = () => { try { localStorage.setItem(TIP, '1'); } catch (e) { /* shown again next time */ } h.hidden = true; };
}
function renderFooter() {
  const pick = re => tools().filter(t => re.test(t.cat || '')).slice(0, 7).map(t => `<button class="fl" data-act="h_flow" data-t="${esc(t.id)}">${esc(t.title)}</button>`).join('');
  const app = [['dx', 'Word editor'], ['dlwin', 'Windows app'], ['shortcuts', 'Keyboard shortcuts'], ['palette', 'Search every command'], ['about', 'About XD3 PDF']].filter(([a]) => hasAct(a)).map(([a, l]) => `<button class="fl" data-act="${a}">${l}</button>`).join('');
  $('#footCols').innerHTML = `<div><h4>Organize</h4>${pick(/organ|optim/i)}</div><div><h4>Convert</h4>${pick(/conv/i)}</div><div><h4>Edit &amp; Sign</h4>${pick(/edit/i)}</div><div><h4>Security</h4>${pick(/secur|sign|protect/i)}</div><div><h4>App</h4>${app}</div>`;
}
function renderTools() { renderNav(); renderChips(); renderGrid(); renderRecentTools(); renderFooter(); markNav(); }

Object.assign(X.ACT, {
  h_need: b => withDoc(() => X.ACT[b.dataset.a](b)),
  h_tool: b => withDoc(() => X.setTool(b.dataset.t)),
  h_menu: b => { const r = b.getBoundingClientRect(); X.showCtx({ clientX: r.left + 14, clientY: r.bottom - 8, target: b }, MENUS[b.dataset.m]()); },
  h_flow: b => openTool(b.dataset.t, b),
  h_wcheck: () => toWriter('issues'),
  h_wfix: () => toWriter('issues', true),
  h_whuman: () => toWriter('human'),
  h_wfile: async () => { const [f] = await pickFiles('.docx,.txt,.md,text/plain'); if (!f) return; X.writing.show(); await X.writing.openFile(f); },
  h_wpdf: () => withDoc(async () => { X.writing.show(); await X.writing.ACTS.pdf(); }),
});
['h_need', 'h_tool', 'h_menu', 'h_flow', 'h_wcheck', 'h_wfix', 'h_whuman', 'h_wfile', 'h_wpdf'].forEach(a => X.NO_DOC.add(a));

/* ───────── Home: the start screen, reachable at any time; open work waits under "Continue" ───────── */
const ago = t => { const d = Math.round((Date.now() - t) / 60000); return d < 2 ? 'just now' : d < 60 ? d + ' min ago' : d < 1440 ? Math.round(d / 60) + ' h ago' : d < 2880 ? 'yesterday' : Math.round(d / 1440) + ' days ago'; };
const mb = n => n < 1048576 ? Math.max(1, Math.round(n / 1024)) + ' KB' : (n / 1048576).toFixed(1) + ' MB';
function goHome() {
  if (X.commitEdit) X.commitEdit(); if (X.hideCtx) X.hideCtx();
  if (X.docx && X.docx.isOpen()) X.docx.leave(); if (X.writing && X.writing.W.open) X.writing.close();
  document.body.classList.add('athome'); render(); const w = $('#welcome'); if (w) w.scrollTop = 0;
}
function resume() { closeSheet(); closeMega(); document.body.classList.remove('athome'); X.updateUI(); if (X.fitZoom && X.S.pages.length) X.layout(); }
function renderContinue() {
  const host = $('#continue'); if (!host) return; const cards = [];
  if (S.pages.length) cards.push(`<button class="cont" data-act="resume"><i data-icon="pdf"></i><span><b>${esc(S.fileName)}${S.dirty ? ' •' : ''}</b><small>PDF · ${S.pages.length} page${S.pages.length > 1 ? 's' : ''}${S.dirty ? ' · not saved yet' : ''}</small></span><em>Continue editing →</em></button>`);
  if (X.docx && X.docx.D.zip && !X.docx.isOpen()) cards.push(`<button class="cont word" data-act="dx"><i data-icon="word"></i><span><b>${esc(X.docx.name())}${X.docx.D.dirty ? ' •' : ''}</b><small>Word document${X.docx.D.dirty ? ' · not saved yet' : ''}</small></span><em>Continue editing →</em></button>`);
  host.hidden = !cards.length; host.innerHTML = cards.length ? `<h3><i data-icon="undo"></i>Continue where you left off</h3><div class="conts">${cards.join('')}</div>` : ''; X.icons(host);
}

/* recent files: the last PDFs and Word files opened on this device, with a picture of the first page */
const RECENT_MAX = 12, RECENT_BYTES = 200e6, RECENT_FILE = 40e6;
async function thumbOf() { try { const c = await X.renderPageImage(S.pages[0], 120 / Math.max(...X.rdims(S.pages[0])), { annots: false }); return c.toDataURL('image/jpeg', 0.7); } catch (e) { return ''; } }
async function noteRecent(files) {
  try {
    const f = [...files].find(x => X.isPdf(x) || X.isDocx(x)); if (!f || f.size > RECENT_FILE) return;
    const kind = X.isPdf(f) ? 'pdf' : 'docx', bytes = new Uint8Array(await f.arrayBuffer()), thumb = kind === 'pdf' && S.pages.length ? await thumbOf() : '';
    let list = (await X.DB.get('recent')) || [];
    list = list.filter(r => r.name !== f.name); list.unshift({ name: f.name, t: Date.now(), bytes, kind, thumb, pages: kind === 'pdf' ? S.pages.length : 0 });
    let total = 0; list = list.filter(r => (total += r.bytes.length) <= RECENT_BYTES).slice(0, RECENT_MAX);
    await X.DB.put('recent', list); renderRecent(list);
  } catch (e) { /* storage unavailable — the list is a convenience */ }
}
async function renderRecent(list) {
  const host = $('#recent'); if (!host) return;
  try { list = list || (await X.DB.get('recent')) || []; } catch (e) { list = []; }
  host.hidden = !list.length; if (!list.length) { host.innerHTML = ''; return; }
  host.innerHTML = `<h3><i data-icon="clock"></i>Recent files<button class="lnk" data-r="clear" title="Forget the list (the files themselves are not touched)">Clear list</button></h3><div class="rgrid">` + list.map((r, i) => {
    const kind = r.kind || 'pdf';
    return `<div class="rcard" data-r="${i}" role="button" tabindex="0" title="Open ${esc(r.name)} again"><div class="rthumb ${kind}">${r.thumb ? `<img src="${r.thumb}" alt="">` : `<i data-icon="${kind === 'docx' ? 'word' : 'pdf'}"></i>`}<span class="rkind">${kind === 'docx' ? 'DOCX' : 'PDF'}</span></div>
      <b>${esc(r.name)}</b><small>${r.pages ? r.pages + ' p · ' : ''}${mb(r.bytes.length)} · ${ago(r.t)}</small><button class="rx" data-rx="${i}" title="Remove from the list">×</button></div>`;
  }).join('') + '</div>';
  X.icons(host);
}
async function openRecent(i) {
  const list = (await X.DB.get('recent')) || [], r = list[i]; if (!r) return;
  if ((r.kind || 'pdf') === 'docx' && X.docx) { document.body.classList.remove('athome'); return X.docx.openFile(new File([r.bytes], r.name, { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' })); }
  X.openFiles([new File([r.bytes], r.name, { type: 'application/pdf' })], false);
}
$('#welcome').addEventListener('click', async e => {
  const x = e.target.closest('[data-rx]'); if (x) { e.stopPropagation(); const list = (await X.DB.get('recent')) || []; list.splice(+x.dataset.rx, 1); await X.DB.put('recent', list); return renderRecent(list); }
  const b = e.target.closest('[data-r]'); if (!b) return;
  if (b.dataset.r === 'clear') { await X.DB.put('recent', []); return renderRecent([]); }
  openRecent(+b.dataset.r);
});
$('#welcome').addEventListener('keydown', e => { const c = e.target.closest('.rcard'); if (c && (e.key === 'Enter' || e.key === ' ')) { e.preventDefault(); openRecent(+c.dataset.r); } });
$('.brand').addEventListener('keydown', e => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); goHome(); } });

function render() { renderContinue(); renderRecent(); renderRecentTools(); if (nav.firstChild) markNav(); }
Object.assign(X.ACT, { home: goHome, resume });
['home', 'resume'].forEach(a => X.NO_DOC.add(a));
X.noteRecent = noteRecent;
renderTools(); renderTip(); render();
// the full tool list (js/flow.js) and the Windows download link load after this file
document.addEventListener('DOMContentLoaded', renderTools);
document.addEventListener('xd3:flowready', renderTools);
// the header marks the tool that is open (#/merge-pdf …); the recent row is fresh when the tool closes
function markNav() { const id = (/^#\/([\w-]+)/.exec(location.hash) || [])[1]; $$('.hn[data-t]', nav).forEach(b => b.classList.toggle('active', b.dataset.t === id)); if (!id) renderRecentTools(); }
window.addEventListener('hashchange', markNav);
X.home = { MENUS, withDoc, render, goHome, resume, noteRecent, tools, cats, openTool, renderTools };
})();
