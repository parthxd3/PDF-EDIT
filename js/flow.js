/* XD3 PDF Editor — one tool per screen, iLovePDF style: landing → files and options → processing → result.
   Routes are #/<tool-id>. A tool runs on its own copy of the files (see sandbox()): the document open in the editor is not touched
   unless the user picks "Open in editor". Tools that are really editor modes (Edit, Sign, Redact…) open the editor with the file. */
(() => {
'use strict';
const X = window.XD3, { S, esc } = X;
const $ = (s, r = document) => r.querySelector(s), $$ = (s, r = document) => [...r.querySelectorAll(s)];
const PJ = { cMapUrl: 'lib/cmaps/', cMapPacked: true, standardFontDataUrl: 'lib/standard_fonts/', isEvalSupported: false };
const NARROW = matchMedia('(max-width: 840px)');
const sizeOf = n => n < 1048576 ? Math.max(1, Math.round(n / 1024)) + ' KB' : (n / 1048576).toFixed(1) + ' MB';
const plural = (n, w, ws) => `${n} ${n === 1 ? w : ws || w + 's'}`;
const baseOf = n => String(n || 'document').replace(/\.[^.\\/]+$/, '') || 'document';
const kindOf = f => X.isPdf(f) ? 'pdf' : X.isDocx(f) ? 'docx' : /^image\//.test(f.type) || /\.(png|jpe?g|gif|webp|bmp|svg)$/i.test(f.name) ? 'img' : /\.(txt|md)$/i.test(f.name) || f.type === 'text/plain' ? 'txt' : 'other';
const MAX_BYTES = 500e6;
const store = { get(k, d) { try { const v = localStorage.getItem(k); return v == null ? d : v; } catch (e) { return d; } }, set(k, v) { try { localStorage.setItem(k, v); } catch (e) { /* not remembered */ } } };

/* ───────── categories and tools ───────── */
const CATS = [
  { id: 'organize', label: 'Organize', color: '#ee6c4d' },
  { id: 'optimize', label: 'Optimize', color: '#8fbc5d' },
  { id: 'convert', label: 'Convert', color: '#295795' },
  { id: 'edit', label: 'Edit', color: '#ab6993' },
  { id: 'security', label: 'Sign & Security', color: '#4a7aab' },
  { id: 'write', label: 'Write & Word', color: '#7253e2' },
];
const GROUPS = [{ id: 'to', cat: 'convert', label: 'Convert to PDF' }, { id: 'from', cat: 'convert', label: 'Convert from PDF' }];
const FMT = { word: '#295795', ppt: '#d04526', excel: '#2e7237', image: '#b7a001', other: '#4a7aab' };
const TOOLS = [], byId = {};
const tool = o => { const t = Object.assign({ kinds: ['pdf'], min: 1, max: 1, mode: 'files', options: [], next: [], keywords: '' }, o); t.color = t.color || CATS.find(c => c.id === t.cat).color; TOOLS.push(t); byId[t.id] = t; };
const ACCEPT = { pdf: '.pdf,application/pdf', docx: '.docx', img: 'image/*', txt: '.txt,.md,text/plain' };
const NOUN = { pdf: ['PDF file', 'PDF files', 'PDFs'], docx: ['Word file', 'Word files', 'Word files'], img: ['image', 'images', 'images'], txt: ['text file', 'text files', 'text files'] };
const accept = t => t.kinds.map(k => ACCEPT[k]).join(',');
const pickLabel = t => t.pick || (t.kinds.length > 1 && !(t.kinds.length === 2 && t.kinds.includes('txt')) ? (t.max > 1 ? 'Select files' : 'Select a file') : 'Select ' + NOUN[t.kinds[0]][t.max > 1 ? 1 : 0]);
const dropLabel = t => 'or drop ' + (t.kinds.length > 1 ? 'files' : NOUN[t.kinds[0]][2]) + ' here';
const thing = t => (t.kinds.length > 1 ? 'files' : NOUN[t.kinds[0]][1]);

// shared option pieces
const RANGE = (label = 'Pages (blank = all)') => ({ t: 'text', k: 'range', label, ph: 'e.g. 1-3, 5, 8-10' });
const QUALITY = [{ v: '110|0.6', title: 'Small file', desc: 'Screen quality, 110 dpi' }, { v: '150|0.72', title: 'Balanced', desc: 'Good for most uses, 150 dpi', badge: 'Recommended' }, { v: '200|0.85', title: 'High quality', desc: 'For printing, 200 dpi' }];
const POS6 = [['bc', 'Bottom centre'], ['br', 'Bottom right'], ['bl', 'Bottom left'], ['tc', 'Top centre'], ['tr', 'Top right'], ['tl', 'Top left']];
const ocrLangs = () => { const s = $('#pOcrLang'); const l = s ? [...s.options].map(o => [o.value, o.textContent]) : []; return l.length ? l : [['eng', 'English']]; };
const slim = it => ({ str: it.str, x: +it.x.toFixed(2), y: +it.y.toFixed(2), w: +it.w.toFixed(2), h: +it.h.toFixed(2), base: +it.base.toFixed(2), size: +it.size.toFixed(2), ...(it.rot ? { rot: it.rot } : {}) });
const info = msg => Object.assign(new Error(msg), { info: true });
const fail = (msg, extra) => Object.assign(new Error(msg), { friendly: true }, extra);
const pdfRes = (bytes, name, o) => Object.assign({ blob: new Blob([bytes], { type: 'application/pdf' }), name, kind: 'pdf' }, o);
async function pack(outs, zipName, mime = 'application/pdf') { // one result → that file; several → one ZIP
  if (outs.length === 1) { const o = outs[0]; return { blob: o.data instanceof Blob ? o.data : new Blob([o.data], { type: mime }), name: o.name, kind: mime === 'application/pdf' ? 'pdf' : 'file' }; }
  const JSZip = await X.loadLib('jszip'), zip = new JSZip(), used = new Set();
  for (const o of outs) { let n = o.name, k = 2; while (used.has(n.toLowerCase())) n = o.name.replace(/(\.\w+)$/, ` (${k++})$1`); used.add(n.toLowerCase()); zip.file(n, o.data); }
  return { blob: await zip.generateAsync({ type: 'blob', compression: 'DEFLATE', mimeType: 'application/zip' }), name: zipName, kind: 'zip', inner: mime === 'application/pdf' ? 'pdf' : /^image\//.test(mime) ? 'img' : 'file' };
}
const pickRange = (pages, txt) => { if (!txt || !txt.trim()) return pages; const idx = X.parseRange(txt, pages.length); if (!idx || !idx.length) throw fail(`That page range isn't valid. Use page numbers 1–${pages.length}, e.g. 1-3, 5.`); return idx.map(i => pages[i]); };
function stamp(p, o, pos, m) { const b = X.bbox(o); o.x = pos[1] === 'c' ? (p.w - b.w) / 2 : pos[1] === 'r' ? p.w - m - b.w : m; o.y = pos[0] === 'b' ? p.h - m - b.h : m; o.align = { l: 'left', c: 'center', r: 'right' }[pos[1]]; p.objs.push(o); }
const toBlob = (c, mime, q) => new Promise(res => c.toBlob(res, mime, q));
function wmPlace(o, W, H, iw, ih) { // where the watermark goes on a W×H page (also drawn as the preview)
  const rot = ((+o.angle % 360) + 360) % 360, k = +o.size || 0.8, cs = Math.abs(Math.cos(rot * Math.PI / 180)), sn = Math.abs(Math.sin(rot * Math.PI / 180));
  const f = Math.min(W * k / (iw * cs + ih * sn), H * k / (iw * sn + ih * cs)), w = iw * f, h = ih * f, y = o.pos === 't' ? Math.max(18, H * 0.08) : o.pos === 'b' ? H - h - Math.max(18, H * 0.08) : (H - h) / 2;
  return { x: (W - w) / 2, y, w, h, rot, opacity: Math.min(1, Math.max(0.02, +o.op || 0.18)) };
}

/* Organize */
tool({ id: 'merge-pdf', title: 'Merge PDF', desc: 'Combine PDFs, Word files and images into one PDF in the order you want.', cat: 'organize', icon: 'merge', keywords: 'combine join append',
  kinds: ['pdf', 'docx', 'img'], min: 2, max: 60, verb: 'Merge PDF', ing: 'Merging PDFs…', side: 'Merge PDF', sort: true,
  options: [{ t: 'info', text: 'To change the order of your files, drag and drop them as you want.' }],
  next: ['compress-pdf', 'organize-pdf', 'add-page-numbers', 'sign-pdf', 'protect-pdf', 'pdf-to-word'],
  run: async c => { const pages = await c.load(); S.fileName = 'merged.pdf'; c.status(`Joining ${plural(pages.length, 'page')}…`); const b = await X.buildPdf(pages);
    return pdfRes(b, 'merged.pdf', { head: 'PDFs have been merged!', sum: `${plural(c.items.length, 'file')} · ${plural(pages.length, 'page')} · ${sizeOf(b.length)}`, dl: 'Download merged PDF' }); } });
tool({ id: 'split-pdf', title: 'Split PDF', desc: 'Split by page ranges, every N pages, or into single pages (ZIP).', cat: 'organize', icon: 'split', keywords: 'separate divide cut',
  verb: 'Split PDF', ing: 'Splitting PDF…', side: 'Split',
  options: [{ t: 'list', k: 'how', def: 'range', items: [{ v: 'range', title: 'Split by range', desc: 'Each range becomes its own PDF' }, { v: 'every', title: 'Every N pages', desc: 'Equal parts of N pages' }, { v: 'single', title: 'Extract all pages', desc: 'Every page as its own PDF' }] },
    { t: 'text', k: 'ranges', label: 'Ranges', ph: 'e.g. 1-3, 4-8, 9', when: o => o.how === 'range' },
    { t: 'check', k: 'one', label: 'Merge all ranges into one PDF', when: o => o.how === 'range' },
    { t: 'num', k: 'n', label: 'Pages in each part', def: 2, min: 1, max: 9999, when: o => o.how === 'every' }],
  next: ['compress-pdf', 'merge-pdf', 'organize-pdf'],
  run: async c => {
    const pages = await c.load(), n = pages.length, o = c.o, base = c.base; let groups;
    if (o.how === 'single') groups = pages.map(p => [p]);
    else if (o.how === 'every') { const k = Math.max(1, +o.n || 1); if (k >= n) throw info(`This PDF has ${plural(n, 'page')} — choose fewer pages per part.`); groups = []; for (let i = 0; i < n; i += k) groups.push(pages.slice(i, i + k)); }
    else {
      const parts = String(o.ranges || '').split(',').map(s => s.trim()).filter(Boolean); if (!parts.length) throw fail('Type the page ranges to split at, e.g. 1-3, 4-8.');
      groups = parts.map(r => { const idx = X.parseRange(r, n); if (!idx) throw fail(`"${r}" isn't a valid range. Use page numbers 1–${n}.`); return idx.map(i => pages[i]); });
      if (o.one) { const b = await X.buildPdf(groups.flat()); return pdfRes(b, base + '-ranges.pdf', { head: 'Your ranges are in one PDF!', sum: `${plural(groups.flat().length, 'page')} · ${sizeOf(b.length)}`, dl: 'Download PDF' }); }
    }
    if (groups.length === 1) { const b = await X.buildPdf(groups[0]); return pdfRes(b, base + '-part.pdf', { head: 'Your PDF has been split!', sum: `${plural(groups[0].length, 'page')} · ${sizeOf(b.length)}`, dl: 'Download PDF' }); }
    const r = await X.tools.splitParts({ pages, groups, progress: c.progress, check: c.check });
    return { blob: r.blob, name: base + '-split.zip', kind: 'zip', inner: 'pdf', head: 'Your PDF has been split!', sum: `${plural(r.parts, 'PDF')} in a ZIP · ${sizeOf(r.blob.size)}`, dl: 'Download ZIP' };
  } });
tool({ id: 'remove-pages', title: 'Remove pages', desc: "Pick the pages you don't need and delete them.", cat: 'organize', icon: 'trash', keywords: 'delete pages',
  mode: 'pages', verb: 'Remove pages', ing: 'Removing pages…', side: 'Remove pages', selVerb: 'remove',
  options: [{ t: 'info', text: 'Click the pages you want to remove.' }, { t: 'pages' }], next: ['compress-pdf', 'merge-pdf', 'add-page-numbers'],
  ready: F => F.pages.some(p => p.sel) ? '' : 'Select the pages to remove',
  run: async c => { const all = await c.load(), keep = all.filter((p, i) => !c.sel[i]); if (!keep.length) throw fail("You can't remove every page — leave at least one."); const b = await X.buildPdf(keep);
    return pdfRes(b, c.base + '-pages-removed.pdf', { head: 'Pages have been removed!', sum: `${plural(all.length - keep.length, 'page')} removed · ${plural(keep.length, 'page')} left · ${sizeOf(b.length)}`, dl: 'Download PDF' }); } });
tool({ id: 'extract-pages', title: 'Extract pages', desc: 'Save only the pages you choose as a new PDF.', cat: 'organize', icon: 'split', keywords: 'select pages save pages',
  mode: 'pages', verb: 'Extract pages', ing: 'Extracting pages…', side: 'Extract pages', selVerb: 'keep',
  options: [{ t: 'info', text: 'Click the pages you want to keep.' }, { t: 'pages' }, { t: 'check', k: 'each', label: 'Save each page as its own PDF (ZIP)' }], next: ['merge-pdf', 'compress-pdf', 'pdf-to-word'],
  ready: F => F.pages.some(p => p.sel) ? '' : 'Select the pages to extract',
  run: async c => { const all = await c.load(), keep = all.filter((p, i) => c.sel[i]);
    if (c.o.each && keep.length > 1) { const r = await X.tools.splitParts({ pages: all, groups: keep.map(p => [p]), progress: c.progress, check: c.check }); return { blob: r.blob, name: c.base + '-pages.zip', kind: 'zip', inner: 'pdf', head: 'Pages have been extracted!', sum: `${plural(r.parts, 'PDF')} in a ZIP · ${sizeOf(r.blob.size)}`, dl: 'Download ZIP' }; }
    const b = await X.buildPdf(keep); return pdfRes(b, c.base + '-extract.pdf', { head: 'Pages have been extracted!', sum: `${plural(keep.length, 'page')} · ${sizeOf(b.length)}`, dl: 'Download PDF' }); } });
tool({ id: 'organize-pdf', title: 'Organize PDF', desc: 'Drag pages into order, rotate, duplicate, replace or insert blank pages.', cat: 'organize', icon: 'grid', keywords: 'rearrange reorder sort pages move',
  mode: 'editor', verb: 'Organize', edit: () => X.ACT.organize() });
tool({ id: 'rotate-pdf', title: 'Rotate PDF', desc: 'Rotate one page, chosen pages or the whole document.', cat: 'organize', icon: 'rotate', keywords: 'turn orientation landscape',
  mode: 'pages', verb: 'Rotate PDF', ing: 'Rotating pages…', side: 'Rotate', rotPages: true,
  options: [{ t: 'rotate' }, { t: 'info', text: 'Hover a page to turn just that one, or select pages and use the buttons above. Nothing selected = every page.' }, { t: 'pages', compact: true }],
  next: ['compress-pdf', 'merge-pdf', 'organize-pdf'],
  ready: F => F.pages.some(p => p.rot) ? '' : 'Turn at least one page first',
  run: async c => { const all = await c.load(), b = await X.buildPdf(all), n = c.pageSel.filter(p => p.rot).length; return pdfRes(b, c.base + '-rotated.pdf', { head: 'Your PDF has been rotated!', sum: `${plural(n, 'page')} turned · ${sizeOf(b.length)}`, dl: 'Download rotated PDF' }); } });
tool({ id: 'remove-blank-pages', title: 'Remove blank pages', desc: 'Find empty scanner pages and remove them in one click.', cat: 'organize', icon: 'trash', keywords: 'empty white scanner', badge: 'New',
  verb: 'Remove blank pages', ing: 'Looking for blank pages…', side: 'Remove blank pages',
  options: [{ t: 'radio', k: 'ink', def: '0.0015', items: [{ v: '0.0005', title: 'Only truly empty', desc: 'Keeps pages with a speck of ink' }, { v: '0.0015', title: 'Normal', desc: 'Ignores scanner dust and specks', badge: 'Recommended' }, { v: '0.004', title: 'Generous', desc: 'Also removes nearly empty pages' }] }],
  next: ['compress-pdf', 'ocr-pdf', 'merge-pdf'],
  run: async c => { const all = await c.load(), blank = await X.tools.findBlank({ pages: all, ink: +c.o.ink, progress: c.progress, check: c.check });
    if (!blank.length) throw info(`No blank pages found in ${c.items[0].name} — it is unchanged.`); if (blank.length === all.length) throw info('Every page looks blank — nothing was removed.');
    const keep = all.filter(p => !blank.includes(p)), b = await X.buildPdf(keep), nums = blank.map(p => all.indexOf(p) + 1);
    return pdfRes(b, c.base + '-no-blanks.pdf', { head: 'Blank pages have been removed!', sum: `${plural(blank.length, 'page')} removed (${nums.slice(0, 12).join(', ')}${nums.length > 12 ? '…' : ''}) · ${plural(keep.length, 'page')} left`, dl: 'Download PDF' }); } });

/* Optimize */
tool({ id: 'compress-pdf', title: 'Compress PDF', desc: 'Make the PDF smaller while text stays sharp and selectable.', cat: 'optimize', icon: 'compress', keywords: 'reduce shrink size smaller optimize email',
  max: 20, verb: 'Compress PDF', ing: 'Compressing PDF…', side: 'Compression level',
  options: [{ t: 'radio', k: 'level', def: 'recommended', items: [{ v: 'extreme', title: 'Extreme compression', desc: 'Less quality, high compression' }, { v: 'recommended', title: 'Recommended compression', desc: 'Good quality, good compression', badge: 'Recommended' }, { v: 'less', title: 'Less compression', desc: 'High quality, less compression' }] },
    { t: 'check', k: 'gray', label: 'Turn pictures grayscale (smaller still)' }, { t: 'check', k: 'strip', label: 'Remove hidden extras (thumbnails, editor data)', def: true }],
  next: ['merge-pdf', 'protect-pdf', 'sign-pdf', 'pdf-to-word'],
  run: async c => {
    const lvl = { extreme: 'max', recommended: 'balanced', less: 'light' }[c.o.level] || 'balanced', outs = []; let before = 0, after = 0, pics = 0;
    for (let i = 0; i < c.items.length; i++) {
      const it = c.items[i]; if (c.items.length > 1) c.progress(i, c.items.length, 'File'); else c.status('Resampling pictures, merging duplicates'); let bytes = it.bytes; if (it.rot) bytes = await X.buildPdf(await c.load([it]));
      const st = await X.compressPdf(bytes, { ...X.COMPRESS_LEVELS[lvl], gray: !!c.o.gray, strip: c.o.strip !== false }); c.check();
      before += it.size; after += Math.min(st.bytes.length, it.size); pics += st.pictures || 0; outs.push({ name: baseOf(it.name) + '-compressed.pdf', data: st.bytes.length < it.size || it.rot ? st.bytes : it.bytes });
    }
    const r = await pack(outs, 'compressed-pdfs.zip'), pct = Math.round((1 - after / before) * 100), one = c.items.length === 1;
    return Object.assign(r, { head: pct > 0 ? `Your PDF${one ? ' is' : 's are'} now ${pct}% smaller!` : `Your PDF${one ? ' is' : 's are'} already compact`, badge: pct > 0 ? `−${pct}%` : '',
      from: before, to: after, sum: `${one ? esc(c.items[0].name) : plural(c.items.length, 'file')}${pics ? ` · ${plural(pics, 'picture')} resampled` : ''}${pct > 0 ? '' : ' · try Extreme compression or grayscale'}`, dl: one ? 'Download compressed PDF' : 'Download ZIP' });
  } });
tool({ id: 'ocr-pdf', title: 'OCR PDF', desc: 'Turn scans into searchable, selectable text — offline, 14 languages.', cat: 'optimize', icon: 'scan', keywords: 'searchable text recognition scan ocr', badge: 'Offline OCR',
  verb: 'Make searchable', ing: 'Reading text…', side: 'OCR',
  options: [{ t: 'select', k: 'lang', label: 'Language of the document', items: ocrLangs, def: () => store.get('xd3.ocr', 'eng') },
    { t: 'radio', k: 'which', def: 'need', items: [{ v: 'need', title: 'Pages without text', desc: 'Scans and photos only — fastest', badge: 'Recommended' }, { v: 'all', title: 'All pages', desc: 'Also reads text inside pictures' }] },
    { t: 'info', text: 'About 2–5 seconds per page. Everything is read on this device.' }],
  next: ['compress-pdf', 'pdf-to-word', 'merge-pdf'],
  run: async c => {
    if (!X.ocrItems) throw fail('Text recognition is not available here (lib/ocr is missing).');
    if (c.o.lang) { store.set('xd3.ocr', c.o.lang); const s = $('#pOcrLang'); if (s) s.value = c.o.lang; }
    const pages = await c.load(); let lines = 0, done = 0;
    for (let i = 0; i < pages.length; i++) { c.progress(i, pages.length); const p = pages[i]; if (c.o.which !== 'all' && (await X.getLines(p)).length) continue; let f; try { f = await X.ocrItems(p, true, true); } catch (e) { throw fail(`Text recognition could not start, so no text was added (${String((e && e.message) || e || 'the engine did not load').replace(/\.$/, '')}).` + (/^https?:$/.test(location.protocol) ? ' If you are offline: open OCR PDF once while connected, and it then works offline too.' : ''), { cause: e }); } p.ocr = f.map(slim); lines += f.length; done++; }
    if (!done) throw info('Every page already has text. Choose "All pages" to also read text inside pictures.');
    const b = await X.buildPdf(pages); return pdfRes(b, c.base + '-searchable.pdf', { head: lines ? 'Your PDF is now searchable!' : 'No printed text was found', sum: `${plural(lines, 'line')} of text recognised on ${plural(done, 'page')} · ${sizeOf(b.length)}`, dl: 'Download searchable PDF' });
  } });
tool({ id: 'batch-process', title: 'Batch process', desc: 'Compress, OCR, watermark or number many PDFs at once.', cat: 'optimize', icon: 'layers', keywords: 'bulk many files at once',
  max: 200, verb: 'Process PDFs', ing: 'Processing PDFs…', side: 'Batch process',
  options: [{ t: 'list', k: 'act', def: 'compress', items: [{ v: 'compress', title: 'Compress', desc: 'Re-encode pictures (recommended level)' }, { v: 'ocr', title: 'Make searchable', desc: 'OCR text layer on scanned pages' }, { v: 'watermark', title: 'Text watermark', desc: 'Diagonal text across every page' }, { v: 'pagenum', title: 'Page numbers', desc: '“1 / N” at the bottom centre' }, { v: 'flatten', title: 'Flatten', desc: 'Pages become pictures, 150 dpi' }] },
    { t: 'text', k: 'wm', label: 'Watermark text', def: 'CONFIDENTIAL', when: o => o.act === 'watermark' }],
  next: ['merge-pdf', 'protect-pdf'],
  run: async c => { const outs = [];
    for (let i = 0; i < c.items.length; i++) { const it = c.items[i]; c.progress(i, c.items.length, 'File'); outs.push({ name: baseOf(it.name) + '-' + c.o.act + '.pdf', data: await X.batchOne(it.bytes, c.o.act, c.o.wm) }); }
    const r = await pack(outs, 'xd3-batch-' + c.o.act + '.zip'), tot = outs.reduce((a, o) => a + o.data.length, 0);
    return Object.assign(r, { head: `${plural(outs.length, 'PDF has', 'PDFs have')} been processed!`, sum: `${plural(outs.length, 'file')} · ${sizeOf(tot)}`, dl: outs.length > 1 ? 'Download ZIP' : 'Download PDF' }); } });

/* Convert to PDF */
tool({ id: 'word-to-pdf', title: 'Word to PDF', desc: 'Turn .docx files into PDFs that keep their fonts and layout.', cat: 'convert', group: 'to', icon: 'word', color: FMT.word, tag: 'DOC', keywords: 'docx doc convert',
  kinds: ['docx'], max: 20, verb: 'Convert to PDF', ing: 'Converting to PDF…', side: 'Word to PDF',
  options: [{ t: 'check', k: 'one', label: 'Join all files into one PDF', when: (o, F) => F.items.length > 1 }],
  next: ['merge-pdf', 'compress-pdf', 'protect-pdf', 'sign-pdf'],
  run: async c => { const outs = [], all = [];
    for (let i = 0; i < c.items.length; i++) { c.progress(i, c.items.length, 'File'); const it = c.items[i], pg = await c.load([it]); if (c.o.one && c.items.length > 1) all.push(...pg); else { S.fileName = baseOf(it.name) + '.pdf'; outs.push({ name: baseOf(it.name) + '.pdf', data: await X.buildPdf(pg) }); } }
    if (all.length) outs.push({ name: c.base + '.pdf', data: await X.buildPdf(all) });
    const r = await pack(outs, 'word-to-pdf.zip'); return Object.assign(r, { head: outs.length > 1 ? 'Your PDFs are ready!' : 'Your PDF is ready!', sum: `${plural(c.items.length, 'Word file')} → ${plural(outs.length, 'PDF')} · ${sizeOf(r.blob.size)}`, dl: outs.length > 1 ? 'Download ZIP' : 'Download PDF' }); } });
tool({ id: 'jpg-to-pdf', title: 'JPG to PDF', desc: 'Turn photos and images into PDF pages, with size and margin options.', cat: 'convert', group: 'to', icon: 'image', color: FMT.image, tag: 'JPG', keywords: 'image png photo picture convert',
  kinds: ['img'], max: 300, verb: 'Convert to PDF', ing: 'Making your PDF…', side: 'Image to PDF', camera: true,
  options: [{ t: 'select', k: 'size', label: 'Page size', def: 'fit', items: [['fit', 'Same as the image'], ['a4', 'A4'], ['letter', 'US Letter']] },
    { t: 'select', k: 'orient', label: 'Orientation', def: 'auto', items: [['auto', 'Automatic'], ['portrait', 'Portrait'], ['landscape', 'Landscape']], when: o => o.size !== 'fit' },
    { t: 'select', k: 'margin', label: 'Margin', def: 'none', items: [['none', 'No margin'], ['small', 'Small'], ['big', 'Big']], when: o => o.size !== 'fit' },
    { t: 'check', k: 'one', label: 'Put all images in one PDF', def: true, when: (o, F) => F.items.length > 1 }],
  next: ['ocr-pdf', 'compress-pdf', 'merge-pdf'],
  run: async c => {
    const o = c.o, M = { none: 0, small: 20, big: 40 }[o.margin] || 0, pages = [];
    for (let i = 0; i < c.items.length; i++) {
      c.progress(i, c.items.length, 'Image'); const it = c.items[i], im = await X.readImage(it.file); let p;
      if (o.size === 'fit') p = X.imagePage(im);
      else { let [W, H] = o.size === 'letter' ? [612, 792] : X.A4; const turned = it.rot % 180, iw = turned ? im.h : im.w, ih = turned ? im.w : im.h; if (o.orient === 'landscape' || (o.orient === 'auto' && iw > ih)) [W, H] = [H, W];
        const [PW, PH] = turned ? [H, W] : [W, H], f = Math.min((PW - 2 * M) / im.w, (PH - 2 * M) / im.h), w = im.w * f, h = im.h * f;
        p = { id: X.uid(), src: -1, w: PW, h: PH, rot: 0, objs: [{ id: X.uid(), type: 'image', asset: X.addAsset(im.url), x: (PW - w) / 2, y: (PH - h) / 2, w, h, opacity: 1 }] }; }
      p.rot = it.rot || 0; p.name = it.name; pages.push(p);
    }
    const outs = o.one !== false || pages.length === 1 ? [{ name: (pages.length === 1 ? c.base : 'images') + '.pdf', data: await X.buildPdf(pages) }] : await Promise.all(pages.map(async p => ({ name: baseOf(p.name) + '.pdf', data: await X.buildPdf([p]) })));
    const r = await pack(outs, 'images-pdf.zip'); return Object.assign(r, { head: 'Your PDF is ready!', sum: `${plural(pages.length, 'image')} → ${plural(outs.length, 'PDF')} · ${sizeOf(r.blob.size)}`, dl: outs.length > 1 ? 'Download ZIP' : 'Download PDF' });
  } });
tool({ id: 'scan-to-pdf', title: 'Scan to PDF', desc: 'Use your camera or photos to make a clean, cropped PDF.', cat: 'convert', group: 'to', icon: 'scan', color: FMT.other, tag: 'PDF', keywords: 'camera photo document scanner',
  mode: 'start', start: 'Start scanning', startIcon: 'scan', sub: 'Point your camera at a page, or import photos — pages are straightened and cropped for you.',
  begin: async () => { if (S.pages.length) { if (!(await guardUnsaved())) return; X.resetDoc(); X.buildPages(); X.updateUI(); } leave(); X.ACT.scan(); } });
tool({ id: 'create-pdf', title: 'Create PDF', desc: 'Start a blank PDF and add text, images and pages.', cat: 'convert', group: 'to', icon: 'plus', color: FMT.other, tag: 'PDF', keywords: 'new blank empty document',
  mode: 'start', start: 'Create blank PDF', startIcon: 'plus', sub: 'A blank A4 page opens in the editor. Add text, pictures, shapes and more pages.',
  begin: async () => { if (!(await guardUnsaved())) return; S.dirty = false; leave(); X.ACT.newdoc(); } });

/* Convert from PDF */
tool({ id: 'pdf-to-word', title: 'PDF to Word', desc: 'Get an editable .docx that looks exactly like the PDF.', cat: 'convert', group: 'from', icon: 'word', color: FMT.word, tag: 'DOC', keywords: 'docx doc convert editable',
  verb: 'Convert to Word', ing: 'Converting to Word…', side: 'Convert to Word',
  options: [{ t: 'radio', k: 'mode', def: 'exact', items: [{ v: 'exact', title: 'Exact layout', desc: 'Looks exactly like the PDF', badge: 'Recommended' }, { v: 'edit', title: 'Editable text', desc: 'Real paragraphs that re-flow as you type' }, { v: 'pic', title: 'Picture', desc: 'Each page as an image — not editable' }] }, RANGE()],
  next: ['word-editor', 'spelling-grammar', 'word-to-pdf'],
  run: async c => { const pages = pickRange(await c.load(), c.o.range); c.check(); c.status(`Converting ${plural(pages.length, 'page')}…`); let blob;
    try { blob = await X.buildDocx(pages, c.o.mode); } catch (e) { if (c.o.mode !== 'exact') throw e; console.warn('exact mode failed, using "same"', e); blob = await X.buildDocx(pages, 'same'); }
    return { blob, name: c.base + '.docx', kind: 'docx', head: 'Your Word file is ready!', sum: `${esc(c.items[0].name)} · ${plural(pages.length, 'page')} → ${sizeOf(blob.size)} .docx`, dl: 'Download Word file' }; } });
tool({ id: 'pdf-to-powerpoint', title: 'PDF to PowerPoint', desc: 'One editable slide for each page.', cat: 'convert', group: 'from', icon: 'slides', color: FMT.ppt, tag: 'PPT', keywords: 'pptx slides presentation convert',
  verb: 'Convert to PowerPoint', ing: 'Making slides…', side: 'Convert to PowerPoint',
  options: [{ t: 'radio', k: 'q', def: '220|0.9', items: [{ v: '220|0.9', title: 'Sharp', desc: '220 dpi — best for projectors', badge: 'Recommended' }, { v: '150|0.85', title: 'Standard', desc: '150 dpi' }, { v: '110|0.8', title: 'Small file', desc: '110 dpi — for email' }] }],
  next: ['pdf-to-word', 'compress-pdf'],
  run: async c => { const pages = await c.load(), [dpi, q] = c.o.q.split('|').map(Number), blob = await X.tools.toPptx({ pages, dpi, q, progress: c.progress, check: c.check });
    return { blob, name: c.base + '.pptx', kind: 'pptx', head: 'Your PowerPoint file is ready!', sum: `${plural(pages.length, 'slide')} · ${sizeOf(blob.size)}`, dl: 'Download PowerPoint file' }; } });
tool({ id: 'pdf-to-excel', title: 'PDF to Excel', desc: 'Tables become rows and columns; numbers stay numbers.', cat: 'convert', group: 'from', icon: 'table', color: FMT.excel, tag: 'XLS', keywords: 'xlsx spreadsheet table convert',
  verb: 'Convert to Excel', ing: 'Reading the tables…', side: 'Convert to Excel',
  options: [{ t: 'radio', k: 'mode', def: 'one', items: [{ v: 'one', title: 'One sheet', desc: 'All pages on one sheet' }, { v: 'each', title: 'A sheet per page', desc: 'Each page on its own sheet' }] }],
  next: ['pdf-to-word', 'ocr-pdf'],
  run: async c => { const pages = await c.load(); let r;
    try { r = await X.tools.toXlsx({ pages, mode: c.o.mode, progress: c.progress, check: c.check }); } catch (e) { if (/No text found/.test(e.message)) throw Object.assign(info('No text found — this looks like a scan. Run OCR PDF first, then convert.'), { fix: { label: 'Run OCR PDF', tool: 'ocr-pdf' } }); throw e; }
    return { blob: r.blob, name: c.base + '.xlsx', kind: 'xlsx', head: 'Your Excel file is ready!', sum: `${plural(r.sheets, 'sheet')} · ${sizeOf(r.blob.size)}`, dl: 'Download Excel file' }; } });
tool({ id: 'pdf-to-jpg', title: 'PDF to JPG', desc: 'Save every page as a PNG or JPG image.', cat: 'convert', group: 'from', icon: 'image', color: FMT.image, tag: 'JPG', keywords: 'image png picture export convert',
  verb: 'Convert to images', ing: 'Rendering pages…', side: 'PDF to images',
  options: [{ t: 'radio', k: 'fmt', def: 'jpg', items: [{ v: 'jpg', title: 'JPG', desc: 'Smaller files, for photos and sharing' }, { v: 'png', title: 'PNG', desc: 'Sharper text and lines, larger files' }] },
    { t: 'select', k: 'dpi', label: 'Resolution', def: '150', items: [['96', '96 dpi (screen)'], ['150', '150 dpi'], ['200', '200 dpi'], ['300', '300 dpi (print)']] }, RANGE()],
  next: ['jpg-to-pdf', 'compress-pdf'],
  run: async c => { const all = await c.load(), pages = pickRange(all, c.o.range), mime = c.o.fmt === 'png' ? 'image/png' : 'image/jpeg', len = String(all.length).length, outs = [];
    for (let i = 0; i < pages.length; i++) { c.progress(i, pages.length); const p = pages[i]; outs.push({ name: `${c.base}-page-${String(all.indexOf(p) + 1).padStart(len, '0')}.${c.o.fmt}`, data: await toBlob(await X.renderPageImage(p, +c.o.dpi / 72), mime, 0.88) }); }
    const r = await pack(outs, c.base + '-images.zip', mime); if (outs.length === 1) r.kind = 'img';
    return Object.assign(r, { head: outs.length > 1 ? 'Your images are ready!' : 'Your image is ready!', sum: `${plural(outs.length, 'page')} · ${c.o.fmt.toUpperCase()} · ${c.o.dpi} dpi · ${sizeOf(r.blob.size)}`, dl: outs.length > 1 ? 'Download ZIP' : 'Download image' }); } });
tool({ id: 'extract-images', title: 'Extract images', desc: 'Save every picture inside the PDF (ZIP).', cat: 'convert', group: 'from', icon: 'image', color: FMT.image, tag: 'PNG', keywords: 'pictures photos get',
  verb: 'Extract images', ing: 'Collecting pictures…', side: 'Extract images', options: [{ t: 'info', text: 'Every picture is saved as a PNG at 200 dpi, in a ZIP.' }],
  next: ['pdf-to-jpg', 'compress-pdf'],
  run: async c => { const pages = await c.load(), r = await X.tools.extractImages({ pages, progress: c.progress, check: c.check }); if (!r.n) throw info('This PDF has no pictures to extract.');
    return { blob: r.blob, name: c.base + '-pictures.zip', kind: 'zip', inner: 'img', head: `${plural(r.n, 'picture has', 'pictures have')} been extracted!`, sum: `${plural(r.n, 'PNG')} in a ZIP · ${sizeOf(r.blob.size)}`, dl: 'Download ZIP' }; } });
tool({ id: 'pdf-to-text', title: 'PDF to Text', desc: 'Export all the text as a plain .txt file.', cat: 'convert', group: 'from', icon: 'text', color: FMT.other, tag: 'TXT', keywords: 'txt plain text export copy',
  verb: 'Convert to text', ing: 'Reading the text…', side: 'PDF to Text', options: [{ t: 'check', k: 'marks', label: 'Mark where each page starts', def: true }],
  next: ['spelling-grammar', 'pdf-to-word'],
  run: async c => { const pages = await c.load(), parts = [];
    for (let i = 0; i < pages.length; i++) { c.progress(i, pages.length); parts.push(X.rowsOf(await X.visibleText(pages[i])).map(r => r.parts.map(q => q.str).join('\t')).join('\r\n')); }
    if (!parts.some(t => t.trim())) throw Object.assign(info('No text found — scanned pages contain pictures of text. Run OCR PDF first.'), { fix: { label: 'Run OCR PDF', tool: 'ocr-pdf' } });
    const text = parts.map((t, i) => c.o.marks && pages.length > 1 ? `—— Page ${i + 1} ——\r\n${t}` : t).join('\r\n\r\n'), words = (text.match(/\S+/g) || []).length;
    return { blob: new Blob(['﻿' + text], { type: 'text/plain;charset=utf-8' }), name: c.base + '.txt', kind: 'txt', text, head: 'Your text file is ready!', sum: `${plural(pages.length, 'page')} · ${words.toLocaleString()} words`, dl: 'Download text file' }; } });

/* Edit */
tool({ id: 'edit-pdf', title: 'Edit PDF', desc: 'Change existing text in its own font, add text, images and shapes.', cat: 'edit', icon: 'edittext', keywords: 'change text modify write', mode: 'editor', edit: () => X.setTool('edittext') });
tool({ id: 'fill-pdf-forms', title: 'Fill PDF forms', desc: 'Type into form fields and tick boxes, then save.', cat: 'edit', icon: 'text', keywords: 'form fields fill in', mode: 'editor', edit: () => { X.setTool('select'); X.toast('Click a form field to type in it, then Save', 'info'); } });
tool({ id: 'annotate-pdf', title: 'Annotate PDF', desc: 'Highlight, underline, strike out, draw and add sticky notes.', cat: 'edit', icon: 'marker', keywords: 'highlight comment note draw markup', mode: 'editor', edit: () => X.setTool('marker') });
tool({ id: 'add-image', title: 'Add image', desc: 'Place a picture or logo anywhere; remove its background.', cat: 'edit', icon: 'image', keywords: 'logo picture photo insert', mode: 'editor', edit: () => X.ACT.image() });
tool({ id: 'add-watermark', title: 'Add watermark', pv: true, desc: 'Stamp text or an image over your pages, at any angle and opacity.', cat: 'edit', icon: 'drop', keywords: 'stamp confidential draft logo',
  verb: 'Add watermark', ing: 'Adding watermark…', side: 'Watermark options',
  options: [{ t: 'list', k: 'type', def: 'text', items: [{ v: 'text', title: 'Text', desc: 'e.g. CONFIDENTIAL or DRAFT' }, { v: 'image', title: 'Image', desc: 'Your logo or a stamp picture' }] },
    { t: 'text', k: 'text', label: 'Text', def: 'CONFIDENTIAL', when: o => o.type === 'text' },
    { t: 'file', k: 'img', label: 'Picture', accept: 'image/*', when: o => o.type === 'image' },
    { t: 'color', k: 'color', label: 'Colour', def: '#dc2626', when: o => o.type === 'text' },
    { t: 'select', k: 'pos', label: 'Position', def: 'c', items: [['c', 'Centre'], ['t', 'Top'], ['b', 'Bottom']] },
    { t: 'select', k: 'angle', label: 'Angle', def: '-35', items: [['0', 'Straight'], ['-35', 'Diagonal up'], ['35', 'Diagonal down'], ['-90', 'Vertical']] },
    { t: 'select', k: 'op', label: 'Opacity', def: '0.18', items: [['0.1', '10%'], ['0.18', '18%'], ['0.3', '30%'], ['0.5', '50%'], ['1', '100%']] },
    { t: 'select', k: 'size', label: 'Size', def: '0.8', items: [['0.35', 'Small'], ['0.55', 'Medium'], ['0.8', 'Large']] }, RANGE()],
  next: ['protect-pdf', 'flatten-pdf', 'compress-pdf'],
  check: o => o.type === 'text' ? (o.text || '').trim() ? '' : 'Type the watermark text' : o.img ? '' : 'Choose a picture for the watermark',
  run: async c => { const o = c.o, all = await c.load(), pages = pickRange(all, o.range);
    const wm = o.type === 'image' ? await X.readImage(o.img) : X.makeWatermark(o.text.trim(), o.color || '#dc2626', 160), iw = o.type === 'image' ? wm.w * 0.75 : wm.w, ih = o.type === 'image' ? wm.h * 0.75 : wm.h;
    const asset = X.addAsset(wm.url);
    for (const p of pages) p.objs.push({ id: X.uid(), type: 'image', asset, ...wmPlace(o, p.w, p.h, iw, ih) });
    const b = await X.buildPdf(all); return pdfRes(b, c.base + '-watermarked.pdf', { head: 'Your watermark has been added!', sum: `${plural(pages.length, 'page')} watermarked · ${sizeOf(b.length)}`, dl: 'Download PDF' }); } });
tool({ id: 'add-page-numbers', title: 'Add page numbers', pv: true, desc: 'Number pages with your choice of position, style and start number.', cat: 'edit', icon: 'hash', keywords: 'numbering paginate footer',
  verb: 'Add page numbers', ing: 'Numbering pages…', side: 'Page numbers',
  options: [{ t: 'select', k: 'pos', label: 'Position', def: 'bc', items: POS6 },
    { t: 'select', k: 'fmt', label: 'Style', def: '{n}', items: [['{n}', '1'], ['Page {n}', 'Page 1'], ['{n} / {N}', '1 / N'], ['Page {n} of {N}', 'Page 1 of N'], ['- {n} -', '- 1 -']] },
    { t: 'num', k: 'start', label: 'Start at', def: 1, min: 0, max: 99999 }, { t: 'num', k: 'size', label: 'Font size', def: 10, min: 5, max: 40 },
    { t: 'select', k: 'font', label: 'Font', def: 'Helvetica', items: [['Helvetica', 'Helvetica (sans)'], ['Times', 'Times (serif)'], ['Courier', 'Courier (mono)']] },
    { t: 'color', k: 'color', label: 'Colour', def: '#333333' }, RANGE('Only these pages (blank = all)'), { t: 'check', k: 'skip1', label: 'Skip the first page (cover)' }],
  next: ['compress-pdf', 'sign-pdf', 'protect-pdf'],
  run: async c => { const o = c.o, all = await c.load(); let idx = o.range && o.range.trim() ? X.parseRange(o.range, all.length) : all.map((_, i) => i); if (!idx) throw fail(`That page range isn't valid. Use page numbers 1–${all.length}.`); if (o.skip1) idx = idx.filter(i => i > 0); if (!idx.length) throw fail('There are no pages left to number.');
    const st = Number.isFinite(+o.start) ? +o.start : 1, N = idx.length + st - 1;
    idx.forEach((pi, k) => stamp(all[pi], X.textObj(o.fmt.replace(/\{n\}/g, k + st).replace(/\{N\}/g, N), +o.size || 10, { color: o.color || '#333333', font: o.font, pn: true }), o.pos, 28));
    const b = await X.buildPdf(all); return pdfRes(b, c.base + '-numbered.pdf', { head: 'Page numbers have been added!', sum: `${plural(idx.length, 'page')} numbered · ${sizeOf(b.length)}`, dl: 'Download PDF' }); } });
tool({ id: 'header-footer', title: 'Header & footer', pv: true, desc: 'Add a title, date or page text to the top and bottom of every page.', cat: 'edit', icon: 'stroke', keywords: 'header footer title date running',
  verb: 'Add header & footer', ing: 'Adding header and footer…', side: 'Header & footer',
  options: [{ t: 'text', k: 'head', label: 'Header', ph: 'e.g. Company name — {date}' }, { t: 'text', k: 'foot', label: 'Footer', ph: 'e.g. Page {page} of {pages}', def: 'Page {page} of {pages}' },
    { t: 'select', k: 'al', label: 'Align', def: 'c', items: [['c', 'Centre'], ['l', 'Left'], ['r', 'Right']] }, { t: 'num', k: 'size', label: 'Font size', def: 10, min: 5, max: 40 }, { t: 'num', k: 'm', label: 'Margin (pt)', def: 28, min: 4, max: 200 },
    { t: 'info', text: 'Placeholders: {page} {pages} {date} {file}' }],
  next: ['compress-pdf', 'sign-pdf', 'protect-pdf'],
  check: o => (o.head || '').trim() || (o.foot || '').trim() ? '' : 'Type a header or a footer',
  run: async c => { const o = c.o, all = await c.load(), N = all.length, fill = (t, i) => t.replace(/\{page\}/g, i + 1).replace(/\{pages\}/g, N).replace(/\{date\}/g, new Date().toLocaleDateString()).replace(/\{file\}/g, c.base);
    all.forEach((p, i) => { if ((o.head || '').trim()) stamp(p, X.textObj(fill(o.head.trim(), i), +o.size || 10), 't' + o.al, +o.m || 28); if ((o.foot || '').trim()) stamp(p, X.textObj(fill(o.foot.trim(), i), +o.size || 10), 'b' + o.al, +o.m || 28); });
    const b = await X.buildPdf(all); return pdfRes(b, c.base + '-header-footer.pdf', { head: 'Header and footer have been added!', sum: `${plural(N, 'page')} · ${sizeOf(b.length)}`, dl: 'Download PDF' }); } });
tool({ id: 'letterhead', title: 'Letterhead', desc: 'Put your letterhead on every page — design one or use your own.', cat: 'edit', icon: 'stamp', keywords: 'company logo header branding', mode: 'editor', edit: () => X.ACT.letterhead() });
tool({ id: 'crop-resize', title: 'Crop & resize', desc: 'Trim margins or change the page size (A4, Letter…).', cat: 'edit', icon: 'crop', keywords: 'trim margins page size a4 letter', mode: 'editor', edit: () => { X.setTool('crop'); X.toast('Drag a box over the part to keep — or right-click a page ▸ This page ▸ Page size', 'info'); } });

/* Sign & Security */
tool({ id: 'sign-pdf', title: 'Sign PDF', desc: "Draw, type or upload your signature, plus initials, stamps and today's date.", cat: 'security', icon: 'sign', keywords: 'signature esign initials', mode: 'editor', edit: () => X.ACT.sign() });
tool({ id: 'protect-pdf', title: 'Protect PDF', desc: 'Add a password with AES encryption.', cat: 'security', icon: 'lock', keywords: 'password encrypt secure lock',
  verb: 'Protect PDF', ing: 'Encrypting PDF…', side: 'Protect PDF',
  options: [{ t: 'pw', k: 'pw', label: 'Password', ac: 'new-password' }, { t: 'pw', k: 'pw2', label: 'Repeat password', ac: 'new-password' }, { t: 'check', k: 'show', label: 'Show passwords' },
    { t: 'head', text: 'People who open it may' }, { t: 'check', k: 'print', label: 'Print', def: true }, { t: 'check', k: 'copy', label: 'Copy text and pictures', def: true }, { t: 'check', k: 'edit', label: 'Edit, annotate and fill forms', def: true },
    { t: 'info', text: 'AES-128 encryption that opens in any PDF reader. A forgotten password cannot be recovered — keep it somewhere safe.' }],
  next: ['compress-pdf', 'sign-pdf', 'merge-pdf'],
  check: o => { if (!(window.crypto && crypto.subtle)) return 'Password protection needs the desktop app or the local start file (a secure page).'; if ((o.pw || '') !== (o.pw2 || '')) return 'The two passwords do not match.'; if (/[^\x20-\xff]/.test(o.pw || '')) return 'Please use only English letters, digits and symbols in the password.'; if (!o.pw && o.print && o.copy && o.edit) return 'Enter a password, or untick something people may not do.'; return ''; },
  run: async c => { const o = c.o, b = await X.buildPdf(await c.load()); c.check(); let P = 0xfffff0c0; if (o.print) P |= 4 | 2048; if (o.edit) P |= 8 | 32 | 256 | 1024; if (o.copy) P |= 16 | 512;
    const out = await X.encryptPdf(b, o.pw || '', ownerPw(), P | 0);
    return pdfRes(out, c.base + '-protected.pdf', { head: o.pw ? 'Your PDF is now protected!' : 'Your PDF is now restricted!', sum: `${esc(c.items[0].name)} · AES-128 · ${o.pw ? 'needs the password to open' : 'opens without a password'} · ${sizeOf(out.length)}`, dl: 'Download protected PDF', noEdit: true }); } });
tool({ id: 'unlock-pdf', title: 'Unlock PDF', desc: 'Open a password-protected PDF and save a copy without the password.', cat: 'security', icon: 'key', keywords: 'remove password decrypt unprotect',
  verb: 'Unlock PDF', ing: 'Unlocking PDF…', side: 'Unlock PDF', unlockTool: true,
  options: [{ t: 'pw', k: 'pw', label: 'Password of the PDF', ac: 'off', when: (o, F) => F.items.some(i => i.locked) }, { t: 'info', text: 'You need to know the password — it is checked on this device, nothing is sent anywhere.' }],
  next: ['compress-pdf', 'merge-pdf', 'edit-pdf'],
  run: async c => { const it = c.items[0];
    if (!it.enc) throw info(`${it.name} has no password — there is nothing to unlock.`);
    let bytes = it.bytes; if (it.locked) { const r = await unlockBytes(it.raw, c.o.pw || ''); if (!r) throw fail('That password is not correct — try again.'); bytes = r.bytes; }
    return pdfRes(bytes, c.base + '-unlocked.pdf', { head: 'Your PDF is unlocked!', sum: `${esc(it.name)} · password and restrictions removed · ${sizeOf(bytes.length)}`, dl: 'Download unlocked PDF' }); } });
tool({ id: 'redact-pdf', title: 'Redact PDF', desc: 'Black out sensitive text and areas — removed from the file for good.', cat: 'security', icon: 'redact', keywords: 'black out hide censor privacy', mode: 'editor', edit: () => X.setTool('redact') });
tool({ id: 'flatten-pdf', title: 'Flatten PDF', desc: 'Make edits, forms and annotations permanent and uneditable.', cat: 'security', icon: 'layers', keywords: 'flatten permanent picture forms',
  verb: 'Flatten PDF', ing: 'Flattening pages…', side: 'Flatten PDF',
  options: [{ t: 'radio', k: 'q', def: '150|0.72', items: QUALITY }, { t: 'check', k: 'gray', label: 'Grayscale' }, { t: 'info', text: 'Every page becomes a picture: text can no longer be selected or searched.' }],
  next: ['compress-pdf', 'protect-pdf', 'ocr-pdf'],
  run: async c => { const pages = await c.load(), [dpi, q] = c.o.q.split('|').map(Number), out = await PDFLib.PDFDocument.create();
    for (let i = 0; i < pages.length; i++) { c.progress(i, pages.length); const p = pages[i], cv = await X.renderPageImage(p, dpi / 72); if (c.o.gray && X.applyFx) X.applyFx(cv, { mode: 'gray' }); const [w, h] = X.rdims(p); out.addPage([w, h]).drawImage(await out.embedJpg(cv.toDataURL('image/jpeg', q)), { x: 0, y: 0, width: w, height: h }); }
    out.setProducer('XD3 PDF Editor (XD3Labs / XDCybertech Pvt Ltd)'); const b = await out.save();
    return pdfRes(b, c.base + '-flattened.pdf', { head: 'Your PDF has been flattened!', sum: `${plural(pages.length, 'page')} · ${dpi} dpi · ${sizeOf(b.length)}`, dl: 'Download flattened PDF' }); } });
tool({ id: 'remove-hidden-data', title: 'Remove hidden data', desc: 'Strip metadata, scripts, attachments and comments.', cat: 'security', icon: 'shield', keywords: 'sanitize metadata privacy clean scripts',
  verb: 'Clean PDF', ing: 'Cleaning PDF…', side: 'Remove hidden data',
  options: [{ t: 'check', k: 'meta', label: 'Title, author, program and other details', def: true }, { t: 'check', k: 'js', label: 'Scripts and automatic actions', def: true }, { t: 'check', k: 'att', label: 'Attached files', def: true }, { t: 'check', k: 'cmt', label: 'Comments, highlights and notes' }, { t: 'check', k: 'extra', label: 'Page thumbnails and editor data', def: true }],
  next: ['protect-pdf', 'compress-pdf'],
  run: async c => { const o = c.o, pages = await c.load(); c.check(); const r = await X.tools.sanitize({ pages, meta: !!o.meta, js: !!o.js, att: !!o.att, cmt: !!o.cmt, extra: !!o.extra });
    return { blob: r.blob, name: c.base + '-clean.pdf', kind: 'pdf', head: 'Hidden data has been removed!', sum: `Removed: ${r.done.join(', ') || 'nothing'} · ${sizeOf(r.blob.size)}`, dl: 'Download clean PDF' }; } });
tool({ id: 'compare-pdf', title: 'Compare PDF', desc: 'See what changed between two versions, side by side.', cat: 'security', icon: 'swap', keywords: 'difference diff versions changes',
  min: 2, max: 2, verb: 'Compare PDFs', ing: 'Comparing pages…', side: 'Compare PDF', pick: 'Select 2 PDF files',
  options: [{ t: 'info', text: 'The first file is the original. In the result, red is only in the first file, green only in the second, grey is the same in both.' }],
  next: ['pdf-to-text'],
  run: async c => { const A = await c.load([c.items[0]]), B = await c.load([c.items[1]]), n = Math.max(A.length, B.length), k = 1.6, out = await PDFLib.PDFDocument.create(), font = await out.embedFont(PDFLib.StandardFonts.Helvetica); let changed = 0;
    for (let i = 0; i < n; i++) { c.progress(i, n); const pa = A[i], pb = B[i], [w, h] = pa ? X.rdims(pa) : X.rdims(pb), W = Math.round(w * k), H = Math.round(h * k);
      const ink = async p => { if (!p) return null; const cv = await X.renderPageImage(p, k), o = document.createElement('canvas'); o.width = W; o.height = H; const g = o.getContext('2d', { willReadFrequently: true }); g.fillStyle = '#fff'; g.fillRect(0, 0, W, H); g.drawImage(cv, 0, 0, W, H); return g.getImageData(0, 0, W, H).data; };
      const da = await ink(pa), db = await ink(pb), cv = document.createElement('canvas'); cv.width = W; cv.height = H; const g = cv.getContext('2d'), img = g.createImageData(W, H), d = img.data; let diff = 0;
      for (let j = 0; j < d.length; j += 4) { const ia = da ? (da[j] + da[j + 1] + da[j + 2]) / 3 < 190 : false, ib = db ? (db[j] + db[j + 1] + db[j + 2]) / 3 < 190 : false; let r = 255, gg = 255, bb = 255;
        if (ia && ib) r = gg = bb = 110; else if (ia) { r = 225; gg = 45; bb = 45; diff++; } else if (ib) { r = 25; gg = 150; bb = 70; diff++; } d[j] = r; d[j + 1] = gg; d[j + 2] = bb; d[j + 3] = 255; }
      g.putImageData(img, 0, 0); const differs = diff > W * H * 0.0005; if (differs) changed++;
      const pg = out.addPage([w, h]); pg.drawImage(await out.embedJpg(cv.toDataURL('image/jpeg', 0.8)), { x: 0, y: 0, width: w, height: h });
      pg.drawText(`Page ${i + 1} — ${differs ? 'differences found' : 'identical'}${!pa ? ' (only in the second file)' : !pb ? ' (only in the first file)' : ''}`, { x: 8, y: h - 16, size: 9, font, color: differs ? PDFLib.rgb(0.73, 0.11, 0.11) : PDFLib.rgb(0.08, 0.5, 0.24) }); }
    const b = await out.save(); return pdfRes(b, 'comparison.pdf', { head: changed ? `${plural(changed, 'page differs', 'pages differ')}` : 'The two PDFs look identical', sum: `${esc(c.items[0].name)} vs ${esc(c.items[1].name)} · ${plural(n, 'page')} compared · red = only in the first, green = only in the second`, dl: 'Download comparison' }); } });

/* Write & Word */
tool({ id: 'word-editor', title: 'Word editor', desc: 'Edit .docx files with their own fonts and styles — Auto Format and Fix Formatting.', cat: 'write', icon: 'word', keywords: 'docx word edit document',
  mode: 'editor', kinds: ['docx'], alt: { label: 'or start a blank document', run: async () => { if (!(await guardWord())) return; leave(); X.docx.newDocument(); } },
  go: async ([f]) => { if (!(await guardWord())) return false; leave(); await X.docx.openFile(f); return true; } });
tool({ id: 'word-templates', title: 'Word templates', desc: 'Letter, résumé, invoice, report and memo templates with your letterhead.', cat: 'write', icon: 'stamp', keywords: 'template letter resume invoice memo',
  mode: 'start', start: 'Choose a template', startIcon: 'stamp', sub: 'Pick a letter, résumé, invoice, report or memo — it opens in the Word editor, ready to fill in.', begin: async () => { if (!(await guardWord())) return; leave(); X.ACT.dxnew(); } });
const pdfWriter = async tab => { X.writing.show(); await X.writing.ACTS.pdf(); if (tab) X.writing.setTab(tab); }; // the PDF open in the editor → the writer
const writerGo = tab => async ([f]) => {
  if (kindOf(f) === 'pdf') return toEditor([f], () => pdfWriter(tab));
  leave(); await X.openInWriter(f); if (tab) X.writing.setTab(tab); return true;
};
tool({ id: 'spelling-grammar', title: 'Spelling & grammar', desc: 'Find and fix mistakes in PDFs, Word and text — offline.', cat: 'write', icon: 'spell', keywords: 'proofread check spell grammar typo',
  mode: 'editor', kinds: ['pdf', 'docx', 'txt'], pick: 'Select a file', alt: { label: 'or type / paste text instead', run: () => { leave(); X.ACT.writer(); } }, tab: '', go: writerGo('') });
tool({ id: 'ai-humanizer', title: 'AI Humanizer', desc: 'Make stiff or AI-sounding text read naturally.', cat: 'write', icon: 'sparkle', keywords: 'rewrite natural tone ai paraphrase', badge: 'New',
  mode: 'editor', kinds: ['pdf', 'docx', 'txt'], pick: 'Select a file', alt: { label: 'or type / paste text instead', run: () => { leave(); X.ACT.writer(); setTimeout(() => X.writing.setTab('human'), 50); } }, tab: 'human', go: writerGo('human') });

/* ───────── unlocking: the Standard security handler (RC4 40/128, AES-128, AES-256), decrypted in place with pdf-lib ───────── */
const PAD = Uint8Array.from([0x28, 0xbf, 0x4e, 0x5e, 0x4e, 0x75, 0x8a, 0x41, 0x64, 0x00, 0x4e, 0x56, 0xff, 0xfa, 0x01, 0x08, 0x2e, 0x2e, 0x00, 0xb6, 0xd0, 0x68, 0x3e, 0x80, 0x2f, 0x0c, 0xa9, 0xfe, 0x64, 0x53, 0x69, 0x7a]);
const cat = (...a) => { const o = new Uint8Array(a.reduce((s, x) => s + x.length, 0)); let k = 0; for (const x of a) { o.set(x, k); k += x.length; } return o; };
const same = (a, b) => a.length === b.length && a.every((v, i) => v === b[i]);
function rc4(key, data) {
  const s = new Uint8Array(256); for (let i = 0; i < 256; i++) s[i] = i;
  for (let i = 0, j = 0; i < 256; i++) { j = (j + s[i] + key[i % key.length]) & 255; const t = s[i]; s[i] = s[j]; s[j] = t; }
  const out = new Uint8Array(data.length);
  for (let k = 0, i = 0, j = 0; k < data.length; k++) { i = (i + 1) & 255; j = (j + s[i]) & 255; const t = s[i]; s[i] = s[j]; s[j] = t; out[k] = data[k] ^ s[(s[i] + s[j]) & 255]; }
  return out;
}
const sha = async (alg, d) => new Uint8Array(await crypto.subtle.digest(alg, d));
const aesKey = (k, use) => crypto.subtle.importKey('raw', k, 'AES-CBC', false, use);
async function aesEnc(key, iv, data) { return new Uint8Array(await crypto.subtle.encrypt({ name: 'AES-CBC', iv }, await aesKey(key, ['encrypt']), data)).slice(0, data.length); } // (no padding: the extra block is dropped)
async function aesDec(key, iv, data, nopad) {
  const k = await aesKey(key, ['encrypt', 'decrypt']);
  if (nopad) data = cat(data, new Uint8Array(await crypto.subtle.encrypt({ name: 'AES-CBC', iv: data.slice(-16) }, k, new Uint8Array(16).fill(16))).slice(0, 16)); // a valid padding block, so WebCrypto accepts it
  return new Uint8Array(await crypto.subtle.decrypt({ name: 'AES-CBC', iv }, k, data));
}
async function hashR6(pw, salt, u, r6) { // ISO 32000-2 algorithm 2.B (R6); R5 is a plain SHA-256
  let K = await sha('SHA-256', cat(pw, salt, u)); if (!r6) return K;
  for (let i = 0, E = null; i < 64 || E[E.length - 1] > i - 32; i++) {
    const one = cat(pw, K, u), K1 = new Uint8Array(one.length * 64); for (let j = 0; j < 64; j++) K1.set(one, j * one.length);
    E = await aesEnc(K.slice(0, 16), K.slice(16, 32), K1); let m = 0; for (let j = 0; j < 16; j++) m += E[j];
    K = await sha(['SHA-256', 'SHA-384', 'SHA-512'][m % 3], E);
  }
  return K.slice(0, 32);
}
async function unlockBytes(bytes, password) { // → { bytes, P, upw } without encryption, or null when the password is wrong; throws when the file cannot be read
  const L = PDFLib, N = n => L.PDFName.of(n), P0 = L.PDFParser.prototype, held = [];
  // object streams are encrypted as a whole: pdf-lib would try (and fail) to read them while loading, so they are set aside and read after decryption.
  // Only this one parser is changed, so other pdf-lib loads running at the same time are not affected.
  const parser = L.PDFParser.forBytesWithOptions(bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes), 100, false, false);
  let last = null;
  parser.parseIndirectObjectHeader = function () { return (last = P0.parseIndirectObjectHeader.call(this)); };
  parser.parseObject = function () { const o = P0.parseObject.call(this); if (last && o instanceof L.PDFRawStream && o.dict.lookup(N('Type')) === N('ObjStm')) { held.push([last, o]); last = null; return L.PDFNull; } return o; };
  const doc = new L.PDFDocument(await parser.parseDocument(), true, false), ctx = doc.context;
  const encRef = ctx.trailerInfo.Encrypt; if (!encRef) return { bytes, plain: true };
  const enc = ctx.lookup(encRef); if (!(enc instanceof L.PDFDict)) throw new Error('Unknown encryption');
  const filt = enc.lookup(N('Filter')); if (filt && filt !== N('Standard')) throw fail('This PDF is locked with a certificate or a special program — it cannot be unlocked here.');
  const num = k => { const v = enc.lookup(N(k)); return v instanceof L.PDFNumber ? v.asNumber() : 0; }, bstr = (d, k) => { const v = d.lookup(N(k)); return v && v.asBytes ? v.asBytes() : new Uint8Array(0); };
  const V = num('V'), R = num('R'), O = bstr(enc, 'O'), U = bstr(enc, 'U'), P = num('P'), em = enc.lookup(N('EncryptMetadata')), encMeta = !(em instanceof L.PDFBool) || em.asBoolean();
  let ids = ctx.trailerInfo.ID; if (ids instanceof L.PDFRef) ids = ctx.lookup(ids); const id0 = ids instanceof L.PDFArray && ids.size() ? ids.lookup(0).asBytes() : new Uint8Array(0);
  // the file key
  let key = null, upw = password; const n = V === 1 ? 5 : Math.max(5, Math.min(16, (num('Length') || 40) >> 3));
  if (R >= 5) {
    const pw = new TextEncoder().encode(password).slice(0, 127), r6 = R >= 6, none = new Uint8Array(0);
    if (same(await hashR6(pw, U.slice(32, 40), none, r6), U.slice(0, 32))) key = await aesDec(await hashR6(pw, U.slice(40, 48), none, r6), new Uint8Array(16), bstr(enc, 'UE'), true);
    else if (same(await hashR6(pw, O.slice(32, 40), U.slice(0, 48), r6), O.slice(0, 32))) key = await aesDec(await hashR6(pw, O.slice(40, 48), U.slice(0, 48), r6), new Uint8Array(16), bstr(enc, 'OE'), true);
  } else {
    const latin = s => Uint8Array.from(s, ch => ch.charCodeAt(0) & 255), padded = b => cat(b.slice(0, 32), PAD).slice(0, 32), le = new Uint8Array(new Int32Array([P]).buffer);
    const keyOf = upw => { let h = X.md5(cat(padded(upw), O.slice(0, 32), le, id0, R >= 4 && !encMeta ? Uint8Array.from([255, 255, 255, 255]) : new Uint8Array(0))); if (R >= 3) for (let i = 0; i < 50; i++) h = X.md5(h.slice(0, n)); return h.slice(0, n); };
    const okU = k => { if (R === 2) return same(rc4(k, PAD), U.slice(0, 32)); let x = rc4(k, X.md5(cat(PAD, id0))); for (let i = 1; i <= 19; i++) x = rc4(k.map(b => b ^ i), x); return same(x.slice(0, 16), U.slice(0, 16)); };
    const pw = latin(password); let k = keyOf(pw);
    if (okU(k)) key = k;
    else { // the owner password: it unlocks the user password stored in O
      let h = X.md5(padded(pw)); if (R >= 3) for (let i = 0; i < 50; i++) h = X.md5(h); const ok = h.slice(0, n); let up = O.slice(0, 32);
      if (R === 2) up = rc4(ok, up); else for (let i = 19; i >= 0; i--) up = rc4(ok.map(b => b ^ i), up);
      k = keyOf(up); if (okU(k)) { key = k; let e = 0; while (e < 32 && !same(up.slice(e), PAD.slice(0, 32 - e))) e++; upw = String.fromCharCode(...up.slice(0, e)); } // (the user password, padded, is what O held)
    }
  }
  if (!key) return null;
  // how strings and streams are encrypted
  const cfm = name => { if (V < 4) return 'rc4'; if (!name || name === N('Identity')) return 'none'; const cf = enc.lookup(N('CF')), d = cf instanceof L.PDFDict ? cf.lookup(name) : null, m = d instanceof L.PDFDict ? d.lookup(N('CFM')) : null; return m === N('AESV2') || m === N('AESV3') ? 'aes' : m === N('V2') ? 'rc4' : 'none'; };
  const strM = cfm(enc.lookup(N('StrF'))), stmM = cfm(enc.lookup(N('StmF')));
  const objKey = (ref, aes) => R >= 5 ? key : X.md5(cat(key, Uint8Array.from([ref.objectNumber & 255, ref.objectNumber >> 8 & 255, ref.objectNumber >> 16 & 255, ref.generationNumber & 255, ref.generationNumber >> 8 & 255]), aes ? Uint8Array.from([0x73, 0x41, 0x6c, 0x54]) : new Uint8Array(0))).slice(0, Math.min(n + 5, 16));
  const dec = async (m, k, d) => { if (m === 'rc4') return rc4(k, d); if (m !== 'aes') return d; if (d.length < 32) return new Uint8Array(0); try { return await aesDec(k, d.slice(0, 16), d.slice(16)); } catch (e) { return d.length % 16 ? d : aesDec(k, d.slice(0, 16), d.slice(16), true); } };
  const hex = b => Array.from(b, x => (x < 16 ? '0' : '') + x.toString(16)).join('');
  const walk = async (o, k) => { // strings anywhere inside an object → decrypted
    if (o instanceof L.PDFString || o instanceof L.PDFHexString) return L.PDFHexString.of(hex(await dec(strM, k, o.asBytes())));
    if (o instanceof L.PDFDict) { for (const [nk, v] of o.entries()) { const nv = await walk(v, k); if (nv !== v) o.set(nk, nv); } return o; }
    if (o instanceof L.PDFArray) { for (let i = 0; i < o.size(); i++) { const v = o.get(i), nv = await walk(v, k); if (nv !== v) o.set(i, nv); } return o; }
    return o;
  };
  const metaPlain = !encMeta, inner = new Set(), asg = ctx.assign;
  for (const [ref] of held) ctx.delete(ref);
  ctx.assign = function (r, o) { inner.add(r.tag); return asg.call(this, r, o); }; // (on this document only: what the object streams hold)
  try {
    for (const [ref, st] of held) { // the objects inside an object stream carry no encryption of their own
      const ns = L.PDFRawStream.of(st.dict, stmM === 'none' ? st.contents : await dec(stmM, objKey(ref, stmM === 'aes'), st.contents));
      try { await L.PDFObjectStreamParser.forStream(ns).parseIntoContext(); } catch (e) { /* unreadable object stream */ }
    }
  } finally { ctx.assign = asg; }
  for (const [ref, obj] of ctx.enumerateIndirectObjects()) {
    if (ref === encRef || (encRef instanceof L.PDFRef && ref.tag === encRef.tag) || inner.has(ref.tag)) continue;
    if (obj instanceof L.PDFRawStream) {
      const type = obj.dict.lookup(N('Type')); if (type === N('XRef')) continue;
      if (strM !== 'none') await walk(obj.dict, objKey(ref, strM === 'aes'));
      if (stmM === 'none' || (metaPlain && type === N('Metadata'))) continue;
      ctx.assign(ref, L.PDFRawStream.of(obj.dict, await dec(stmM, objKey(ref, stmM === 'aes'), obj.contents)));
    } else if (strM !== 'none') await walk(obj, objKey(ref, strM === 'aes'));
  }
  ctx.trailerInfo.Encrypt = undefined; if (encRef instanceof L.PDFRef) ctx.delete(encRef);
  return { bytes: await doc.save({ useObjectStreams: false, updateFieldAppearances: false }), P, upw };
}
const ownerPw = () => Array.from(crypto.getRandomValues(new Uint8Array(24)), x => String.fromCharCode(33 + x % 90)).join('');
async function relock(r, lock) { // a result made from a password-protected PDF gets the same password and permissions again
  const enc = b => X.encryptPdf(b, lock.pw, ownerPw(), lock.P | 0);
  if (r.kind === 'pdf') { r.blob = new Blob([await enc(new Uint8Array(await r.blob.arrayBuffer()))], { type: 'application/pdf' }); return true; }
  if (r.kind !== 'zip' || r.inner !== 'pdf') return false;
  const JSZip = await X.loadLib('jszip'), z = await JSZip.loadAsync(r.blob);
  for (const n of Object.keys(z.files)) if (!z.files[n].dir && /\.pdf$/i.test(n)) z.file(n, await enc(await z.files[n].async('uint8array')));
  r.blob = await z.generateAsync({ type: 'blob', compression: 'DEFLATE', mimeType: 'application/zip' }); return true;
}

/* ───────── the editor's state, swapped out while a tool runs, so the open document is never touched ───────── */
const DOC_KEYS = ['sources', 'pages', 'assets', 'assetVer', 'xfonts', 'outline', 'undo', 'redo', 'sel', 'cur', 'meta', 'metaOrig', 'dirty', 'find', 'fileName', 'editing'];
let boxed = false;
async function sandbox(fn) {
  if (boxed) throw new Error('Another tool is still running');
  if (X.commitEdit) X.commitEdit();
  const keep = {}; DOC_KEYS.forEach(k => { keep[k] = S[k]; });
  Object.assign(S, { sources: [], pages: [], assets: {}, xfonts: {}, outline: [], undo: [], redo: [], sel: null, cur: 0, meta: {}, metaOrig: null, dirty: false, find: { q: '', hits: [], cur: -1 }, fileName: 'document.pdf', editing: null });
  boxed = true;
  try { return await fn(); } finally {
    for (const s of S.sources) { try { s.pdf.destroy(); } catch (e) { /* already closed */ } }
    DOC_KEYS.forEach(k => { S[k] = keep[k]; }); boxed = false;
    S.srcVer++; S.assetVer = (S.assetVer || 0) + 1; // the next autosave writes the open document again
    if (S.pages.length) X.scheduleSave();
    if (X.updateUI) X.updateUI(); if (visible && F.tool) document.title = F.tool.title + ' — XD3 PDF Editor';
  }
}

/* ───────── state ───────── */
const F = { tool: null, screen: 'land', items: [], pages: [], opts: {}, alerts: [], result: null, prog: null, cancel: false, running: false, carry: null, sideOpen: false };
let seq = 0, root = null, visible = false;
const optsFor = t => { const o = {}; for (const d of t.options) if (d.k) o[d.k] = typeof d.def === 'function' ? d.def() : d.def !== undefined ? d.def : d.t === 'check' ? false : d.t === 'radio' || d.t === 'list' ? d.items[0].v : ''; return o; };
function freeItem(it) { try { if (it.pdf) it.pdf.destroy(); } catch (e) { /* gone */ } if (it.url) URL.revokeObjectURL(it.url); }
function resetFiles() { F.items.forEach(freeItem); F.items = []; F.pages = []; F.alerts = []; F.result = null; if (pageIO) pageIO.disconnect(); }

/* ───────── files ───────── */
const hasEncrypt = b => { const d = new TextDecoder('latin1'), n = b.length, cut = 1 << 20; return /\/Encrypt\b/.test(d.decode(b.subarray(0, Math.min(n, cut)))) || (n > cut && /\/Encrypt\b/.test(d.decode(b.subarray(Math.max(cut, n - cut))))); };
async function prep(it) { // bytes, encryption, page count and the first-page picture
  if (it.kind === 'img') { it.url = URL.createObjectURL(it.file); it.n = 1; return; }
  if (it.kind !== 'pdf') return;
  it.raw = it.bytes = new Uint8Array(await it.file.arrayBuffer());
  if (hasEncrypt(it.bytes)) {
    it.enc = true;
    try { const r = await unlockBytes(it.raw, ''); if (r && !r.plain) { it.bytes = r.bytes; it.lock = { pw: '', P: r.P }; } else if (!r) it.locked = true; else it.enc = false; }
    catch (e) { if (e.friendly) { it.err = e.message; return; } it.locked = true; }
  }
  if (it.locked) return;
  await openPdf(it);
}
async function openPdf(it) {
  try { it.pdf = await pdfjsLib.getDocument({ data: it.bytes.slice(), ...PJ }).promise; it.n = it.pdf.numPages; }
  catch (e) { if (e && e.name === 'PasswordException') it.locked = true; else it.err = "We couldn't read this file — it may be damaged. Try Repair by re-saving it in the editor."; }
}
async function unlockItem(it, pw) {
  let r = null; try { r = await unlockBytes(it.raw, pw); } catch (e) { r = null; }
  if (!r) return false;
  it.bytes = r.bytes; it.locked = false; it.pw = pw; it.lock = { pw: r.upw, P: r.P }; await openPdf(it); return true;
}
async function addFiles(list) {
  const t = F.tool; if (!t || F.running) return;
  list = [...list]; if (!list.length) return;
  if (t.mode === 'editor') return goEditor(list);
  F.alerts = F.alerts.filter(a => a.keep);
  const ok = [];
  for (const f of list) {
    const k = kindOf(f);
    if (!t.kinds.includes(k)) {
      const what = t.kinds.length > 1 ? 'a ' + t.kinds.map(x => ({ pdf: 'PDF', docx: 'Word file', img: 'image', txt: 'text file' }[x])).join(', ').replace(/, ([^,]*)$/, ' or $1') : t.kinds[0] === 'pdf' ? 'a PDF' : 'a ' + NOUN[t.kinds[0]][0];
      const conv = k === 'docx' && t.kinds.includes('pdf') ? 'word-to-pdf' : k === 'img' && t.kinds.includes('pdf') ? 'jpg-to-pdf' : null;
      F.alerts.push({ kind: 'error', text: `${f.name} isn't ${what}. Choose ${t.kinds[0] === 'pdf' ? 'a .pdf file' : what}${conv ? ` — or use ${byId[conv].title} first.` : '.'}`, acts: conv ? [{ label: 'Convert it', fn: () => open(conv, [f]) }] : [] });
      continue;
    }
    if (f.size > MAX_BYTES) { F.alerts.push({ kind: 'error', text: `${f.name}: this file is too large to process here (max ~ 500 MB). Try Compress or Split first.` }); continue; }
    ok.push(f);
  }
  if (t.max === 1 && ok.length) { const al = F.alerts; resetFiles(); F.alerts = al; ok.splice(1); }
  const room = t.max - F.items.length; if (ok.length > room) { F.alerts.push({ kind: 'info', text: `${t.title} takes up to ${plural(t.max, 'file')} — ${plural(ok.length - room, 'file was', 'files were')} left out.` }); ok.splice(room); }
  if (!ok.length) { render(); return; }
  const fresh = ok.map(f => ({ id: ++seq, file: f, name: f.name, size: f.size, kind: kindOf(f), rot: 0 }));
  F.items.push(...fresh); F.screen = 'work'; render();
  await Promise.all(fresh.map(async it => { try { await prep(it); } catch (e) { it.err = "We couldn't read this file — it may be damaged. Try Repair by re-saving it in the editor."; } }));
  if (t.mode === 'pages' && F.items[0]) { const it = F.items[0]; F.pages = Array.from({ length: it.n || 0 }, (_, i) => ({ i, rot: 0, sel: false })); }
  render();
}
const pickFiles = (acc, multi, capture) => new Promise(res => { const i = document.createElement('input'); i.type = 'file'; i.accept = acc; i.multiple = !!multi; if (capture) i.setAttribute('capture', 'environment'); i.onchange = () => res([...i.files]); i.addEventListener('cancel', () => res([])); i.click(); });
const canCam = t => t.mode !== 'start' && (t.kinds.includes('img') || t.kinds.includes('pdf'));
async function choose(capture) {
  const t = F.tool; let files = await pickFiles(capture ? 'image/*' : accept(t), t.max > 1 || capture, capture); if (!files.length) return;
  if (capture && !t.kinds.includes('img')) { // photos → one PDF, made aside like a tool run
    if (F.running || boxed) return;
    try { files = [await sandbox(async () => { const pages = []; for (const f of files) pages.push(X.imagePage(await X.readImage(f))); return new File([await X.buildPdf(pages)], `scan-${new Date().toISOString().slice(0, 10)}.pdf`, { type: 'application/pdf' }); })]; }
    catch (e) { F.alerts = [{ kind: 'error', text: `We couldn't read those photos — ${e.message || e}` }]; return render(); }
  }
  addFiles(files);
}
async function useOpenDoc() {
  const name = /\.pdf$/i.test(S.fileName) ? S.fileName : baseOf(S.fileName) + '.pdf';
  const t = F.tool;
  if (t.mode === 'editor') { leave(); X.ACT.resume(); if (t.edit) await t.edit(); else await pdfWriter(t.tab); return; }
  let bytes; try { bytes = await X.buildPdf(S.pages); } catch (e) { F.alerts.push({ kind: 'error', text: e.message || String(e) }); return render(); }
  addFiles([new File([bytes], name, { type: 'application/pdf' })]);
}
async function addRecent(btn) {
  let list = []; try { list = (await X.DB.get('recent')) || []; } catch (e) { /* none */ }
  list = list.filter(r => F.tool.kinds.includes(r.kind === 'docx' ? 'docx' : 'pdf'));
  const box = document.createElement('div'); box.className = 'xf-pop';
  box.innerHTML = list.length ? '<b>Recent files</b>' + list.slice(0, 8).map((r, i) => `<button type="button" data-ri="${i}">${r.thumb ? `<img src="${r.thumb}" alt="">` : '<i data-icon="pdf"></i>'}<span>${esc(r.name)}</span></button>`).join('') : '<p>No recent files yet.</p>';
  X.icons(box); root.appendChild(box); const b = btn.getBoundingClientRect(), rr = root.getBoundingClientRect();
  box.style.left = Math.max(8, Math.min(b.left - rr.left - 120, rr.width - 300)) + 'px'; box.style.top = (b.bottom - rr.top + 8) + 'px';
  const close = e => { if (e && box.contains(e.target) && !e.target.closest('[data-ri]')) return; box.remove(); document.removeEventListener('pointerdown', close, true); };
  setTimeout(() => document.addEventListener('pointerdown', close, true), 0);
  box.onclick = e => { const b2 = e.target.closest('[data-ri]'); if (!b2) return; const r = list[+b2.dataset.ri]; close(); addFiles([new File([r.bytes], r.name, { type: r.kind === 'docx' ? 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' : 'application/pdf' })]); };
}

/* ───────── leaving for the editor ───────── */
function askUnsaved(name, what, save, discard) { // Keep editing / Download first / Discard → true to go on
  let dl = false;
  return X.dialog(`You have unsaved changes in ${esc(name)}`, `<p>${what}</p><div class="row"><button type="button" class="btn" data-xf="dl">Download first</button></div>`,
    { ok: 'Discard', onOpen: d => { const b = d.querySelector('[data-xf=dl]'); if (b) b.onclick = () => { dl = true; d.close('ok'); }; const c = d.querySelector('button[value=cancel]'); if (c) c.textContent = 'Keep editing'; } })
    .then(async r => { if (!r) return false; if (dl) await save(); discard(); return true; });
}
function guardUnsaved() { // the PDF editor's document
  if (!S.dirty || !S.pages.length) return Promise.resolve(true);
  return askUnsaved(S.fileName, 'Opening another file in the editor closes the document you are working on.', () => X.ACT.save(), () => { S.dirty = false; });
}
function guardWord() { // the Word editor's document (it stays open in the background while a tool is shown)
  const D = X.docx && X.docx.D; if (!D || !D.zip || !D.dirty) return Promise.resolve(true);
  return askUnsaved(X.docx.name(), 'Opening another document in the Word editor closes the one you are working on.', async () => X.download(await X.docx.buildDocx(), X.docx.name()), () => { D.dirty = false; });
}
async function toEditor(files, after) {
  if (!(await guardUnsaved())) return false;
  const before = S.pages; await X.openFiles(files, false);
  if (S.pages === before || !S.pages.length) return false;
  leave(); if (after) await after(); return true;
}
async function goEditor(files) {
  const t = F.tool, f = files.find(x => t.kinds.includes(kindOf(x)));
  if (!f) { F.alerts = [{ kind: 'error', text: `${files[0].name} isn't ${t.kinds[0] === 'pdf' ? 'a PDF. Choose a .pdf file — or use Word to PDF first.' : 'a file this tool can open.'}`, acts: kindOf(files[0]) === 'docx' && t.kinds[0] === 'pdf' ? [{ label: 'Convert it', fn: () => open('word-to-pdf', [files[0]]) }] : [] }]; return render(); }
  if (t.go) return t.go([f]);
  return toEditor([f], t.edit);
}

/* ───────── running a tool ───────── */
const CANCEL = { cancelled: true };
async function run() {
  const t = F.tool; if (F.running) return;
  const why = blocker(); if (why) { flashAlert(why); return; }
  const err = t.check && t.check(F.opts, F); if (err) { F.alerts = [{ kind: 'error', text: err, opt: true }]; if (NARROW.matches) F.sideOpen = true; return render(); }
  F.alerts = []; F.cancel = F.userCancel = F.leaveAfter = false; F.after = null; F.running = true; F.prog = null; F.sideOpen = false;
  const t0 = performance.now(), items = F.items.slice(), opts = { ...F.opts }, sel = F.pages.map(p => p.sel), pageSel = F.pages.map(p => ({ ...p }));
  root.classList.add('xf-busy'); blockPage(true);
  const timer = setTimeout(() => { if (F.running) { F.screen = 'proc'; render(); } }, 400); // quick jobs go straight to the result
  const ctx = {
    items, o: opts, sel, pageSel, base: baseOf(items[0] && items[0].name),
    progress: (i, n, what = 'Page') => { ctx.check(); F.prog = { i, n, what }; paintProg(); },
    status: txt => { F.prog = { txt }; paintProg(); },
    check: () => { if (F.cancel) throw CANCEL; },
    load: async (list = items) => { // the files as editor pages (inside the sandbox)
      const out = [];
      for (let k = 0; k < list.length; k++) {
        const it = list[k]; if (list.length > 1) ctx.status(`Reading ${it.name} (${k + 1} of ${list.length})…`); ctx.check(); let pg;
        if (it.kind === 'pdf') pg = await X.loadSource(it.bytes); else if (it.kind === 'docx') pg = await X.docxToPages(it.file); else pg = [X.imagePage(await X.readImage(it.file))];
        if (k === 0 && it.kind === 'pdf' && list === items) { try { const i = (await S.sources[pg[0].src].pdf.getMetadata()).info || {}; S.meta = { title: i.Title || '', author: i.Author || '', subject: i.Subject || '', keywords: i.Keywords || '' }; S.metaOrig = { ...S.meta }; } catch (e) { /* no info */ } }
        if (it.kind === 'pdf' && pg.length && X.readOutline) { // the bookmarks, as the editor keeps them (pages that stay keep theirs)
          try { const ol = await X.readOutline(S.sources[pg[0].src].pdf, pg); S.outline = list === items ? (S.outline || []).concat(ol) : ol; } catch (e) { /* no bookmarks */ } }
        if (t.mode === 'pages' && list === items) pg = pageSel.map(e => { const p = pg[e.i]; p.rot = ((p.rot || 0) + e.rot + 360) % 360; return p; });
        else if (it.rot) pg.forEach(p => { p.rot = ((p.rot || 0) + it.rot) % 360; });
        out.push(...pg);
      }
      return out;
    },
  };
  let res = null, error = null;
  try {
    res = await sandbox(async () => { S.fileName = ctx.base + '.pdf'; return t.run(ctx); });
    if (res && !F.cancel) await keepLock(t, res, items);
  }
  catch (e) { error = e; }
  finally { clearTimeout(timer); F.running = false; root.classList.remove('xf-busy'); blockPage(false); }
  if (F.cancel && !error) error = CANCEL; // (a step that could not stop: its result is dropped)
  if (F.leaveAfter) { // Back, Home… were pressed while it ran
    F.leaveAfter = false; F.cancel = false; const id = route(), after = F.after; F.after = null;
    if (after) { after(); return; }
    if (id !== t.id) { if (id) sync(); else hide(); return; }
  }
  if (!visible || F.tool !== t) return;
  if (error === CANCEL) { F.screen = 'work'; F.alerts = [{ kind: 'info', text: 'Stopped — nothing was saved.' }]; return render(); }
  if (error) {
    F.screen = 'work';
    const expected = error.info || error.friendly || /encrypt|password/i.test(error.message); // a message for the user (wrong password, no blank pages…), not a fault
    if (!expected) console.error(error); else if (error.cause) console.warn(error.message, error.cause);
    const p = F.prog && F.prog.n ? ` ${F.prog.what === 'Page' ? 'page' : F.prog.what.toLowerCase()} ${F.prog.i + 1}` : '';
    if (error.info) F.alerts = [{ kind: 'info', text: error.message, acts: error.fix ? [{ label: error.fix.label, fn: () => open(error.fix.tool, items.map(i => i.file)) }] : [] }];
    else if (error.friendly) F.alerts = [{ kind: 'error', text: error.message }];
    else if (/encrypt|password/i.test(error.message)) F.alerts = [{ kind: 'error', text: 'This PDF is password protected. Enter its password to open it.' }];
    else if (error instanceof RangeError || /memory|allocation/i.test(error.message)) F.alerts = [{ kind: 'error', text: 'This file is too large to process here (max ~ 500 MB). Try Compress or Split first.' }];
    else F.alerts = [{ kind: 'error', text: `Something went wrong while ${(t.ing || 'working').replace(/…$/, '').toLowerCase()}${p ? ' (at' + p + ')' : ''}.`, detail: error.message, acts: [{ label: 'Try again', fn: run, primary: true }, { label: 'Report', fn: () => report(t, error) }] }];
    return render();
  }
  F.result = Object.assign(res, { ms: performance.now() - t0, file: new File([res.blob], res.name, { type: res.blob.type }) });
  F.screen = 'done'; render();
}
async function keepLock(t, r, items) { // results of password-protected PDFs stay protected (same password and permissions)
  const locks = items.filter(i => i.lock).map(i => i.lock), pdfish = r.kind === 'pdf' || (r.kind === 'zip' && r.inner === 'pdf');
  if (t.id === 'protect-pdf') return;
  if (t.unlockTool) { r.name = r.name.replace(/-protected(?=[-.])/i, ''); return; }
  if (!locks.length) return;
  let done = false;
  if (pdfish && locks.every(l => l.pw === locks[0].pw && l.P === locks[0].P)) { try { done = await relock(r, locks[0]); } catch (e) { console.warn('could not protect the result again', e); } }
  if (done) { r.sum += locks[0].pw ? ' · still protected with the same password' : ' · same restrictions as the original'; r.locked = true; return; }
  r.name = r.name.replace(/-protected(?=[-.])/i, '');
  if (pdfish) r.warn = 'This result has no password: your files did not share the same one. Use Protect PDF to add a password.';
}
function report(t, e) {
  const txt = `XD3 PDF Editor — ${t.title} failed\n${e && e.stack || e}\n${navigator.userAgent}`;
  (navigator.clipboard ? navigator.clipboard.writeText(txt) : Promise.reject()).then(() => X.toast('Error details copied — paste them into an email to XD3Labs support', 'info'), () => X.toast(txt.slice(0, 300), 'info'));
}
function blocker() {
  const t = F.tool, n = F.items.length;
  if (n < t.min) return n ? `Add at least ${t.min} ${thing(t)}` : `Add ${t.min > 1 ? 'at least ' + t.min : 'a'} ${t.min > 1 ? thing(t) : NOUN[t.kinds[0]][0]}`;
  if (F.items.some(i => !i.pdf && i.kind === 'pdf' && !i.locked && !i.err)) return 'Reading files…';
  if (F.items.some(i => i.err)) return 'Remove the file that cannot be read';
  if (!t.unlockTool && F.items.some(i => i.locked)) return 'Unlock the password-protected file first';
  if (t.ready) return t.ready(F);
  return '';
}
function flashAlert(why) { const a = $('.xf-add', root); if (a) X.flash(a, 'xf-pulse'); X.toast(why, 'info'); }
let veil = null;
function blockPage(on) { // while a tool runs, clicks on the header (Open file, Home…) would act on the swapped-out editor state
  if (on) { if (!veil) { veil = document.createElement('div'); veil.className = 'xf-block'; document.body.appendChild(veil); } veil.hidden = false; document.body.classList.add('xflow-busy'); }
  else { if (veil) veil.hidden = true; document.body.classList.remove('xflow-busy'); }
}

/* ───────── rendering ───────── */
const ico = (t, big) => `<span class="xf-ico${big ? ' big' : ''}" style="--c:${t.color}"><i data-icon="${t.icon}"></i>${t.tag ? `<b>${t.tag}</b>` : ''}</span>`;
const SVG_ARROW = '<svg class="xf-arrow" viewBox="0 0 24 24" aria-hidden="true"><circle cx="12" cy="12" r="10"/><path d="M8 12h8M13 8l4 4-4 4"/></svg>';
const SVG_SHIELD = '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M12 3l8 3v6c0 4.5-3.4 8.3-8 9-4.6-.7-8-4.5-8-9V6z"/><path d="M8.5 12l2.5 2.5 4.5-5"/></svg>';
function render() {
  if (!root || !F.tool) return;
  const t = F.tool; root.dataset.screen = F.screen; root.style.setProperty('--cat', t.color);
  const h = F.screen === 'land' ? landing(t) : F.screen === 'work' ? workspace(t) : F.screen === 'proc' ? processing(t) : done(t);
  root.querySelector('.xf-body').innerHTML = h; X.icons(root);
  if (F.screen === 'work') { F.items.forEach(drawThumb); if (t.mode === 'pages') watchPages(); }
  if (F.screen === 'proc') paintProg();
}
const FAQ = { // a few short answers under each tool, by category (the first one is shared)
  all: ['Are my files uploaded?', 'No. Everything happens in this app, on your device — it even works offline. Nothing is sent anywhere.'],
  organize: [['Is my original file changed?', 'No. You download a new PDF; the file you added stays as it was.'], ['Are bookmarks kept?', 'Yes — bookmarks of the pages that stay are kept, and Merge keeps the bookmarks of every file.']],
  optimize: [['How much smaller will it get?', 'It depends on the pictures inside. Scans and photos shrink the most; text-only PDFs are often small already.'], ['Does OCR need the internet?', 'No. The text recognition runs on this device, in 14 languages.']],
  convert: [['What about scanned pages?', 'A scan is a picture of text. Run OCR PDF first so the words can be converted.'], ['Is my original file changed?', 'No. You get a new file to download; the one you added stays as it was.']],
  edit: [['Can I see it before I download?', 'The first page shows a preview where the tool allows it. You can also open the result in the editor.'], ['Is my original file changed?', 'No. You download a new PDF; the file you added stays as it was.']],
  security: [['What if I forget the password?', 'A forgotten password cannot be recovered — keep it somewhere safe.'], ['Do protected files stay protected?', 'Yes. A result made from a password-protected PDF gets the same password and permissions — only Unlock PDF removes them.']],
  write: [['Does it work offline?', 'Yes. Spelling, grammar and rewriting run on this device.'], ['Which files can I use?', 'PDF, Word (.docx) and plain text files — or type and paste text straight in.']],
};
const faqHtml = t => { const l = [FAQ.all, ...(FAQ[t.cat] || [])]; return `<section class="xf-faq" aria-label="Questions"><h2>Questions</h2>${l.map(([q, a]) => `<details><summary>${esc(q)}</summary><p>${esc(a)}</p></details>`).join('')}</section>`; };
function landing(t) {
  const hasOpen = S.pages.length && t.kinds.includes('pdf') && t.mode !== 'start';
  const how = t.mode === 'start' ? '' : `<ol class="xf-how"><li><b>1</b><span>Select or drop your ${t.kinds.length > 1 ? 'files' : NOUN[t.kinds[0]][t.max > 1 ? 1 : 0]}</span></li><li><b>2</b><span>${t.mode === 'editor' ? 'It opens in the editor with the right tool' : `Choose your options and click <em>${esc(t.verb)}</em>`}</span></li><li><b>3</b><span>${t.mode === 'editor' ? 'Save — nothing leaves this device' : 'Download — nothing leaves this device'}</span></li></ol>`;
  return `<div class="xf-land">
    ${ico(t, true)}<h1>${esc(t.title)}</h1><p class="xf-sub">${esc(t.sub || t.desc)}</p>
    ${alertsHtml()}
    <div class="xf-pick">
      ${t.mode === 'start' ? `<button type="button" class="xf-xl" data-x="begin"><i data-icon="${t.startIcon || 'plus'}"></i>${esc(t.start)}</button>` : `<button type="button" class="xf-xl" data-x="pick">${esc(pickLabel(t))}</button>
      <span class="xf-rounds"><button type="button" class="xf-round" data-x="recent" title="Open a recent file" aria-label="Open a recent file"><i data-icon="clock"></i></button>${canCam(t) ? '<button type="button" class="xf-round" data-x="camera" title="Scan with the camera, or use photos" aria-label="Scan with the camera"><i data-icon="scan"></i></button>' : ''}</span>`}
    </div>
    ${t.mode === 'start' ? '' : `<p class="xf-or">${esc(dropLabel(t))}</p>`}
    ${hasOpen ? `<button type="button" class="xf-link" data-x="useopen"><i data-icon="pdf"></i>Use the open document — ${esc(S.fileName)}</button>` : ''}
    ${t.alt ? `<button type="button" class="xf-link" data-x="alt">${esc(t.alt.label)}</button>` : ''}
    <p class="xf-trust">${SVG_SHIELD}Your files stay on this device — nothing is uploaded.</p>
    ${how}
    ${faqHtml(t)}
  </div>`;
}
function alertsHtml(only) {
  return F.alerts.map((a, i) => only && !only(a) ? '' : `<div class="xf-alert ${a.kind}" role="${a.kind === 'error' ? 'alert' : 'status'}"><i data-icon="${a.kind === 'error' ? 'x' : 'info'}"></i><div><p>${esc(a.text)}</p>${a.detail ? `<small>${esc(a.detail)}</small>` : ''}${a.acts && a.acts.length ? `<div class="xf-aacts">${a.acts.map((b, j) => `<button type="button" class="xf-btn${b.primary || j === 0 ? ' red' : ''}" data-alert="${i}" data-j="${j}">${esc(b.label)}</button>`).join('')}</div>` : ''}</div></div>`).join('');
}
function cardHtml(it, i, t) {
  const st = it.locked ? '<span class="xf-lock"><i data-icon="lock"></i>Locked</span>' : it.err ? '<span class="xf-lock bad"><i data-icon="x"></i>Unreadable</span>' : '';
  return `<div class="xf-card${it.locked || it.err ? ' off' : ''}" data-i="${i}" draggable="${F.items.length > 1}" tabindex="0" aria-label="${esc(it.name)}">
    ${it.n > 1 ? `<span class="xf-pages">${plural(it.n, 'page')}</span>` : ''}
    <div class="xf-acts"><button type="button" data-c="rot" title="Rotate" aria-label="Rotate"><i data-icon="rotate"></i></button>${it.kind !== 'docx' && !it.locked ? '<button type="button" data-c="zoom" title="Preview" aria-label="Preview"><i data-icon="zoomin"></i></button>' : ''}<button type="button" data-c="del" title="Remove" aria-label="Remove"><i data-icon="x"></i></button></div>
    <div class="xf-thumb" style="transform:rotate(${it.rot}deg)">${it.kind === 'img' && it.url ? `<img src="${it.url}" alt="">` : it.kind === 'docx' ? '<span class="xf-doc"><i data-icon="word"></i>DOCX</span>' : '<span class="xf-wait"></span>'}</div>
    ${st}<div class="xf-name" title="${esc(it.name)}">${esc(it.name)}</div></div>`;
}
function pageCardHtml(e, k, t) {
  return `<div class="xf-pcard${e.sel ? ' sel' : ''}${t.selVerb === 'remove' && e.sel ? ' gone' : ''}" data-p="${k}" tabindex="0" role="checkbox" aria-checked="${e.sel}" aria-label="Page ${e.i + 1}">
    <span class="xf-tick"><i data-icon="check"></i></span>${t.rotPages ? '<div class="xf-acts"><button type="button" data-c="prot" title="Rotate this page" aria-label="Rotate this page"><i data-icon="rotate"></i></button></div>' : ''}
    <div class="xf-thumb" style="transform:rotate(${e.rot}deg)"><span class="xf-wait"></span></div><div class="xf-pnum">${e.i + 1}</div></div>`;
}
function workspace(t) {
  const why = blocker(), locked = F.items.filter(i => i.locked), minAlert = F.items.length < t.min && F.items.length ? `<div class="xf-alert info"><i data-icon="info"></i><div><p>Please select more ${thing(t)} by clicking the + button again.</p></div></div>` : '';
  const grid = t.mode === 'pages' && F.items[0] && F.items[0].pdf ? `<div class="xf-grid pages">${F.pages.map((e, k) => pageCardHtml(e, k, t)).join('')}</div>` : `<div class="xf-grid">${F.items.map((it, i) => cardHtml(it, i, t)).join('')}</div>`;
  const lockHtml = t.unlockTool ? '' : locked.map(it => `<div class="xf-alert error"><i data-icon="lock"></i><div><p>${esc(it.name)} is password protected. Enter its password to open it.</p><div class="xf-unlock"><input type="password" data-unlock="${it.id}" placeholder="Password" autocomplete="off" aria-label="Password for ${esc(it.name)}"><button type="button" class="xf-btn red" data-x="unlock" data-id="${it.id}">Unlock</button></div>${it.bad ? '<small>That password is not correct — try again.</small>' : ''}</div></div>`).join('');
  const narrow = NARROW.matches, msgs = alertsHtml(narrow ? a => !a.opt : null) + lockHtml + minAlert; // (on phones the messages go above the files: the options panel is usually closed)
  return `<div class="xf-work${F.sideOpen ? ' side-open' : ''}">
    <div class="xf-area">
      <div class="xf-addbar">${t.mode === 'pages' ? '' : `<span class="xf-addw"><button type="button" class="xf-add${F.items.length < t.min ? ' xf-pulse' : ''}" data-x="pick" title="${t.max > 1 ? 'Add more files' : 'Choose another file'}" aria-label="Add files"><i data-icon="plus"></i>${F.items.length && t.max > 1 ? `<b>${F.items.length}</b>` : ''}</button><span class="xf-addx"><button type="button" class="xf-round xs" data-x="recent" title="Add a recent file" aria-label="Add a recent file"><i data-icon="clock"></i></button>${canCam(t) ? '<button type="button" class="xf-round xs" data-x="camera" title="Scan with the camera, or use photos" aria-label="Scan with the camera"><i data-icon="scan"></i></button>' : ''}</span></span>`}${t.sort && F.items.length > 1 ? '<button type="button" class="xf-round sm" data-x="sort" title="Sort A → Z" aria-label="Sort by name">A→Z</button>' : ''}</div>
      ${narrow && msgs ? `<div class="xf-msgs">${msgs}</div>` : ''}${grid}
    </div>
    <aside class="xf-side" aria-label="Options"><h2>${esc(t.side || t.title)}</h2><div class="xf-opts">${narrow ? alertsHtml(a => a.opt) : msgs}${t.options.map(optHtml).join('')}</div></aside>
    <button type="button" class="xf-gear" data-x="side" title="${F.sideOpen ? 'Close options' : 'Options'}" aria-label="Options"><i data-icon="${F.sideOpen ? 'x' : 'sliders'}"></i></button>
    <button type="button" class="xf-cta" data-x="run"${why ? ` aria-disabled="true" title="${esc(why)}"` : ''}>${esc(t.verb)}${SVG_ARROW}</button>
  </div>`;
}
function optHtml(d) {
  const o = F.opts, t = F.tool;
  if (d.when && !d.when(o, F)) return '';
  const id = 'xfo-' + (d.k || d.t);
  switch (d.t) {
    case 'info': return `<div class="xf-alert info soft"><i data-icon="info"></i><div><p>${esc(d.text)}</p></div></div>`;
    case 'head': return `<h3 class="xf-oh">${esc(d.text)}</h3>`;
    case 'radio': return `<div class="xf-radios" role="radiogroup">${d.items.map(it => `<label class="xf-rcard${o[d.k] === it.v ? ' on' : ''}"><input type="radio" name="${d.k}" value="${esc(it.v)}" data-k="${d.k}"${o[d.k] === it.v ? ' checked' : ''}><span class="xf-rt">${esc(it.title)}${it.badge ? `<em>${esc(it.badge)}</em>` : ''}</span><span class="xf-rd">${esc(it.desc)}</span></label>`).join('')}</div>`;
    case 'list': return `<div class="xf-list" role="radiogroup">${d.items.map(it => `<label class="xf-lrow${o[d.k] === it.v ? ' on' : ''}"><input type="radio" name="${d.k}" value="${esc(it.v)}" data-k="${d.k}"${o[d.k] === it.v ? ' checked' : ''}><b>${esc(it.title)}</b><span>${esc(it.desc)}</span></label>`).join('')}</div>`;
    case 'check': return `<label class="xf-check"><input type="checkbox" data-k="${d.k}"${o[d.k] ? ' checked' : ''}><span>${esc(d.label)}</span></label>`;
    case 'num': return `<label class="xf-field" for="${id}"><span>${esc(d.label)}</span><input id="${id}" type="number" data-k="${d.k}" value="${esc(o[d.k])}" min="${d.min}" max="${d.max}"></label>`;
    case 'text': return `<label class="xf-field" for="${id}"><span>${esc(d.label)}</span><input id="${id}" type="text" data-k="${d.k}" value="${esc(o[d.k] || '')}" placeholder="${esc(d.ph || '')}" autocomplete="off"></label>`;
    case 'pw': return `<label class="xf-field" for="${id}"><span>${esc(d.label)}</span><input id="${id}" type="${o.show ? 'text' : 'password'}" data-k="${d.k}" value="${esc(o[d.k] || '')}" autocomplete="${d.ac || 'off'}"></label>`;
    case 'color': return `<label class="xf-field color" for="${id}"><span>${esc(d.label)}</span><input id="${id}" type="color" data-k="${d.k}" value="${esc(o[d.k])}"></label>`;
    case 'select': { const items = typeof d.items === 'function' ? d.items() : d.items; return `<label class="xf-field" for="${id}"><span>${esc(d.label)}</span><select id="${id}" data-k="${d.k}">${items.map(([v, l]) => `<option value="${esc(v)}"${String(o[d.k]) === String(v) ? ' selected' : ''}>${esc(l)}</option>`).join('')}</select></label>`; }
    case 'file': return `<div class="xf-field"><span>${esc(d.label)}</span><button type="button" class="xf-btn grey" data-x="ofile" data-k="${d.k}" data-acc="${esc(d.accept)}"><i data-icon="upload"></i>${o[d.k] ? esc(o[d.k].name) : 'Choose a picture…'}</button></div>`;
    case 'rotate': return '<div class="xf-rotbtns"><button type="button" class="xf-btn grey" data-x="rotsel" data-d="-90"><i data-icon="rotatel"></i>Left</button><button type="button" class="xf-btn grey" data-x="rotsel" data-d="90"><i data-icon="rotate"></i>Right</button><button type="button" class="xf-btn grey" data-x="rotreset">Reset</button></div>';
    case 'pages': { const n = F.pages.length, s = F.pages.filter(p => p.sel).length;
      return `<div class="xf-psel"><div class="xf-pills">${[['all', 'All'], ['none', 'None'], ['odd', 'Odd'], ['even', 'Even']].map(([v, l]) => `<button type="button" class="xf-pill" data-x="psel" data-v="${v}">${l}</button>`).join('')}</div>
        ${d.compact ? '' : `<label class="xf-field" for="xfo-prange"><span>Pages</span><input id="xfo-prange" type="text" data-x="prange" value="${esc(rangeText())}" placeholder="e.g. 1-3, 7, 10-12" autocomplete="off"></label>`}
        <p class="xf-count">${n ? `${s} of ${plural(n, 'page')} selected${t.selVerb ? ` to ${t.selVerb}` : ''}` : ''}</p></div>`; }
  }
  return '';
}
const rangeText = () => { const out = []; let a = -1; F.pages.forEach((p, k) => { if (p.sel && a < 0) a = k; if ((!p.sel || k === F.pages.length - 1) && a >= 0) { const b = p.sel ? k : k - 1; out.push(a === b ? `${a + 1}` : `${a + 1}-${b + 1}`); a = -1; } }); return out.join(', '); };
function processing(t) {
  const stop = F.cancel; // (Cancel, Back or Home pressed: the current step is finishing)
  return `<div class="xf-proc"><h2>${esc(stop ? 'Stopping…' : t.ing || 'Working…')}</h2><div class="xf-bar${F.prog && F.prog.n ? '' : ' ind'}"><i></i><span></span></div><p class="xf-step">on this device · no upload</p><button type="button" class="xf-btn grey" data-x="cancel"${stop ? ' disabled' : ''}>${stop ? 'Stopping…' : 'Cancel'}</button>${stop ? '<p class="xf-note">Finishing the current step — nothing will be saved.</p>' : ''}</div>`;
}
function paintProg() {
  if (!root || F.screen !== 'proc') return; const bar = $('.xf-bar', root), step = $('.xf-step', root); if (!bar) return; const p = F.prog || {};
  if (p.n) { const pct = Math.round((p.i + 0.5) / p.n * 100); bar.classList.remove('ind'); bar.firstChild.style.width = pct + '%'; bar.lastChild.textContent = pct > 6 ? pct + '%' : ''; step.textContent = `${p.what} ${p.i + 1} of ${p.n} · on this device · no upload`; }
  else if (p.txt) { bar.classList.add('ind'); step.textContent = p.txt + ' · on this device · no upload'; }
}
function done(t) {
  const r = F.result, secs = r.ms < 1000 ? (r.ms / 1000).toFixed(2) : (r.ms / 1000).toFixed(1), canEdit = !r.noEdit && (r.kind === 'pdf' || r.kind === 'docx');
  let next = (t.next || []).map(id => byId[id]).filter(Boolean); if (r.warn && byId['protect-pdf']) next = [byId['protect-pdf'], ...next.filter(n => n.id !== 'protect-pdf')];
  return `<div class="xf-done"><div class="xf-dmain">
    <div class="xf-dhead"><button type="button" class="xf-round" data-x="back" title="Back to the options" aria-label="Back"><i data-icon="left"></i></button><h2>${esc(r.head)}</h2></div>
    ${r.warn ? `<div class="xf-alert info xf-dwarn" role="status"><i data-icon="info"></i><div><p>${esc(r.warn)}</p></div></div>` : ''}
    ${r.badge ? `<span class="xf-badge">${esc(r.badge)}</span>` : ''}
    <p class="xf-sum">${r.sum || ''}</p>
    ${r.from ? `<div class="xf-sizes"><span class="from">${sizeOf(r.from)}</span><i data-icon="right"></i><span class="to">${sizeOf(r.to)}</span></div>` : ''}
    <div class="xf-pick"><button type="button" class="xf-xl" data-x="download"><i data-icon="download"></i>${esc(r.dl || 'Download')}</button>
      <span class="xf-rounds">${canEdit ? `<button type="button" class="xf-round" data-x="edit" title="${r.kind === 'docx' ? 'Open in the Word editor' : 'Open in editor'}" aria-label="Open in editor"><i data-icon="edittext"></i></button>` : ''}${r.kind === 'pdf' && !r.noEdit ? '<button type="button" class="xf-round" data-x="print" title="Print" aria-label="Print"><i data-icon="print"></i></button>' : ''}${r.kind === 'txt' || r.kind === 'img' ? '<button type="button" class="xf-round" data-x="copy" title="Copy to clipboard" aria-label="Copy to clipboard"><i data-icon="copy"></i></button>' : ''}<button type="button" class="xf-round" data-x="restart" title="Start over" aria-label="Start over"><i data-icon="undo"></i></button></span></div>
    <div class="xf-trustcard">${SVG_SHIELD}<p>Processed on your device in ${secs} s. Nothing was uploaded; close this window and nothing remains.</p></div>
  </div>
  ${next.length ? `<aside class="xf-next"><h3>Continue to…</h3>${next.map(n => { const w = carryOf(n); return `<button type="button" class="xf-nrow" data-next="${n.id}">${ico(n)}<span class="xf-nt">${esc(n.title)}${w.from === 'input' ? `<small>with ${esc(w.label)}</small>` : ''}</span><i data-icon="right"></i></button>`; }).join('')}<button type="button" class="xf-nrow all" data-x="all"><span class="xf-nt">See all tools →</span></button></aside>` : ''}
  </div>`;
}

/* thumbnails */
const dpr = () => Math.min(2, window.devicePixelRatio || 1);
async function pageCanvas(pdf, n, W, H) {
  const pg = await pdf.getPage(n), v1 = pg.getViewport({ scale: 1 }), s = Math.min(W / v1.width, H / v1.height), vp = pg.getViewport({ scale: s * dpr() }), c = document.createElement('canvas');
  c.width = Math.max(1, Math.round(vp.width)); c.height = Math.max(1, Math.round(vp.height)); c.style.width = Math.round(vp.width / dpr()) + 'px'; c.style.height = Math.round(vp.height / dpr()) + 'px';
  const g = c.getContext('2d'); g.fillStyle = '#fff'; g.fillRect(0, 0, c.width, c.height); await pg.render({ canvasContext: g, viewport: vp, intent: 'print' }).promise;
  c.pw = v1.width; c.ph = v1.height; return c;
}
async function drawThumb(it) {
  if (!root || F.tool.mode === 'pages') return;
  const host = () => $(`.xf-card[data-i="${F.items.indexOf(it)}"] .xf-thumb`, root);
  if (!it.pdf) return;
  try { if (!it.thumb) it.thumb = await pageCanvas(it.pdf, 1, 150, 170); const h = host(); if (h && !h.contains(it.thumb)) { h.textContent = ''; h.appendChild(it.thumb); } } catch (e) { /* no picture */ }
  paintPreview(it);
}
// a live preview of the stamp (watermark, page number, header / footer) on the first page
const PV_FONT = { Helvetica: 'Arial, Helvetica, sans-serif', Times: "'Times New Roman', Times, serif", Courier: "'Courier New', Courier, monospace" };
async function pvItems(t, o, W, H) { // → pieces in page points
  if (t.id === 'add-watermark') {
    if (o.type === 'image') { if (!o.img) return []; if (!F.pvImg || F.pvImg.f !== o.img) F.pvImg = { f: o.img, im: await X.readImage(o.img) }; const im = F.pvImg.im; return [{ img: im.url, ...wmPlace(o, W, H, im.w * 0.75, im.h * 0.75) }]; }
    const tx = (o.text || '').trim(); if (!tx) return [];
    const key = tx + '|' + o.color; if (!F.pvWm || F.pvWm.k !== key) F.pvWm = { k: key, wm: X.makeWatermark(tx, o.color || '#dc2626', 160) };
    return [{ img: F.pvWm.wm.url, ...wmPlace(o, W, H, F.pvWm.wm.w, F.pvWm.wm.h) }];
  }
  const it = F.items[0], N = (it && it.n) || 1;
  if (t.id === 'add-page-numbers') { const st = Number.isFinite(+o.start) ? +o.start : 1; return [{ text: String(o.fmt).replace(/\{n\}/g, st).replace(/\{N\}/g, N + st - 1), pos: o.pos, m: 28, size: +o.size || 10, color: o.color || '#333333', font: o.font }]; }
  if (t.id === 'header-footer') {
    const fill = s => s.replace(/\{page\}/g, 1).replace(/\{pages\}/g, N).replace(/\{date\}/g, new Date().toLocaleDateString()).replace(/\{file\}/g, baseOf(it && it.name)), out = [], m = +o.m || 28;
    if ((o.head || '').trim()) out.push({ text: fill(o.head.trim()), pos: 't' + o.al, m, size: +o.size || 10 });
    if ((o.foot || '').trim()) out.push({ text: fill(o.foot.trim()), pos: 'b' + o.al, m, size: +o.size || 10 });
    return out;
  }
  return [];
}
async function paintPreview(it) {
  const t = F.tool; if (!t || !t.pv || !root || !it.thumb) return;
  const th = $(`.xf-card[data-i="${F.items.indexOf(it)}"] .xf-thumb`, root), c = it.thumb; if (!th || !th.contains(c)) return;
  let list = []; try { list = await pvItems(t, F.opts, c.pw, c.ph); } catch (e) { list = []; }
  const s = c.clientWidth / c.pw; if (!s || !th.contains(c)) return;
  let ov = $('.xf-pv', th); if (!ov) { ov = document.createElement('div'); ov.className = 'xf-pv'; ov.setAttribute('aria-hidden', 'true'); th.appendChild(ov); }
  ov.style.cssText = `left:${c.offsetLeft}px;top:${c.offsetTop}px;width:${c.clientWidth}px;height:${c.clientHeight}px`;
  const px = v => (v * s).toFixed(1) + 'px';
  ov.innerHTML = list.map(p => p.img ? `<img src="${p.img}" alt="" style="left:${px(p.x)};top:${px(p.y)};width:${px(p.w)};height:${px(p.h)};opacity:${p.opacity};transform:rotate(${p.rot}deg)">`
    : `<span style="${p.pos[0] === 't' ? 'top' : 'bottom'}:${px(p.m)};${p.pos[1] === 'c' ? 'left:50%;transform:translateX(-50%)' : p.pos[1] === 'l' ? 'left:' + px(p.m) : 'right:' + px(p.m)};font-size:${Math.max(4.5, p.size * s).toFixed(1)}px;color:${esc(p.color || '#333')};font-family:${PV_FONT[p.font] || PV_FONT.Helvetica}">${esc(p.text)}</span>`).join('');
}
let pvT = 0;
const paintPreviews = () => { clearTimeout(pvT); pvT = setTimeout(() => { if (visible && F.screen === 'work') F.items.forEach(paintPreview); }, 60); };
let pageIO = null;
function watchPages() {
  const it = F.items[0]; if (!it || !it.pdf) return; it.pcache = it.pcache || {};
  if (pageIO) pageIO.disconnect();
  const show = async el => { const k = +el.dataset.p, e = F.pages[k]; if (!e) return; const th = $('.xf-thumb', el);
    try { const c = it.pcache[e.i] || (it.pcache[e.i] = await pageCanvas(it.pdf, e.i + 1, 112, 150)); if (!th.contains(c)) { th.textContent = ''; th.appendChild(c); } } catch (er) { /* no picture */ } };
  pageIO = 'IntersectionObserver' in window ? new IntersectionObserver(es => es.forEach(en => { if (en.isIntersecting) { pageIO.unobserve(en.target); show(en.target); } }), { root: $('.xf-area', root), rootMargin: '400px' }) : null;
  $$('.xf-pcard', root).forEach(el => { if (pageIO) pageIO.observe(el); else show(el); });
}
async function zoom(it, pageNo = 1) {
  const box = document.createElement('div'); box.className = 'xf-zoom'; box.innerHTML = '<button type="button" class="xf-round" aria-label="Close preview"><i data-icon="x"></i></button><div class="xf-zimg"></div>'; X.icons(box); root.appendChild(box);
  const close = () => box.remove(); box.onclick = close; box.tabIndex = -1; box.focus(); box.onkeydown = e => { if (e.key === 'Escape') { e.stopPropagation(); close(); } };
  try { const host = $('.xf-zimg', box); if (it.kind === 'img') host.innerHTML = `<img src="${it.url}" alt="" style="transform:rotate(${it.rot}deg)">`; else { const c = await pageCanvas(it.pdf, pageNo, Math.min(innerWidth - 80, 900), innerHeight - 120); c.style.transform = `rotate(${it.rot}deg)`; host.appendChild(c); } } catch (e) { close(); }
}

/* ───────── events ───────── */
function onClick(e) {
  const t = F.tool; if (!t || F.running && !e.target.closest('[data-x=cancel]')) return;
  const al = e.target.closest('[data-alert]'); if (al) { const a = F.alerts[+al.dataset.alert], b = a && a.acts[+al.dataset.j]; if (b) b.fn(); return; }
  const nx = e.target.closest('[data-next]'); if (nx) { continueTo(byId[nx.dataset.next]); return; }
  if (F.sideOpen && e.target.closest('.xf-area')) { toggleSide(false); return; }
  const card = e.target.closest('.xf-card'), pc = e.target.closest('.xf-pcard'), c = e.target.closest('[data-c]');
  if (card && c) { const it = F.items[+card.dataset.i]; if (!it) return;
    if (c.dataset.c === 'del') { freeItem(it); F.items.splice(F.items.indexOf(it), 1); if (!F.items.length) { F.screen = 'land'; F.alerts = []; X.toast('All files removed', 'info'); } render(); }
    else if (c.dataset.c === 'rot') { it.rot = (it.rot + 90) % 360; const th = $('.xf-thumb', card); if (th) th.style.transform = `rotate(${it.rot}deg)`; }
    else if (c.dataset.c === 'zoom') zoom(it);
    return; }
  if (pc) { const k = +pc.dataset.p, en = F.pages[k];
    if (c && c.dataset.c === 'prot') { en.rot = (en.rot + 90) % 360; $('.xf-thumb', pc).style.transform = `rotate(${en.rot}deg)`; syncSide(); return; }
    en.sel = !en.sel; pc.classList.toggle('sel', en.sel); pc.classList.toggle('gone', t.selVerb === 'remove' && en.sel); pc.setAttribute('aria-checked', en.sel); syncSide(); return; }
  const b = e.target.closest('[data-x]'); if (!b) return;
  const x = b.dataset.x;
  if (x === 'pick') choose();
  else if (x === 'camera') choose(true);
  else if (x === 'recent') addRecent(b);
  else if (x === 'useopen') useOpenDoc();
  else if (x === 'alt') t.alt.run();
  else if (x === 'begin') t.begin();
  else if (x === 'run') { if (b.getAttribute('aria-disabled') === 'true') { if (F.items.length < t.min) flashAlert(blocker()); else X.toast(blocker(), 'info'); } else run(); }
  else if (x === 'cancel') stopRun();
  else if (x === 'side') toggleSide();
  else if (x === 'sort') { F.items.sort((a, c2) => a.name.localeCompare(c2.name, undefined, { numeric: true, sensitivity: 'base' })); render(); }
  else if (x === 'unlock') doUnlock(+b.dataset.id);
  else if (x === 'ofile') pickFiles(b.dataset.acc, false).then(([f]) => { if (f) { F.opts[b.dataset.k] = f; render(); } });
  else if (x === 'psel') { const v = b.dataset.v; F.pages.forEach((p, k) => { p.sel = v === 'all' || (v === 'odd' && k % 2 === 0) || (v === 'even' && k % 2 === 1); }); render(); }
  else if (x === 'rotsel') { const d = +b.dataset.d, any = F.pages.some(p => p.sel); F.pages.forEach(p => { if (!any || p.sel) p.rot = ((p.rot + d) % 360 + 360) % 360; }); render(); }
  else if (x === 'rotreset') { F.pages.forEach(p => { p.rot = 0; }); render(); }
  else if (x === 'download') X.download(F.result.blob, F.result.name);
  else if (x === 'edit') openResult();
  else if (x === 'print') printResult();
  else if (x === 'copy') copyResult();
  else if (x === 'restart') { resetFiles(); F.opts = optsFor(t); F.screen = 'land'; render(); }
  else if (x === 'back') { F.screen = 'work'; F.result = null; render(); }
  else if (x === 'all') { leave(); }
}
function stopRun(nav) { // Cancel; or leaving while it runs (the tool closes once its work has ended and the editor's document is back)
  if (!F.running) return; F.cancel = true; if (nav) F.leaveAfter = true; else F.userCancel = true;
  if (F.screen !== 'proc') F.screen = 'proc'; render();
}
// "Continue to…": the next tool gets the result; if it cannot take the result, the original files (and says so); else nothing
const fits = (n, k) => n.mode === 'start' ? false : n.kinds.includes(k);
function carryOf(n) {
  const r = F.result; if (!r) return { from: null };
  if (r.kind === 'zip' ? fits(n, r.inner) : fits(n, kindOf(r.file))) return { from: 'result' };
  const ins = F.items.filter(i => fits(n, i.kind)); if (ins.length && ins.length === F.items.length) return { from: 'input', label: ins.length === 1 ? ins[0].name : 'the original files' };
  return { from: null };
}
async function continueTo(n) {
  const r = F.result, w = carryOf(n); let files = null;
  if (w.from === 'result' && r.kind === 'zip') { const JSZip = await X.loadLib('jszip'), z = await JSZip.loadAsync(r.blob), mime = { pdf: 'application/pdf', img: 'image/png' }[r.inner] || '';
    files = []; for (const k of Object.keys(z.files)) if (!z.files[k].dir) { const nm = k.split('/').pop(), f = new File([await z.files[k].async('blob')], nm, { type: /\.jpe?g$/i.test(nm) ? 'image/jpeg' : mime }); if (fits(n, kindOf(f))) files.push(f); } }
  else if (w.from === 'result') files = [r.file];
  else if (w.from === 'input') files = F.items.map(i => i.file);
  open(n.id, files && files.length ? files : null);
}
function toggleSide(on = !F.sideOpen) {
  F.sideOpen = on; const w = $('.xf-work', root), g = $('.xf-gear', root); if (w) w.classList.toggle('side-open', on);
  if (g) { g.innerHTML = `<i data-icon="${on ? 'x' : 'sliders'}"></i>`; g.title = on ? 'Close options' : 'Options'; X.icons(g); }
}
async function doUnlock(id) {
  const it = F.items.find(i => i.id === id), inp = $(`[data-unlock="${id}"]`, root); if (!it || !inp) return;
  const okk = await unlockItem(it, inp.value); it.bad = !okk;
  if (okk) { if (F.tool.mode === 'pages' && F.items[0] === it) F.pages = Array.from({ length: it.n || 0 }, (_, i) => ({ i, rot: 0, sel: false })); }
  render(); if (okk) drawThumb(it); else { const i2 = $(`[data-unlock="${id}"]`, root); if (i2) i2.focus(); }
}
function syncSide() { // the selection count, range field and button state after a page click
  const t = F.tool, n = F.pages.length, s = F.pages.filter(p => p.sel).length, cnt = $('.xf-count', root), rg = $('[data-x=prange]', root), cta = $('.xf-cta', root);
  if (cnt) cnt.textContent = `${s} of ${plural(n, 'page')} selected${t.selVerb ? ` to ${t.selVerb}` : ''}`; if (rg && document.activeElement !== rg) rg.value = rangeText();
  if (cta) { const why = blocker(); if (why) { cta.setAttribute('aria-disabled', 'true'); cta.title = why; } else { cta.removeAttribute('aria-disabled'); cta.removeAttribute('title'); } }
}
function onInput(e) {
  const el = e.target;
  if (el.dataset.x === 'prange') { const idx = el.value.trim() ? X.parseRange(el.value.replace(/\s*,\s*$/, ''), F.pages.length) : []; if (idx) { F.pages.forEach((p, k) => { p.sel = idx.includes(k); }); $$('.xf-pcard', root).forEach((c, k) => { c.classList.toggle('sel', F.pages[k].sel); c.classList.toggle('gone', F.tool.selVerb === 'remove' && F.pages[k].sel); }); syncSide(); } el.classList.toggle('bad', !idx); return; }
  const k = el.dataset.k; if (!k) return;
  F.opts[k] = el.type === 'checkbox' ? el.checked : el.value;
  if (F.tool.pv) paintPreviews();
  const d = F.tool.options.find(o => o.k === k), redraw = el.type === 'radio' || el.type === 'checkbox' || el.tagName === 'SELECT';
  if (k === 'show') $$('.xf-side input[type=password], .xf-side input[data-k=pw], .xf-side input[data-k=pw2]', root).forEach(i => { i.type = el.checked ? 'text' : 'password'; });
  if (redraw && e.type === 'change' && d) { const side = $('.xf-opts', root), y = side ? side.scrollTop : 0; render(); const s2 = $('.xf-opts', root); if (s2) s2.scrollTop = y; const back = $(`[data-k="${k}"]${el.type === 'radio' ? `[value="${CSS.escape(el.value)}"]` : ''}`, root); if (back) back.focus(); }
}
// drag to reorder file cards (mouse), Ctrl+←/→ (keyboard)
let dragI = -1;
function onDragStart(e) { if (td) { e.preventDefault(); return; } const c = e.target.closest && e.target.closest('.xf-card'); if (!c) return; dragI = +c.dataset.i; c.classList.add('dragging'); e.dataTransfer.effectAllowed = 'move'; e.dataTransfer.setData('text/x-xd3-card', String(dragI)); }
function onDragOver(e) { if (dragI < 0) return; e.preventDefault(); const c = e.target.closest('.xf-card'); $$('.drop-before,.drop-after', root).forEach(x => x.classList.remove('drop-before', 'drop-after')); if (c && +c.dataset.i !== dragI) { const r = c.getBoundingClientRect(); c.classList.add(e.clientX < r.left + r.width / 2 ? 'drop-before' : 'drop-after'); } }
function onDragDrop(e) { if (dragI < 0) return; e.preventDefault(); e.stopPropagation(); const c = e.target.closest('.xf-card'); if (c) { const to = +c.dataset.i, after = c.classList.contains('drop-after'); move(dragI, after ? to + (to < dragI ? 1 : 0) : to - (to > dragI ? 1 : 0)); } dragI = -1; }
function onDragEnd() { dragI = -1; if (root) $$('.dragging,.drop-before,.drop-after', root).forEach(x => x.classList.remove('dragging', 'drop-before', 'drop-after')); }
// touch: hold a card for a moment, then drag it (otherwise the page scrolls as usual)
let td = null;
const clearDrop = keep => $$('.drop-before,.drop-after', root).forEach(x => { if (x !== keep) x.classList.remove('drop-before', 'drop-after'); });
function tdEnd(drop) {
  if (!td) return; clearTimeout(td.timer); const d = td; td = null;
  d.el.classList.remove('dragging', 'lift'); d.el.style.transform = '';
  const tgt = root && $('.drop-before, .drop-after', root), after = !!tgt && tgt.classList.contains('drop-after'); if (root) clearDrop();
  if (drop && d.on && tgt) { const to = +tgt.dataset.i; move(d.i, after ? to + (to < d.i ? 1 : 0) : to - (to > d.i ? 1 : 0)); }
}
function onPDown(e) {
  if (e.pointerType === 'mouse' || F.running || F.screen !== 'work' || F.items.length < 2) return;
  const c = e.target.closest('.xf-card'); if (!c || e.target.closest('[data-c]')) return;
  tdEnd(false); td = { i: +c.dataset.i, el: c, x: e.clientX, y: e.clientY, on: false };
  td.timer = setTimeout(() => { if (!td) return; td.on = true; c.classList.add('dragging', 'lift'); try { if (navigator.vibrate) navigator.vibrate(12); } catch (er) { /* none */ } }, 250);
}
function onPMove(e) {
  if (!td) return; const dx = e.clientX - td.x, dy = e.clientY - td.y;
  if (!td.on) { if (Math.abs(dx) + Math.abs(dy) > 10) tdEnd(false); return; } // (a scroll)
  e.preventDefault(); td.el.style.transform = `translate(${dx}px, ${dy}px) scale(1.04)`;
  td.el.style.pointerEvents = 'none'; const under = document.elementFromPoint(e.clientX, e.clientY); td.el.style.pointerEvents = ''; // (the carried card is under the finger)
  const c = under && under.closest('.xf-card');
  clearDrop(c);
  if (c && c !== td.el) { const r = c.getBoundingClientRect(), before = e.clientX < r.left + r.width / 2; c.classList.toggle('drop-before', before); c.classList.toggle('drop-after', !before); }
}
function move(from, to) { to = Math.max(0, Math.min(F.items.length - 1, to)); if (from === to) return render(); const [it] = F.items.splice(from, 1); F.items.splice(to, 0, it); render(); const c = $(`.xf-card[data-i="${to}"]`, root); if (c) c.focus(); }
function onKey(e) {
  if (!visible) return; e.stopPropagation(); // the editor behind must not see these keys
  const k = e.key, typing = /INPUT|TEXTAREA|SELECT/.test(e.target.tagName);
  if ((e.ctrlKey || e.metaKey) && k.toLowerCase() === 'o') { e.preventDefault(); if (F.screen === 'land' || F.screen === 'work') choose(); return; }
  if ((e.ctrlKey || e.metaKey) && /^[spfhzy]$/i.test(k)) { e.preventDefault(); return; }
  if (F.running) return;
  const card = e.target.closest && e.target.closest('.xf-card'), pc = e.target.closest && e.target.closest('.xf-pcard');
  if (card && (e.ctrlKey || e.metaKey) && (k === 'ArrowLeft' || k === 'ArrowRight')) { e.preventDefault(); const i = +card.dataset.i; move(i, i + (k === 'ArrowLeft' ? -1 : 1)); return; }
  if (card && (k === 'Delete' || k === 'Backspace')) { e.preventDefault(); const del = $('[data-c=del]', card); if (del) del.click(); return; }
  if (pc && (k === ' ' || k === 'Enter')) { e.preventDefault(); pc.click(); return; }
  if (k === 'Escape' && !typing) { e.preventDefault(); back(); return; }
  if (k === 'Enter' && !typing && F.screen === 'land' && !(e.target.closest && e.target.closest('button'))) { e.preventDefault(); if (F.tool.mode === 'start') F.tool.begin(); else choose(); }
}
function back() { // Esc: one screen back — but never drops the files that were added
  const ov = $('.xf-zoom, .xf-pop', root); if (ov) { ov.remove(); return; }
  if (F.screen === 'done') { F.screen = 'work'; F.result = null; render(); }
  else if (F.screen === 'work') { if (F.sideOpen) toggleSide(false); }
  else leave();
}
function openResult() {
  const r = F.result; if (!r) return;
  if (r.kind === 'docx') { guardWord().then(ok => { if (ok) { leave(); X.docx.openFile(r.file); } }); return; }
  toEditor([r.file]);
}
function printResult() {
  const url = URL.createObjectURL(F.result.blob); let f = document.getElementById('printframe'); if (f) f.remove();
  f = document.createElement('iframe'); f.id = 'printframe'; f.style.cssText = 'position:fixed;right:0;bottom:0;width:0;height:0;border:0';
  f.onload = () => setTimeout(() => { try { f.contentWindow.focus(); f.contentWindow.print(); } catch (e) { window.open(url, '_blank'); } }, 300); f.src = url; document.body.appendChild(f);
}
async function copyResult() {
  const r = F.result;
  try { if (r.kind === 'txt') await navigator.clipboard.writeText(r.text); else { let b = r.blob; if (b.type !== 'image/png') { const bm = await createImageBitmap(b), c = document.createElement('canvas'); c.width = bm.width; c.height = bm.height; c.getContext('2d').drawImage(bm, 0, 0); b = await toBlob(c, 'image/png'); } await navigator.clipboard.write([new ClipboardItem({ 'image/png': b })]); } X.toast('Copied to the clipboard', 'success'); }
  catch (e) { X.toast('Copying is not available here — use Download instead', 'warn'); }
}

/* ───────── showing / hiding, routes ───────── */
function build() {
  root = document.createElement('section'); root.id = 'xflow'; root.className = 'xf'; root.hidden = true; root.tabIndex = -1; root.setAttribute('aria-label', 'Tool');
  root.innerHTML = '<div class="xf-body"></div><div class="xf-veil" hidden><div>Drop to add</div></div>';
  document.body.appendChild(root);
  root.addEventListener('click', onClick); root.addEventListener('input', onInput); root.addEventListener('change', onInput); root.addEventListener('keydown', onKey);
  root.addEventListener('keydown', e => { if (e.target.dataset && e.target.dataset.unlock && e.key === 'Enter') { e.preventDefault(); doUnlock(+e.target.dataset.unlock); } }, true);
  root.addEventListener('contextmenu', e => { if (e.target.closest('input, textarea')) return; e.preventDefault(); e.stopPropagation(); }); // (no editor menu over a tool)
  root.addEventListener('pointerdown', onPDown); root.addEventListener('pointermove', onPMove); root.addEventListener('pointerup', () => tdEnd(true)); root.addEventListener('pointercancel', () => tdEnd(false));
  root.addEventListener('touchmove', e => { if (td && td.on) e.preventDefault(); }, { passive: false }); // (no scrolling while a card is carried)
  root.addEventListener('dragstart', onDragStart); root.addEventListener('dragover', onDragOver); root.addEventListener('drop', onDragDrop); root.addEventListener('dragend', onDragEnd);
  // while a tool is shown, keys, dropped files and pasted pictures are ours — not the hidden editor's
  window.addEventListener('keydown', e => { if (boxed && !root.contains(e.target) && e.key !== 'Tab') { e.preventDefault(); e.stopImmediatePropagation(); return; } // (the editor's document is swapped out)
    if (!visible || root.contains(e.target) || (document.getElementById('dlg') || {}).open) return; const pal = document.getElementById('palette'); if (pal && !pal.hidden) return; if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'k') return; if (e.target !== document.body && e.target !== document.documentElement) return; onKey(e); e.stopImmediatePropagation(); }, true);
  const files = e => e.dataTransfer && [...e.dataTransfer.types].includes('Files'); let depth = 0;
  const v = () => $('.xf-veil', root), dropOk = () => visible && !F.running && (F.screen === 'land' || F.screen === 'work') && F.tool.mode !== 'start';
  window.addEventListener('dragenter', e => { if (!visible || !files(e)) return; e.stopImmediatePropagation(); if (dropOk()) { depth++; v().hidden = false; } }, true);
  window.addEventListener('dragleave', e => { if (!visible || !files(e)) return; e.stopImmediatePropagation(); if (--depth <= 0) { depth = 0; v().hidden = true; } }, true);
  window.addEventListener('dragover', e => { if ((!visible && !boxed) || !files(e)) return; e.preventDefault(); e.stopImmediatePropagation(); }, true);
  window.addEventListener('drop', e => { if ((!visible && !boxed) || !files(e)) return; e.preventDefault(); e.stopImmediatePropagation(); depth = 0; v().hidden = true; const dv = document.getElementById('dropveil'); if (dv) dv.classList.remove('on'); if (dropOk()) addFiles([...e.dataTransfer.files]); }, true);
  window.addEventListener('paste', e => { if (!visible && !boxed) return; if (/INPUT|TEXTAREA/.test((document.activeElement || {}).tagName || '')) { e.stopImmediatePropagation(); return; } e.stopImmediatePropagation(); const fl = [...(e.clipboardData ? e.clipboardData.files : [])]; if (fl.length && dropOk()) { e.preventDefault(); addFiles(fl); } }, true);
  window.addEventListener('resize', place);
  const re = () => { if (visible && F.screen === 'work') render(); }; window.addEventListener('resize', () => { if (F.tool && F.tool.pv) paintPreviews(); }); if (NARROW.addEventListener) NARROW.addEventListener('change', re); else NARROW.addListener(re);
  document.addEventListener('xd3:opened', () => { if (visible && !F.running) leave(); });
  new MutationObserver(() => { if (!visible || F.running) return; const b = document.body.classList; if (b.contains('dxopen')) leave(); else if (b.contains('modal')) { const w = document.getElementById('writer'), o = document.getElementById('org'); if ((w && !w.hidden) || (o && !o.hidden)) leave(); } }).observe(document.body, { attributes: true, attributeFilter: ['class'] });
}
function place() { // below the app header, whatever its height
  if (!root) return; const hd = document.getElementById('topbar') || document.querySelector('body > header'); let top = 0;
  if (hd) { const r = hd.getBoundingClientRect(), cs = getComputedStyle(hd); if (cs.display !== 'none' && r.height) top = Math.max(0, r.bottom); }
  root.style.top = top + 'px'; root.style.setProperty('--xf-top', top + 'px');
}
let goHome0 = null;
function show(id, files) {
  const t = byId[id]; if (!t) return;
  if (!root) build();
  if (!visible) { const gh = goHome0 || (X.home && X.home.goHome); if (gh && !document.body.classList.contains('athome')) gh(); }
  if (F.tool !== t || files || !visible) { resetFiles(); F.tool = t; F.opts = optsFor(t); F.screen = 'land'; F.sideOpen = false; } // (a tool opened afresh starts empty)
  visible = true; root.hidden = false; document.body.classList.add('xflow'); place(); noteRecent(id);
  document.title = t.title + ' — XD3 PDF Editor';
  render(); try { root.focus({ preventScroll: true }); } catch (e) { root.focus(); }
  if (files && files.length) addFiles(files);
}
function hide() {
  if (!visible) return;
  if (F.running) return stopRun(true);
  visible = false; if (root) { root.hidden = true; $$('.xf-pop,.xf-zoom', root).forEach(x => x.remove()); }
  document.body.classList.remove('xflow'); resetFiles(); F.sideOpen = false;
  document.title = 'XD3 PDF Editor — Free PDF Editor by XD3Labs'; if (X.updateUI) X.updateUI();
}
function leave() { // close the tool; the URL loses its #/tool part
  hide(); if (route()) history.replaceState(null, '', location.pathname + location.search);
}
const route = () => { const m = /^#\/([\w-]+)/.exec(location.hash); return m && byId[m[1]] ? m[1] : null; };
function sync() {
  const id = route();
  if (F.running) { // not while a tool runs: it stops first (see run())
    if (id === F.tool.id) { if (F.leaveAfter) { F.leaveAfter = false; F.cancel = !!F.userCancel; render(); } }
    else stopRun(true);
    return;
  }
  if (id) show(id, F.carry); else hide(); F.carry = null;
}
function open(id, files) {
  if (!byId[id]) return false; if (F.running) return false;
  if (route() === id) { show(id, files); return true; }
  F.carry = files || null; location.hash = '#/' + id; return true;
}
function noteRecent(id) { let l = []; try { l = JSON.parse(store.get('xd3.recentTools', '[]')) || []; } catch (e) { l = []; } l = [id, ...l.filter(x => x !== id && byId[x])].slice(0, 6); store.set('xd3.recentTools', JSON.stringify(l)); }
function recent() { let l = []; try { l = JSON.parse(store.get('xd3.recentTools', '[]')) || []; } catch (e) { l = []; } return l.map(id => byId[id]).filter(Boolean); }

window.addEventListener('hashchange', sync);
addEventListener('DOMContentLoaded', () => {
  goHome0 = X.home && X.home.goHome;
  const h = X.ACT.home; if (h) X.ACT.home = (...a) => { if (F.running) { F.after = () => { leave(); h(...a); }; return stopRun(true); } if (visible) leave(); return h(...a); }; // the logo / Home button closes the tool
  if (route()) sync();
});

X.flow = { CATS, GROUPS, TOOLS, byId, open, recent, close: () => leave(), add: files => addFiles(files), run: () => run(), state: F, unlockBytes, isOpen: () => visible };
document.dispatchEvent(new CustomEvent('xd3:flowready', { detail: X.flow }));
})();
