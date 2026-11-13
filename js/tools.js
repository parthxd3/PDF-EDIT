/* XD3 PDF Editor — more tools for the tabs: page batch operations (reverse, rotate all, remove blank pages, split every N pages),
   PDF → PowerPoint and PDF → Excel, picture extraction, and removing hidden data (metadata, scripts, attachments, comments).
   Each tool also has a dialog-free core (called with an options object) that returns its result instead of downloading it — used by flow.js. */
(() => {
'use strict';
const X = window.XD3, { S, toast, busy } = X, dialog = (...a) => X.dialog(...a), download = (...a) => X.download(...a);
const xesc = s => String(s).replace(/[\x00-\x08\x0b\x0c\x0e-\x1f]/g, '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const H = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const zipBlob = (zip, mime) => zip.generateAsync({ type: 'blob', compression: 'DEFLATE', mimeType: mime });
const put = (zip, n, d, o) => zip.file(n, d, { createFolders: false, ...o });
// options of the cores: pages (default: the open document), progress(i, n) before each page, check() throws to cancel
const pagesOf = o => (o && o.pages) || S.pages;
const step = (o, i, n) => { if (o && o.check) o.check(); if (o && o.progress) o.progress(i, n); };

/* ───────── pages ───────── */
function reversePages(opt) { if (opt && opt.pages) return opt.pages.reverse(); if (S.pages.length < 2) return toast('There is only one page'); X.pageOp(() => S.pages.reverse()); toast('Page order reversed — Ctrl+Z to undo', 'success'); }
const turn = (pages, d) => pages.forEach(p => { p.rot = (((p.rot || 0) + d) % 360 + 360) % 360; });
function rotateAll(d = 90, opt) { if (opt && opt.pages) return turn(opt.pages, d); X.pageOp(() => turn(S.pages, d)); toast('All pages rotated', 'success'); }
async function findBlank(opt) { // → the pages with (almost) no ink; opt.ink: the share of inked pixels below which a page counts as blank
  const blank = [], pages = pagesOf(opt), lim = opt && opt.ink != null ? opt.ink : 0.0015;
  for (let k = 0; k < pages.length; k++) {
    step(opt, k, pages.length); const p = pages[k]; if (p.objs.length) continue;
    const c = await X.renderPageImage(p, 0.35), d = c.getContext('2d', { willReadFrequently: true }).getImageData(0, 0, c.width, c.height).data; let ink = 0;
    for (let i = 0; i < d.length; i += 4) if (d[i] < 235 || d[i + 1] < 235 || d[i + 2] < 235) ink++;
    if (ink < c.width * c.height * lim) blank.push(p);
  }
  return blank;
}
async function removeBlank() {
  let blank = [];
  await busy(async () => { blank = await findBlank(); }, 'Looking for blank pages…');
  if (!blank.length) return toast('No blank pages found', 'info');
  if (blank.length === S.pages.length) return toast('Every page looks blank — nothing was removed', 'info');
  if (!confirm(`Remove ${blank.length} blank page(s)?`)) return;
  X.pageOp(() => { S.pages = S.pages.filter(p => !blank.includes(p)); S.cur = Math.min(S.cur, S.pages.length - 1); }); X.updateUI(); toast(`${blank.length} blank page(s) removed — Ctrl+Z to undo`, 'success');
}
async function splitParts(opt) { // opt.n pages in each part, or opt.groups [[page…]…] → { blob (ZIP of PDFs), parts }
  const pages = pagesOf(opt), JSZip = await X.loadLib('jszip'), zip = new JSZip(), n = Math.max(1, +opt.n || 1);
  const groups = opt.groups || Array.from({ length: Math.ceil(pages.length / n) }, (_, i) => pages.slice(i * n, i * n + n)), len = String(groups.length).length;
  for (let i = 0; i < groups.length; i++) {
    step(opt, i, groups.length); const g = groups[i], a = pages.indexOf(g[0]) + 1, b = pages.indexOf(g[g.length - 1]) + 1;
    put(zip, `${X.baseName()}-part${String(i + 1).padStart(len, '0')} (pages ${a === b ? a : a + '-' + b}).pdf`, await X.buildPdf(g));
  }
  return { blob: await zipBlob(zip, 'application/zip'), parts: groups.length };
}
async function splitEvery() {
  const r = await dialog('Split into parts', `<label>Pages in each part<input type="number" name="n" value="${Math.min(10, Math.max(1, Math.ceil(S.pages.length / 2)))}" min="1" max="${S.pages.length}"></label><p class="mut">${S.pages.length} pages → saved as a ZIP of PDFs. To save chosen pages as one PDF, use Pages ▸ Extract.</p>`, { ok: 'Split' });
  if (!r) return; const n = Math.max(1, Math.min(S.pages.length, +r.n || 1)); if (n >= S.pages.length) return toast('That is the whole document already', 'info');
  await busy(async () => { const o = await splitParts({ n }); download(o.blob, X.baseName() + '-parts.zip'); toast(`Split into ${o.parts} PDFs (ZIP)`, 'success'); }, 'Splitting…');
}

/* ───────── PDF → PowerPoint: one slide per page ───────── */
async function toPptx(opt) { // with opt { dpi, q, pages… }: no dialog, → the .pptx Blob
  if (opt) return pptxOf(opt);
  const r = await dialog('PDF → PowerPoint (.pptx)', `<label>Quality<select name="q"><option value="150|0.85">Standard (150 dpi)</option><option value="220|0.9" selected>Sharp (220 dpi)</option><option value="110|0.8">Small file (110 dpi)</option></select></label><p class="mut">Every page becomes a slide, shown exactly as in the PDF. To change the text, use PDF → Word.</p>`, { ok: 'Convert' });
  if (!r) return; const [dpi, q] = r.q.split('|').map(Number);
  await busy(async () => { download(await pptxOf({ dpi, q }), X.baseName() + '.pptx'); toast(`PowerPoint file saved — ${S.pages.length} slide(s)`, 'success'); }, 'Making slides…');
}
async function pptxOf(opt) {
  const pages = pagesOf(opt), dpi = +opt.dpi || 220, q = +opt.q || 0.9, JSZip = await X.loadLib('jszip');
  const zip = new JSZip(), [W0, H0] = X.rdims(pages[0]), cx = Math.round(W0 * 12700), cy = Math.round(H0 * 12700), P = 'http://schemas.openxmlformats.org/presentationml/2006/main', A = 'http://schemas.openxmlformats.org/drawingml/2006/main', R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships', RT = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/', CT = 'application/vnd.openxmlformats-officedocument.presentationml';
  const n = pages.length;
  put(zip, '[Content_Types].xml', `${H}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="jpeg" ContentType="image/jpeg"/><Override PartName="/ppt/presentation.xml" ContentType="${CT}.presentation.main+xml"/><Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="${CT}.slideMaster+xml"/><Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="${CT}.slideLayout+xml"/><Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>${Array.from({ length: n }, (_, i) => `<Override PartName="/ppt/slides/slide${i + 1}.xml" ContentType="${CT}.slide+xml"/>`).join('')}</Types>`);
  put(zip, '_rels/.rels', `${H}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${RT}officeDocument" Target="ppt/presentation.xml"/></Relationships>`);
  put(zip, 'ppt/presentation.xml', `${H}<p:presentation xmlns:a="${A}" xmlns:r="${R}" xmlns:p="${P}" saveSubsetFonts="1"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst><p:sldIdLst>${Array.from({ length: n }, (_, i) => `<p:sldId id="${256 + i}" r:id="rId${10 + i}"/>`).join('')}</p:sldIdLst><p:sldSz cx="${cx}" cy="${cy}"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>`);
  put(zip, 'ppt/_rels/presentation.xml.rels', `${H}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${RT}slideMaster" Target="slideMasters/slideMaster1.xml"/><Relationship Id="rId2" Type="${RT}theme" Target="theme/theme1.xml"/>${Array.from({ length: n }, (_, i) => `<Relationship Id="rId${10 + i}" Type="${RT}slide" Target="slides/slide${i + 1}.xml"/>`).join('')}</Relationships>`);
  const tree = inner => `<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>${inner}</p:spTree></p:cSld>`;
  put(zip, 'ppt/slideMasters/slideMaster1.xml', `${H}<p:sldMaster xmlns:a="${A}" xmlns:r="${R}" xmlns:p="${P}">${tree('')}<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/><p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst></p:sldMaster>`);
  put(zip, 'ppt/slideMasters/_rels/slideMaster1.xml.rels', `${H}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${RT}slideLayout" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="rId2" Type="${RT}theme" Target="../theme/theme1.xml"/></Relationships>`);
  put(zip, 'ppt/slideLayouts/slideLayout1.xml', `${H}<p:sldLayout xmlns:a="${A}" xmlns:r="${R}" xmlns:p="${P}" type="blank" preserve="1">${tree('')}<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`);
  put(zip, 'ppt/slideLayouts/_rels/slideLayout1.xml.rels', `${H}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${RT}slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>`);
  const clr = (n2, v) => `<a:${n2}><a:srgbClr val="${v}"/></a:${n2}>`, fill = '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>', ln = '<a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>';
  put(zip, 'ppt/theme/theme1.xml', `${H}<a:theme xmlns:a="${A}" name="XD3"><a:themeElements><a:clrScheme name="XD3"><a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>${clr('dk2', '44546A')}${clr('lt2', 'E7E6E6')}${clr('accent1', '4472C4')}${clr('accent2', 'ED7D31')}${clr('accent3', 'A5A5A5')}${clr('accent4', 'FFC000')}${clr('accent5', '5B9BD5')}${clr('accent6', '70AD47')}${clr('hlink', '0563C1')}${clr('folHlink', '954F72')}</a:clrScheme><a:fontScheme name="XD3"><a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme><a:fmtScheme name="XD3"><a:fillStyleLst>${fill.repeat(3)}</a:fillStyleLst><a:lnStyleLst>${ln.repeat(3)}</a:lnStyleLst><a:effectStyleLst>${'<a:effectStyle><a:effectLst/></a:effectStyle>'.repeat(3)}</a:effectStyleLst><a:bgFillStyleLst>${fill.repeat(3)}</a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>`);
  for (let i = 0; i < n; i++) {
    step(opt, i, n);
    const p = pages[i], c = await X.renderPageImage(p, dpi / 72), b = new Uint8Array(await (await new Promise(res => c.toBlob(res, 'image/jpeg', q))).arrayBuffer());
    const [w, h] = X.rdims(p), f = Math.min(cx / (w * 12700), cy / (h * 12700)), pw = Math.round(w * 12700 * f), ph = Math.round(h * 12700 * f);
    put(zip, `ppt/media/page${i + 1}.jpeg`, b);
    put(zip, `ppt/slides/slide${i + 1}.xml`, `${H}<p:sld xmlns:a="${A}" xmlns:r="${R}" xmlns:p="${P}">${tree(`<p:pic><p:nvPicPr><p:cNvPr id="2" name="Page ${i + 1}"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rId2"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="${Math.round((cx - pw) / 2)}" y="${Math.round((cy - ph) / 2)}"/><a:ext cx="${pw}" cy="${ph}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`)}<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`);
    put(zip, `ppt/slides/_rels/slide${i + 1}.xml.rels`, `${H}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${RT}slideLayout" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="rId2" Type="${RT}image" Target="../media/page${i + 1}.jpeg"/></Relationships>`);
  }
  return zipBlob(zip, 'application/vnd.openxmlformats-officedocument.presentationml.presentation');
}

/* ───────── PDF → Excel: every line of text a row, its pieces in columns lined up across the page ───────── */
const colLetter = n => { let s = ''; n++; while (n) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = (n - m - 1) / 26; } return s; };
async function toXlsx(opt) { // with opt { mode: 'one' | 'each', pages… }: no dialog, → { blob, sheets } (throws when there is no text)
  if (opt) return xlsxOf(opt);
  const r = await dialog('PDF → Excel (.xlsx)', `<label>Sheets<select name="mode"><option value="one">All pages on one sheet</option><option value="each">A sheet for each page</option></select></label><p class="mut">Rows of the PDF become rows of the sheet; text that lines up in columns (tables) goes into the same column. Numbers become real numbers you can add up.</p>`, { ok: 'Convert' });
  if (!r) return;
  await busy(async () => { const o = await xlsxOf({ mode: r.mode }); download(o.blob, X.baseName() + '.xlsx'); toast(`Excel file saved — ${o.sheets} sheet(s)`, 'success'); }, 'Reading the tables…');
}
async function xlsxOf(opt) {
  const r = { mode: opt.mode || 'one' }, pages = pagesOf(opt), JSZip = await X.loadLib('jszip');
  const sheets = [];
  for (let i = 0; i < pages.length; i++) {
    step(opt, i, pages.length);
    const rows = X.rowsOf(await X.visibleText(pages[i])); if (!rows.length && r.mode !== 'each') continue;
    // a line of plain text sits in column A; in a run of table rows the n-th piece goes to the n-th column, and a row with empty cells
    // puts each piece under the nearest column of the fullest row of the run (so right-aligned numbers stay under their headings)
    const grid = [], mid = q => q.x + q.w / 2;
    for (let a = 0; a < rows.length;) {
      if (rows[a].parts.length < 2) { grid.push([rows[a].parts[0].str.trim()]); a++; continue; }
      let b = a; while (b < rows.length && rows[b].parts.length > 1) b++;
      const run = rows.slice(a, b), ref = run.reduce((m, r2) => r2.parts.length > m.parts.length ? r2 : m), n = ref.parts.length;
      const edge = (q, k) => { const c = ref.parts[k]; return Math.min(Math.abs(q.x - c.x), Math.abs(q.x + q.w - c.x - c.w), Math.abs(mid(q) - mid(c))); };
      for (const row of run) {
        if (row.parts.length >= n) { grid.push(row.parts.map(q => q.str.trim())); continue; }
        const cells = [], used = new Set();
        for (const q of row.parts) { let k = -1, d = Infinity; for (let j = 0; j < n; j++) if (!used.has(j) && edge(q, j) < d) { d = edge(q, j); k = j; } used.add(k); cells[k] = q.str.trim(); }
        grid.push(cells);
      }
      a = b;
    }
    if (r.mode === 'each' || !sheets.length) sheets.push({ name: r.mode === 'each' ? `Page ${i + 1}` : X.baseName().slice(0, 31) || 'Sheet1', rows: [] });
    const sh = sheets[sheets.length - 1]; if (r.mode !== 'each' && sh.rows.length) sh.rows.push([]); sh.rows.push(...grid);
  }
  if (!sheets.length || sheets.every(s => !s.rows.length)) throw new Error('No text found — scanned pages need OCR first (Convert ▸ Make searchable)');
  const num = v => { const t = v.replace(/^[₹$€£¥]\s*/, '').replace(/,/g, ''); return /^[-+−]?\d+(\.\d+)?$/.test(t) ? t.replace('−', '-') : null; };
  const zip = new JSZip(), CT = 'application/vnd.openxmlformats-officedocument.spreadsheetml', RT = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
  put(zip, '[Content_Types].xml', `${H}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="${CT}.sheet.main+xml"/>${sheets.map((_, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="${CT}.worksheet+xml"/>`).join('')}</Types>`);
  put(zip, '_rels/.rels', `${H}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${RT}officeDocument" Target="xl/workbook.xml"/></Relationships>`);
  const used = new Set(), sname = n => { let b = n.replace(/[\\/?*[\]:]/g, ' ').slice(0, 31) || 'Sheet', s = b, k = 2; while (used.has(s.toLowerCase())) s = b.slice(0, 28) + ' ' + k++; used.add(s.toLowerCase()); return s; };
  put(zip, 'xl/workbook.xml', `${H}<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${RT.slice(0, -1)}"><sheets>${sheets.map((s, i) => `<sheet name="${xesc(sname(s.name))}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('')}</sheets></workbook>`);
  put(zip, 'xl/_rels/workbook.xml.rels', `${H}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${sheets.map((_, i) => `<Relationship Id="rId${i + 1}" Type="${RT}worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join('')}</Relationships>`);
  sheets.forEach((s, i) => {
    const widths = []; s.rows.forEach(row => row.forEach((v, k) => { if (v) widths[k] = Math.max(widths[k] || 8, Math.min(60, v.length + 2)); }));
    const rowsXml = s.rows.map((row, ri) => `<row r="${ri + 1}">${row.map((v, k) => { if (v == null || v === '') return ''; const ref = colLetter(k) + (ri + 1), nv = num(v); return nv != null ? `<c r="${ref}"><v>${nv}</v></c>` : `<c r="${ref}" t="inlineStr"><is><t xml:space="preserve">${xesc(v)}</t></is></c>`; }).join('')}</row>`).join('');
    put(zip, `xl/worksheets/sheet${i + 1}.xml`, `${H}<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${widths.length ? `<cols>${widths.map((w, k) => w ? `<col min="${k + 1}" max="${k + 1}" width="${w}" customWidth="1"/>` : '').join('')}</cols>` : ''}<sheetData>${rowsXml}</sheetData></worksheet>`);
  });
  return { blob: await zipBlob(zip, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'), sheets: sheets.length };
}

/* ───────── pictures out of the PDF ───────── */
async function extractImages(opt) { // with opt { pages… }: → { blob (ZIP, null when there are none), n }, nothing downloaded
  if (opt) return imagesOf(opt);
  await busy(async () => {
    const o = await imagesOf({}); if (!o.n) return toast('This PDF has no pictures', 'info');
    download(o.blob, X.baseName() + '-pictures.zip'); toast(`${o.n} picture(s) saved as a ZIP`, 'success');
  }, 'Collecting pictures…');
}
async function imagesOf(opt) {
  const JSZip = await X.loadLib('jszip'), pages = pagesOf(opt), zip = new JSZip(); let n = 0;
  for (let i = 0; i < pages.length; i++) {
    step(opt, i, pages.length);
    const p = pages[i], boxes = p.src < 0 ? [] : (await X.scanPage(p)).all.filter(b => b.w > 24 && b.h > 24); if (!boxes.length) continue;
    const k = 200 / 72, c = await X.renderPageImage({ ...p, rot: 0, objs: [] }, k, { noText: true });
    for (const b of boxes) { const o = document.createElement('canvas'); o.width = Math.round(b.w * k); o.height = Math.round(b.h * k); o.getContext('2d').drawImage(c, Math.round(b.x * k), Math.round(b.y * k), o.width, o.height, 0, 0, o.width, o.height); put(zip, `page${i + 1}-picture${++n}.png`, await new Promise(res => o.toBlob(res, 'image/png'))); }
  }
  return { blob: n ? await zipBlob(zip, 'application/zip') : null, n };
}

/* ───────── hidden data ───────── */
async function sanitize(opt) { // with opt { meta, js, att, cmt, extra, pages… }: no dialog, → { blob, done }
  if (opt) return cleanOf(opt);
  const r = await dialog('Remove hidden data', `<p>Saves a clean copy without the things a reader does not see but could find in the file.</p>
    <label class="inl"><input type="checkbox" name="meta" checked>Title, author, program and other details (metadata)</label>
    <label class="inl"><input type="checkbox" name="js" checked>Scripts and automatic actions</label>
    <label class="inl"><input type="checkbox" name="att" checked>Attached files</label>
    <label class="inl"><input type="checkbox" name="cmt">Comments, highlights and notes</label>
    <label class="inl"><input type="checkbox" name="extra" checked>Page thumbnails and editing-program data</label>`, { ok: 'Save clean copy' });
  if (!r) return;
  await busy(async () => {
    const o = await cleanOf({ meta: !!r.meta, js: !!r.js, att: !!r.att, cmt: !!r.cmt, extra: !!r.extra }); download(o.blob, X.baseName() + '-clean.pdf');
    toast(`Clean copy saved — removed: ${o.done.join(', ') || 'nothing'}`, 'success');
  }, 'Cleaning…');
}
async function cleanOf(r) {
  const L = PDFLib, N = n => L.PDFName.of(n);
  const doc = await L.PDFDocument.load(await X.buildPdf(pagesOf(r)), { updateMetadata: false }), cat = doc.catalog, names = cat.lookup(N('Names')), done = [];
  if (r.meta) { doc.context.trailerInfo.Info = undefined; if (cat.has(N('Metadata'))) cat.delete(N('Metadata')); done.push('details'); }
  if (r.js) { if (names instanceof L.PDFDict && names.has(N('JavaScript'))) names.delete(N('JavaScript')); for (const k of ['OpenAction', 'AA']) if (cat.has(N(k))) cat.delete(N(k)); done.push('scripts'); }
  if (r.att && names instanceof L.PDFDict && names.has(N('EmbeddedFiles'))) { names.delete(N('EmbeddedFiles')); done.push('attachments'); }
  const CMT = new Set(['Text', 'FreeText', 'Highlight', 'Underline', 'StrikeOut', 'Squiggly', 'Ink', 'Square', 'Circle', 'Line', 'Polygon', 'PolyLine', 'Stamp', 'Caret', 'Popup', 'FileAttachment', 'Sound']);
  for (const page of doc.getPages()) {
    const node = page.node; if (r.extra) for (const k of ['Thumb', 'PieceInfo']) if (node.has(N(k))) node.delete(N(k)); if (r.js && node.has(N('AA'))) node.delete(N('AA'));
    const an = node.lookup(N('Annots')); if (!(an instanceof L.PDFArray)) continue; const keep = [];
    for (let i = 0; i < an.size(); i++) {
      const ref = an.get(i), a = an.lookup(i); if (!(a instanceof L.PDFDict)) { keep.push(ref); continue; }
      const sub = a.lookup(N('Subtype')), name = sub instanceof L.PDFName ? sub.asString().slice(1) : '';
      if (r.cmt && CMT.has(name)) continue; if (r.att && name === 'FileAttachment') continue;
      if (r.js) { const act = a.lookup(N('A')); if (act instanceof L.PDFDict && act.lookup(N('S')) === N('JavaScript')) a.delete(N('A')); if (a.has(N('AA'))) a.delete(N('AA')); }
      keep.push(ref);
    }
    if (keep.length !== an.size()) node.set(N('Annots'), doc.context.obj(keep));
  }
  if (r.cmt) done.push('comments'); if (r.extra) { if (cat.has(N('PieceInfo'))) cat.delete(N('PieceInfo')); done.push('editor data'); }
  if (X.native) X.native.collectGarbage(doc);
  return { blob: new Blob([await doc.save({ updateFieldAppearances: false })], { type: 'application/pdf' }), done };
}

Object.assign(X.ACT, { pgreverse: reversePages, pgrotall: () => rotateAll(90), pgblank: removeBlank, pgsplit: splitEvery, topptx: () => toPptx(), toxlsx: () => toXlsx(), extractimg: () => extractImages(), sanitize: () => sanitize() });
X.tools = { reversePages, rotateAll, removeBlank, findBlank, splitEvery, splitParts, toPptx, toXlsx, extractImages, sanitize };
})();
