/* XD3 PDF Editor — automated checks. Open index.html?selftest=1 (or run "Run Tests.bat"); the desktop build runs them with XD3_SELFTEST=1.
   Each check opens a sample from tests/samples and verifies the editor's behaviour; results show in a dialog and in window.__selftest. */
(() => {
'use strict';
const X = window.XD3, { S } = X, results = [], t0 = performance.now();
const ok = (name, cond, detail = '') => { results.push({ name, ok: !!cond, detail: String(detail) }); console[cond ? 'log' : 'error'](`${cond ? 'PASS' : 'FAIL'} ${name} ${detail}`); };
const sleep = ms => new Promise(r => setTimeout(r, ms));
const load = async (name, type = 'application/pdf') => { const res = await fetch('tests/samples/' + name); if (!res.ok) throw new Error('sample missing: ' + name); S.dirty = false; await X.openFiles([new File([await res.blob()], name, { type })], false); const p = S.pages[0]; X.visible.add(p.id); return p; };
const items = async p => { X.setTool('edittext'); await X.ocrItems(p); await sleep(50); await X.drawTI(p); return X.PE.get(p.id).items || []; };
const reopen = bytes => pdfjsLib.getDocument({ data: bytes.slice(), cMapUrl: 'lib/cmaps/', cMapPacked: true, standardFontDataUrl: 'lib/standard_fonts/', isEvalSupported: false, fontExtraProperties: true }).promise;
const pixelDiff = async (a, b, n = 1) => { // pages rendered at 3× and compared: number of clearly different pixels
  const draw = async bytes => { const pdf = await reopen(bytes), pg = await pdf.getPage(n), vp = pg.getViewport({ scale: 3 }), c = document.createElement('canvas'); c.width = vp.width; c.height = vp.height; const x = c.getContext('2d', { willReadFrequently: true }); x.fillStyle = '#fff'; x.fillRect(0, 0, c.width, c.height); await pg.render({ canvasContext: x, viewport: vp, intent: 'print' }).promise; return x.getImageData(0, 0, c.width, c.height).data; };
  const A = await draw(a), B = await draw(b); if (A.length !== B.length) return -1; let d = 0; for (let i = 0; i < A.length; i += 4) if (Math.abs(A[i] - B[i]) + Math.abs(A[i + 1] - B[i + 1]) + Math.abs(A[i + 2] - B[i + 2]) > 60) d++; return d;
};
const fontNames = async pdf => { const pg = await pdf.getPage(1), tc = await pg.getTextContent(); await pg.getOperatorList(); const m = {}; for (const it of tc.items) if (it.str.trim()) { try { m[it.str] = pg.commonObjs.get(it.fontName).name; } catch (e) { /* */ } } return m; };
const textOf = async (pdf, i = 1) => (await (await pdf.getPage(i)).getTextContent()).items.map(t => t.str).join(' ');

async function run() {
  // start from the default view, whatever this browser remembers from earlier use (reading mode, two pages, night colours)
  for (const k of ['xd3.reader', 'xd3.spread', 'xd3.night']) { try { localStorage.removeItem(k); } catch (e) { /* */ } }
  document.body.classList.remove('reader', 'spread', 'night');
  // ── native text: fonts are reused, line width kept, saved file re-embeds them ──
  let p = await load('native.pdf'), its = await items(p);
  ok('native: text lines found', its.filter(i => !i.ocr && !i.field).length >= 9, its.length);
  ok('native: bookmarks read', S.outline.length === 2 && S.outline[0].title === 'Invoice', JSON.stringify(S.outline.map(e => e.title)));
  const acme = its.find(i => i.str === 'Acme Trading'); let o = acme && await X.editOriginal(p, acme);
  ok('native: embedded bold font reused', o && /Lato/.test(X.fontLabel(o.font)) && o.bold, o && X.fontLabel(o.font));
  ok('native: width matches original', o && Math.abs(X.bbox(o).w - acme.w) < 0.6, o && (X.bbox(o).w - acme.w).toFixed(2));
  const just = its.find(i => i.str.startsWith('This line is justified')); o = just && await X.editOriginal(p, just);
  ok('native: justified spacing kept', o && o.ws > 0.3 && Math.abs(X.bbox(o).w - just.w) < 0.6, o && o.ws);
  const rot = its.find(i => i.rot); o = rot && await X.editOriginal(p, rot); ok('native: rotated text editable', o && o.rot === rot.rot, rot && rot.rot);
  const pic = its.find(i => i.ocr && /Text inside/.test(i.str)); ok('native: text inside picture recognised', !!pic, its.filter(i => i.ocr).map(i => i.str).join(' | '));
  if (pic) { o = await X.editOriginal(p, pic); ok('native: picture text fitted bold', o.bold && o.size > 18 && o.size < 25, `${o.font} ${o.size}`); o.text = 'Changed'; }
  // a sticky note + a rectangle → PDF comments
  p.objs.push(X.textObj('Please review', 12, { x: 60, y: 300, note: true, bg: '#fef08a' }), { id: X.uid(), type: 'rect', x: 50, y: 400, w: 200, h: 60, color: '#dc2626', fill: 'none', width: 2, opacity: 1 });
  let bytes = await X.buildPdf(S.pages), pdf = await reopen(bytes), pg = await pdf.getPage(1), ann = await pg.getAnnotations();
  ok('save: file rebuilt', bytes.length > 50000 && pdf.numPages === 2, bytes.length);
  ok('save: comments written', ann.some(a => a.subtype === 'Text') && ann.some(a => a.subtype === 'Square'), ann.map(a => a.subtype).join(','));
  ok('save: bookmarks written', (await pdf.getOutline() || []).length === 2);
  const txt = await textOf(pdf); ok('save: edited text searchable', /Changed/.test(txt) && /Acme Trading/.test(txt));
  await pg.getOperatorList(); const names = new Set(); for (const it of (await pg.getTextContent()).items) { try { names.add(pg.commonObjs.get(it.fontName).name); } catch (e) { /* */ } }
  ok('save: document font embedded again', [...names].some(n => /Lato/.test(n)), [...names].join(','));

  // ── no trace of editing: unchanged edits are invisible, changed text uses the original font, file info is kept ──
  const origBytes = new Uint8Array(await (await fetch('tests/samples/native.pdf')).arrayBuffer());
  p = await load('native.pdf'); its = await items(p);
  for (const it of its.filter(i => !i.ocr)) await X.editOriginal(p, it);
  bytes = await X.buildPdf(S.pages); ok('no trace: re-typed lines are pixel-identical to the original', await pixelDiff(origBytes, bytes) === 0);
  const before = await fontNames(await reopen(origBytes));
  p.objs.find(o => o.type === 'text' && o.text === 'Acme Trading').text = 'Acme Holdings';
  p.objs.find(o => o.type === 'text' && o.text === 'Standard Times text that is not embedded').text = 'Standard Times text, edited';
  bytes = await X.buildPdf(S.pages); pdf = await reopen(bytes); let tx = await textOf(pdf);
  ok('no trace: old words removed from the file', !/Acme Trading|not embedded/.test(tx) && /Acme Holdings/.test(tx) && /text, edited/.test(tx), tx.slice(0, 80));
  const after = await fontNames(pdf);
  ok('no trace: new words use the original font resource', after['Acme Holdings'] === before['Acme Trading'] && after['Standard Times text, edited'] === 'Times-Roman', `${after['Acme Holdings']} / ${before['Acme Trading']}`);
  ok('no trace: no extra fonts embedded', new Set(Object.values(after)).size <= new Set(Object.values(before)).size, [...new Set(Object.values(after))].join(','));
  ok('no trace: producer and file name kept', (await pdf.getMetadata()).info.Producer === (await (await reopen(origBytes)).getMetadata()).info.Producer && S.fileName === 'native.pdf', (await pdf.getMetadata()).info.Producer);
  ok('no trace: saved file is not bigger than needed', bytes.length < origBytes.length * 1.15, `${bytes.length} vs ${origBytes.length}`);
  // a real-world file (Chrome / Edge print, two-byte CID fonts, kerning, ligatures): every line rewritten stays visually the same
  if ((await fetch('tests/samples/edge.pdf', { method: 'HEAD' })).ok) {
    const eb = new Uint8Array(await (await fetch('tests/samples/edge.pdf')).arrayBuffer()); p = await load('edge.pdf'); its = await items(p); const eo = [];
    for (const it of its.filter(i => !i.ocr)) eo.push(await X.editOriginal(p, it)); eo.forEach(o => { o.orig = null; }); // (forces every line to be written anew)
    bytes = await X.buildPdf(S.pages); const dpx = await pixelDiff(eb, bytes);
    ok('CID fonts: rewritten page matches the original', dpx >= 0 && dpx < 612 * 792 * 9 * 0.001, dpx + ' px (anti-aliasing only)'); // (< 0.1 % of the page: one-pixel snapping, invisible)
    eo.find(o => /^Quarterly/.test(o.text)).text = 'Quarterly Report 2027'; bytes = await X.buildPdf(S.pages); tx = (await textOf(await reopen(bytes))).replace(/\s+/g, '');
    ok('CID fonts: changed heading saved, old one gone, no cover boxes', /Report2027/.test(tx) && !/2026/.test(tx), tx.slice(0, 40));
  }
  // delete a line through the right-click action, delete page 2 → neither survives in the file
  p = await load('native.pdf'); its = await items(p); S.ctxTI = { pid: p.id, i: its.findIndex(i => i.str === 'dated today') }; await X.ACT.ti_del();
  S.pages.splice(1, 1); X.buildPages(); bytes = await X.buildPdf(S.pages); pdf = await reopen(bytes);
  ok('delete text: line removed from the file', !/dated today/.test(await textOf(pdf)));
  let pageObjs = 0; for (const [, o] of (await PDFLib.PDFDocument.load(bytes, { updateMetadata: false })).context.enumerateIndirectObjects()) if (o instanceof PDFLib.PDFDict && String(o.get(PDFLib.PDFName.of('Type'))) === '/Page') pageObjs++;
  ok('delete page: removed page is not left inside the file', pdf.numPages === 1 && pageObjs === 1, pageObjs);
  // links and crop
  p = S.pages[0]; p.objs.push({ id: X.uid(), type: 'link', x: 58, y: 530, w: 100, h: 14, url: 'https://xd3labs.example/', opacity: 1 }); p.crop = { x: 40, y: 40, w: 400, h: 500 };
  pdf = await reopen(await X.buildPdf(S.pages)); pg = await pdf.getPage(1); ann = await pg.getAnnotations();
  ok('link: saved as a link annotation', ann.some(a => a.subtype === 'Link' && a.url === 'https://xd3labs.example/'), ann.filter(a => a.subtype === 'Link').map(a => a.url).join(','));
  const vw = pg.getViewport({ scale: 1 }); ok('crop: page trimmed to the box', Math.round(vw.width) === 400 && Math.round(vw.height) === 500, `${vw.width}×${vw.height}`);

  // ── blackout: the words under the box are removed from the file; the rest stays real text ──
  p = await load('native.pdf'); p.objs.push({ id: X.uid(), type: 'rect', cover: true, redact: true, x: 55, y: 50, w: 300, h: 18, color: 'none', fill: '#000000', width: 0, opacity: 1 });
  pdf = await reopen(await X.buildPdf(S.pages)); tx = await textOf(pdf);
  ok('redaction: blacked-out words removed, page stays text', !/Invoice|Acme/.test(tx) && /quick brown/.test(tx), tx.slice(0, 50));
  p.objs.push({ id: X.uid(), type: 'rect', cover: true, redact: true, x: 70, y: 370, w: 120, h: 30, color: 'none', fill: '#000000', width: 0, opacity: 1 });
  pdf = await reopen(await X.buildPdf(S.pages)); ok('redaction: blackout over a picture makes that page a picture', (await textOf(pdf)).trim() === '');
  pdf = await reopen(await X.buildPdf(S.pages, { redact: 'cover' })); ok('redaction: "keep editable" keeps text', /Invoice/.test(await textOf(pdf)));

  // ── scans: OCR, font fit, searchable layer ──
  p = await load('scan.pdf'); its = await items(p);
  ok('scan: all lines recognised', its.length === 6, its.map(i => i.str).join(' | '));
  const tot = its.find(i => /Total amount/.test(i.str)); o = tot && await X.editOriginal(p, tot);
  ok('scan: sans line fitted regular', o && !o.bold && o.size > 16 && o.size < 19, o && `${o.font} ${o.bold} ${o.size}`);
  const code = its.find(i => /Reference code/.test(i.str)); o = code && await X.editOriginal(p, code); ok('scan: monospace line fitted', o && /Mono|Courier/.test(o.font), o && o.font);
  await X.makeSearchable; p.ocr = (await X.ocrItems(p, true)).map(it => ({ str: it.str, x: it.x, y: it.y, w: it.w, h: it.h, base: it.base, size: it.size }));
  pdf = await reopen(await X.buildPdf(S.pages)); const layer = await textOf(pdf);
  ok('scan: searchable layer written', /agreement/i.test(layer) && /12,500/.test(layer), layer.slice(0, 60));

  p = await load('skew.pdf'); its = await items(p);
  ok('skew: crooked scan read', its.length >= 5 && its.every(i => Math.abs((i.rot || 0) + 2) < 0.6), its.map(i => (i.rot || 0) + ':' + i.str.slice(0, 10)).join(' | '));
  if (its[1]) { o = await X.editOriginal(p, its[1]); ok('skew: line placed at its angle', o && Math.abs(o.rot - its[1].rot) < 0.01 && o.size > 14 && o.size < 19, o && `${o.rot} ${o.size}`); }

  p = await load('searchable.pdf'); its = await items(p);
  ok('hidden layer: invisible text detected', its.length === 2 && its.every(i => i.hidden), its.map(i => i.hidden).join(','));

  // ── forms ──
  p = await load('form.pdf'); its = await items(p); const fields = its.filter(i => i.field);
  ok('form: fields outlined', fields.length === 4, fields.map(f => f.ft + ':' + f.name).join(','));
  const city = fields.find(f => f.name === 'city'); if (city) { await X.fillField(p, city); X.commitEdit(); const t = p.objs.find(x => x.type === 'text'); ok('form: text field prefilled', t && t.text === 'Pune', t && t.text); }
  const agree = fields.find(f => f.name === 'agree'); if (agree) { await X.fillField(p, agree); ok('form: checkbox ticked', p.objs.some(x => x.text === '✓')); }
  pdf = await reopen(await X.buildPdf(S.pages)); const wid = (await (await pdf.getPage(1)).getAnnotations()).filter(a => a.subtype === 'Widget'), val = n => (wid.find(a => a.fieldName === n) || {}).fieldValue;
  ok('form: values saved into the real fields', val('city') === 'Pune' && val('agree') && val('agree') !== 'Off' && wid.length === 4, `${val('city')} ${val('agree')} ${wid.length}`);

  // ── reading mode: selectable text, links, highlight from a selection; Add text: click, type, click again ──
  p = await load('native.pdf'); X.setZoom(1); if (!document.body.classList.contains('reader')) X.ACT.reader(); await sleep(1200);
  let pe = X.PE.get(p.id), tl = pe.wrap.querySelector('.textLayer');
  ok('reader: text layer rendered', tl && tl.querySelectorAll('span').length > 20, tl && tl.querySelectorAll('span').length);
  ok('reader: links placed', pe.wrap.querySelectorAll('.plink').length === 2, pe.wrap.querySelectorAll('.plink').length);
  if (tl) { X.goToPage(0); const sp = [...tl.querySelectorAll('span')].find(s => /quick brown/.test(s.textContent)), rg = document.createRange(); rg.selectNodeContents(sp); getSelection().removeAllRanges(); getSelection().addRange(rg); await sleep(400);
    ok('reader: selection pop-up', !document.getElementById('selpop').hidden); document.querySelector('#selpop [data-p=hl]').click(); await sleep(100);
    const hl = p.objs.find(o => o.blend); ok('reader: highlight from selection', hl && Math.abs(hl.x - 60) < 4 && hl.w > 280, hl && `${hl.x.toFixed(1)} ${hl.w.toFixed(1)}`); }
  X.ACT.reader(); ok('reader: back to editing', !document.body.classList.contains('reader') && S.tool === 'select');
  X.setTool('text'); pe = X.PE.get(p.id); const down = (x, y) => { const r = pe.svg.getBoundingClientRect(); pe.svg.dispatchEvent(new PointerEvent('pointerdown', { bubbles: true, button: 0, clientX: r.left + x, clientY: r.top + y })); };
  down(200, 600); await sleep(50); document.execCommand('insertText', false, 'First box'); down(200, 650); await sleep(50); document.execCommand('insertText', false, 'Second box'); X.commitEdit();
  const boxes = p.objs.filter(o => o.type === 'text' && /box$/.test(o.text)).map(o => o.text);
  ok('add text: a click while typing finishes the box and starts the next', boxes.length === 2 && boxes[1] === 'Second box', boxes.join(','));

  // ── PDF → Word "same as the PDF": the page design behind, every piece of text in a text box at its exact place ──
  if ((await fetch('tests/samples/layout.pdf', { method: 'HEAD' })).ok) {
    const JSZip = await X.loadLib('jszip'), docOf = async blob => { const z = await JSZip.loadAsync(blob); return { z, xml: await z.file('word/document.xml').async('string') }; };
    p = await load('layout.pdf');
    let { z, xml } = await docOf(await X.buildDocx(S.pages)); const boxes = (xml.match(/<wps:txbx>/g) || []).length, media = Object.keys(z.files).filter(n => n.startsWith('word/media/') && !n.endsWith('/'));
    const words = s => s.replace(/<[^>]+>/g, '').replace(/\s+/g, ''); // (justified lines are split into words, so compare without spaces)
    ok('to Word: text boxes, page design and all text', boxes > 15 && media.length === 1 && ['AnnualPerformanceReview', 'CONFIDENTIAL', '4,250,000', 'Smallsteps'].every(t => words(xml).includes(t)) && /w:pgSz w:w="11899"/.test(xml), `${boxes} boxes, ${media.length} picture(s)`);
    const tc = await (await X.getPdfPage(p)).getTextContent(), it = tc.items.find(t => t.str.startsWith('1. Summary')), m = /<wp:positionH relativeFrom="page"><wp:posOffset>(-?\d+)<\/wp:posOffset><\/wp:positionH><wp:positionV relativeFrom="page"><wp:posOffset>(-?\d+)<\/wp:posOffset>(?:(?!<wp:positionH).)*?w:line="(\d+)"(?:(?!<wp:positionH).)*?>1\.(?:(?!<wp:positionH).)*?Summary/.exec(xml);
    const bx = m && +m[1] / 12700, base = m && +m[2] / 12700 + 0.8 * +m[3] / 20, pb = it && 842 - it.transform[5];
    ok('to Word: a heading lands on its exact spot', m && it && Math.abs(bx - it.transform[4]) < 0.5 && Math.abs(base - pb) < 0.5, m && it && `dx ${(bx - it.transform[4]).toFixed(2)}, dy ${(base - pb).toFixed(2)}`);
    const line = (await X.getLines(p)).find(l => /Customer satisfaction/.test(l.str)); await X.editOriginal(p, line, 'Customer happiness reached a record');
    ({ xml } = await docOf(await X.buildDocx(S.pages)));
    ok('to Word: edits are converted, replaced text is gone', words(xml).includes('Customerhappinessreachedarecord') && !words(xml).includes('Customersatisfactionrose'));
    p = await load('searchable.pdf'); ({ xml } = await docOf(await X.buildDocx(S.pages)));
    ok('to Word: a scan keeps its look, its hidden text stays hidden', /w14:alpha w14:val="100000"/.test(xml) && words(xml).includes('Totalamountpayable') && /Page design/.test(xml));
  }
  // ── compress: pictures resampled to where they are shown, duplicates merged, transparency and text untouched ──
  {
    const L = PDFLib, doc = await L.PDFDocument.create(), canvasBytes = async (w, h, draw, type, q) => { const c = document.createElement('canvas'); c.width = w; c.height = h; draw(c.getContext('2d')); return new Uint8Array(await (await new Promise(r => c.toBlob(r, type, q))).arrayBuffer()); };
    const photo = await canvasBytes(1800, 1200, x => { for (let i = 0; i < 1500; i++) { x.fillStyle = `hsl(${i % 360},70%,${30 + (i * 7) % 50}%)`; x.fillRect((i * 733) % 1800, (i * 397) % 1200, 40, 40); } }, 'image/jpeg', 0.95);
    const logo = await canvasBytes(700, 700, x => { x.fillStyle = '#16a34a'; x.beginPath(); x.arc(350, 350, 300, 0, 7); x.fill(); }, 'image/png'), font = await doc.embedFont(L.StandardFonts.Helvetica);
    for (let i = 0; i < 2; i++) { const pg = doc.addPage([595, 842]); pg.drawRectangle({ x: 0, y: 600, width: 595, height: 242, color: L.rgb(0.95, 0.85, 0.2) }); pg.drawImage(await doc.embedPng(logo), { x: 40, y: 640, width: 160, height: 160 }); pg.drawImage(await doc.embedJpg(photo), { x: 60, y: 260, width: 300, height: 200 }); pg.drawText('Compression check ' + (i + 1), { x: 60, y: 220, size: 14, font }); }
    const src = await doc.save(), st = await X.compressPdf(src), back = await L.PDFDocument.load(st.bytes), imgs = [];
    for (const [, o] of back.context.enumerateIndirectObjects()) if (o instanceof L.PDFRawStream && o.dict.lookup(L.PDFName.of('Subtype')) === L.PDFName.of('Image')) imgs.push(o);
    const masks = imgs.filter(o => !o.dict.has(L.PDFName.of('SMask')) && o.dict.lookup(L.PDFName.of('ColorSpace')) === L.PDFName.of('DeviceGray'));
    const draw1 = async b => { const pg = await (await reopen(b)).getPage(1), vp = pg.getViewport({ scale: 1 }), c = document.createElement('canvas'); c.width = vp.width; c.height = vp.height; const x = c.getContext('2d', { willReadFrequently: true }); await pg.render({ canvasContext: x, viewport: vp, intent: 'print' }).promise; return x.getImageData(0, 0, c.width, c.height).data; };
    const A1 = await draw1(src), B1 = await draw1(st.bytes); let d = 0; for (let i = 0; i < A1.length; i += 4) if (Math.abs(A1[i] - B1[i]) + Math.abs(A1[i + 1] - B1[i + 1]) + Math.abs(A1[i + 2] - B1[i + 2]) > 90) d++; // (at screen size: a 150 dpi copy looks the same there)
    ok('compress: pictures resampled, duplicates merged', st.after < src.length * 0.2 && imgs.length === 3 && st.dups >= 3, `${src.length} → ${st.after} bytes, ${imgs.length} pictures left`);
    ok('compress: transparency and text kept', masks.length === 1 && (await textOf(await reopen(st.bytes), 2)) === 'Compression check 2' && d < 0.01 * A1.length / 4, `${d} of ${A1.length / 4} pixels differ`);
    const scan = new Uint8Array(await (await fetch('tests/samples/scan.pdf')).arrayBuffer()), ss = await X.compressPdf(scan);
    ok('compress: a 200 dpi scan shrinks to screen resolution', ss.after < scan.length * 0.5 && (await textOf(await reopen(ss.bytes))) === (await textOf(await reopen(scan))), `${scan.length} → ${ss.after}`);
    const form = new Uint8Array(await (await fetch('tests/samples/form.pdf')).arrayBuffer()), fs = await X.compressPdf(form), fields = async b => (await (await (await reopen(b)).getPage(1)).getAnnotations()).filter(a => a.fieldType).length;
    ok('compress: form fields survive', (await fields(fs.bytes)) === (await fields(form)) && (await fields(form)) > 0);
  }
  // ── writing tools: offline spelling & grammar, humanize, Word round trip with formatting kept, fixes written back into the PDF ──
  if (X.writing && (await fetch('tests/samples/writing.docx', { method: 'HEAD' })).ok) {
    const WR = X.writing, W = WR.W, settle = async () => { await sleep(200); for (let k = 0; k < 100 && W.busyLint; k++) await sleep(100); };
    X.ACT.writer(); await WR.openFile(new File([await (await fetch('tests/samples/writing.docx')).blob()], 'writing.docx')); await settle();
    const found = W.paras.flatMap(q => (q.issues || []).map(i => i.word));
    ok('writing: spelling and grammar found offline', ['Quartely', 'importent', 'sentense', 'The the'].every(w => found.includes(w)), found.join(', '));
    for (let k = 0; k < 12; k++) { const btn = document.querySelector('#wrPane .wcard .wps button'); if (!btn) break; btn.click(); await sleep(150); } await settle();
    const xml = await (await (await X.loadLib('jszip')).loadAsync(await WR.buildDocx())).file('word/document.xml').async('string');
    ok('writing: Word file fixed with its formatting kept', /<w:b\/><\/w:rPr><w:t[^>]*>important<\/w:t>/.test(xml) && /Quarterly Report/.test(xml) && /<w:i\/><\/w:rPr><w:t[^>]*>utilize<\/w:t>/.test(xml) && /Heading1/.test(xml));
    const rw = WR.rewrites('natural').map(r => r.from + '>' + r.to);
    ok('writing: humanize suggests natural wording', ['Furthermore>Also', 'utilize>use', 'in order to>to'].every(x => rw.includes(x)), rw.join(' | '));
    W.paras[3].text = "The report isn't finished and we can't wait."; const fm = WR.rewrites('formal').filter(r => r.i === 3).map(r => r.to).join(','); ok('writing: formal tone expands contractions', fm === 'is not,cannot', fm);
    WR.close(); p = await load('native.pdf'); X.ACT.writer(); document.querySelector('[data-w=pdf]').click(); await settle(); await settle();
    const q = W.paras.find(x => /lazy dog/.test(x.text)); q.text = q.text.replace('lazy dog', 'sleepy dog'); await WR.applyToPdf(); WR.close();
    tx = await textOf(await reopen(await X.buildPdf(S.pages)));
    ok('writing: fix written back into the PDF in its own font', /sleepy dog/.test(tx) && !/lazy dog/.test(tx) && /running through the quiet forest/.test(tx));
  }
  // ── viewing: two pages side by side ──
  X.ACT.spread(); ok('view: two pages side by side', document.body.classList.contains('spread')); X.ACT.spread();

  // ── Word editor: a full document opens, saves untouched run for run, edits land in the file, Fix Formatting and Auto Format ──
  if (X.docx && (await fetch('tests/samples/complex.docx', { method: 'HEAD' })).ok) {
    const DX = X.docx, JSZip = await X.loadLib('jszip'), WNS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main', ser = new XMLSerializer();
    const sample = async n => new File([await (await fetch('tests/samples/' + n, { cache: 'no-store' })).blob()], n);
    const docXml = async blob => (await JSZip.loadAsync(blob)).file('word/document.xml').async('string');
    const runsSig = xml => { // every paragraph as its formatted pieces of text (runs with the same formatting side by side count as one)
      const d = new DOMParser().parseFromString(xml, 'application/xml'), clean = s => s.replace(/ xmlns(:\w+)?="[^"]*"/g, '').replace(/ w:rsid\w*="[^"]*"/g, '');
      return [...d.getElementsByTagNameNS(WNS, 'p')].map(p => { const out = []; for (const r of p.getElementsByTagNameNS(WNS, 'r')) { const rp = [...r.children].find(c => c.localName === 'rPr'), k = rp ? clean(ser.serializeToString(rp)) : '', t = [...r.children].filter(c => c.localName !== 'rPr').map(c => c.localName === 't' ? c.textContent : '<' + c.localName + '>').join(''); const last = out[out.length - 1]; if (last && last[0] === k) last[1] += t; else out.push([k, t]); } return JSON.stringify(out); });
    };
    const P = () => DX.allParas(), find = t => P().findIndex(p => DX.ptext(p).startsWith(t)), wait = sleep;
    DX.D.dirty = false; await DX.openFile(await sample('complex.docx')); const dxDoc = document.querySelector('#dxDoc');
    ok('word: full document opens', P().length > 30 && dxDoc.querySelectorAll('table.dx-t').length === 1 && dxDoc.querySelectorAll('.dx-img img').length === 1 && dxDoc.querySelector('.dx-txbx') && document.querySelector('#dxHdr').textContent.includes('XD3 Test Header') && P().some(p => p.dataset.label === '2.'), `${P().length} paragraphs`);
    const orig = await docXml(await sample('complex.docx')), same = await docXml(await DX.buildDocx()), A = runsSig(orig), B = runsSig(same);
    const diff = A.map((s, i) => s === B[i] ? null : i).filter(i => i !== null);
    ok('word: untouched save keeps every run and paragraph', A.length === B.length && !diff.length, `${A.length}/${B.length} ${diff.slice(0, 3).map(i => B[i]).join(' ')}`);
    let i = find('This report has'); DX.setSel({ a: [i, 5], b: [i, 11] }); DX.setRunProp('b', true, true);
    i = find('Between the lists'); DX.setSel({ a: [i, 0], b: [i, 0] }); DX.setStyle('heading 2');
    i = find('Third bullet'); DX.setSel({ a: [i, 12], b: [i, 12] }); document.execCommand('insertParagraph'); await wait(30); document.execCommand('insertText', false, 'Fourth bullet'); await wait(30);
    i = find('Text on the second page'); DX.setSel({ a: [i, 24], b: [i, 24] }); document.execCommand('insertParagraph'); await wait(30);
    i = find('A sentence with a'); DX.setSel({ a: [i, 0], b: [i, 26] }); document.execCommand('delete'); await wait(50);
    let x = await docXml(await DX.buildDocx());
    ok('word: edits saved in the Word file', /<w:b\/><w:bCs\/><\/w:rPr><w:t[^>]*>report<\/w:t>/.test(x) && /<w:pStyle w:val="Heading2"\/><\/w:pPr><w:r><w:t[^>]*>Between the lists/.test(x) && /<w:numId w:val="1"\/><\/w:numPr><\/w:pPr><w:r><w:t[^>]*>Fourth bullet/.test(x));
    const count = (re, s) => (s.match(re) || []).length, ids = (n, s) => [...s.matchAll(new RegExp(`<w:${n} [^>]*w:id="(\\d+)"`, 'g'))].map(m => m[1]).sort().join(',');
    ok('word: fields, bookmarks and sections stay whole', count(/fldCharType="begin"/g, x) === count(/fldCharType="end"/g, x) && ids('bookmarkStart', x) === ids('bookmarkEnd', x) && ids('commentRangeStart', x) === ids('commentRangeEnd', x) && count(/<w:sectPr/g, x) === 2, `${count(/<w:sectPr/g, x)} sections`);
    DX.undo(); DX.undo(); ok('word: undo', P().some(p => DX.ptext(p).replace(/￼/g, '').startsWith('A sentence with a comment')), P().length);
    // spelling and grammar right in the page
    DX.D.dirty = false; await DX.newDocument(); document.querySelector('#dxDoc').focus(); document.execCommand('insertText', false, 'This is an importent sentense.'); await wait(50); await DX.lintAll(); await wait(100);
    const iss = (P()[0]._issues || []).map(s => s.word); ok('word: spelling marked in the page', iss.includes('importent') && (!window.Highlight || CSS.highlights.has('dx-spell')), iss.join(','));
    // Fix Formatting and Auto Format
    DX.D.dirty = false; await DX.openFile(await sample('messy.docx'));
    const keys = DX.checks().map(c => c.key); ok('word: Fix Formatting finds the problems', ['headings', 'lists', 'fonts', 'sizes', 'spaces', 'empty', 'indent', 'breaks', 'quotes', 'align'].every(k => keys.includes(k)), keys.join(','));
    DX.applyFixes(DX.checks()); await wait(100); x = await docXml(await DX.buildDocx());
    ok('word: Fix all leaves a tidy document', !DX.checks().length && /<w:pStyle w:val="Title"\/>/.test(x) && /<w:pStyle w:val="Heading1"\/>/.test(x) && count(/<w:numPr>/g, x) >= 5 && !/ {2,}/.test(x.replace(/<[^>]+>/g, '')), DX.checks().map(c => c.key).join(','));
    DX.D.dirty = false; await DX.openFile(await sample('messy.docx'));
    let st = 'pending'; DX.autoFormat().then(() => { st = 'done'; }); await wait(150); const dlgEl = document.querySelector('#dlg');
    dlgEl.querySelector('[name=set]').value = 'office'; dlgEl.querySelector('[name=colors]').checked = true; dlgEl.querySelector('button[value=ok]').click(); await wait(80);
    if (st === 'pending') dlgEl.dispatchEvent(new Event('close')); // (a hidden test window does not deliver the dialog's close event)
    for (let k = 0; k < 60 && st === 'pending'; k++) await wait(250);
    const hs = [...document.querySelectorAll('#dxDoc p[data-h]')].map(p => p.dataset.h + ':' + p.textContent).join('|');
    ok('word: Auto Format styles and tidies the document', st === 'done' && /0:ANNUAL SUMMARY/.test(hs) && /1:Background/.test(hs) && /1:Key Points/.test(hs) && !DX.checks().length, hs);
    DX.close(true);
  }
  // ── desktop app: the Windows bridge answers (only inside the .exe) ──
  if (window.xd3desk) {
    const info = await xd3desk.info(), st = await xd3desk.pdfDefault(); ok('desktop: Windows bridge', info && info.os && st && typeof st.mine === 'boolean', JSON.stringify({ os: info && info.os.name, pdf: st && st.app }));
    X.ACT.defapp(); for (let k = 0; k < 20 && !document.querySelector('#dlg[open] #daStat'); k++) await sleep(100);
    const da = document.querySelector('#dlg[open] #daStat'); ok('desktop: Default PDF app dialog', da && /XD3 PDF Editor opens|PDFs now open with/.test(da.textContent), da && da.textContent);
    const dlgEl = document.querySelector('#dlg'); if (dlgEl.open) { dlgEl.close(); dlgEl.dispatchEvent(new Event('close')); }
    if (X.docx) { // the Word editor's "PDF" button: the page printed by the app itself
      X.docx.D.dirty = false; await X.docx.openFile(new File([await (await fetch('tests/samples/complex.docx')).blob()], 'complex.docx'));
      document.documentElement.classList.add('dxprint'); await sleep(200);
      let pdf = null, t = ''; try { pdf = await reopen(new Uint8Array(await xd3desk.printToPdf({ width: 8.5, height: 11, margins: { top: 1, bottom: 1, left: 1, right: 1 } }))); for (let k = 1; k <= pdf.numPages; k++) t += await textOf(pdf, k) + ' '; } finally { document.documentElement.classList.remove('dxprint'); }
      ok('desktop: Word document saved as an exact PDF', pdf && pdf.numPages >= 2 && /Project Report/.test(t) && /landscape sec/.test(t) && !/Auto Format|Spelling/.test(t), pdf && pdf.numPages + ' pages: ' + t.replace(/\s+/g, ' ').slice(-160)); // (the Calibri 'ti' ligature reads back as one odd character)
      X.docx.close(true);
    }
  }

  // ── session round trip ──
  await X.openFiles; S.dirty = true; X.scheduleSave(); await sleep(7000); await X.ACT.restore(); await sleep(300);
  ok('session: restored with edits', S.pages.length === 2 && S.pages[0].objs.some(o => /sleepy dog/.test(o.text || '')), S.pages[0] && S.pages[0].objs.length);
}
window.__selftest = run().catch(e => { ok('suite crashed', false, e && e.stack || e); }).then(() => {
  const fails = results.filter(r => !r.ok), secs = ((performance.now() - t0) / 1000).toFixed(1);
  const html = `<p><b>${results.length - fails.length} / ${results.length} checks passed</b> in ${secs} s.</p><div class="keys" style="grid-template-columns:auto 1fr">${results.map(r => `<kbd style="color:${r.ok ? '#16a34a' : '#dc2626'}">${r.ok ? 'PASS' : 'FAIL'}</kbd><span>${X.esc(r.name)}${r.detail ? ` <small style="color:var(--mut)">— ${X.esc(r.detail.slice(0, 160))}</small>` : ''}</span>`).join('')}</div>`;
  X.dialog(fails.length ? 'Self-test: problems found' : 'Self-test passed', html, { cancel: false });
  return { passed: results.length - fails.length, total: results.length, fails: fails.map(f => f.name + (f.detail ? ' — ' + f.detail : '')) };
});
})();
