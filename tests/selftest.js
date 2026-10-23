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
