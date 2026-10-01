/* XD3 PDF Editor — converters: PDF ⇄ Word (.docx), PDF → images, text export, flatten / compress, split. All offline. */
(() => {
'use strict';
const X = window.XD3, { S, toast, busy, dialog, clamp, uid, download } = X;
const pad = (n, len) => String(n).padStart(len, '0');
const twip = v => Math.round(v * 20), emu = v => Math.round(v * 12700);
const xesc = s => String(s).replace(/[\x00-\x08\x0b\x0c\x0e-\x1f]/g, '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const toBlob = (c, type, q) => new Promise(res => c.toBlob(res, type, q));
const pickPages = range => { const t = (range || '').trim(); if (!t) return S.pages.slice(); const idx = X.parseRange(t, S.pages.length); if (!idx) { toast(`Invalid page range. Use numbers 1–${S.pages.length}, e.g. 1-3, 5`, true); return null; } return idx.map(i => S.pages[i]); };
const zipSave = async (zip, name, mime) => download(await zip.generateAsync({ type: 'blob', compression: 'DEFLATE', mimeType: mime || 'application/zip' }), name);

/* ───────── what is on a page ───────── */
function rowsOf(lines) { // lines sharing a baseline → one row with its parts left-to-right (table cells, columns)
  const L = [...lines].sort((a, b) => a.base - b.base || a.x - b.x), rows = [];
  for (const l of L) {
    const r = rows[rows.length - 1];
    if (r && Math.abs(r.base - l.base) < 0.35 * Math.min(r.size, l.size)) { r.parts.push(l); r.size = Math.max(r.size, l.size); } else rows.push({ base: l.base, size: l.size, parts: [l] });
  }
  for (const r of rows) { r.parts.sort((a, b) => a.x - b.x); r.x = r.parts[0].x; r.right = Math.max(...r.parts.map(q => q.x + q.w)); r.top = r.base - 0.95 * r.size; r.bottom = r.top + 1.2 * r.size; }
  return rows;
}
function inkColor(ctx, k, l) { // text colour of an original line, sampled from a render
  try {
    const x = Math.max(0, Math.floor(l.x * k)), y = Math.max(0, Math.floor(l.y * k)), w = Math.min(ctx.canvas.width - x, Math.ceil(l.w * k)), h = Math.min(ctx.canvas.height - y, Math.ceil(l.h * k));
    if (w < 2 || h < 2) return '#000000'; const d = ctx.getImageData(x, y, w, h).data, b = [d[0], d[1], d[2]]; let md = 0, c = [0, 0, 0];
    for (let i = 0; i < d.length; i += 4) { const dd = Math.abs(d[i] - b[0]) + Math.abs(d[i + 1] - b[1]) + Math.abs(d[i + 2] - b[2]); if (dd > md) { md = dd; c = [d[i], d[i + 1], d[i + 2]]; } }
    if (md < 90 || Math.max(...c) < 70) return '#000000'; return '#' + c.map(v => v.toString(16).padStart(2, '0')).join('');
  } catch (e) { return '#000000'; }
}
async function pageContent(p) { // visible text lines (with style) + images, in page view coordinates
  const lines = await X.visibleText(p), images = [], K = 2;
  const orig = lines.filter(l => l.src === 'orig'), boxes = p.src >= 0 ? (await X.getImages(p)).filter(b => !X.covered(p, b)) : [];
  if (p.src >= 0 && (orig.length || boxes.length)) {
    const canvas = await X.renderPageImage({ ...p, rot: 0, objs: [] }, K), ctx = canvas.getContext('2d', { willReadFrequently: true }); // rendering also loads the fonts
    for (const l of orig) { const fs = await X.fontStyle(p, l.fontName); l.bold = fs.bold; l.italic = fs.italic; if (fs.family) l.font = fs.family; l.color = inkColor(ctx, K, l); }
    for (const b of boxes) {
      const c = document.createElement('canvas'); c.width = Math.max(1, Math.round(b.w * K)); c.height = Math.max(1, Math.round(b.h * K));
      c.getContext('2d').drawImage(canvas, Math.round(b.x * K), Math.round(b.y * K), c.width, c.height, 0, 0, c.width, c.height);
      images.push({ ...b, url: c.toDataURL('image/jpeg', 0.9) });
    }
  }
  for (const o of p.objs) if (o.type === 'image' && S.assets[o.asset]) images.push({ x: o.x, y: o.y, w: o.w, h: o.h, url: S.assets[o.asset] });
  return { lines, images };
}

/* ───────── PDF → DOCX ───────── */
const FAM = { Helvetica: 'Arial', Times: 'Times New Roman', Courier: 'Courier New' };
function runXml(l, text) {
  const f = FAM[l.font] || (X.BUNDLED[l.font] ? l.font : 'Arial'), col = l.color && l.color !== '#000000' ? `<w:color w:val="${l.color.slice(1).toUpperCase()}"/>` : '', sz = Math.max(2, Math.round(l.size * 2));
  const rpr = `<w:rPr><w:rFonts w:ascii="${f}" w:hAnsi="${f}" w:cs="${f}"/>${l.bold ? '<w:b/>' : ''}${l.italic ? '<w:i/>' : ''}${l.strike ? '<w:strike/>' : ''}${col}<w:sz w:val="${sz}"/><w:szCs w:val="${sz}"/>${l.underline ? '<w:u w:val="single"/>' : ''}</w:rPr>`;
  return `<w:r>${rpr}<w:t xml:space="preserve">${xesc(text)}</w:t></w:r>`;
}
function docxBuilder() {
  const zip = new JSZip(), rels = [], seen = new Map(); let n = 0, dp = 0;
  const media = url => { // → relationship id for an image data-URL (each distinct image stored once)
    if (seen.has(url)) return seen.get(url);
    const ext = /^data:image\/jpe?g/.test(url) ? 'jpeg' : 'png', id = 'rId' + (++n), name = `image${n}.${ext}`;
    zip.file('word/media/' + name, url.split(',')[1], { base64: true }); rels.push(`<Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/${name}"/>`);
    seen.set(url, id); return id;
  };
  const pic = (url, w, h, at) => { // inline picture, or floating at a page position when `at` is given
    const rid = media(url), id = ++dp, ext = `cx="${emu(w)}" cy="${emu(h)}"`;
    const g = `<wp:docPr id="${id}" name="Picture ${id}"/><wp:cNvGraphicFramePr/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="${id}" name="Picture ${id}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext ${ext}/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic>`;
    if (!at) return `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent ${ext}/><wp:effectExtent l="0" t="0" r="0" b="0"/>${g}</wp:inline></w:drawing></w:r>`;
    return `<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="${id}" behindDoc="1" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="page"><wp:posOffset>${emu(at.x)}</wp:posOffset></wp:positionH><wp:positionV relativeFrom="page"><wp:posOffset>${emu(at.y)}</wp:posOffset></wp:positionV><wp:extent ${ext}/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>${g}</wp:anchor></w:drawing></w:r>`;
  };
  const sect = (w, h, m) => `<w:sectPr><w:pgSz w:w="${twip(w)}" w:h="${twip(h)}"${w > h ? ' w:orient="landscape"' : ''}/><w:pgMar w:top="${twip(m.t)}" w:right="${twip(m.r)}" w:bottom="${twip(m.b)}" w:left="${twip(m.l)}" w:header="0" w:footer="0" w:gutter="0"/></w:sectPr>`;
  const finish = async body => {
    const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"';
    const H = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
    zip.file('[Content_Types].xml', `${H}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Default Extension="jpeg" ContentType="image/jpeg"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`);
    zip.file('_rels/.rels', `${H}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`);
    zip.file('word/_rels/document.xml.rels', `${H}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels.join('')}</Relationships>`);
    zip.file('word/document.xml', `${H}<w:document ${NS}><w:body>${body}</w:body></w:document>`);
    return zip.generateAsync({ type: 'blob', compression: 'DEFLATE', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
  };
  return { pic, sect, finish };
}
// each builder returns { paras: [{ ppr, body }], margins } for one page; the section break is attached to the page's last paragraph
function flowPage(pc, p, B) { // flowing, editable text — paragraphs are re-joined so Word can re-wrap them
  const rows = rowsOf(pc.lines), els = [...rows.map(r => ({ top: r.top, bottom: r.bottom, r })), ...pc.images.map(im => ({ top: im.y, bottom: im.y + im.h, im }))].sort((a, b) => a.top - b.top);
  const xs = [...rows.map(r => r.x), ...pc.images.map(i => i.x)], ML = clamp(xs.length ? Math.min(...xs) : 56, 14, p.w / 2);
  const CR = Math.min(p.w - 8, Math.max(ML + 120, ...rows.map(r => r.right), ...pc.images.map(i => i.x + i.w))), MT = clamp(els.length ? els[0].top : 56, 14, p.h / 2);
  const sig = l => [Math.round(l.size * 2), l.font, !!l.bold, !!l.italic, l.color || ''].join('|'), paras = []; let prevBottom = MT, cur = null;
  for (const e of els) {
    if (e.im) { cur = null; paras.push({ ppr: `<w:spacing w:before="${twip(clamp(e.top - prevBottom, 0, 500))}" w:after="0"/><w:ind w:left="${twip(Math.max(0, e.im.x - ML))}"/>`, body: B.pic(e.im.url, Math.min(e.im.w, CR - ML), e.im.h * Math.min(1, (CR - ML) / e.im.w)) }); prevBottom = Math.max(prevBottom, e.bottom); continue; }
    const r = e.r, f = r.parts[0], single = r.parts.length === 1;
    if (cur && single && cur.sig === sig(f) && r.base > cur.lastBase && r.base - cur.lastBase < 1.6 * r.size && cur.lastRight > CR - 0.15 * (CR - ML)
      && (cur.contX == null ? Math.abs(r.x - cur.x) < 40 : Math.abs(r.x - cur.contX) < 3)) { // the next line of the same paragraph
      cur.text += (/[-­]$/.test(cur.text) ? '' : ' ') + f.str; cur.gaps.push(r.base - cur.lastBase); cur.contX = r.x; cur.lastBase = r.base; cur.lastRight = r.right; prevBottom = r.bottom; continue;
    }
    let align = 'left'; const mid = (r.x + r.right) / 2;
    if (single && Math.abs(mid - p.w / 2) < 0.02 * p.w && r.x - ML > 24) align = 'center'; else if (single && Math.abs(r.right - CR) < 2 && r.x - ML > 0.4 * (CR - ML)) align = 'right';
    const P = { row: r, text: f.str, gaps: [], x: r.x, contX: null, lastBase: r.base, lastRight: r.right, sig: single && align === 'left' ? sig(f) : null, align, before: clamp(r.top - prevBottom, 0, 500) };
    paras.push(P); cur = P.sig ? P : null; prevBottom = r.bottom;
  }
  return {
    margins: { t: MT, l: ML, r: Math.max(8, p.w - CR - 6), b: 10 },
    paras: paras.map(P => {
      if (P.body) return P;
      const r = P.row, lineH = P.gaps.length ? P.gaps.reduce((a, b) => a + b, 0) / P.gaps.length : 1.2 * r.size, multi = r.parts.length > 1;
      const tabs = multi ? `<w:tabs>${r.parts.slice(1).map(q => `<w:tab w:val="left" w:pos="${twip(Math.max(1, q.x - ML))}"/>`).join('')}</w:tabs>` : '';
      const left = P.align !== 'left' ? 0 : P.contX != null ? P.contX - ML : P.x - ML, fl = P.align === 'left' && P.contX != null ? P.x - P.contX : 0;
      const ind = `<w:ind w:left="${twip(Math.max(0, left))}"${fl > 1 ? ` w:firstLine="${twip(fl)}"` : fl < -1 ? ` w:hanging="${twip(Math.min(-fl, left))}"` : ''}/>`;
      const jc = P.align !== 'left' ? `<w:jc w:val="${P.align}"/>` : '';
      return { ppr: `${tabs}<w:spacing w:before="${twip(P.before)}" w:after="0" w:line="${twip(lineH)}" w:lineRule="atLeast"/>${ind}${jc}`, body: multi ? r.parts.map((q, i) => (i ? '<w:r><w:tab/></w:r>' : '') + runXml(q, q.str)).join('') : runXml(r.parts[0], P.text) };
    }),
  };
}
function framePage(pc, p, B) { // every line pinned at its exact position (text frames) — looks identical, edits line by line
  const paras = pc.lines.map(l => ({
    ppr: `<w:framePr w:w="${twip(l.w * 1.1 + 8)}" w:h="${twip(l.size * 1.25)}" w:hRule="exact" w:wrap="none" w:vAnchor="page" w:hAnchor="page" w:x="${twip(Math.max(0, l.x))}" w:y="${twip(Math.max(0, l.base - 0.95 * l.size))}"/><w:spacing w:before="0" w:after="0" w:line="${twip(l.size * 1.2)}" w:lineRule="exact"/>`,
    body: runXml(l, l.str),
  }));
  paras.push({ ppr: '<w:spacing w:before="0" w:after="0"/>', body: pc.images.map(im => B.pic(im.url, im.w, im.h, im)).join('') });
  return { paras, margins: { t: 10, l: 10, r: 10, b: 10 } };
}
async function pdfToDocx() {
  const r = await dialog('PDF → Word (.docx)', `
    <label>Layout<select name="mode">
      <option value="flow">Editable text — flowing paragraphs (best for editing)</option>
      <option value="frame">Exact layout — text boxes at their original positions</option>
      <option value="pic">Picture — each page as an image (looks identical, not editable)</option></select></label>
    <label>Pages (blank = all), e.g. 1-3, 5<input type="text" name="range" autocomplete="off"></label>
    <p class="mut">Includes your edits. Text, fonts, sizes, colours and pictures are converted; drawn shapes, tables borders and complex layouts are approximated. Scanned pages have no text to convert — use “Picture”.</p>`, { ok: 'Convert' });
  if (!r) return; const pages = pickPages(r.range); if (!pages) return;
  await busy(async () => {
    const B = docxBuilder(); let body = '';
    for (let i = 0; i < pages.length; i++) {
      const p = pages[i]; let out, w = p.w, h = p.h;
      if (r.mode === 'pic') { [w, h] = X.rdims(p); const c = await X.renderPageImage(p, 150 / 72); out = { paras: [{ ppr: '<w:spacing w:before="0" w:after="0"/>', body: B.pic(c.toDataURL('image/jpeg', 0.85), w, h, { x: 0, y: 0 }) }], margins: { t: 10, l: 10, r: 10, b: 10 } }; }
      else { const pc = await pageContent(p); out = r.mode === 'frame' ? framePage(pc, p, B) : flowPage(pc, p, B); }
      if (!out.paras.length) out.paras.push({ ppr: '', body: '' });
      const sect = B.sect(w, h, out.margins), last = out.paras.length - 1;
      out.paras.forEach((P, k) => { body += k === last && i < pages.length - 1 ? `<w:p><w:pPr>${P.ppr}${sect}</w:pPr>${P.body}</w:p>` : `<w:p><w:pPr>${P.ppr}</w:pPr>${P.body}</w:p>`; });
      if (i === pages.length - 1) body += sect;
    }
    download(await B.finish(body), X.baseName() + '.docx'); toast('Word document saved to your downloads');
  }, 'Converting to Word…');
}

/* ───────── DOCX → pages ───────── */
const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const kids = (el, n) => el ? [...el.children].filter(c => c.localName === n) : [];
const kid = (el, n) => { if (!el) return null; for (const c of el.children) if (c.localName === n) return c; return null; };
const av = (el, n) => el ? (el.getAttributeNS(W_NS, n) || el.getAttribute('w:' + n)) : null;
const onoff = el => { const v = av(el, 'val'); return !(v === '0' || v === 'false' || v === 'off'); };
const famOf = f => Object.keys(X.BUNDLED).find(n => n.toLowerCase() === f.toLowerCase()) || (/courier|consol|mono/i.test(f) ? 'Courier' : /times|georgia|cambria|garamond|antiqua|palatino|minion|bookman|century|serif/i.test(f) && !/sans/i.test(f) ? 'Times' : 'Helvetica');
const BASE_R = { size: 10, font: 'Helvetica', color: '#000000', bold: false, italic: false, underline: false, strike: false };
function parseRPr(el) {
  const o = {}; if (!el) return o; let e;
  if ((e = kid(el, 'b'))) o.bold = onoff(e); if ((e = kid(el, 'i'))) o.italic = onoff(e); if ((e = kid(el, 'strike'))) o.strike = onoff(e); if ((e = kid(el, 'caps'))) o.caps = onoff(e);
  if ((e = kid(el, 'u'))) o.underline = (av(e, 'val') || 'single') !== 'none';
  if ((e = kid(el, 'sz')) && +av(e, 'val')) o.size = +av(e, 'val') / 2;
  if ((e = kid(el, 'color'))) { const v = av(e, 'val'); if (v && /^[0-9a-f]{6}$/i.test(v)) o.color = '#' + v.toLowerCase(); else if (v === 'auto') o.color = '#000000'; }
  if ((e = kid(el, 'rFonts'))) { const f = av(e, 'ascii') || av(e, 'hAnsi') || av(e, 'cs'); if (f) o.font = famOf(f); }
  return o;
}
function parsePPr(el) {
  const o = {}; if (!el) return o; let e;
  if ((e = kid(el, 'pStyle'))) o.style = av(e, 'val');
  if ((e = kid(el, 'jc'))) o.align = { center: 'center', right: 'right', end: 'right' }[av(e, 'val')] || 'left';
  if ((e = kid(el, 'spacing'))) {
    if (av(e, 'before') != null) o.before = +av(e, 'before') / 20; if (av(e, 'after') != null) o.after = +av(e, 'after') / 20;
    if (av(e, 'line')) { const rule = av(e, 'lineRule'); if (!rule || rule === 'auto') { o.lineMult = +av(e, 'line') / 240; o.lineExact = 0; } else { o.lineExact = +av(e, 'line') / 20; o.lineRule = rule; } }
  }
  if ((e = kid(el, 'ind'))) {
    const l = av(e, 'left') || av(e, 'start'), r = av(e, 'right') || av(e, 'end');
    if (l != null) o.indL = +l / 20; if (r != null) o.indR = +r / 20;
    if (av(e, 'hanging') != null) { o.hang = +av(e, 'hanging') / 20; o.first = 0; } else if (av(e, 'firstLine') != null) { o.first = +av(e, 'firstLine') / 20; o.hang = 0; }
  }
  if ((e = kid(el, 'tabs'))) o.tabs = kids(e, 'tab').filter(t => av(t, 'val') !== 'clear').map(t => +av(t, 'pos') / 20).sort((a, b) => a - b);
  if ((e = kid(el, 'pageBreakBefore'))) o.pbb = onoff(e);
  if ((e = kid(el, 'numPr'))) o.num = { id: av(kid(e, 'numId'), 'val'), lvl: av(kid(e, 'ilvl'), 'val') || '0' };
  return o;
}
const hasBorders = el => el ? [...el.children].some(c => { const v = av(c, 'val'); return v && v !== 'nil' && v !== 'none'; }) : null;
const fmtNum = (n, fmt) => {
  if (fmt === 'lowerLetter' || fmt === 'upperLetter') { const s = String.fromCharCode(97 + (n - 1) % 26); return fmt === 'upperLetter' ? s.toUpperCase() : s; }
  if (fmt === 'lowerRoman' || fmt === 'upperRoman') { let s = '', v = n; for (const [a, r] of [[1000, 'm'], [900, 'cm'], [500, 'd'], [400, 'cd'], [100, 'c'], [90, 'xc'], [50, 'l'], [40, 'xl'], [10, 'x'], [9, 'ix'], [5, 'v'], [4, 'iv'], [1, 'i']]) while (v >= a) { s += r; v -= a; } return fmt === 'upperRoman' ? s.toUpperCase() : s; }
  return String(n);
};

async function docxToPages(file) {
  let zip; try { zip = await JSZip.loadAsync(await file.arrayBuffer()); } catch (e) { throw new Error('Could not read this Word file — is it a valid .docx?'); }
  const xml = async n => { const f = zip.file(n); return f ? new DOMParser().parseFromString(await f.async('string'), 'application/xml') : null; };
  const doc = await xml('word/document.xml'); if (!doc || !doc.documentElement || doc.getElementsByTagName('parsererror').length) throw new Error('This is not a valid .docx file');
  const stylesDoc = await xml('word/styles.xml'), numDoc = await xml('word/numbering.xml'), relsDoc = await xml('word/_rels/document.xml.rels');

  // styles
  const raw = {}, defaults = { r: {}, p: {} }, tblBorders = {}, cache = {}; let defPara = null;
  if (stylesDoc) {
    const root = stylesDoc.documentElement, dd = kid(root, 'docDefaults');
    defaults.r = parseRPr(kid(kid(dd, 'rPrDefault'), 'rPr')); defaults.p = parsePPr(kid(kid(dd, 'pPrDefault'), 'pPr'));
    for (const st of kids(root, 'style')) {
      const id = av(st, 'styleId'), type = av(st, 'type'); raw[id] = { based: av(kid(st, 'basedOn'), 'val'), r: parseRPr(kid(st, 'rPr')), p: parsePPr(kid(st, 'pPr')) }; delete raw[id].p.style;
      if (type === 'paragraph' && av(st, 'default') === '1') defPara = id;
      if (type === 'table') tblBorders[id] = hasBorders(kid(kid(st, 'tblPr'), 'tblBorders'));
    }
  }
  const style = (id, depth = 0) => { if (!id || !raw[id] || depth > 12) return { r: {}, p: {} }; if (cache[id]) return cache[id]; const b = style(raw[id].based, depth + 1); return cache[id] = { r: { ...b.r, ...raw[id].r }, p: { ...b.p, ...raw[id].p } }; };

  // numbering (bullets / numbered lists)
  const nums = {}, abs = {}, counters = {};
  if (numDoc) {
    const root = numDoc.documentElement;
    for (const a of kids(root, 'abstractNum')) { const lv = abs[av(a, 'abstractNumId')] = {}; for (const l of kids(a, 'lvl')) { const ind = kid(kid(l, 'pPr'), 'ind'); lv[av(l, 'ilvl')] = { fmt: av(kid(l, 'numFmt'), 'val') || 'decimal', text: av(kid(l, 'lvlText'), 'val') || '', start: +(av(kid(l, 'start'), 'val') || 1), indL: ind ? +(av(ind, 'left') || av(ind, 'start') || 0) / 20 : null, hang: ind ? +(av(ind, 'hanging') || 0) / 20 : 0 }; } }
    for (const n of kids(root, 'num')) nums[av(n, 'numId')] = av(kid(n, 'abstractNumId'), 'val');
  }
  const label = num => {
    const lv = +num.lvl || 0, a = abs[nums[num.id]]; if (!a || !a[lv]) return null;
    const c = counters[num.id] || (counters[num.id] = []); c[lv] = (c[lv] === undefined ? a[lv].start - 1 : c[lv]) + 1; c.length = lv + 1;
    const d = a[lv]; return { d, text: d.fmt === 'bullet' ? '•' : d.fmt === 'none' ? '' : d.text.replace(/%(\d)/g, (m, k) => fmtNum(c[k - 1] === undefined ? 1 : c[k - 1], (a[k - 1] || d).fmt)) };
  };

  // pictures
  const images = {};
  if (relsDoc) for (const r of [...relsDoc.documentElement.children]) {
    if (!/\/image$/.test(r.getAttribute('Type') || '')) continue;
    const t = r.getAttribute('Target') || '', f = zip.file(t.startsWith('/') ? t.slice(1) : 'word/' + t), ext = (t.split('.').pop() || '').toLowerCase();
    const mime = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', webp: 'image/webp', svg: 'image/svg+xml' }[ext]; if (!f || !mime) continue;
    try { const im = await X.readImage(new File([await f.async('uint8array')], t, { type: mime })); images[r.getAttribute('Id')] = { asset: X.addAsset(im.url), pw: im.w, ph: im.h }; } catch (e) { /* unsupported picture — skipped */ }
  }
  const imageOf = el => {
    const blip = el.getElementsByTagNameNS('*', 'blip')[0]; if (!blip) return null;
    const rid = blip.getAttribute('r:embed') || blip.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'embed'), im = images[rid]; if (!im) return null;
    const ext = el.getElementsByTagNameNS('*', 'extent')[0], w = ext ? +ext.getAttribute('cx') / 12700 : im.pw * 0.75, h = ext ? +ext.getAttribute('cy') / 12700 : im.ph * 0.75;
    return { asset: im.asset, w: w || im.pw * 0.75, h: h || im.ph * 0.75 };
  };

  // document → blocks
  function parsePara(pEl) {
    const pPr = kid(pEl, 'pPr'), direct = parsePPr(pPr), stl = style(direct.style || defPara), P = { ...defaults.p, ...stl.p, ...direct }, Rb = { ...BASE_R, ...defaults.r, ...stl.r }, runs = []; let pb = false;
    const walk = el => {
      for (const c of el.children) {
        const n = c.localName;
        if (n === 'r') {
          const rp = kid(c, 'rPr'), rs = av(kid(rp, 'rStyle'), 'val'), st = { ...Rb, ...(rs ? style(rs).r : {}), ...parseRPr(rp) };
          st.key = [st.size, st.font, st.bold, st.italic, st.underline, st.strike, st.color].join('|');
          for (const t of c.children) {
            const tn = t.localName;
            if (tn === 't') runs.push({ text: st.caps ? t.textContent.toUpperCase() : t.textContent, st });
            else if (tn === 'tab') runs.push({ text: '\t', st });
            else if (tn === 'br' || tn === 'cr') { if (av(t, 'type') === 'page') pb = true; else runs.push({ text: '\n', st }); }
            else if (tn === 'noBreakHyphen') runs.push({ text: '-', st });
            else if (tn === 'drawing' || tn === 'AlternateContent') { const im = imageOf(t); if (im) runs.push({ img: im, st }); }
          }
        } else if (/^(hyperlink|ins|smartTag|sdt|sdtContent|fldSimple)$/.test(n)) walk(c);
      }
    };
    walk(pEl);
    const b = { t: 'p', P, direct, runs, mark: { ...Rb, ...parseRPr(kid(pPr, 'rPr')) }, pb, sect: !!kid(pPr, 'sectPr') };
    if (P.num && P.num.id && P.num.id !== '0') {
      const lb = label(P.num);
      if (lb) { b.label = lb.text; if (direct.indL === undefined && lb.d.indL != null) { P.indL = lb.d.indL; P.hang = lb.d.hang; P.first = 0; } else if (P.indL === undefined) { P.indL = 36 * ((+P.num.lvl || 0) + 1); P.hang = 18; P.first = 0; } }
    }
    return b;
  }
  const flat = blocks => blocks.flatMap(b => b.t === 'tbl' ? b.rows.flatMap(r => r.flatMap(c => c.blocks)) : [b]);
  function parseTbl(tEl) {
    const tblPr = kid(tEl, 'tblPr'), ts = av(kid(tblPr, 'tblStyle'), 'val'); let borders = hasBorders(kid(tblPr, 'tblBorders')); if (borders === null) borders = !!(ts && tblBorders[ts]);
    return { t: 'tbl', borders, styled: !!ts, grid: kids(kid(tEl, 'tblGrid'), 'gridCol').map(g => (+av(g, 'w') || 0) / 20),
      rows: kids(tEl, 'tr').map(tr => kids(tr, 'tc').map(tc => {
        const pr = kid(tc, 'tcPr'), sh = av(kid(pr, 'shd'), 'fill'), vm = kid(pr, 'vMerge');
        return { span: +av(kid(pr, 'gridSpan'), 'val') || 1, cont: !!vm && av(vm, 'val') !== 'restart', shade: sh && /^[0-9a-f]{6}$/i.test(sh) && sh.toLowerCase() !== 'ffffff' ? '#' + sh.toLowerCase() : null, blocks: flat(blocksOf(tc)) };
      })) };
  }
  function blocksOf(el) { const out = []; for (const c of el.children) { const n = c.localName; if (n === 'p') out.push(parsePara(c)); else if (n === 'tbl') out.push(parseTbl(c)); else if (n === 'sdt') out.push(...blocksOf(kid(c, 'sdtContent') || c)); } return out; }
  const body = kid(doc.documentElement, 'body'); if (!body) throw new Error('This Word file has no document body');
  const blocks = blocksOf(body), sp = kid(body, 'sectPr'), sz = kid(sp, 'pgSz'), mar = kid(sp, 'pgMar');
  const PW = (+av(sz, 'w') || 12240) / 20, PH = (+av(sz, 'h') || 15840) / 20, m = k => { const v = av(mar, k); return v == null ? 72 : Math.max(0, +v / 20); };
  const ML = m('left'), MR = m('right'), MT = m('top'), MB = m('bottom'), CW = Math.max(100, PW - ML - MR), maxH = PH - MT - MB;

  // layout
  function layPara(b, width) {
    const P = b.P, indL = P.indL || 0, indR = P.indR || 0, first = P.hang ? -P.hang : (P.first || 0), lines = []; let line = null, pend = null;
    const avail = () => Math.max(40, width - indL - indR - (lines.length === 1 ? first : 0));
    const nl = () => { line = { segs: [], w: 0, max: 0, imgH: 0 }; lines.push(line); pend = null; };
    const put = (text, st, w) => { const l = line.segs[line.segs.length - 1]; if (l && !l.img && l.st.key === st.key && Math.abs(l.x + l.w - line.w) < 0.01) { l.text += text; l.w += w; } else line.segs.push({ text, st, x: line.w, w }); line.w += w; if (st.size > line.max) line.max = st.size; };
    nl();
    if (b.label) { const st = { ...b.mark, underline: false }; st.key = 'label'; put(b.label, st, X.measure(b.label, st)); line.w = Math.max(line.w + 4, P.hang || 0); }
    const toks = [];
    for (const r of b.runs) {
      if (r.img) { toks.push(r); continue; }
      for (const piece of r.text.split(/( +|\t|\n)/)) { if (!piece) continue; toks.push(piece === '\t' ? { tab: 1, st: r.st } : piece === '\n' ? { br: 1, st: r.st } : { text: piece, sp: piece[0] === ' ', st: r.st, w: X.measure(piece, r.st) }); }
    }
    for (let i = 0; i < toks.length; i++) {
      const t = toks[i];
      if (t.sp) { pend = pend ? { text: pend.text + t.text, st: pend.st, w: pend.w + t.w } : t; continue; }
      if (t.br) { if (!line.max) line.max = t.st.size; nl(); continue; }
      if (t.tab) { // jump to the paragraph's next tab stop (measured from the margin), else the default half-inch grid
        if (pend) put(pend.text, pend.st, pend.w); pend = null;
        const off = indL + (lines.length === 1 ? first : 0), at = off + line.w, stop = (P.tabs || []).find(s => s > at + 0.5);
        line.w = (stop !== undefined ? stop : (Math.floor((at + 0.5) / 36) + 1) * 36) - off; if (t.st.size > line.max) line.max = t.st.size; continue;
      }
      if (t.img) { let w = t.img.w, h = t.img.h; const f = Math.min(1, avail() / w, maxH / h); w *= f; h *= f; if (line.segs.length && line.w + w > avail()) nl(); line.segs.push({ img: t.img, x: line.w, w, h }); line.w += w; line.imgH = Math.max(line.imgH, h); pend = null; continue; }
      let j = i, cw = 0; while (j < toks.length && toks[j].text !== undefined && !toks[j].sp) cw += toks[j++].w; // words glued across runs wrap together
      const sw = pend ? pend.w : 0;
      if (line.segs.length && line.w + sw + cw > avail()) nl(); else if (pend && line.segs.length) put(pend.text, pend.st, sw);
      pend = null; for (; i < j; i++) put(toks[i].text, toks[i].st, toks[i].w); i--;
    }
    for (const l of lines) {
      const mx = l.max || b.mark.size, th = P.lineExact ? (P.lineRule === 'atLeast' ? Math.max(P.lineExact, mx * 1.2) : P.lineExact) : mx * 1.2 * (P.lineMult || 1);
      l.h = l.max || !l.imgH ? Math.max(th, l.imgH ? l.imgH + 2 : 0) : l.imgH + 2; l.base = l.h - 0.27 * mx;
    }
    return { lines, before: P.before || 0, after: P.after || 0, indL, indR, first, k: P.align === 'center' ? 0.5 : P.align === 'right' ? 1 : 0, width };
  }
  function emitLine(pg, L, ln, i, x0, y) {
    const sx = x0 + L.indL + (i === 0 ? L.first : 0), off = (L.width - L.indL - L.indR - (i === 0 ? L.first : 0) - ln.w) * L.k;
    for (const s of ln.segs) {
      if (s.img) pg.objs.push({ id: uid(), type: 'image', asset: s.img.asset, x: sx + off + s.x, y: y + ln.h - s.h - 1, w: s.w, h: s.h, opacity: 1 });
      else if (s.text.trim()) pg.objs.push(X.textObj(s.text.replace(/\s+$/, ''), s.st.size, { x: sx + off + s.x, y: y + ln.base - X.BASE * s.st.size, color: s.st.color, font: s.st.font, bold: !!s.st.bold, italic: !!s.st.italic, underline: !!s.st.underline, strike: !!s.st.strike }));
    }
  }
  const pages = []; let pg = null, y = 0, brk = false;
  const newPage = () => { pg = { id: uid(), src: -1, w: PW, h: PH, rot: 0, objs: [] }; pages.push(pg); y = MT; };
  const room = h => { if (!pg || (y + h > PH - MB && y > MT + 1)) newPage(); };
  for (const b of blocks) {
    if (b.t === 'p') {
      if ((brk || b.P.pbb) && pg && y > MT + 1) newPage(); brk = false;
      const L = layPara(b, CW); if (!pg) newPage(); if (y > MT + 1) y += L.before;
      L.lines.forEach((ln, i) => { room(ln.h); emitLine(pg, L, ln, i, ML, y); y += ln.h; });
      y += L.after; if (b.pb || b.sect) brk = true;
    } else {
      const nCols = Math.max(1, ...b.rows.map(r => r.reduce((a, c) => a + c.span, 0))); let cols = b.grid.filter(w => w > 0);
      if (cols.length !== nCols) cols = Array(nCols).fill(CW / nCols);
      const tot = cols.reduce((a, c) => a + c, 0), sc = tot > CW ? CW / tot : 1, xs = [ML]; cols.forEach(c => xs.push(xs[xs.length - 1] + c * sc));
      const PX = 5, PY = 3;
      for (const row of b.rows) {
        let ci = 0;
        const cells = row.map(c => {
          const x = xs[Math.min(ci, nCols)], w = xs[Math.min(ci + c.span, nCols)] - x; ci += c.span;
          const lays = c.cont ? [] : c.blocks.map(pb => { if (b.styled && pb.direct.after === undefined && !pb.direct.style) pb.P = { ...pb.P, after: 0, lineMult: 1 }; return layPara(pb, Math.max(20, w - 2 * PX)); });
          let h = 2 * PY; lays.forEach((L, i) => { h += (i ? L.before : 0) + L.lines.reduce((a, l) => a + l.h, 0) + (i < lays.length - 1 ? L.after : 0); });
          return { c, x, w, lays, h };
        });
        const rh = Math.max(14, ...cells.map(c => c.h)); room(rh);
        for (const ce of cells) {
          if (ce.w <= 0) continue;
          if (ce.c.shade) pg.objs.push({ id: uid(), type: 'rect', x: ce.x, y, w: ce.w, h: rh, color: 'none', fill: ce.c.shade, width: 0, opacity: 1 });
          if (b.borders) pg.objs.push({ id: uid(), type: 'rect', x: ce.x, y, w: ce.w, h: rh, color: '#444444', fill: 'none', width: 0.6, opacity: 1 });
          let cy = y + PY; ce.lays.forEach((L, i) => { if (i) cy += L.before; L.lines.forEach((ln, li) => { emitLine(pg, L, ln, li, ce.x + PX, cy); cy += ln.h; }); cy += L.after; });
        }
        y += rh;
      }
      y += 8;
    }
  }
  if (!pages.length) newPage();
  return pages;
}

/* ───────── other exports ───────── */
async function pdfToImages() {
  const r = await dialog('PDF → Images', `
    <div class="row"><label>Format<select name="fmt"><option value="png">PNG (sharp, larger)</option><option value="jpg">JPG (smaller)</option></select></label>
    <label>Resolution<select name="dpi"><option value="96">96 dpi (screen)</option><option value="150" selected>150 dpi</option><option value="200">200 dpi</option><option value="300">300 dpi (print)</option></select></label></div>
    <label>Pages (blank = all), e.g. 1-3, 5<input type="text" name="range" autocomplete="off"></label>
    <p class="mut">Several pages are saved together as one ZIP file.</p>`, { ok: 'Export' });
  if (!r) return; const pages = pickPages(r.range); if (!pages) return;
  await busy(async () => {
    const zip = new JSZip(), mime = r.fmt === 'jpg' ? 'image/jpeg' : 'image/png', len = String(S.pages.length).length;
    for (const p of pages) {
      const blob = await toBlob(await X.renderPageImage(p, +r.dpi / 72), mime, 0.88), name = `${X.baseName()}-page-${pad(S.pages.indexOf(p) + 1, len)}.${r.fmt}`;
      if (pages.length === 1) return download(blob, name); zip.file(name, blob);
    }
    await zipSave(zip, X.baseName() + '-images.zip');
  }, 'Rendering pages…');
}
async function splitZip() {
  if (S.pages.length < 2) return toast('This document has only one page');
  await busy(async () => {
    const zip = new JSZip(), len = String(S.pages.length).length;
    for (let i = 0; i < S.pages.length; i++) zip.file(`${X.baseName()}-page-${pad(i + 1, len)}.pdf`, await X.buildPdf([S.pages[i]]));
    await zipSave(zip, X.baseName() + '-split.zip'); toast(`${S.pages.length} single-page PDFs saved as a ZIP`);
  }, 'Splitting…');
}
async function exportText() {
  await busy(async () => {
    const parts = [];
    for (let i = 0; i < S.pages.length; i++) parts.push(rowsOf(await X.visibleText(S.pages[i])).map(r => r.parts.map(q => q.str).join('\t')).join('\r\n'));
    if (!parts.some(t => t.trim())) return toast('No text found — scanned pages contain pictures of text, not text', true);
    download(new Blob(['﻿' + parts.map((t, i) => S.pages.length > 1 ? `—— Page ${i + 1} ——\r\n${t}` : t).join('\r\n\r\n')], { type: 'text/plain;charset=utf-8' }), X.baseName() + '.txt');
  }, 'Extracting text…');
}
async function flatten() {
  const r = await dialog('Flatten (permanent redaction)', `
    <p>Turns every page into a picture and saves a copy. It <b>permanently bakes in</b> whiteout, blackout and edited text — the hidden originals are removed.</p>
    <p class="mut">Text in the copy can no longer be selected or searched. Your open document is not changed.</p>
    <div class="row"><label>Quality<select name="q"><option value="110|0.6">Small file (screen)</option><option value="150|0.72" selected>Balanced</option><option value="200|0.85">High (print)</option></select></label>
    <label>Colour<select name="gray"><option value="">Keep colours</option><option value="1">Grayscale</option></select></label></div>`, { ok: 'Save copy' });
  if (!r) return; const [dpi, q] = r.q.split('|').map(Number);
  await busy(async () => {
    const out = await PDFLib.PDFDocument.create();
    for (const p of S.pages) {
      const c = await X.renderPageImage(p, dpi / 72); if (r.gray) X.applyFx(c, { mode: 'gray' });
      const [w, h] = X.rdims(p); out.addPage([w, h]).drawImage(await out.embedJpg(c.toDataURL('image/jpeg', q)), { x: 0, y: 0, width: w, height: h });
    }
    out.setProducer('XD3 PDF Editor (XD3Labs / XDCybertech Pvt Ltd)');
    const bytes = await out.save(); download(new Blob([bytes], { type: 'application/pdf' }), X.baseName() + '-flattened.pdf');
    toast(`Saved — ${(bytes.length / 1048576).toFixed(2)} MB`);
  }, 'Flattening pages…');
}

/* ───────── compress: re-encode the pictures inside the PDF — text and vector graphics are left untouched ───────── */
function unpredict(d, w, comps) { // undo PNG row filters (DecodeParms /Predictor 10–15)
  const row = w * comps, rows = Math.floor(d.length / (row + 1)), out = new Uint8Array(rows * row);
  for (let y = 0; y < rows; y++) {
    const ft = d[y * (row + 1)], src = y * (row + 1) + 1, dst = y * row;
    for (let x = 0; x < row; x++) {
      const a = x >= comps ? out[dst + x - comps] : 0, b = y ? out[dst - row + x] : 0, c = y && x >= comps ? out[dst - row + x - comps] : 0; let v = d[src + x];
      if (ft === 1) v += a; else if (ft === 2) v += b; else if (ft === 3) v += (a + b) >> 1;
      else if (ft === 4) { const p = a + b - c, pa = Math.abs(p - a), pb = Math.abs(p - b), pc = Math.abs(p - c); v += pa <= pb && pa <= pc ? a : pb <= pc ? b : c; }
      out[dst + x] = v;
    }
  }
  return out;
}
async function shrinkImage(doc, ref, obj, opt) { // → true when the picture was replaced by a smaller one
  const L = PDFLib, N = L.PDFName.of.bind(L.PDFName), dict = obj.dict, num = k => { const v = dict.lookup(N(k)); return v instanceof L.PDFNumber ? v.asNumber() : 0; };
  if (dict.lookup(N('Subtype')) !== N('Image') || dict.has(N('ImageMask')) || dict.has(N('Decode')) || dict.has(N('Mask'))) return false;
  const w = num('Width'), h = num('Height'); if (num('BitsPerComponent') !== 8 || w * h < 96 * 96 || obj.contents.length < 12000) return false;
  let filter = dict.lookup(N('Filter')); if (filter instanceof L.PDFArray) { if (filter.size() !== 1) return false; filter = filter.lookup(0); }
  const cs = dict.lookup(N('ColorSpace')); let comps = cs === N('DeviceRGB') ? 3 : cs === N('DeviceGray') ? 1 : 0;
  if (!comps && cs instanceof L.PDFArray && cs.lookup(0) === N('ICCBased')) { const icc = cs.lookup(1), n = icc && icc.dict && icc.dict.lookup(N('N')); comps = n instanceof L.PDFNumber ? n.asNumber() : 0; }
  if (comps !== 1 && comps !== 3) return false; // CMYK, indexed and other special colour spaces are left alone
  let src;
  if (filter === N('DCTDecode')) src = await createImageBitmap(new Blob([obj.contents], { type: 'image/jpeg' }));
  else if (filter === N('FlateDecode')) {
    const parms = dict.lookup(N('DecodeParms')), pv = parms instanceof L.PDFDict ? parms.lookup(N('Predictor')) : null, pred = pv instanceof L.PDFNumber ? pv.asNumber() : 1;
    if (pred !== 1 && pred < 10) return false;
    let raw = L.decodePDFRawStream(obj).decode(); if (pred >= 10) raw = unpredict(raw, w, comps);
    if (raw.length < w * h * comps) return false;
    const id = new ImageData(w, h), d = id.data;
    for (let i = 0, j = 0, k = 0; i < w * h; i++, j += comps, k += 4) { d[k] = raw[j]; d[k + 1] = comps === 3 ? raw[j + 1] : raw[j]; d[k + 2] = comps === 3 ? raw[j + 2] : raw[j]; d[k + 3] = 255; }
    src = document.createElement('canvas'); src.width = w; src.height = h; src.getContext('2d').putImageData(id, 0, 0);
  } else return false;
  const f = Math.min(1, opt.max / Math.max(w, h)), c = document.createElement('canvas'); c.width = Math.max(1, Math.round(w * f)); c.height = Math.max(1, Math.round(h * f));
  const x = c.getContext('2d', { willReadFrequently: true }); x.fillStyle = '#fff'; x.fillRect(0, 0, c.width, c.height); x.imageSmoothingQuality = 'high'; x.drawImage(src, 0, 0, c.width, c.height);
  if (opt.gray) X.applyFx(c, { mode: 'gray' });
  const bytes = new Uint8Array(await (await toBlob(c, 'image/jpeg', opt.q)).arrayBuffer());
  if (bytes.length > obj.contents.length * 0.9) return false; // not worth it — keep the original
  dict.set(N('Width'), L.PDFNumber.of(c.width)); dict.set(N('Height'), L.PDFNumber.of(c.height)); dict.set(N('ColorSpace'), N('DeviceRGB')); dict.set(N('BitsPerComponent'), L.PDFNumber.of(8));
  dict.set(N('Filter'), N('DCTDecode')); dict.delete(N('DecodeParms'));
  doc.context.assign(ref, L.PDFRawStream.of(dict, bytes)); return true;
}
async function compress() {
  const r = await dialog('Compress PDF', `
    <p>Makes the file smaller by re-encoding the pictures inside it. Text stays sharp, selectable and searchable.</p>
    <label>Compression<select name="lvl"><option value="2400|0.82">Light — best picture quality</option><option value="1600|0.68" selected>Balanced</option><option value="1100|0.5">Strong — smallest file</option></select></label>
    <label class="inl"><input type="checkbox" name="gray">Turn pictures grayscale (smaller still)</label>
    <p class="mut">Saves a compressed copy; the document you have open is not changed. Files that are mostly text are already small.</p>`, { ok: 'Compress' });
  if (!r) return; const [max, q] = r.lvl.split('|').map(Number);
  await busy(async () => {
    const before = await X.buildPdf(S.pages), doc = await PDFLib.PDFDocument.load(before, { updateMetadata: false }); let n = 0;
    for (const [ref, obj] of doc.context.enumerateIndirectObjects()) {
      if (!(obj instanceof PDFLib.PDFRawStream)) continue;
      try { if (await shrinkImage(doc, ref, obj, { max, q, gray: !!r.gray })) n++; } catch (e) { /* this picture stays as it is */ }
    }
    const after = n ? await doc.save() : before, best = after.length < before.length ? after : before, mb = v => (v / 1048576).toFixed(2) + ' MB';
    download(new Blob([best], { type: 'application/pdf' }), X.baseName() + '-compressed.pdf');
    toast(best === after && n ? `Compressed: ${mb(before.length)} → ${mb(after.length)} (${Math.round((1 - after.length / before.length) * 100)}% smaller, ${n} picture(s) re-encoded)` : `Already compact (${mb(before.length)}) — no pictures left to shrink`);
  }, 'Compressing…');
}

Object.assign(X.ACT, { compress, pdf2docx: pdfToDocx, pdf2img: pdfToImages, splitzip: splitZip, exporttxt: exportText, flatten });
X.NO_DOC.add('imgpages');
Object.assign(X, { docxToPages, pageContent, rowsOf, shrinkImage });
})();
