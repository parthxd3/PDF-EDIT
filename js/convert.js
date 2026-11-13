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
    for (const l of orig) { const fs = await X.fontStyle(p, l.fontName); l.bold = fs.bold; l.italic = fs.italic; if (fs.family || fs.real) l.font = fs.family || fs.real; l.color = inkColor(ctx, K, l); }
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
const FAM = { Helvetica: 'Arial', Times: 'Times New Roman', Courier: 'Courier New', Carlito: 'Calibri', Caladea: 'Cambria', Gelasio: 'Georgia', 'DejaVu Sans': 'Verdana' };
function runXml(l, text) {
  const f = FAM[l.font] || X.esc(X.fontLabel(l.font || '').replace(/ (Bold|Italic)\b/g, '')) || 'Arial', col = l.color && l.color !== '#000000' ? `<w:color w:val="${l.color.slice(1).toUpperCase()}"/>` : '', sz = Math.max(2, Math.round(l.size * 2));
  const rpr = `<w:rPr><w:rFonts w:ascii="${f}" w:hAnsi="${f}" w:cs="${f}"/>${l.bold ? '<w:b/>' : ''}${l.italic ? '<w:i/>' : ''}${l.strike ? '<w:strike/>' : ''}${col}<w:sz w:val="${sz}"/><w:szCs w:val="${sz}"/>${l.underline ? '<w:u w:val="single"/>' : ''}</w:rPr>`;
  return `<w:r>${rpr}<w:t xml:space="preserve">${xesc(text)}</w:t></w:r>`;
}
const fnv = u8 => { let h = 0x811c9dc5; for (let i = 0; i < u8.length; i++) { h ^= u8[i]; h = Math.imul(h, 0x01000193); } return (h >>> 0).toString(36) + '.' + u8.length; };
function docxBuilder() {
  const zip = new JSZip(), rels = [], seen = new Map(), put = (name, data, o) => zip.file(name, data, { createFolders: false, ...o }); let n = 0, dp = 0; // (no folder entries in the package)
  const media = src => { // a data URL, or { bytes, ext, key } → relationship id (each distinct picture is stored once)
    const url = typeof src === 'string', key = url ? src : src.key; if (key && seen.has(key)) return seen.get(key);
    const ext = url ? (/^data:image\/jpe?g/.test(src) ? 'jpeg' : 'png') : src.ext, id = 'rId' + (100 + ++n), name = `image${n}.${ext}`;
    put('word/media/' + name, url ? src.split(',')[1] : src.bytes, url ? { base64: true } : {}); rels.push(`<Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/${name}"/>`);
    if (key) seen.set(key, id); return id;
  };
  const pic = (src, w, h, at) => { // inline picture, or floating behind the text at a page position when `at` is given
    const rid = media(src), id = ++dp, ext = `cx="${emu(w)}" cy="${emu(h)}"`;
    const g = `<wp:docPr id="${id}" name="${at ? 'Page design' : 'Picture'} ${id}"/><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="${id}" name="Picture ${id}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext ${ext}/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic>`;
    if (!at) return `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent ${ext}/><wp:effectExtent l="0" t="0" r="0" b="0"/>${g}</wp:inline></w:drawing></w:r>`;
    return `<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="${id}" behindDoc="1" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="page"><wp:posOffset>${emu(at.x)}</wp:posOffset></wp:positionH><wp:positionV relativeFrom="page"><wp:posOffset>${emu(at.y)}</wp:posOffset></wp:positionV><wp:extent ${ext}/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>${g}</wp:anchor></w:drawing></w:r>`;
  };
  const tbox = (x, y, w, h, rot, inner) => { // an invisible text box pinned to the page (x, y: its top-left before turning; it turns around its centre)
    const id = ++dp, ext = `cx="${emu(w)}" cy="${emu(h)}"`, r = Math.round((((rot % 360) + 360) % 360) * 60000);
    return `<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="${5000 + id}" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="page"><wp:posOffset>${emu(x)}</wp:posOffset></wp:positionH><wp:positionV relativeFrom="page"><wp:posOffset>${emu(y)}</wp:posOffset></wp:positionV><wp:extent ${ext}/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/><wp:docPr id="${id}" name="Text ${id}"/><wp:cNvGraphicFramePr/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm${r ? ` rot="${r}"` : ''}><a:off x="0" y="0"/><a:ext ${ext}/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></wps:spPr><wps:txbx><w:txbxContent>${inner}</w:txbxContent></wps:txbx><wps:bodyPr rot="0" vert="horz" wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" anchor="t" anchorCtr="0"><a:spAutoFit/></wps:bodyPr></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>`;
  };
  const sect = (w, h, m) => `<w:sectPr><w:pgSz w:w="${twip(w)}" w:h="${twip(h)}"${w > h ? ' w:orient="landscape"' : ''}/><w:pgMar w:top="${twip(m.t)}" w:right="${twip(m.r)}" w:bottom="${twip(m.b)}" w:left="${twip(m.l)}" w:header="0" w:footer="0" w:gutter="0"/></w:sectPr>`;
  const finish = async (body, title, extra = {}) => { // extra: { styles, numbering, header, footer } — whole XML parts written by an editable conversion
    const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" mc:Ignorable="w14"';
    const H = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>', W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"', CT = 'application/vnd.openxmlformats-officedocument', now = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
    put('[Content_Types].xml', `${H}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Default Extension="jpeg" ContentType="image/jpeg"/><Override PartName="/word/document.xml" ContentType="${CT}.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="${CT}.wordprocessingml.styles+xml"/><Override PartName="/word/settings.xml" ContentType="${CT}.wordprocessingml.settings+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>${extra.numbering ? `<Override PartName="/word/numbering.xml" ContentType="${CT}.wordprocessingml.numbering+xml"/>` : ''}${extra.header ? `<Override PartName="/word/header1.xml" ContentType="${CT}.wordprocessingml.header+xml"/>` : ''}${extra.footer ? `<Override PartName="/word/footer1.xml" ContentType="${CT}.wordprocessingml.footer+xml"/>` : ''}</Types>`);
    put('_rels/.rels', `${H}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>`);
    put('docProps/core.xml', `${H}<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>${xesc(title || '')}</dc:title><dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified></cp:coreProperties>`);
    put('word/_rels/document.xml.rels', `${H}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>${extra.numbering ? '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>' : ''}${extra.header ? '<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>' : ''}${extra.footer ? '<Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>' : ''}${rels.join('')}</Relationships>`);
    if (extra.numbering) put('word/numbering.xml', H + extra.numbering);
    if (extra.header) put('word/header1.xml', `${H}<w:hdr ${NS}>${extra.header}</w:hdr>`);
    if (extra.footer) put('word/footer1.xml', `${H}<w:ftr ${NS}>${extra.footer}</w:ftr>`);
    if (extra.styles) put('word/styles.xml', H + extra.styles); else put('word/styles.xml', `${H}<w:styles ${W}><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="en-US"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style></w:styles>`);
    put('word/settings.xml', `${H}<w:settings ${W}><w:zoom w:percent="100"/><w:defaultTabStop w:val="720"/><w:characterSpacingControl w:val="doNotCompress"/><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>`);
    put('word/document.xml', `${H}<w:document ${NS}><w:body>${body}</w:body></w:document>`);
    return zip.generateAsync({ type: 'blob', compression: 'DEFLATE', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
  };
  return { pic, tbox, sect, finish };
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

/* "Same as the PDF": the page's own drawing (rules, table lines, shading, pictures, form fields) becomes one picture behind the text,
   and every piece of text goes into an invisible text box at its exact place — in the font Word has that matches, at the exact width.
   Text the PDF itself does not show (the hidden layer of a scanned page, text drawn as shapes) is kept as see-through, searchable text. */
const DK = 2.5; // the page design is rendered at 180 dpi
const probe = document.createElement('canvas').getContext('2d'); try { probe.fontKerning = 'none'; probe.textRendering = 'optimizeSpeed'; } catch (e) { /* older Chromium: kerning stays on */ }
const hasFont = (() => { // is this font installed? (it measures differently from every generic family)
  const memo = {}, t = 'mmmmmmmmmmlliWWQ@#0123456789ag';
  return name => {
    if (name in memo) return memo[name]; let ok = false; const q = name.replace(/["\\]/g, '');
    for (const g of ['monospace', 'serif', 'sans-serif']) { probe.font = `40px ${g}`; const a = probe.measureText(t).width; probe.font = `40px "${q}", ${g}`; if (Math.abs(probe.measureText(t).width - a) > 0.2) { ok = true; break; } }
    return memo[name] = ok;
  };
})();
const STD_FONT = { helvetica: 'Arial', arial: 'Arial', arimo: 'Arial', liberationsans: 'Arial', nimbussans: 'Arial', nimbussansl: 'Arial', times: 'Times New Roman', timesroman: 'Times New Roman', timesnewroman: 'Times New Roman', tinos: 'Times New Roman', liberationserif: 'Times New Roman', nimbusroman: 'Times New Roman', courier: 'Courier New', couriernew: 'Courier New', cousine: 'Courier New', liberationmono: 'Courier New', nimbusmono: 'Courier New', carlito: 'Calibri', caladea: 'Cambria', gelasio: 'Georgia', dejavusans: 'Verdana', symbol: 'Symbol', zapfdingbats: 'Wingdings' };
const WEIGHT_RE = /(Semi[Bb]old|Demi[Bb]old|Extra[Bb]old|Ultra[Bb]old|Semi[Ll]ight|Extra[Ll]ight|Ultra[Ll]ight|Light|Thin|Medium|Black|Heavy|Condensed|Narrow)/;
const fontMemo = new Map();
function wordFont(ps, cls) { // a PDF font name → { name: the font Word should use, exact: Word has that very font, weight: the family name already carries the weight }
  const key = ps + '|' + cls; if (fontMemo.has(key)) return fontMemo.get(key);
  const done = v => (fontMemo.set(key, v), v), sp = s => s.replace(/([a-z])([A-Z])/g, '$1 $2').replace(/([A-Z]+)([A-Z][a-z])/g, '$1 $2').trim();
  const raw = String(ps || '').replace(/^[A-Z]{6}\+/, ''), parts = raw.split(/[-,]/), base = (parts[0] || '').replace(/(PS)?MT$|PS$/, ''), wt = ((parts[1] || '').match(WEIGHT_RE) || [])[1] || '';
  const bare = base.replace(/(Bold|Italic|Oblique|Regular|Book|Normal)+$/, '') || base, cand = [];
  if (wt) cand.push([sp(bare) + ' ' + wt.replace(/^(Semi|Demi|Extra|Ultra)([BL])/, (m, a, b) => a + b.toLowerCase()), 1], [sp(bare) + ' ' + wt, 1]);
  cand.push([sp(bare), 0], [sp(base), WEIGHT_RE.test(base) ? 1 : 0], [bare, 0]);
  for (const [n, w] of cand) { if (!n) continue; const std = STD_FONT[n.toLowerCase().replace(/[^a-z]/g, '')], nm = std || n; if (hasFont(nm)) return done({ name: nm, exact: true, weight: !!w && !std }); }
  const alt = STD_FONT[String(cls || '').toLowerCase().replace(/[^a-z]/g, '')];
  if (alt && hasFont(alt)) return done({ name: alt, exact: false, weight: false });
  return done({ name: /courier|mono/i.test(cls) ? 'Courier New' : /times|serif/i.test(cls) && !/sans/i.test(cls) ? 'Times New Roman' : 'Arial', exact: false, weight: false });
}
const natWidth = (text, st) => { probe.font = `${st.italic ? 'italic ' : ''}${st.bold ? 'bold ' : ''}100px "${st.font.replace(/["\\]/g, '')}"`; return probe.measureText(text).width * st.sz / 200; }; // st.sz: half-points, as Word stores sizes
const GHOST = '<w14:textFill><w14:solidFill><w14:srgbClr w14:val="000000"><w14:alpha w14:val="100000"/></w14:srgbClr></w14:solidFill></w14:textFill>';
function rprXml(st, w100, sp) {
  const f = xesc(st.font), col = st.color && /^#[0-9a-f]{6}$/i.test(st.color) && st.color.toLowerCase() !== '#000000' && !st.ghost ? `<w:color w:val="${st.color.slice(1).toUpperCase()}"/>` : '';
  return `<w:rPr><w:rFonts w:ascii="${f}" w:hAnsi="${f}" w:eastAsia="${f}" w:cs="${f}"/>${st.bold ? '<w:b/><w:bCs/>' : ''}${st.italic ? '<w:i/><w:iCs/>' : ''}${st.strike ? '<w:strike/>' : ''}${st.ghost ? '<w:noProof/>' : ''}${col}${sp ? `<w:spacing w:val="${sp}"/>` : ''}${w100 && w100 !== 100 ? `<w:w w:val="${w100}"/>` : ''}${st.pos ? `<w:position w:val="${st.pos}"/>` : ''}<w:sz w:val="${st.sz}"/><w:szCs w:val="${st.sz}"/>${st.underline ? '<w:u w:val="single"/>' : ''}${st.ghost ? GHOST : ''}</w:rPr>`;
}
const runOf = (text, st, w100, sp) => `<w:r>${rprXml(st, w100, sp)}<w:t xml:space="preserve">${xesc(text)}</w:t></w:r>`;
function fitRuns(text, target, st) { // text set so that it takes up exactly `target` points: extra room of a justified line goes into its word gaps, small differences are spread over the letters
  const chars = [...text].length; if (!chars) return { xml: '', w: 0 };
  const nat = natWidth(text, st); let w100 = 100;
  if (!st.exact && nat > 0.5) w100 = clamp(Math.round(target / nat * 100), 40, 250); // a stand-in font: scaled to the original's width
  const nw = nat * w100 / 100, extra = target - nw, nSp = (text.match(/ /g) || []).length, size = st.sz / 2;
  if (nSp && extra > 0.3 + 0.006 * nw && extra / nSp < 3 * size) {
    const s = Math.round(extra / nSp * 20); let xml = '';
    for (const part of text.split(/( +)/)) if (part) xml += runOf(part, st, w100, part[0] === ' ' ? s : 0);
    return { xml, w: nw + nSp * s / 20 };
  }
  const s = clamp(Math.round(extra / chars * 20), Math.round(-6 * size), Math.round(20 * size)); // twentieths of a point after every character
  return { xml: runOf(text, st, w100, s), w: nw + chars * s / 20 };
}
function linesOf(ps) { // pieces sharing a baseline and standing close together → one line, its pieces left to right
  const rows = [], out = [];
  for (const q of [...ps].sort((a, b) => a.base - b.base || a.x - b.x)) {
    let row = null; for (let i = rows.length - 1; i >= 0 && i >= rows.length - 30; i--) { const r = rows[i]; if (Math.abs(r.base - q.base) < 0.3 * Math.min(r.size, q.size)) { row = r; break; } }
    if (row) row.ps.push(q); else rows.push({ base: q.base, size: q.size, ps: [q] });
  }
  for (const r of rows) {
    r.ps.sort((a, b) => a.x - b.x); let cur = null, end = 0;
    for (const q of r.ps) {
      const gap = q.x - end;
      if (cur && gap < 1.2 * Math.max(q.size, cur.size) && gap > -0.3 * q.size && !q.ghost === !cur.ghost) { cur.runs.push(q); cur.size = Math.max(cur.size, q.size); end = Math.max(end, q.x + q.w); }
      else { out.push(cur = { runs: [q], size: q.size, ghost: !!q.ghost }); end = q.x + q.w; }
    }
  }
  for (const l of out) { const big = l.runs.reduce((a, b) => (b.size > a.size ? b : a)); Object.assign(l, { base: big.base, x: l.runs[0].x, right: Math.max(...l.runs.map(q => q.x + q.w)) }); }
  return out;
}
function blocksOf(lines) { // lines below one another at an even spacing → one text box (paragraphs stay editable together)
  const blocks = [];
  for (const l of [...lines].sort((a, b) => a.base - b.base || a.x - b.x)) {
    let hit = null;
    for (let i = blocks.length - 1; i >= 0 && i >= blocks.length - 16; i--) {
      const b = blocks[i], last = b.lines[b.lines.length - 1], gap = l.base - last.base, sz = Math.max(l.size, last.size);
      if (b.ghost !== l.ghost || Math.abs(l.size - b.size) > 0.15 * b.size || gap < 0.95 * sz || gap > 2.2 * sz) continue;
      if (b.gap && Math.abs(gap - b.gap) > Math.max(0.35, 0.05 * b.gap)) continue;
      const ov = Math.min(l.right, last.right) - Math.max(l.x, last.x); if (ov < 0.25 * Math.min(l.right - l.x, last.right - last.x)) continue;
      hit = b; break;
    }
    if (hit) { if (!hit.gap) hit.gap = l.base - hit.lines[hit.lines.length - 1].base; hit.lines.push(l); } else blocks.push({ lines: [l], size: l.size, ghost: l.ghost, gap: 0 });
  }
  return blocks;
}
function alignOf(L) {
  if (L.length < 2) return 'left';
  const same = (f, tol) => L.every(l => Math.abs(f(l) - f(L[0])) < tol);
  return same(l => l.x, 1) ? 'left' : same(l => (l.x + l.right) / 2, 1.5) ? 'center' : same(l => l.right, 1.5) ? 'right' : 'left';
}
function styleOf(q, l, sc) {
  const pos = Math.round((l.base - q.base) * sc * 2);
  return { font: q.font, exact: q.exact, bold: q.bold, italic: q.italic, underline: q.underline, strike: q.strike, color: q.color, ghost: q.ghost, sz: clamp(Math.round(q.size * sc * 2), 2, 3276), pos: Math.abs(pos) >= 1 ? pos : 0 };
}
function lineXml(l, sc) { // every piece starts exactly where it starts in the PDF, the last one ends where it ends
  let out = '', carry = 0;
  l.runs.forEach((q, i) => {
    const nx = l.runs[i + 1], st = styleOf(q, l, sc); let text = q.str.replace(/[\r\n\t]/g, ' '), target = (nx ? nx.x - q.x : q.w) * sc + carry;
    if (nx && nx.x - (q.x + q.w) > 0.1 * q.size && !/\s$/.test(text) && !/^\s/.test(nx.str)) text += ' ';
    if (!nx) { const t = text.replace(/\s+$/, ''); if (t !== text) { target -= natWidth(text.slice(t.length), st); text = t; } }
    const r = fitRuns(text, target, st); out += r.xml; carry = target - r.w;
  });
  return out;
}
function blockXml(b, B, place, prot, sc) {
  const lines = b.lines, L = b.gap || 1.2 * b.size, al = alignOf(lines), x0 = Math.min(...lines.map(l => l.x)), r0 = Math.max(...lines.map(l => l.right));
  const span = r0 - x0, slack = 0.03 * span + 0.6 * b.size + 2; let bx = x0, bw = span + slack;
  if (al === 'center') { const c = lines.reduce((a, l) => a + (l.x + l.right) / 2, 0) / lines.length, half = Math.max(...lines.map(l => Math.max(c - l.x, l.right - c))) + slack / 2; bx = c - half; bw = 2 * half; }
  else if (al === 'right') { bx = x0 - slack; bw = r0 - bx; }
  // Word puts the baseline 0.8 of an exact line height below the line's top. Lines at the paragraph's spacing follow with a line break;
  // a line whose spacing differs starts a new paragraph whose line height lands its baseline exactly (0.2 of the last height + 0.8 of the new one).
  const exact = v => Math.max(1, twip(v * sc)) / 20 / sc, top = lines[0].base - 0.8 * exact(L); let xml = '', para = null, bh = 0, wb = 0; // wb: where Word puts the previous baseline
  const close = () => { if (para) xml += `<w:p><w:pPr>${para.ppr}</w:pPr>${para.body}</w:p>`; };
  lines.forEach((l, i) => {
    if (para && Math.abs(l.base - (wb + para.L)) < 0.12 && (al !== 'left' || Math.abs(l.x - lines[i - 1].x) < 0.5)) { para.body += '<w:r><w:br/></w:r>' + lineXml(l, sc); bh += para.L; wb += para.L; return; }
    const Lp = exact(i ? Math.max(0.9 * l.size, (l.base - wb - 0.2 * para.L) / 0.8) : L); wb = i ? wb + 0.2 * para.L + 0.8 * Lp : top + 0.8 * Lp; close(); bh += Lp;
    const ind = l.x - bx, st = styleOf(l.runs[0], l, sc);
    para = { L: Lp, ppr: `<w:spacing w:before="0" w:after="0" w:line="${twip(Lp * sc)}" w:lineRule="exact"/>${al === 'left' && ind > 0.05 ? `<w:ind w:left="${twip(ind * sc)}"/>` : ''}${al !== 'left' ? `<w:jc w:val="${al}"/>` : ''}${rprXml({ ...st, pos: 0 }, 100, 0)}`, body: lineXml(l, sc) };
  });
  close();
  const q = lines[0].runs[0], [cx, cy] = place(q, bx + bw / 2, top + bh / 2);
  return B.tbox((cx - bw / 2) * sc, (cy - bh / 2) * sc, bw * sc, bh * sc, (q.rot || 0) + prot, xml);
}
async function pageModel(p, K) { // what is on a page: its drawing without text (bg), the full render (ref), and every piece of text with font, colour and the colour behind it
  const [PW, PH] = X.rdims(p), rot = (((p.rot || 0) % 360) + 360) % 360;
  const bg = await X.renderPageImage(p, K, { noText: true, annots: false }), ref = await X.renderPageImage(p, K, { annots: false });
  const W = bg.width, H = bg.height, bx = bg.getContext('2d', { willReadFrequently: true }), bd = bx.getImageData(0, 0, W, H), d = bd.data, rd = ref.getContext('2d', { willReadFrequently: true }).getImageData(0, 0, W, H).data;
  if (p.src >= 0) { // form fields, stamps and other annotations keep their look: they go into the page design
    let an = []; try { an = await (await X.getPdfPage(p)).getAnnotations({ intent: 'print' }); } catch (e) { /* none readable */ }
    if (an.length) {
      const fd = (await X.renderPageImage(p, K)).getContext('2d', { willReadFrequently: true }).getImageData(0, 0, W, H).data; let ch = 0;
      for (let i = 0; i < d.length; i += 4) if (fd[i] !== rd[i] || fd[i + 1] !== rd[i + 1] || fd[i + 2] !== rd[i + 2]) { d[i] = fd[i]; d[i + 1] = fd[i + 1]; d[i + 2] = fd[i + 2]; ch++; }
      if (ch) bx.putImageData(bd, 0, 0);
    }
  }
  const toPage = (x, y) => (rot === 90 ? [p.h - y, x] : rot === 180 ? [p.w - x, p.h - y] : rot === 270 ? [y, p.w - x] : [x, y]);
  const place = (q, x, y) => { // a point in a piece's own (upright) frame → the page as Word shows it
    if (q.rot) { const a = q.rot * Math.PI / 180, dx = x - q.pv[0], dy = y - q.pv[1]; x = q.pv[0] + dx * Math.cos(a) - dy * Math.sin(a); y = q.pv[1] + dx * Math.sin(a) + dy * Math.cos(a); }
    return toPage(x, y);
  };
  const pieces = [], styles = {};
  if (p.src >= 0) { // the PDF's own text pieces (not the merged editing lines), so every word gap — tabs, justified lines — keeps its exact place
    const pg = await X.getPdfPage(p), vt = pg.getViewport({ scale: 1 }).transform, tc = await pg.getTextContent(), segs = [];
    for (const it of tc.items) {
      if (!it.str || !it.str.trim() || !(it.width > 0) || (tc.styles[it.fontName] || {}).vertical) continue;
      const m = pdfjsLib.Util.transform(vt, it.transform), size = Math.hypot(m[2], m[3]), ang = Math.atan2(m[1], m[0]); if (size < 1) continue;
      const rot = Math.abs(ang) > 0.02 ? Math.round(ang * 1800 / Math.PI) / 10 : 0, c = [m[4] + Math.cos(ang) * it.width / 2 + Math.sin(ang) * 0.3 * size, m[5] + Math.sin(ang) * it.width / 2 - Math.cos(ang) * 0.3 * size];
      if (X.covered(p, { x: c[0] - 0.5, y: c[1] - 0.5, w: 1, h: 1 })) continue; // replaced by an edit (its new text comes with the editor's objects below)
      segs.push({ str: it.str, x: m[4], base: m[5], w: it.width, size, rot, fontName: it.fontName });
    }
    for (const l of X.ocrCached ? X.ocrCached(p) : []) if (l.str && l.str.trim() && !X.covered(p, l)) segs.push({ str: l.str, x: l.x, base: l.base, w: l.w, size: l.size, rot: 0 });
    for (const s of segs) {
      const f = s.fontName ? await (styles[s.fontName] || (styles[s.fontName] = X.origFont(p, { fontName: s.fontName }, true).catch(() => ({ font: 'Helvetica' })))) : { font: 'Helvetica' };
      const wf = wordFont(f.ps || f.real || '', f.font);
      pieces.push({ str: s.str, x: s.x, base: s.base, w: s.w, size: s.size, rot: s.rot, pv: s.rot ? [s.x, s.base] : null, font: wf.name, exact: wf.exact, bold: !!f.bold && !wf.weight, italic: !!f.italic, sample: true });
    }
  }
  for (const o of p.objs) {
    if (o.type !== 'text' || !o.text || !o.text.trim()) continue;
    const xf = S.xfonts[o.font], BF = X.BUNDLED[o.font], cls = xf ? xf.fb : BF ? ({ serif: 'Times', monospace: 'Courier' }[BF[2]] || 'Helvetica') : o.font;
    const wf = wordFont(xf ? xf.ps || xf.name : o.font, cls), pv = o.rot ? X.center(X.bbox(o)) : null;
    for (const l of X.textLines(o)) if (l.ln.trim()) pieces.push({ str: l.ln, x: l.x, base: l.y, w: l.w, size: o.size, rot: o.rot || 0, pv, font: wf.name, exact: wf.exact, bold: !!o.bold && !wf.weight, italic: !!o.italic, underline: !!o.underline, strike: !!o.strike, color: o.color, sample: false });
  }
  const hex3 = c => '#' + c.map(v => Math.round(v).toString(16).padStart(2, '0')).join('');
  for (const q of pieces) { // colour of the text (difference between the page with and without it), and the colour behind it
    const pts = [[q.x, q.base - 0.8 * q.size], [q.x + q.w, q.base - 0.8 * q.size], [q.x, q.base + 0.2 * q.size], [q.x + q.w, q.base + 0.2 * q.size]].map(v => place(q, v[0], v[1]));
    const xa = clamp(Math.floor(Math.min(...pts.map(v => v[0])) * K), 0, W - 1), xb = clamp(Math.ceil(Math.max(...pts.map(v => v[0])) * K), 0, W - 1), ya = clamp(Math.floor(Math.min(...pts.map(v => v[1])) * K), 0, H - 1), yb = clamp(Math.ceil(Math.max(...pts.map(v => v[1])) * K), 0, H - 1);
    const cnt = new Map(); let total = 0;
    for (let y = ya; y <= yb; y += 2) for (let x = xa, i = (y * W + xa) * 4; x <= xb; x += 2, i += 8) { const k = (d[i] >> 3) << 10 | (d[i + 1] >> 3) << 5 | d[i + 2] >> 3; cnt.set(k, (cnt.get(k) || 0) + 1); total++; }
    let best = 0, bk = 0; for (const [k, n] of cnt) if (n > best) { best = n; bk = k; }
    const bgc = [(bk >> 10) << 3 | 4, ((bk >> 5) & 31) << 3 | 4, (bk & 31) << 3 | 4]; q.bgc = best > total * 0.5 && Math.min(...bgc) < 236 ? hex3(bgc) : null;
    if (!q.sample) continue;
    let md = 0; const hits = [];
    for (let y = ya; y <= yb; y++) for (let x = xa, i = (y * W + xa) * 4; x <= xb; x++, i += 4) { const dd = Math.abs(rd[i] - d[i]) + Math.abs(rd[i + 1] - d[i + 1]) + Math.abs(rd[i + 2] - d[i + 2]); if (dd > 40) hits.push(dd, i); if (dd > md) md = dd; }
    if (md < 60) { q.ghost = true; continue; }
    let r = 0, g = 0, bl = 0, n = 0; for (let k = 0; k < hits.length; k += 2) if (hits[k] >= 0.8 * md) { const i = hits[k + 1]; r += rd[i]; g += rd[i + 1]; bl += rd[i + 2]; n++; }
    const c = [r / n, g / n, bl / n].map(Math.round); q.color = Math.max(...c) < 70 ? '#000000' : hex3(c);
  }
  return { p, PW, PH, K, W, H, bg, ref, d, rd, pieces, place, rot };
}
async function samePage(p, B) {
  const M = await pageModel(p, DK), { PW, PH, K, W, H, bg, d, pieces, place, rot } = M, sc = Math.min(1, 1584 / Math.max(PW, PH)); // (Word pages are at most 22 inches)
  // the page design, cropped to where something is drawn (a page of plain text needs no picture at all)
  let x0 = W, y0 = H, x1 = -1, y1 = -1, back = '';
  for (let y = 0; y < H; y++) for (let x = 0, i = y * W * 4; x < W; x++, i += 4) if (d[i] < 248 || d[i + 1] < 248 || d[i + 2] < 248) { if (x < x0) x0 = x; if (x > x1) x1 = x; if (y < y0) y0 = y; y1 = y; }
  if (x1 >= 0) {
    const c = document.createElement('canvas'); c.width = x1 - x0 + 1; c.height = y1 - y0 + 1; c.getContext('2d').drawImage(bg, x0, y0, c.width, c.height, 0, 0, c.width, c.height);
    let blob = await toBlob(c, 'image/png'), ext = 'png';
    if (blob.size > 200000) { const j = await toBlob(c, 'image/jpeg', 0.9); if (j && j.size < blob.size * 0.6) { blob = j; ext = 'jpeg'; } } // photos: JPEG; line art stays crisp PNG
    const bytes = new Uint8Array(await blob.arrayBuffer());
    back = B.pic({ bytes, ext, key: ext + fnv(bytes) }, c.width / K * sc, c.height / K * sc, { x: x0 / K * sc, y: y0 / K * sc });
  }
  const groups = new Map(); // text set at the same angle around the same point is laid out together
  for (const q of pieces) { const k = q.rot ? q.rot + '|' + q.pv.map(v => v.toFixed(1)).join(',') : '0'; if (!groups.has(k)) groups.set(k, []); groups.get(k).push(q); }
  let boxes = '';
  for (const g of groups.values()) for (const b of blocksOf(linesOf(g))) boxes += blockXml(b, B, place, rot, sc);
  const m = 18 * sc;
  return { w: PW * sc, h: PH * sc, margins: { t: m, l: m, r: m, b: m }, paras: [{ ppr: '<w:spacing w:before="0" w:after="0" w:line="20" w:lineRule="exact"/><w:rPr><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr>', body: back + boxes }] };
}

/* ───────── PDF → DOCX, "editable": real Word paragraphs, headings, lists and tables in reading order ───────── */
const EK = 1.5; // render scale for colours and pictures
const BULLET_RE = /^\s*([•●○◦▪■□►▸‣⁃∙·\-–—*✓✔➢➤])\s+/, NUM_RE = /^\s*(\d{1,3}|[a-zA-Z]|[ivxlcdm]{1,5})([.)])\s+/;
function graphicsOf(M, boxesToSkip) { // drawn things on the page (no text): rules, filled boxes, charts, logos, photos → [{ x, y, w, h, kind, fill }]
  const { d, W, H, K } = M, C = 2, gw = Math.ceil(W / C), gh = Math.ceil(H / C), on = new Uint8Array(gw * gh);
  for (let y = 0; y < H; y++) for (let x = 0, i = y * W * 4; x < W; x++, i += 4) if (d[i] < 238 || d[i + 1] < 238 || d[i + 2] < 238) on[(y / C | 0) * gw + (x / C | 0)] = 1;
  for (const t of boxesToSkip) for (let y = Math.max(0, Math.floor((t.y - 3) * K / C)); y < Math.min(gh, Math.ceil((t.y + t.h + 3) * K / C)); y++) for (let x = Math.max(0, Math.floor((t.x - 3) * K / C)); x < Math.min(gw, Math.ceil((t.x + t.w + 3) * K / C)); x++) on[y * gw + x] = 0; // a table draws its own frame and shading
  const R = Math.max(1, Math.round(6 * K / C)), dil = new Uint8Array(gw * gh); // things closer than ~6 pt belong together (the bars of one chart)
  for (let y = 0; y < gh; y++) for (let x = 0; x < gw; x++) if (on[y * gw + x]) for (let yy = Math.max(0, y - R); yy <= Math.min(gh - 1, y + R); yy++) { const row = yy * gw; for (let xx = Math.max(0, x - R); xx <= Math.min(gw - 1, x + R); xx++) dil[row + xx] = 1; }
  const seen = new Uint8Array(gw * gh), out = [], stack = [];
  for (let s0 = 0; s0 < gw * gh; s0++) {
    if (!dil[s0] || seen[s0]) continue; let x0 = gw, y0 = gh, x1 = 0, y1 = 0, cnt = 0; stack.push(s0); seen[s0] = 1;
    while (stack.length) { const s = stack.pop(), x = s % gw, y = s / gw | 0; if (on[s]) { cnt++; if (x < x0) x0 = x; if (x > x1) x1 = x; if (y < y0) y0 = y; if (y > y1) y1 = y; }
      for (const n of [s - 1, s + 1, s - gw, s + gw]) if (n >= 0 && n < gw * gh && !seen[n] && dil[n] && Math.abs((n % gw) - x) <= 1) { seen[n] = 1; stack.push(n); } }
    if (!cnt) continue;
    const b = { x: x0 * C / K, y: y0 * C / K, w: (x1 - x0 + 1) * C / K, h: (y1 - y0 + 1) * C / K };
    if (b.w < 4 && b.h < 4) continue;
    const tb = boxesToSkip.find(t => b.y >= t.y - 8 && b.y + b.h <= t.y + t.h + 8 && b.x < t.x + t.w + 14 && b.x + b.w > t.x - 14);
    if (tb) { const x1 = Math.max(tb.x + tb.w, b.x + b.w); tb.x = Math.min(tb.x, b.x); tb.w = x1 - tb.x; continue; } // more of the table's frame (columns whose text is short): the table grows to it
    const ov = t => Math.max(0, Math.min(b.x + b.w, t.x + t.w) - Math.max(b.x, t.x)) * Math.max(0, Math.min(b.y + b.h, t.y + t.h) - Math.max(b.y, t.y));
    if (boxesToSkip.some(t => ov(t) > 0.5 * Math.min(b.w * b.h, t.w * t.h))) continue; // (a table's own frame, lines and shading are drawn by the table)
    // one colour filling most of the box → a coloured band (it becomes paragraph shading); a thin long box → a rule; anything else → a picture
    const cntc = new Map(); let tot = 0; for (let y = Math.floor(b.y * K); y < (b.y + b.h) * K; y += 3) for (let x = Math.floor(b.x * K); x < (b.x + b.w) * K; x += 3) { const i = (y * W + x) * 4, k = (d[i] >> 4) << 8 | (d[i + 1] >> 4) << 4 | d[i + 2] >> 4; cntc.set(k, (cntc.get(k) || 0) + 1); tot++; }
    let best = 0, bk = 0; for (const [k, n] of cntc) if (n > best) { best = n; bk = k; }
    const fill = '#' + [bk >> 8, (bk >> 4) & 15, bk & 15].map(v => (v * 17).toString(16).padStart(2, '0')).join('');
    b.kind = b.h <= 3.5 && b.w >= 30 ? 'rule' : b.w <= 3.5 && b.h >= 30 ? 'vrule' : best > tot * 0.7 && Math.max(...[bk >> 8, (bk >> 4) & 15, bk & 15]) * 17 < 250 ? 'fill' : 'pic'; b.fill = fill;
    out.push(b);
  }
  return out;
}
function tablesOf(lines, PW) { // two or more rows in a row whose cells line up in columns → a table: { x, y, w, h, cols: [x…], rows: [[cell lines…]…] }
  const rows = [];
  for (const l of [...lines].sort((a, b) => a.base - b.base || a.x - b.x)) { const r = rows[rows.length - 1]; if (r && Math.abs(r.base - l.base) < 0.35 * Math.min(r.size, l.size)) { r.cells.push(l); r.size = Math.max(r.size, l.size); } else rows.push({ base: l.base, size: l.size, cells: [l] }); }
  rows.forEach(r => r.cells.sort((a, b) => a.x - b.x));
  const multi = r => r.cells.length >= 2, out = [], fits = (r, cols) => r.cells.every(c => cols.some(k => c.x < k.r + 4 && c.right > k.l - 4)) && r.cells.length <= cols.length;
  for (let i = 0; i < rows.length; i++) {
    if (!multi(rows[i])) continue;
    const grp = [rows[i]]; let cols = rows[i].cells.map(c => ({ l: c.x, r: c.right }));
    for (let j = i + 1; j < rows.length; j++) {
      const r = rows[j], gap = r.base - grp[grp.length - 1].base; if (gap > 3.2 * Math.max(r.size, grp[0].size)) break;
      if (multi(r) && (fits(r, cols) || r.cells.length === cols.length)) { grp.push(r); if (r.cells.length > cols.length) cols = r.cells.map(c => ({ l: c.x, r: c.right })); else r.cells.forEach(c => { const k = cols.find(k => c.x < k.r + 4 && c.right > k.l - 4); if (k) { k.l = Math.min(k.l, c.x); k.r = Math.max(k.r, c.right); } }); }
      else if (!multi(r) && cols.some(k => r.cells[0].x >= k.l - 2 && r.cells[0].right <= k.r + 30) && gap < 1.6 * r.size) grp.push(r); // a cell's second line
      else break;
    }
    const realRows = grp.filter(multi); if (realRows.length < 2 || cols.length < 2) continue;
    const cellW = median(realRows.flatMap(r => r.cells.map(c => c.right - c.x))), cellChars = median(realRows.flatMap(r => r.cells.map(c => c.runs.map(q => q.str).join('').length)));
    if (cols.length <= 3 && cellW > 0.26 * PW && cellChars > 28) continue; // (side-by-side columns of running text)
    const T = { cols: cols.sort((a, b) => a.l - b.l), rows: [] };
    for (const r of grp) {
      if (multi(r) || !T.rows.length) { const row = cols.map(() => []); for (const c of r.cells) { let k = cols.findIndex(k => c.x < k.r + 4 && c.right > k.l - 4); if (k < 0) k = cols.reduce((bi, k2, ki) => Math.abs(k2.l - c.x) < Math.abs(cols[bi].l - c.x) ? ki : bi, 0); row[k].push(c); } T.rows.push({ base: r.base, size: r.size, cells: row }); }
      else { const c = r.cells[0], k = Math.max(0, cols.findIndex(k => c.x < k.r + 30 && c.right > k.l - 4)); T.rows[T.rows.length - 1].cells[k].push(c); }
    }
    const all = grp.flatMap(r => r.cells); T.x = Math.min(...all.map(c => c.x)) - 6; T.y = Math.min(...all.map(c => c.base - c.size)) - 4; T.w = Math.max(...all.map(c => c.right)) + 6 - T.x; T.h = Math.max(...all.map(c => c.base + 0.3 * c.size)) + 4 - T.y;
    T.lines = new Set(all); out.push(T); i += grp.length - 1;
  }
  return out;
}
function readingOrder(items) { // XY-cut: columns first (left before right), else split at the widest gap between rows, and again inside each part
  if (items.length <= 1) return items;
  const s = [...items].sort((a, b) => a.x - b.x), cols = [[s[0]]]; let end = s[0].x + s[0].w;
  for (const it of s.slice(1)) { if (it.x - end >= 9) cols.push([]); cols[cols.length - 1].push(it); end = Math.max(end, it.x + it.w); }
  if (cols.length > 1) return cols.flatMap(readingOrder);
  const r = [...items].sort((a, b) => a.y - b.y); let best = -1, at = -1, bot = r[0].y + r[0].h;
  for (let i = 1; i < r.length; i++) { const g = r[i].y - bot; if (g > best) { best = g; at = i; } bot = Math.max(bot, r[i].y + r[i].h); }
  if (best >= 1) return [...readingOrder(r.slice(0, at)), ...readingOrder(r.slice(at))];
  return r.sort((a, b) => a.y - b.y || a.x - b.x);
}
const median = a => { if (!a.length) return 0; const s = [...a].sort((x, y) => x - y); return s[s.length >> 1]; };
function paragraphsOf(lines, bodyPt = 99) { // lines in reading order → paragraphs (re-joined, so Word can re-wrap them), one column at a time
  const runs = []; for (const l of lines) { const g = runs[runs.length - 1], pl = g && g[g.length - 1]; if (g && l.base > pl.base && l.x < pl.right && l.right > pl.x) g.push(l); else runs.push([l]); }
  return runs.flatMap(g => paragraphsOfColumn(g, bodyPt));
}
function paragraphsOfColumn(L, bodyPt) {
  const out = []; let cur = null;
  const left = Math.min(...L.map(l => l.x)), right = Math.max(...L.map(l => l.right)), width = right - left;
  const sig = l => [Math.round(l.size), l.runs[0].bold, l.runs[0].font].join('|');
  for (const l of L) {
    const text = l.runs.map(r => r.str).join(' '), bullet = l.size < bodyPt * 1.12 && (l.dot ? ['•'] : BULLET_RE.exec(text) || NUM_RE.exec(text)); // ("1. Summary" set large is a numbered heading, not a list)
    const gap = cur ? l.base - cur.lines[cur.lines.length - 1].base : 0, prev = cur && cur.lines[cur.lines.length - 1];
    const cont = cur && !bullet && sig(l) === cur.sig && gap > 0.6 * l.size && gap < 1.75 * l.size && (cur.gap == null || Math.abs(gap - cur.gap) < 0.25 * l.size)
      && (cur.list ? l.x > cur.lines[0].x + 2 && Math.abs(l.x - (cur.lines[1] || l).x) < 4 : Math.abs(l.x - (cur.lines[1] || l).x) < 4 || cur.lines.length === 1 && l.x < cur.lines[0].x)
      && (prev.right > right - Math.max(3 * l.size, 0.18 * width) || cur.center);
    if (cont) { if (cur.gap == null) cur.gap = gap; cur.lines.push(l); continue; }
    cur = { lines: [l], sig: sig(l), list: bullet ? (l.dot || BULLET_RE.test(text) ? 'bullet' : 'number') : null, num: bullet && !l.dot && !BULLET_RE.test(text) ? bullet[1] : null, gap: null, left, right };
    out.push(cur);
  }
  for (const P of out) {
    const ls = P.lines, c0 = (left + right) / 2, centred = ls.every(l => Math.abs((l.x + l.right) / 2 - c0) < 4) && ls.some(l => l.x > left + 8);
    P.align = centred ? 'center' : ls.length > 1 && ls.every(l => Math.abs(l.right - right) < 2) && ls[0].x > left + 8 ? 'right' : ls.length > 2 && ls.slice(0, -1).every(l => Math.abs(l.right - right) < 2.5) ? 'both' : 'left';
    P.size = median(ls.map(l => l.size)); P.y = ls[0].base - ls[0].size; P.bottom = ls[ls.length - 1].base + 0.3 * P.size; P.x = Math.min(...ls.map(l => l.x)); P.colLeft = left;
  }
  return out;
}
function runsOf(P, ctx) { // the paragraph's text as runs, a run per change of style; lines are joined (a hyphen at a line end joins the word)
  const out = []; let prevLine = null;
  for (const l of P.lines) {
    l.runs.forEach((q, i) => {
      let t = q.str; const nx = l.runs[i + 1];
      if (i === 0 && prevLine) { const last = out[out.length - 1]; if (last && /[-‐]$/.test(last.t) && /^\p{Ll}/u.test(t)) last.t = last.t.replace(/[-‐]$/, ''); else if (last && !/\s$/.test(last.t)) last.t += ' '; }
      if (nx && nx.x - (q.x + q.w) > 0.12 * q.size && !/\s$/.test(t) && !/^\s/.test(nx.str)) t += ' ';
      const st = { font: q.font, sz: Math.round(q.size * 2), b: !!q.bold, i: !!q.italic, u: !!q.underline, s: !!q.strike, color: q.ghost ? null : q.color, shd: q.bgc && !P.shd ? q.bgc : null, pos: Math.abs(q.base - l.base) > 0.2 * l.size ? Math.round((l.base - q.base) * 2) : 0 };
      const last = out[out.length - 1];
      if (last && JSON.stringify(last.st) === JSON.stringify(st)) last.t += t; else out.push({ t, st });
    });
    prevLine = l;
  }
  if (out.length) { out[0].t = out[0].t.replace(/^\s+/, ''); if (P.list && !P.lines[0].dot) out[0].t = out[0].t.replace(P.list === 'bullet' ? BULLET_RE : NUM_RE, ''); out[out.length - 1].t = out[out.length - 1].t.replace(/\s+$/, ''); }
  return out.filter(r => r.t);
}
function rprOf(st, base) { // only what differs from the paragraph's style
  const f = st.font !== base.font ? `<w:rFonts w:ascii="${xesc(st.font)}" w:hAnsi="${xesc(st.font)}" w:eastAsia="${xesc(st.font)}" w:cs="${xesc(st.font)}"/>` : '';
  const col = st.color && st.color.toLowerCase() !== (base.color || '#000000') ? `<w:color w:val="${st.color.slice(1).toUpperCase()}"/>` : '';
  const x = f + (st.b !== !!base.b ? (st.b ? '<w:b/><w:bCs/>' : '<w:b w:val="0"/><w:bCs w:val="0"/>') : '') + (st.i !== !!base.i ? (st.i ? '<w:i/><w:iCs/>' : '<w:i w:val="0"/>') : '') + (st.s ? '<w:strike/>' : '') + col
    + (st.pos ? `<w:position w:val="${st.pos}"/>` : '') + (st.sz !== base.sz ? `<w:sz w:val="${st.sz}"/><w:szCs w:val="${st.sz}"/>` : '') + (st.u ? '<w:u w:val="single"/>' : '') + (st.shd ? `<w:shd w:val="clear" w:color="auto" w:fill="${st.shd.slice(1).toUpperCase()}"/>` : '');
  return x ? `<w:rPr>${x}</w:rPr>` : '';
}
const runsXml = (runs, base) => runs.map(r => `<w:r>${rprOf(r.st, base)}<w:t xml:space="preserve">${xesc(r.t)}</w:t></w:r>`).join('');
async function editModel(p) { // pass 1 for one page: lines, tables, drawn things and pictures, in reading order
  const M = await pageModel(p, EK), scan = M.pieces.length && M.pieces.every(q => q.ghost);
  if (scan) M.pieces.forEach(q => { q.ghost = false; q.color = '#000000'; q.bgc = null; }); // a scan with its text recognised: the text is what becomes editable
  const straight = M.pieces.filter(q => !q.rot && !q.ghost), turned = M.pieces.filter(q => q.rot);
  const lines = linesOf(straight).map(l => ({ ...l, h: 1.2 * l.size, y: l.base - l.size, w: l.right - l.x, t: 'line' }));
  const tables = tablesOf(lines, M.PW);
  const gfx = scan ? [] : graphicsOf(M, tables);
  const items = [];
  for (const T of tables) {
    const shade = (x0, y0, x1, y1) => { const { d, W, K } = M, cnt = new Map(); let tot = 0; for (let y = Math.floor(y0 * K); y < y1 * K; y += 2) for (let x = Math.floor(x0 * K); x < x1 * K; x += 2) { const i = (y * W + x) * 4, k = (d[i] >> 4) << 8 | (d[i + 1] >> 4) << 4 | d[i + 2] >> 4; cnt.set(k, (cnt.get(k) || 0) + 1); tot++; } let best = 0, bk = 0xfff; for (const [k, n] of cnt) if (n > best) { best = n; bk = k; } return Math.min(bk >> 8, (bk >> 4) & 15, bk & 15) * 17 < 236 ? '#' + [bk >> 8, (bk >> 4) & 15, bk & 15].map(v => (v * 17).toString(16).padStart(2, '0')).join('') : null; };
    const edges = T.cols.map((c, i) => i ? (T.cols[i - 1].r + c.l) / 2 : T.x).concat([T.x + T.w]);
    T.rows.forEach((r, ri) => { const y0 = ri ? (T.rows[ri - 1].base + r.base - r.size) / 2 + 2 : T.y, y1 = ri < T.rows.length - 1 ? (r.base + T.rows[ri + 1].base - T.rows[ri + 1].size) / 2 : T.y + T.h; r.shd = r.cells.map((c, ci) => shade(edges[ci] + 2, y0 + 1, edges[ci + 1] - 2, y1 - 1)); });
    const { d, W, K } = M, dark = (x, y) => { const i = ((Math.round(y * K)) * W + Math.round(x * K)) * 4; return d[i] < 200 || d[i + 1] < 200 || d[i + 2] < 200; };
    let hit = 0, n = 0; for (let x = T.x + 8; x < T.x + T.w - 8; x += 6) { n++; for (let dy = -6; dy <= 6; dy++) if (dark(x, T.rows[0].base + 0.3 * T.rows[0].size + 2 + dy) || dark(x, T.y + T.h - 4 + dy)) { hit++; break; } }
    T.borders = n && hit > n * 0.6; T.edges = edges;
    items.push({ t: 'table', x: T.x, y: T.y, w: T.w, h: T.h, T });
  }
  const flow = lines.filter(l => !tables.some(T => T.lines.has(l)));
  const pics = [];
  for (const g of gfx) {
    if (g.kind === 'vrule') continue;
    if (g.kind === 'rule') { items.push({ t: 'rule', x: g.x, y: g.y, w: g.w, h: Math.max(g.h, 1), color: g.fill }); continue; }
    const inside = flow.filter(l => l.x >= g.x - 2 && l.right <= g.x + g.w + 2 && l.y >= g.y - 2 && l.base <= g.y + g.h + 2);
    if (g.kind === 'fill') { for (const l of inside) l.fill = g.fill; if (!inside.length && g.w * g.h > 400 && !flow.some(l => l.x < g.x + g.w && l.right > g.x && l.y < g.y + g.h && l.base > g.y)) pics.push(g); continue; } // a coloured band: its text gets the shading (behind part of a line: that text is highlighted instead)
    const marks = g.w <= 9 ? flow.filter(l => l.y < g.y + g.h && l.base > g.y && l.x > g.x + g.w && l.x - (g.x + g.w) < 22) : [];
    if (marks.length) { for (const l of marks) l.dot = true; continue; } // drawn list bullets (a web page's <ul>): the lines become a bulleted list
    if (g.w * g.h < 36) continue;
    g.text = inside; pics.push(g);
  }
  for (const g of pics) { // a picture: a photo, logo or chart — the labels inside a chart go with it
    const c = document.createElement('canvas'); c.width = Math.max(1, Math.round(g.w * M.K)); c.height = Math.max(1, Math.round(g.h * M.K));
    c.getContext('2d').drawImage(g.text && g.text.length ? M.ref : M.bg, Math.round(g.x * M.K), Math.round(g.y * M.K), c.width, c.height, 0, 0, c.width, c.height);
    let blob = await toBlob(c, 'image/png'), ext = 'png'; if (blob.size > 150000) { const j = await toBlob(c, 'image/jpeg', 0.88); if (j && j.size < blob.size * 0.6) { blob = j; ext = 'jpeg'; } }
    const bytes = new Uint8Array(await blob.arrayBuffer()); items.push({ t: 'pic', x: g.x, y: g.y, w: g.w, h: g.h, bytes, ext });
    if (g.text) for (const l of g.text) l.inPic = true;
  }
  for (const l of flow) if (!l.inPic) items.push(l);
  return { p, PW: M.PW, PH: M.PH, items: readingOrder(items), turned, place: M.place };
}
function hfOf(models) { // the same text at the top (bottom) of most pages → the header (footer); page numbers become Word page fields
  if (models.length < 2) return { header: null, footer: null };
  const key = l => l.runs.map(r => r.str).join(' ').replace(/\b\d{1,4}\b/g, '#').trim(), pick = (zone) => {
    const seen = new Map();
    for (const m of models) for (const it of m.items) if (it.t === 'line' && zone(it, m)) { const k = key(it); if (!seen.has(k)) seen.set(k, new Set()); seen.get(k).add(m); }
    const keys = [...seen].filter(([, s]) => s.size >= Math.max(2, models.length * 0.6)).map(([k]) => k); if (!keys.length) return null;
    for (const m of models) m.items = m.items.filter(it => !(it.t === 'line' && zone(it, m) && keys.includes(key(it))));
    return keys;
  };
  return { header: pick((l, m) => l.base < m.PH * 0.09), footer: pick((l, m) => l.y > m.PH * 0.9) };
}
const hfXml = keys => (keys || []).map(k => `<w:p><w:pPr><w:jc w:val="center"/></w:pPr>${k.split('#').map((t, i, a) => (t ? `<w:r><w:t xml:space="preserve">${xesc(t)}</w:t></w:r>` : '') + (i < a.length - 1 ? `<w:fldSimple w:instr=" ${/of\s*$/i.test(t) ? 'NUMPAGES' : 'PAGE'} "><w:r><w:t>1</w:t></w:r></w:fldSimple>` : '')).join('')}</w:p>`).join('');
async function buildEditable(pages, B) { // → { body, extra } for docxBuilder.finish
  const models = []; for (const p of pages) models.push(await editModel(p));
  const allL = models.flatMap(m => m.items.filter(i => i.t === 'line').concat(m.items.filter(i => i.t === 'table').flatMap(i => [...i.T.lines])));
  const chars = new Map(), fonts = new Map(), cols = new Map(), add = (m, k, n) => m.set(k, (m.get(k) || 0) + n);
  for (const l of allL) for (const q of l.runs) { const n = q.str.length; add(chars, Math.round(q.size * 2), n); add(fonts, q.font, n); add(cols, q.color || '#000000', n); }
  const top = m => [...m].sort((a, b) => b[1] - a[1])[0] || [];
  const body = { sz: top(chars)[0] || 22, font: top(fonts)[0] || 'Calibri', color: top(cols)[0] || '#000000', b: false, i: false }, bodyPt = body.sz / 2;
  // heading sizes: noticeably bigger than the body text (or bold lines at body size standing alone)
  const hsizes = [...new Set(allL.filter(l => l.size > bodyPt * 1.18).map(l => Math.round(l.size * 2)))].sort((a, b) => b - a);
  const title = hsizes[0] && hsizes[0] >= body.sz * 1.7 && hsizes.length > 1 ? hsizes.shift() : null, levels = hsizes.slice(0, 3);
  const hf = hfOf(models), heads = {}, firstOf = (sz, bold) => allL.find(l => Math.round(l.size * 2) === sz && (bold == null || !!l.runs[0].bold === bold));
  const styleRun = l => l ? { font: l.runs[0].font, sz: Math.round(l.size * 2), b: !!l.runs[0].bold, i: !!l.runs[0].italic, color: l.runs[0].color || '#000000' } : null;
  if (title) heads.Title = styleRun(firstOf(title));
  levels.forEach((sz, i) => { heads['Heading' + (i + 1)] = styleRun(firstOf(sz)); });
  const boldH = allL.some(l => Math.abs(l.size - bodyPt) < 0.6 && l.runs.every(q => q.bold)) && !heads.Heading3 ? 'Heading' + (levels.length + 1) : null;
  if (boldH) heads[boldH] = { ...body, b: true };
  const styleFor = P => { const sz = Math.round(P.size * 2), l = P.lines[0]; if (P.lines.length > 2) return null; if (title && sz === title) return 'Title'; const k = levels.indexOf(sz); if (k >= 0) return 'Heading' + (k + 1); if (boldH && P.lines.length === 1 && l.runs.every(q => q.bold) && Math.abs(P.size - bodyPt) < 0.6 && l.runs.map(q => q.str).join('').length < 90) return boldH; return null; };
  const nums = []; let numId = 2, out = '', firstPage = true;
  const cw = m => m.PW - 2 * m.margin;
  for (const m of models) {
    const ls = m.items.filter(i => i.t === 'line'), lft = ls.length ? Math.min(...ls.map(l => l.x)) : 54;
    m.margin = clamp(Math.min(lft, m.items.length ? Math.min(...m.items.map(i => i.x)) : lft), 18, m.PW / 4);
    let prevBottom = null, pending = [], list = null, firstPara = true;
    const ppr = (extra, before, P) => `<w:pPr>${extra}${firstPara && !firstPage ? '<w:pageBreakBefore/>' : ''}<w:spacing w:before="${twip(clamp(before, 0, 48))}" w:after="0"${P && P.line ? ` w:line="${P.line}" w:lineRule="auto"` : ''}/>${P && P.ind ? P.ind : ''}${P && P.align && P.align !== 'left' ? `<w:jc w:val="${P.align}"/>` : ''}</w:pPr>`;
    const flush = () => {
      if (!pending.length) return;
      for (const P of paragraphsOf(pending, bodyPt)) {
        const before = prevBottom == null ? 0 : P.y - prevBottom - 2, st = styleFor(P), base = st ? heads[st] : body;
        const fill = P.lines.every(l => l.fill) ? P.lines[0].fill : null; if (fill) P.shd = fill;
        const runs = runsOf(P); if (!runs.length) continue;
        P.line = P.gap && !st ? clamp(Math.round(P.gap / (1.17 * P.size) * 240), 200, 600) : 0;
        const ref = P.colLeft > m.margin + 0.25 * cw(m) ? P.colLeft : m.margin, indL = P.x - ref, first = P.lines.length > 1 ? P.lines[0].x - P.lines[1].x : 0; // (a right-hand column flows on from the margin)
        P.ind = P.align === 'left' || P.align === 'both' ? (indL > 3 || Math.abs(first) > 3 ? `<w:ind w:left="${twip(Math.max(0, P.list ? indL + 18 : P.lines.length > 1 ? P.lines[1].x - ref : indL))}"${P.list ? ' w:hanging="360"' : first > 3 ? ` w:firstLine="${twip(first)}"` : ''}/>` : P.list ? '<w:ind w:left="720" w:hanging="360"/>' : '') : '';
        let numPr = '';
        if (P.list) { if (!list || list.kind !== P.list) { list = { kind: P.list, id: ++numId, start: P.num && /^\d+$/.test(P.num) ? +P.num : 1 }; nums.push(list); } numPr = `<w:numPr><w:ilvl w:val="0"/><w:numId w:val="${list.id}"/></w:numPr>`; } else list = null;
        const pStyle = st ? `<w:pStyle w:val="${st}"/>` : P.list ? '<w:pStyle w:val="ListParagraph"/>' : '';
        out += `<w:p>${ppr(pStyle + (st ? '<w:keepNext/>' : '') + numPr + (P.shd ? `<w:shd w:val="clear" w:color="auto" w:fill="${P.shd.slice(1).toUpperCase()}"/>` : ''), before, P)}${runsXml(runs, { ...base, color: base.color || '#000000' })}</w:p>`;
        firstPara = false; prevBottom = P.bottom;
      }
      pending = [];
    };
    for (const it of m.items) {
      if (it.t === 'line') { pending.push(it); continue; }
      flush(); list = null; const before = prevBottom == null ? 0 : it.y - prevBottom;
      if (it.t === 'rule') { out += `<w:p>${ppr(`<w:pBdr><w:bottom w:val="single" w:sz="${clamp(Math.round(it.h * 8), 4, 24)}" w:space="1" w:color="${(it.color || '#000000').slice(1).toUpperCase()}"/></w:pBdr>`, before)}</w:p>`; }
      else if (it.t === 'pic') { const f = Math.min(1, cw(m) / it.w), cx = (it.x + it.w / 2 - m.PW / 2) / m.PW; out += `<w:p>${ppr(Math.abs(cx) < 0.06 ? '<w:jc w:val="center"/>' : it.x - m.margin > 6 ? `<w:ind w:left="${twip(Math.min(it.x - m.margin, cw(m) - it.w * f))}"/>` : '', before)}${B.pic({ bytes: it.bytes, ext: it.ext, key: it.ext + fnv(it.bytes) }, it.w * f, it.h * f)}</w:p>`; }
      else if (it.t === 'table') out += tableXml(it.T, m, body, before, firstPara && !firstPage);
      firstPara = false; prevBottom = it.y + it.h;
    }
    flush();
    for (const g of [...new Map(m.turned.map(q => [q.rot + '|' + q.pv.join(','), null])).keys()]) { // turned text (a margin label) stays a small text box at its place
      const grp = m.turned.filter(q => q.rot + '|' + q.pv.join(',') === g); for (const b of blocksOf(linesOf(grp))) out += `<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="20" w:lineRule="exact"/></w:pPr>${blockXml(b, B, m.place, 0, 1)}</w:p>`;
    }
    if (firstPara) out += `<w:p>${!firstPage ? '<w:pPr><w:pageBreakBefore/></w:pPr>' : ''}</w:p>`;
    firstPage = false;
  }
  const m0 = models[0], mar = Math.round(m0 ? m0.margin : 54), hRef = (hf.header ? '<w:headerReference w:type="default" r:id="rId4"/>' : '') + (hf.footer ? '<w:footerReference w:type="default" r:id="rId5"/>' : '');
  out += `<w:sectPr>${hRef}<w:pgSz w:w="${twip(m0 ? m0.PW : 595)}" w:h="${twip(m0 ? m0.PH : 842)}"${m0 && m0.PW > m0.PH ? ' w:orient="landscape"' : ''}/><w:pgMar w:top="${twip(Math.max(36, mar))}" w:right="${twip(mar)}" w:bottom="${twip(Math.max(36, mar))}" w:left="${twip(mar)}" w:header="${twip(Math.min(28, mar / 2))}" w:footer="${twip(Math.min(28, mar / 2))}" w:gutter="0"/></w:sectPr>`;
  const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"', rpr = (st, def) => { if (!st) return ''; const f = xesc(st.font); return `<w:rPr><w:rFonts w:ascii="${f}" w:hAnsi="${f}" w:eastAsia="${f}" w:cs="${f}"/>${st.b ? '<w:b/><w:bCs/>' : def ? '' : '<w:b w:val="0"/>'}${st.i ? '<w:i/><w:iCs/>' : ''}${st.color && st.color !== '#000000' ? `<w:color w:val="${st.color.slice(1).toUpperCase()}"/>` : ''}<w:sz w:val="${st.sz}"/><w:szCs w:val="${st.sz}"/></w:rPr>`; };
  const hst = (id, name, lvl) => heads[id] ? `<w:style w:type="paragraph" w:styleId="${id}"><w:name w:val="${name}"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="0" w:after="0"/>${lvl != null ? `<w:outlineLvl w:val="${lvl}"/>` : ''}</w:pPr>${rpr(heads[id])}</w:style>` : '';
  const styles = `<w:styles ${W}><w:docDefaults><w:rPrDefault>${rpr(body, true).replace('<w:rPr>', '<w:rPr>').replace('</w:rPr>', '<w:lang w:val="en-US"/></w:rPr>')}</w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>`
    + `<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>` + hst('Title', 'Title') + [1, 2, 3, 4].map(n => hst('Heading' + n, 'heading ' + n, n - 1)).join('')
    + `<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:ind w:left="720"/><w:contextualSpacing/></w:pPr></w:style>`
    + `<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders></w:tblPr></w:style></w:styles>`;
  const lvl = (fmt, txt, font) => `<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="${fmt}"/><w:lvlText w:val="${txt}"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>${font ? `<w:rPr><w:rFonts w:ascii="${font}" w:hAnsi="${font}" w:hint="default"/></w:rPr>` : ''}</w:lvl>`;
  const numbering = nums.length ? `<w:numbering ${W}><w:abstractNum w:abstractNumId="1"><w:multiLevelType w:val="singleLevel"/>${lvl('bullet', '•', 'Symbol').replace('w:val="•"', 'w:val=""')}</w:abstractNum><w:abstractNum w:abstractNumId="2"><w:multiLevelType w:val="singleLevel"/>${lvl('decimal', '%1.')}</w:abstractNum>${nums.map(n => `<w:num w:numId="${n.id}"><w:abstractNumId w:val="${n.kind === 'bullet' ? 1 : 2}"/>${n.kind !== 'bullet' && n.start !== 1 ? `<w:lvlOverride w:ilvl="0"><w:startOverride w:val="${n.start}"/></w:lvlOverride>` : ''}</w:num>`).join('')}</w:numbering>` : null;
  return { body: out, extra: { styles, numbering, header: hf.header ? hfXml(hf.header) : null, footer: hf.footer ? hfXml(hf.footer) : null } };
}
function tableXml(T, m, body, before, pageBreak) {
  const widths = T.edges.slice(1).map((e, i) => Math.max(18, e - T.edges[i])), tw = widths.reduce((a, b) => a + b, 0), f = Math.min(1, (m.PW - 2 * m.margin) / tw), bd = T.borders ? 'single' : 'nil';
  const border = `<w:tblBorders>${['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map(k => `<w:${k} w:val="${bd}" w:sz="4" w:space="0" w:color="A6A6A6"/>`).join('')}</w:tblBorders>`;
  const cellP = (lines, align) => { if (!lines.length) return '<w:p><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr></w:p>'; const L = [...lines].sort((a, b) => a.base - b.base || a.x - b.x), P = { lines: L, size: median(L.map(l => l.size)), shd: 'cell' }; return `<w:p><w:pPr><w:spacing w:before="0" w:after="0"/>${align !== 'left' ? `<w:jc w:val="${align}"/>` : ''}</w:pPr>${runsXml(runsOf(P), body)}</w:p>`; };
  const rows = T.rows.map((r, ri) => `<w:tr>${ri === 0 ? '<w:trPr><w:tblHeader/></w:trPr>' : ''}${r.cells.map((c, ci) => { const k = T.cols[ci], cellW = widths[ci], align = c.length && k && c.every(l => l.right > T.edges[ci + 1] - 14 && l.x > T.edges[ci] + cellW * 0.25) ? 'right' : 'left'; return `<w:tc><w:tcPr><w:tcW w:w="${twip(cellW * f)}" w:type="dxa"/>${r.shd && r.shd[ci] ? `<w:shd w:val="clear" w:color="auto" w:fill="${r.shd[ci].slice(1).toUpperCase()}"/>` : ''}</w:tcPr>${cellP(c, align)}</w:tc>`; }).join('')}</w:tr>`).join('');
  return `${pageBreak ? '<w:p><w:pPr><w:pageBreakBefore/><w:spacing w:before="0" w:after="0"/></w:pPr></w:p>' : before > 4 ? `<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="${twip(clamp(before, 2, 30))}" w:lineRule="exact"/></w:pPr></w:p>` : ''}<w:tbl><w:tblPr><w:tblW w:w="${twip(tw * f)}" w:type="dxa"/>${T.x - m.margin > 4 ? `<w:tblInd w:w="${twip(T.x - m.margin)}" w:type="dxa"/>` : ''}${border}<w:tblLayout w:type="fixed"/><w:tblCellMar><w:left w:w="100" w:type="dxa"/><w:right w:w="100" w:type="dxa"/></w:tblCellMar><w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/></w:tblPr><w:tblGrid>${widths.map(w => `<w:gridCol w:w="${twip(w * f)}"/>`).join('')}</w:tblGrid>${rows}</w:tbl>`;
}
async function buildDocx(pages, mode = 'same') { // → the .docx as a Blob
  await X.loadLib('jszip'); const B = docxBuilder();
  if (mode === 'exact' && X.pdf2word) return X.pdf2word.build(pages); // (js/pdf2word.js; the Blob carries .fallback: pages set as text boxes)
  if (mode === 'edit' || mode === 'flow') { const r = await buildEditable(pages, B); return B.finish(r.body, X.baseName(), r.extra); }
  let body = '';
  for (let i = 0; i < pages.length; i++) {
    const p = pages[i]; let out, w = p.w, h = p.h;
    if (mode === 'pic') { [w, h] = X.rdims(p); const c = await X.renderPageImage(p, 150 / 72); out = { paras: [{ ppr: '<w:spacing w:before="0" w:after="0"/>', body: B.pic(c.toDataURL('image/jpeg', 0.85), w, h, { x: 0, y: 0 }) }], margins: { t: 10, l: 10, r: 10, b: 10 } }; }
    else { out = await samePage(p, B); w = out.w; h = out.h; }
    if (!out.paras.length) out.paras.push({ ppr: '', body: '' });
    const sect = B.sect(w, h, out.margins), last = out.paras.length - 1;
    out.paras.forEach((P, k) => { body += k === last && i < pages.length - 1 ? `<w:p><w:pPr>${P.ppr}${sect}</w:pPr>${P.body}</w:p>` : `<w:p><w:pPr>${P.ppr}</w:pPr>${P.body}</w:p>`; });
    if (i === pages.length - 1) body += sect;
  }
  return B.finish(body, X.baseName());
}
async function pdfToDocx() {
  await X.loadLib('jszip');
  let last = 'exact'; try { last = localStorage.getItem('xd3.docxmode2') || 'exact'; } catch (e) { /* default */ }
  const r = await dialog('PDF → Word (.docx)', `
    <label>Layout<select name="mode">
      <option value="exact"${last === 'exact' ? ' selected' : ''}>Exact layout — looks the same, fully editable</option>
      <option value="same"${last === 'same' ? ' selected' : ''}>Fixed layout (text boxes)</option>
      <option value="edit"${last === 'edit' ? ' selected' : ''}>Simple flowing text</option>
      <option value="pic"${last === 'pic' ? ' selected' : ''}>Picture — each page as an image (not editable)</option></select></label>
    <label>Pages (blank = all), e.g. 1-3, 5<input type="text" name="range" autocomplete="off"></label>
    <p class="mut">Includes your edits. <b>Exact layout</b> keeps every line where it is in the PDF, with real paragraphs, tables and styles you can type into (best opened in Microsoft Word; each PDF page stays one Word page). <b>Fixed layout</b> puts the text in boxes at their places. <b>Simple flowing text</b> rebuilds plain paragraphs, headings, lists and tables. Scanned pages: run OCR first.</p>`, { ok: 'Convert' });
  if (!r) return; const pages = pickPages(r.range); if (!pages) return; try { localStorage.setItem('xd3.docxmode2', r.mode); } catch (e) { /* not remembered */ }
  await busy(async () => {
    const blob = await buildDocx(pages, r.mode), fw = blob.fallbackWhy || (blob.fallback || []).map(page => ({ page, why: 'complex' })); download(blob, X.baseName() + '.docx');
    const list = k => fw.filter(f => f.why === k).map(f => f.page), pl = a => `Page${a.length > 1 ? 's' : ''} ${a.join(', ')}`, notes = [];
    if (list('complex').length) notes.push(`${pl(list('complex'))}: too crowded to flow (text over text), so text boxes`);
    if (list('big').length) notes.push(`${pl(list('big'))}: larger than Word allows, so text boxes`);
    if (list('failed').length) notes.push(`${pl(list('failed'))}: could not be laid out, so text boxes`);
    if (list('upright').length) notes.push(`${pl(list('upright'))}: turned in the PDF, upright in Word so the text can flow`);
    if (list('redacted').length) notes.push(`${pl(list('redacted'))}: has blacked-out text, so it is a picture`);
    toast(notes.length ? `Word document saved. ${notes.join('. ')}` : 'Word document saved to your downloads');
  }, 'Converting to Word…');
}

/* ───────── DOCX → pages: laid out by Word's own rules (Word 2013+, measured), so a Word file keeps its pages ─────────
   Sections (page size, margins, columns, continuous breaks with balanced columns), headers / footers with PAGE / NUMPAGES,
   exact / at-least / multiple line spacing (exact L: baseline 0.8·L below the line top), spacing before / after (kept at the
   top of a page only after a hard break), indents, tab stops (left / right / centre / decimal, leaders), justified lines
   (spaces may shrink ~25%), letter spacing, paragraph borders and shading, keep-with-next / keep-lines / widow control,
   tables (grid, fixed widths, cell margins, row heights, borders, shading, merged cells, split and header rows),
   inline and anchored pictures and shapes (page / margin / column / paragraph relative, wrap), text boxes, embedded fonts. */
const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main', R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const kids = (el, n) => el ? [...el.children].filter(c => c.localName === n) : [];
const kid = (el, n) => { if (!el) return null; for (const c of el.children) if (c.localName === n) return c; return null; };
const av = (el, n) => el ? (el.getAttributeNS(W_NS, n) || el.getAttribute('w:' + n)) : null;
const anum = (el, n) => { const v = av(el, n); return v == null || v === '' || !isFinite(+v) ? null : +v; };
const onoff = el => { const v = av(el, 'val'); return !(v === '0' || v === 'false' || v === 'off'); };
const dxRid = (el, n) => el ? (el.getAttributeNS(R_NS, n) || el.getAttribute('r:' + n)) : null;
const dxDeep = (el, n) => el ? [...el.getElementsByTagNameNS('*', n)] : [];
const dxDeep1 = (el, n) => el ? el.getElementsByTagNameNS('*', n)[0] || null : null;
const dxHex = v => v && /^[0-9a-f]{6}$/i.test(v) ? '#' + v.toLowerCase() : null;
const DX_EMU = 12700, DX_RULE_OUT = 1.45, DX_SQUEEZE = 0.25;
const DX_WEIGHT = /\s+(thin|hairline|extralight|ultralight|light|semilight|book|regular|normal|medium|semibold|semi bold|demibold|demi|bold|extrabold|ultrabold|black|heavy)$/i;
function famOf(f) { // a Word font name → { fam: the family the page draws with, heavy: the name itself says bold }
  const find = n => Object.keys(X.BUNDLED).find(k => k.toLowerCase() === n.toLowerCase());
  let n = String(f || '').trim(), heavy = false, m, b = find(n); if (b) return { fam: b, heavy };
  while ((m = n.match(DX_WEIGHT))) { if (/semi|demi|bold|black|heavy/i.test(m[1])) heavy = true; n = n.slice(0, m.index); if ((b = find(n))) return { fam: b, heavy }; }
  return { fam: /courier|consol|mono|menlo/i.test(n) ? 'Courier' : /times|georgia|cambria|garamond|antiqua|palatino|minion|bookman|century|serif/i.test(n) && !/sans/i.test(n) ? 'Times' : 'Helvetica', heavy };
}
// Word's single line (hhea ascent + descent + line gap) and its descent, in em — the fonts Word documents use most
const DX_MET = { calibri: [1.2207, 0.2686], carlito: [1.2207, 0.2686], arial: [1.1499, 0.2119], arimo: [1.1499, 0.2119], helvetica: [1.1499, 0.2119], 'times new roman': [1.1499, 0.2163], tinos: [1.1499, 0.2163], times: [1.1499, 0.2163],
  georgia: [1.1362, 0.2192], gelasio: [1.1362, 0.2192], cambria: [1.1719, 0.2222], caladea: [1.1719, 0.2222], verdana: [1.2153, 0.2100], tahoma: [1.2070, 0.2065], 'segoe ui': [1.3301, 0.2510],
  'courier new': [1.1328, 0.3003], cousine: [1.1328, 0.3003], courier: [1.1328, 0.3003] };
const DXC = document.createElement('canvas').getContext('2d'); try { DXC.fontKerning = 'none'; DXC.textRendering = 'optimizeSpeed'; } catch (e) { /* older Chromium */ }
let dxFont = ''; const dxW = new Map(), dxLocal = new Map(), dxBox = new Map();
function dxW100(text, css) { const k = css + '\u0001' + text; let w = dxW.get(k); if (w == null) { if (dxFont !== css) { DXC.font = css; dxFont = css; } w = DXC.measureText(text).width; if (dxW.size > 250000) dxW.clear(); dxW.set(k, w); } return w; }
function dxInstalled(name) { // a font this computer has: Word's own face is measured then, whatever face draws it
  if (dxLocal.has(name)) return dxLocal.get(name);
  const q = `"${name.replace(/["\\]/g, '')}"`, t = 'mmmmmmmmmmlli WQ@#1abc', w = f => { DXC.font = '60px ' + f; dxFont = ''; return DXC.measureText(t).width; };
  const r = !/^(serif|sans-serif|monospace|cursive|fantasy)$/i.test(name) && (w(`${q}, monospace`) !== w('monospace') || w(`${q}, serif`) !== w('serif'));
  dxLocal.set(name, r); return r;
}
function dxBoxMet(fam) { // ascent + descent and descent (em) of a face the canvas has, for fonts not in the table
  if (dxBox.has(fam)) return dxBox.get(fam);
  DXC.font = '100px ' + fam; dxFont = ''; const m = DXC.measureText('Hg'); let r = [1.2, 0.22];
  if (m.fontBoundingBoxAscent > 0 && m.fontBoundingBoxDescent >= 0) r = [(m.fontBoundingBoxAscent + m.fontBoundingBoxDescent) / 100, m.fontBoundingBoxDescent / 100];
  dxBox.set(fam, r); return r;
}
const dxW_ = (text, st) => dxW100(text, st.css) * st.msize / 100 * st.scale + (st.sp ? [...text].length * st.sp : 0); // Word's width of a piece (no kerning, no ligatures)
function dxRgbHsl(r, g, b) { const mx = Math.max(r, g, b), mn = Math.min(r, g, b), l = (mx + mn) / 2; let h = 0, s = 0; if (mx !== mn) { const d = mx - mn; s = l > 0.5 ? d / (2 - mx - mn) : d / (mx + mn); h = mx === r ? (g - b) / d + (g < b ? 6 : 0) : mx === g ? (b - r) / d + 2 : (r - g) / d + 4; h /= 6; } return [h, s, l]; }
function dxHslRgb(h, s, l) { if (!s) return [l, l, l]; const f = (p, q, t) => { if (t < 0) t += 1; if (t > 1) t -= 1; return t < 1 / 6 ? p + (q - p) * 6 * t : t < 1 / 2 ? q : t < 2 / 3 ? p + (q - p) * (2 / 3 - t) * 6 : p; }; const q = l < 0.5 ? l * (1 + s) : l + s - l * s, p = 2 * l - q; return [f(p, q, h + 1 / 3), f(p, q, h), f(p, q, h - 1 / 3)]; }
function dxColor(holder, theme) { // a DrawingML colour (inside solidFill, ln, fillRef …) → { c: '#rrggbb', a: alpha } or null
  if (!holder) return null; const c = [...holder.children].find(k => /^(srgbClr|schemeClr|sysClr|prstClr|scrgbClr)$/.test(k.localName)); if (!c) return null;
  const v = c.getAttribute('val') || '', map = { tx1: 'dk1', bg1: 'lt1', tx2: 'dk2', bg2: 'lt2', phClr: 'accent1' };
  let hex = c.localName === 'srgbClr' ? v : c.localName === 'sysClr' ? c.getAttribute('lastClr') || (v === 'window' ? 'ffffff' : '000000') : c.localName === 'schemeClr' ? theme.clr[map[v] || v] : { black: '000000', white: 'ffffff', red: 'ff0000', blue: '0000ff', green: '008000', yellow: 'ffff00', gray: '808080' }[v];
  if (!hex || !/^[0-9a-f]{6}$/i.test(hex)) hex = '000000';
  let rgb = [0, 2, 4].map(i => parseInt(hex.slice(i, i + 2), 16) / 255), a = 1;
  for (const m of c.children) {
    const k = +(m.getAttribute('val') || 0) / 100000;
    if (m.localName === 'alpha') a = k;
    else if (m.localName === 'lumMod' || m.localName === 'lumOff') { const hsl = dxRgbHsl(...rgb); hsl[2] = m.localName === 'lumMod' ? hsl[2] * k : hsl[2] + k; hsl[2] = clamp(hsl[2], 0, 1); rgb = dxHslRgb(...hsl); }
    else if (m.localName === 'tint') rgb = rgb.map(x => x * k + (1 - k));
    else if (m.localName === 'shade') rgb = rgb.map(x => x * k);
  }
  return { c: '#' + rgb.map(x => Math.round(clamp(x, 0, 1) * 255).toString(16).padStart(2, '0')).join(''), a };
}
const DX_HL = { yellow: '#ffff00', green: '#00ff00', cyan: '#00ffff', magenta: '#ff00ff', blue: '#0000ff', red: '#ff0000', darkBlue: '#000080', darkCyan: '#008080', darkGreen: '#008000', darkMagenta: '#800080', darkRed: '#800000', darkYellow: '#808000', darkGray: '#808080', lightGray: '#c0c0c0', black: '#000000', white: '#ffffff' };
function parseRPr(el, theme) {
  const o = {}; if (!el) return o; let e;
  for (const n of ['b', 'i', 'strike', 'dstrike', 'caps', 'smallCaps', 'vanish', 'webHidden']) if ((e = kid(el, n))) o[n] = onoff(e);
  if ((e = kid(el, 'u'))) o.u = (av(e, 'val') || 'single') !== 'none';
  if ((e = kid(el, 'sz')) && anum(e, 'val')) o.sz = anum(e, 'val') / 2;
  if ((e = kid(el, 'color'))) { const v = av(e, 'val'); if (dxHex(v)) o.color = dxHex(v); else if (v === 'auto') o.color = '#000000'; }
  if ((e = kid(el, 'rFonts'))) {
    const th = n => { const t = av(e, n); return t && theme ? theme.font(t) : null; };
    const f = av(e, 'ascii') || th('asciiTheme') || av(e, 'hAnsi') || th('hAnsiTheme') || av(e, 'cs') || av(e, 'eastAsia'); if (f) o.font = f;
  }
  if ((e = kid(el, 'spacing')) && anum(e, 'val') != null) o.sp = anum(e, 'val') / 20;
  if ((e = kid(el, 'position')) && anum(e, 'val') != null) o.pos = anum(e, 'val') / 2;
  if ((e = kid(el, 'vertAlign'))) o.vert = av(e, 'val');
  if ((e = kid(el, 'w')) && anum(e, 'val')) o.scale = anum(e, 'val') / 100;
  if ((e = kid(el, 'highlight')) && DX_HL[av(e, 'val')]) o.hl = DX_HL[av(e, 'val')];
  else if ((e = kid(el, 'shd')) && dxHex(av(e, 'fill')) && av(e, 'fill').toLowerCase() !== 'ffffff') o.hl = dxHex(av(e, 'fill'));
  if ((e = kid(el, 'rStyle'))) o.rStyle = av(e, 'val');
  return o;
}
function dxBorder(e) { const v = e && av(e, 'val'); if (!e || !v || v === 'nil' || v === 'none') return null; const sz = anum(e, 'sz') || 4, c = dxHex(av(e, 'color')); return { w: Math.max(0.25, (v === 'double' ? 3 : /thick/.test(v) ? 2 : 1) * sz / 8), space: anum(e, 'space') || 0, color: c || '#000000', dbl: v === 'double' }; }
function dxTabs(el) { return kids(el, 'tab').map(t => ({ pos: (anum(t, 'pos') || 0) / 20, val: av(t, 'val') || 'left', leader: av(t, 'leader') || 'none' })); }
function dxMergeTabs(a, b) { // a style's tab stops, then the paragraph's own (a "clear" takes one away)
  if (!b || !b.length) return a || []; const m = new Map((a || []).map(t => [Math.round(t.pos * 20), t]));
  for (const t of b) { const k = Math.round(t.pos * 20); if (t.val === 'clear') m.delete(k); else m.set(k, t); }
  return [...m.values()].sort((x, y) => x.pos - y.pos);
}
function parsePPr(el) {
  const o = {}; if (!el) return o; let e;
  if ((e = kid(el, 'pStyle'))) o.style = av(e, 'val');
  if ((e = kid(el, 'jc'))) o.align = { center: 'center', right: 'right', end: 'right', both: 'both', distribute: 'distribute', lowKashida: 'both', mediumKashida: 'both', highKashida: 'both', thaiDistribute: 'both' }[av(e, 'val')] || 'left';
  if ((e = kid(el, 'spacing'))) {
    if (anum(e, 'before') != null) o.before = anum(e, 'before') / 20; if (anum(e, 'after') != null) o.after = anum(e, 'after') / 20;
    if (av(e, 'beforeAutospacing') != null) o.beforeAuto = onoffAttr(e, 'beforeAutospacing'); if (av(e, 'afterAutospacing') != null) o.afterAuto = onoffAttr(e, 'afterAutospacing');
    if (anum(e, 'beforeLines') != null) o.beforeLines = anum(e, 'beforeLines') / 100; if (anum(e, 'afterLines') != null) o.afterLines = anum(e, 'afterLines') / 100;
    if (anum(e, 'line')) { const rule = av(e, 'lineRule'); if (!rule || rule === 'auto') { o.lineMult = anum(e, 'line') / 240; o.lineExact = 0; o.lineRule = 'auto'; } else { o.lineExact = anum(e, 'line') / 20; o.lineRule = rule; } }
  }
  if ((e = kid(el, 'ind'))) {
    const l = anum(e, 'left') ?? anum(e, 'start'), r = anum(e, 'right') ?? anum(e, 'end');
    if (l != null) o.indL = l / 20; if (r != null) o.indR = r / 20;
    if (anum(e, 'hanging') != null) { o.hang = anum(e, 'hanging') / 20; o.first = 0; } else if (anum(e, 'firstLine') != null) { o.first = anum(e, 'firstLine') / 20; o.hang = 0; }
  }
  if ((e = kid(el, 'tabs'))) o.tabs = dxTabs(e);
  for (const n of ['pageBreakBefore', 'keepNext', 'keepLines', 'widowControl', 'contextualSpacing', 'snapToGrid']) if ((e = kid(el, n))) o[n] = onoff(e);
  if ((e = kid(el, 'numPr'))) o.num = { id: av(kid(e, 'numId'), 'val'), lvl: av(kid(e, 'ilvl'), 'val') || '0' };
  if ((e = kid(el, 'pBdr'))) { const B = {}; for (const s of ['top', 'bottom', 'left', 'right', 'between']) { const b = dxBorder(kid(e, s)); if (b) B[s] = b; } o.bdr = B; }
  if ((e = kid(el, 'shd'))) { const f = dxHex(av(e, 'fill')); o.shd = f && f !== '#ffffff' ? f : null; }
  return o;
}
function onoffAttr(e, n) { const v = av(e, n); return !(v === '0' || v === 'false' || v === 'off'); }
const fmtNum = (n, fmt) => {
  if (fmt === 'lowerLetter' || fmt === 'upperLetter') { const s = String.fromCharCode(97 + (n - 1) % 26).repeat(Math.floor((n - 1) / 26) + 1); return fmt === 'upperLetter' ? s.toUpperCase() : s; }
  if (fmt === 'lowerRoman' || fmt === 'upperRoman') { let s = '', v = n; for (const [a, r] of [[1000, 'm'], [900, 'cm'], [500, 'd'], [400, 'cd'], [100, 'c'], [90, 'xc'], [50, 'l'], [40, 'xl'], [10, 'x'], [9, 'ix'], [5, 'v'], [4, 'iv'], [1, 'i']]) while (v >= a) { s += r; v -= a; } return fmt === 'upperRoman' ? s.toUpperCase() : s; }
  if (fmt === 'decimalZero') return String(n).padStart(2, '0');
  return String(n);
};
const dxSym = (ch, font) => { // a symbol-font character (Symbol, Wingdings) → the Unicode look-alike
  const c = ch.charCodeAt(0), k = c >= 0xf000 ? c - 0xf000 : c, f = String(font || '').toLowerCase();
  if (/wingdings/.test(f)) return { 0xa7: '▪', 0xa8: '◻', 0xfc: '✓', 0xfb: '✗', 0xd8: '➢', 0x76: '❖', 0x6c: '●', 0x6e: '■', 0x71: '❑', 0x77: '◆', 0xe8: '➔', 0xa1: '○', 0x9f: '•' }[k] || '•';
  if (/symbol/.test(f)) return { 0xb7: '•', 0xa7: '♣', 0xae: '→', 0xb0: '°', 0xb1: '±', 0xd7: '×', 0xb4: '×', 0x2d: '−', 0xa5: '∞' }[k] || (k >= 0x20 && k < 0x7f ? String.fromCharCode(k) : '•');
  return c >= 0xf000 ? '•' : ch;
};
function dxDeobf(u8, key) { // an embedded .odttf → the font file (its first 32 bytes are XORed with the GUID, read backwards)
  const hex = String(key || '').replace(/[^0-9a-f]/gi, ''); if (hex.length !== 32) return u8;
  const k = []; for (let i = 0; i < 16; i++) k.push(parseInt(hex.substr(30 - 2 * i, 2), 16));
  const out = new Uint8Array(u8); for (let i = 0; i < 32 && i < out.length; i++) out[i] ^= k[i % 16]; return out;
}

async function docxToPages(file) {
  await X.loadLib('jszip');
  let zip; try { zip = await JSZip.loadAsync(await file.arrayBuffer()); } catch (e) { throw new Error('Could not read this Word file — is it a valid .docx?'); }
  const xml = async n => { const f = zip.file(n); if (!f) return null; const d = new DOMParser().parseFromString(await f.async('string'), 'application/xml'); return d.getElementsByTagName('parsererror').length ? null : d; };
  const doc = await xml('word/document.xml'); if (!doc || !doc.documentElement) throw new Error('This is not a valid .docx file');
  const body = kid(doc.documentElement, 'body'); if (!body) throw new Error('This Word file has no document body');
  const warn = [];

  // ── theme (fonts and colours), settings
  const theme = { clr: {}, major: 'Calibri Light', minor: 'Calibri', font(t) { return /^major/.test(t) ? this.major : /^minor/.test(t) ? this.minor : null; } };
  const thDoc = await xml('word/theme/theme1.xml');
  if (thDoc) {
    const cs = dxDeep1(thDoc, 'clrScheme'); if (cs) for (const c of cs.children) { const v = dxDeep1(c, 'srgbClr'), s = dxDeep1(c, 'sysClr'); theme.clr[c.localName] = v ? v.getAttribute('val') : s ? s.getAttribute('lastClr') : null; }
    const lat = n => { const f = dxDeep1(thDoc, n), l = f && kid(f, 'latin'); return l && l.getAttribute('typeface'); };
    theme.major = lat('majorFont') || theme.major; theme.minor = lat('minorFont') || theme.minor;
  }
  const setDoc = await xml('word/settings.xml'), setRoot = setDoc && setDoc.documentElement;
  const defTab = Math.max(6, (anum(kid(setRoot, 'defaultTabStop'), 'val') || 720) / 20), evenOdd = !!(kid(setRoot, 'evenAndOddHeaders') && onoff(kid(setRoot, 'evenAndOddHeaders')));
  const noExpandBr = !!dxDeep(setRoot, 'doNotExpandShiftReturn').length;

  // ── styles
  const stylesDoc = await xml('word/styles.xml'), raw = {}, defaults = { r: {}, p: {} }, tblStyles = {}, cache = {}; let defPara = null, defTbl = null;
  const tblProps = tp => { const o = {}; if (!tp) return o; const bd = kid(tp, 'tblBorders'); if (bd) { o.bd = {}; for (const s of ['top', 'left', 'start', 'bottom', 'right', 'end', 'insideH', 'insideV']) { const e = kid(bd, s); if (e) o.bd[s === 'start' ? 'left' : s === 'end' ? 'right' : s] = dxBorder(e); } }
    const cm = kid(tp, 'tblCellMar'); if (cm) { o.mar = {}; for (const s of ['top', 'left', 'start', 'bottom', 'right', 'end']) { const e = kid(cm, s); if (e && anum(e, 'w') != null) o.mar[s === 'start' ? 'left' : s === 'end' ? 'right' : s] = anum(e, 'w') / 20; } }
    const ind = kid(tp, 'tblInd'); if (ind && anum(ind, 'w') != null) o.ind = anum(ind, 'w') / 20;
    const jc = kid(tp, 'jc'); if (jc) o.jc = av(jc, 'val'); const lay = kid(tp, 'tblLayout'); if (lay) o.fixed = av(lay, 'type') === 'fixed';
    const w = kid(tp, 'tblW'); if (w) o.w = { v: anum(w, 'w') || 0, type: av(w, 'type') || 'dxa' };
    return o; };
  if (stylesDoc) {
    const root = stylesDoc.documentElement, dd = kid(root, 'docDefaults');
    defaults.r = parseRPr(kid(kid(dd, 'rPrDefault'), 'rPr'), theme); defaults.p = parsePPr(kid(kid(dd, 'pPrDefault'), 'pPr'));
    for (const st of kids(root, 'style')) {
      const id = av(st, 'styleId'), type = av(st, 'type');
      raw[id] = { based: av(kid(st, 'basedOn'), 'val'), r: parseRPr(kid(st, 'rPr'), theme), p: parsePPr(kid(st, 'pPr')), t: tblProps(kid(st, 'tblPr')) }; delete raw[id].p.style;
      if (type === 'paragraph' && av(st, 'default') === '1') defPara = id;
      if (type === 'table' && av(st, 'default') === '1') defTbl = id;
      if (type === 'table') tblStyles[id] = true;
    }
  }
  const style = (id, depth = 0) => {
    if (!id || !raw[id] || depth > 12) return { r: {}, p: {}, t: {} }; if (cache[id]) return cache[id]; const b = style(raw[id].based, depth + 1), s = raw[id];
    const t = { ...b.t, ...s.t, bd: { ...(b.t.bd || {}), ...(s.t.bd || {}) }, mar: { ...(b.t.mar || {}), ...(s.t.mar || {}) } };
    return cache[id] = { r: { ...b.r, ...s.r }, p: { ...b.p, ...s.p, tabs: dxMergeTabs(b.p.tabs, s.p.tabs) }, t };
  };

  // ── fonts: embedded ones (fontTable) are measured and drawn with their own faces
  const embedded = {}; // Word name → { measure family, draw id per [bold][italic] }
  const ftDoc = await xml('word/fontTable.xml');
  if (ftDoc) {
    const rels = await relsOf('word/fontTable.xml');
    for (const f of kids(ftDoc.documentElement, 'font')) {
      const name = av(f, 'name'); if (!name || X.BUNDLED[famOf(name).fam] && famOf(name).fam.toLowerCase() === name.toLowerCase()) continue; // (a bundled family is used as it is)
      for (const [tagN, bold, italic] of [['embedRegular', 0, 0], ['embedBold', 1, 0], ['embedItalic', 0, 1], ['embedBoldItalic', 1, 1]]) {
        const e = kid(f, tagN), r = e && rels[dxRid(e, 'id')], zf = r && zip.file(r.target); if (!zf) continue;
        try {
          const bytes = dxDeobf(await zf.async('uint8array'), av(e, 'fontKey'));
          const ff = new FontFace('XD3W ' + name, bytes, { weight: bold ? '700' : '400', style: italic ? 'italic' : 'normal' }); document.fonts.add(ff); await ff.load();
          const id = 'dx' + (name + bold + italic + bytes.length).split('').reduce((h, c) => Math.imul(h ^ c.charCodeAt(0), 16777619), 2166136261 >>> 0).toString(36).replace('-', 'n');
          S.xfonts[id] = { name, ps: name.replace(/\s+/g, ''), bytes, bold: !!bold, italic: !!italic, fb: famOf(name).fam };
          const E = embedded[name] || (embedded[name] = { css: `"XD3W ${name.replace(/["\\]/g, '')}"`, ids: {} }); E.ids[bold * 2 + italic] = id;
        } catch (err) { console.warn('embedded font skipped', name, err); }
      }
    }
  }
  const stCache = new Map(), usedFams = new Set(), allSt = [];
  const BASE_RUN = { sz: 10, font: theme.minor || 'Calibri', color: '#000000' };
  function mkSt(o) { // merged run properties → the style object the layout uses (shared, cached)
    const key = JSON.stringify(o); let st = stCache.get(key); if (st) return st;
    const wname = o.font || 'Calibri', ff = famOf(wname), E = embedded[wname], bold = !!o.b || ff.heavy, italic = !!o.i;
    const sup = o.vert === 'superscript' || o.vert === 'subscript', size = o.sz || 10, msize = (sup ? size * 0.65 : size) * (o.smallCaps ? 0.8 : 1);
    let fam = ff.fam; if (E) { const id = E.ids[(bold ? 2 : 0) + (italic ? 1 : 0)] || E.ids[0] || E.ids[2] || E.ids[1]; if (id) fam = id; }
    let mf; const nm = wname.replace(/["\\]/g, ''), base = wname.replace(DX_WEIGHT, '');
    if (E) mf = `${E.css}, ${X.FONTS[ff.fam] || 'sans-serif'}`;
    else if (dxInstalled(wname)) mf = `"${nm}", ${X.FONTS[ff.fam] || 'sans-serif'}`;
    else if (base !== wname && dxInstalled(base)) mf = `"${base.replace(/["\\]/g, '')}", ${X.FONTS[ff.fam] || 'sans-serif'}`;
    else mf = X.FONTS[ff.fam] || 'Helvetica, Arial, sans-serif';
    const wt = /semibold|semi bold|demi/i.test(wname) && !E ? '600' : bold ? 'bold' : /light/i.test(wname) && !E && !dxInstalled(wname) ? '300' : 'normal';
    const lk = wname.toLowerCase();
    st = { size, msize, wname, fam, bold, italic, color: o.color || '#000000', underline: !!o.u, strike: !!(o.strike || o.dstrike), caps: !!(o.caps || o.smallCaps), hidden: !!(o.vanish || o.webHidden),
      sp: o.sp || 0, rise: (o.pos || 0) + (o.vert === 'superscript' ? size * 0.35 : o.vert === 'subscript' ? -size * 0.14 : 0), scale: o.scale || 1, hl: o.hl || null, mf, wt,
      css: `${italic ? 'italic ' : ''}${wt} 100px ${mf}`, met: DX_MET[lk] || DX_MET[lk.replace(DX_WEIGHT, '')] || DX_MET[String(fam).toLowerCase()] || null };
    if (!st.met) st.met = dxBoxMet(mf);
    st.key = key; stCache.set(key, st); allSt.push(st); usedFams.add(fam); return st;
  }

  // ── numbering (bullets / numbered lists)
  const numDoc = await xml('word/numbering.xml'), nums = {}, abs = {}, counters = {}, overrides = {}, seenNum = new Set();
  if (numDoc) {
    const root = numDoc.documentElement;
    for (const a of kids(root, 'abstractNum')) { const lv = abs[av(a, 'abstractNumId')] = {}; for (const l of kids(a, 'lvl')) { const pp = parsePPr(kid(l, 'pPr')); lv[av(l, 'ilvl')] = { fmt: av(kid(l, 'numFmt'), 'val') || 'decimal', text: av(kid(l, 'lvlText'), 'val') ?? '', start: +(av(kid(l, 'start'), 'val') || 1), p: pp, r: parseRPr(kid(l, 'rPr'), theme), suff: av(kid(l, 'suff'), 'val') || 'tab', jc: av(kid(l, 'lvlJc'), 'val') || 'left' }; } }
    for (const n of kids(root, 'num')) { const id = av(n, 'numId'); nums[id] = av(kid(n, 'abstractNumId'), 'val'); for (const o of kids(n, 'lvlOverride')) { const so = kid(o, 'startOverride'); if (so) (overrides[id] = overrides[id] || {})[av(o, 'ilvl') || '0'] = +av(so, 'val'); } }
  }
  const label = num => {
    const lv = +num.lvl || 0, aid = nums[num.id], a = abs[aid]; if (!a || !a[lv]) return null;
    const c = counters[aid] || (counters[aid] = []);
    if (!seenNum.has(num.id)) { seenNum.add(num.id); const ov = overrides[num.id]; if (ov) for (const k in ov) { c[+k] = ov[k] - 1; c.length = Math.max(c.length, +k + 1); } }
    c[lv] = (c[lv] === undefined ? a[lv].start - 1 : c[lv]) + 1; c.length = lv + 1;
    const d = a[lv], text = d.fmt === 'bullet' ? (d.text ? [...d.text].map(ch => dxSym(ch, d.r.font)).join('') : '•') : d.fmt === 'none' ? '' : d.text.replace(/%(\d)/g, (m, k) => fmtNum(c[k - 1] === undefined ? (a[k - 1] ? a[k - 1].start : 1) : c[k - 1], (a[k - 1] || d).fmt));
    return { d, text };
  };

  // ── parts and their pictures
  async function relsOf(part) {
    const dir = part.slice(0, part.lastIndexOf('/') + 1), d = await xml(dir + '_rels/' + part.slice(dir.length) + '.rels'), map = {};
    if (d) for (const r of [...d.documentElement.children]) {
      const t = r.getAttribute('Target') || '', ext = r.getAttribute('TargetMode') === 'External';
      let p = t; if (!ext) { if (t.startsWith('/')) p = t.slice(1); else { const segs = (dir + t).split('/'), out = []; for (const s of segs) { if (s === '..') out.pop(); else if (s !== '.') out.push(s); } p = out.join('/'); } }
      map[r.getAttribute('Id')] = { type: (r.getAttribute('Type') || '').split('/').pop(), target: p, ext };
    }
    return map;
  }
  async function partCtx(part) {
    const rels = await relsOf(part), imgs = {};
    for (const id in rels) {
      const r = rels[id]; if (r.type !== 'image' || r.ext) continue;
      const f = zip.file(r.target), ext = (r.target.split('.').pop() || '').toLowerCase();
      const mime = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', jpe: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', webp: 'image/webp', svg: 'image/svg+xml' }[ext]; if (!f || !mime) continue;
      try { const im = await X.readImage(new File([await f.async('uint8array')], r.target.split('/').pop(), { type: mime })); imgs[id] = { asset: X.addAsset(im.url), pw: im.w, ph: im.h }; } catch (e) { /* unsupported picture: its room is kept, nothing is drawn */ }
    }
    return { part, rels, imgs };
  }

  // ── drawings: pictures, shapes, groups, text boxes
  function spOf(spPr, styleEl, C) { // shape properties → { geom, fill, line, rot, flipH, flipV, img }
    const o = { geom: 'rect', fill: null, line: null, rot: 0, flipH: false, flipV: false, adj: null };
    const xf = kid(spPr, 'xfrm'); if (xf) { o.rot = (+xf.getAttribute('rot') || 0) / 60000; o.flipH = xf.getAttribute('flipH') === '1'; o.flipV = xf.getAttribute('flipV') === '1'; }
    const pg = kid(spPr, 'prstGeom'); if (pg) { o.geom = pg.getAttribute('prst') || 'rect'; const gd = dxDeep1(pg, 'gd'); if (gd) { const m = /val\s+(-?\d+)/.exec(gd.getAttribute('fmla') || ''); if (m) o.adj = +m[1] / 100000; } }
    else if (kid(spPr, 'custGeom')) o.geom = 'rect';
    const sf = kid(spPr, 'solidFill'), gf = kid(spPr, 'gradFill'), bf = kid(spPr, 'blipFill');
    if (sf) o.fill = dxColor(sf, theme); else if (gf) o.fill = dxColor(dxDeep1(gf, 'gs'), theme);
    else if (bf) { const b = dxDeep1(bf, 'blip'), im = b && C.imgs[dxRid(b, 'embed')]; if (im) o.img = im.asset; }
    else if (!kid(spPr, 'noFill') && styleEl) { const fr = kid(styleEl, 'fillRef'); if (fr && +fr.getAttribute('idx') > 0) o.fill = dxColor(fr, theme); }
    const ln = kid(spPr, 'ln');
    if (ln) { if (!kid(ln, 'noFill')) { const c = dxColor(kid(ln, 'solidFill'), theme) || (styleEl && dxColor(kid(styleEl, 'lnRef'), theme)); if (c) o.line = { c: c.c, a: c.a, w: (+ln.getAttribute('w') || 9525) / DX_EMU }; } }
    else if (styleEl) { const lr = kid(styleEl, 'lnRef'); if (lr && +lr.getAttribute('idx') > 0) { const c = dxColor(lr, theme); if (c) o.line = { c: c.c, a: c.a, w: 0.75 }; } }
    return o;
  }
  function graphicOf(gd, C, w, h) { // a:graphicData → shapes in the drawing's own box [{ x, y, w, h, kind, … }]
    if (!gd) return []; const out = [];
    const one = (el, x, y, cw, ch) => {
      const n = el.localName;
      if (n === 'pic') { const b = dxDeep1(el, 'blip'), im = b && C.imgs[dxRid(b, 'embed')], sp = spOf(kid(el, 'spPr') || el, null, C); out.push({ kind: 'pic', x, y, w: cw, h: ch, asset: im ? im.asset : null, rot: sp.rot, flipH: sp.flipH, flipV: sp.flipV, line: sp.line }); }
      else if (n === 'wsp') {
        const sp = spOf(kid(el, 'spPr'), kid(el, 'style'), C), tx = kid(el, 'txbx'), bp = kid(el, 'bodyPr'), s = { kind: 'shape', x, y, w: cw, h: ch, ...sp };
        if (tx) { const tc = kid(tx, 'txbxContent'); if (tc) s.txbx = { blocks: blocksOf(tc, C), ins: ['lIns', 'tIns', 'rIns', 'bIns'].map((k, i) => bp && bp.getAttribute(k) != null ? +bp.getAttribute(k) / DX_EMU : i % 2 ? 3.6 : 7.2), vert: bp && bp.getAttribute('vert') || 'horz', anchor: bp && bp.getAttribute('anchor') || 't', auto: !!(bp && kid(bp, 'spAutoFit')) }; }
        out.push(s);
      } else if (n === 'wgp' || n === 'grpSp') {
        const gx = kid(kid(el, 'grpSpPr'), 'xfrm'), ch0 = gx && kid(gx, 'chOff'), che = gx && kid(gx, 'chExt');
        const ox = ch0 ? +ch0.getAttribute('x') : 0, oy = ch0 ? +ch0.getAttribute('y') : 0, ex = che ? +che.getAttribute('cx') || 1 : cw * DX_EMU, ey = che ? +che.getAttribute('cy') || 1 : ch * DX_EMU, kx = cw / ex, ky = ch / ey;
        for (const c of el.children) {
          if (!/^(wsp|pic|grpSp|wgp)$/.test(c.localName)) continue;
          const xf = c.localName === 'grpSp' ? kid(kid(c, 'grpSpPr'), 'xfrm') : kid(kid(c, 'spPr'), 'xfrm'), off = xf && kid(xf, 'off'), ext = xf && kid(xf, 'ext'); if (!off || !ext) continue;
          one(c, x + (+off.getAttribute('x') - ox) * kx, y + (+off.getAttribute('y') - oy) * ky, +ext.getAttribute('cx') * kx, +ext.getAttribute('cy') * ky);
        }
      }
    };
    for (const c of gd.children) one(c, 0, 0, w, h);
    return out;
  }
  function drawingOf(dr, C) { // w:drawing → { inline, w, h, shapes, pos… }
    const a = kid(dr, 'anchor'), box = a || kid(dr, 'inline'); if (!box) return null;
    const ext = kid(box, 'extent'), w = ext ? (+ext.getAttribute('cx') || 0) / DX_EMU : 0, h = ext ? (+ext.getAttribute('cy') || 0) / DX_EMU : 0;
    const d = { inline: !a, w, h, shapes: graphicOf(dxDeep1(box, 'graphicData'), C, w, h) };
    if (a) {
      const pos = el => { if (!el) return null; const o = el.getAttribute('relativeFrom') || 'column', p = kid(el, 'posOffset'), al = kid(el, 'align'), pc = kid(el, 'pctPosHOffset') || kid(el, 'pctPosVOffset'); return { rel: o, off: p ? +p.textContent / DX_EMU : 0, align: al ? al.textContent.trim() : null, pct: pc ? +pc.textContent / 100000 : null }; };
      d.H = pos(kid(a, 'positionH')) || { rel: 'column', off: 0 }; d.V = pos(kid(a, 'positionV')) || { rel: 'paragraph', off: 0 };
      if (a.getAttribute('simplePos') === '1') { const sp = kid(a, 'simplePos'); d.H = { rel: 'page', off: (+sp.getAttribute('x') || 0) / DX_EMU }; d.V = { rel: 'page', off: (+sp.getAttribute('y') || 0) / DX_EMU }; }
      d.behind = a.getAttribute('behindDoc') === '1'; d.z = +a.getAttribute('relativeHeight') || 0; d.inCell = a.getAttribute('layoutInCell') !== '0';
      d.wrap = ['wrapSquare', 'wrapTight', 'wrapThrough', 'wrapTopAndBottom'].map(n => kid(a, n) && n).find(Boolean) || 'none';
      d.side = (kid(a, 'wrapSquare') || kid(a, 'wrapTight') || kid(a, 'wrapThrough') || a).getAttribute('wrapText') || 'bothSides';
      d.dist = ['distT', 'distB', 'distL', 'distR'].map(k => (+a.getAttribute(k) || 0) / DX_EMU);
      const ee = kid(a, 'effectExtent'); d.eff = ee ? ['l', 't', 'r', 'b'].map(k => (+ee.getAttribute(k) || 0) / DX_EMU) : [0, 0, 0, 0];
    }
    return d;
  }
  function vmlOf(pict, C) { // an old VML picture / text box (w:pict) → a drawing
    const sh = [...pict.children].find(c => c.localName === 'shape' || c.localName === 'rect' || c.localName === 'roundrect'); if (!sh) return null;
    const css = {}; (sh.getAttribute('style') || '').split(';').forEach(p => { const i = p.indexOf(':'); if (i > 0) css[p.slice(0, i).trim()] = p.slice(i + 1).trim(); });
    const pt = v => { const m = /^(-?[\d.]+)(pt|in|cm|mm|px)?$/.exec(v || ''); if (!m) return 0; return +m[1] * ({ pt: 1, in: 72, cm: 72 / 2.54, mm: 72 / 25.4, px: 0.75 }[m[2] || 'px'] || 1); };
    const w = pt(css.width), h = pt(css.height), im = dxDeep1(sh, 'imagedata'), tb = dxDeep1(sh, 'txbxContent'), shapes = [];
    if (im) { const r = C.imgs[dxRid(im, 'id')]; shapes.push({ kind: 'pic', x: 0, y: 0, w, h, asset: r ? r.asset : null }); }
    else { const fc = sh.getAttribute('fillcolor'), sc = sh.getAttribute('strokecolor'), fill = /^#?[0-9a-f]{6}$/i.test(fc || '') ? { c: '#' + fc.replace('#', '').toLowerCase(), a: 1 } : sh.getAttribute('filled') === 'f' ? null : { c: '#ffffff', a: 1 };
      shapes.push({ kind: 'shape', x: 0, y: 0, w, h, geom: sh.localName === 'roundrect' ? 'roundRect' : 'rect', fill, line: sh.getAttribute('stroked') === 'f' ? null : { c: /^#?[0-9a-f]{6}$/i.test(sc || '') ? '#' + sc.replace('#', '').toLowerCase() : '#000000', a: 1, w: pt(sh.getAttribute('strokeweight')) || 0.75 }, rot: 0, txbx: tb ? { blocks: blocksOf(tb, C), ins: [7.2, 3.6, 7.2, 3.6], vert: 'horz', anchor: 't' } : null }); }
    const d = { inline: css.position !== 'absolute', w, h, shapes };
    if (!d.inline) { d.H = { rel: { page: 'page', margin: 'margin', text: 'column' }[css['mso-position-horizontal-relative']] || 'column', off: pt(css['margin-left']) + pt(css.left) }; d.V = { rel: { page: 'page', margin: 'margin', text: 'paragraph', line: 'line' }[css['mso-position-vertical-relative']] || 'paragraph', off: pt(css['margin-top']) + pt(css.top) }; d.behind = /^-/.test(css['z-index'] || ''); d.wrap = dxDeep1(sh, 'wrap') ? ({ square: 'wrapSquare', tight: 'wrapTight', topAndBottom: 'wrapTopAndBottom' }[dxDeep1(sh, 'wrap').getAttribute('type')] || 'none') : 'none'; d.side = 'bothSides'; d.dist = [0, 0, 9, 9]; d.eff = [0, 0, 0, 0]; d.inCell = true; }
    return d;
  }

  // ── document → blocks
  function parsePara(pEl, C, tbl) {
    const pPr = kid(pEl, 'pPr'), direct = parsePPr(pPr), sid = direct.style || defPara, stl = style(sid), tp = tbl ? style(tbl).p : {};
    const P = { ...defaults.p, ...tp, ...stl.p, ...direct, tabs: dxMergeTabs(dxMergeTabs(dxMergeTabs(defaults.p.tabs, tp.tabs), stl.p.tabs), direct.tabs) };
    if (tbl && !direct.style && sid === defPara) for (const k of ['after', 'before', 'lineMult', 'lineExact', 'lineRule']) if (tp[k] !== undefined && direct[k] === undefined) P[k] = tp[k];
    const Rb = { ...BASE_RUN, ...defaults.r, ...(tbl ? style(tbl).r : {}), ...stl.r }; delete Rb.rStyle;
    const runs = [], anchors = []; let fld = [];
    const push = (r, inFld) => { if (fld.some(f => f.hide)) return; runs.push(r); };
    const walk = (el, href) => {
      for (const c of el.children) {
        const n = c.localName;
        if (n === 'r') {
          const rp = kid(c, 'rPr'), d = parseRPr(rp, theme), rs = d.rStyle ? style(d.rStyle).r : {}; delete d.rStyle; const o = { ...Rb, ...rs, ...d }; delete o.rStyle; const st = mkSt(o);
          const runKids = el2 => {
            for (const t of el2.children) {
              const tn = t.localName;
              if (tn === 'fldChar') { const ty = av(t, 'fldCharType'); if (ty === 'begin') fld.push({ instr: '', hide: true, sep: false }); else if (ty === 'separate' && fld.length) { const f = fld[fld.length - 1], k = /^\s*(PAGE|NUMPAGES|SECTIONPAGES)\b/i.exec(f.instr); f.sep = true; f.hide = !!k; if (k && !fld.slice(0, -1).some(g => g.hide)) runs.push({ fld: k[1].toUpperCase(), st }); } else if (ty === 'end' && fld.length) { const f = fld.pop(); const k = /^\s*(PAGE|NUMPAGES|SECTIONPAGES)\b/i.exec(f.instr); if (k && !f.sep && !fld.some(g => g.hide)) runs.push({ fld: k[1].toUpperCase(), st }); } continue; }
              if (tn === 'instrText') { if (fld.length) fld[fld.length - 1].instr += t.textContent; continue; }
              if (fld.some(f => f.hide)) continue;
              if (tn === 't') push({ text: st.caps ? t.textContent.toUpperCase() : t.textContent, st, href });
              else if (tn === 'tab' || tn === 'ptab') push({ tab: 1, st, href });
              else if (tn === 'br' || tn === 'cr') { const ty = av(t, 'type'); push({ br: ty === 'page' ? 'page' : ty === 'column' ? 'column' : 'line', st }); }
              else if (tn === 'noBreakHyphen') push({ text: '\u2011', st, href });
              else if (tn === 'sym') { const ch = String.fromCharCode(parseInt(av(t, 'char') || '2022', 16)); push({ text: dxSym(ch, av(t, 'font')), st, href }); }
              else if (tn === 'footnoteReference' || tn === 'endnoteReference') { const fs = mkSt({ ...o, vert: 'superscript' }); push({ text: String(av(t, 'id') || ''), st: fs, href }); }
              else if (tn === 'drawing') { const d2 = drawingOf(t, C); if (d2) { if (d2.inline) push({ img: d2, st }); else anchors.push(d2); } }
              else if (tn === 'pict' || tn === 'object') { const d2 = vmlOf(t, C); if (d2) { if (d2.inline) push({ img: d2, st }); else anchors.push(d2); } }
              else if (tn === 'AlternateContent') { const ch = kid(t, 'Choice'), fb = kid(t, 'Fallback'), n0 = runs.length + anchors.length; if (ch) runKids(ch); if (runs.length + anchors.length === n0 && fb) runKids(fb); }
            }
          };
          runKids(c);
        } else if (n === 'hyperlink') { const r = C.rels[dxRid(c, 'id')]; walk(c, r && r.ext ? r.target : href); }
        else if (n === 'fldSimple') { const k = /^\s*(PAGE|NUMPAGES|SECTIONPAGES)\b/i.exec(av(c, 'instr') || ''); if (k) { const r0 = c.getElementsByTagNameNS('*', 'r')[0], o = { ...Rb, ...parseRPr(kid(r0, 'rPr'), theme) }; delete o.rStyle; runs.push({ fld: k[1].toUpperCase(), st: mkSt(o) }); } else walk(c, href); }
        else if (/^(ins|smartTag|sdt|sdtContent|customXml|moveTo|dir|bdo)$/.test(n)) walk(c, href);
        else if (n === 'oMath' || n === 'oMathPara') { const t = dxDeep(c, 't').map(e => e.textContent).join(''); if (t) runs.push({ text: t, st: mkSt(Rb) }); }
      }
    };
    walk(pEl);
    const mo = { ...Rb, ...parseRPr(kid(pPr, 'rPr'), theme) }; delete mo.rStyle;
    const b = { t: 'p', P, sid, direct, runs, anchors, mark: mkSt(mo), sect: kid(pPr, 'sectPr') };
    if (P.num && P.num.id && P.num.id !== '0') {
      const lb = label(P.num);
      if (lb) {
        const lp = lb.d.p; if (direct.indL === undefined && lp.indL !== undefined) P.indL = lp.indL; if (direct.indR === undefined && lp.indR !== undefined) P.indR = lp.indR; if (direct.hang === undefined && direct.first === undefined) { if (lp.hang !== undefined) { P.hang = lp.hang; P.first = 0; } else if (lp.first !== undefined) { P.first = lp.first; P.hang = 0; } }
        if (lp.tabs) P.tabs = dxMergeTabs(lp.tabs, P.tabs);
        const lo = { ...mo, ...lb.d.r }; delete lo.rStyle; if (lb.d.r.font && /symbol|wingdings/i.test(lb.d.r.font)) lo.font = mo.font;
        if (lb.text) b.label = { text: lb.text, st: mkSt({ ...lo, u: false }), suff: lb.d.suff, jc: lb.d.jc };
      }
    }
    return b;
  }
  function parseTbl(tEl, C) {
    const tblPr = kid(tEl, 'tblPr'), ts = av(kid(tblPr, 'tblStyle'), 'val') || defTbl, TS = style(ts).t, own = tblProps(tblPr);
    const T = { t: 'tbl', sid: ts, bd: { ...(TS.bd || {}), ...(own.bd || {}) }, mar: { top: 0, bottom: 0, left: 5.4, right: 5.4, ...(TS.mar || {}), ...(own.mar || {}) }, ind: own.ind ?? TS.ind ?? 0, jc: own.jc || TS.jc || 'left', fixed: own.fixed ?? TS.fixed, w: own.w || TS.w,
      grid: kids(kid(tEl, 'tblGrid'), 'gridCol').map(g => (anum(g, 'w') || 0) / 20), rows: [] };
    for (const tr of kids(tEl, 'tr')) {
      const trPr = kid(tr, 'trPr'), th = kid(trPr, 'trHeight'), row = { cells: [], h: th ? (anum(th, 'val') || 0) / 20 : 0, hRule: th ? av(th, 'hRule') || 'atLeast' : 'atLeast', cant: !!(kid(trPr, 'cantSplit') && onoff(kid(trPr, 'cantSplit'))), head: !!(kid(trPr, 'tblHeader') && onoff(kid(trPr, 'tblHeader'))), before: anum(kid(trPr, 'gridBefore'), 'val') || 0 };
      const cellsOf = el => { for (const tc of el.children) { if (tc.localName === 'tc') {
        const pr = kid(tc, 'tcPr'), sh = dxHex(av(kid(pr, 'shd'), 'fill')), vm = kid(pr, 'vMerge'), tb = kid(pr, 'tcBorders'), tm = kid(pr, 'tcMar'), bd = {}, mar = {};
        if (tb) for (const s of ['top', 'left', 'start', 'bottom', 'right', 'end']) { const e = kid(tb, s); if (e) bd[s === 'start' ? 'left' : s === 'end' ? 'right' : s] = dxBorder(e) || 'none'; }
        if (tm) for (const s of ['top', 'left', 'start', 'bottom', 'right', 'end']) { const e = kid(tm, s); if (e && anum(e, 'w') != null) mar[s === 'start' ? 'left' : s === 'end' ? 'right' : s] = anum(e, 'w') / 20; }
        row.cells.push({ span: anum(kid(pr, 'gridSpan'), 'val') || 1, vm: vm ? (av(vm, 'val') === 'restart' ? 'start' : 'cont') : null, shade: sh && sh !== '#ffffff' && av(kid(pr, 'shd'), 'fill').toLowerCase() !== 'auto' ? sh : null, bd, mar, va: av(kid(pr, 'vAlign'), 'val') || 'top', blocks: blocksOf(tc, C, ts) });
      } else if (tc.localName === 'sdt' || tc.localName === 'customXml') cellsOf(kid(tc, 'sdtContent') || tc); } };
      cellsOf(tr); T.rows.push(row);
    }
    return T;
  }
  function blocksOf(el, C, tbl) { const out = []; for (const c of el.children) { const n = c.localName; if (n === 'p') out.push(parsePara(c, C, tbl)); else if (n === 'tbl') out.push(parseTbl(c, C)); else if (n === 'sdt') out.push(...blocksOf(kid(c, 'sdtContent') || c, C, tbl)); else if (n === 'customXml') out.push(...blocksOf(c, C, tbl)); } return out; }

  const DC = await partCtx('word/document.xml'), blocks = blocksOf(body, DC);
  // sections: each sectPr ends one; the body's own is the last
  const secProps = (sp, prev) => {
    const sz = kid(sp, 'pgSz'), mar = kid(sp, 'pgMar'), cols = kid(sp, 'cols'), pn = kid(sp, 'pgNumType'), dg = kid(sp, 'docGrid');
    const PW = (anum(sz, 'w') || 12240) / 20, PH = (anum(sz, 'h') || 15840) / 20, m = (k, d) => { const v = anum(mar, k); return v == null ? d : v / 20; };
    const o = { PW, PH, ML: Math.max(0, m('left', 72)), MR: Math.max(0, m('right', 72)), MT: m('top', 72), MB: m('bottom', 72), HD: m('header', 36), FD: m('footer', 36), type: av(kid(sp, 'type'), 'val') || 'nextPage', titlePg: !!(kid(sp, 'titlePg') && onoff(kid(sp, 'titlePg'))), start: pn && anum(pn, 'start') != null ? anum(pn, 'start') : null, refs: {} , grid: dg && /^lines/.test(av(dg, 'type') || '') ? (anum(dg, 'linePitch') || 0) / 20 : 0 };
    o.MT = Math.abs(o.MT); o.MB = Math.abs(o.MB);
    const n = Math.max(1, anum(cols, 'num') || 1), space = (anum(cols, 'space') ?? 720) / 20, list = kids(cols, 'col'), W = Math.max(36, PW - o.ML - o.MR);
    if (list.length && av(cols, 'equalWidth') === '0') { let x = 0; o.cols = list.map(c => { const w = (anum(c, 'w') || 0) / 20, r = { x, w }; x += w + (anum(c, 'space') || 0) / 20; return r; }); }
    else { const w = Math.max(18, (W - space * (n - 1)) / n); o.cols = Array.from({ length: n }, (_, i) => ({ x: i * (w + space), w })); }
    for (const kind of ['header', 'footer']) for (const r of kids(sp, kind + 'Reference')) o.refs[kind + ':' + (av(r, 'type') || 'default')] = dxRid(r, 'id');
    if (prev) for (const k in prev.refs) if (!(k in o.refs)) o.refs[k] = prev.refs[k]; // (a section without its own header uses the one before)
    return o;
  };
  const sections = []; let cur = [], prevP = null;
  for (const b of blocks) { cur.push(b); if (b.t === 'p' && b.sect) { prevP = secProps(b.sect, prevP); sections.push({ blocks: cur, props: prevP }); cur = []; } }
  prevP = secProps(kid(body, 'sectPr'), prevP); sections.push({ blocks: cur, props: prevP });
  // header / footer parts
  const hfParts = {};
  for (const s of sections) for (const k in s.props.refs) {
    const id = s.props.refs[k], r = DC.rels[id]; if (!r || hfParts[id] !== undefined) continue; hfParts[id] = null;
    const d = await xml(r.target); if (!d) continue; const C = await partCtx(r.target); hfParts[id] = blocksOf(d.documentElement, C);
  }

  // ── fonts must be loaded before anything is measured; a missing font file is said, and a standard PDF font stands in
  const bundledUsed = [...usedFams].filter(f => X.BUNDLED[f]), failed = new Set();
  await Promise.all(bundledUsed.map(f => Promise.resolve(X.ensureFont(f)).catch(() => {})));
  await Promise.all(Object.keys(S.xfonts).filter(id => usedFams.has(id)).map(id => Promise.resolve(X.ensureFont(id)).catch(() => {})));
  const faces = [...document.fonts];
  for (const f of bundledUsed) { const mine = faces.filter(ff => ff.family.replace(/["']/g, '') === 'XD3 ' + f); if (!mine.length || mine.some(ff => ff.status !== 'loaded')) failed.add(f); }
  if (failed.size) {
    for (const st of allSt) if (failed.has(st.fam)) { const g = X.BUNDLED[st.fam][2]; st.fam = g === 'serif' ? 'Times' : g === 'monospace' ? 'Courier' : 'Helvetica'; }
    warn.push(`Some fonts could not be loaded (${[...failed].join(', ')}): the font files in lib/fonts are missing or blocked. Standard PDF fonts were used instead, so their text may look and wrap a little differently.`);
  }
  dxFont = ''; dxW.clear();

  // ── line breaking (Word's rules)
  const tokCache = new Map();
  function toks(b) { // a paragraph's runs → words, spaces, tabs, breaks, pictures, fields
    let T = tokCache.get(b); if (T) return T; T = [];
    for (const r of b.runs) {
      if (r.img) { T.push({ k: 'img', d: r.img, w: r.img.w, h: r.img.h, st: r.st, bk: true }); continue; }
      if (r.tab) { T.push({ k: 'tab', st: r.st, href: r.href }); continue; }
      if (r.br) { T.push({ k: 'br', br: r.br, st: r.st }); continue; }
      if (r.fld) { const t = r.fld === 'PAGE' ? '8' : '88'; T.push({ k: 'fld', fld: r.fld, t, st: r.st, w: dxW_(t, r.st) }); continue; }
      if (r.st.hidden || !r.text) continue;
      for (const piece of r.text.replace(/[\u00ad\u200b]/g, '').split(/( +)/)) {
        if (!piece) continue;
        if (piece[0] === ' ') { T.push({ k: 'sp', t: piece, st: r.st, w: dxW_(piece, r.st), href: r.href }); continue; }
        const parts = /[\u2e80-\u9fff\uac00-\ud7af\uf900-\ufaff\uff00-\uffef]/.test(piece) ? [...piece] : piece.split(/(?<=[-\u2013\u2014](?=\S))/);
        parts.forEach((t, i) => T.push({ k: 'w', t, st: r.st, w: dxW_(t, r.st), href: r.href, cut: i > 0 }));
      }
    }
    for (let i = 0; i < T.length; i++) { const t = T[i], p = T[i - 1]; if (t.k === 'w' || t.k === 'fld') t.bk = !p || p.k === 'sp' || p.k === 'tab' || p.k === 'br' || p.k === 'img' || !!t.cut || (p.k === 'w' && /[-\u2013\u2014]$/.test(p.t) && /^\S/.test(t.t || '')); }
    tokCache.set(b, T); return T;
  }
  const lineSp = (P, mark) => P.lineRule === 'exact' ? P.lineExact : P.lineRule === 'atLeast' ? Math.max(P.lineExact, mark.size * mark.met[0]) : mark.size * mark.met[0] * (P.lineMult || 1);
  const leadCh = { dot: '.', hyphen: '-', underscore: '_', middleDot: '\u00b7', heavy: '_' };
  // one line from token `from`, `avail` points wide; `org`: where the line starts, from the text area's left edge (tab stops count from there)
  function buildLine(b, from, avail, org, first, secGrid) {
    const T = toks(b), P = b.P, segs = [], mets = [];
    let i = Math.floor(from + 1e-9), skip = Math.round((from - i) * 1e4), x = 0, cw = 0, grp = null, tabAt = 0, bp = null, brk = null, sp = 0, spW = 0, tabbed = false, to = null;
    const add = (t, w) => { const l = segs[segs.length - 1]; if (l && l.k === 't' && t.k === 'w' || l && l.k === 't' && t.k === 'sp') { if (l.st === t.st && l.href === t.href && Math.abs(l.x + l.w - x) < 1e-6) { l.t += t.t; l.w += w; x += w; return; } } segs.push({ k: t.k === 'img' ? 'img' : t.k === 'fld' ? 'fld' : 't', t: t.t, st: t.st, x, w, d: t.d, h: t.h, fld: t.fld, href: t.href }); x += w; };
    const endGroup = () => { // a right / centre / decimal tab: the text after it ends at (centres on) the stop
      if (!grp) return; const gw = cw - grp.x0; let ns = grp.t.val === 'center' ? grp.t.pos - org - gw / 2 : grp.t.val === 'decimal' ? grp.t.pos - org - grp.dec() : grp.t.pos - org - gw; ns = Math.max(ns, grp.x0); const dx = ns - grp.x0;
      if (dx) { for (let k = grp.seg; k < segs.length; k++) segs[k].x += dx; x += dx; cw += dx; }
      if (grp.lead && ns - grp.x0 > 1) segs.splice(grp.seg, 0, { k: 'lead', x: grp.x0, w: ns - grp.x0, st: grp.st, ch: grp.lead });
      grp = null;
    };
    const nextStop = at => { const stops = (P.tabs || []).filter(s => s.val !== 'clear' && s.val !== 'bar' && s.val !== 'num'), hs = P.hang ? (P.indL || 0) : null; let s = stops.find(s => s.pos > at + 0.01); if (hs != null && hs > at + 0.01 && (!s || hs < s.pos)) s = { pos: hs, val: 'left', leader: 'none' }; return s || null; };
    if (first && b.label) { // the list number, then its tab (the hanging indent is a tab stop)
      const L = b.label, lw = dxW_(L.text, L.st), lx = L.jc === 'right' ? -lw : L.jc === 'center' ? -lw / 2 : 0;
      segs.push({ k: 't', t: L.text, st: L.st, x: lx, w: lw, lbl: true }); x = cw = Math.max(0, lx + lw);
      if (L.suff === 'space') x += dxW_(' ', L.st);
      else if (L.suff !== 'nothing') { const at = org + x, s = nextStop(at); x = (s ? s.pos : (Math.floor((at + 0.01) / defTab) + 1) * defTab) - org; }
      cw = x; tabAt = segs.length;
    }
    const tokAt = k => { const t = T[k]; if (k !== Math.floor(from + 1e-9) || !skip || t.k !== 'w') return t; const tail = [...t.t].slice(skip).join(''); return { ...t, t: tail, w: dxW_(tail, t.st), bk: true }; };
    while (i < T.length) {
      const t = tokAt(i);
      if (t.k === 'br') { brk = t.br; mets.push(t.st); i++; break; }
      if (t.k === 'sp') { add(t, t.w); sp += t.t.length; spW += t.w; i++; continue; }
      if (t.k === 'tab') {
        endGroup(); mets.push(t.st);
        const at = org + cw; let s = nextStop(at);
        if (!s) { const p = (Math.floor((at + 0.01) / defTab) + 1) * defTab; if (p - org > avail + 0.01 && (segs.some(q => !q.lbl) || tabbed)) break; s = { pos: p, val: 'left', leader: 'none' }; }
        const lead = leadCh[s.leader] || null;
        if (s.val === 'left' || s.val === 'start' || s.val === 'num') { const nx = Math.max(cw, s.pos - org); if (lead && nx - cw > 1) segs.push({ k: 'lead', x: cw, w: nx - cw, st: t.st, ch: lead }); x = cw = nx; }
        else { x = cw; const g = { t: s, x0: cw, seg: segs.length, lead, st: t.st }; g.dec = () => { let w = 0; for (let k = g.seg; k < segs.length; k++) { const q = segs[k]; if (q.k !== 't') { w += q.w; continue; } const j = q.t.search(/[.,]/); if (j >= 0) return w + dxW_(q.t.slice(0, j), q.st); w += q.w; } return w; }; grp = g; }
        tabAt = segs.length; sp = 0; spW = 0; tabbed = true; i++; continue;
      }
      // a word (with the pieces glued to it), a picture or a field: does it fit?
      let j = i, w = 0; do { w += j === i ? t.w : T[j].w; j++; } while (j < T.length && (T[j].k === 'w' || T[j].k === 'fld') && !T[j].bk);
      const sq = P.align === 'both' || P.align === 'distribute' ? DX_SQUEEZE * spW : 0;
      const end = grp ? (grp.t.val === 'center' ? Math.max(grp.t.pos - org + (cw + w - grp.x0) / 2, cw + w) : Math.max(grp.t.pos - org, cw + w)) : x + w;
      const fits = end - sq <= avail + (sq ? 0.15 : 0.02), has = segs.some(q => !q.lbl) || tabbed;
      if (!fits && has) {
        if (t.bk || t.k === 'img') break;
        if (bp) { segs.length = bp.n; if (bp.last) Object.assign(segs[bp.n - 1], bp.last); x = bp.x; cw = bp.cw; i = bp.i; break; }
      }
      if (!fits && !has && t.k === 'w' && j === i + 1 && x + t.w > avail) { // one word wider than the line: it is cut between letters
        const ch = [...t.t]; let k = 1; while (k < ch.length && x + dxW_(ch.slice(0, k + 1).join(''), t.st) <= avail) k++;
        if (k < ch.length) { const head = ch.slice(0, k).join(''); add({ ...t, t: head }, dxW_(head, t.st)); cw = x; to = i + (skip + k) / 1e4; break; }
      }
      if (t.bk || t.k === 'img') { const l = segs[segs.length - 1]; bp = { n: segs.length, x, cw, i, last: l ? { t: l.t, w: l.w } : null }; }
      add(t, t.w); for (let q = i + 1; q < j; q++) add(T[q], T[q].w); i = j; cw = x;
    }
    endGroup();
    if (to == null) to = i;
    const last = to >= T.length && !brk;
    let nsp = 0; segs.forEach((s, k) => { if (s.k !== 't' || k < tabAt || s.lbl) return; const rest = segs.slice(k + 1).every(q => q.k === 't' && !q.t.trim()); nsp += ((rest ? s.t.replace(/ +$/, '') : s.t).match(/ /g) || []).length; });
    // height and baseline
    let a = 0, d = 0, img = 0, any = false;
    const met = st => { const s = st.size; a = Math.max(a, (st.met[0] - st.met[1]) * s); d = Math.max(d, st.met[1] * s); any = true; };
    for (const s of segs) { if (s.k === 'img') { img = Math.max(img, s.h); met(s.st); } else met(s.st); }
    mets.forEach(met); if (!any) met(b.mark);
    const aI = Math.max(a, img), nat = a + d; let h, base;
    if (P.lineRule === 'exact') { h = P.lineExact; base = 0.8 * h; }
    else if (P.lineRule === 'atLeast') { h = Math.max(P.lineExact, aI + d); base = h - d; }
    else { // multiple spacing: the extra room goes below the text; a picture's line is the picture (it is not multiplied)
      const m = P.lineMult || 1; h = Math.max(nat * m, aI + d); if (secGrid && P.snapToGrid !== false && !img) h = Math.ceil(h / secGrid - 0.01) * secGrid;
      base = m >= 1 || img > a ? aI : h - d;
    }
    return { segs, from, to, w: cw, nsp, tabAt, brk, last, h, base };
  }
  // all lines of a paragraph at one width (cells, headers, text boxes, look-ahead)
  const linesCache = new Map();
  function linesAt(b, width, grid) {
    const key = width.toFixed(2) + '|' + grid; let m = linesCache.get(b); if (!m) linesCache.set(b, m = new Map()); let r = m.get(key); if (r) return r; r = [];
    const P = b.P, indL = P.indL || 0, indR = P.indR || 0, fi = P.hang ? -P.hang : (P.first || 0); let from = 0, k = 0;
    for (;;) { const o = indL + (k === 0 ? fi : 0), ln = buildLine(b, from, width - o - indR, o, k === 0, grid); ln.org = o; ln.avail = width - o - indR; r.push(ln); k++; if (ln.last || k > 5000) break; if (ln.to === from && !ln.brk) { ln.last = true; break; } from = ln.to; }
    m.set(key, r); return r;
  }
  const spaceBefore = (b, prev, top) => { const P = b.P; if (prev && prev.t === 'p' && P.contextualSpacing && prev.sid === b.sid) return 0; return P.beforeAuto ? 14 : P.beforeLines ? P.beforeLines * b.mark.size * 1.2 : (P.before || 0); };
  const spaceAfter = (b, next) => { const P = b.P; if (next && next.t === 'p' && P.contextualSpacing && next.sid === b.sid) return 0; return P.afterAuto ? 14 : P.afterLines ? P.afterLines * b.mark.size * 1.2 : (P.after || 0); };
  const sameBdr = (a, b) => a && b && a.t === 'p' && b.t === 'p' && a.P.bdr && b.P.bdr && JSON.stringify(a.P.bdr) === JSON.stringify(b.P.bdr) && (a.P.indL || 0) === (b.P.indL || 0) && (a.P.indR || 0) === (b.P.indR || 0);
  const bdrTop = (b, prev) => b.P.bdr && b.P.bdr.top && !sameBdr(prev, b) ? b.P.bdr.top.space + b.P.bdr.top.w : sameBdr(prev, b) && b.P.bdr.between ? b.P.bdr.between.space + b.P.bdr.between.w : 0;
  const bdrBot = (b, next) => b.P.bdr && b.P.bdr.bottom && !sameBdr(b, next) ? b.P.bdr.bottom.space + b.P.bdr.bottom.w : 0;

  // ── drawing out
  const fixups = [];
  const rect = (L, x, y, w, h, fill, op) => { if (w > 0.01 && h > 0.01) L.push({ id: uid(), type: 'rect', x, y, w, h, color: 'none', fill, width: 0, opacity: op ?? 1 }); };
  function textOut(L, text, st, x, base, target) {
    if (!text.trim()) return;
    const size = st.msize, y = base - st.rise - X.BASE * size, common = { color: st.color, font: st.fam, bold: st.bold, italic: st.italic, underline: st.underline, strike: st.strike };
    if (Math.abs(st.sp) >= 0.5 && [...text].length <= 80) { // letter-spaced text: letter by letter
      let cx = x; for (const ch of text) { const cw = dxW100(ch, st.css) * size / 100 * st.scale; if (ch !== ' ') L.push(X.textObj(ch, size, { x: cx, y, ...common })); cx += cw + st.sp; }
      return;
    }
    const o = X.textObj(text, size, { x, y, ...common }), n = (text.match(/ /g) || []).length;
    if (target != null && target > 0) {
      const drawn = X.measure(text, o); let ws = n ? clamp((target - drawn) / (n * size), -0.08, 3) : 0; if (Math.abs(ws) > 0.002) o.ws = +ws.toFixed(4); else ws = 0;
      const k = target / (drawn + n * ws * size); if (k < 0.985) { o.size = +(size * Math.max(0.85, k)).toFixed(3); o.y = base - st.rise - X.BASE * o.size; if (o.ws) o.ws = +(o.ws * size / o.size).toFixed(4); }
    }
    L.push(o); return o;
  }
  function emitLine(E, b, ln, x0, top, justify) { // x0: the text area's left edge; top: the line's top
    const P = b.P, room = ln.avail, k = P.align === 'center' ? 0.5 : P.align === 'right' ? 1 : 0;
    let off = 0, ex = 0;
    const just = (P.align === 'both' && (!ln.last && !(ln.brk && noExpandBr)) || P.align === 'distribute') && ln.nsp > 0 && justify !== false;
    if (just) ex = (room - ln.w) / ln.nsp; else off = Math.max(-ln.org, (room - ln.w) * k);
    const base = top + ln.base, L = E.L.main;
    let acc = 0;
    ln.segs.forEach((s, si) => {
      const sx = x0 + ln.org + off + s.x + acc;
      if (s.k === 'img') { const d = s.d; if (d) drawShapes(E, d.shapes, sx, base - s.h, d.w, d.h, L); return; }
      if (s.k === 'lead') { const cw = dxW_(s.ch, s.st); if (cw > 0.5) { const n = Math.floor((s.w - 1) / cw); if (n > 0) textOut(L, s.ch.repeat(n), s.st, sx + s.w - n * cw - 0.5, base); } return; }
      if (s.k === 'fld') { const o = textOut(L, s.fld === 'PAGE' ? String(E.num) : '0', s.st, sx, base); if (o && s.fld !== 'PAGE') fixups.push({ o, f: s.fld, sec: E.secIdx }); return; }
      let t = s.t; if (si === ln.segs.length - 1 || ln.segs.slice(si + 1).every(q => q.k === 't' && !q.t.trim())) t = t.replace(/ +$/, '');
      const lead = (t.match(/^ +/) || [''])[0].length, n = (t.match(/ /g) || []).length, inJ = just && si >= ln.tabAt;
      const w = dxW_(t, s.st) + (inJ ? ex * n : 0);
      if (s.st.hl && t.trim()) rect(E.L.main, sx, top + ln.base - (s.st.met[0] - s.st.met[1]) * s.st.size, w, s.st.met[0] * s.st.size, s.st.hl);
      if (t.trim()) {
        const lx = lead ? dxW_(t.slice(0, lead), s.st) + (inJ ? ex * lead : 0) : 0, body = t.slice(lead);
        textOut(L, body, s.st, sx + lx, base, w - lx);
        if (s.href && E.L.links) E.L.links.push({ id: uid(), type: 'link', x: sx, y: base - s.st.size, w, h: s.st.size * 1.25, url: s.href });
      }
      if (inJ) acc += ex * (s.t.match(/ /g) || []).length;
    });
  }
  function paraDeco(E, b, x0, width, y0, y1, prev, next) { // borders and shading of one paragraph (or its part on this page)
    const P = b.P, B = P.bdr || {}, l = x0 + (P.indL || 0) - DX_RULE_OUT - (B.left ? B.left.space + B.left.w : 0), r = x0 + width - (P.indR || 0) + DX_RULE_OUT + (B.right ? B.right.space + B.right.w : 0);
    if (P.shd) E.L.main.push({ id: uid(), type: 'rect', x: l, y: y0, w: r - l, h: y1 - y0, color: 'none', fill: P.shd, width: 0, opacity: 1, _bg: 1 });
    const L = E.L.main;
    if (B.top && !sameBdr(prev, b)) rect(L, l, y0, r - l, B.top.w, B.top.color);
    else if (B.between && sameBdr(prev, b)) rect(L, l, y0, r - l, B.between.w, B.between.color);
    if (B.bottom && !sameBdr(b, next)) rect(L, l, y1 - B.bottom.w, r - l, B.bottom.w, B.bottom.color);
    if (B.left) rect(L, l, y0, B.left.w, y1 - y0, B.left.color);
    if (B.right) rect(L, r - B.right.w, y0, B.right.w, y1 - y0, B.right.color);
  }
  function raster(s, w, h) { // a shape that is not a plain box → a picture of it
    const k = Math.min(3, 2400 / Math.max(1, w, h)), lw = s.line ? s.line.w : 0, pad = lw / 2 + 1, c = document.createElement('canvas');
    c.width = Math.max(1, Math.ceil((w + 2 * pad) * k)); c.height = Math.max(1, Math.ceil((h + 2 * pad) * k)); const g = c.getContext('2d'); g.scale(k, k); g.translate(pad, pad);
    g.beginPath(); const geo = s.geom;
    if (geo === 'ellipse' || geo === 'flowChartConnector') g.ellipse(w / 2, h / 2, Math.max(0.1, w / 2), Math.max(0.1, h / 2), 0, 0, Math.PI * 2);
    else if (/^(line|straightConnector1|bentConnector\d|curvedConnector\d)$/.test(geo)) { const fx = s.flipH, fy = s.flipV; g.moveTo(fx ? w : 0, fy ? h : 0); g.lineTo(fx ? 0 : w, fy ? 0 : h); }
    else if (geo === 'roundRect' || geo === 'flowChartAlternateProcess') { const r = Math.min(w, h) * (s.adj != null ? s.adj : 0.16667); g.moveTo(r, 0); g.arcTo(w, 0, w, h, r); g.arcTo(w, h, 0, h, r); g.arcTo(0, h, 0, 0, r); g.arcTo(0, 0, w, 0, r); g.closePath(); }
    else if (geo === 'triangle' || geo === 'flowChartExtract') { g.moveTo(w / 2, 0); g.lineTo(w, h); g.lineTo(0, h); g.closePath(); }
    else if (geo === 'rtTriangle') { g.moveTo(0, 0); g.lineTo(w, h); g.lineTo(0, h); g.closePath(); }
    else if (geo === 'diamond' || geo === 'flowChartDecision') { g.moveTo(w / 2, 0); g.lineTo(w, h / 2); g.lineTo(w / 2, h); g.lineTo(0, h / 2); g.closePath(); }
    else g.rect(0, 0, w, h);
    if (s.fill && !/^(line|straightConnector1)$/.test(geo)) { g.globalAlpha = s.fill.a; g.fillStyle = s.fill.c; g.fill(); }
    if (s.line && s.line.w > 0) { g.globalAlpha = s.line.a; g.strokeStyle = s.line.c; g.lineWidth = s.line.w; g.stroke(); }
    return { url: c.toDataURL('image/png'), pad };
  }
  function rotateInto(L, objs, cx, cy, deg) { // objects drawn unturned → turned by deg around (cx, cy)
    const a = deg * Math.PI / 180, cs = Math.cos(a), sn = Math.sin(a);
    for (const o of objs) {
      const w = o.type === 'text' ? X.measure(o.text, o) : o.w, h = o.type === 'text' ? (o.lh || X.LH) * o.size : o.h, ox = o.x + w / 2 - cx, oy = o.y + h / 2 - cy;
      o.x = cx + ox * cs - oy * sn - w / 2; o.y = cy + ox * sn + oy * cs - h / 2; o.rot = ((o.rot || 0) + deg) % 360; L.push(o);
    }
  }
  function drawShapes(E, shapes, X0, Y0, W, H, L) { // a drawing's shapes at (X0, Y0)
    for (const s of shapes) {
      const x = X0 + s.x, y = Y0 + s.y, w = s.w, h = s.h, op = s.fill ? s.fill.a : 1;
      if (s.kind === 'pic') { if (s.asset) L.push({ id: uid(), type: 'image', asset: s.asset, x, y, w, h, opacity: 1, ...(s.rot ? { rot: s.rot } : {}) }); continue; }
      if (s.img) L.push({ id: uid(), type: 'image', asset: s.img, x, y, w, h, opacity: 1, ...(s.rot ? { rot: s.rot } : {}) });
      const plain = (s.geom === 'rect' || s.geom === 'flowChartProcess') && !s.rot;
      if (plain) {
        if (s.fill && !s.img) rect(L, x, y, w, h, s.fill.c, op);
        if (s.line && s.line.w > 0) { const lw = s.line.w; if (h <= lw * 1.5) rect(L, x, y - lw / 2 + h / 2, w, lw, s.line.c, s.line.a); else if (w <= lw * 1.5) rect(L, x - lw / 2 + w / 2, y, lw, h, s.line.c, s.line.a); else { rect(L, x - lw / 2, y - lw / 2, w + lw, lw, s.line.c, s.line.a); rect(L, x - lw / 2, y + h - lw / 2, w + lw, lw, s.line.c, s.line.a); rect(L, x - lw / 2, y + lw / 2, lw, h - lw, s.line.c, s.line.a); rect(L, x + w - lw / 2, y + lw / 2, lw, h - lw, s.line.c, s.line.a); } }
      } else if ((s.fill || s.line) && w + h > 0.5) {
        const horiz = /^(line|straightConnector1)$/.test(s.geom) && s.line && !s.rot;
        if (horiz && h < 0.01) rect(L, x, y - s.line.w / 2, w, s.line.w, s.line.c, s.line.a);
        else if (horiz && w < 0.01) rect(L, x - s.line.w / 2, y, s.line.w, h, s.line.c, s.line.a);
        else { const r = raster(s, w, h); L.push({ id: uid(), type: 'image', asset: X.addAsset(r.url), x: x - r.pad, y: y - r.pad, w: w + 2 * r.pad, h: h + 2 * r.pad, opacity: 1, ...(s.rot ? { rot: s.rot } : {}) }); }
      }
      if (s.txbx && s.txbx.blocks.length) {
        const [li, ti, ri, bi] = s.txbx.ins, vert = s.txbx.vert, turn = vert === 'vert' || vert === 'eaVert' ? 90 : vert === 'vert270' ? 270 : 0;
        const bw = (turn ? h : w) - li - ri, bh = (turn ? w : h) - ti - bi, lay = layBox(s.txbx.blocks, Math.max(10, bw));
        const dy = s.txbx.anchor === 'ctr' ? (bh - lay.h) / 2 : s.txbx.anchor === 'b' ? bh - lay.h : 0, sink = [], E2 = { ...E, L: { ...E.L, main: sink, behind: sink, front: sink, links: null }, colX: x, colW: w, cellX: x };
        const bx = turn ? x + w / 2 - (bw + li + ri) / 2 : x, by = turn ? y + h / 2 - (bh + ti + bi) / 2 : y;
        for (const it of lay.items) it.draw(E2, bx + li, by + ti + Math.max(0, dy) + it.y);
        const rot = ((s.rot || 0) + turn) % 360; if (rot) rotateInto(L, sink, x + w / 2, y + h / 2, rot); else L.push(...sink);
      }
    }
  }
  function placeFloat(E, d, paraTop, lineTop) { // an anchored drawing → page position; wrapping ones keep text away
    const pw = E.PW, ph = E.PH, sp = E.sp, W = d.w, H = d.h;
    const hFrame = { page: [0, pw], margin: [sp.ML, pw - sp.MR], column: [E.colX, E.colW], character: [E.colX, E.colW], leftMargin: [0, sp.ML], rightMargin: [pw - sp.MR, sp.MR], insideMargin: [0, sp.ML], outsideMargin: [pw - sp.MR, sp.MR] }[d.H.rel] || [E.colX, E.colW];
    if (E.cellX != null && d.inCell && (d.H.rel === 'column' || d.H.rel === 'character')) { hFrame[0] = E.cellX; hFrame[1] = E.cellW || E.colW; }
    if (d.H.rel === 'margin') hFrame[1] = pw - sp.MR - sp.ML; else if (d.H.rel === 'page') hFrame[1] = pw;
    let x = d.H.align ? hFrame[0] + ({ left: 0, inside: 0, center: (hFrame[1] - W) / 2, right: hFrame[1] - W, outside: hFrame[1] - W }[d.H.align] || 0) : d.H.pct != null ? hFrame[0] + d.H.pct * hFrame[1] : hFrame[0] + d.H.off;
    const vFrame = { page: [0, ph], margin: [sp.MT, ph - sp.MT - sp.MB], topMargin: [0, sp.MT], bottomMargin: [ph - sp.MB, sp.MB], insideMargin: [0, sp.MT], outsideMargin: [ph - sp.MB, sp.MB], paragraph: [paraTop, 0], line: [lineTop, 0] }[d.V.rel] || [paraTop, 0];
    let y = d.V.align ? vFrame[0] + ({ top: 0, inside: 0, center: (vFrame[1] - H) / 2, bottom: vFrame[1] - H, outside: vFrame[1] - H }[d.V.align] || 0) : d.V.pct != null ? vFrame[0] + d.V.pct * vFrame[1] : vFrame[0] + d.V.off;
    drawShapes(E, d.shapes, x, y, W, H, d.behind ? E.L.behind : E.L.front);
    if (d.wrap !== 'none' && E.excl) E.excl.push({ x0: x - d.dist[2] - d.eff[0], x1: x + W + d.dist[3] + d.eff[2], y0: y - d.dist[0] - d.eff[1], y1: y + H + d.dist[1] + d.eff[3], tb: d.wrap === 'wrapTopAndBottom', side: d.side });
  }

  // ── boxes that do not break across pages: table cells, headers, footers, text boxes → items { y, h, draw(E, x, y) }
  function layBox(bl, width, grid) {
    const items = []; let y = 0;
    bl.forEach((b, i) => {
      const prev = bl[i - 1], next = bl[i + 1];
      if (b.t === 'tbl') { const TL = tableLayout(b, width); TL.rows.forEach((R, ri) => { const yy = y; items.push({ y: yy, h: R.h, draw: (E, x, top) => drawRow(E, TL, ri, x + TL.x, top, R.h, 0) }); y += R.h; }); return; }
      const top = y, before = i === 0 && false ? 0 : spaceBefore(b, prev), bt = bdrTop(b, prev), lines = linesAt(b, width, grid || 0);
      const pTop = top; y += before + bt; const yText = y;
      const ls = lines.map(ln => { const it = { y, h: ln.h, ln }; y += ln.h; return it; });
      y += bdrBot(b, next); const yEnd = y; y += spaceAfter(b, next);
      items.push({ y: pTop, h: 0, draw: (E, x, yy) => { if ((b.P.bdr && Object.keys(b.P.bdr).length) || b.P.shd) paraDeco(E, b, x, width, yy + before, yy + before + (yEnd - pTop - before), prev, next); for (const d of b.anchors) placeFloat({ ...E, colX: E.cellX != null ? E.cellX : x, colW: E.cellW || width }, d, yy, yy + before); } });
      for (const it of ls) items.push({ y: it.y, h: it.h, draw: (E, x, yy) => emitLine(E, b, it.ln, x, yy) });
    });
    return { h: y, items };
  }

  // ── tables
  const tlCache = new Map();
  function tableLayout(T, avail) {
    const key = avail.toFixed(2); let m = tlCache.get(T); if (!m) tlCache.set(T, m = new Map()); if (m.get(key)) return m.get(key);
    const nCols = Math.max(1, ...T.rows.map(r => r.before + r.cells.reduce((a, c) => a + c.span, 0)));
    let cols = T.grid.slice(0, nCols); if (cols.length < nCols || cols.some(w => !(w > 0))) { const tot = T.w && T.w.type === 'dxa' && T.w.v ? T.w.v / 20 : T.w && T.w.type === 'pct' ? avail * T.w.v / 5000 : avail - T.ind; cols = Array(nCols).fill(Math.max(10, tot / nCols)); }
    let tot = cols.reduce((a, c) => a + c, 0);
    if (!T.fixed && tot > avail - T.ind + 0.5 && tot > 0) { const k = (avail - T.ind) / tot; cols = cols.map(c => c * k); tot = avail - T.ind; }
    const xs = [0]; cols.forEach(c => xs.push(xs[xs.length - 1] + c));
    const x = T.jc === 'center' ? (avail - tot) / 2 : T.jc === 'right' || T.jc === 'end' ? avail - tot : T.ind;
    const nR = T.rows.length, cellAt = [];
    const rows = T.rows.map((row, ri) => {
      let ci = row.before; const cells = [];
      row.cells.forEach(c => {
        const x0 = xs[Math.min(ci, nCols)], x1 = xs[Math.min(ci + c.span, nCols)], first = ci === 0, lastC = ci + c.span >= nCols, mar = { ...T.mar, ...c.mar };
        const edge = (s, outer, inner) => { const v = c.bd[s]; return v === 'none' ? null : v || (outer ? T.bd[s] : T.bd[inner]) || null; };
        const bd = { top: edge('top', ri === 0, 'insideH'), bottom: edge('bottom', ri === nR - 1, 'insideH'), left: edge('left', first, 'insideV'), right: edge('right', lastC, 'insideV') };
        const cw = Math.max(4, x1 - x0 - mar.left - mar.right), lay = c.vm === 'cont' ? { h: 0, items: [] } : layBox(c.blocks, cw);
        cells.push({ c, x0, x1, mar, bd, lay, ci, span: c.span }); ci += c.span;
      });
      const tw = ri === 0 ? Math.max(0, ...cells.map(c => c.bd.top ? c.bd.top.w : 0)) : 0, bw = Math.max(0, ...cells.map(c => c.bd.bottom ? c.bd.bottom.w : 0));
      const need = Math.max(0, ...cells.filter(c => c.c.vm !== 'start').map(c => c.lay.h + c.mar.top + c.mar.bottom));
      const inner = row.hRule === 'exact' && row.h ? row.h : Math.max(row.h, need);
      return { cells, tw, bw, inner, h: tw + inner + bw, row };
    });
    // merged cells: the rows they span grow to hold their text
    rows.forEach((R, ri) => R.cells.forEach(ce => {
      if (ce.c.vm !== 'start') return; let k = ri + 1; while (k < rows.length && rows[k].cells.some(o => o.ci === ce.ci && o.c.vm === 'cont')) k++;
      ce.rows = k - ri; const have = rows.slice(ri, k).reduce((a, r) => a + r.h, 0) - R.tw - rows[k - 1].bw, need = ce.lay.h + ce.mar.top + ce.mar.bottom;
      if (need > have + 0.01) { const L = rows[k - 1]; if (L.row.hRule !== 'exact') { L.inner += need - have; L.h += need - have; } }
    }));
    const r = { x, w: tot, rows, xs, nCols }; m.set(key, r); return r;
  }
  function drawRow(E, TL, ri, x, top, h, cut, limit) { // one row at (x, top), h high; split rows: only what lies from `cut` (to cut + limit) on this page
    const R = TL.rows[ri], L = E.L.main, whole = !cut && limit == null;
    for (const ce of R.cells) {
      let ch = h; if (ce.rows > 1 && whole) ch = TL.rows.slice(ri, ri + ce.rows).reduce((a, r) => a + r.h, 0);
      if (ce.c.shade) rect(L, x + ce.x0, top, ce.x1 - ce.x0, ce.c.vm === 'cont' ? h : ch, ce.c.shade);
      if (ce.c.vm !== 'cont') {
        const innerH = ce.rows > 1 ? ch - R.tw - (TL.rows[ri + ce.rows - 1] || R).bw : R.inner, free = innerH - ce.lay.h - ce.mar.top - ce.mar.bottom;
        const dy = whole ? (ce.c.va === 'center' ? Math.max(0, free / 2) : ce.c.va === 'bottom' ? Math.max(0, free) : 0) : 0, E2 = { ...E, cellX: x + ce.x0, cellW: ce.x1 - ce.x0 };
        for (const it of ce.lay.items) {
          const iy = R.tw + ce.mar.top + dy + it.y; if (iy < (cut || 0) - 0.01) continue; if (limit != null && iy + it.h > cut + limit + 0.01) continue;
          it.draw(E2, x + ce.x0 + ce.mar.left, top + iy - (cut || 0));
        }
      }
      const b = ce.bd, l = x + ce.x0, r = x + ce.x1;
      if (b.top && ri === 0 && !cut) rect(L, l, top, r - l, b.top.w, b.top.color);
      if (b.bottom && (!ce.rows || ce.rows === 1) && limit == null) rect(L, l, top + h - b.bottom.w, r - l, b.bottom.w, b.bottom.color);
      if (b.left) rect(L, l - b.left.w / 2, top, b.left.w, h, b.left.color);
      if (b.right) rect(L, r - b.right.w / 2, top, b.right.w, h, b.right.color);
    }
  }
  // ── pages
  const pages = []; let PR = null, sec = null, secIdx = -1, cols = [], colI = 0, regTop = 0, y = 0, bottom = 0, placed = false, natural = false, pageNo = 0, colMax = [], trial = null;
  const hfLay = new Map();
  const hfOf = (s, kind, first, num) => { const P = s.props, t = P.titlePg && first ? 'first' : evenOdd && num % 2 === 0 ? 'even' : 'default', id = P.refs[kind + ':' + t]; return id && hfParts[id] ? { id, blocks: hfParts[id] } : null; };
  const hfLayOf = (hf, s) => { const w = s.props.PW - s.props.ML - s.props.MR, k = hf.id + '|' + w.toFixed(1); let l = hfLay.get(k); if (!l) hfLay.set(k, l = layBox(hf.blocks, w)); return l; };
  const OVER = { over: true };
  function setCols(s, top) { sec = s; cols = s.props.cols.map(c => ({ x: s.props.ML + c.x, w: c.w })); colI = 0; regTop = top; y = top; colMax = cols.map(() => top); placed = false; }
  function newPage(s, hard) {
    if (trial) throw OVER;
    const P = s.props, first = !pages.some(p => p.sec === s); pageNo = first && P.start != null ? P.start : pageNo + 1;
    PR = { sec: s, secIdx, first, num: pageNo, L: { behind: [], main: [], front: [], links: [] }, excl: [], pg: { id: uid(), src: -1, w: P.PW, h: P.PH, rot: 0, objs: [] } };
    pages.push(PR);
    const hd = hfOf(s, 'header', first, pageNo), ft = hfOf(s, 'footer', first, pageNo);
    const top = Math.max(P.MT, hd ? P.HD + hfLayOf(hd, s).h : 0); bottom = P.PH - Math.max(P.MB, ft ? P.FD + hfLayOf(ft, s).h : 0);
    if (bottom - top < 36) bottom = Math.min(P.PH, top + 36);
    setCols(s, top); natural = !hard;
  }
  function nextRegion(hard) { colMax[colI] = Math.max(colMax[colI], y); if (colI < cols.length - 1) { colI++; y = regTop; placed = false; natural = !hard; } else newPage(sec, hard); }
  const E0 = () => ({ L: PR.L, PW: PR.pg.w, PH: PR.pg.h, sp: sec.props, colX: cols[colI].x, colW: cols[colI].w, num: PR.num, secIdx, excl: PR.excl });
  const room = (x0, w, yy, h) => { // a line's room beside the drawings text wraps around
    let l = x0, r = x0 + w, jump = null;
    for (const e of PR.excl) {
      if (e.y1 <= yy + 0.01 || e.y0 >= yy + h - 0.01 || e.x1 <= l || e.x0 >= r) continue;
      if (e.tb) { jump = Math.max(jump || 0, e.y1); continue; }
      const lf = e.x0 - l, rf = r - e.x1;
      if (e.side === 'left' || (e.side !== 'right' && lf >= rf && e.side !== 'bothSides') || (e.side === 'bothSides' && lf >= rf)) r = Math.min(r, e.x0); else l = Math.max(l, e.x1);
      if (r - l < 18) jump = Math.max(jump || 0, e.y1);
    }
    return { l, r, jump };
  };
  function placePara(b, prev, next, keepReq) {
    const P = b.P, col = () => cols[colI], grid = sec.props.grid;
    if (P.pageBreakBefore && placed) newPage(sec, true);
    const bt = bdrTop(b, prev), bb = bdrBot(b, next), after = spaceAfter(b, next), bef = () => (!placed && natural ? 0 : spaceBefore(b, prev));
    if (placed) { // can it start here (with what has to stay with it)?
      const ls = linesAt(b, col().w, grid), all = bef() + bt + ls.reduce((a, l) => a + l.h, 0) + bb, wc = P.widowControl !== false && ls.length >= 2;
      let need = bef() + bt + ls[0].h + (wc ? ls[1].h : 0);
      if (P.keepLines) need = all; if (keepReq && (P.keepLines || ls.length <= 3)) need = all + after + keepReq;
      if (y + need > bottom + 0.02 && (need <= bottom - regTop + 0.02 || y + bef() + bt + ls[0].h > bottom + 0.02)) nextRegion(false);
    }
    const before = bef(), startY = y, E = E0();
    for (const d of b.anchors) placeFloat(E, d, y, y + before);
    y += before + bt;
    let from = 0, k = 0, partTop = y, first = true;
    for (let guard = 0; guard < 5000; guard++) {
      const c = col(), ls = []; let yy = y, f = from, kk = k;
      for (;;) { // the rest of the paragraph, line by line, in this column (beside the drawings text wraps around)
        const o = (P.indL || 0) + (kk === 0 ? (P.hang ? -P.hang : P.first || 0) : 0), w = c.w - o - (P.indR || 0);
        let rm = PR.excl.length ? room(c.x + o, w, yy, b.mark.size * 1.2) : null; if (rm && rm.jump != null && rm.jump > yy) { yy = rm.jump; rm = room(c.x + o, w, yy, b.mark.size * 1.2); }
        const av2 = rm ? rm.r - rm.l : w, ln = buildLine(b, f, av2, o, kk === 0, grid); ln.org = o; ln.avail = av2; ln.dx = rm ? rm.l - (c.x + o) : 0; ln.top = yy;
        ls.push(ln); yy += ln.h; kk++; if (ln.last || ln.brk === 'page' || ln.brk === 'column' || (ln.to === f && !ln.brk)) break; f = ln.to;
      }
      let n = 0; while (n < ls.length && ls[n].top + ls[n].h <= bottom + 0.02) n++;
      if (n < ls.length) {
        const all = ls[ls.length - 1].last;
        if (P.widowControl !== false) { if (all && ls.length - n === 1 && n >= 2) n--; if (k === 0 && n === 1 && ls.length >= 2) n = 0; }
        if (P.keepLines && k === 0 && placed) n = 0;
        if (n === 0 && !placed) n = 1; // (taller than the whole column: it goes here anyway)
      }
      if (n === 0) { nextRegion(false); partTop = y; continue; }
      const Ex = E0();
      for (let q = 0; q < n; q++) { const ln = ls[q]; emitLine(Ex, b, ln, c.x + ln.dx, ln.top); y = ln.top + ln.h; placed = true; }
      const lastLn = ls[n - 1], done = n === ls.length && (lastLn.last || (lastLn.to === lastLn.from && !lastLn.brk)), forced = n === ls.length && !done && (lastLn.brk === 'page' || lastLn.brk === 'column');
      if ((P.bdr && Object.keys(P.bdr).length) || P.shd) paraDeco(Ex, b, c.x, c.w, first ? startY + before : partTop, done ? y + bb : y, first ? prev : b, done ? next : b);
      if (done || (n === ls.length && !forced)) break;
      if (forced) { if (lastLn.brk === 'column') nextRegion(true); else newPage(sec, true); from = lastLn.to; }
      else { nextRegion(false); from = ls[n].from; }
      k += n; partTop = y; first = false;
    }
    y += bb; colMax[colI] = Math.max(colMax[colI], y); y += after;
  }
  function placeTable(T, prev, next) {
    const TL = tableLayout(T, cols[colI].w), heads = []; for (const R of TL.rows) { if (R.row.head) heads.push(R); else break; }
    const head = ri => { if (ri < heads.length) return; for (let hi = 0; hi < heads.length; hi++) { drawRow(E0(), TL, hi, cols[colI].x + TL.x, y, heads[hi].h, 0); y += heads[hi].h; placed = true; } };
    for (let ri = 0; ri < TL.rows.length; ri++) {
      const R = TL.rows[ri], h = R.h; let cut = 0;
      for (let guard = 0; guard < 500; guard++) {
        const rem = h - cut, full = bottom - regTop;
        if (y + rem <= bottom + 0.02 || (!placed && (R.row.cant || rem <= full + 0.02))) { drawRow(E0(), TL, ri, cols[colI].x + TL.x, y, rem, cut, null); y += rem; placed = true; break; }
        if (placed && (R.row.cant || (cut === 0 && R.row.hRule === 'exact'))) { nextRegion(false); head(ri); continue; }
        // the row breaks here: what fits stays, the rest goes on
        const avail = bottom - y; let c2 = cut + avail;
        for (const ce of R.cells) for (const it of ce.lay.items) { const iy = R.tw + ce.mar.top + it.y; if (it.h && iy < c2 && iy + it.h > c2 + 0.01) c2 = Math.min(c2, iy); }
        if (c2 - cut < 1) { if (placed) { nextRegion(false); head(ri); continue; } c2 = cut + avail; }
        drawRow(E0(), TL, ri, cols[colI].x + TL.x, y, avail, cut, c2 - cut); placed = true; y += avail; cut = c2;
        nextRegion(false); head(ri);
      }
    }
    colMax[colI] = Math.max(colMax[colI], y);
  }
  function keepNeed(bl, i) { // how much of what follows a keep-with-next paragraph has to come along
    const nb = bl[i + 1]; if (!nb) return 0;
    if (nb.t === 'tbl') { const TL = tableLayout(nb, cols[colI].w); return TL.rows.length ? TL.rows[0].h : 0; }
    const ls = linesAt(nb, cols[colI].w, sec.props.grid), body = ls.reduce((a, l) => a + l.h, 0);
    return spaceBefore(nb, bl[i]) + bdrTop(nb, bl[i]) + (nb.P.keepNext && ls.length <= 3 ? body + spaceAfter(nb, bl[i + 2]) + keepNeed(bl, i + 1) : ls[0].h + (ls.length > 1 && nb.P.widowControl !== false ? ls[1].h : 0));
  }
  function placeBlocks(bl) {
    bl.forEach((b, i) => {
      const prev = bl[i - 1], next = bl[i + 1];
      if (b.t === 'p') placePara(b, prev, next, b.P.keepNext ? keepNeed(bl, i) : 0); else placeTable(b, prev, next);
    });
  }
  sections.forEach((s, si) => {
    secIdx = si; const P = s.props, type = si === 0 ? 'nextPage' : P.type;
    if (!PR) newPage(s, true);
    else if (type === 'continuous' && Math.abs(PR.pg.w - P.PW) < 1 && Math.abs(PR.pg.h - P.PH) < 1) { colMax[colI] = Math.max(colMax[colI], y); const top = cols.length > 1 ? Math.max(...colMax) : Math.max(y, ...colMax); setCols(s, top); placed = true; natural = false; }
    else if (type === 'nextColumn' && colI < cols.length - 1) { nextRegion(true); sec = s; }
    else { newPage(s, true); if ((type === 'evenPage' && PR.num % 2) || (type === 'oddPage' && PR.num % 2 === 0)) newPage(s, true); /* (Word leaves the page before blank) */ }
    const nx = sections[si + 1], balance = cols.length > 1 && nx && nx.props.type === 'continuous';
    if (!balance) { placeBlocks(s.blocks); return; }
    // columns before a continuous break are balanced: the lowest bottom that keeps it all on this page
    const snap = { pages: pages.length, PR, L: { behind: PR.L.behind.length, main: PR.L.main.length, front: PR.L.front.length, links: PR.L.links.length }, excl: PR.excl.length, fix: fixups.length, y, regTop, bottom, placed, natural, colI, colMax: colMax.slice(), pageNo };
    const restore = () => { pages.length = snap.pages; PR = snap.PR; for (const k in snap.L) PR.L[k].length = snap.L[k]; PR.excl.length = snap.excl; fixups.length = snap.fix; y = snap.y; regTop = snap.regTop; bottom = snap.bottom; placed = snap.placed; natural = snap.natural; colI = snap.colI; colMax = snap.colMax.slice(); pageNo = snap.pageNo; sec = s; };
    const fitsWithin = lim => { restore(); bottom = lim; trial = true; try { placeBlocks(s.blocks); return true; } catch (e) { if (e !== OVER) throw e; return false; } finally { trial = null; } };
    if (fitsWithin(snap.bottom)) {
      let lo = snap.regTop, hi = snap.bottom; for (let it = 0; it < 14 && hi - lo > 0.5; it++) { const mid = (lo + hi) / 2; if (fitsWithin(mid)) hi = mid; else lo = mid; }
      restore(); bottom = hi + 0.01; trial = true; try { placeBlocks(s.blocks); } finally { trial = null; } bottom = snap.bottom;
    } else { restore(); placeBlocks(s.blocks); }
  });
  if (!pages.length) newPage(sections[sections.length - 1], true);

  // ── headers, footers and page numbers, then the layers in Word's order
  const total = pages.length;
  for (const f of fixups) f.o.text = String(f.f === 'SECTIONPAGES' ? pages.filter(p => p.secIdx === f.sec).length : total);
  for (const p of pages) {
    const s = p.sec, P = s.props, E = { L: { behind: [], main: [], front: [], links: p.L.links }, PW: p.pg.w, PH: p.pg.h, sp: P, colX: P.ML, colW: P.PW - P.ML - P.MR, num: p.num, secIdx: p.secIdx, excl: null };
    const hdrBehind = E.L.behind, hdrMain = E.L.main, hdrFront = E.L.front, fx = fixups.length;
    const hd = hfOf(s, 'header', p.first, p.num), ft = hfOf(s, 'footer', p.first, p.num);
    if (hd) { const l = hfLayOf(hd, s); for (const it of l.items) it.draw(E, P.ML, P.HD + it.y); }
    if (ft) { const l = hfLayOf(ft, s), top = P.PH - P.FD - l.h; for (const it of l.items) it.draw(E, P.ML, top + it.y); }
    for (const f of fixups.slice(fx)) f.o.text = String(f.f === 'SECTIONPAGES' ? pages.filter(q => q.secIdx === f.sec).length : total);
    const bg = p.L.main.filter(o => o._bg), main = p.L.main.filter(o => !o._bg); bg.forEach(o => delete o._bg);
    p.pg.objs = [...hdrBehind, ...hdrMain, ...hdrFront, ...p.L.behind, ...bg, ...main, ...p.L.front, ...p.L.links];
  }
  const out = pages.map(p => p.pg);
  if (warn.length) { out.warn = warn.join(' '); console.warn(out.warn); try { toast(out.warn, 'warn'); } catch (e) { /* no toast here */ } }
  return out;
}

/* ───────── other exports ───────── */
async function pdfToImages() {
  await X.loadLib('jszip');
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
  await X.loadLib('jszip');
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

/* ───────── compress: pictures resampled to the resolution they are shown at, duplicates merged, waste removed — text and drawings untouched ───────── */
const LEVELS = { light: { dpi: 220, q: 0.85 }, balanced: { dpi: 150, q: 0.72 }, strong: { dpi: 110, q: 0.58 }, max: { dpi: 72, q: 0.45 } };
const mulM = (a, b) => [a[0] * b[0] + a[1] * b[2], a[0] * b[1] + a[1] * b[3], a[2] * b[0] + a[3] * b[2], a[2] * b[1] + a[3] * b[3], a[4] * b[0] + a[5] * b[2] + b[4], a[4] * b[1] + a[5] * b[3] + b[5]];
const deflate = async u8 => new Uint8Array(await new Response(new Blob([u8]).stream().pipeThrough(new CompressionStream('deflate'))).arrayBuffer());
const PN = n => PDFLib.PDFName.of(n);
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
function imageUse(doc) { // image object → the largest size (points) it is drawn at anywhere: { w, h }
  const L = PDFLib, use = new Map(), visits = new Map();
  const bytesOf = s => { try { return s instanceof L.PDFRawStream ? L.decodePDFRawStream(s).decode() : s.getContents(); } catch (e) { return null; } };
  const walk = (bytes, res, ctm0, depth) => {
    const xo = res instanceof L.PDFDict ? res.lookup(PN('XObject')) : null; if (!(xo instanceof L.PDFDict) || !bytes) return;
    let ops; try { ops = X.native.parseOps(bytes); } catch (e) { return; }
    let ctm = ctm0; const st = [];
    for (const o of ops) {
      if (o.op === 'q') st.push(ctm); else if (o.op === 'Q') ctm = st.length ? st.pop() : ctm0;
      else if (o.op === 'cm' && o.args.length === 6 && o.args.every(a => typeof a === 'number')) ctm = mulM(o.args, ctm);
      else if (o.op === 'Do' && o.args[0] && o.args[0].n) {
        const ref = xo.get(PN(o.args[0].n)); if (!(ref instanceof L.PDFRef)) continue;
        const s = doc.context.lookup(ref); if (!(s instanceof L.PDFStream)) continue; const sub = s.dict.lookup(PN('Subtype'));
        if (sub === PN('Image')) { const u = use.get(ref.tag) || { w: 0, h: 0 }; u.w = Math.max(u.w, Math.hypot(ctm[0], ctm[1])); u.h = Math.max(u.h, Math.hypot(ctm[2], ctm[3])); use.set(ref.tag, u); }
        else if (sub === PN('Form') && depth < 8) {
          const v = (visits.get(ref.tag) || 0) + 1; visits.set(ref.tag, v); if (v > 40) continue; // (a form drawn hundreds of times: its first uses tell enough)
          const m = s.dict.lookup(PN('Matrix')), mm = m instanceof L.PDFArray && m.size() === 6 ? m.asArray().map(x => (x instanceof L.PDFNumber ? x.asNumber() : 0)) : [1, 0, 0, 1, 0, 0];
          walk(bytesOf(s), s.dict.lookup(PN('Resources')) || res, mulM(mm, ctm), depth + 1);
        }
      }
    }
  };
  for (const page of doc.getPages()) {
    const node = page.node; let res = null; for (let n = node, k = 0; n && k < 32 && !res; n = n.lookup(PN('Parent')), k++) { const r = n.lookup(PN('Resources')); if (r instanceof L.PDFDict) res = r; }
    const c = node.lookup(PN('Contents')), list = c instanceof L.PDFArray ? c.asArray().map(r => doc.context.lookup(r)) : c ? [c] : [], parts = list.filter(s => s instanceof L.PDFStream).map(bytesOf);
    if (parts.some(x => !x)) continue;
    const all = new Uint8Array(parts.reduce((a, x) => a + x.length + 1, 0)); let o = 0; for (const x of parts) { all.set(x, o); o += x.length; all[o++] = 10; }
    walk(all, res, [1, 0, 0, 1, 0, 0], 0);
  }
  return use;
}
const isDefaultDecode = a => a.asArray().every((v, i) => v instanceof PDFLib.PDFNumber && v.asNumber() === (i % 2 ? 1 : 0));
function lastLayer(doc, obj) { // a stream's data under all but its last filter (e.g. a JPEG wrapped in ASCII85) → { filter, parms, bytes, peeled }, or null
  const L = PDFLib, d = obj.dict, f = d.lookup(PN('Filter')), p = d.lookup(PN('DecodeParms'));
  if (!(f instanceof L.PDFArray)) return { filter: f || null, parms: p, bytes: obj.contents, peeled: false };
  const n = f.size(), last = n ? f.lookup(n - 1) : null, parms = p instanceof L.PDFArray ? (n ? p.lookup(n - 1) : null) : p;
  if (n <= 1) return { filter: last, parms, bytes: obj.contents, peeled: false };
  if (p instanceof L.PDFArray && p.asArray().slice(0, -1).some(v => v instanceof L.PDFDict)) return null; // (outer layers with their own parameters: rare, left alone)
  try {
    const t = d.clone(doc.context); t.set(PN('Filter'), doc.context.obj(f.asArray().slice(0, -1))); t.delete(PN('DecodeParms'));
    return { filter: last, parms, bytes: L.decodePDFRawStream(L.PDFRawStream.of(t, obj.contents)).decode(), peeled: true };
  } catch (e) { return null; }
}
async function shrinkImage(doc, ref, obj, opt) { // opt { f: scale (≤ 1), q: JPEG quality, gray, mask: it is a soft mask (kept lossless) } → true when the picture was replaced by a smaller one
  const L = PDFLib, dict = obj.dict, num = k => { const v = dict.lookup(PN(k)); return v instanceof L.PDFNumber ? v.asNumber() : 0; };
  if (dict.lookup(PN('Subtype')) !== PN('Image') || dict.has(PN('ImageMask')) || dict.has(PN('Mask'))) return false;
  const dec = dict.lookup(PN('Decode')); if (dec && !(dec instanceof L.PDFArray && isDefaultDecode(dec))) return false;
  const w = num('Width'), h = num('Height'), f = Math.min(1, opt.f || 1); if (num('BitsPerComponent') !== 8 || w * h < 64 * 64 || obj.contents.length < 3000) return false;
  if (opt.mask && f >= 1) return false; // a transparency mask is only ever made smaller in size, never re-encoded with losses
  const ly = lastLayer(doc, obj); if (!ly) return false; const filter = ly.filter;
  const cs = dict.lookup(PN('ColorSpace')); let comps = cs === PN('DeviceRGB') || cs === PN('CalRGB') ? 3 : cs === PN('DeviceGray') || cs === PN('CalGray') ? 1 : 0;
  if (!comps && cs instanceof L.PDFArray && cs.lookup(0) === PN('ICCBased')) { const icc = cs.lookup(1), n = icc && icc.dict && icc.dict.lookup(PN('N')); comps = n instanceof L.PDFNumber ? n.asNumber() : 0; }
  if (!comps && opt.mask && !cs) comps = 1;
  if (comps !== 1 && comps !== 3) return false; // CMYK, indexed and other special colour spaces are left alone
  let src;
  if (filter === PN('DCTDecode')) src = await createImageBitmap(new Blob([ly.bytes], { type: 'image/jpeg' }), { colorSpaceConversion: 'none' });
  else if (filter === PN('FlateDecode') || !filter) {
    const parms = ly.parms, pv = parms instanceof L.PDFDict ? parms.lookup(PN('Predictor')) : null, pred = pv instanceof L.PDFNumber ? pv.asNumber() : 1;
    if (pred !== 1 && pred < 10) return false;
    let raw = filter ? L.decodePDFRawStream(obj).decode() : obj.contents; if (pred >= 10) raw = unpredict(raw, w, comps);
    if (raw.length < w * h * comps) return false;
    const id = new ImageData(w, h), d = id.data;
    for (let i = 0, j = 0, k = 0; i < w * h; i++, j += comps, k += 4) { d[k] = raw[j]; d[k + 1] = comps === 3 ? raw[j + 1] : raw[j]; d[k + 2] = comps === 3 ? raw[j + 2] : raw[j]; d[k + 3] = 255; }
    src = document.createElement('canvas'); src.width = w; src.height = h; src.getContext('2d').putImageData(id, 0, 0);
  } else return false;
  const c = document.createElement('canvas'); c.width = Math.max(1, Math.round(w * f)); c.height = Math.max(1, Math.round(h * f));
  const x = c.getContext('2d', { willReadFrequently: true }); x.fillStyle = '#fff'; x.fillRect(0, 0, c.width, c.height); x.imageSmoothingQuality = 'high'; x.drawImage(src, 0, 0, c.width, c.height);
  if (src.close) src.close();
  const set = (bytes, filterName, space) => {
    dict.set(PN('Width'), L.PDFNumber.of(c.width)); dict.set(PN('Height'), L.PDFNumber.of(c.height)); dict.set(PN('BitsPerComponent'), L.PDFNumber.of(8)); if (space) dict.set(PN('ColorSpace'), space);
    dict.set(PN('Filter'), PN(filterName)); dict.delete(PN('DecodeParms')); dict.delete(PN('Decode'));
    doc.context.assign(ref, L.PDFRawStream.of(dict, bytes)); return true;
  };
  if (opt.mask) { // soft mask: grey levels, compressed without loss
    const d = x.getImageData(0, 0, c.width, c.height).data, g = new Uint8Array(c.width * c.height); for (let i = 0; i < g.length; i++) g[i] = d[i * 4];
    const z = await deflate(g); return z.length < obj.contents.length * 0.9 ? set(z, 'FlateDecode', PN('DeviceGray')) : false;
  }
  if (opt.gray) X.applyFx(c, { mode: 'gray' });
  const bytes = new Uint8Array(await (await toBlob(c, 'image/jpeg', opt.q)).arrayBuffer());
  if (bytes.length > obj.contents.length * (f < 1 ? 0.95 : 0.85)) return false; // not worth it — keep the original
  return set(bytes, 'DCTDecode', comps === 3 && !opt.gray ? null : PN('DeviceRGB')); // (an RGB picture keeps its own colour profile)
}
function stripExtras(doc) { // page thumbnails, private editor data (Illustrator / InDesign PieceInfo) and per-object metadata — never shown, often large
  const L = PDFLib; let n = 0;
  if (doc.catalog.has(PN('PieceInfo'))) { doc.catalog.delete(PN('PieceInfo')); n++; }
  for (const [, o] of doc.context.enumerateIndirectObjects()) {
    const d = o instanceof L.PDFDict ? o : o instanceof L.PDFStream ? o.dict : null; if (!d) continue;
    const type = d.lookup(PN('Type')), sub = d.lookup(PN('Subtype')), keys = type === PN('Page') ? ['Thumb', 'PieceInfo'] : sub === PN('Image') || sub === PN('Form') ? ['PieceInfo', 'Metadata'] : [];
    for (const k of keys) if (d.has(PN(k))) { d.delete(PN(k)); n++; }
  }
  return n;
}
function dedupe(doc) { // identical streams (the same picture or font embedded again and again) → one shared copy
  const L = PDFLib, ctx = doc.context, groups = new Map(), map = new Map();
  for (const [ref, o] of ctx.enumerateIndirectObjects()) {
    if (!(o instanceof L.PDFRawStream) || o.contents.length < 64) continue;
    const key = o.dict.entries().filter(([k]) => k.toString() !== '/Length').map(([k, v]) => k.toString() + ' ' + v.toString()).sort().join(' ') + '|' + fnv(o.contents), g = groups.get(key);
    if (!g) { groups.set(key, { ref, o }); continue; }
    const a = g.o.contents, b = o.contents; let same = a.length === b.length; for (let i = 0; same && i < a.length; i++) if (a[i] !== b[i]) same = false;
    if (same) map.set(ref.tag, g.ref);
  }
  if (!map.size) return 0;
  const fix = v => (v instanceof L.PDFRef && map.has(v.tag) ? map.get(v.tag) : null);
  const walk = o => {
    if (o instanceof L.PDFDict) { for (const [k, v] of o.entries()) { const r = fix(v); if (r) o.set(k, r); else if (v instanceof L.PDFDict || v instanceof L.PDFArray) walk(v); } }
    else if (o instanceof L.PDFArray) { for (let i = 0; i < o.size(); i++) { const v = o.get(i), r = fix(v); if (r) o.set(i, r); else if (v instanceof L.PDFDict || v instanceof L.PDFArray) walk(v); } }
    else if (o instanceof L.PDFStream) walk(o.dict);
  };
  for (const [, o] of ctx.enumerateIndirectObjects()) walk(o);
  for (const [ref] of ctx.enumerateIndirectObjects()) if (map.has(ref.tag)) ctx.delete(ref);
  return map.size;
}
async function recompress(doc) { // streams stored without compression, or with weak old codecs (LZW, ASCII85, hex, run-length) → Flate
  const L = PDFLib, weak = [PN('LZWDecode'), PN('ASCII85Decode'), PN('ASCIIHexDecode'), PN('RunLengthDecode')]; let n = 0;
  for (const [ref, o] of doc.context.enumerateIndirectObjects()) {
    if (!(o instanceof L.PDFRawStream) || o.contents.length < 200) continue;
    const d = o.dict; if (d.lookup(PN('Type')) === PN('Metadata')) continue; // (XMP is meant to stay readable)
    let filter = d.lookup(PN('Filter'));
    if (filter instanceof L.PDFArray && filter.size() > 1) { // text / LZW layers around a real codec (e.g. ASCII85 + JPEG): taken off without touching the data inside
      const fs = filter.asArray().map((v, i) => filter.lookup(i)); if (!fs.slice(0, -1).every(v => weak.includes(v))) continue;
      const last = fs[fs.length - 1];
      if (!weak.includes(last)) {
        const ly = lastLayer(doc, o); if (!ly || !ly.peeled || ly.bytes.length >= o.contents.length) continue;
        d.set(PN('Filter'), last); if (ly.parms instanceof L.PDFDict) d.set(PN('DecodeParms'), ly.parms); else d.delete(PN('DecodeParms'));
        doc.context.assign(ref, L.PDFRawStream.of(d, ly.bytes)); n++; continue;
      }
      if (d.has(PN('DecodeParms'))) continue; filter = last; // (only weak layers: decoded fully and stored as Flate below)
    } else if (filter instanceof L.PDFArray) { if (filter.size() !== 1) continue; filter = filter.lookup(0); }
    if (filter && (!weak.includes(filter) || d.has(PN('DecodeParms')))) continue;
    let raw; try { raw = filter ? L.decodePDFRawStream(o).decode() : o.contents; } catch (e) { continue; }
    const z = await deflate(raw); if (z.length >= o.contents.length * 0.95) continue;
    d.set(PN('Filter'), PN('FlateDecode')); d.delete(PN('DecodeParms')); doc.context.assign(ref, L.PDFRawStream.of(d, z)); n++;
  }
  return n;
}
async function compressPdf(bytes, opt = {}) { // → { bytes (the smaller of original and result), before, after, pictures, dups, streams, extras }
  const L = PDFLib, o = { ...LEVELS.balanced, strip: true, ...opt }, before = bytes.length;
  let doc; try { doc = await L.PDFDocument.load(bytes, { updateMetadata: false }); }
  catch (e) { throw new Error(/encrypt/i.test(e.message) ? 'This PDF is password-protected — remove the password first, then compress it.' : 'Could not read this PDF: ' + e.message); }
  const st = { before, pictures: 0, dups: 0, streams: 0, extras: o.strip ? stripExtras(doc) : 0 };
  const merge = () => { for (let k = 0; k < 3; k++) { const n = dedupe(doc); st.dups += n; if (!n) break; } }; // (repeated: pictures sharing a merged mask become identical too)
  merge(); // first, so every picture is only resampled once
  let use = new Map(); try { if (X.native) use = imageUse(doc); } catch (e) { /* sizes unknown — the pixel-count rule below is used */ }
  const items = [], masks = new Map();
  for (const [ref, obj] of doc.context.enumerateIndirectObjects()) {
    if (!(obj instanceof L.PDFRawStream) || obj.dict.lookup(PN('Subtype')) !== PN('Image')) continue; items.push([ref, obj]);
    const sm = obj.dict.get(PN('SMask')); if (sm instanceof L.PDFRef) masks.set(sm.tag, ref.tag);
  }
  const scale = (tag, obj) => { // how much this picture can shrink and still have `dpi` pixels per inch where it is shown
    const num = k => { const v = obj.dict.lookup(PN(k)); return v instanceof L.PDFNumber ? v.asNumber() : 0; }, w = num('Width'), h = num('Height'), u = use.get(tag) || use.get(masks.get(tag));
    const f = u && u.w > 1 && u.h > 1 ? o.dpi / Math.min(w * 72 / u.w, h * 72 / u.h) : o.dpi * 12 / Math.max(w, h, 1); // (never seen drawn: assume up to a 12-inch picture)
    return f < 0.9 ? f : 1;
  };
  for (const [ref, obj] of items) {
    try { if (await shrinkImage(doc, ref, obj, { f: scale(ref.tag, obj), q: o.q, gray: !!o.gray, mask: masks.has(ref.tag) })) st.pictures++; } catch (e) { /* this picture stays as it is */ }
  }
  merge();
  try { st.streams = await recompress(doc); } catch (e) { /* CompressionStream missing — streams stay as they are */ }
  if (X.native) X.native.collectGarbage(doc);
  const out = await doc.save({ useObjectStreams: true, updateFieldAppearances: false });
  st.after = Math.min(out.length, before); st.bytes = out.length < before ? out : bytes;
  return st;
}
async function compress() {
  const r = await dialog('Compress PDF', `
    <p>Makes the file smaller: pictures are resampled to the resolution they are shown at and re-encoded, duplicate copies are merged and unused data is removed. Text and drawings stay sharp, selectable and searchable.</p>
    <label>Compression<select name="lvl">
      <option value="light">Light — print quality (pictures at 220 dpi)</option>
      <option value="balanced" selected>Balanced — screen and email (150 dpi)</option>
      <option value="strong">Strong — small file (110 dpi)</option>
      <option value="max">Maximum — smallest file (72 dpi, lower picture quality)</option></select></label>
    <label class="inl"><input type="checkbox" name="gray">Turn pictures grayscale (smaller still)</label>
    <label class="inl"><input type="checkbox" name="strip" checked>Remove hidden extras (page thumbnails, editing-program data)</label>
    <p class="mut">Saves a compressed copy; the document you have open is not changed.</p>`, { ok: 'Compress' });
  if (!r) return;
  await busy(async () => {
    const st = await compressPdf(await X.buildPdf(S.pages), { ...(LEVELS[r.lvl] || LEVELS.balanced), gray: !!r.gray, strip: !!r.strip }), mb = v => (v / 1048576).toFixed(2) + ' MB';
    download(new Blob([st.bytes], { type: 'application/pdf' }), X.baseName() + '-compressed.pdf');
    const what = [st.pictures && `${st.pictures} picture${st.pictures > 1 ? 's' : ''} resampled`, st.dups && `${st.dups} duplicate${st.dups > 1 ? 's' : ''} merged`].filter(Boolean).join(', ');
    toast(st.after < st.before ? `Compressed: ${mb(st.before)} → ${mb(st.after)} (${Math.max(1, Math.round((1 - st.after / st.before) * 100))}% smaller${what ? ' — ' + what : ''})` : `Already compact (${mb(st.before)}) — try a stronger level or grayscale`);
  }, 'Compressing…');
}

Object.assign(X.ACT, { compress, pdf2docx: pdfToDocx, pdf2img: pdfToImages, splitzip: splitZip, exporttxt: exportText, flatten });
X.NO_DOC.add('imgpages');
Object.assign(X, { docxToPages, pageContent, rowsOf, shrinkImage, compressPdf, buildDocx, COMPRESS_LEVELS: LEVELS, _editModel: p => editModel(p) });
Object.assign(X, { _cv: { samePage, blockXml, linesOf, blocksOf, wordFont, hasFont } }); // (used by pdf2word.js)
})();
