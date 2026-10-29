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
  const finish = async (body, title) => {
    const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" mc:Ignorable="w14"';
    const H = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>', W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"', CT = 'application/vnd.openxmlformats-officedocument', now = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
    put('[Content_Types].xml', `${H}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Default Extension="jpeg" ContentType="image/jpeg"/><Override PartName="/word/document.xml" ContentType="${CT}.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="${CT}.wordprocessingml.styles+xml"/><Override PartName="/word/settings.xml" ContentType="${CT}.wordprocessingml.settings+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>`);
    put('_rels/.rels', `${H}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>`);
    put('docProps/core.xml', `${H}<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>${xesc(title || '')}</dc:title><dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified></cp:coreProperties>`);
    put('word/_rels/document.xml.rels', `${H}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>${rels.join('')}</Relationships>`);
    put('word/styles.xml', `${H}<w:styles ${W}><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="en-US"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style></w:styles>`);
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
async function samePage(p, B) {
  const [PW, PH] = X.rdims(p), sc = Math.min(1, 1584 / Math.max(PW, PH)), K = DK, rot = (((p.rot || 0) % 360) + 360) % 360; // (Word pages are at most 22 inches)
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
  // the text
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
  for (const q of pieces) { // colour of the text, from the difference between the page with and without it; text that makes no difference is not shown by the PDF either
    if (!q.sample) continue;
    const pts = [[q.x, q.base - 0.8 * q.size], [q.x + q.w, q.base - 0.8 * q.size], [q.x, q.base + 0.2 * q.size], [q.x + q.w, q.base + 0.2 * q.size]].map(v => place(q, v[0], v[1]));
    const xa = clamp(Math.floor(Math.min(...pts.map(v => v[0])) * K), 0, W - 1), xb = clamp(Math.ceil(Math.max(...pts.map(v => v[0])) * K), 0, W - 1), ya = clamp(Math.floor(Math.min(...pts.map(v => v[1])) * K), 0, H - 1), yb = clamp(Math.ceil(Math.max(...pts.map(v => v[1])) * K), 0, H - 1);
    let md = 0; const hits = [];
    for (let y = ya; y <= yb; y++) for (let x = xa, i = (y * W + xa) * 4; x <= xb; x++, i += 4) { const dd = Math.abs(rd[i] - d[i]) + Math.abs(rd[i + 1] - d[i + 1]) + Math.abs(rd[i + 2] - d[i + 2]); if (dd > 40) hits.push(dd, i); if (dd > md) md = dd; }
    if (md < 60) { q.ghost = true; continue; }
    let r = 0, g = 0, bl = 0, n = 0; for (let k = 0; k < hits.length; k += 2) if (hits[k] >= 0.8 * md) { const i = hits[k + 1]; r += rd[i]; g += rd[i + 1]; bl += rd[i + 2]; n++; }
    const c = [r / n, g / n, bl / n].map(Math.round); q.color = Math.max(...c) < 70 ? '#000000' : '#' + c.map(v => v.toString(16).padStart(2, '0')).join('');
  }
  const groups = new Map(); // text set at the same angle around the same point is laid out together
  for (const q of pieces) { const k = q.rot ? q.rot + '|' + q.pv.map(v => v.toFixed(1)).join(',') : '0'; if (!groups.has(k)) groups.set(k, []); groups.get(k).push(q); }
  let boxes = '';
  for (const g of groups.values()) for (const b of blocksOf(linesOf(g))) boxes += blockXml(b, B, place, rot, sc);
  const m = 18 * sc;
  return { w: PW * sc, h: PH * sc, margins: { t: m, l: m, r: m, b: m }, paras: [{ ppr: '<w:spacing w:before="0" w:after="0" w:line="20" w:lineRule="exact"/><w:rPr><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr>', body: back + boxes }] };
}
async function buildDocx(pages, mode = 'same') { // → the .docx as a Blob
  await X.loadLib('jszip'); const B = docxBuilder(); let body = '';
  for (let i = 0; i < pages.length; i++) {
    const p = pages[i]; let out, w = p.w, h = p.h;
    if (mode === 'pic') { [w, h] = X.rdims(p); const c = await X.renderPageImage(p, 150 / 72); out = { paras: [{ ppr: '<w:spacing w:before="0" w:after="0"/>', body: B.pic(c.toDataURL('image/jpeg', 0.85), w, h, { x: 0, y: 0 }) }], margins: { t: 10, l: 10, r: 10, b: 10 } }; }
    else if (mode === 'flow') out = flowPage(await pageContent(p), p, B);
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
  const r = await dialog('PDF → Word (.docx)', `
    <label>Layout<select name="mode">
      <option value="same">Same as the PDF — exact look, editable text (recommended)</option>
      <option value="flow">Flowing text — plain paragraphs that re-wrap, layout approximated</option>
      <option value="pic">Picture — each page as an image (not editable)</option></select></label>
    <label>Pages (blank = all), e.g. 1-3, 5<input type="text" name="range" autocomplete="off"></label>
    <p class="mut">Includes your edits. “Same as the PDF” keeps every line, table border, colour and picture exactly where it is, and puts the text in editable text boxes in the same fonts. Scanned pages keep their look; run OCR first to make their text selectable.</p>`, { ok: 'Convert' });
  if (!r) return; const pages = pickPages(r.range); if (!pages) return;
  await busy(async () => { download(await buildDocx(pages, r.mode), X.baseName() + '.docx'); toast('Word document saved to your downloads'); }, 'Converting to Word…');
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
  await X.loadLib('jszip');
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
Object.assign(X, { docxToPages, pageContent, rowsOf, shrinkImage, compressPdf, buildDocx, COMPRESS_LEVELS: LEVELS });
})();
