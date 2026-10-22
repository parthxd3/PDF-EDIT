/* XD3 PDF Editor — filling in PDF forms. Fields are outlined in Edit text mode; what you type becomes page content and the empty field is dropped when saving. */
(() => {
'use strict';
const X = window.XD3, { S, toast, dialog, clamp, esc } = X;

async function formItems(p) { // the form fields on a page, as clickable items (page view space); fields already filled in are left out
  if (p.src < 0) return [];
  const s = S.sources[p.src]; s.forms = s.forms || {};
  if (!s.forms[p.idx]) s.forms[p.idx] = (async () => {
    const out = [];
    try {
      const pg = await X.getPdfPage(p), vp = pg.getViewport({ scale: 1 });
      for (const a of await pg.getAnnotations()) {
        if (a.subtype !== 'Widget' || a.hidden || !a.rect || !a.fieldType || a.fieldType === 'Sig' || a.readOnly) continue;
        const r = vp.convertToViewportRectangle(a.rect), x = Math.min(r[0], r[2]), y = Math.min(r[1], r[3]), w = Math.abs(r[2] - r[0]), h = Math.abs(r[3] - r[1]);
        if (w < 4 || h < 4) continue;
        const m = /([\d.]+)\s+Tf/.exec(a.defaultAppearance || ''), da = m ? +m[1] : 0;
        out.push({ field: true, ft: a.fieldType, x, y, w, h, str: a.fieldName || 'Form field', name: a.fieldName || '', value: typeof a.fieldValue === 'string' ? a.fieldValue : Array.isArray(a.fieldValue) ? a.fieldValue[0] || '' : '',
          check: !!a.checkBox, radio: !!a.radioButton, on: a.exportValue || a.buttonValue || 'Yes', multi: !!a.multiLine, options: (a.options || []).map(o => o.displayValue || o.exportValue), da, rect: a.rect.map(v => Math.round(v * 100) / 100) });
      }
    } catch (e) { console.warn('form fields', e); }
    return out;
  })();
  const all = await s.forms[p.idx], done = p.filled || [];
  return all.filter(f => !done.some(r => r.every((v, i) => Math.abs(v - f.rect[i]) < 0.6)));
}
async function fillField(p, f) { // click on a field → a text object sized to it, ready for typing (or a tick / dot for boxes)
  const before = X.snap(); let text = f.value || '', size = f.da || 0;
  if (f.check || f.radio) text = f.radio ? '●' : '✓';
  else if (f.options && f.options.length) {
    const r = await dialog(esc(f.name || 'Choose'), `<label>Value<select name="v">${f.options.map(o => `<option${o === text ? ' selected' : ''}>${esc(o)}</option>`).join('')}</select></label>`, { ok: 'Fill in' });
    if (!r) return; text = r.v;
  }
  if (!size) size = f.check || f.radio ? f.h * 0.75 : f.multi ? 10 : clamp(f.h * 0.6, 6, 14);
  size = clamp(size, 4, 72);
  const o = X.textObj(text, +size.toFixed(1), { color: '#000000', font: 'Helvetica' });
  if (f.check || f.radio) { const b = X.bbox(o); o.x = f.x + (f.w - b.w) / 2; o.y = f.y + (f.h - b.h) / 2; }
  else if (f.multi) { o.x = f.x + 2; o.y = f.y + 1; }
  else { o.x = f.x + 2; o.y = f.y + f.h / 2 + size * 0.35 - X.baseOff(o); } // sits on the field's centre line
  o.field = { name: f.name, kind: f.check ? 'check' : f.radio ? 'radio' : f.options && f.options.length ? 'choice' : 'text', on: f.on, rect: f.rect }; // (saved into the real field when the file is edited in place)
  p.filled = [...(p.filled || []), f.rect]; p.objs.push(o);
  if (f.check || f.radio) { X.pushUndo(before); X.drawObjs(p); X.drawTI(p); }
  else X.startEdit(p, o, before, !!text);
}
function hideFields(page, p) { // export: drop the widgets whose value is now page content, so readers do not draw an empty box over it
  const L = PDFLib, N = L.PDFName.of, annots = page.node.Annots(); if (!annots) return;
  const keep = [], ctx = page.doc.context;
  for (let i = 0; i < annots.size(); i++) {
    const ref = annots.get(i), d = ctx.lookup(ref); let drop = false;
    try {
      const rect = d && d.lookup && d.lookup(N('Rect')); const v = rect && rect.asArray ? rect.asArray().map(n => n.asNumber()) : null;
      if (v && v.length === 4) { const nv = [Math.min(v[0], v[2]), Math.min(v[1], v[3]), Math.max(v[0], v[2]), Math.max(v[1], v[3])]; drop = p.filled.some(r => { const nr = [Math.min(r[0], r[2]), Math.min(r[1], r[3]), Math.max(r[0], r[2]), Math.max(r[1], r[3])]; return nr.every((q, k) => Math.abs(q - nv[k]) < 0.6); }); }
    } catch (e) { /* keep it */ }
    if (!drop) keep.push(ref);
  }
  page.node.set(N('Annots'), ctx.obj(keep));
}
Object.assign(X, { formItems, fillField, hideFields });
})();
