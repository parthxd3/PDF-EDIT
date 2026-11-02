/* XD3 PDF Editor — the home screen: every feature card starts its tool (asking for a file first when the tool needs one),
   and the writing assistant checks spelling and grammar, auto-corrects, or humanizes pasted text right from the start page. */
(() => {
'use strict';
const X = window.XD3, { S, $, toast, mi } = X;

function pickFiles(accept) { // → the chosen files ([] when the picker is closed)
  return new Promise(res => { const i = document.createElement('input'); i.type = 'file'; i.accept = accept; i.onchange = () => res([...i.files]); i.addEventListener('cancel', () => res([])); i.click(); });
}
async function withDoc(run) { // a tool that works on an open document: with none open, the file is asked for first, then the tool starts
  if (!S.pages.length) {
    const files = await pickFiles('.pdf,application/pdf,.docx,image/*'); if (!files.length) return;
    await X.openFiles(files, false); if (!S.pages.length) return;
  }
  await run();
}
const MENUS = {
  convert: () => [mi('PDF → Word (.docx)…', 'word', 'h_need', { d: { a: 'pdf2docx' } }), mi('Word → PDF…', 'word', 'docx2pdf'), '-',
    mi('PDF → images…', 'image', 'h_need', { d: { a: 'pdf2img' } }), mi('Images → PDF…', 'image', 'imgpages'), '-', mi('Export text (.txt)', 'text', 'h_need', { d: { a: 'exporttxt' } })],
  protect: () => [mi('Password protect…', 'lock', 'h_need', { d: { a: 'protect' } }), mi('Black out sensitive text', 'redact', 'h_tool', { d: { t: 'redact' } }), mi('Flatten (permanent redaction)…', 'layers', 'h_need', { d: { a: 'flatten' } })],
  scan: () => [mi('Scan with the camera or photos', 'scan', 'scan'), mi('Make a PDF searchable (OCR)…', 'scan', 'h_need', { d: { a: 'ocrall' } }), '-',
    mi('Compare two PDFs…', 'swap', 'h_need', { d: { a: 'compare' } }), mi('Batch process many PDFs…', 'layers', 'batch')],
  sign: () => [mi('Sign…', 'sign', 'h_need', { d: { a: 'sign' } }), mi("Stamp today's date", 'stamp', 'h_need', { d: { a: 'datestamp' } }), '-',
    mi('Highlight', 'marker', 'h_tool', { d: { t: 'marker' } }), mi('Sticky note', 'note', 'h_tool', { d: { t: 'note' } }), mi('Arrows and shapes', 'arrow', 'h_tool', { d: { t: 'arrow' } })],
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

Object.assign(X.ACT, {
  h_need: b => withDoc(() => X.ACT[b.dataset.a](b)),
  h_tool: b => withDoc(() => X.setTool(b.dataset.t)),
  h_menu: b => { const r = b.getBoundingClientRect(); X.showCtx({ clientX: r.left + 14, clientY: r.bottom - 8, target: b }, MENUS[b.dataset.m]()); },
  h_wcheck: () => toWriter('issues'),
  h_wfix: () => toWriter('issues', true),
  h_whuman: () => toWriter('human'),
  h_wfile: async () => { const [f] = await pickFiles('.docx,.txt,.md,text/plain'); if (!f) return; X.writing.show(); await X.writing.openFile(f); },
  h_wpdf: () => withDoc(async () => { X.writing.show(); await X.writing.ACTS.pdf(); }),
});
['h_need', 'h_tool', 'h_menu', 'h_wcheck', 'h_wfix', 'h_whuman', 'h_wfile', 'h_wpdf'].forEach(a => X.NO_DOC.add(a));
X.home = { MENUS, withDoc };
})();
