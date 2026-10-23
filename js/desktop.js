/* XD3 PDF Editor — Windows app extras: make XD3 the PDF app of the PC (Windows shows its own chooser on 8, 10 and 11),
   choose what a double-clicked PDF opens in (reading view or the editor), and a one-time offer on the start screen.
   Uses the small bridge in desktop/preload.js; on the website none of this appears. */
(() => {
'use strict';
const X = window.XD3, { S, $, $$, dialog, toast, icons, esc } = X, desk = window.xd3desk;
const store = { get: k => { try { return localStorage.getItem(k); } catch (e) { return null; } }, set: (k, v) => { try { localStorage.setItem(k, v); } catch (e) { /* not remembered */ } } };
X.idle = () => !S.pages.length && !S.dirty && !(X.docx && X.docx.isOpen()) && !(X.writing && X.writing.W.open); // the desktop shell opens the next double-clicked file in an idle window
X.shellView = () => store.get('xd3.shellview') !== 'edit'; // PDFs from File Explorer open in reading view unless the user prefers the editor
if (!desk) return;
X.desk = desk; document.documentElement.classList.add('desk');

const statusHtml = st => !st ? '<span class="mut">Not available on this computer.</span>'
  : st.mine ? '<b class="daok"><i data-icon="check"></i>XD3 PDF Editor opens your PDFs</b>'
  : `<span>PDFs now open with <b>${esc(st.app || 'another program')}</b></span>`;
let watch = 0;
function watchStatus(onChange) { // after Windows' chooser: look again whenever the window comes back into focus
  clearInterval(watch); let n = 0;
  const check = async () => { const st = await desk.pdfDefault().catch(() => null); if (st && st.mine) { clearInterval(watch); window.removeEventListener('focus', check); } onChange(st); if (++n > 40) clearInterval(watch); };
  window.addEventListener('focus', check); watch = setInterval(check, 3000);
}
async function makeDefault(onChange) {
  const r = await desk.makeDefault().catch(e => ({ error: e.message }));
  if (!r || r.error) return toast(r && r.error ? r.error : 'Windows did not allow the change', true);
  if (r.done) { toast('Done — PDFs now open with XD3 PDF Editor', 'success'); onChange(await desk.pdfDefault()); return; }
  toast('In the Windows window, choose “XD3 PDF Editor” and then “Always”', 'info'); watchStatus(onChange);
}
async function defaultDlg() {
  const [info, st] = await Promise.all([desk.info(), desk.pdfDefault().catch(() => null)]);
  const win7 = info.os.major === 6 && info.os.minor <= 1, view = X.shellView() ? 'read' : 'edit';
  await dialog('Default PDF app', `<div class="defapp">
      <div class="dastat" id="daStat">${statusHtml(st)}</div>
      <p>Make XD3 the program that opens PDFs on this PC: double-click any PDF and it opens here — fast, offline and private.</p>
      <div class="darow"><button type="button" class="btn primary" id="daMake"${!info.canRegister ? ' disabled' : ''}><i data-icon="check"></i><span>Make XD3 PDF Editor the default</span></button>
      <button type="button" class="btn line" id="daSet"><i data-icon="sliders"></i><span>Windows settings</span></button></div>
      <p class="mut">${win7 ? 'Windows 7 lets XD3 change this directly.' : `${esc(info.os.name)} asks you itself: in its window pick <b>XD3 PDF Editor</b>, then <b>Always</b> (or tick “Always use this app”). The test PDF it opens confirms the choice.`}</p>
      <h4>A PDF opened from File Explorer</h4>
      <label class="inl"><input type="radio" name="view" value="read"${view === 'read' ? ' checked' : ''}>Opens in reading view — click <b>Edit</b> to change it</label>
      <label class="inl"><input type="radio" name="view" value="edit"${view === 'edit' ? ' checked' : ''}>Opens ready to edit</label>
      ${info.portable ? '<p class="mut">This is the portable copy: Windows remembers the folder it runs from, so set it again if you move the folder.</p>' : ''}
    </div>`, { ok: 'Done', cancel: false, onOpen: d => {
      icons(d); const upd = s => { const el = $('#daStat', d); if (el) { el.innerHTML = statusHtml(s); icons(el); } };
      $('#daMake', d).onclick = () => makeDefault(upd);
      $('#daSet', d).onclick = () => { desk.openSettings(); watchStatus(upd); };
      $$('input[name=view]', d).forEach(r => r.onchange = () => store.set('xd3.shellview', r.value));
    }, onClose: () => clearInterval(watch) });
}
X.ACT.defapp = defaultDlg; X.NO_DOC.add('defapp');

// the ⋯ menu gets the entry; the website's "Download for Windows" makes no sense inside the app
const more = $('#mMore'), dl = more && $('[data-act=dlwin]', more);
if (dl) { const b = document.createElement('button'); b.dataset.act = 'defapp'; b.innerHTML = '<i data-icon="windows"></i>Default PDF app…'; dl.replaceWith(b); icons(b); }

// a quiet offer on the start screen until the user answers it
(async () => {
  const ask = store.get('xd3.defask'); if (ask === 'never' || (ask && Date.now() - +ask < 14 * 864e5)) return;
  const info = await desk.info().catch(() => null); if (!info || !info.canRegister) return;
  const st = await desk.pdfDefault().catch(() => null); if (!st || st.mine) return;
  const hero = $('#welcome .hero'), bar = document.createElement('div'); bar.className = 'defbar';
  bar.innerHTML = `<i data-icon="windows"></i><div><b>Open PDFs with XD3 PDF Editor?</b><span>Make it this PC's PDF app — double-click a PDF to read it, one click to edit.</span></div>
    <button class="btn primary" data-dd="make">Make default</button><button class="btn line" data-dd="later">Not now</button><button class="btn line" data-dd="never">Don't ask again</button>`;
  hero.insertBefore(bar, $('.dropzone', hero)); icons(bar);
  bar.addEventListener('click', e => {
    const b = e.target.closest('[data-dd]'); if (!b) return; const a = b.dataset.dd;
    if (a === 'make') makeDefault(s => { if (s && s.mine) bar.remove(); });
    else { store.set('xd3.defask', a === 'never' ? 'never' : String(Date.now())); bar.remove(); }
  });
})();
})();
