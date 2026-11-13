/* XD3 PDF Editor — interface polish: click ripples, theme + accent colours, the notification panel and the command palette. */
(() => {
'use strict';
const X = window.XD3, { S, $, $$, esc, icons } = X, root = document.documentElement;
const store = { get: k => { try { return localStorage.getItem(k); } catch (e) { return null; } }, set: (k, v) => { try { localStorage.setItem(k, v); } catch (e) { /* storage unavailable */ } } };

/* ───────── ripple on press ───────── */
document.addEventListener('pointerdown', e => {
  const b = e.target.closest && e.target.closest('.btn, .tool, .tgl, .tab, .menu button, .chips button'); if (!b || b.disabled) return;
  const r = b.getBoundingClientRect(), d = Math.max(r.width, r.height), s = document.createElement('span');
  s.className = 'rip'; s.style.cssText = `width:${d}px;height:${d}px;left:${e.clientX - r.left - d / 2}px;top:${e.clientY - r.top - d / 2}px`;
  b.appendChild(s); setTimeout(() => s.remove(), 600);
}, true);

/* ───────── theme + accent ───────── */
function setAccent(v) { if (v && v !== 'red') root.dataset.accent = v; else delete root.dataset.accent; $$('.swatches button').forEach(b => b.classList.toggle('on', b.dataset.v === (v || 'red'))); }
function animateTheme() { root.classList.add('theming'); clearTimeout(animateTheme.t); animateTheme.t = setTimeout(() => root.classList.remove('theming'), 450); } // cross-fade instead of a hard switch
const themeIcon = () => { // the button shows the theme it switches to
  const dark = root.dataset.theme === 'dark', meta = $('meta[name="theme-color"]'); if (meta) meta.content = dark ? '#24262e' : '#ffffff';
  $$('[data-act="theme"]').forEach(b => { const i = $('i[data-icon]', b); if (i) { i.dataset.icon = dark ? 'sun' : 'moon'; i.textContent = ''; icons(b); } if (b.closest('#topbar')) b.title = dark ? 'Switch to the light theme' : 'Switch to the dark theme'; });
};
X.ACT.theme = () => { animateTheme(); root.dataset.theme = root.dataset.theme === 'dark' ? 'light' : 'dark'; store.set('xd3.theme', root.dataset.theme); themeIcon(); };
X.ACT.accent = b => { animateTheme(); setAccent(b.dataset.v); store.set('xd3.accent', b.dataset.v); };
setAccent(store.get('xd3.accent')); themeIcon();

/* ───────── hosted copy: works offline after the first visit and can be installed like an app; ?selftest=1 runs the checks ───────── */
const params = new URLSearchParams(location.search);
if (params.has('selftest')) { const s = document.createElement('script'); s.src = 'tests/selftest.js'; document.head.appendChild(s); }
else if ('serviceWorker' in navigator && /^https?:$/.test(location.protocol)) {
  if (/^(localhost|127\.0\.0\.1|\[::1\])$/.test(location.hostname)) navigator.serviceWorker.getRegistrations().then(rs => rs.forEach(r => r.unregister())).catch(() => {}); // a copy run on this computer is always current — no offline cache to go stale
  else {
    navigator.serviceWorker.register('sw.js').catch(() => { /* not available on this host */ });
    // once the editor has loaded, the service worker quietly keeps the rest of the fonts and the OCR engine for offline use (skipped on a metered "save data" connection)
    const fontFiles = () => { const B = X.BUNDLED || {}, out = []; for (const n in B) { out.push(`lib/fonts/${B[n][0]}-Regular.ttf`); if (B[n][1]) out.push(`lib/fonts/${B[n][0]}-Bold.ttf`); } return [...new Set(out)]; };
    const ocrFiles = () => (X.OCR_FILES ? X.OCR_FILES() : []);
    const warm = (urls, first) => navigator.serviceWorker.ready.then(r => { if (r.active && urls.length) r.active.postMessage({ type: 'warm', urls, first: !!first }); }).catch(() => {});
    const saveData = () => !!(navigator.connection && navigator.connection.saveData);
    const later = () => setTimeout(() => { if (saveData()) return; const o = ocrFiles(); warm([...o.slice(0, 5), ...fontFiles(), ...o.slice(5)]); }, 6000);
    if (document.readyState === 'complete') later(); else window.addEventListener('load', later, { once: true });
    // a tool that needs them: its files go to the front of the queue as soon as it is opened while online
    const need = () => { const h = location.hash; if (/ocr/.test(h)) warm(ocrFiles(), true); else if (/word-to-pdf|docx|word-editor/.test(h)) warm(fontFiles(), true); };
    window.addEventListener('hashchange', need); need();
  }
}

/* ───────── notification panel (the bell) ───────── */
const ICON = { success: 'check', error: 'x', warn: 'info', info: 'info' };
const ago = t => { const s = Math.round((Date.now() - t) / 1000); return s < 45 ? 'now' : s < 3600 ? Math.round(s / 60) + ' min' : Math.round(s / 3600) + ' h'; };
function renderNotes() {
  const m = $('#mNotes');
  m.innerHTML = '<div class="mhead">Notifications</div>' + (S.notes.length
    ? S.notes.map((n, i) => `<div class="note t-${n.type}" style="animation-delay:${Math.min(i, 8) * 25}ms"><i data-icon="${ICON[n.type]}"></i><div>${esc(n.msg)}</div><small>${ago(n.t)}</small></div>`).join('') + '<hr><button data-act="notesclear"><i data-icon="trash"></i><span>Clear all</span></button>'
    : '<div class="empty">Nothing yet — messages about saves, uploads and errors are kept here.</div>');
  icons(m); S.unread = 0; X.updateBell();
}
$('#bBell').addEventListener('click', renderNotes); // runs before the menu opens, so the list is fresh
X.ACT.notesclear = () => { S.notes.length = 0; S.unread = 0; X.updateBell(); };

/* ───────── command palette (Ctrl+K) ───────── */
const pal = $('#palette'), input = $('#palInput'), list = $('#palList'); let cmds = [], shown = [], at = 0;
function collect() { // every button in the bars and menus becomes a searchable command
  const seen = new Set(), out = [], editing = document.body.classList.contains('hasdoc') && !document.body.classList.contains('athome');
  for (const el of $$('#topbar [data-act], #toolbar [data-act], #toolbar [data-tool], #status [data-act], .tbm [data-act], .tbm [data-tool], #mMore [data-act], #mArrange [data-act], #mSave [data-act]')) {
    if (el.closest('.brand, .hn') || el.dataset.act === 'palette' || el.dataset.act === 'readertog' || el.hidden) continue; // (the header shortcuts repeat the tool menu)
    if (editing && el.closest('.mega, #navSheet')) continue; // in the editor, commands work on the open file; the home tool cards would leave it
    if (!S.pages.length && (el.dataset.tool || (el.dataset.act && !X.NO_DOC.has(el.dataset.act)))) continue; // with no file open these would only say "Open a PDF first": the tool pages ("Add watermark"…) are offered instead
    const key = (el.dataset.act || 'tool:' + el.dataset.tool) + '|' + (el.dataset.t || el.dataset.m || el.dataset.v || ''); if (seen.has(key)) continue; seen.add(key);
    const span = $$('span', el).find(s => !s.closest('.tico')), title = el.title || '', text = (span ? span.textContent : el.textContent).replace(/[▾\s]+$/, '').trim();
    const sc = (title.match(/\(([^()]{1,10})\)\s*$/) || [])[1] || '', plain = title.replace(/\s*\([^()]{1,10}\)\s*$/, '');
    let name = text || plain; if (el.dataset.act === 'stamp') name = 'Stamp: ' + name;
    const head = el.closest('#mSave') && headOf(el); if (head) name = head + ': ' + name; // "Export as: Word (.docx)", "Save a copy: Compressed"
    out.push({ el, name, hint: plain && plain !== name ? plain : '', sc, icon: ($('i[data-icon]', el) || { dataset: {} }).dataset.icon || 'bolt' });
  }
  return out;
}
const headOf = el => { for (let p = el.previousElementSibling; p; p = p.previousElementSibling) if (p.classList.contains('mhead')) return p.textContent.trim(); return ''; };
function draw() {
  const q = input.value.toLowerCase().split(/\s+/).filter(Boolean);
  const rank = c => { const n = c.name.toLowerCase(); return n.startsWith(q[0]) ? 2 : (' ' + n).includes(' ' + q[0]) ? 1 : 0; }; // the name starts with the word, then a word in the name does
  shown = cmds.filter(c => { const hay = (c.name + ' ' + c.hint).toLowerCase(); return q.every(w => hay.includes(w)); })
    .map((c, i) => [c, q.length ? rank(c) : 0, i]).sort((a, b) => b[1] - a[1] || a[2] - b[2]).map(x => x[0]).slice(0, 200);
  at = Math.min(at, Math.max(0, shown.length - 1));
  list.innerHTML = shown.length ? shown.map((c, i) => `<div class="pitem${i === at ? ' on' : ''}" data-i="${i}"><i data-icon="${c.icon}"></i><b>${esc(c.name)}</b><span>${esc(c.hint)}</span>${c.sc ? `<small><kbd>${esc(c.sc)}</kbd></small>` : ''}</div>`).join('') : '<div class="empty">No matching command</div>';
  icons(list); const on = $('.pitem.on', list); if (on) on.scrollIntoView({ block: 'nearest' });
}
function openPal() { X.commitEdit(); cmds = collect(); at = 0; input.value = ''; pal.hidden = false; document.body.classList.add('modal'); draw(); input.focus(); }
function closePal() { pal.hidden = true; document.body.classList.remove('modal'); }
function run(i) { const c = shown[i]; if (!c) return; closePal(); c.el.click(); }
X.ACT.palette = openPal;
input.addEventListener('input', () => { at = 0; draw(); });
input.addEventListener('keydown', e => {
  e.stopPropagation();
  if (e.key === 'ArrowDown') { e.preventDefault(); at = (at + 1) % Math.max(1, shown.length); draw(); }
  else if (e.key === 'ArrowUp') { e.preventDefault(); at = (at - 1 + shown.length) % Math.max(1, shown.length); draw(); }
  else if (e.key === 'Enter') { e.preventDefault(); run(at); }
  else if (e.key === 'Escape') closePal();
});
list.addEventListener('mousemove', e => { const it = e.target.closest('.pitem'); if (it && +it.dataset.i !== at) { at = +it.dataset.i; $$('.pitem', list).forEach((p, i) => p.classList.toggle('on', i === at)); } });
list.addEventListener('click', e => { const it = e.target.closest('.pitem'); if (it) run(+it.dataset.i); });
pal.addEventListener('mousedown', e => { if (e.target === pal) closePal(); });
window.addEventListener('keydown', e => {
  if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'k' && !document.body.classList.contains('dxopen')) { e.preventDefault(); e.stopImmediatePropagation(); if (pal.hidden) { if (!$('#dlg').open) openPal(); } else closePal(); }
}, true);

/* ───────── guidance: the tip for the current tool sits in the toolbar (where people look), not only in the status bar ───────── */
const hint = $('#hint'), ctxhint = $('#ctxhint'), copyHint = () => { ctxhint.textContent = hint.textContent; ctxhint.title = hint.textContent; };
new MutationObserver(copyHint).observe(hint, { childList: true, characterData: true, subtree: true }); copyHint();
// the toolbar rows never wrap: when they are wider than the window, the mouse wheel scrolls them sideways
// a row wider than the window fades out at the edge that has more
const rows = $$('.tools, #props'), fade = row => { const over = row.scrollWidth > row.clientWidth + 2, set = (c, v) => { if (row.classList.contains(c) !== v) row.classList.toggle(c, v); }; set('fade-l', over && row.scrollLeft > 2); set('fade-r', over && row.scrollLeft + row.clientWidth < row.scrollWidth - 2); }; // (only real changes, so the observer below never loops)
const fadeAll = () => requestAnimationFrame(() => rows.forEach(fade));
for (const row of rows) { row.addEventListener('scroll', () => fade(row), { passive: true }); new MutationObserver(fadeAll).observe(row, { attributes: true, subtree: true, attributeFilter: ['class', 'data-ctx', 'disabled', 'hidden'] }); }
new MutationObserver(fadeAll).observe(document.body, { attributes: true, attributeFilter: ['class', 'data-mode'] }); window.addEventListener('resize', fadeAll); fadeAll();
for (const row of rows) row.addEventListener('wheel', e => { if (row.scrollWidth > row.clientWidth + 2 && Math.abs(e.deltaY) > Math.abs(e.deltaX)) { row.scrollLeft += e.deltaY; e.preventDefault(); } }, { passive: false });

/* ───────── touch screens: press and hold a button to see what it does ───────── */
let holdT = 0, tip = null;
const hideTip = () => { clearTimeout(holdT); if (tip) { tip.remove(); tip = null; } };
document.addEventListener('touchstart', e => {
  hideTip(); const b = e.target.closest && e.target.closest('button[title], [data-tool][title]'); if (!b) return;
  holdT = setTimeout(() => {
    const r = b.getBoundingClientRect(); tip = document.createElement('div'); tip.className = 'holdtip'; tip.textContent = b.title.replace(/\s*\([^()]{1,12}\)\s*/, ' ').trim(); document.body.appendChild(tip);
    tip.style.left = Math.max(8, Math.min(innerWidth - tip.offsetWidth - 8, r.left + r.width / 2 - tip.offsetWidth / 2)) + 'px'; tip.style.top = (r.bottom + 8 + tip.offsetHeight > innerHeight ? r.top - tip.offsetHeight - 8 : r.bottom + 8) + 'px';
    setTimeout(hideTip, 2600);
  }, 450);
}, { passive: true });
document.addEventListener('touchmove', () => clearTimeout(holdT), { passive: true });
document.addEventListener('touchend', () => clearTimeout(holdT), { passive: true });

/* ───────── first time a PDF is opened: three short tips ───────── */
const SEEN = 'xd3.tips1', seen = () => { try { return localStorage.getItem(SEEN) === '1'; } catch (e) { return true; } };
function firstTips(force) {
  if ((!force && seen()) || $('#coach') || /selftest/.test(location.search)) return;
  const c = document.createElement('div'); c.id = 'coach'; c.setAttribute('role', 'dialog'); c.setAttribute('aria-label', 'Quick tips');
  c.innerHTML = `<b>Quick tips</b><ol><li><b>Click any text</b> to change it — it keeps its own font.</li><li>The tool row has the rest — <b>Pages ▾</b> and <b>More ▾</b> hold every other tool, and <b>Ctrl K</b> finds any of them.</li><li>Click <b>Download</b> when you are done — your file never leaves this device.</li></ol><button class="btn primary" type="button">Got it</button>`;
  document.body.appendChild(c);
  $('button', c).onclick = () => { try { localStorage.setItem(SEEN, '1'); } catch (e) { /* shown again next time */ } c.remove(); };
}
document.addEventListener('xd3:opened', () => setTimeout(firstTips, 600));
X.ACT.tips = () => firstTips(true); X.NO_DOC.add('tips');

/* ───────── More menu: each switch shows whether its option is on, every time the menu opens ───────── */
const more = $('#mMore'), body = document.body;
const TOG = { reader: () => body.classList.contains('reader'), sidebar: () => (innerWidth > 980 ? !body.classList.contains('nosb') : body.classList.contains('sbopen')), spread: () => body.classList.contains('spread'), night: () => body.classList.contains('night'),
  fullscreen: () => !!document.fullscreenElement, annots: () => { try { return localStorage.getItem('xd3.annots') !== 'off'; } catch (e) { return true; } } };
const syncTog = () => $$('[data-tog]', more).forEach(b => b.classList.toggle('on', TOG[b.dataset.tog]()));
new MutationObserver(() => { if (more.classList.contains('open')) syncTog(); }).observe(more, { attributes: true, attributeFilter: ['class'] });
X.ACT.readertog = () => X.ACT.reader(); // (its own name, so reading view's "Edit" label stays on the header button only)
more.addEventListener('click', e => { if (e.target.closest('[data-tog]')) { setTimeout(() => { more.classList.add('open'); syncTog(); }, 0); setTimeout(syncTog, 350); } }); // a switch flips in place: the menu stays open

/* ───────── tool row: a menu button lights up (and shows the tool's icon) while one of its tools is in use ───────── */
const grpIcon = (b, name) => { const i = $('i[data-icon]', b); if (i && i.dataset.icon !== name) { i.dataset.icon = name; i.textContent = ''; icons(b); } };
function syncGroups() {
  const m = body.dataset.mode;
  $$('.grpbtn[data-tools]').forEach(b => {
    const on = b.dataset.tools.split(' ').includes(m); b.classList.toggle('active', on);
    const src = on && $(`#${b.dataset.menu} [data-tool="${m}"] i[data-icon]`); if (src && /^m(Mark|Shapes)$/.test(b.dataset.menu)) grpIcon(b, src.dataset.icon);
  });
}
new MutationObserver(syncGroups).observe(body, { attributes: true, attributeFilter: ['data-mode'] }); syncGroups();
// an item that opens its own menu (Letterhead) opens it under the tool-row button, since its own menu has just closed
document.addEventListener('click', e => {
  const b = e.target.closest && e.target.closest('.tbm [data-act="letterhead"]'); if (!b) return;
  const menu = b.closest('.tbm'), trg = $(`#toolbar [data-menu="${menu.id}"]`); e.stopPropagation(); menu.classList.remove('open');
  if (!S.pages.length) return X.toast('Open a PDF first'); X.commitEdit(); X.ACT.letterhead(trg && trg.offsetParent ? trg : null);
}, true);
})();
