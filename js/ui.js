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
function setAccent(v) { if (v && v !== 'violet') root.dataset.accent = v; else delete root.dataset.accent; $$('.swatches button').forEach(b => b.classList.toggle('on', b.dataset.v === (v || 'violet'))); }
function animateTheme() { root.classList.add('theming'); clearTimeout(animateTheme.t); animateTheme.t = setTimeout(() => root.classList.remove('theming'), 450); } // cross-fade instead of a hard switch
const themeIcon = () => $$('[data-act="theme"] i[data-icon]').forEach(i => { i.dataset.icon = root.dataset.theme === 'dark' ? 'moon' : 'sun'; i.textContent = ''; icons(i.parentNode); });
X.ACT.theme = () => { animateTheme(); root.dataset.theme = root.dataset.theme === 'dark' ? 'light' : 'dark'; store.set('xd3.theme', root.dataset.theme); themeIcon(); };
X.ACT.accent = b => { animateTheme(); setAccent(b.dataset.v); store.set('xd3.accent', b.dataset.v); };
setAccent(store.get('xd3.accent')); themeIcon();

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
  const seen = new Set(), out = [];
  for (const el of $$('#topbar [data-act], #toolbar [data-act], #toolbar [data-tool], #status [data-act], #mMore [data-act], #mStamps [data-act], #mArrange [data-act]')) {
    if (el.closest('.brand') || el.dataset.act === 'palette') continue;
    const key = (el.dataset.act || 'tool:' + el.dataset.tool) + '|' + (el.dataset.t || el.dataset.m || el.dataset.v || ''); if (seen.has(key)) continue; seen.add(key);
    const span = $('span', el), title = el.title || '', text = (span ? span.textContent : el.textContent).replace(/[▾\s]+$/, '').trim();
    const sc = (title.match(/\(([^()]{1,10})\)\s*$/) || [])[1] || '', plain = title.replace(/\s*\([^()]{1,10}\)\s*$/, '');
    let name = text || plain; if (el.closest('#mStamps')) name = 'Stamp: ' + name;
    out.push({ el, name, hint: plain && plain !== name ? plain : '', sc, icon: ($('i[data-icon]', el) || { dataset: {} }).dataset.icon || 'bolt' });
  }
  return out;
}
function draw() {
  const q = input.value.toLowerCase().split(/\s+/).filter(Boolean);
  shown = cmds.filter(c => { const hay = (c.name + ' ' + c.hint).toLowerCase(); return q.every(w => hay.includes(w)); })
    .sort((a, b) => q.length ? (b.name.toLowerCase().startsWith(q[0]) - a.name.toLowerCase().startsWith(q[0])) : 0).slice(0, 60);
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
  if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'k') { e.preventDefault(); e.stopImmediatePropagation(); if (pal.hidden) { if (!$('#dlg').open) openPal(); } else closePal(); }
}, true);
})();
