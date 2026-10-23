/* XD3 PDF Editor — writing tools: spelling and grammar (Harper, bundled in lib/harper, fully offline), natural-tone rewrites,
   clarity checks, readability statistics and text clean-up — for Word (.docx) files, text files, pasted text and the open PDF.
   Word files keep their formatting: only the changed words are rewritten inside the document's own runs. */
(() => {
'use strict';
const X = window.XD3, { S, $, $$, toast, busy, dialog, download, esc, icons, clamp } = X;
const store = { get: k => { try { return localStorage.getItem(k); } catch (e) { return null; } }, set: (k, v) => { try { localStorage.setItem(k, v); } catch (e) { /* not remembered */ } } };

/* ───────── state ───────── */
const W = {
  paras: [], kind: 'text', name: 'document', zip: null, xml: null, pdf: null, open: false,
  dialect: store.get('xd3.dialect') || 'American', style: store.get('xd3.style') !== 'off', tab: 'issues', tone: 'natural',
  dict: JSON.parse(store.get('xd3.dict') || '[]'), ignored: new Set(), undo: [], redo: [], rewrites: [], busyLint: false,
};

/* ───────── workspace ───────── */
const box = document.createElement('section'); box.id = 'writer'; box.hidden = true;
box.innerHTML = `<header>
    <b><i data-icon="spell"></i>Writing tools</b><span id="wrName" class="mut"></span>
    <div class="spacer"></div>
    <button class="btn line" data-w="open" title="Open a Word (.docx) or text (.txt) file"><i data-icon="folder"></i><span>Open</span></button>
    <button class="btn line" data-w="paste" title="Start with text from the clipboard, or type"><i data-icon="pages"></i><span>Paste / type</span></button>
    <button class="btn line" data-w="pdf" title="Check the text of the PDF that is open"><i data-icon="edittext"></i><span>This PDF</span></button>
    <span class="sep"></span>
    <select id="wrDialect" title="Variety of English">${['American', 'British', 'Australian', 'Canadian', 'Indian'].map(d => `<option value="${d}">${d} English</option>`).join('')}</select>
    <button class="btn ico" data-w="undo" title="Undo (Ctrl+Z)"><i data-icon="undo"></i></button>
    <button class="btn ico" data-w="redo" title="Redo (Ctrl+Y)"><i data-icon="redo"></i></button>
    <span class="sep"></span>
    <button class="btn primary" data-w="save" title="Save the corrected text"><i data-icon="download"></i><span>Save</span></button>
    <button class="btn ico" data-w="close" title="Close (Esc)"><i data-icon="x"></i></button>
  </header>
  <div class="wr-body">
    <div class="wr-doc" id="wrDoc"><div class="wr-page" id="wrPage"></div></div>
    <aside class="wr-side">
      <nav class="wr-tabs"><button data-wt="issues" class="on">Issues <b id="wrCount">0</b></button><button data-wt="human">Humanize</button><button data-wt="stats">Stats</button><button data-wt="tools">Tools</button></nav>
      <div class="wr-pane" id="wrPane"></div>
    </aside>
  </div>
  <div id="wrPop" hidden></div>`;
document.body.appendChild(box); icons(box);
const page = $('#wrPage'), pane = $('#wrPane'), pop = $('#wrPop');
$('#wrDialect').value = W.dialect;

/* ───────── the checking engine (Harper, loaded on first use) ───────── */
let H = null, HB = null, linter = null, lintDialect = null, engineFail = null;
let starting = null;
async function engine() {
  if (engineFail) throw engineFail;
  try {
    if (!linter) await (starting || (starting = (async () => { // (one start-up, however many callers ask at once)
      const base = new URL('lib/harper/', document.baseURI).href; H = await import(base + 'index.js'); HB = (await import(base + 'binary.js')).binary;
      const L = new H.LocalLinter({ binary: HB, dialect: H.Dialect[W.dialect] }); await L.setup(); if (W.dict.length) await L.importWords(W.dict); lintDialect = W.dialect; linter = L;
    })().finally(() => { starting = null; })));
    if (lintDialect !== W.dialect) { await linter.setDialect(H.Dialect[W.dialect]); lintDialect = W.dialect; }
    return linter;
  } catch (e) { engineFail = new Error('The grammar checker could not start (lib/harper is missing or this browser is too old)'); console.error(e); throw engineFail; }
}
const cpIndex = text => { // Harper counts characters (code points); JavaScript strings count UTF-16 units
  if (!/[\ud800-\udfff]/.test(text)) return i => i;
  const map = []; let u = 0; for (const ch of text) { map.push(u); u += ch.length; } map.push(u); return i => map[Math.min(i, map.length - 1)];
};

/* ───────── extra clarity checks (shown with the grammar issues when "style" is on) ───────── */
const WEAK = ['very', 'really', 'extremely', 'basically', 'actually', 'literally', 'quite', 'totally', 'definitely', 'simply', 'just', 'somewhat', 'rather'];
const CLICHES = ['at the end of the day', 'think outside the box', 'low-hanging fruit', 'move the needle', 'in this day and age', 'the fact of the matter', 'needless to say', 'last but not least', 'each and every', 'first and foremost', 'game changer', 'game-changer'];
const PASSIVE = /\b(?:am|is|are|was|were|be|been|being)\s+(?:\w+ly\s+)?(\w+(?:ed|en)|made|done|built|sent|found|shown|held|kept|told|paid|brought|thought|left|bought|caught|taught|sold|spent|set|put|cut|hit|let|read|led|lost|meant|met|run|said|seen|sung|won)\b/gi;
function styleIssues(text) {
  const out = [];
  for (const m of text.matchAll(/[^.!?]+[.!?]*/g)) { const words = m[0].trim().split(/\s+/).filter(Boolean).length; if (words > 32) out.push({ kind: 'Clarity', pretty: 'Long sentence', msg: `This sentence has ${words} words — consider splitting it.`, a: m.index + (m[0].length - m[0].trimStart().length), b: m.index + m[0].trimEnd().length, sug: [] }); }
  for (const m of text.matchAll(PASSIVE)) if (!/^(?:need|seed|feed|speed|bleed|breed|weed|indeed|tired|interested|excited|bored|based|used|supposed)$/i.test(m[1])) out.push({ kind: 'Clarity', pretty: 'Passive voice', msg: 'Passive voice — an active sentence is usually clearer (say who does it).', a: m.index, b: m.index + m[0].length, sug: [] });
  for (const w of WEAK) for (const m of text.matchAll(new RegExp(`\\b${w}\\b ?`, 'gi'))) out.push({ kind: 'Clarity', pretty: 'Weak word', msg: `“${m[0].trim()}” rarely adds meaning — try removing it or using a stronger word.`, a: m.index, b: m.index + m[0].length, sug: [''] });
  for (const c of CLICHES) for (const m of text.matchAll(new RegExp(`\\b${c.replace(/[-]/g, '[- ]')}\\b`, 'gi'))) out.push({ kind: 'Clarity', pretty: 'Cliché', msg: 'A cliché — say it in plainer words.', a: m.index, b: m.index + m[0].length, sug: [] });
  return out;
}

/* ───────── checking ───────── */
const keyOf = is => is.kind + ':' + is.word.toLowerCase();
async function lintPara(p) {
  const text = p.text; p.issues = []; if (!text.trim()) return;
  const L = await engine(), lints = await L.lint(text, { language: 'plaintext' }), at = cpIndex(text);
  for (const l of lints) {
    const s = l.span(), a = at(s.start), b = at(s.end), word = text.slice(a, b), kind = l.lint_kind();
    const sug = [...new Set(l.suggestions().map(x => { const k = x.kind(), r = x.get_replacement_text(); x.free(); return k === 1 ? '' : k === 2 ? word + r : r; }))].slice(0, 5);
    p.issues.push({ kind, pretty: l.lint_kind_pretty(), msg: l.message(), a, b, word, sug }); l.free();
  }
  if (W.style) for (const is of styleIssues(text)) { is.word = text.slice(is.a, is.b); if (!p.issues.some(o => o.a < is.b && is.a < o.b && o.kind !== 'Clarity')) p.issues.push(is); }
  p.issues = p.issues.filter(is => !W.ignored.has(keyOf(is))).sort((x, y) => x.a - y.a || y.b - x.b);
}
async function lintAll() {
  if (W.busyLint) return; W.busyLint = true; const badge = $('#wrCount'); badge.textContent = '…';
  try {
    await engine(); const t0 = performance.now();
    for (let i = 0; i < W.paras.length; i++) { await lintPara(W.paras[i]); if (performance.now() - t0 > 40) { await new Promise(r => setTimeout(r)); } } // (stays responsive on long documents)
    W.paras.forEach((p, i) => renderPara(i)); renderSide();
  } catch (e) { toast(e.message, true); badge.textContent = '!'; }
  finally { W.busyLint = false; }
}
const CAT = k => k === 'Spelling' || k === 'Typo' ? 'spell' : k === 'Clarity' ? 'style' : /Style|Readability|Redundancy|Repetition|WordChoice|Enhancement/.test(k) ? 'improve' : 'grammar';
const CATNAME = { spell: 'Spelling', grammar: 'Grammar', improve: 'Word choice', style: 'Clarity' };

/* ───────── the document view ───────── */
function renderPara(i) {
  const p = W.paras[i], el = page.children[i]; if (!el) return;
  let html = '', at = 0;
  for (const [n, is] of (p.issues || []).entries()) { if (is.a < at) continue; html += esc(p.text.slice(at, is.a)) + `<mark class="wm wm-${CAT(is.kind)}" data-n="${n}">${esc(p.text.slice(is.a, is.b)) || '&#8203;'}</mark>`; at = is.b; }
  html += esc(p.text.slice(at)); el.innerHTML = html || '<br>';
}
function renderDoc() {
  page.textContent = '';
  W.paras.forEach((p, i) => { const el = document.createElement('p'); el.className = 'wp' + (p.head ? ' wh' : ''); el.contentEditable = 'true'; el.spellcheck = false; el.dataset.i = i; page.appendChild(el); renderPara(i); });
  if (!W.paras.length) page.innerHTML = '<div class="empty">Open a Word or text file, paste text, or check the open PDF.</div>';
  $('#wrName').textContent = W.paras.length ? `${W.name} · ${W.kind === 'docx' ? 'Word document' : W.kind === 'pdf' ? 'text of the open PDF' : 'text'}` : '';
}
const caretOf = el => { const s = getSelection(); if (!s.rangeCount || !el.contains(s.anchorNode)) return null; const r = document.createRange(); r.selectNodeContents(el); r.setEnd(s.anchorNode, s.anchorOffset); return r.toString().length; };
function caretTo(el, off) {
  const w = document.createTreeWalker(el, NodeFilter.SHOW_TEXT); let n, left = off;
  while ((n = w.nextNode())) { if (left <= n.length) { const r = document.createRange(); r.setStart(n, left); r.collapse(true); const s = getSelection(); s.removeAllRanges(); s.addRange(r); return; } left -= n.length; }
}
const textOf = el => el.innerText.replace(/​/g, '').replace(/\n$/, '');

/* ───────── undo ───────── */
const snapW = () => W.paras.map(p => p.text);
function remember() { W.undo.push(snapW()); if (W.undo.length > 100) W.undo.shift(); W.redo.length = 0; }
async function restoreW(t) { W.paras.forEach((p, i) => { p.text = t[i] ?? p.text; }); for (const p of W.paras) await lintPara(p).catch(() => {}); W.paras.forEach((p, i) => renderPara(i)); renderSide(); }
async function undoW() { if (!W.undo.length) return; W.redo.push(snapW()); await restoreW(W.undo.pop()); }
async function redoW() { if (!W.redo.length) return; W.undo.push(snapW()); await restoreW(W.redo.pop()); }

/* ───────── typing ───────── */
let typingSnap = 0;
page.addEventListener('input', e => {
  const el = e.target.closest('.wp'); if (!el) return; const i = +el.dataset.i, p = W.paras[i];
  if (Date.now() - typingSnap > 1500) remember(); typingSnap = Date.now();
  p.text = textOf(el); clearTimeout(el._t);
  el._t = setTimeout(async () => { const c = caretOf(el); await lintPara(p).catch(() => {}); if (document.activeElement === el && textOf(el) === p.text) { renderPara(i); if (c != null) caretTo(el, c); } renderSide(); }, 700);
});
page.addEventListener('keydown', e => {
  const el = e.target.closest('.wp'); if (!el) return;
  if (e.key === 'Enter') { // a new paragraph (Word files: paragraphs are kept as they are — add new ones in Word)
    e.preventDefault(); if (W.kind !== 'text') return toast(W.kind === 'docx' ? 'Paragraphs of a Word file are kept as they are — use Shift+Enter for a line break' : 'Each line of the PDF stays a line', 'info');
    const i = +el.dataset.i, c = caretOf(el) ?? W.paras[i].text.length, t = W.paras[i].text; remember();
    W.paras[i].text = t.slice(0, c); W.paras.splice(i + 1, 0, { text: t.slice(c), issues: [] }); renderDoc(); const nx = page.children[i + 1]; nx.focus(); caretTo(nx, 0); lintPara(W.paras[i]).then(() => renderPara(i)); return;
  }
  if ((e.ctrlKey || e.metaKey) && /^[zy]$/i.test(e.key)) { e.preventDefault(); (e.key.toLowerCase() === 'y' || e.shiftKey ? redoW : undoW)(); }
});
page.addEventListener('paste', e => { e.preventDefault(); document.execCommand('insertText', false, (e.clipboardData.getData('text/plain') || '').replace(/\r/g, '')); }); // (plain text only)

/* ───────── fixing ───────── */
async function applyAt(i, a, b, rep, keepUndo) {
  const p = W.paras[i]; if (!keepUndo) remember();
  p.text = p.text.slice(0, a) + rep + p.text.slice(b); await lintPara(p).catch(() => {}); renderPara(i); renderSide();
}
function showPop(mark) {
  const el = mark.closest('.wp'), i = +el.dataset.i, is = W.paras[i].issues[+mark.dataset.n]; if (!is) return;
  const r = mark.getBoundingClientRect();
  pop.innerHTML = `<div class="wpk wm-${CAT(is.kind)}">${esc(is.pretty)}</div><div class="wpm">${esc(is.msg)}</div>
    <div class="wps">${is.sug.map((s, k) => `<button data-s="${k}">${s ? esc(s) : '<i>remove</i>'}</button>`).join('') || '<span class="mut">No automatic fix — edit the text directly.</span>'}</div>
    <div class="wpa">${CAT(is.kind) === 'spell' ? '<button data-a="dict">Add to dictionary</button>' : ''}<button data-a="ignore">Ignore</button></div>`;
  pop.hidden = false; pop.style.left = clamp(r.left, 8, innerWidth - pop.offsetWidth - 8) + 'px'; pop.style.top = (r.bottom + 6 + pop.offsetHeight > innerHeight ? r.top - pop.offsetHeight - 6 : r.bottom + 6) + 'px';
  pop.onclick = async e => {
    const s = e.target.closest('[data-s]'), a = e.target.closest('[data-a]'); if (!s && !a) return; pop.hidden = true;
    if (s) return applyAt(i, is.a, is.b, is.sug[+s.dataset.s]);
    if (a.dataset.a === 'dict') return addWord(is.word);
    W.ignored.add(keyOf(is)); for (const q of W.paras) q.issues = (q.issues || []).filter(o => keyOf(o) !== keyOf(is)); W.paras.forEach((q, k) => renderPara(k)); renderSide();
  };
}
async function addWord(word) {
  const w = word.trim(); if (!w) return; W.dict = [...new Set([...W.dict, w])]; store.set('xd3.dict', JSON.stringify(W.dict));
  try { await (await engine()).importWords([w]); } catch (e) { /* */ } await lintAll(); toast(`“${w}” added to your dictionary`);
}
page.addEventListener('click', e => { const m = e.target.closest('mark.wm'); if (m) showPop(m); else pop.hidden = true; });
document.addEventListener('pointerdown', e => { if (!pop.hidden && !e.target.closest('#wrPop, mark.wm')) pop.hidden = true; }, true);

/* ───────── side panel ───────── */
function allIssues() { const out = []; W.paras.forEach((p, i) => (p.issues || []).forEach((is, n) => out.push({ i, n, is }))); return out; }
function renderSide() {
  const list = allIssues(); $('#wrCount').textContent = list.length;
  if (W.tab === 'issues') {
    const by = {}; for (const x of list) (by[CAT(x.is.kind)] = by[CAT(x.is.kind)] || []).push(x);
    const quick = list.filter(x => x.is.sug.length === 1 && x.is.sug[0] && CAT(x.is.kind) !== 'style');
    pane.innerHTML = `<div class="wchips">${Object.keys(CATNAME).map(c => `<span class="wchip wm-${c}">${CATNAME[c]} <b>${(by[c] || []).length}</b></span>`).join('')}</div>
      <label class="inl wopt"><input type="checkbox" id="wrStyle"${W.style ? ' checked' : ''}>Clarity suggestions (long sentences, passive voice, weak words, clichés)</label>
      ${quick.length ? `<button class="btn line wfull" data-w="fixall">Fix ${quick.length} clear-cut issue(s)</button>` : ''}
      <div class="wcards">${list.length ? list.slice(0, 400).map(({ i, n, is }) => `<div class="wcard wm-${CAT(is.kind)}" data-i="${i}" data-n="${n}"><div class="wck wm-${CAT(is.kind)}">${esc(is.pretty)}</div><div class="wcw">${esc(is.word.length > 60 ? is.word.slice(0, 60) + '…' : is.word)}</div><div class="wcm">${esc(is.msg)}</div>${is.sug.length ? `<div class="wps">${is.sug.slice(0, 3).map((s, k) => `<button data-s="${k}">${s ? esc(s) : '<i>remove</i>'}</button>`).join('')}</div>` : ''}</div>`).join('') : `<div class="empty">${W.paras.length ? 'No issues found. 🎉' : 'Nothing to check yet.'}</div>`}</div>`;
    $('#wrStyle').onchange = e => { W.style = e.target.checked; store.set('xd3.style', W.style ? 'on' : 'off'); lintAll(); };
  } else if (W.tab === 'human') renderHuman();
  else if (W.tab === 'stats') renderStats();
  else renderTools();
}
pane.addEventListener('click', async e => {
  const card = e.target.closest('.wcard');
  if (card && W.tab === 'issues') {
    const i = +card.dataset.i, is = W.paras[i].issues[+card.dataset.n], s = e.target.closest('[data-s]'); if (!is) return;
    if (s) return applyAt(i, is.a, is.b, is.sug[+s.dataset.s]);
    const m = page.children[i].querySelector(`mark[data-n="${card.dataset.n}"]`); if (m) { m.scrollIntoView({ block: 'center', behavior: 'smooth' }); X.flash(m, 'wflash'); setTimeout(() => showPop(m), 250); }
  }
});
$$('.wr-tabs button', box).forEach(b => b.onclick = () => { W.tab = b.dataset.wt; $$('.wr-tabs button', box).forEach(x => x.classList.toggle('on', x === b)); renderSide(); });

/* ───────── humanize: plain, natural wording ───────── */
const RE = s => new RegExp(`\\b${s}\\b`, 'gi');
const WORDY = [ // [pattern, replacement, why] — clearer, more natural wording
  ['it is (?:important|worth|crucial|essential) to (?:note|mention|remember) that ', '', 'filler opening'], ["it's (?:important|worth) (?:noting|to note|mentioning) that ", '', 'filler opening'],
  ['in order to', 'to', 'wordy'], ['due to the fact that', 'because', 'wordy'], ['owing to the fact that', 'because', 'wordy'], ['in the event that', 'if', 'wordy'], ['at this point in time', 'now', 'wordy'],
  ['at the present time', 'now', 'wordy'], ['for the purpose of', 'for', 'wordy'], ['in spite of the fact that', 'although', 'wordy'], ['with regards? to', 'about', 'wordy'], ['in regard to', 'about', 'wordy'],
  ['a (?:large|significant|considerable) number of', 'many', 'wordy'], ['a (?:myriad|plethora) of', 'many', 'overused'], ['has the ability to', 'can', 'wordy'], ['have the ability to', 'can', 'wordy'],
  ['is able to', 'can', 'wordy'], ['are able to', 'can', 'wordy'], ['prior to', 'before', 'formal'], ['subsequent to', 'after', 'formal'], ['in close proximity to', 'near', 'wordy'],
  ["in today's (?:fast-paced|digital|modern|ever-changing) (?:world|landscape|age)", 'today', 'overused'], ['navigate the complexities of', 'handle', 'overused'], ['a testament to', 'proof of', 'overused'],
  ['plays a (?:crucial|pivotal|vital|key|significant) role in', 'matters for', 'overused'], ['delves? into', 'looks at', 'overused'], ['delving into', 'looking at', 'overused'], ['delved into', 'looked at', 'overused'],
  ['embarks? on', 'starts', 'overused'], ['embarked on', 'started', 'overused'], ['the realm of', 'the world of', 'overused'], ['a rich tapestry of', 'a mix of', 'overused'], ['unlock the (?:full )?potential of', 'make the most of', 'overused'],
];
const WORDS = { utilize: 'use', utilizes: 'uses', utilized: 'used', utilizing: 'using', utilise: 'use', utilises: 'uses', utilised: 'used', utilising: 'using', utilization: 'use', utilisation: 'use',
  facilitate: 'help', facilitates: 'helps', facilitated: 'helped', facilitating: 'helping', commence: 'start', commences: 'starts', commenced: 'started', commencing: 'starting',
  endeavor: 'try', endeavors: 'tries', endeavour: 'try', endeavours: 'tries', subsequently: 'later', numerous: 'many', approximately: 'about', demonstrate: 'show', demonstrates: 'shows', demonstrated: 'showed',
  assist: 'help', assists: 'helps', assisted: 'helped', assisting: 'helping', obtain: 'get', obtains: 'gets', obtained: 'got', purchase: 'buy', purchased: 'bought', regarding: 'about', individuals: 'people',
  furthermore: 'also', moreover: 'also', additionally: 'also', consequently: 'so', nevertheless: 'still', nonetheless: 'still', thus: 'so', hence: 'so', showcase: 'show', showcases: 'shows', showcased: 'showed',
  crucial: 'key', pivotal: 'key', paramount: 'vital', seamless: 'smooth', seamlessly: 'smoothly', robust: 'strong', optimal: 'best', enhance: 'improve', enhances: 'improves', enhanced: 'improved', enhancing: 'improving',
  comprehensive: 'thorough', terminate: 'end', terminated: 'ended', sufficient: 'enough', additional: 'more', initiate: 'start', initiated: 'started', ascertain: 'find out', endeavoring: 'trying', leverage: 'use', leveraging: 'using', leveraged: 'used' };
const CONTRACT = [['do not', "don't"], ['does not', "doesn't"], ['did not', "didn't"], ['is not', "isn't"], ['are not', "aren't"], ['was not', "wasn't"], ['were not', "weren't"], ['cannot', "can't"], ['can not', "can't"],
  ['will not', "won't"], ['would not', "wouldn't"], ['should not', "shouldn't"], ['could not', "couldn't"], ['have not', "haven't"], ['has not', "hasn't"], ['had not', "hadn't"], ['let us', "let's"],
  ['I am', "I'm"], ['we are', "we're"], ['they are', "they're"], ['you are', "you're"], ['it is', "it's"], ['that is', "that's"], ['there is', "there's"], ['what is', "what's"], ['we will', "we'll"], ['you will', "you'll"], ['I will', "I'll"], ['they will', "they'll"], ['we have', "we've"], ['I have', "I've"], ['you have', "you've"], ['they have', "they've"]];
const CASUAL = { purchase: 'buy', inquire: 'ask', require: 'need', requires: 'needs', reside: 'live', receive: 'get', received: 'got', commence: 'start', 'a lot of': 'lots of', therefore: 'so', however: 'but' };
const FORMAL = { "a lot of": 'many', 'lots of': 'many', 'kind of': 'somewhat', 'sort of': 'somewhat', okay: 'acceptable', ok: 'acceptable', kids: 'children', stuff: 'material', gonna: 'going to', wanna: 'want to', gotta: 'have to', 'get rid of': 'remove', 'find out': 'determine', 'figure out': 'determine', 'look into': 'investigate', 'set up': 'establish', big: 'large', really: 'truly' };
const caseLike = (src, rep) => !rep ? rep : src[0] === src[0].toUpperCase() && src[0] !== src[0].toLowerCase() ? rep[0].toUpperCase() + rep.slice(1) : rep;
function rewrites(tone) { // → [{ i, a, b, from, to, why }] for the whole document, without overlaps
  const out = [];
  W.paras.forEach((p, i) => {
    const t = p.text, taken = [], add = (a, b, to, why) => { if (taken.some(([x, y]) => a < y && x < b)) return; taken.push([a, b]); out.push({ i, a, b, from: t.slice(a, b), to: caseLike(t.slice(a, b), to), why }); };
    if (tone !== 'formal') {
      for (const [src, rep, why] of WORDY) for (const m of t.matchAll(RE(src))) add(m.index, m.index + m[0].length, rep, why === 'wordy' ? 'shorter, plainer' : why === 'overused' ? 'a stock phrase that sounds machine-written' : why === 'formal' ? 'less stiff' : 'adds nothing');
      for (const w in WORDS) for (const m of t.matchAll(RE(w))) add(m.index, m.index + m[0].length, WORDS[w], 'simpler word');
      for (const m of t.matchAll(/\s?—\s?/g)) add(m.index, m.index + m[0].length, ', ', 'long dashes everywhere read as machine-written');
      for (const [a, b] of CONTRACT) for (const m of t.matchAll(new RegExp(`\\b${a}\\b(?![.,;:!?)]|\\s*$)`, 'gi'))) add(m.index, m.index + m[0].length, b, 'contraction — sounds like a person talking');
      if (tone === 'casual') for (const w in CASUAL) for (const m of t.matchAll(RE(w))) add(m.index, m.index + m[0].length, CASUAL[w], 'more casual');
    } else {
      for (const [a, b] of CONTRACT) for (const m of t.matchAll(new RegExp(`\\b${b.replace("'", "['’]")}\\b`, 'gi'))) add(m.index, m.index + m[0].length, a === 'can not' ? 'cannot' : a, 'formal writing avoids contractions');
      for (const w in FORMAL) for (const m of t.matchAll(RE(w))) add(m.index, m.index + m[0].length, FORMAL[w], 'more formal');
      for (const m of t.matchAll(/\s?—\s?/g)) add(m.index, m.index + m[0].length, ', ', 'a comma reads more evenly');
    }
    // sentences that start with the same word as the one before → flag only
  });
  return out.sort((x, y) => x.i - y.i || x.a - y.a);
}
function renderHuman() {
  W.rewrites = rewrites(W.tone);
  pane.innerHTML = `<p class="mut">Suggestions to make the writing sound more natural and personal: plainer words, fewer stock phrases, contractions${W.tone === 'formal' ? ' — or, in Formal, the opposite' : ''}. Nothing changes until you accept it.</p>
    <div class="wtone">${[['natural', 'Natural'], ['casual', 'Casual'], ['formal', 'Formal']].map(([v, l]) => `<button data-tone="${v}" class="${W.tone === v ? 'on' : ''}">${l}</button>`).join('')}</div>
    ${W.rewrites.length ? `<button class="btn primary wfull" data-w="acceptall">Accept all ${W.rewrites.length} change(s)</button>` : ''}
    <div class="wcards">${W.rewrites.length ? W.rewrites.slice(0, 400).map((r, k) => `<div class="wcard wrw" data-k="${k}"><div class="wcw"><s>${esc(r.from)}</s> → <b>${r.to ? esc(r.to) : '<i>remove</i>'}</b></div><div class="wcm">${esc(r.why)}</div><div class="wps"><button data-acc="${k}">Accept</button><button data-skip="${k}">Skip</button></div></div>`).join('') : '<div class="empty">Nothing to change in this tone.</div>'}</div>`;
}
pane.addEventListener('click', async e => {
  const tb = e.target.closest('[data-tone]'); if (tb) { W.tone = tb.dataset.tone; return renderHuman(); }
  const acc = e.target.closest('[data-acc]'), skip = e.target.closest('[data-skip]');
  if (acc) { const r = W.rewrites[+acc.dataset.acc]; if (r && W.paras[r.i].text.slice(r.a, r.b) === r.from) await applyAt(r.i, r.a, r.b, r.to); else renderHuman(); }
  if (skip) { const card = skip.closest('.wcard'); card.remove(); }
  const c = !acc && !skip && e.target.closest('.wrw'); if (c) { const r = W.rewrites[+c.dataset.k]; const el = page.children[r.i]; if (el) { el.scrollIntoView({ block: 'center', behavior: 'smooth' }); X.flash(el, 'wflash'); } }
});
async function acceptAll() {
  remember(); const list = rewrites(W.tone), byPara = {}; for (const r of list) (byPara[r.i] = byPara[r.i] || []).push(r);
  for (const i in byPara) { const p = W.paras[i]; for (const r of byPara[i].sort((x, y) => y.a - x.a)) p.text = p.text.slice(0, r.a) + r.to + p.text.slice(r.b); p.text = tidy(p.text); }
  for (const i in byPara) { await lintPara(W.paras[i]).catch(() => {}); renderPara(+i); }
  renderSide(); toast(`${list.length} change(s) made — Ctrl+Z to undo`);
}
const tidy = t => t.replace(/ {2,}/g, ' ').replace(/ ([,.;:!?])/g, '$1').replace(/(^|[.!?]\s+)([a-z])/g, (m, a, b) => a + b.toUpperCase()).replace(/,\s*,/g, ',');

/* ───────── statistics ───────── */
const syllables = w => { w = w.toLowerCase().replace(/[^a-z]/g, ''); if (!w) return 0; if (w.length <= 3) return 1; w = w.replace(/(?:[^laeiouy]es|ed|[^laeiouy]e)$/, '').replace(/^y/, ''); const m = w.match(/[aeiouy]{1,2}/g); return Math.max(1, m ? m.length : 1); };
const STOP = new Set('the a an and or but if of to in on at by for with from as is are was were be been being it its this that these those i you he she we they me him her us them my your his our their not no so do does did have has had will would can could should may might must there here what which who whom whose when where why how all any each few more most other some such than too very just also into over under about after before again then once only own same both up down out off'.split(' '));
function renderStats() {
  const text = W.paras.map(p => p.text).join('\n'), words = text.match(/[\p{L}\p{N}'’-]+/gu) || [], sents = (text.match(/[^.!?\n]+[.!?]+|[^.!?\n]+$/gm) || []).filter(s => /\p{L}/u.test(s));
  const nW = words.length, nS = Math.max(1, sents.length), syl = words.reduce((a, w) => a + syllables(w), 0), fre = 206.835 - 1.015 * (nW / nS) - 84.6 * (syl / Math.max(1, nW)), fk = 0.39 * (nW / nS) + 11.8 * (syl / Math.max(1, nW)) - 15.59;
  const level = fre >= 80 ? 'Very easy' : fre >= 70 ? 'Easy' : fre >= 60 ? 'Plain English' : fre >= 50 ? 'Fairly difficult' : fre >= 30 ? 'Difficult' : 'Very difficult';
  const freq = {}; for (const w of words) { const k = w.toLowerCase(); if (k.length > 2 && !STOP.has(k)) freq[k] = (freq[k] || 0) + 1; }
  const top = Object.entries(freq).sort((a, b) => b[1] - a[1]).slice(0, 10), long = sents.filter(s => s.split(/\s+/).filter(Boolean).length > 25).length, passive = (text.match(PASSIVE) || []).length;
  const mins = n => n < 1 ? '< 1 min' : Math.round(n) + ' min';
  pane.innerHTML = `<div class="wstats">
    ${[['Words', nW], ['Characters', text.replace(/\n/g, '').length], ['Without spaces', text.replace(/\s/g, '').length], ['Sentences', sents.length], ['Paragraphs', W.paras.filter(p => p.text.trim()).length], ['Reading time', mins(nW / 238)], ['Speaking time', mins(nW / 150)], ['Words per sentence', (nW / nS).toFixed(1)]].map(([k, v]) => `<div><small>${k}</small><b>${v}</b></div>`).join('')}
    </div>
    <div class="wread"><div class="wbar"><span style="width:${clamp(fre, 0, 100)}%"></span></div><b>${level}</b> · reading ease ${Math.round(fre)} / 100 · school grade ${Math.max(1, Math.round(fk))}</div>
    <p class="mut">${long} long sentence(s) (over 25 words) · ${passive} passive construction(s) · ${new Set(words.map(w => w.toLowerCase())).size} different words</p>
    <h4>Most used words</h4><div class="wtop">${top.map(([w, n]) => `<span>${esc(w)} <b>${n}</b></span>`).join('') || '<span class="mut">—</span>'}</div></div>`;
}

/* ───────── text tools ───────── */
function renderTools() {
  pane.innerHTML = `<p class="mut">Works on the selected text, or on the whole document when nothing is selected.</p>
    <h4>Change case</h4><div class="wgrid"><button data-tool="upper">UPPERCASE</button><button data-tool="lower">lowercase</button><button data-tool="title">Title Case</button><button data-tool="sentence">Sentence case</button></div>
    <h4>Clean up</h4><div class="wgrid"><button data-tool="spaces">Fix spacing</button><button data-tool="curly">“Smart” quotes</button><button data-tool="straight">"Straight" quotes</button><button data-tool="empty">Remove empty lines</button><button data-tool="join">Join broken lines</button><button data-tool="dups">Remove duplicate lines</button><button data-tool="sort">Sort lines A–Z</button><button data-tool="trim">Trim lines</button></div>
    <h4>Find &amp; replace</h4><div class="wfind"><input id="wrF" placeholder="Find"><input id="wrR" placeholder="Replace with"><label class="inl"><input type="checkbox" id="wrCase">Match case</label><label class="inl"><input type="checkbox" id="wrWord">Whole words</label><button class="btn line" data-tool="replace">Replace all</button></div>
    <h4>Your dictionary</h4><div class="wdict">${W.dict.length ? W.dict.map((w, k) => `<span>${esc(w)}<button data-del="${k}" title="Remove">×</button></span>`).join('') : '<span class="mut">Words you add are kept here.</span>'}</div>`;
}
const sel = () => { const s = getSelection(); if (!s.rangeCount || s.isCollapsed) return null; const a = s.anchorNode && s.anchorNode.parentElement && s.anchorNode.parentElement.closest('.wp'), b = s.focusNode && s.focusNode.parentElement && s.focusNode.parentElement.closest('.wp'); if (!a || a !== b) return null; const lo = Math.min(caretOf(a) ?? 0, (() => { const r = document.createRange(); r.selectNodeContents(a); r.setEnd(s.focusNode, s.focusOffset); return r.toString().length; })()), hi = s.toString().length + lo; return { i: +a.dataset.i, a: lo, b: hi }; };
let lastSel = null; document.addEventListener('selectionchange', () => { if (!box.hidden && document.activeElement && document.activeElement.closest && document.activeElement.closest('.wp')) lastSel = sel(); });
async function runTool(t) {
  const title = async s => linter ? await linter.toTitleCase(s) : s.replace(/\b\p{L}/gu, c => c.toUpperCase());
  const map = async s => t === 'upper' ? s.toUpperCase() : t === 'lower' ? s.toLowerCase() : t === 'title' ? await title(s) : t === 'sentence' ? s.toLowerCase().replace(/(^\s*|[.!?]\s+)(\p{L})/gu, (m, a, b) => a + b.toUpperCase())
    : t === 'spaces' ? tidy(s.replace(/[ \t]+/g, ' ').trim()) : t === 'curly' ? s.replace(/(^|[\s(\[{])"/g, '$1“').replace(/"/g, '”').replace(/(^|[\s(\[{])'/g, '$1‘').replace(/'/g, '’') : t === 'straight' ? s.replace(/[“”]/g, '"').replace(/[‘’]/g, "'") : t === 'trim' ? s.trim() : s;
  remember(); let n = 0;
  if (['upper', 'lower', 'title', 'sentence', 'spaces', 'curly', 'straight', 'trim'].includes(t)) {
    if (lastSel && lastSel.b > lastSel.a) { const p = W.paras[lastSel.i]; p.text = p.text.slice(0, lastSel.a) + await map(p.text.slice(lastSel.a, lastSel.b)) + p.text.slice(lastSel.b); n = 1; }
    else for (const p of W.paras) { const v = await map(p.text); if (v !== p.text) { p.text = v; n++; } }
  } else if (t === 'replace') {
    const f = $('#wrF').value; if (!f) return toast('Type the words to find'); const rx = new RegExp(($('#wrWord').checked ? '\\b' : '') + f.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + ($('#wrWord').checked ? '\\b' : ''), $('#wrCase').checked ? 'g' : 'gi');
    for (const p of W.paras) { const v = p.text.replace(rx, () => { n++; return $('#wrR').value; }); p.text = v; }
    toast(`${n} replacement(s)`);
  } else if (W.kind !== 'text') return toast('This works on plain text — Word and PDF documents keep their own lines', 'info');
  else {
    const lines = W.paras.map(p => p.text); let nl = lines;
    if (t === 'empty') nl = lines.filter(l => l.trim()); else if (t === 'dups') nl = lines.filter((l, k) => !l.trim() || lines.indexOf(l) === k); else if (t === 'sort') nl = [...lines].sort((a, b) => a.localeCompare(b));
    else if (t === 'join') { nl = []; for (const l of lines) { const last = nl[nl.length - 1]; if (last != null && last.trim() && l.trim() && !/[.!?:;]["”’)]?$/.test(last.trim())) nl[nl.length - 1] = last.replace(/-$/, '') + (last.endsWith('-') ? '' : ' ') + l.trim(); else nl.push(l); } }
    n = lines.length - nl.length; W.paras = nl.map(text => ({ text, issues: [] })); renderDoc();
  }
  await lintAll(); if (t !== 'replace') toast(n ? 'Done' : 'Nothing to change');
}
pane.addEventListener('click', e => {
  const t = e.target.closest('[data-tool]'); if (t) return runTool(t.dataset.tool);
  const d = e.target.closest('[data-del]'); if (d) { W.dict.splice(+d.dataset.del, 1); store.set('xd3.dict', JSON.stringify(W.dict)); (linter ? linter.clearWords().then(() => linter.importWords(W.dict)) : Promise.resolve()).then(lintAll); renderTools(); }
});
pane.addEventListener('mousedown', e => { if (e.target.closest('[data-tool]')) e.preventDefault(); }); // (keeps the text selection)

/* ───────── loading ───────── */
const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
async function loadDocx(file) {
  const JSZip = await X.loadLib('jszip'), zip = await JSZip.loadAsync(await file.arrayBuffer()), xmlText = await zip.file('word/document.xml').async('string');
  const xml = new DOMParser().parseFromString(xmlText, 'application/xml'), paras = [];
  for (const wp of xml.getElementsByTagNameNS(W_NS, 'p')) {
    const segs = []; let text = '';
    for (const el of wp.getElementsByTagNameNS(W_NS, '*')) {
      if (el.localName === 't' && el.parentNode.localName === 'r') { segs.push({ el, a: text.length, len: el.textContent.length }); text += el.textContent; }
      else if (el.localName === 'tab' && el.parentNode.localName === 'r') text += '\t'; else if ((el.localName === 'br' || el.localName === 'cr') && el.parentNode.localName === 'r') text += '\n';
    }
    const style = wp.getElementsByTagNameNS(W_NS, 'pStyle')[0], head = style && /heading|title/i.test(style.getAttributeNS(W_NS, 'val') || style.getAttribute('w:val') || '');
    paras.push({ text, orig: text, segs, head, issues: [] });
  }
  return { zip, xml, paras };
}
function changes(old, nw) { // word-level differences between two versions → [{ a, b, rep }] in the old text's positions
  if (old === nw) return [];
  const tok = t => t.match(/\s+|[\p{L}\p{N}'’]+|[^\s\p{L}\p{N}]/gu) || [], A = tok(old), B = tok(nw), n = A.length, m = B.length;
  if (n * m > 4e6) { let a = 0; while (a < old.length && a < nw.length && old[a] === nw[a]) a++; let b = 0; while (b < old.length - a && b < nw.length - a && old[old.length - 1 - b] === nw[nw.length - 1 - b]) b++; return [{ a, b: old.length - b, rep: nw.slice(a, nw.length - b) }]; }
  const dp = Array.from({ length: n + 1 }, () => new Uint16Array(m + 1));
  for (let i = n - 1; i >= 0; i--) for (let j = m - 1; j >= 0; j--) dp[i][j] = A[i] === B[j] ? dp[i + 1][j + 1] + 1 : Math.max(dp[i + 1][j], dp[i][j + 1]);
  const out = []; let i = 0, j = 0, pa = 0, cur = null;
  const flush = () => { if (cur) { out.push(cur); cur = null; } };
  while (i < n || j < m) {
    if (i < n && j < m && A[i] === B[j]) { flush(); pa += A[i].length; i++; j++; }
    else if (j < m && (i >= n || dp[i][j + 1] >= dp[i + 1][j])) { cur = cur || { a: pa, b: pa, rep: '' }; cur.rep += B[j++]; }
    else { cur = cur || { a: pa, b: pa, rep: '' }; pa += A[i].length; cur.b = pa; i++; }
  }
  flush(); return out;
}
function splice(pieces, ch) { // apply changes (last first) to pieces [{ a, len, get(), set(t) }] laid end to end over the old text; formatting of each piece is kept
  for (const c of [...ch].sort((x, y) => y.a - x.a)) {
    const host = pieces.find(q => c.a >= q.a && c.a < q.a + q.len) || [...pieces].reverse().find(q => c.a >= q.a && c.a <= q.a + q.len) || pieces[pieces.length - 1]; let done = false;
    for (const q of pieces) {
      const lo = Math.max(c.a, q.a), hi = Math.min(c.b, q.a + q.len), mine = q === host;
      if (!mine && hi <= lo) continue;
      const t = q.get(), from = clamp(lo - q.a, 0, t.length), to = Math.max(from, clamp(hi - q.a, 0, t.length));
      q.set(t.slice(0, mine ? clamp(c.a - q.a, 0, t.length) : from) + (mine && !done ? c.rep : '') + t.slice(mine ? Math.max(clamp(c.a - q.a, 0, t.length), to) : to)); if (mine) done = true;
    }
  }
}
function writeDocxPara(p) { // each change goes into the run where it sits, so bold, italic, colour and the rest stay exactly where they were
  if (p.text === p.orig || !p.segs.length) return;
  const XS = 'http://www.w3.org/XML/1998/namespace';
  splice(p.segs.map(g => ({ a: g.a, len: g.len, get: () => g.el.textContent, set: t => { g.el.textContent = t; g.el.setAttributeNS(XS, 'xml:space', 'preserve'); } })), changes(p.orig, p.text));
}
async function buildDocx() { // → Blob
  const JSZip = await X.loadLib('jszip');
  if (W.kind === 'docx') { for (const p of W.paras) writeDocxPara(p); W.zip.file('word/document.xml', new XMLSerializer().serializeToString(W.xml)); const blob = await W.zip.generateAsync({ type: 'blob', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' }); W.paras.forEach(p => { p.orig = p.text; }); return blob; }
  const zip = new JSZip(), x = s => esc(s).replace(/'/g, '&apos;');
  zip.file('[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  zip.file('_rels/.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  zip.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="${W_NS}"><w:body>${W.paras.map(p => `<w:p><w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr><w:t xml:space="preserve">${x(p.text)}</w:t></w:r></w:p>`).join('')}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>`);
  return zip.generateAsync({ type: 'blob', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
}
async function loadPdf() { // the open document's text: paragraphs of original lines, and text boxes added in the editor
  const paras = [];
  for (const p of S.pages) {
    const covered = it => X.covered(p, it);
    if (p.src >= 0) for (const b of await X.getBlocks(p)) { const L = (b.lines || [b]).filter(l => !covered(l)); if (!L.length) continue; let text = '', parts = []; for (const l of L) { if (text) text += ' '; parts.push({ it: l, a: text.length, b: text.length + l.str.length }); text += l.str; } paras.push({ text, orig: text, pid: p.id, parts, issues: [] }); }
    for (const o of p.objs) if (o.type === 'text' && o.text.trim()) paras.push({ text: o.text, orig: o.text, pid: p.id, obj: o.id, issues: [] });
  }
  return paras;
}
async function applyToPdf() {
  const changed = W.paras.filter(p => p.text !== p.orig); if (!changed.length) return toast('Nothing has changed');
  await busy(async () => {
    X.pushUndo(); let n = 0;
    for (const q of changed) {
      const p = S.pages.find(x => x.id === q.pid); if (!p) continue;
      if (q.obj) { const o = p.objs.find(x => x.id === q.obj); if (o) { o.text = q.text; n++; } continue; }
      const lines = q.parts.map(pt => ({ a: pt.a, len: pt.b - pt.a, it: pt.it, text: q.orig.slice(pt.a, pt.b) })), before = lines.map(l => l.text);
      splice(lines.map(l => ({ a: l.a, len: l.len, get: () => l.text, set: t => { l.text = t; } })), changes(q.orig, q.text));
      for (const [k, l] of lines.entries()) { if (l.text === before[k]) continue; const o = await X.editOriginal(p, l.it, l.text.trim() ? l.text : ''); if (!l.text.trim()) p.objs = p.objs.filter(x => x !== o); n++; } // (only the lines a change touches are rewritten)
      q.orig = q.text;
    }
    S.pages.forEach(X.drawObjs); X.updateUI(); toast(`${n} line(s) updated in the PDF — they are rewritten in the document's own font when you save`);
  }, 'Updating the PDF…');
}
async function openWith(paras, kind, name, extra = {}) {
  Object.assign(W, { paras, kind, name, zip: null, xml: null, undo: [], redo: [], ...extra }); W.ignored.clear(); renderDoc(); renderSide(); lintAll();
}
function pick(accept) { return new Promise(res => { const i = document.createElement('input'); i.type = 'file'; i.accept = accept; i.onchange = () => res(i.files[0] || null); i.click(); }); }
async function openFile(file) {
  if (!file) return;
  await busy(async () => {
    if (/\.docx$/i.test(file.name)) { const d = await loadDocx(file); await openWith(d.paras, 'docx', file.name.replace(/\.docx$/i, ''), { zip: d.zip, xml: d.xml }); }
    else { const t = (await file.text()).replace(/\r\n?/g, '\n'); await openWith(t.split('\n').map(text => ({ text, orig: text, issues: [] })), 'text', file.name.replace(/\.\w+$/, '')); }
  }, 'Opening…');
}
function show() { W.open = true; box.hidden = false; document.body.classList.add('modal'); renderSide(); engine().catch(e => toast(e.message, true)); }
function close() { W.open = false; box.hidden = true; pop.hidden = true; document.body.classList.remove('modal'); }

/* ───────── header actions ───────── */
const ACTS = {
  open: async () => openFile(await pick('.docx,.txt,.md,.csv,text/plain')),
  paste: async () => { let t = ''; try { t = await navigator.clipboard.readText(); } catch (e) { /* no clipboard access */ } await openWith((t || '').replace(/\r\n?/g, '\n').split('\n').map(text => ({ text, orig: text, issues: [] })), 'text', 'text'); if (!t) { const el = page.querySelector('.wp'); if (el) el.focus(); toast('Type or paste your text', 'info'); } },
  pdf: async () => { if (!S.pages.length) return toast('Open a PDF first'); await busy(async () => openWith(await loadPdf(), 'pdf', X.baseName()), 'Reading the PDF…'); },
  undo: undoW, redo: redoW, close,
  fixall: async () => { remember(); let n = 0; for (const [i, p] of W.paras.entries()) { const q = (p.issues || []).filter(is => is.sug.length === 1 && is.sug[0] && CAT(is.kind) !== 'style').sort((x, y) => y.a - x.a); for (const is of q) { p.text = p.text.slice(0, is.a) + is.sug[0] + p.text.slice(is.b); n++; } if (q.length) { await lintPara(p).catch(() => {}); renderPara(i); } } renderSide(); toast(`${n} issue(s) fixed — Ctrl+Z to undo`); },
  acceptall: acceptAll,
  save: () => {
    if (!W.paras.length) return toast('Nothing to save yet');
    const items = [W.kind === 'pdf' && X.mi('Apply the changes to the PDF', 'edittext', 'w_applypdf'), X.mi('Save as Word (.docx)', 'word', 'w_docx'), X.mi('Save as text (.txt)', 'text', 'w_txt'), X.mi('Copy all text', 'copy', 'w_copy'), '-', X.mi('Make a PDF from this text', 'pages', 'w_pdf')];
    const b = $('[data-w=save]', box).getBoundingClientRect(); X.showCtx({ clientX: b.left, clientY: b.bottom + 4, target: box }, items);
  },
};
box.querySelector('header').addEventListener('click', e => { const b = e.target.closest('[data-w]'); if (b && ACTS[b.dataset.w]) ACTS[b.dataset.w](); });
pane.addEventListener('click', e => { const b = e.target.closest('[data-w]'); if (b && ACTS[b.dataset.w]) ACTS[b.dataset.w](); });
$('#wrDialect').onchange = e => { W.dialect = e.target.value; store.set('xd3.dialect', W.dialect); lintAll(); };
window.addEventListener('keydown', e => { if (!W.open || $('#dlg').open) return; if (e.key === 'Escape' && pop.hidden) close(); else if (e.key === 'Escape') pop.hidden = true; }, true);
Object.assign(X.ACT, {
  writer: () => show(), w_applypdf: async () => { await applyToPdf(); close(); },
  w_docx: async () => busy(async () => download(await buildDocx(), (W.name || 'document') + '.docx'), 'Saving…'),
  w_txt: () => download(new Blob(['﻿' + W.paras.map(p => p.text).join('\r\n')], { type: 'text/plain;charset=utf-8' }), (W.name || 'document') + '.txt'),
  w_copy: () => navigator.clipboard.writeText(W.paras.map(p => p.text).join('\n')).then(() => toast('Copied'), () => toast('Copy is not available here', true)),
  w_pdf: async () => { const blob = await buildDocx(); close(); await X.openFiles([new File([blob], (W.name || 'document') + '.docx')], false); },
});
X.NO_DOC.add('writer'); ['w_docx', 'w_txt', 'w_copy', 'w_pdf'].forEach(a => X.NO_DOC.add(a));
X.openInWriter = async file => { show(); await openFile(file); };
X.writing = { W, engine, lintAll, loadDocx, buildDocx, changes, splice, openWith, rewrites, styleIssues, applyToPdf, show, close, openFile };
})();
