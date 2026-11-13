/* XD3 PDF Editor — offline AI writing assistant: the page side. The model (Qwen2.5-0.5B-Instruct, Apache 2.0) runs in js/ai-worker.js,
   loaded from the app's own files the first time an AI button is used; text never leaves the device.
   X.ai.run(task, text, { onToken }) → the new text. Long texts are done a few sentences at a time; the answer is cleaned
   (no "Here is…" preambles or quotes) and checked (an answer that drifts away from the text is refused, not shown). */
(() => {
'use strict';
const X = window.XD3, { toast, esc } = X;

const TASKS = { // what each button asks for; an example answer keeps the small model on track where it needs one
  humanize: { label: 'Humanize', icon: 'sparkle', hint: 'Natural, human-sounding wording', ask: 'Rewrite this so it sounds human and conversational. Output only the new text.',
    shot: ['Moreover, it is crucial to leverage innovative solutions in order to facilitate seamless collaboration across teams.', 'We also need good tools so our teams can work together without friction.'] },
  fix: { label: 'Fix grammar', icon: 'check', hint: 'Spelling, grammar and punctuation', ask: 'Correct the grammar and spelling mistakes in the following text. Output only the corrected text.' },
  formal: { label: 'Formal', icon: 'stamp', hint: 'Polite and professional', ask: 'Rewrite the following message in a polite, professional tone. Keep it short. Output only the rewritten message.' },
  friendly: { label: 'Friendly', icon: 'note', hint: 'Warm and casual', ask: 'Rewrite this in a warm, friendly and casual tone. Output only the new text.',
    shot: ['Please be advised that the office will be closed on Monday.', 'Just a heads-up: the office is closed on Monday.'] },
  shorter: { label: 'Shorter', icon: 'minus', hint: 'Same message, fewer words', ask: 'Make this shorter. Keep the key facts. Output only the new text.',
    shot: ['We would like to take this opportunity to let you know that your order has now been shipped and should arrive within the next few days.', 'Your order has shipped and should arrive in a few days.'] },
  longer: { label: 'Longer', icon: 'plus', hint: 'More detail, same facts', ask: 'Write two or three sentences that explain the following statement in more detail. Output only those sentences.' },
  summary: { label: 'Summarize', icon: 'outline', hint: 'The main point in a sentence or two', ask: 'Summarize the following text in one or two sentences. Output only the summary.' },
  bullets: { label: 'Bullet points', icon: 'listbul', hint: 'The key points as a list', ask: 'List the key points of the following text as short bullet points, one per line, each starting with "- ". Output only the list.' },
  continue: { label: 'Continue writing', icon: 'pen', hint: 'Adds a sentence or two', ask: 'Continue the following text with one or two more sentences in the same style. Output only the new sentences.' },
};
const A = { state: 'off', pct: 0, err: '', threads: 0 }; // off · loading · ready · error
let worker = null, seq = 0, ready = null;
const jobs = new Map(), emit = () => document.dispatchEvent(new CustomEvent('xd3:ai', { detail: A }));
const supported = () => typeof WebAssembly === 'object' && typeof Worker === 'function';
function start() {
  if (worker) return; worker = new Worker('js/ai-worker.js', { type: 'module' });
  worker.onmessage = ({ data: d }) => {
    if (d.type === 'progress' && d.total) { A.pct = Math.min(99, Math.round(d.loaded / d.total * 100)); emit(); return; }
    const j = jobs.get(d.id); if (!j) return;
    if (d.type === 'token') j.onToken && j.onToken(d.text);
    else if (d.type === 'ready' || d.type === 'done') { jobs.delete(d.id); j.res(d); }
    else if (d.type === 'error') { jobs.delete(d.id); j.rej(new Error(d.message)); }
  };
  worker.onerror = e => { const err = new Error('The AI could not start here' + (e && e.message ? ': ' + e.message : '')); for (const j of jobs.values()) j.rej(err); jobs.clear(); A.state = 'error'; A.err = err.message; emit(); worker = null; ready = null; };
}
const post = (m, onToken) => new Promise((res, rej) => { start(); const id = ++seq; jobs.set(id, { res, rej, onToken }); worker.postMessage({ id, ...m }); });
function ensure() { // load the model once (about 480 MB from the app's own files; a few seconds after the first time)
  if (!supported()) return Promise.reject(new Error('This browser cannot run the offline AI'));
  return ready || (ready = (async () => { A.state = 'loading'; A.pct = 0; emit(); const d = await post({ op: 'load' }); A.state = 'ready'; A.threads = d.threads; A.pct = 100; emit(); })()
    .catch(e => { ready = null; A.state = 'error'; A.err = e.message; emit(); throw e; }));
}
function stop() { if (worker) worker.postMessage({ id: 0, op: 'stop' }); }

/* ── cleaning and checking an answer ── */
const words = t => (t.toLowerCase().match(/[a-z0-9']+/g) || []);
function clean(out, input, task) {
  let t = String(out || '').replace(/\r/g, '').trim();
  t = t.replace(/^(?:sure|certainly|of course|okay|ok)[,!.]?\s*/i, '').replace(/^(?:here(?:'s| is| are)|the (?:corrected|rewritten|shortened|new|revised|expanded|summari[sz]ed) (?:text|version|message)[^:\n]*|(?:corrected|rewritten|revised) (?:text|version)|summary|task)[^:\n]{0,80}:\s*/i, '');
  t = t.replace(/^"{3}\s*|\s*"{3}$/g, '').replace(/^```\w*\s*|\s*```$/g, '').trim();
  if (/^["“].*["”]$/s.test(t) && !/^["“]/.test(input.trim())) t = t.slice(1, -1).trim();
  t = t.replace(/\n{3,}/g, '\n\n');
  if (task === 'bullets') t = t.split('\n').map(l => l.replace(/^\s*(?:[-*•]|\d+[.)])\s*/, '').trim()).filter(Boolean).map(l => '• ' + l).join('\n');
  if (!/\n/.test(input) && task !== 'bullets' && task !== 'longer') t = t.replace(/\s*\n+\s*/g, ' ');
  if (input.trim() && /^[a-z]/.test(input.trim()) && task !== 'summary' && task !== 'bullets') { /* keep a lower-case start the same as the original */ }
  return t.trim();
}
function sane(out, input, task) { // → '' when the answer is fine, else why it was refused
  if (!out) return 'no answer';
  const a = words(input), b = words(out); if (!b.length) return 'no answer';
  if (task === 'continue' || task === 'longer' || task === 'summary' || task === 'bullets') return b.length > 400 ? 'too long' : '';
  const A2 = new Set(a), kept = b.filter(w => A2.has(w)).length, keep = kept / b.length, newWords = b.length - kept;
  if (task === 'fix' && ((keep < 0.75 && newWords > Math.max(3, a.length * 0.15)) || b.length > a.length * 1.3 + 4)) return 'changed too much'; // (a short sentence may need a few new words)
  if (b.length > a.length * 2.2 + 12) return 'too long';
  if (keep < 0.2 && a.length > 6) return 'off topic';
  return '';
}
const humanPass = t => { // the rule-based humanizer tidies what the model left (stock phrases, "utilize", …)
  try { const r = X.writing.rewrites('natural', [{ text: t }]); for (const x of [...r].sort((p, q) => q.a - p.a)) t = t.slice(0, x.a) + x.to + t.slice(x.b); return t.replace(/\s{2,}/g, ' ').replace(/^\s*[a-z]/, c => c.toUpperCase()); } catch (e) { return t; }
};
function chunks(text) { // a few sentences at a time keeps the small model accurate and fast
  if (text.length <= 700) return [text];
  const s = text.match(/[^.!?]+(?:[.!?]+["')\]]*\s*|$)/g) || [text], out = []; let cur = '';
  for (const x of s) { if (cur && (cur + x).length > 600) { out.push(cur); cur = ''; } cur += x; } if (cur) out.push(cur); return out;
}
async function once(task, text, onToken, temp) {
  const T = TASKS[task], m = [], fence = t => '"""\n' + t + '\n"""';
  if (T.shot) m.push({ role: 'user', content: T.ask + '\n\n' + fence(T.shot[0]) }, { role: 'assistant', content: T.shot[1] });
  m.push({ role: 'user', content: T.ask + '\n\n' + fence(text) });
  const max = task === 'longer' || task === 'continue' ? 160 : task === 'bullets' ? Math.round(text.length / 3) + 60 : Math.round(text.length / 2) + 40;
  const d = await post({ op: 'run', messages: m, max: Math.min(700, max), temp: temp || 0 }, onToken);
  let out = clean(d.text, text, task); if (task === 'humanize' || task === 'friendly') out = humanPass(out);
  return out;
}
async function run(task, text, opt = {}) { // → { text, refused } ; opt.onToken(piece) streams, opt.retry asks for a different wording
  if (!TASKS[task]) throw new Error('Unknown AI task'); await ensure();
  const src = String(text || '').trim(); if (!src) return { text: '', refused: 'no text' };
  if (task === 'summary' || task === 'bullets' || task === 'continue') {
    const t = await once(task, src.length > 2400 ? src.slice(task === 'continue' ? -2400 : 0, task === 'continue' ? undefined : 2400) : src, opt.onToken, opt.retry ? 0.7 : 0);
    return { text: t, refused: sane(t, src, task) };
  }
  const parts = chunks(src), outs = [];
  for (const [i, c] of parts.entries()) {
    if (opt.onToken && i) opt.onToken(' ');
    let t = await once(task, c, opt.onToken, opt.retry ? 0.7 : 0), why = sane(t, c, task);
    if (why && !opt.retry) { t = await once(task, c, null, 0.6); why = sane(t, c, task); } // one more try with a little variety
    outs.push(why ? c : t); if (why && parts.length === 1) return { text: t, refused: why };
  }
  const out = outs.join(' ').replace(/\s{2,}/g, ' ').trim();
  return { text: out, refused: out === src ? 'no change' : '' };
}

/* ── the progress line any panel can show ── */
function statusHtml() {
  if (A.state === 'loading') return `<div class="ai-load"><span>Starting the offline AI… ${A.pct ? A.pct + '%' : ''}</span><i style="--p:${A.pct}%"></i><small>First time only takes a few seconds; it runs on this device.</small></div>`;
  if (A.state === 'error') return `<div class="ai-load err">The offline AI could not start here: ${esc(A.err)}</div>`;
  return '';
}

X.ai = { TASKS, run, ensure, stop, state: () => A, supported, statusHtml, clean, sane };
})();
