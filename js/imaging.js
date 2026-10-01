/* XD3 PDF Editor — image tools: filters / crop / flip, offline background remover, document scanner. */
(() => {
'use strict';
const X = window.XD3, { S, $, $$, toast, busy, dialog, clamp, uid } = X;

const loadImg = url => new Promise((res, rej) => { const im = new Image(); im.onload = () => res(im); im.onerror = () => rej(new Error('Could not load image')); im.src = url; });
function canvasOf(src, maxDim = 0) {
  const w0 = src.naturalWidth || src.videoWidth || src.width, h0 = src.naturalHeight || src.videoHeight || src.height, f = maxDim ? Math.min(1, maxDim / Math.max(w0, h0)) : 1;
  const c = document.createElement('canvas'); c.width = Math.max(1, Math.round(w0 * f)); c.height = Math.max(1, Math.round(h0 * f));
  c.getContext('2d', { willReadFrequently: true }).drawImage(src, 0, 0, c.width, c.height); return c;
}
function geom(c, { rot = 0, flipH, flipV, crop } = {}) { // rot = quarter turns clockwise; crop = normalised rect in the turned image
  const sw = rot % 2 ? c.height : c.width, sh = rot % 2 ? c.width : c.height;
  let o = document.createElement('canvas'); o.width = sw; o.height = sh;
  const x = o.getContext('2d', { willReadFrequently: true }); x.translate(sw / 2, sh / 2); x.scale(flipH ? -1 : 1, flipV ? -1 : 1); x.rotate(rot * Math.PI / 2); x.drawImage(c, -c.width / 2, -c.height / 2);
  if (crop) {
    const k = document.createElement('canvas'); k.width = Math.max(1, Math.round(crop.w * sw)); k.height = Math.max(1, Math.round(crop.h * sh));
    k.getContext('2d', { willReadFrequently: true }).drawImage(o, Math.round(crop.x * sw), Math.round(crop.y * sh), k.width, k.height, 0, 0, k.width, k.height); o = k;
  }
  return o;
}

/* ───────── filters ───────── */
function paperMap(d, w, h) { // per-pixel estimate of the paper / lighting colour (brightest value per cell, smoothed)
  const cs = Math.max(8, Math.round(Math.max(w, h) / 40)), gw = Math.ceil(w / cs), gh = Math.ceil(h / cs);
  const g = document.createElement('canvas'); g.width = gw; g.height = gh; const gx = g.getContext('2d'), gi = gx.createImageData(gw, gh);
  for (let cy = 0; cy < gh; cy++) for (let cx = 0; cx < gw; cx++) {
    let r = 0, gg = 0, b = 0; const x1 = Math.min(w, (cx + 1) * cs), y1 = Math.min(h, (cy + 1) * cs);
    for (let y = cy * cs; y < y1; y += 2) for (let x = cx * cs; x < x1; x += 2) { const i = (y * w + x) * 4; if (d[i] > r) r = d[i]; if (d[i + 1] > gg) gg = d[i + 1]; if (d[i + 2] > b) b = d[i + 2]; }
    const j = (cy * gw + cx) * 4; gi.data[j] = r; gi.data[j + 1] = gg; gi.data[j + 2] = b; gi.data[j + 3] = 255;
  }
  gx.putImageData(gi, 0, 0);
  const o = document.createElement('canvas'); o.width = w; o.height = h; const ox = o.getContext('2d', { willReadFrequently: true });
  ox.imageSmoothingEnabled = true; ox.imageSmoothingQuality = 'high'; ox.drawImage(g, 0, 0, gw * cs, gh * cs);
  return ox.getImageData(0, 0, w, h).data;
}
function applyFx(c, fx) { // in place. fx: { mode: none|doc|bw|gray|sepia|invert, bright, contrast, sat (-100..100), sharpen }
  const ctx = c.getContext('2d', { willReadFrequently: true }), w = c.width, h = c.height, id = ctx.getImageData(0, 0, w, h), d = id.data, mode = fx.mode || 'none';
  const bg = mode === 'doc' || mode === 'bw' ? paperMap(d, w, h) : null;
  const br = (fx.bright || 0) * 1.6, cv = (fx.contrast || 0) * 1.8, cf = (259 * (cv + 255)) / (255 * (259 - cv)), sf = 1 + (fx.sat || 0) / 100;
  for (let i = 0; i < d.length; i += 4) {
    let r = d[i], g = d[i + 1], b = d[i + 2];
    if (bg) { // flatten lighting / shadows, then stretch so paper goes white and ink goes dark
      r = (Math.min(255, r * 255 / Math.max(bg[i], 60)) - 30) * 1.2; g = (Math.min(255, g * 255 / Math.max(bg[i + 1], 60)) - 30) * 1.2; b = (Math.min(255, b * 255 / Math.max(bg[i + 2], 60)) - 30) * 1.2;
    }
    if (mode === 'bw') { r = g = b = (0.299 * r + 0.587 * g + 0.114 * b - 135) * 6; }
    else if (mode === 'gray') { r = g = b = 0.299 * r + 0.587 * g + 0.114 * b; }
    else if (mode === 'sepia') { const R = 0.393 * r + 0.769 * g + 0.189 * b, G = 0.349 * r + 0.686 * g + 0.168 * b, B = 0.272 * r + 0.534 * g + 0.131 * b; r = R; g = G; b = B; }
    else if (mode === 'invert') { r = 255 - r; g = 255 - g; b = 255 - b; }
    if (br || cv) { r = cf * (r - 128) + 128 + br; g = cf * (g - 128) + 128 + br; b = cf * (b - 128) + 128 + br; }
    if (sf !== 1) { const y = 0.299 * r + 0.587 * g + 0.114 * b; r = y + (r - y) * sf; g = y + (g - y) * sf; b = y + (b - y) * sf; }
    d[i] = r; d[i + 1] = g; d[i + 2] = b;
  }
  if (fx.sharpen) { // 3×3 unsharp kernel
    const s = new Uint8ClampedArray(d), row = w * 4;
    for (let y = 1; y < h - 1; y++) for (let x = 1; x < w - 1; x++) { const i = (y * w + x) * 4; for (let k = 0; k < 3; k++) d[i + k] = 5 * s[i + k] - s[i + k - 4] - s[i + k + 4] - s[i + k - row] - s[i + k + row]; }
  }
  ctx.putImageData(id, 0, 0); return c;
}
const MODES = [['none', 'Original'], ['doc', 'Enhance'], ['gray', 'Grayscale'], ['bw', 'Black & white'], ['sepia', 'Sepia'], ['invert', 'Invert']];
const selImage = () => { const s = X.selOne(); if (!s || s.o.type !== 'image') { toast('Select an image first (use the Select tool, or Edit object to pick one up from the page)'); return null; } return s; };
function setAsset(s, canvas, forcePng) { // bake an edited canvas back into the selected image object (undoable)
  const o = s.o, jpeg = !forcePng && /^data:image\/jpe?g/.test(S.assets[o.asset] || '');
  X.pushUndo(); o.asset = X.addAsset(canvas.toDataURL(jpeg ? 'image/jpeg' : 'image/png', 0.9)); X.drawObjs(s.p);
}

async function imgEdit() {
  const s = selImage(); if (!s) return;
  const img = await loadImg(S.assets[s.o.asset]), small = canvasOf(img, 720);
  const st = { rot: 0, flipH: false, flipV: false, crop: null, fx: { mode: 'none', bright: 0, contrast: 0, sat: 0, sharpen: false } };
  const r = await dialog('Edit image', `
    <div class="imgprev"><canvas id="iePrev"></canvas></div>
    <div class="chips" id="ieModes">${MODES.map(([v, l]) => `<button type="button" data-m="${v}">${l}</button>`).join('')}</div>
    <div class="row"><label>Brightness<input type="range" min="-100" max="100" value="0" data-k="bright"></label>
      <label>Contrast<input type="range" min="-100" max="100" value="0" data-k="contrast"></label>
      <label>Saturation<input type="range" min="-100" max="100" value="0" data-k="sat"></label></div>
    <div class="chips"><button type="button" data-g="rot">⟳ Rotate</button><button type="button" data-g="flipH">⇋ Flip horizontal</button><button type="button" data-g="flipV">⇵ Flip vertical</button>
      <button type="button" data-g="sharpen">Sharpen</button><button type="button" data-g="nocrop">Clear crop</button><button type="button" data-g="reset">Reset all</button></div>
    <p class="mut">Drag on the picture to crop it.</p>`, {
    ok: 'Apply', wide: true, onOpen: d => {
      const pv = $('#iePrev', d), px = pv.getContext('2d'); let base, fxc, dragA = null;
      const paint = () => {
        px.clearRect(0, 0, pv.width, pv.height); px.drawImage(fxc, 0, 0);
        if (st.crop) { const c = st.crop, x = c.x * pv.width, y = c.y * pv.height, w = c.w * pv.width, h = c.h * pv.height;
          px.fillStyle = 'rgba(0,0,0,.55)'; px.beginPath(); px.rect(0, 0, pv.width, pv.height); px.rect(x, y, w, h); px.fill('evenodd');
          px.strokeStyle = '#fff'; px.lineWidth = 2; px.setLineDash([6, 4]); px.strokeRect(x, y, w, h); px.setLineDash([]); }
      };
      const refx = () => { fxc = applyFx(geom(base), st.fx); paint(); };
      const rebase = () => { base = geom(small, { rot: st.rot, flipH: st.flipH, flipV: st.flipV }); pv.width = base.width; pv.height = base.height; refx(); };
      const sync = () => { $$('#ieModes button', d).forEach(b => b.classList.toggle('active', b.dataset.m === st.fx.mode)); $('[data-g=sharpen]', d).classList.toggle('active', st.fx.sharpen); $$('[data-k]', d).forEach(i => i.value = st.fx[i.dataset.k]); };
      $('#ieModes', d).onclick = e => { const b = e.target.closest('[data-m]'); if (!b) return; st.fx.mode = b.dataset.m; sync(); refx(); };
      $$('[data-k]', d).forEach(i => i.oninput = () => { st.fx[i.dataset.k] = +i.value; refx(); });
      $$('[data-g]', d).forEach(b => b.onclick = () => {
        const g = b.dataset.g;
        if (g === 'rot') { st.rot = (st.rot + 1) % 4; st.crop = null; rebase(); } else if (g === 'flipH' || g === 'flipV') { st[g] = !st[g]; st.crop = null; rebase(); }
        else if (g === 'sharpen') { st.fx.sharpen = !st.fx.sharpen; sync(); refx(); } else if (g === 'nocrop') { st.crop = null; paint(); }
        else { Object.assign(st, { rot: 0, flipH: false, flipV: false, crop: null }); st.fx = { mode: 'none', bright: 0, contrast: 0, sat: 0, sharpen: false }; sync(); rebase(); }
      });
      const norm = e => { const b = pv.getBoundingClientRect(); return [clamp((e.clientX - b.left) / b.width, 0, 1), clamp((e.clientY - b.top) / b.height, 0, 1)]; };
      pv.onpointerdown = e => { dragA = norm(e); pv.setPointerCapture(e.pointerId); };
      pv.onpointermove = e => { if (!dragA) return; const p = norm(e), x = Math.min(dragA[0], p[0]), y = Math.min(dragA[1], p[1]), w = Math.abs(p[0] - dragA[0]), h = Math.abs(p[1] - dragA[1]); st.crop = w > 0.03 && h > 0.03 ? { x, y, w, h } : null; paint(); };
      pv.onpointerup = pv.onpointercancel = () => dragA = null;
      sync(); rebase();
    }
  });
  if (!r) return;
  await busy(async () => {
    const full = applyFx(geom(canvasOf(img, 4000), st), st.fx), o = s.o, cx = o.x + o.w / 2, cy = o.y + o.h / 2;
    let w = o.w, h = o.h; if (st.rot % 2) [w, h] = [h, w]; if (st.crop) { w *= st.crop.w; h *= st.crop.h; }
    setAsset(s, full); Object.assign(o, { w, h, x: cx - w / 2, y: cy - h / 2 }); X.drawObjs(s.p);
  }, 'Applying…');
}

/* ───────── background remover (colour based, fully offline) ───────── */
function bgAlpha(d, W, H, st) { // → Uint8ClampedArray alpha per pixel. st: { mode: edge|all, tol (%), soft (px), seeds: [[x,y] normalised] }
  const n = W * H, A = new Uint8ClampedArray(n).fill(255), tol = st.tol / 100 * 441.7, t2 = tol * tol;
  const avgAt = (px, py) => { let r = 0, g = 0, b = 0, c = 0; for (let y = Math.max(0, py - 1); y <= Math.min(H - 1, py + 1); y++) for (let x = Math.max(0, px - 1); x <= Math.min(W - 1, px + 1); x++) { const i = (y * W + x) * 4; r += d[i]; g += d[i + 1]; b += d[i + 2]; c++; } return [r / c, g / c, b / c]; };
  // dominant colour along the image border = the background
  const bins = {}; let best = null;
  const edge = (x, y) => { const i = (y * W + x) * 4, k = (d[i] >> 4) << 8 | (d[i + 1] >> 4) << 4 | d[i + 2] >> 4, e = bins[k] || (bins[k] = [0, 0, 0, 0]); e[0]++; e[1] += d[i]; e[2] += d[i + 1]; e[3] += d[i + 2]; if (!best || e[0] > best[0]) best = e; };
  for (let x = 0; x < W; x++) { edge(x, 0); edge(x, H - 1); } for (let y = 0; y < H; y++) { edge(0, y); edge(W - 1, y); }
  const ref0 = [best[1] / best[0], best[2] / best[0], best[3] / best[0]];
  const seeds = st.seeds.map(([sx, sy]) => { const px = clamp(Math.round(sx * W), 0, W - 1), py = clamp(Math.round(sy * H), 0, H - 1); return { i: py * W + px, c: avgAt(px, py) }; });
  const dist2 = (p, c) => { const j = p * 4, a = d[j] - c[0], b = d[j + 1] - c[1], e = d[j + 2] - c[2]; return a * a + b * b + e * e; };
  if (st.mode === 'all') { // remove the colour wherever it appears, with a soft ramp (ideal for signatures, stamps, logos)
    const refs = [ref0, ...seeds.map(s => s.c)], hi = tol * 1.7;
    for (let p = 0; p < n; p++) { let m = 1e9; for (const c of refs) { const v = dist2(p, c); if (v < m) m = v; } m = Math.sqrt(m); A[p] = m <= tol ? 0 : m >= hi ? 255 : (m - tol) / (hi - tol) * 255; }
  } else { // flood from the borders (and from every clicked point) so matching colours inside the subject are kept
    const stack = new Int32Array(n); let sp = 0;
    const flood = (starts, c) => {
      sp = 0; for (const p of starts) if (A[p] && dist2(p, c) <= t2) { A[p] = 0; stack[sp++] = p; }
      while (sp) {
        const p = stack[--sp], x = p % W;
        if (x > 0 && A[p - 1] && dist2(p - 1, c) <= t2) { A[p - 1] = 0; stack[sp++] = p - 1; }
        if (x < W - 1 && A[p + 1] && dist2(p + 1, c) <= t2) { A[p + 1] = 0; stack[sp++] = p + 1; }
        if (p >= W && A[p - W] && dist2(p - W, c) <= t2) { A[p - W] = 0; stack[sp++] = p - W; }
        if (p < n - W && A[p + W] && dist2(p + W, c) <= t2) { A[p + W] = 0; stack[sp++] = p + W; }
      }
    };
    const border = []; for (let x = 0; x < W; x++) border.push(x, (H - 1) * W + x); for (let y = 0; y < H; y++) border.push(y * W, y * W + W - 1);
    flood(border, ref0);
    for (const s of seeds) { A[s.i] = 255; flood([s.i], s.c); }
    const r = Math.round(st.soft);
    if (r > 0) { // box-blur the mask, then pull the edge inwards a little: feathers it and trims the colour fringe
      const T = new Float32Array(n), win = 2 * r + 1;
      for (let y = 0; y < H; y++) { let s = 0; const o = y * W; for (let x = -r; x <= r; x++) s += A[o + clamp(x, 0, W - 1)]; for (let x = 0; x < W; x++) { T[o + x] = s / win; s += A[o + Math.min(W - 1, x + r + 1)] - A[o + Math.max(0, x - r)]; } }
      for (let x = 0; x < W; x++) { let s = 0; for (let y = -r; y <= r; y++) s += T[clamp(y, 0, H - 1) * W + x]; for (let y = 0; y < H; y++) { const v = s / win; A[y * W + x] = Math.min(A[y * W + x], (v - 110) * 255 / 145); s += T[Math.min(H - 1, y + r + 1) * W + x] - T[Math.max(0, y - r) * W + x]; } }
    }
  }
  for (let p = 0; p < n; p++) A[p] = A[p] * d[p * 4 + 3] / 255; // keep any transparency the image already had
  return A;
}
function cutOut(c, st) { const x = c.getContext('2d', { willReadFrequently: true }), id = x.getImageData(0, 0, c.width, c.height), A = bgAlpha(id.data, c.width, c.height, st); for (let p = 0; p < A.length; p++) id.data[p * 4 + 3] = A[p]; x.putImageData(id, 0, 0); return c; }

async function bgRemove() {
  const s = selImage(); if (!s) return;
  const img = await loadImg(S.assets[s.o.asset]), small = canvasOf(img, 800), st = { mode: 'edge', tol: 16, soft: 2, seeds: [] };
  const r = await dialog('Remove background', `
    <div class="imgprev"><canvas id="bgPrev" class="pick"></canvas></div>
    <div class="chips" id="bgModes"><button type="button" data-m="edge">Around the subject</button><button type="button" data-m="all">Colour everywhere (signatures, logos)</button></div>
    <div class="row"><label>Tolerance<input type="range" min="2" max="60" value="16" data-k="tol"></label><label>Smooth edges<input type="range" min="0" max="6" value="2" data-k="soft"></label></div>
    <div class="chips"><button type="button" data-g="undo">Undo last click</button><button type="button" data-g="reset">Reset</button></div>
    <p class="mut">Click any leftover area in the picture to remove that colour too. Works best on plain or evenly lit backgrounds; the checkerboard shows what becomes transparent.</p>`, {
    ok: 'Apply', wide: true, onOpen: d => {
      const pv = $('#bgPrev', d); pv.width = small.width; pv.height = small.height; let t = 0;
      const paint = () => { const c = canvasOf(small); cutOut(c, st); const x = pv.getContext('2d'); x.clearRect(0, 0, pv.width, pv.height); x.drawImage(c, 0, 0); };
      const later = () => { clearTimeout(t); t = setTimeout(paint, 30); };
      const sync = () => $$('#bgModes button', d).forEach(b => b.classList.toggle('active', b.dataset.m === st.mode));
      $('#bgModes', d).onclick = e => { const b = e.target.closest('[data-m]'); if (!b) return; st.mode = b.dataset.m; sync(); paint(); };
      $$('[data-k]', d).forEach(i => i.oninput = () => { st[i.dataset.k] = +i.value; later(); });
      $('[data-g=undo]', d).onclick = () => { st.seeds.pop(); paint(); };
      $('[data-g=reset]', d).onclick = () => { st.seeds = []; paint(); };
      pv.onclick = e => { const b = pv.getBoundingClientRect(); st.seeds.push([(e.clientX - b.left) / b.width, (e.clientY - b.top) / b.height]); paint(); };
      sync(); paint();
    }
  });
  if (!r) return;
  await busy(async () => {
    const full = canvasOf(img, 2400), k = full.width / small.width; // same settings, scaled to full resolution
    setAsset(s, cutOut(full, { ...st, soft: st.soft * Math.max(1, k * 0.6) }), true);
  }, 'Removing background…');
  toast('Background removed');
}

/* ───────── document scanner ───────── */
function solve(M, v) { // Gaussian elimination, 8×8
  const n = v.length, A = M.map((r, i) => [...r, v[i]]);
  for (let c = 0; c < n; c++) {
    let p = c; for (let r = c + 1; r < n; r++) if (Math.abs(A[r][c]) > Math.abs(A[p][c])) p = r;
    [A[c], A[p]] = [A[p], A[c]]; if (Math.abs(A[c][c]) < 1e-12) return null;
    for (let r = 0; r < n; r++) if (r !== c) { const f = A[r][c] / A[c][c]; for (let k = c; k <= n; k++) A[r][k] -= f * A[c][k]; }
  }
  return A.map((r, i) => r[n] / r[i]);
}
function warp(c, q) { // q = corners TL,TR,BR,BL (normalised) → straightened rectangle
  const P = q.map(([x, y]) => [x * c.width, y * c.height]), dist = (a, b) => Math.hypot(a[0] - b[0], a[1] - b[1]);
  let W = Math.round(Math.max(dist(P[0], P[1]), dist(P[3], P[2]))), H = Math.round(Math.max(dist(P[0], P[3]), dist(P[1], P[2])));
  const f = Math.min(1, 2600 / Math.max(W, H)); W = Math.max(8, Math.round(W * f)); H = Math.max(8, Math.round(H * f));
  const D = [[0, 0], [W, 0], [W, H], [0, H]], M = [], v = [];
  for (let i = 0; i < 4; i++) { const [x, y] = D[i], [u, w] = P[i]; M.push([x, y, 1, 0, 0, 0, -x * u, -y * u]); v.push(u); M.push([0, 0, 0, x, y, 1, -x * w, -y * w]); v.push(w); }
  const h = solve(M, v); if (!h) return c;
  const sd = c.getContext('2d', { willReadFrequently: true }).getImageData(0, 0, c.width, c.height).data, o = document.createElement('canvas'); o.width = W; o.height = H;
  const ox = o.getContext('2d', { willReadFrequently: true }), od = ox.createImageData(W, H), dd = od.data, sw = c.width, sh = c.height;
  for (let y = 0; y < H; y++) for (let x = 0; x < W; x++) {
    const z = h[6] * x + h[7] * y + 1, u = clamp((h[0] * x + h[1] * y + h[2]) / z, 0, sw - 1.001), w = clamp((h[3] * x + h[4] * y + h[5]) / z, 0, sh - 1.001);
    const x0 = u | 0, y0 = w | 0, fx = u - x0, fy = w - y0, i = (y0 * sw + x0) * 4, j = (y * W + x) * 4;
    for (let k = 0; k < 3; k++) dd[j + k] = (sd[i + k] * (1 - fx) + sd[i + 4 + k] * fx) * (1 - fy) + (sd[i + sw * 4 + k] * (1 - fx) + sd[i + sw * 4 + 4 + k] * fx) * fy;
    dd[j + 3] = 255;
  }
  ox.putImageData(od, 0, 0); return o;
}
async function scan() {
  let stream = null, added = 0, queue = [], cur = null, corners, firstNew = -1; const fresh = !S.pages.length;
  const FULL = () => [[0, 0], [1, 0], [1, 1], [0, 1]];
  await dialog('Scan document', `
    <div id="scCam"><div class="imgprev"><video id="scVideo" autoplay playsinline muted></video></div><p id="scMsg" class="mut"></p>
      <div class="row"><button type="button" class="btn primary" id="scShot">Capture page</button><button type="button" class="btn line" id="scImport">Import photos…</button></div></div>
    <div id="scAdj" hidden><div class="imgprev"><div class="scwrap"><canvas id="scPrev"></canvas><svg viewBox="0 0 100 100" preserveAspectRatio="none"><polygon id="scPoly"/></svg>
        <i class="sch"></i><i class="sch"></i><i class="sch"></i><i class="sch"></i></div></div>
      <div class="row"><label>Look<select id="scMode"><option value="doc">Enhanced colour</option><option value="bw">Black &amp; white</option><option value="gray">Grayscale</option><option value="none">Original photo</option></select></label>
        <button type="button" class="btn line" id="scRot">⟳ Rotate</button><button type="button" class="btn line" id="scFull">Reset corners</button></div>
      <p class="mut">Drag the four corners onto the edges of the paper to straighten and crop it.</p>
      <div class="row"><button type="button" class="btn primary" id="scAdd">Add page</button><button type="button" class="btn line" id="scSkip">Discard</button></div></div>
    <p id="scCount" class="mut"></p><input type="file" id="scFile" accept="image/*" multiple hidden>`, {
    ok: 'Done', cancel: false, wide: true,
    onOpen: d => {
      const video = $('#scVideo', d), msg = $('#scMsg', d), cam = $('#scCam', d), adj = $('#scAdj', d), pv = $('#scPrev', d), hs = $$('.sch', d), poly = $('#scPoly', d), mode = $('#scMode', d);
      const startCam = async () => {
        try { stream = await navigator.mediaDevices.getUserMedia({ video: { facingMode: { ideal: 'environment' }, width: { ideal: 2560 }, height: { ideal: 1440 } } }); video.srcObject = stream; msg.textContent = 'Hold the page flat and fill the frame, then capture.'; }
        catch (e) { video.parentNode.hidden = true; $('#scShot', d).hidden = true; msg.textContent = 'No camera available here — use “Import photos” to scan from pictures instead.'; }
      };
      const place = () => { hs.forEach((h, i) => { h.style.left = corners[i][0] * 100 + '%'; h.style.top = corners[i][1] * 100 + '%'; }); poly.setAttribute('points', corners.map(c => c[0] * 100 + ',' + c[1] * 100).join(' ')); };
      const draw = () => { const c = applyFx(canvasOf(cur, 760), { mode: mode.value }); pv.width = c.width; pv.height = c.height; pv.getContext('2d').drawImage(c, 0, 0); place(); };
      const adjust = c => { cur = canvasOf(c, 2600); corners = FULL(); cam.hidden = true; adj.hidden = false; draw(); };
      const next = async () => {
        $('#scCount', d).textContent = added ? `${added} page(s) added${queue.length ? ` · ${queue.length} photo(s) waiting` : ''}` : '';
        if (queue.length) { try { adjust(await loadImg(URL.createObjectURL(queue.shift()))); } catch (e) { toast('Could not read that image', true); next(); } }
        else { adj.hidden = true; cam.hidden = false; }
      };
      hs.forEach((h, i) => {
        h.onpointerdown = e => { h.setPointerCapture(e.pointerId); e.preventDefault(); };
        h.onpointermove = e => { if (!h.hasPointerCapture(e.pointerId)) return; const b = pv.getBoundingClientRect(); corners[i] = [clamp((e.clientX - b.left) / b.width, 0, 1), clamp((e.clientY - b.top) / b.height, 0, 1)]; place(); };
      });
      $('#scShot', d).onclick = () => { if (video.videoWidth) adjust(video); };
      $('#scImport', d).onclick = () => $('#scFile', d).click();
      $('#scFile', d).onchange = e => { queue.push(...e.target.files); e.target.value = ''; if (adj.hidden) next(); else $('#scCount', d).textContent = `${added} page(s) added · ${queue.length} photo(s) waiting`; };
      mode.onchange = draw;
      $('#scRot', d).onclick = () => { cur = geom(cur, { rot: 1 }); corners = FULL(); draw(); };
      $('#scFull', d).onclick = () => { corners = FULL(); place(); };
      $('#scSkip', d).onclick = next;
      $('#scAdd', d).onclick = () => {
        const moved = corners.some((c, i) => Math.abs(c[0] - FULL()[i][0]) + Math.abs(c[1] - FULL()[i][1]) > 0.004);
        const out = applyFx(moved ? warp(cur, corners) : canvasOf(cur), { mode: mode.value });
        if (!added) { if (fresh) { X.resetDoc(); S.fileName = 'scan.pdf'; } else X.pushUndo(); firstNew = S.pages.length; }
        const W = out.width > out.height ? X.A4[1] : X.A4[0], H = W * out.height / out.width;
        S.pages.push({ id: uid(), src: -1, w: W, h: H, rot: 0, objs: [{ id: uid(), type: 'image', asset: X.addAsset(out.toDataURL('image/jpeg', 0.86)), x: 0, y: 0, w: W, h: H, opacity: 1 }] });
        added++; next();
      };
      startCam();
    },
    onClose: () => { if (stream) stream.getTracks().forEach(t => t.stop()); },
  });
  if (!added) return;
  S.dirty = true; X.buildPages();
  if (fresh) { S.zoom = clamp(X.fitZoom(), 0.25, 1.25); X.layout(); X.setTool('select'); }
  X.goToPage(firstNew); X.syncControls(); X.updateUI(); X.scheduleSave(); toast(`${added} scanned page(s) added`);
}

Object.assign(X.ACT, { imgedit: imgEdit, bgremove: bgRemove, scan });
Object.assign(X, { applyFx, canvasOf, loadImg, cutOut, warp });
})();
