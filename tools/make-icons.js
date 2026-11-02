/* Renders every app icon from icons/logo.svg (and icons/logo-small.svg for 16–32 px) into tests/out/.
   Use: run serve.py, open the editor, load this file with a <script src="tools/make-icons.js"> tag (or paste it into the console),
   then run  python tools/make-ico.py  — it copies the PNGs into icons/ and builds desktop/icon.ico. */
(async () => {
  const load = async svg => { const img = new Image(); img.src = URL.createObjectURL(new Blob([svg], { type: 'image/svg+xml' })); await img.decode(); return img; };
  const full = await (await fetch('icons/logo.svg?' + Date.now())).text(), small = await (await fetch('icons/logo-small.svg?' + Date.now())).text();
  // full-bleed version for maskable / Apple icons: square background, artwork shrunk into the safe zone
  const bleed = k => full.replace('rx="14" fill="url(#xbg)"/>', `fill="url(#xbg)"/><g transform="translate(${32 - 32 * k} ${32 - 32 * k}) scale(${k})">`).replace('</svg>', '</g></svg>');
  const F = await load(full), Sm = await load(small), M = await load(bleed(0.78)), A = await load(bleed(0.9));
  const png = (img, n) => { const c = document.createElement('canvas'); c.width = c.height = n; const x = c.getContext('2d'); x.imageSmoothingQuality = 'high'; x.drawImage(img, 0, 0, n, n); return new Promise(r => c.toBlob(r, 'image/png')); };
  const put = async (name, blob) => (await fetch('tests/out/' + name, { method: 'PUT', body: blob })).status === 204;
  const ok = [await put('icon-192.png', await png(F, 192)), await put('icon-512.png', await png(F, 512)), await put('icon-512-maskable.png', await png(M, 512)), await put('apple-touch-icon.png', await png(A, 180))];
  for (const n of [16, 20, 24, 32, 40, 48, 64, 96, 128, 256]) ok.push(await put(`ico-${n}.png`, await png(n <= 32 ? Sm : F, n)));
  console.log(ok.every(Boolean) ? 'Icons written to tests/out — now run: python tools/make-ico.py' : 'Some icons could not be written (is serve.py running?)');
})();
