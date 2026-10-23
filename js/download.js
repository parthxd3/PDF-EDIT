/* XD3 PDF Editor — "Download for Windows": offers the installer / portable ZIP that the website carries in downloads/
   (made by Build EXE.bat, listed in downloads/manifest.json). Shown only on the hosted copy, never inside the desktop app itself. */
(() => {
'use strict';
const X = window.XD3, { $, dialog, esc, icons } = X;
if (!/^https?:$/.test(location.protocol)) return; // the desktop app and file:// copies have nothing to offer
const isWin = /Windows/i.test(navigator.userAgent), is64 = /WOW64|Win64|x64|x86_64/i.test(navigator.userAgent) || (navigator.userAgentData && navigator.userAgentData.platform === 'Windows');
const mb = n => (n / 1048576).toFixed(0) + ' MB';
let manifest = null;
async function load() {
  if (manifest !== null) return manifest;
  try { const r = await fetch('downloads/manifest.json', { cache: 'no-cache' }); manifest = r.ok ? await r.json() : false; } catch (e) { manifest = false; }
  if (manifest) { // only the files that were actually uploaded (people often put just the installers on the site)
    const here = await Promise.all((manifest.files || []).map(f => fetch('downloads/' + encodeURIComponent(f.file), { method: 'HEAD', cache: 'no-cache' }).then(r => r.ok, () => false)));
    manifest.files = (manifest.files || []).filter((f, i) => here[i]);
    if (!manifest.files.length) manifest = false;
  }
  return manifest;
}
function row(f) {
  const label = f.kind === 'installer' ? 'Installer' : 'Portable ZIP', bits = f.arch === 'ia32' ? '32-bit' : '64-bit';
  const hint = f.kind === 'installer' ? 'Installs with Start-menu and desktop shortcuts; PDFs get an "Open with XD3 PDF Editor" entry.' : 'No installation: unzip anywhere and run "XD3 PDF Editor.exe". Good for USB sticks and locked-down PCs.';
  const rec = isWin && ((is64 && f.arch === 'x64') || (!is64 && f.arch === 'ia32')) && f.kind === 'installer';
  return `<a class="dlrow${rec ? ' rec' : ''}" href="downloads/${encodeURIComponent(f.file)}" download><i data-icon="${f.kind === 'installer' ? 'download' : 'folder'}"></i><div><b>${label} · ${bits}</b>${rec ? '<span class="dlrec">Recommended for this PC</span>' : ''}<small>${hint}</small></div><em>${mb(f.size)}</em></a>`;
}
async function show() {
  const m = await load();
  if (!m) return dialog('Download for Windows', `<p>The Windows app is not on this site yet.</p><p class="mut">Run <b>Build EXE.bat</b> on the project, then <b>Build Hostinger ZIP.bat</b> and upload — the installer and portable ZIP are then offered here.</p>`, { cancel: false });
  const files = [...m.files].sort((a, b) => (a.arch === 'x64' ? 0 : 1) - (b.arch === 'x64' ? 0 : 1) || (a.kind === 'installer' ? 0 : 1) - (b.kind === 'installer' ? 0 : 1));
  await dialog(`Download for Windows <small class="mut" style="font-weight:400">v${esc(m.version)}</small>`, `
    <p>The same editor as a program on your PC — opens PDFs by double-click, works without a browser or internet, and runs on <b>Windows 7, 8, 10 and 11</b>.</p>
    <div class="dllist">${files.map(row).join('')}</div>
    <p class="mut" style="margin-top:10px">Not sure? Most PCs made after 2010 are 64-bit. The installer is unsigned, so Windows SmartScreen may ask once — choose <i>More info ▸ Run anyway</i>. Everything stays on your device; the app never connects to the internet.</p>`,
    { ok: 'Close', cancel: false, wide: true, onOpen: d => icons(d) });
}
X.ACT.dlwin = show; X.NO_DOC.add('dlwin');
load().then(m => { if (!m) $$('[data-act=dlwin]').forEach(b => { b.hidden = true; }); }); // (no files on this site → the button is not shown)
function $$(s) { return [...document.querySelectorAll(s)]; }
})();
