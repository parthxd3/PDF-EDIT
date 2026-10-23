// XD3 PDF Editor — desktop shell (Electron). © XD3Labs / XDCybertech Pvt Ltd.
// Serves the bundled web app from a private app:// origin, so it runs fully offline with no local web server.
// One running copy: PDFs and Word files double-clicked while it is open come to it (an idle window, or a new one).
const { app, BrowserWindow, protocol, Menu, shell, session, dialog, ipcMain } = require('electron');
const path = require('path'), fs = require('fs'), os = require('os'), crypto = require('crypto');

const ROOT = path.join(__dirname, 'www');
const OPENABLE = /\.(pdf|docx|txt|md|png|jpe?g|gif|bmp|webp)$/i;
const SELFTEST = !!process.env.XD3_SELFTEST;
const fileIn = (argv, cwd) => { const a = argv.find(x => OPENABLE.test(x) && !x.startsWith('-')); return a ? path.resolve(cwd || process.cwd(), a) : ''; };
const opened = new Map(); // token → file path; the page fetches it from app://xd3/__open__/<token>/<name>
const portable = !fs.existsSync(path.join(path.dirname(process.execPath), 'Uninstall XD3 PDF Editor.exe'));
const winreg = require('./winreg.js')({ exe: process.execPath, root: process.env.XD3_REG_ROOT || 'HKCU\\Software' });
const canRegister = process.platform === 'win32' && (app.isPackaged || !!process.env.XD3_REG_ROOT);
const osInfo = () => { const [major, minor, build] = os.release().split('.').map(Number); return { major, minor, build, name: major >= 10 ? (build >= 22000 ? 'Windows 11' : 'Windows 10') : { 1: 'Windows 7', 2: 'Windows 8', 3: 'Windows 8.1' }[minor] || 'Windows' }; };

protocol.registerSchemesAsPrivileged([{ scheme: 'app', privileges: { standard: true, secure: true, supportFetchAPI: true, stream: true, corsEnabled: true } }]);
if (process.platform === 'win32') app.setAppUserModelId('in.xdcybertech.xd3pdfeditor');

function urlFor(file) {
  if (!file) return 'app://xd3/index.html' + (SELFTEST ? '?selftest=1' : '');
  const t = crypto.randomBytes(9).toString('hex'); opened.set(t, file);
  return `app://xd3/index.html?open=${encodeURIComponent(path.basename(file))}&f=${t}`;
}
function createWindow(file) {
  const last = BrowserWindow.getFocusedWindow() || BrowserWindow.getAllWindows().pop(), at = last && !last.isMaximized() ? last.getPosition() : null;
  const win = new BrowserWindow({
    width: 1360, height: 860, minWidth: 420, minHeight: 480, backgroundColor: '#0e1016', autoHideMenuBar: true, show: false,
    ...(at ? { x: at[0] + 28, y: at[1] + 28 } : {}), title: 'XD3 PDF Editor', icon: path.join(__dirname, 'icon.ico'),
    webPreferences: { contextIsolation: true, nodeIntegration: false, sandbox: true, plugins: true, preload: path.join(__dirname, 'preload.js') },
  });
  win.once('ready-to-show', () => win.show());
  win.loadURL(urlFor(file));
  win.webContents.setWindowOpenHandler(({ url }) => {
    if (/^(https?|mailto):/i.test(url)) { shell.openExternal(url); return { action: 'deny' }; }
    return { action: url.startsWith('blob:') ? 'allow' : 'deny' }; // blob: = the print preview of the edited PDF
  });
  win.webContents.on('will-navigate', (e, url) => { if (!url.startsWith('app://xd3/')) e.preventDefault(); }); // the window only ever shows the bundled app
  win.webContents.on('will-prevent-unload', e => { // the page asks to stay open when there are unsaved edits
    const leave = dialog.showMessageBoxSync(win, { type: 'question', buttons: ['Close without saving', 'Keep editing'], defaultId: 1, cancelId: 1, title: 'XD3 PDF Editor', message: 'You have unsaved changes.', detail: 'Close anyway?' }) === 0;
    if (leave) e.preventDefault();
  });
  if (SELFTEST) { // build check: load the app, report whether it started, quit
    win.webContents.once('did-finish-load', async () => {
      const r = await win.webContents.executeJavaScript(file ? `(async () => { // a file given on the command line: report how it opened
        for (let k = 0; k < 80 && !(XD3.S.pages.length || (XD3.docx && XD3.docx.D.zip)); k++) await new Promise(r => setTimeout(r, 250));
        await new Promise(r => setTimeout(r, 500));
        return JSON.stringify({ app: typeof XD3, pages: XD3.S.pages.length, reader: document.body.classList.contains('reader'), word: !!(XD3.docx && XD3.docx.isOpen()), title: document.title, desk: typeof xd3desk });
      })()` : `(async () => {
        if (window.__selftest) { const t = await window.__selftest; return JSON.stringify({ chrome: (navigator.userAgent.match(/Chrome.(\\d+)/) || [])[1], desk: typeof xd3desk, ...t }); } // the full suite from tests/, when bundled
        await XD3.ensureFont('Roboto');
        const d = await PDFLib.PDFDocument.create(); d.addPage([200, 200]);
        const pdf = await pdfjsLib.getDocument({ data: await d.save(), isEvalSupported: false }).promise;
        return JSON.stringify({ app: typeof XD3, desk: typeof xd3desk, pdfPages: pdf.numPages, fontkit: typeof fontkit, zip: typeof JSZip, font: document.fonts.check('12px "XD3 Roboto"'), crypto: !!(window.crypto && crypto.subtle), chrome: (navigator.userAgent.match(/Chrome.(\\d+)/) || [])[1] });
      })()`).catch(e => 'error: ' + e);
      console.log('XD3_SELFTEST ' + r); app.exit(/"fails":\[\]|"app":"object"/.test(r) ? 0 : 1);
    });
  }
  return win;
}
async function openPath(file) { // a file from a second launch: into a window that has nothing open, else a new window
  if (!file) { const w = BrowserWindow.getAllWindows()[0]; if (w) { if (w.isMinimized()) w.restore(); w.focus(); } else createWindow(''); return; }
  for (const w of BrowserWindow.getAllWindows()) {
    const idle = await w.webContents.executeJavaScript('!!(window.XD3 && XD3.idle && XD3.idle())').catch(() => false);
    if (idle) { w.loadURL(urlFor(file)); if (w.isMinimized()) w.restore(); w.focus(); return; }
  }
  createWindow(file).focus();
}

const ok = e => e.senderFrame && /^app:\/\/xd3\//.test(e.senderFrame.url); // only the bundled page may ask
function bridge() {
  ipcMain.handle('desk:info', e => ok(e) && { version: app.getVersion(), packaged: app.isPackaged, portable, canRegister, os: osInfo() });
  ipcMain.handle('desk:status', async e => ok(e) && process.platform === 'win32' ? { ...(await winreg.status()), canRegister } : null);
  ipcMain.handle('desk:make-default', async e => { if (!ok(e)) return null; if (!canRegister) return { error: 'This copy is not installed (development build)' }; return winreg.makeDefault(osInfo()); });
  ipcMain.handle('desk:settings', async e => {
    if (!ok(e)) return; const v = osInfo();
    if (v.major >= 10) return shell.openExternal('ms-settings:defaultapps?registeredAppUser=' + encodeURIComponent(winreg.NAME));
    require('child_process').execFile('control.exe', ['/name', 'Microsoft.DefaultPrograms', '/page', 'pageDefaultProgram']);
  });
  ipcMain.handle('desk:unregister', async e => { if (ok(e) && canRegister) await winreg.unregister(); return true; });
  ipcMain.handle('desk:print-pdf', async (e, o = {}) => { // the page as printed, with the paper size and margins the document asks for (inches)
    if (!ok(e)) return null; const m = o.margins || {};
    return e.sender.printToPDF({ printBackground: true, landscape: false, pageSize: o.width && o.height ? { width: o.width, height: o.height } : 'A4', margins: { top: m.top || 0, bottom: m.bottom || 0, left: m.left || 0, right: m.right || 0 } });
  });
}

if (!SELFTEST && !app.requestSingleInstanceLock()) app.quit();
else {
  app.on('second-instance', (e, argv, cwd) => openPath(fileIn(argv.slice(1), cwd)));
  app.whenReady().then(() => {
    protocol.registerFileProtocol('app', (request, callback) => {
      let p = decodeURIComponent(new URL(request.url).pathname);
      if (p.startsWith('/__open__/')) { const f = opened.get(p.split('/')[2]); return callback(f ? { path: f } : { error: -6 }); }
      if (p === '/' || p === '') p = '/index.html';
      const file = path.normalize(path.join(ROOT, p));
      callback(file.startsWith(ROOT) ? { path: file } : { error: -10 });
    });
    session.defaultSession.setPermissionRequestHandler((wc, permission, cb) => cb(['media', 'clipboard-read', 'clipboard-sanitized-write'].includes(permission))); // camera for Scan, clipboard for the right-click menu
    Menu.setApplicationMenu(null);
    bridge();
    createWindow(fileIn(process.argv.slice(1)));
    if (canRegister && !SELFTEST) setTimeout(async () => { // keep "Open with" pointing at this copy: always for the installed app; a portable copy only takes over when the old one is gone
      try { const st = await winreg.status(); if (st.registered) return; if (!portable || (st.registeredExe && !fs.existsSync(st.registeredExe))) await winreg.register(); } catch (err) { /* not essential */ }
    }, 4000);
  });
  app.on('window-all-closed', () => app.quit());
}
