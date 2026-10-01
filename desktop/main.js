// XD3 PDF Editor — desktop shell (Electron). © XD3Labs / XDCybertech Pvt Ltd.
// Serves the bundled web app from a private app:// origin, so it runs fully offline with no local web server.
const { app, BrowserWindow, protocol, Menu, shell, session, dialog } = require('electron');
const path = require('path');

const ROOT = path.join(__dirname, 'www');
const OPENABLE = /\.(pdf|docx|png|jpe?g|gif|bmp|webp)$/i;
const fileArg = (process.argv.slice(1).find(a => OPENABLE.test(a) && !a.startsWith('-')) || '');

protocol.registerSchemesAsPrivileged([{ scheme: 'app', privileges: { standard: true, secure: true, supportFetchAPI: true, stream: true } }]);

function createWindow() {
  const win = new BrowserWindow({
    width: 1360, height: 860, minWidth: 420, minHeight: 480, backgroundColor: '#0e1016', autoHideMenuBar: true,
    title: 'XD3 PDF Editor', icon: path.join(__dirname, 'icon.ico'),
    webPreferences: { contextIsolation: true, nodeIntegration: false, sandbox: true, plugins: true },
  });
  // a file passed on the command line ("Open with…") is handed to the page, which fetches it from app://xd3/__open__/
  win.loadURL('app://xd3/index.html' + (fileArg ? '?open=' + encodeURIComponent(path.basename(fileArg)) : ''));
  win.webContents.setWindowOpenHandler(({ url }) => {
    if (/^https?:/i.test(url)) { shell.openExternal(url); return { action: 'deny' }; }
    return { action: url.startsWith('blob:') ? 'allow' : 'deny' }; // blob: = the print preview of the edited PDF
  });
  win.webContents.on('will-navigate', (e, url) => { if (!url.startsWith('app://xd3/')) e.preventDefault(); }); // the window only ever shows the bundled app
  win.webContents.on('will-prevent-unload', e => { // the page asks to stay open when there are unsaved edits
    const leave = dialog.showMessageBoxSync(win, { type: 'question', buttons: ['Close without saving', 'Keep editing'], defaultId: 1, cancelId: 1, title: 'XD3 PDF Editor', message: 'You have unsaved changes.', detail: 'Close the editor anyway?' }) === 0;
    if (leave) e.preventDefault();
  });
  if (process.env.XD3_SELFTEST) { // build check: load the app, report whether it started, quit
    win.webContents.once('did-finish-load', async () => {
      const r = await win.webContents.executeJavaScript(`(async () => {
        await XD3.ensureFont('Roboto');
        const d = await PDFLib.PDFDocument.create(); d.addPage([200, 200]);
        const pdf = await pdfjsLib.getDocument({ data: await d.save(), isEvalSupported: false }).promise;
        return JSON.stringify({ app: typeof XD3, pdfPages: pdf.numPages, fontkit: typeof fontkit, zip: typeof JSZip, font: document.fonts.check('12px "XD3 Roboto"'), crypto: !!(window.crypto && crypto.subtle), chrome: navigator.userAgent.match(/Chrome.(\\d+)/)[1] });
      })()`).catch(e => 'error: ' + e);
      console.log('XD3_SELFTEST ' + r); app.exit(0);
    });
  }
}

app.whenReady().then(() => {
  protocol.registerFileProtocol('app', (request, callback) => {
    let p = decodeURIComponent(new URL(request.url).pathname);
    if (p.startsWith('/__open__/')) return callback(fileArg ? { path: path.resolve(fileArg) } : { error: -6 });
    if (p === '/' || p === '') p = '/index.html';
    const file = path.normalize(path.join(ROOT, p));
    callback(file.startsWith(ROOT) ? { path: file } : { error: -10 });
  });
  session.defaultSession.setPermissionRequestHandler((wc, permission, cb) => cb(['media', 'clipboard-read', 'clipboard-sanitized-write'].includes(permission))); // camera for Scan, clipboard for the right-click menu
  Menu.setApplicationMenu(null);
  createWindow();
});
app.on('window-all-closed', () => app.quit());
