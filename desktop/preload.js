// XD3 PDF Editor — desktop bridge. The page gets only these few Windows features; no file system or Node.js access.
const { contextBridge, ipcRenderer } = require('electron');
contextBridge.exposeInMainWorld('xd3desk', {
  info: () => ipcRenderer.invoke('desk:info'), // { version, packaged, portable, os: { name, major, minor, build } }
  pdfDefault: () => ipcRenderer.invoke('desk:status'), // which program opens PDFs: { mine, app, registered }
  makeDefault: () => ipcRenderer.invoke('desk:make-default'), // registers XD3, then Windows 7 sets it; Windows 8–11 show their own chooser
  openSettings: () => ipcRenderer.invoke('desk:settings'), // Windows' Default apps page for XD3
  unregister: () => ipcRenderer.invoke('desk:unregister'), // removes XD3 from "Open with" (portable copies)
  printToPdf: opts => ipcRenderer.invoke('desk:print-pdf', opts), // the page as it prints (Word editor → exact PDF)
});
