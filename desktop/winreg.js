// XD3 PDF Editor — Windows file types, per user (no administrator rights needed).
// Lists the app in "Open with" and in Default apps for PDF and Word files, tells which program opens PDFs now, and opens
// Windows' own chooser. Windows 8, 10 and 11 only let the person pick the default app, so the app asks Windows to show its
// chooser instead of changing the setting itself; Windows 7 allows a program to set it, which is done only when asked.
'use strict';
const { execFile } = require('child_process'), fs = require('fs'), os = require('os'), path = require('path');

module.exports = function winreg({ exe, root = 'HKCU\\Software', tmp = os.tmpdir() }) {
  const NAME = 'XD3 PDF Editor', PDF = 'XD3PDFEditor.pdf', DOCX = 'XD3PDFEditor.docx', exeName = path.basename(exe);
  const CL = root + '\\Classes', CAPS = 'XD3Labs\\XD3 PDF Editor\\Capabilities', FE = root + '\\Microsoft\\Windows\\CurrentVersion\\Explorer\\FileExts\\.pdf';
  const command = `"${exe}" "%1"`, icon = `"${exe}",0`;
  const run = (cmd, args, opt = {}) => new Promise(res => execFile(cmd, args, { windowsHide: true, ...opt }, (err, out) => res(err ? null : String(out || ''))));
  const reg = args => run('reg.exe', args);

  async function get(key, name) { // one value (null = the key's default value); null when it is not there
    const out = await reg(['query', key, ...(name == null ? ['/ve'] : ['/v', name])]); if (out == null) return null;
    const m = out.match(/\s(REG_[A-Z_]+)[ \t]*(.*?)\s*$/m); if (!m) return null;
    const v = m[2]; return /^\(.*\)$/.test(v) && name == null ? null : v; // "(value not set)", in the language of Windows
  }
  async function put(entries) { // [[key, name or null for the default value, text]] → one .reg file, imported in one go
    const q = s => '"' + String(s).replace(/\\/g, '\\\\').replace(/"/g, '\\"') + '"';
    let txt = 'Windows Registry Editor Version 5.00\r\n', last = '';
    for (const [k, n, v] of entries) { if (k !== last) { txt += `\r\n[${k.replace(/^HKCU\\/, 'HKEY_CURRENT_USER\\')}]\r\n`; last = k; } txt += (n == null ? '@' : q(n)) + '=' + q(v) + '\r\n'; }
    const f = path.join(tmp, `xd3-${process.pid}-${Date.now()}.reg`); fs.writeFileSync(f, Buffer.from('﻿' + txt, 'utf16le'));
    try { return (await reg(['import', f])) !== null; } finally { fs.unlink(f, () => {}); }
  }
  function entries() {
    const e = [];
    for (const [id, desc, ext] of [[PDF, 'PDF Document', '.pdf'], [DOCX, 'Word Document', '.docx']]) {
      const k = `${CL}\\${id}`;
      e.push([k, null, desc], [k, 'FriendlyTypeName', desc], [k + '\\DefaultIcon', null, icon],
        [k + '\\Application', 'ApplicationName', NAME], [k + '\\Application', 'ApplicationCompany', 'XDCybertech Pvt Ltd'],
        [k + '\\shell', null, 'open'], [k + '\\shell\\open', 'FriendlyAppName', NAME], [k + '\\shell\\open\\command', null, command],
        [`${CL}\\${ext}\\OpenWithProgids`, id, '']);
    }
    const A = `${CL}\\Applications\\${exeName}`;
    e.push([A, 'FriendlyAppName', NAME], [A + '\\shell\\open\\command', null, command], [A + '\\SupportedTypes', '.pdf', ''], [A + '\\SupportedTypes', '.docx', ''],
      [`${root}\\${CAPS}`, 'ApplicationName', NAME], [`${root}\\${CAPS}`, 'ApplicationDescription', 'Free offline PDF viewer and editor, with a Word editor'], [`${root}\\${CAPS}`, 'ApplicationIcon', icon],
      [`${root}\\${CAPS}\\FileAssociations`, '.pdf', PDF], [`${root}\\${CAPS}\\FileAssociations`, '.docx', DOCX],
      [`${root}\\RegisteredApplications`, NAME, `Software\\${CAPS}`]);
    return e;
  }
  const register = () => put(entries());
  async function unregister() {
    const del = (k, v) => reg(['delete', k, ...(v ? ['/v', v] : []), '/f']);
    if ((await get(`${CL}\\.pdf`, null)) === PDF) await reg(['delete', `${CL}\\.pdf`, '/ve', '/f']); // (only the value set on Windows 7, never the key)
    await Promise.all([del(`${CL}\\${PDF}`), del(`${CL}\\${DOCX}`), del(`${CL}\\.pdf\\OpenWithProgids`, PDF), del(`${CL}\\.docx\\OpenWithProgids`, DOCX),
      del(`${CL}\\Applications\\${exeName}`), del(`${root}\\XD3Labs\\XD3 PDF Editor`), del(`${root}\\RegisteredApplications`, NAME)]);
  }
  async function nameOf(id) { // a readable name for the program behind a file type
    if (!id) return '';
    const known = [[/^MSEdge/i, 'Microsoft Edge'], [/^Chrome/i, 'Google Chrome'], [/^Acro/i, 'Adobe Acrobat'], [/^Firefox/i, 'Firefox'], [/^Foxit/i, 'Foxit PDF Reader'], [/^SumatraPDF/i, 'SumatraPDF'], [/^AppX/i, 'a Microsoft Store app']];
    for (const [rx, n] of known) if (rx.test(id)) return n;
    for (const [k, v] of [['\\Application', 'ApplicationName'], ['\\shell\\open', 'FriendlyAppName']]) { const n = await get(`HKCR\\${id}${k}`, v); if (n && !n.startsWith('@')) return n; }
    const cmd = await get(`HKCR\\${id}\\shell\\open\\command`, null), m = cmd && cmd.match(/([^\\"/]+)\.exe/i); if (m) return m[1];
    const d = await get(`HKCR\\${id}`, null); return d && !d.startsWith('@') ? d : id;
  }
  async function status() {
    const id = (await get(FE + '\\UserChoiceLatest', 'ProgId')) || (await get(FE + '\\UserChoice', 'ProgId')) || (await get(`${CL}\\.pdf`, null)) || (await get('HKLM\\Software\\Classes\\.pdf', null)) || '';
    const mine = id === PDF || id.toLowerCase() === ('applications\\' + exeName).toLowerCase();
    const cur = await get(`${CL}\\${PDF}\\shell\\open\\command`, null);
    return { progId: id, mine, app: mine ? NAME : await nameOf(id), registered: cur === command, registeredExe: cur ? (cur.match(/^"([^"]+)"/) || [])[1] || '' : '' };
  }
  function samplePdf() { // the one-page PDF shown by Windows' chooser — opening it in XD3 confirms the choice
    const lines = [['F1', 26, 72, 740, 'XD3 PDF Editor'], ['F2', 13, 72, 706, 'If you can read this in XD3 PDF Editor, PDFs on this PC now open here.'], ['F2', 13, 72, 686, 'You can close this file. Double-click any PDF to read it; click Edit to change it.'], ['F2', 10, 72, 640, 'XD3Labs - XDCybertech Pvt Ltd']];
    const stream = lines.map(([f, s, x, y, t]) => `BT /${f} ${s} Tf ${x} ${y} Td (${t.replace(/[\\()]/g, '\\$&')}) Tj ET`).join('\n');
    const objs = ['<< /Type /Catalog /Pages 2 0 R >>', '<< /Type /Pages /Kids [3 0 R] /Count 1 >>', '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] /Resources << /Font << /F1 4 0 R /F2 5 0 R >> >> /Contents 6 0 R >>',
      '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold >>', '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>', `<< /Length ${stream.length} >>\nstream\n${stream}\nendstream`];
    let out = '%PDF-1.4\n'; const at = [];
    objs.forEach((o, i) => { at.push(out.length); out += `${i + 1} 0 obj\n${o}\nendobj\n`; });
    const x = out.length; out += `xref\n0 ${objs.length + 1}\n0000000000 65535 f \n` + at.map(n => String(n).padStart(10, '0') + ' 00000 n \n').join('');
    return Buffer.from(out + `trailer\n<< /Size ${objs.length + 1} /Root 1 0 R >>\nstartxref\n${x}\n%%EOF\n`, 'latin1');
  }
  async function makeDefault(ver) {
    if (!(await register())) return { error: 'Windows did not accept the registration' };
    if (ver.major === 6 && ver.minor <= 1) { // Windows 7: a program may set the default itself
      await reg(['delete', FE + '\\UserChoice', '/f']); await put([[`${CL}\\.pdf`, null, PDF]]);
      if ((await status()).mine) return { done: true };
    }
    const sample = path.join(tmp, 'XD3-PDF-Editor.pdf'); fs.writeFileSync(sample, samplePdf());
    execFile('rundll32.exe', ['shell32.dll,OpenAs_RunDLL ' + sample], { windowsVerbatimArguments: true }, () => {}); // Windows' "Open with" chooser
    return { chooser: true };
  }
  return { register, unregister, status, makeDefault, samplePdf, NAME, PDF };
};
