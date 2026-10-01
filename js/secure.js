/* XD3 PDF Editor — password protection. Encrypts the saved PDF with the standard security handler (AES-128, PDF 1.6),
   which every mainstream PDF reader can open. Runs locally; the password never leaves the device. */
(() => {
'use strict';
const X = window.XD3, { S, toast, busy, dialog, download } = X;

function md5(bytes) { // RFC 1321 → Uint8Array(16)
  const R = [7, 12, 17, 22, 7, 12, 17, 22, 7, 12, 17, 22, 7, 12, 17, 22, 5, 9, 14, 20, 5, 9, 14, 20, 5, 9, 14, 20, 5, 9, 14, 20, 4, 11, 16, 23, 4, 11, 16, 23, 4, 11, 16, 23, 4, 11, 16, 23, 6, 10, 15, 21, 6, 10, 15, 21, 6, 10, 15, 21, 6, 10, 15, 21];
  const n = bytes.length, blocks = ((n + 8) >> 6) + 1, M = new Uint8Array(blocks * 64), dv = new DataView(M.buffer);
  M.set(bytes); M[n] = 0x80; dv.setUint32(blocks * 64 - 8, (n * 8) >>> 0, true); dv.setUint32(blocks * 64 - 4, Math.floor(n / 0x20000000), true);
  let a0 = 0x67452301, b0 = 0xefcdab89 | 0, c0 = 0x98badcfe | 0, d0 = 0x10325476;
  for (let o = 0; o < M.length; o += 64) {
    let A = a0, B = b0, C = c0, D = d0;
    for (let i = 0; i < 64; i++) {
      let F, g;
      if (i < 16) { F = (B & C) | (~B & D); g = i; } else if (i < 32) { F = (D & B) | (~D & C); g = (5 * i + 1) % 16; } else if (i < 48) { F = B ^ C ^ D; g = (3 * i + 5) % 16; } else { F = C ^ (B | ~D); g = (7 * i) % 16; }
      F = (F + A + Math.floor(Math.abs(Math.sin(i + 1)) * 4294967296) + dv.getUint32(o + g * 4, true)) | 0;
      A = D; D = C; C = B; B = (B + ((F << R[i]) | (F >>> (32 - R[i])))) | 0;
    }
    a0 = (a0 + A) | 0; b0 = (b0 + B) | 0; c0 = (c0 + C) | 0; d0 = (d0 + D) | 0;
  }
  const out = new Uint8Array(16), ov = new DataView(out.buffer); [a0, b0, c0, d0].forEach((v, i) => ov.setInt32(i * 4, v, true)); return out;
}
function rc4(key, data) {
  const s = new Uint8Array(256), out = new Uint8Array(data.length); for (let i = 0; i < 256; i++) s[i] = i;
  for (let i = 0, j = 0; i < 256; i++) { j = (j + s[i] + key[i % key.length]) & 255; const t = s[i]; s[i] = s[j]; s[j] = t; }
  for (let k = 0, i = 0, j = 0; k < data.length; k++) { i = (i + 1) & 255; j = (j + s[i]) & 255; const t = s[i]; s[i] = s[j]; s[j] = t; out[k] = data[k] ^ s[(s[i] + s[j]) & 255]; }
  return out;
}
const concat = (...parts) => { const out = new Uint8Array(parts.reduce((a, p) => a + p.length, 0)); let o = 0; for (const p of parts) { out.set(p, o); o += p.length; } return out; };
const PAD = new Uint8Array([0x28, 0xbf, 0x4e, 0x5e, 0x4e, 0x75, 0x8a, 0x41, 0x64, 0x00, 0x4e, 0x56, 0xff, 0xfa, 0x01, 0x08, 0x2e, 0x2e, 0x00, 0xb6, 0xd0, 0x68, 0x3e, 0x80, 0x2f, 0x0c, 0xa9, 0xfe, 0x64, 0x53, 0x69, 0x7a]);
const padPw = pw => concat(Uint8Array.from([...pw].slice(0, 32), c => c.charCodeAt(0)), PAD).slice(0, 32);

async function encryptPdf(bytes, userPw, ownerPw, P) { // ISO 32000-1 §7.6 — V4 / R4, crypt filter AESV2
  const L = PDFLib, doc = await L.PDFDocument.load(bytes, { updateMetadata: false }), ctx = doc.context;
  const id = crypto.getRandomValues(new Uint8Array(16)), up = padPw(userPw), xor = (k, i) => k.map(b => b ^ i);
  let h = md5(padPw(ownerPw)); for (let i = 0; i < 50; i++) h = md5(h);                              // Algorithm 3: the O entry
  let O = rc4(h, up); for (let i = 1; i <= 19; i++) O = rc4(xor(h, i), O);
  const pb = new Uint8Array(4); new DataView(pb.buffer).setInt32(0, P, true);
  let key = md5(concat(up, O, pb, id)); for (let i = 0; i < 50; i++) key = md5(key);                  // Algorithm 2: the file key
  let U = rc4(key, md5(concat(PAD, id))); for (let i = 1; i <= 19; i++) U = rc4(xor(key, i), U);       // Algorithm 5: the U entry
  U = concat(U, new Uint8Array(16));

  const hexS = b => L.PDFHexString.of(Array.from(b, v => v.toString(16).padStart(2, '0')).join(''));
  const aes = async (data, num, gen) => { // Algorithm 1: every object gets its own key and a random IV
    const ok = md5(concat(key, Uint8Array.of(num & 255, (num >> 8) & 255, (num >> 16) & 255, gen & 255, (gen >> 8) & 255, 0x73, 0x41, 0x6c, 0x54)));
    const iv = crypto.getRandomValues(new Uint8Array(16)), ck = await crypto.subtle.importKey('raw', ok, 'AES-CBC', false, ['encrypt']);
    return concat(iv, new Uint8Array(await crypto.subtle.encrypt({ name: 'AES-CBC', iv }, ck, data)));
  };
  const isStr = v => v instanceof L.PDFString || v instanceof L.PDFHexString;
  const walk = async (v, num, gen) => { // strings held directly inside an object are encrypted with that object's key
    if (v instanceof L.PDFDict) { for (const [k, x] of v.entries()) { if (isStr(x)) v.set(k, hexS(await aes(x.asBytes(), num, gen))); else await walk(x, num, gen); } }
    else if (v instanceof L.PDFArray) { for (let i = 0; i < v.size(); i++) { const x = v.get(i); if (isStr(x)) v.set(i, hexS(await aes(x.asBytes(), num, gen))); else await walk(x, num, gen); } }
  };
  for (const [ref, obj] of ctx.enumerateIndirectObjects()) {
    const num = ref.objectNumber, gen = ref.generationNumber;
    if (isStr(obj)) ctx.assign(ref, hexS(await aes(obj.asBytes(), num, gen)));
    else if (obj instanceof L.PDFStream) { const raw = obj instanceof L.PDFRawStream ? obj.contents : obj.getContents(); await walk(obj.dict, num, gen); ctx.assign(ref, L.PDFRawStream.of(obj.dict, await aes(raw, num, gen))); }
    else await walk(obj, num, gen);
  }
  const enc = ctx.obj({ Filter: 'Standard', V: 4, R: 4, Length: 128, P, CF: { StdCF: { AuthEvent: 'DocOpen', CFM: 'AESV2', Length: 16 } }, StmF: 'StdCF', StrF: 'StdCF' });
  enc.set(L.PDFName.of('O'), hexS(O)); enc.set(L.PDFName.of('U'), hexS(U)); // the encryption dictionary itself stays readable
  ctx.trailerInfo.Encrypt = ctx.register(enc); ctx.trailerInfo.ID = ctx.obj([hexS(id), hexS(id)]);
  return doc.save({ useObjectStreams: false, updateFieldAppearances: false }); // nothing may be added after encryption
}

async function protect() {
  if (!(window.crypto && crypto.subtle)) return toast('Password protection needs a secure page — start the editor with “Start PDF Editor.bat” or the desktop app', true);
  const r = await dialog('Password protect', `
    <label>Password needed to open the PDF<input type="password" name="pw" autocomplete="new-password"></label>
    <label>Repeat the password<input type="password" name="pw2" autocomplete="new-password"></label>
    <label class="inl"><input type="checkbox" id="pwShow">Show passwords</label>
    <div class="mhead" style="padding:6px 0 0">People who open it may</div>
    <label class="inl"><input type="checkbox" name="print" checked>Print</label>
    <label class="inl"><input type="checkbox" name="copy" checked>Copy text and pictures</label>
    <label class="inl"><input type="checkbox" name="edit" checked>Edit, annotate and fill forms</label>
    <p class="mut">Saves an AES-encrypted copy that opens in any PDF reader. Leave the password empty to only set the restrictions. A forgotten password cannot be recovered — keep it somewhere safe.</p>`, {
    ok: 'Save protected copy', onOpen: d => { d.querySelector('#pwShow').onchange = e => d.querySelectorAll('[name=pw],[name=pw2]').forEach(i => i.type = e.target.checked ? 'text' : 'password'); },
  });
  if (!r) return;
  if (r.pw !== r.pw2) return toast('The two passwords do not match', true);
  if (/[^\x20-\xff]/.test(r.pw)) return toast('Please use only English letters, digits and symbols in the password', true);
  if (!r.pw && r.print && r.copy && r.edit) return toast('Enter a password, or untick something people may not do', true);
  let P = 0xfffff0c0; if (r.print) P |= 4 | 2048; if (r.edit) P |= 8 | 32 | 256 | 1024; if (r.copy) P |= 16 | 512;
  const owner = Array.from(crypto.getRandomValues(new Uint8Array(24)), b => String.fromCharCode(33 + b % 90)).join(''); // a separate owner key keeps the restrictions in force
  await busy(async () => {
    const out = await encryptPdf(await X.buildPdf(S.pages), r.pw, owner, P | 0);
    download(new Blob([out], { type: 'application/pdf' }), X.baseName() + '-protected.pdf');
    toast(r.pw ? 'Protected copy saved — it now needs the password to open' : 'Restricted copy saved');
  }, 'Encrypting…');
}

X.ACT.protect = protect;
Object.assign(X, { encryptPdf, md5 });
})();
