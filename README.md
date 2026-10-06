# XD3 PDF Editor

A free, fast, private PDF editor by **XD3Labs / XDCybertech Pvt Ltd**. Everything runs on your own device — in the browser, as a hosted
web app, or as a Windows program — and nothing is ever uploaded.

## What it does

- **Edit text in place** — click any line or paragraph and retype it. The document's own fonts are reused (embedded fonts are extracted and
  re-embedded on save), word spacing and justified lines are kept, rotated text is editable, and a bold word inside a line keeps its weight.
- **Edit scans and pictures** — text in scanned pages, flattened PDFs and images is recognised offline (Tesseract OCR, 14 languages) and
  becomes editable; crooked scans are straightened first; the closest typeface, size and colour are matched automatically.
- **Edit objects** — pick up pictures (at full resolution, with transparency) or any region, then move, resize, rotate or delete it.
- **Forms** — fill in text fields, checkboxes, radio buttons and dropdowns.
- **Annotate** — pen, highlighter, shapes, arrows, sticky notes; saved as real PDF comments other readers can edit (or flattened, your choice).
- **Blackout** — pages with blackout boxes are saved as pictures, so the hidden text really is gone.
- **Make searchable** — OCR a scanned document and keep the words as an invisible layer that any reader can search and copy.
- **Pages** — merge PDF / Word / images, reorder, rotate, replace, extract, resize, page numbers, headers and footers, bookmarks.
- **Convert** — PDF ⇄ Word, PDF → images, images → PDF, text export; compress; split; compare two versions; batch processing.
- **Protect** — AES password protection, permanent redaction (flatten).
- **Insert** — images, signatures (drawn or typed), stamps, dates, watermarks; scan with the camera.

## Running it

| How | Steps |
| --- | --- |
| On your PC (browser) | Double-click `Start PDF Editor.bat` (needs Python or Node.js) and the editor opens at `http://localhost:5391`. |
| Windows program | `Build EXE.bat` builds a portable app for Windows 7 / 8 / 10 / 11 into `dist\`. |
| Website (Hostinger or any Apache / LiteSpeed host) | `Build Hostinger ZIP.bat` makes `dist\xd3-pdf-editor-hostinger.zip`; upload and extract it into `public_html`. The site works offline after the first visit and can be installed like an app. |

## Testing

`Run Tests.bat` opens the editor with `?selftest=1`, which runs the checks in `tests/selftest.js` against the sample PDFs in
`tests/samples` (rebuild them with `python tests/make_samples.py`; needs `pip install reportlab pillow`). The desktop build runs the
same suite when started with `XD3_SELFTEST=1`.

## Layout

```
index.html            the app shell (strict Content-Security-Policy: only its own files)
css/style.css
js/app.js             viewer, tools, text editing, export (pdf-lib)
js/ocr.js             offline OCR, deskew, typeface matching
js/forms.js           form filling          js/extras.js   searchable PDF, bookmarks, compare, batch
js/insert.js          stamps, watermark, page numbers, signature
js/imaging.js         image filters, background remover, scanner
js/convert.js         Word ⇄ PDF, images, compress, flatten        js/secure.js   password protection
js/organize.js        page organizer        js/ui.js       theme, notifications, command palette, service worker
lib/                  pdf.js, pdf-lib, fontkit, JSZip, fonts, OCR engine + language data (all bundled, no CDN)
desktop/              Electron shell for the Windows build
sw.js, manifest.webmanifest, icons/   installable offline web app
.htaccess             hosting settings (caching, MIME types, security headers)
```

## Licences

The editor is © XDCybertech Pvt Ltd. Bundled components: pdf.js (Apache 2.0), pdf-lib and fontkit (MIT), JSZip (MIT), Tesseract.js and
Tesseract language data (Apache 2.0), fonts from Google Fonts (SIL OFL, Apache 2.0, Ubuntu Font Licence) and DejaVu (Bitstream Vera
licence) — see `lib/fonts/README.txt` and `lib/ocr/README.txt`.
