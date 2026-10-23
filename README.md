# XD3 PDF Editor

A free, fast, private PDF editor by **XD3Labs / XDCybertech Pvt Ltd**. Everything runs on your own device — in the browser, as a hosted
web app, or as a Windows program — and nothing is ever uploaded.

## What it does

- **Edit text in place, invisibly** — click any line or paragraph and retype it. On save the old words are removed from the page's own
  instructions and the new ones are written with the document's own font resource, in the same place in the drawing order (same clip, same
  tag), keeping the original's word spacing, kerning and ligatures. A line you open but do not change is left byte-for-byte untouched. Only
  letters the document's (subset) font lacks are borrowed from a metric-compatible look-alike (Carlito for Calibri, Gelasio for Georgia, …).
  The saved file keeps its name, producer, dates, forms, bookmarks, links and tags — nothing marks it as edited (Properties ▸ untick to brand it).
- **Edit scans and pictures** — text in scanned pages, flattened PDFs and images is recognised offline (Tesseract OCR, 14 languages) and
  becomes editable; crooked scans are straightened first; the closest typeface, size and colour are matched automatically.
- **Edit objects** — pick up pictures (at full resolution, with transparency) or any region, then move, resize, rotate or delete it.
- **Forms** — fill in text fields, checkboxes, radio buttons and dropdowns; values go into the real fields, which stay fillable.
- **Annotate** — pen, highlighter, shapes, arrows, sticky notes; saved as real PDF comments other readers can edit (or flattened, your choice).
- **Blackout / redaction** — the text under a blackout box is removed from the file (the page stays real text); a page is turned into a
  picture only when a blackout covers a picture. *Find ▸ Black out all* redacts every match. Deleted pages leave nothing behind in the file.
- **Make searchable** — OCR a scanned document and keep the words as an invisible layer that any reader can search and copy.
- **Pages** — merge PDF / Word / images, reorder, rotate, replace, extract, resize, page numbers, headers and footers, bookmarks.
- **Convert** — PDF ⇄ Word, PDF → images, images → PDF, text export; compress; split; compare two versions; batch processing.
- **Protect** — AES password protection, permanent redaction (flatten).
- **Insert** — images, signatures (drawn or typed), stamps, dates, watermarks; scan with the camera.
- **Links and crop** — add links to web pages or other pages; crop pages (one, all, or all of the same size).
- **Writing tools** — spelling and grammar checking (Harper, fully offline; US, UK, Canadian, Australian and Indian English), clarity
  hints (long sentences, passive voice, weak words, clichés), *Humanize* rewrites in a natural, casual or formal tone, readability statistics
  and text clean-up — for Word files (formatting kept: each fix goes into its own run), text files, pasted text and the open PDF (fixes are
  written back in the PDF's own font). Your dictionary is kept on the device.
- **View** — *View PDF* opens a file straight into reading mode; two pages side by side.
- **Read** — a reading mode with selectable text, clickable links, bookmarks in the side panel, night colours and full screen; select text to highlight, comment or copy it. Recently opened files are one click away on the start screen.

## Running it

| How | Steps |
| --- | --- |
| On your PC (browser) | Double-click `Start PDF Editor.bat` (needs Python or Node.js) and the editor opens at `http://localhost:5391` (served by `serve.py`, which never lets the browser keep a stale copy). |
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
js/native.js          content-stream engine: finds, removes and rewrites the page's own text (invisible edits, true redaction)
js/ocr.js             offline OCR, deskew, typeface matching
js/forms.js           form filling          js/extras.js   searchable PDF, bookmarks, compare, batch
js/insert.js          stamps, watermark, page numbers, signature
js/imaging.js         image filters, background remover, scanner
js/convert.js         Word ⇄ PDF, images, compress, flatten        js/secure.js   password protection
js/reader.js          reading mode (text layer, links, bookmarks panel, night, selection pop-up, two-page view)
js/writing.js         writing tools: spelling, grammar, humanize, readability, text tools (Word / text / PDF)
js/organize.js        page organizer        js/ui.js       theme, notifications, command palette, service worker
lib/                  pdf.js, pdf-lib, fontkit, JSZip, fonts, OCR engine + language data, Harper grammar checker (all bundled, no CDN)
desktop/              Electron shell for the Windows build
sw.js, manifest.webmanifest, icons/   installable offline web app
.htaccess             hosting settings (caching, MIME types, security headers)
```

## Licences

The editor is © XDCybertech Pvt Ltd. Bundled components: pdf.js (Apache 2.0), pdf-lib and fontkit (MIT), JSZip (MIT), Tesseract.js and
Tesseract language data (Apache 2.0), Harper grammar checker (Apache 2.0, lib/harper/LICENSE), fonts from Google Fonts (SIL OFL, Apache 2.0, Ubuntu Font Licence) and DejaVu (Bitstream Vera
licence) — see `lib/fonts/README.txt` and `lib/ocr/README.txt`.
