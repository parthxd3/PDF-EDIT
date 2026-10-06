Offline text recognition (OCR) bundled with XD3 PDF Editor. Nothing here is downloaded at run time.

tesseract.min.js, worker.min.js            Tesseract.js 6.0.1          Apache License 2.0  (LICENSE_TESSERACT_JS)
tesseract-core-simd-lstm.wasm.js           tesseract.js-core 6.1.2     Apache License 2.0  (WebAssembly build of Tesseract OCR)
tesseract-core-lstm.wasm.js                same, for processors without SIMD
*.traineddata.gz                           Tesseract language data (tessdata_best, integerised) - Apache License 2.0:
                                           eng English, hin Hindi, mar Marathi, guj Gujarati, ben Bengali, pan Punjabi, tam Tamil, tel Telugu,
                                           kan Kannada, mal Malayalam, ara Arabic, spa Spanish, fra French, deu German

To add another language: put its <code>.traineddata.gz (from the @tesseract.js-data/<code> npm package, folder 4.0.0_best_int)
in this folder and add it to the LANGS list at the top of js/ocr.js.
