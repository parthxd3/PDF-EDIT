@echo off
title XD3 PDF Editor - self-test
cd /d "%~dp0"
set PORT=5392

where python >nul 2>nul
if not %errorlevel%==0 (
  echo Python is needed to run the tests (it serves the files and builds the samples). Install it from https://python.org
  pause
  goto :eof
)
if not exist "tests\samples\native.pdf" (
  echo Building the sample PDFs...
  python tests\make_samples.py || (echo. & echo Could not build the samples. Run:  pip install reportlab pillow & pause & goto :eof)
)
echo Opening the editor with the self-test at http://localhost:%PORT%/index.html?selftest=1
echo A dialog lists every check; details are also in the browser console (F12). Close this window when you are done.
start "" "http://localhost:%PORT%/index.html?selftest=1"
python -m http.server %PORT% --bind 127.0.0.1
