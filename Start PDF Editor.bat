@echo off
title XD3 PDF Editor - XD3Labs / XDCybertech Pvt Ltd
cd /d "%~dp0"
set PORT=5391

where python >nul 2>nul
if %errorlevel%==0 (
  echo Starting XD3 PDF Editor at http://localhost:%PORT%
  echo Keep this window open while using the editor. Close it to stop.
  start "" "http://localhost:%PORT%"
  python -m http.server %PORT% --bind 127.0.0.1
  goto :eof
)

where npx >nul 2>nul
if %errorlevel%==0 (
  echo Starting XD3 PDF Editor at http://localhost:%PORT%
  echo Keep this window open while using the editor. Close it to stop.
  start "" "http://localhost:%PORT%"
  npx --yes http-server -p %PORT% -a 127.0.0.1 -c-1
  goto :eof
)

echo Python or Node.js not found - opening the editor directly in your browser.
start "" "%~dp0index.html"
