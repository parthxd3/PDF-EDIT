@echo off
setlocal EnableExtensions
title Build XD3 PDF Editor (.exe)
cd /d "%~dp0"

rem Electron 22 is the last version that still runs on Windows 7 / 8 / 8.1 - the same build also runs on Windows 10 and 11.
set "ELECTRON_VERSION=22.3.27"
set "STAGE=%TEMP%\xd3-pdf-editor-build"
if not defined XD3_OUT set "XD3_OUT=%~dp0dist"
if not defined XD3_ARCH set "XD3_ARCH=x64 ia32"

echo.
echo  XD3 PDF Editor - EXE builder   (XD3Labs / XDCybertech Pvt Ltd)
echo  ---------------------------------------------------------------
echo  Builds the Windows app (Windows 7, 8, 10 and 11): portable folders, an installer and a portable ZIP.
echo  Needs Node.js and an internet connection the first time (downloads Electron, ~100 MB per build).
echo.

where node >nul 2>nul
if errorlevel 1 (
  echo  ERROR: Node.js was not found. Install the LTS version from https://nodejs.org and run this file again.
  goto :end
)

echo  [1/3] Preparing app files...
if exist "%STAGE%" rmdir /s /q "%STAGE%"
mkdir "%STAGE%\www"
copy /y "desktop\main.js" "%STAGE%\" >nul
copy /y "desktop\package.json" "%STAGE%\" >nul
copy /y "index.html" "%STAGE%\www\" >nul
xcopy /e /i /q /y "css" "%STAGE%\www\css" >nul
xcopy /e /i /q /y "js" "%STAGE%\www\js" >nul
xcopy /e /i /q /y "lib" "%STAGE%\www\lib" >nul
if exist "icons" xcopy /e /i /q /y "icons" "%STAGE%\www\icons" >nul
if defined XD3_WITH_TESTS xcopy /e /i /q /y "tests" "%STAGE%\www\tests" >nul
set "ICON="
if exist "desktop\icon.ico" copy /y "desktop\icon.ico" "%STAGE%\" >nul
if exist "desktop\icon.ico" set ICON=--icon="%STAGE%\icon.ico"

echo  [2/3] Building...
for %%A in (%XD3_ARCH%) do (
  echo        - %%A
  call npx --yes electron-packager@17.1.2 "%STAGE%" "XD3 PDF Editor" --platform=win32 --arch=%%A --electron-version=%ELECTRON_VERSION% --out="%XD3_OUT%" --overwrite --asar %ICON% --app-copyright="Copyright XDCybertech Pvt Ltd" --win32metadata.CompanyName="XDCybertech Pvt Ltd" --win32metadata.ProductName="XD3 PDF Editor" --win32metadata.FileDescription="XD3 PDF Editor"
  if errorlevel 1 goto :fail
)
rmdir /s /q "%STAGE%" >nul 2>nul

echo  [3/3] Making the downloads (installer + portable ZIP)...
if defined XD3_NO_DOWNLOADS goto :done
powershell -NoProfile -ExecutionPolicy Bypass -File "tools\make-downloads.ps1" -Out "%XD3_OUT%" -Arches "%XD3_ARCH%"
if errorlevel 1 echo  (The installer could not be made - the portable folders below still work.)

:done
echo.
echo  Done. Your apps are in:
echo    %XD3_OUT%\XD3 PDF Editor-win32-x64\XD3 PDF Editor.exe     (64-bit Windows 7 / 8 / 10 / 11)
echo    %XD3_OUT%\XD3 PDF Editor-win32-ia32\XD3 PDF Editor.exe    (32-bit Windows)
echo.
echo  Each folder is a complete portable app: copy or zip the WHOLE folder to share it. No installation needed.
echo.
echo  Files to give people (also offered by the website's "Download for Windows" button):
echo    %XD3_OUT%\downloads\XD3-PDF-Editor-Setup-*.exe       installer (shortcuts, Open with for PDFs)
echo    %XD3_OUT%\downloads\XD3-PDF-Editor-Portable-*.zip    unzip and run, no installation
echo  Run "Build Hostinger ZIP.bat" afterwards to put them on your website.
if not defined XD3_NOPAUSE start "" explorer "%XD3_OUT%"
goto :end

:fail
echo.
echo  BUILD FAILED - see the messages above. Check your internet connection and try again.

:end
echo.
if not defined XD3_NOPAUSE pause
endlocal
