@echo off
setlocal EnableExtensions
title Build XD3 PDF Editor - Hostinger ZIP
cd /d "%~dp0"

set "STAGE=%TEMP%\xd3-pdf-editor-web"
if not defined XD3_OUT set "XD3_OUT=%~dp0dist"
set "ZIP=%XD3_OUT%\xd3-pdf-editor-hostinger.zip"

echo.
echo  XD3 PDF Editor - web hosting package   (XD3Labs / XDCybertech Pvt Ltd)
echo  ---------------------------------------------------------------------
echo  Packs the editor into one ZIP that can be uploaded to Hostinger (or any Apache / LiteSpeed host).
echo.

if not exist "index.html" (
  echo  ERROR: index.html was not found. Keep this file in the XD3 PDF Editor folder and run it from there.
  goto :end
)

echo  [1/2] Collecting the app files...
if exist "%STAGE%" rmdir /s /q "%STAGE%"
mkdir "%STAGE%"
copy /y "index.html" "%STAGE%\" >nul
if exist ".htaccess" copy /y ".htaccess" "%STAGE%\" >nul
xcopy /e /i /q /y "css" "%STAGE%\css" >nul
xcopy /e /i /q /y "js" "%STAGE%\js" >nul
xcopy /e /i /q /y "lib" "%STAGE%\lib" >nul
xcopy /e /i /q /y "icons" "%STAGE%\icons" >nul
copy /y "sw.js" "%STAGE%\" >nul
copy /y "manifest.webmanifest" "%STAGE%\" >nul
if not exist "%STAGE%\lib\pdf.min.js" goto :fail
rem (the editor's own scripts are joined into one file, so the page needs one request instead of twelve)
powershell -NoProfile -ExecutionPolicy Bypass -File "tools/bundle-js.ps1" "%STAGE%"
if errorlevel 1 goto :fail

echo  [2/2] Creating the ZIP...
if not exist "%XD3_OUT%" mkdir "%XD3_OUT%"
if exist "%ZIP%" del /q "%ZIP%"
rem (written entry by entry with forward slashes, so the folders unpack correctly on the Linux server)
powershell -NoProfile -ExecutionPolicy Bypass -Command "$ErrorActionPreference='Stop'; Add-Type -AssemblyName System.IO.Compression; Add-Type -AssemblyName System.IO.Compression.FileSystem; $s=$env:STAGE; $z=[IO.Compression.ZipFile]::Open($env:ZIP,'Create'); try { Get-ChildItem -LiteralPath $s -Recurse -File -Force | ForEach-Object { $n=$_.FullName.Substring($s.Length+1).Replace('\','/'); [void][IO.Compression.ZipFileExtensions]::CreateEntryFromFile($z,$_.FullName,$n,'Optimal') } } finally { $z.Dispose() }"
if errorlevel 1 goto :fail
if not exist "%ZIP%" goto :fail
rmdir /s /q "%STAGE%" >nul 2>nul

echo.
echo  Done:  %ZIP%
echo.
echo  How to put it online with Hostinger:
echo    1. hPanel ^> Websites ^> Manage ^> File Manager, and open the public_html folder
echo       (or a sub-folder such as public_html/pdf if the editor should live at yourdomain.com/pdf).
echo    2. Upload xd3-pdf-editor-hostinger.zip there, right-click it and choose Extract, extracting into that same folder.
echo    3. Delete the ZIP from the server and open your domain. index.html must sit directly in the folder you chose.
echo    4. Turn SSL on (hPanel ^> Security ^> SSL), then remove the # from the four HTTPS lines in .htaccess:
echo       the camera scanner only works over https.
echo.
echo  To update the site later: run this file again, upload the new ZIP and extract it over the old files.
if not defined XD3_NOPAUSE start "" explorer "%XD3_OUT%"
goto :end

:fail
echo.
echo  FAILED - the package could not be created. Check that the css, js and lib folders are next to this file
echo  and that the dist folder is not open in another program.

:end
echo.
if not defined XD3_NOPAUSE pause
endlocal
