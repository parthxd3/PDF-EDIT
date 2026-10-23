@echo off
setlocal EnableExtensions
title Build XD3 PDF Editor - Hostinger ZIP with the Windows downloads
cd /d "%~dp0"
rem The website ZIP with everything inside: the editor plus the Windows installers and portable ZIPs (dist\downloads),
rem so one upload puts the whole site online, "Download for Windows" included. Makes dist\xd3-pdf-editor-hostinger-full.zip.
if exist "dist\downloads\manifest.json" goto :pack
echo.
echo  The Windows app has not been built yet (dist\downloads is empty).
if defined XD3_NOPAUSE goto :pack
choice /c YN /m " Build it now with Build EXE.bat (takes a few minutes, needs Node.js)"
if errorlevel 2 goto :pack
set "XD3_NOPAUSE=1"
call "%~dp0Build EXE.bat"
set "XD3_NOPAUSE="
:pack
set "XD3_WITH_DOWNLOADS=1"
call "%~dp0Build Hostinger ZIP.bat"
endlocal
