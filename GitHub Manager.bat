@echo off
setlocal EnableExtensions EnableDelayedExpansion
title XD3 PDF Editor - GitHub Manager
cd /d "%~dp0"

rem ---------------------------------------------------------------------------
rem  GitHub Manager for XD3 PDF Editor  (XD3Labs / XDCybertech Pvt Ltd)
rem  Double-click = AUTOMATIC upload: add everything, commit, merge what is on GitHub, push. No questions asked.
rem  Other ways to run it:
rem     "GitHub Menu.bat"                                 the full menu (pull, branches, tags, undo, ...)
rem     "GitHub Manager.bat" watch                        keep uploading changes automatically every 2 minutes
rem     "GitHub Manager.bat" upload "commit message"      one upload with your own message
rem     "GitHub Manager.bat" status ^| commit "msg" ^| push ^| pull ^| fetch ^| log ^| get
rem ---------------------------------------------------------------------------
if not defined XD3_REMOTE set "XD3_REMOTE=https://github.com/parthxd3/PDF-EDIT.git"
set "REMOTE=%XD3_REMOTE%"
set "BRANCH=main"
set "MSGFILE=%TEMP%\xd3_commit_message.txt"

where git >nul 2>nul
if errorlevel 1 (
  echo.
  echo  Git is not installed. Download it from https://git-scm.com/download/win , install it, then run this file again.
  goto :finish
)

call :setup
if "%~1"=="" goto :auto
if /i "%~1"=="auto" goto :auto
if /i "%~1"=="menu" goto :menu
if /i "%~1"=="watch" goto :watch
goto :cli

rem ============================== AUTOMATIC =================================
:auto
set "AUTO=1"
echo.
echo   XD3 PDF Editor - automatic upload to GitHub
echo   Repository : !REMOTE!
echo.
call :upload ""
set "RC=!errorlevel!"
echo.
if "!RC!"=="0" (
  echo   Finished. This window closes by itself in 10 seconds. For more options run "GitHub Menu.bat".
  if not defined XD3_NOPAUSE timeout /t 10 >nul 2>nul
) else (
  echo   The upload did not finish - read the message above.
  if not defined XD3_NOPAUSE pause
)
endlocal & exit /b %RC%

:watch
set "AUTO=1"
echo.
echo   Auto-sync is ON for !REMOTE!
echo   Changes in this folder are uploaded every 2 minutes. Close this window to stop.
:watch_loop
set "CHANGED=0"
for /f %%n in ('git status --porcelain 2^>nul ^| find /c /v ""') do set "CHANGED=%%n"
if not "!CHANGED!"=="0" (
  echo.
  echo   !TIME:~0,5!  changed files: !CHANGED! - uploading
  call :upload ""
)
timeout /t 120 /nobreak >nul 2>nul
if errorlevel 1 ping -n 121 127.0.0.1 >nul
goto :watch_loop

rem ============================== MENU ======================================
:menu
call :refresh
cls
echo.
echo   XD3 PDF Editor - GitHub Manager
echo   ------------------------------------------------------------------
echo   Repository : !REMOTE!
echo   Branch     : !BRANCH!      Changed files: !CHANGED!      To upload: !AHEAD!   To download: !BEHIND!
echo   ------------------------------------------------------------------
echo.
echo    EVERYDAY
echo     1  Upload everything      add + commit + pull + push in one go
echo     2  Status                 what changed on this PC
echo     3  Commit                 save a snapshot locally, no upload
echo     4  Push                   upload commits to GitHub
echo     5  Pull                   download and merge the latest from GitHub
echo     6  Fetch                  check GitHub for news without changing files
echo.
echo    HISTORY
echo     7  History                recent commits
echo     8  Show changes           what is different since the last commit
echo     9  Undo last commit       keeps your files, removes the commit
echo.
echo    BRANCHES AND RELEASES
echo    10  Branches               list, create, switch, merge, delete
echo    11  Tag a release          for example v1.0.0
echo    12  Stash                  park unfinished changes and bring them back
echo.
echo    CAREFUL
echo    13  Get latest             REPLACE local files with the GitHub version
echo    14  Discard local changes  throw away everything not committed
echo    15  Force push             overwrite GitHub with this PC's history
echo.
echo    SETUP
echo    16  Settings               name, e-mail, repository address
echo    17  Open on GitHub         in your browser
echo    18  Clone a fresh copy     into another folder
echo    19  Auto-sync              upload changes by itself every 2 minutes
echo     0  Exit
echo.
set "CH="
set /p "CH=  Choose a number: "
if "!CH!"=="0" goto :finish_quiet
rem several empty answers in a row means there is nobody typing (input ended) - leave instead of looping forever
if not defined CH (
  set /a EMPTY+=1
  if !EMPTY! GEQ 5 goto :finish_quiet
  goto :menu
)
set "EMPTY=0"
echo.
if "!CH!"=="1" call :upload ""
if "!CH!"=="2" call :status
if "!CH!"=="3" call :commit_all ""
if "!CH!"=="4" call :push
if "!CH!"=="5" call :pull
if "!CH!"=="6" call :fetch
if "!CH!"=="7" call :history
if "!CH!"=="8" call :changes
if "!CH!"=="9" call :undo_commit
if "!CH!"=="10" call :branches
if "!CH!"=="11" call :tag
if "!CH!"=="12" call :stash
if "!CH!"=="13" call :get_latest
if "!CH!"=="14" call :discard
if "!CH!"=="15" call :force_push
if "!CH!"=="16" call :settings
if "!CH!"=="17" call :open_web
if "!CH!"=="18" call :clone_copy
if "!CH!"=="19" goto :watch
echo.
pause
goto :menu

rem ============================== COMMAND LINE ==============================
:cli
set "CMD=%~1"
set "RC=0"
if /i "!CMD!"=="upload" ( call :upload "%~2" & set "RC=!errorlevel!" & goto :cli_done )
if /i "!CMD!"=="commit" ( call :commit_all "%~2" & set "RC=!errorlevel!" & goto :cli_done )
if /i "!CMD!"=="status" ( call :status & goto :cli_done )
if /i "!CMD!"=="push" ( call :push & set "RC=!errorlevel!" & goto :cli_done )
if /i "!CMD!"=="pull" ( call :pull & set "RC=!errorlevel!" & goto :cli_done )
if /i "!CMD!"=="fetch" ( call :fetch & set "RC=!errorlevel!" & goto :cli_done )
if /i "!CMD!"=="log" ( call :history & goto :cli_done )
if /i "!CMD!"=="get" ( call :get_latest & set "RC=!errorlevel!" & goto :cli_done )
echo Unknown command "!CMD!". Use: auto, menu, watch, upload, commit, status, push, pull, fetch, log, get
set "RC=2"
:cli_done
endlocal & exit /b %RC%

rem ============================== FIRST-TIME SETUP ==========================
:setup
if not exist ".git\" (
  git init -b main >nul 2>nul
  if errorlevel 1 (
    git init >nul
    git symbolic-ref HEAD refs/heads/main
  )
  echo  Created a new Git repository in this folder.
)
git remote get-url origin >nul 2>nul
if errorlevel 1 git remote add origin "!REMOTE!"
for /f "delims=" %%u in ('git remote get-url origin 2^>nul') do set "REMOTE=%%u"
if not exist ".gitignore" call :make_ignore
if not exist ".gitattributes" call :make_attributes
git config core.safecrlf false
git config user.name >nul 2>nul
if errorlevel 1 call :identity
git config user.email >nul 2>nul
if errorlevel 1 call :identity
exit /b 0

:make_ignore
rem dist\ holds the built EXE folders - far larger than GitHub's 100 MB per-file limit
(
echo # build output and local tool settings - not part of the source
echo dist/
echo node_modules/
echo .claude/
echo *.log
echo Thumbs.db
echo desktop.ini
)> ".gitignore"
echo  Created .gitignore  - the dist folder with the built EXE is not uploaded.
exit /b 0

:make_attributes
(
echo * text=auto
echo *.bat text eol=crlf
echo *.ttf binary
echo *.ico binary
echo *.bcmap binary
echo *.pfb binary
)> ".gitattributes"
exit /b 0

:identity
echo.
echo  Git needs your name and e-mail to label your commits.
set "GN="
set "GE="
set /p "GN=  Your name   : "
set /p "GE=  Your e-mail : "
if defined GN git config user.name "!GN!"
if defined GE git config user.email "!GE!"
exit /b 0

:refresh
for /f "delims=" %%b in ('git symbolic-ref --short -q HEAD 2^>nul') do set "BRANCH=%%b"
set "CHANGED=0"
for /f %%n in ('git status --porcelain 2^>nul ^| find /c /v ""') do set "CHANGED=%%n"
set "AHEAD=?"
set "BEHIND=?"
for /f "tokens=1,2" %%a in ('git rev-list --left-right --count "origin/!BRANCH!...HEAD" 2^>nul') do (
  set "BEHIND=%%a"
  set "AHEAD=%%b"
)
exit /b 0

:confirm
rem  %1 = warning text. Returns 0 only when the user types YES.
echo  %~1
set "ANS="
set /p "ANS=  Type YES to continue: "
if /i "!ANS!"=="YES" exit /b 0
echo  Cancelled - nothing was changed.
exit /b 1

rem ============================== EVERYDAY ==================================
:upload
call :commit_all "%~1"
if errorlevel 1 exit /b 1
call :sync
exit /b !errorlevel!

:commit_all
call :refresh
git add -A
rem GitHub refuses any file over 100 MB - leave those out instead of failing the whole upload
set "BIG="
for /f "delims=" %%F in ('git diff --cached --name-only --diff-filter=AM') do (
  if exist "%%F" for %%S in ("%%F") do if %%~zS GTR 99000000 (
    echo  SKIPPED - larger than 100 MB, GitHub would reject it: %%F
    git reset -q -- "%%F"
    set "BIG=1"
  )
)
git diff --cached --quiet
if not errorlevel 1 (
  echo  Nothing new to commit - your files match the last commit.
  exit /b 0
)
set "MSG=%~1"
if not defined MSG if defined AUTO call :auto_message
if not defined MSG (
  echo  Files to be committed:
  git diff --cached --name-status
  echo.
  set /p "MSG=  Commit message, or just press Enter for an automatic one: "
)
if not defined MSG set "MSG=Update %DATE% %TIME:~0,5%"
> "!MSGFILE!" echo(!MSG!
git commit -q -F "!MSGFILE!"
set "RC=!errorlevel!"
del "!MSGFILE!" >nul 2>nul
if not "!RC!"=="0" (
  echo  The commit failed - see the message above.
  exit /b 1
)
echo  Committed: !MSG!
exit /b 0

:auto_message
rem  e.g.  Auto update 2026-10-02 14:05 - 12 files
set "STAMP="
for /f "usebackq delims=" %%t in (`powershell -NoProfile -Command "Get-Date -Format 'yyyy-MM-dd HH:mm'" 2^>nul`) do set "STAMP=%%t"
if not defined STAMP set "STAMP=%DATE% %TIME:~0,5%"
set "NFILES=0"
for /f %%n in ('git diff --cached --name-only ^| find /c /v ""') do set "NFILES=%%n"
set "MSG=Auto update !STAMP! - !NFILES! files"
exit /b 0

:sync
call :refresh
echo  Contacting GitHub...
git fetch origin
if errorlevel 1 (
  echo.
  echo  Could not reach the repository. Check your internet connection, that the repository exists,
  echo  and that you are signed in to GitHub - a sign-in window opens the first time.
  exit /b 1
)
git rev-parse --verify --quiet "refs/remotes/origin/!BRANCH!" >nul
if not errorlevel 1 (
  rem GitHub already has commits on this branch: merge them in before pushing
  set "XOPT="
  if defined AUTO set "XOPT=-X ours"
  git merge-base HEAD "origin/!BRANCH!" >nul 2>nul
  if errorlevel 1 (
    git pull origin "!BRANCH!" --no-rebase --no-edit --allow-unrelated-histories !XOPT!
  ) else (
    git pull origin "!BRANCH!" --no-rebase --no-edit !XOPT!
  )
  if errorlevel 1 if defined AUTO (
    git merge --abort >nul 2>nul
    echo.
    echo  GitHub's version could not be merged automatically, so nothing was uploaded and your files are untouched.
    echo  Run "GitHub Menu.bat": option 15 overwrites GitHub with this PC, option 13 replaces this PC with GitHub.
    exit /b 1
  )
  if errorlevel 1 (
    echo.
    echo  GitHub has changes that clash with yours - a merge conflict.
    echo  Open the files listed above, fix the parts between the conflict markers, then choose Upload again.
    echo  To give up the merge instead, run:  git merge --abort
    exit /b 1
  )
)
git push -u origin "!BRANCH!"
if errorlevel 1 (
  echo.
  echo  Push failed - see the message above.
  exit /b 1
)
echo.
echo  Done - everything is on GitHub.
exit /b 0

:status
call :refresh
echo  Branch !BRANCH!  -  changed files: !CHANGED!,  commits to upload: !AHEAD!,  commits to download: !BEHIND!
echo.
git status --short --branch
exit /b 0

:push
call :refresh
git push -u origin "!BRANCH!"
if errorlevel 1 (
  echo.
  echo  Push was rejected. If GitHub has newer commits, choose Pull first - or use Upload, which does both.
  exit /b 1
)
exit /b 0

:pull
call :refresh
git fetch origin
if errorlevel 1 exit /b 1
git rev-parse --verify --quiet "refs/remotes/origin/!BRANCH!" >nul
if errorlevel 1 (
  echo  GitHub has no "!BRANCH!" branch yet - nothing to pull.
  exit /b 0
)
git merge-base HEAD "origin/!BRANCH!" >nul 2>nul
if errorlevel 1 (
  git pull origin "!BRANCH!" --no-rebase --no-edit --allow-unrelated-histories
) else (
  git pull origin "!BRANCH!" --no-rebase --no-edit
)
exit /b !errorlevel!

:fetch
git fetch origin --prune --tags
if errorlevel 1 exit /b 1
call :refresh
echo  Commits on GitHub but not here: !BEHIND!     Commits here but not on GitHub: !AHEAD!
if not "!BEHIND!"=="0" if not "!BEHIND!"=="?" git log --oneline "HEAD..origin/!BRANCH!"
exit /b 0

rem ============================== HISTORY ===================================
:history
git --no-pager log --graph --decorate --date=short --pretty=format:"%%C(yellow)%%h%%Creset %%ad %%C(cyan)%%an%%Creset %%s%%C(auto)%%d" -25
echo.
exit /b 0

:changes
git --no-pager diff --stat HEAD
echo.
set "ANS="
set /p "ANS=  Show the full line-by-line changes? (y/N): "
if /i "!ANS!"=="y" git diff HEAD
exit /b 0

:undo_commit
git rev-parse --verify --quiet HEAD~1 >nul
if errorlevel 1 (
  echo  There is no earlier commit to go back to.
  exit /b 1
)
echo  Last commit:
git --no-pager log -1 --oneline
call :confirm "This removes that commit but keeps all your files exactly as they are."
if errorlevel 1 exit /b 1
git reset --soft HEAD~1
echo  Commit removed. Your changes are still here, ready to be committed again.
exit /b 0

rem ============================== BRANCHES / RELEASES =======================
:branches
echo  Branches  - the star marks the one you are on:
git --no-pager branch -a -vv
echo.
echo   1  Create a new branch and switch to it
echo   2  Switch to another branch
echo   3  Merge another branch into this one
echo   4  Delete a local branch
echo   5  Upload this branch to GitHub
echo   Enter  Back
set "BC="
set /p "BC=  Choose: "
if not defined BC exit /b 0
set "BN="
if "!BC!"=="5" (
  call :refresh
  git push -u origin "!BRANCH!"
  exit /b 0
)
set /p "BN=  Branch name: "
if not defined BN exit /b 0
if "!BC!"=="1" git switch -c "!BN!"
if "!BC!"=="2" git switch "!BN!"
if "!BC!"=="3" git merge --no-edit "!BN!"
if "!BC!"=="4" git branch -d "!BN!"
exit /b 0

:tag
echo  Existing tags:
git --no-pager tag --sort=-creatordate
echo.
set "TG="
set "TM="
set /p "TG=  New tag name, for example v1.0.0 : "
if not defined TG exit /b 0
set /p "TM=  Short description               : "
if not defined TM set "TM=Release !TG!"
git tag -a "!TG!" -m "!TM!"
if errorlevel 1 exit /b 1
git push origin "!TG!"
if errorlevel 1 exit /b 1
echo  Tag !TG! is on GitHub. Create a release from it under Releases on the repository page.
exit /b 0

:stash
echo  Parked changes:
git --no-pager stash list
echo.
echo   1  Park my current changes
echo   2  Bring back the most recent parked changes
echo   3  Delete the most recent parked changes
echo   Enter  Back
set "SC="
set /p "SC=  Choose: "
if "!SC!"=="1" git stash push -u -m "Parked %DATE% %TIME:~0,5%"
if "!SC!"=="2" git stash pop
if "!SC!"=="3" git stash drop
exit /b 0

rem ============================== CAREFUL ===================================
:get_latest
call :refresh
git fetch origin
if errorlevel 1 exit /b 1
git rev-parse --verify --quiet "refs/remotes/origin/!BRANCH!" >nul
if errorlevel 1 (
  echo  GitHub has no "!BRANCH!" branch yet - there is nothing to get.
  exit /b 1
)
call :confirm "This REPLACES your local files with the GitHub version. Uncommitted work and local-only commits are lost."
if errorlevel 1 exit /b 1
git reset --hard "origin/!BRANCH!"
echo  This folder now matches GitHub. Files Git does not track, such as dist, were left alone.
exit /b 0

:discard
git status --short
call :confirm "This throws away ALL changes since the last commit, including new files. It cannot be undone."
if errorlevel 1 exit /b 1
git reset --hard
git clean -fd
echo  Local changes discarded.
exit /b 0

:force_push
call :refresh
call :confirm "This OVERWRITES the history on GitHub with this PC's history. Commits that exist only on GitHub are lost."
if errorlevel 1 exit /b 1
git push --force-with-lease -u origin "!BRANCH!"
exit /b !errorlevel!

rem ============================== SETUP =====================================
:settings
echo  Name       :
git config user.name
echo  E-mail     :
git config user.email
echo  Repository : !REMOTE!
echo.
echo   1  Change name and e-mail for this project
echo   2  Change the repository address
echo   3  Sign out of GitHub on this PC  - you are asked to sign in on the next push
echo   Enter  Back
set "ST="
set /p "ST=  Choose: "
if "!ST!"=="1" call :identity
if "!ST!"=="2" (
  set "NU="
  set /p "NU=  New repository address: "
  if defined NU (
    git remote set-url origin "!NU!"
    set "REMOTE=!NU!"
    echo  Repository address changed.
  )
)
if "!ST!"=="3" call :signout
exit /b 0

:signout
> "%TEMP%\gh_signout.txt" echo protocol=https
>> "%TEMP%\gh_signout.txt" echo host=github.com
>> "%TEMP%\gh_signout.txt" echo.
git credential reject < "%TEMP%\gh_signout.txt"
del "%TEMP%\gh_signout.txt" >nul 2>nul
echo  Signed out of GitHub on this PC.
exit /b 0

:open_web
set "WEB=!REMOTE!"
if /i "!WEB:~-4!"==".git" set "WEB=!WEB:~0,-4!"
start "" "!WEB!"
exit /b 0

:clone_copy
set "CD2="
set /p "CD2=  Folder for the fresh copy, for example C:\Projects\PDF-EDIT : "
if not defined CD2 exit /b 0
git clone "!REMOTE!" "!CD2!"
exit /b !errorlevel!

rem ============================== EXIT ======================================
:finish
echo.
pause
:finish_quiet
endlocal
exit /b 0
