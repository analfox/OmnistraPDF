@echo off
rem ---------------------------------------------------------------------------
rem  Commit everything, push it to GitHub and, if asked, release a version.
rem
rem  A release is a tag vX.Y.Z. GitHub Actions (.github/workflows/release.yml)
rem  then builds OmnistraPDF.exe and publishes it, and every copy of the
rem  program offers the update on its next start.
rem
rem  No token file: anyone can READ a public repository, but only you can
rem  push, and git signs in with the saved Windows login (Git Credential
rem  Manager). The very first push opens a GitHub sign-in window once.
rem
rem  Messages are in English on purpose: Russian text stored in a .bat shows
rem  up as mojibake under some console code pages.
rem ---------------------------------------------------------------------------
setlocal
cd /d "%~dp0"
chcp 65001 >nul <nul
set "REPO=https://github.com/analfox/OmnistraPDF.git"

echo [1/5] Checking git...
git --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: git is not installed.
    pause
    exit /b 1
)
if not exist ".git" git init -b main
git remote get-url origin >nul 2>&1
if errorlevel 1 (git remote add origin "%REPO%") else (git remote set-url origin "%REPO%")

rem The current version is the VERSION line in pdf_tools.py.
rem \x22 is a double quote: a real one inside this command would confuse cmd.
rem <nul on chcp and powershell: they would otherwise swallow input meant for set /p.
set "CUR="
for /f "usebackq delims=" %%v in (`powershell -NoProfile -Command "if ([IO.File]::ReadAllText((Join-Path $PWD 'pdf_tools.py')) -match '(?m)^VERSION = \x22([^\x22]+)\x22') { $Matches[1] }" ^<nul`) do set "CUR=%%v"
if not defined CUR (
    echo ERROR: no VERSION line in pdf_tools.py.
    pause
    exit /b 1
)

echo.
echo [2/5] Current version: %CUR%
echo To release a version, type its number, for example the next one after %CUR%.
echo Just press Enter to push the changes without a release.
set "NEW="
set /p "NEW=Version to release: "
if not defined NEW goto :stage

echo %NEW%| findstr /r /x "[0-9][0-9]*\.[0-9][0-9]*\.[0-9][0-9]*" >nul
if errorlevel 1 (
    echo ERROR: "%NEW%" is not a version number. Use three numbers, like 1.2.3
    pause
    exit /b 1
)
powershell -NoProfile -Command "exit [int]([version]$env:NEW -lt [version]$env:CUR)" <nul
if errorlevel 1 (
    echo ERROR: %NEW% is lower than the current version %CUR%.
    pause
    exit /b 1
)
git rev-parse -q --verify "refs/tags/v%NEW%" >nul
if not errorlevel 1 (
    echo ERROR: version %NEW% has already been released. Pick a higher number.
    pause
    exit /b 1
)
rem Same number as in the code: the first release of it, nothing to change.
if "%NEW%"=="%CUR%" goto :stage
echo Writing VERSION = %NEW% into pdf_tools.py...
powershell -NoProfile -Command "$p = Join-Path $PWD 'pdf_tools.py'; $q = [char]34; $t = [IO.File]::ReadAllText($p) -replace '(?m)^VERSION = \x22[^\x22]*\x22', ('VERSION = ' + $q + $env:NEW + $q); [IO.File]::WriteAllText($p, $t, (New-Object Text.UTF8Encoding $false))" <nul
if errorlevel 1 (
    echo ERROR: could not write the version into pdf_tools.py.
    pause
    exit /b 1
)

:stage
echo.
echo [3/5] Staging changes...
git add -A
git diff --cached --quiet
if not errorlevel 1 (
    echo Nothing to commit - the working tree is clean.
    goto :tag
)
echo.
echo What will be committed:
git status --short
echo.

rem The date has to be computed OUTSIDE any if-block: cmd expands a whole
rem block before running it, so a variable set inside would be empty.
for /f %%i in ('powershell -NoProfile -Command "Get-Date -Format yyyy-MM-dd_HH-mm" ^<nul') do set "DT=%%i"
set "AUTO=update %DT%"
if defined NEW set "AUTO=Version %NEW%"
set "MSG="
set /p "MSG=Commit message (Enter for: %AUTO%): "
if not defined MSG set "MSG=%AUTO%"

echo [4/5] Committing: %MSG%
git commit -q -m "%MSG%"
if errorlevel 1 (
    echo ERROR: the commit failed.
    pause
    exit /b 1
)

:tag
if not defined NEW goto :push
git tag -a "v%NEW%" -m "OmnistraPDF %NEW%"
if errorlevel 1 (
    echo ERROR: could not create the tag v%NEW%.
    pause
    exit /b 1
)

:push
echo [5/5] Pushing to GitHub...
set "GIT_TERMINAL_PROMPT=0"
set "REFS=main"
if defined NEW set "REFS=main v%NEW%"
rem --atomic: the branch and the release tag go together or not at all.
rem A normal push, never --force: that would delete commits made elsewhere.
git push --atomic origin %REFS% 2>"%TEMP%\omnistrapdf_push.log"
set "RC=%errorlevel%"
type "%TEMP%\omnistrapdf_push.log"
if "%RC%"=="0" goto :pushed
rem The tag stays only on success, otherwise the next run would say "already released".
if defined NEW git tag -d "v%NEW%" >nul
findstr /l /c:"[rejected]" "%TEMP%\omnistrapdf_push.log" >nul && goto :rejected

echo.
echo === PUSH FAILED - nothing was sent. ===
echo.
echo If the text above says "Authentication failed" or "403": sign in to GitHub
echo as analfox in the window git opens. If no window appears, remove the old
echo github.com entry in Windows Credential Manager and run this again.
echo Otherwise check the internet connection.
pause
exit /b 1

:pushed
echo.
echo === Done. Everything is on GitHub. ===
if not defined NEW goto :end
echo.
echo GitHub is building OmnistraPDF.exe %NEW%, it takes about 5 minutes:
echo     https://github.com/analfox/OmnistraPDF/actions
echo When it is ready, the programs offer the update on their next start.
:end
pause
exit /b 0

:rejected
echo.
echo === PUSH REJECTED - nothing was sent. ===
echo.
echo GitHub has commits that you do not have locally. This happens when the
echo project was pushed from another machine, or edited on the site.
echo.
echo Nothing is lost. Pull them in first:
echo.
echo     git pull --rebase origin main
echo.
echo then run this script again. Do NOT add --force: it would delete the
echo commits that are on GitHub but not here.
pause
exit /b 1
