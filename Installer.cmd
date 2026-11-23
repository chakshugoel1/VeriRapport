@echo off
setlocal enabledelayedexpansion

rem ============================================================================
rem  VeriRapport - one-file bootstrap installer
rem ============================================================================
rem  Double-click this file on a brand-new machine. It has no dependency on
rem  the rest of this repo being present - it downloads bootstrap-install.ps1
rem  from GitHub and runs it, which then clones the repo and runs the full
rem  installer (install.ps1). This file is meant to be copied/emailed/shared
rem  on its own; everything else it needs, it fetches itself.
rem
rem  Requires: PowerShell (built into Windows) and internet access to GitHub.
rem  If your network blocks raw.githubusercontent.com, ask for
rem  bootstrap-install.ps1 directly instead and run it with:
rem    powershell -NoProfile -ExecutionPolicy Bypass -File bootstrap-install.ps1
rem ============================================================================

set "SCRIPT_URL=https://raw.githubusercontent.com/chakshugoel1/VeriRapport/main/bootstrap-install.ps1"
set "TEMP_SCRIPT=%TEMP%\verirapport-bootstrap-install.ps1"

echo ===============================================================
echo   VeriRapport - Bootstrap Installer
echo ===============================================================
echo.
echo This will:
echo   - clone/update the repo to %%USERPROFILE%%\verirapport
echo   - install Node.js and Git if missing (via winget)
echo   - install npm dependencies and the Playwright browser
echo   - ask for your test-account credentials and save a login session
echo   - run a quick end-to-end check
echo.
echo You can close this window at any time; nothing runs until you continue.
echo.
pause

echo.
echo Downloading installer...
rem Windows PowerShell 5.1 sometimes defaults to an older TLS version than
rem GitHub requires, which closes the connection before any HTTP response
rem ("The underlying connection was closed: An unexpected error occurred on
rem a send."). Force TLS 1.2 for this process before downloading.
powershell -NoProfile -ExecutionPolicy Bypass -Command "[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; Invoke-WebRequest -UseBasicParsing -Uri '%SCRIPT_URL%' -OutFile '%TEMP_SCRIPT%'"
if errorlevel 1 goto DOWNLOAD_FAILED
if not exist "%TEMP_SCRIPT%" goto DOWNLOAD_FAILED

echo Download OK. Running installer...
echo.
powershell -NoProfile -ExecutionPolicy Bypass -File "%TEMP_SCRIPT%"
set "EXITCODE=%errorlevel%"

echo.
if "%EXITCODE%"=="0" (
  echo ===============================================================
  echo   Setup finished successfully.
  echo ===============================================================
) else (
  echo ===============================================================
  echo   Setup did not finish ^(exit code %EXITCODE%^).
  echo   Scroll up to see what failed. Fixing that and double-clicking
  echo   this file again will pick up from where it left off - steps
  echo   already completed are skipped automatically.
  echo ===============================================================
)
pause
exit /b %EXITCODE%

:DOWNLOAD_FAILED
echo.
echo ===============================================================
echo   Could not download the installer.
echo ===============================================================
echo Check your internet connection, or whether this network blocks
echo raw.githubusercontent.com (some corporate proxies do).
echo.
echo If it's blocked, ask whoever sent you this file for
echo bootstrap-install.ps1 directly, put it next to this file, and run:
echo   powershell -NoProfile -ExecutionPolicy Bypass -File bootstrap-install.ps1
echo.
pause
exit /b 1
