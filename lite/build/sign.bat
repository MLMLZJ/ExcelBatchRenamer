@echo off
setlocal enabledelayedexpansion
chcp 65001 >nul
title Sign ExcelPdfRename.exe

rem Sign dist\ExcelPdfRename.exe
rem Usage: build\sign.bat

cd /d "%~dp0.."

echo ============================================
echo   Sign dist\ExcelPdfRename.exe
echo ============================================
echo.

set SIGNTOOL=
for %%P in (
  "%ProgramFiles(x86)%\Windows Kits\10\bin\10.0.22621.0\x64\signtool.exe"
  "%ProgramFiles(x86)%\Windows Kits\10\bin\10.0.19041.0\x64\signtool.exe"
  "%ProgramFiles(x86)%\Windows Kits\10\Tools\bin\i386\signtool.exe"
  "%ProgramFiles(x86)%\Windows Kits\10\Tools\bin\x64\signtool.exe"
  "%ProgramFiles(x86)%\Windows Kits\8.1\bin\x64\signtool.exe"
) do (
  if exist %%P if not defined SIGNTOOL set SIGNTOOL=%%~P
)
if not defined SIGNTOOL (
  for /f "delims=" %%F in ('where signtool.exe 2^>nul') do (
    if not defined SIGNTOOL set SIGNTOOL=%%F
  )
)
if not defined SIGNTOOL (
  echo [ERROR] signtool.exe not found.
  echo Install Windows SDK with "Windows SDK Signing Tools for Desktop Apps".
  pause
  exit /b 1
)
echo [1/4] signtool: %SIGNTOOL%
echo.

set EXE=dist\ExcelPdfRename.exe
if not exist "%EXE%" (
  echo [ERROR] %EXE% not found. Run build\build.bat first.
  pause
  exit /b 1
)
echo [2/4] Target: %CD%\%EXE%
for /f "tokens=1" %%H in ('certutil -hashfile "%EXE%" SHA256 ^| findstr /v ":"') do (
  if not defined BEFORE set BEFORE=%%H
)
echo     SHA256 before signing: !BEFORE!
echo.

echo [3/4] Certificate source
echo   1 = PFX file
echo   2 = installed cert, by thumbprint
echo   3 = auto-pick from store
set /p MODE=Choose [1/2/3]:

set SIGNARGS=
if "%MODE%"=="1" (
  set /p PFXPATH=PFX file path:
  set /p PFXPASS=PFX password:
  set SIGNARGS=/f "!PFXPATH!" /p "!PFXPASS!"
) else if "%MODE%"=="2" (
  powershell -NoProfile -Command "Get-ChildItem Cert:\CurrentUser\My -CodeSigningCert | ForEach-Object { Write-Host ($_.Thumbprint + '  ' + $_.Subject) }"
  set /p THUMB=Thumbprint:
  set SIGNARGS=/sha1 !THUMB!
) else if "%MODE%"=="3" (
  set SIGNARGS=/a
) else (
  echo [ERROR] invalid choice
  pause
  exit /b 1
)

echo.
echo [4/4] Signing...
"%SIGNTOOL%" sign %SIGNARGS% /fd sha256 /td sha256 /tr http://timestamp.digicert.com /v "%EXE%"
if errorlevel 1 (
  echo [ERROR] Signing failed.
  pause
  exit /b 1
)

echo.
echo ============================================
echo   Signed. Verifying...
echo ============================================
"%SIGNTOOL%" verify /pa /v "%EXE%" > "%TEMP%\sv.txt" 2>&1
findstr /i /c:"Issued to" /c:"Number of files successfully Verified" /c:"Error" "%TEMP%\sv.txt"
echo   Note: self-signed / internal-CA certs report "untrusted root" here. Normal.
del "%TEMP%\sv.txt" 2>nul
echo.
for /f "tokens=1" %%H in ('certutil -hashfile "%EXE%" SHA256 ^| findstr /v ":"') do (
  if not defined AFTER set AFTER=%%H
)
echo     SHA256 after signing: !AFTER!
echo.
echo ^>^>^> Use THIS hash for VirusTotal / whitelisting ^<^<^<
pause
