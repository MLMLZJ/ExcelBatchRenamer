@echo off
setlocal
chcp 65001 >nul
title Build ExcelPdfRename.exe (Lite / .NET Framework 4.8)

rem Build the Lite variant. Run from the lite\ folder:
rem   build\build.bat
rem Output: dist\ExcelPdfRename.exe

cd /d "%~dp0.."

echo ============================================
echo   Build Lite variant (ExcelPdfRename.exe)
echo ============================================
echo.

set CSC=%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\csc.exe
if not exist "%CSC%" set CSC=%WINDIR%\Microsoft.NET\Framework\v4.0.30319\csc.exe
if not exist "%CSC%" (
    echo [ERROR] csc.exe not found. .NET Framework 4.x is required.
    echo         Windows 10 / 11 already ship it.
    pause
    exit /b 1
)

if not exist "dist" mkdir dist

echo [1/2] Compiler: %CSC%
echo [2/2] Compiling...
echo.

"%CSC%" /nologo /codepage:65001 /langversion:5 /target:winexe /platform:anycpu /optimize+ /warn:4 ^
  /win32icon:resources\app.ico ^
  /win32manifest:resources\app.manifest ^
  /resource:resources\app.ico,app.ico ^
  /out:dist\ExcelPdfRename.exe ^
  /r:System.dll /r:System.Core.dll /r:System.Xml.dll /r:System.Xml.Linq.dll ^
  /r:System.Windows.Forms.dll /r:System.Drawing.dll ^
  /r:System.IO.Compression.dll /r:System.IO.Compression.FileSystem.dll ^
  src\Core.cs src\Program.cs src\UiKit.cs src\MainForm.cs src\BatchForm.cs src\AssemblyInfo.cs

if errorlevel 1 (
    echo.
    echo [ERROR] Build failed.
    pause
    exit /b 1
)

echo.
echo ============================================
echo   Done: %CD%\dist\ExcelPdfRename.exe
echo ============================================
echo   Plain .NET Framework assembly, no dependencies.
echo   Next: sign it with build\sign.bat
echo ============================================
dir /b dist\ExcelPdfRename.exe
pause
