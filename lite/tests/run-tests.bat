@echo off
setlocal enabledelayedexpansion
chcp 65001 >nul
title Run Lite tests

cd /d "%~dp0.."

set CSC=%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\csc.exe
if not exist "%CSC%" set CSC=%WINDIR%\Microsoft.NET\Framework\v4.0.30319\csc.exe
if not exist "%CSC%" (
    echo [ERROR] csc.exe not found.
    pause
    exit /b 1
)

if not exist "tests\bin" mkdir "tests\bin"

echo ============================================
echo   Building test runners
echo ============================================
set REFS=/r:System.dll /r:System.Core.dll /r:System.Xml.dll /r:System.Xml.Linq.dll /r:System.IO.Compression.dll /r:System.IO.Compression.FileSystem.dll

"%CSC%" /nologo /codepage:65001 /langversion:5 /target:exe /optimize+ /out:tests\bin\Tests.exe %REFS% src\Core.cs tests\TestRunner.cs
if errorlevel 1 goto :fail
"%CSC%" /nologo /codepage:65001 /langversion:5 /target:exe /optimize+ /out:tests\bin\EdgeTests.exe %REFS% src\Core.cs tests\EdgeTest.cs
if errorlevel 1 goto :fail
"%CSC%" /nologo /codepage:65001 /langversion:5 /target:exe /optimize+ /out:tests\bin\FixTests.exe %REFS% src\Core.cs tests\FixTest.cs
if errorlevel 1 goto :fail

echo.
echo ============================================
echo   Running tests
echo ============================================
echo [1/3] Main tests
tests\bin\Tests.exe
set R1=%errorlevel%
if "!R1!"=="2" set R1=0
echo [2/3] Edge case tests
tests\bin\EdgeTests.exe
set R2=%errorlevel%
if "!R2!"=="2" set R2=0
echo [3/3] Fix regression tests
tests\bin\FixTests.exe
set R3=%errorlevel%
if "!R3!"=="2" set R3=0

echo.
if "!R1!!R2!!R3!"=="000" (
    echo ============================================
    echo   ALL TESTS PASSED
    echo ============================================
    echo   [SKIP] means fixture data missing. Generate with:
    echo     python tests\make_testdata.py
    echo     python tests\make_edge.py
    pause
    exit /b 0
) else (
    echo ============================================
    echo   SOME TESTS FAILED  ^(main=!R1! edge=!R2! fix=!R3!^)
    echo ============================================
    pause
    exit /b 1
)

:fail
echo [ERROR] Build failed.
pause
exit /b 1
