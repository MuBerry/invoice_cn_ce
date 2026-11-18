@echo off
rem SPDX-License-Identifier: Apache-2.0
rem =====================================================================
rem  invoice_cn_ce batch launcher - drag invoice files / folders / ZIPs here
rem  Output goes to  <this folder>\out\
rem  (ASCII-only text so Chinese paths are never mangled by cmd.)
rem =====================================================================
setlocal
chcp 65001 >nul 2>&1

if "%~1"=="" (
    echo.
    echo   Drag PDF / XML / OFD files, folders or ZIP archives onto this .bat.
    echo   Or run:  python -m invoice_cn_ce.batch ^<path^> -o out
    echo.
    pushd "%~dp0"
    set "PYTHONPATH=%~dp0src"
    python -m invoice_cn_ce.batch --help
    popd
    echo.
    if not defined INVOICE_CN_CE_NO_PAUSE pause
    exit /b 2
)

rem --- prefer a project-local venv when present -----------------------
set "PY=python"
if exist "%~dp0.venv\Scripts\python.exe" set "PY=%~dp0.venv\Scripts\python.exe"

pushd "%~dp0"
rem  run straight from the source tree: put src on PYTHONPATH
set "PYTHONPATH=%~dp0src"
"%PY%" -m invoice_cn_ce.batch %* -o "%~dp0out"
set RC=%ERRORLEVEL%
popd

if not "%RC%"=="0" (
    echo.
    echo   [FAILED] exit code %RC% - see messages above.
)

if exist "%~dp0out" start "" "%~dp0out"

echo.
if not defined INVOICE_CN_CE_NO_PAUSE pause
exit /b %RC%
