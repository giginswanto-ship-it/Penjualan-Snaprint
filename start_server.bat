@echo off
title Snaprint Digital - Server API Simpan Perubahan Data
cd /d "%~dp0"

echo ============================================================
echo   SNAPRINT DIGITAL - SERVER API SIMPAN PERUBAHAN DATA       
echo   Mendukung Penyimpanan Otomatis ke File & Database SQL     
echo ============================================================

set "UV_EXE=%USERPROFILE%\.local\bin\uv.exe"
set "UV_PY=%APPDATA%\uv\python\cpython-3.11.17-windows-x86_64-none\python.exe"

if exist "%UV_EXE%" (
    "%UV_EXE%" run python "%~dp0app.py"
    goto end
)

if exist "%UV_PY%" (
    "%UV_PY%" "%~dp0app.py"
    goto end
)

python "%~dp0app.py"

:end
if %ERRORLEVEL% NEQ 0 (
    pause
)
