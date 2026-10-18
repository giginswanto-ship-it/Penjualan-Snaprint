@echo off
title Snaprint Digital - Sistem Dashboard Penjualan & Laba Rugi
cd /d "%~dp0"

echo ============================================================
echo   SNAPRINT DIGITAL - DASHBOARD PENJUALAN & LABA RUGI        
echo   Multi-Tahun: 2024, 2025, ^& 2026 (s/d Agustus)             
echo ============================================================

set "UV_EXE=%USERPROFILE%\.local\bin\uv.exe"
set "UV_PY=%APPDATA%\uv\python\cpython-3.11.17-windows-x86_64-none\python.exe"

if exist "%UV_EXE%" (
    echo   [INFO] Menjalankan Server Interaktif Python via UV...
    echo   [INFO] Fitur Simpan Perubahan Langsung ke File Aktif!
    start "" "%UV_EXE%" run python "%~dp0app.py"
    exit
)

if exist "%UV_PY%" (
    echo   [INFO] Menjalankan Server Interaktif Python...
    echo   [INFO] Fitur Simpan Perubahan Langsung ke File Aktif!
    start "" "%UV_PY%" "%~dp0app.py"
    exit
)

where python >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    echo   [INFO] Menjalankan Server Interaktif Python...
    echo   [INFO] Fitur Simpan Perubahan Langsung ke File Aktif!
    start "" python "%~dp0app.py"
    exit
)

echo   [INFO] Membuka Dashboard Langsung di Browser Default...
start "" "%~dp0index.html"
exit
