@echo off
title Snaprint Digital - Sistem Dashboard Penjualan & Laba Rugi
cd /d "%~dp0"

echo ============================================================
echo   SNAPRINT DIGITAL - DASHBOARD PENJUALAN & LABA RUGI        
echo   Multi-Tahun: 2024, 2025, ^& 2026 (s/d Agustus)             
echo ============================================================

where python >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    echo   [INFO] Menjalankan Server Interaktif Python (Port 8080)...
    echo   [INFO] Fitur Simpan Perubahan Langsung ke File Aktif!
    start "" python "%~dp0app.py"
) else (
    echo   [INFO] Membuka Dashboard Langsung di Browser Default...
    start "" "%~dp0index.html"
)

exit
