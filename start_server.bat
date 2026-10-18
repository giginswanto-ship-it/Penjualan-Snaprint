@echo off
title Snaprint Digital - Server API Simpan Perubahan Data
cd /d "%~dp0"

echo ============================================================
echo   SNAPRINT DIGITAL - SERVER API SIMPAN PERUBAHAN DATA       
echo   Mendukung Penyimpanan Otomatis ke File & Database SQL     
echo ============================================================

python "%~dp0app.py"
if %ERRORLEVEL% NEQ 0 (
    pause
)
