@echo off
chcp 65001 > nul
title POP SIPGN Autofill - SPPG Kalisapu

echo =======================================================
echo     POP SIPGN Autofill - SPPG Kalisapu (Langkah 2)
echo =======================================================
echo.

:: Periksa apakah Python sudah terpasang
where python >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python belum terdeteksi di komputer ini!
    echo Silakan install Python 3 terlebih dahulu dari https://www.python.org/
    echo dan pastikan centang opsi "Add Python to PATH" saat instalasi.
    pause
    exit /b 1
)

:: Periksa apakah library pandas dan openpyxl sudah terpasang
python -c "import pandas, openpyxl" >nul 2>&1
if errorlevel 1 (
    echo [INFO] Modul belum lengkap. Sedang menginstal pustaka yang dibutuhkan...
    pip install -r requirements.txt
    if errorlevel 1 (
        echo [ERROR] Gagal menginstal dependensi.
        pause
        exit /b 1
    )
    echo [SUKSES] Seluruh modul berhasil diinstal!
    echo.
)

:: Jalankan aplikasi utama
echo [INFO] Menjalankan aplikasi autofill...
echo.
python main.py

if errorlevel 1 (
    echo.
    echo [ERROR] Aplikasi berhenti karena terjadi kesalahan.
    pause
)
