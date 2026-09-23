@echo off
REM ==========================================================================
REM Script Jalan Otomatis - Aplikasi Penggabung Proposal PKKPRL (Windows)
REM ==========================================================================
REM Cara pakai:
REM   1. Double-click file ini, ATAU buka folder ini lewat Command Prompt
REM      lalu ketik:  jalankan_windows.bat
REM   2. Buka browser ke http://localhost:5000
REM
REM Script ini akan otomatis:
REM   - Membuat virtual environment (folder "venv") kalau belum ada
REM   - Mengaktifkan virtual environment tersebut
REM   - Menginstall semua dependency dari requirements.txt di dalam venv
REM   - Menjalankan aplikasinya (python app.py)
REM ==========================================================================

cd /d "%~dp0"

where python >nul 2>nul
if %errorlevel% neq 0 (
    echo ERROR: Python tidak ditemukan di PATH.
    echo Silakan install Python 3.10+ dari https://www.python.org/downloads/
    echo PENTING: saat instalasi, centang kotak "Add Python to PATH".
    pause
    exit /b 1
)

echo [1/4] Python ditemukan:
python --version

if not exist venv (
    echo [2/4] Membuat virtual environment baru di folder "venv"...
    python -m venv venv
) else (
    echo [2/4] Virtual environment "venv" sudah ada, dipakai ulang.
)

echo [3/4] Mengaktifkan virtual environment ^& menginstall dependency...
call venv\Scripts\activate.bat
python -m pip install --upgrade pip -q
pip install -r requirements.txt -q

echo [4/4] Menjalankan aplikasi...
echo.
echo ==================================================
echo  Aplikasi siap. Buka browser ke: http://localhost:5000
echo  Tekan CTRL+C di jendela ini untuk menghentikan server.
echo ==================================================
echo.
python app.py

pause
