#!/bin/bash
# ==========================================================================
# Script Jalan Otomatis - Aplikasi Penggabung Proposal PKKPRL (Mac / Linux)
# ==========================================================================
# Cara pakai:
#   1. Buka Terminal di folder ini (folder yang berisi file app.py)
#   2. Jalankan:  bash jalankan_mac_linux.sh
#   3. Buka browser ke http://localhost:5000
#
# Script ini akan otomatis:
#   - Membuat virtual environment (folder "venv") kalau belum ada
#   - Mengaktifkan virtual environment tersebut
#   - Menginstall semua dependency dari requirements.txt di dalam venv
#   - Menjalankan aplikasinya (python3 app.py)
# ==========================================================================

set -e
cd "$(dirname "$0")"

PYTHON_BIN="python3"
if ! command -v python3 &> /dev/null; then
    if command -v python &> /dev/null; then
        PYTHON_BIN="python"
    else
        echo "ERROR: Python tidak ditemukan. Silakan install Python 3.10+ dari https://www.python.org/downloads/"
        exit 1
    fi
fi

echo "[1/4] Menggunakan Python: $($PYTHON_BIN --version)"

if [ ! -d "venv" ]; then
    echo "[2/4] Membuat virtual environment baru di folder 'venv'..."
    "$PYTHON_BIN" -m venv venv
else
    echo "[2/4] Virtual environment 'venv' sudah ada, dipakai ulang."
fi

echo "[3/4] Mengaktifkan virtual environment & menginstall dependency..."
source venv/bin/activate
pip install --upgrade pip -q
pip install -r requirements.txt -q

echo "[4/4] Menjalankan aplikasi..."
echo ""
echo "=================================================="
echo " Aplikasi siap. Buka browser ke: http://localhost:5000"
echo " Tekan CTRL+C di sini untuk menghentikan server."
echo "=================================================="
echo ""
python3 app.py
