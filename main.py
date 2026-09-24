# -*- coding: utf-8 -*-
"""
main.py -- Entry point gabungan (Flask + FastAPI) dalam SATU proses ASGI
=========================================================================

Railway/Render memberi SATU port dan SATU start command per service, jadi
aplikasi web (app.py, Flask) dan REST API JSON (api_fastapi.py, FastAPI)
digabung jadi satu aplikasi ASGI:

- Route /api/v1/*  -> ditangani FastAPI (api_fastapi.py), lengkap dengan
                      autentikasi X-API-Key, validasi, dan response JSON.
- Semua route lain  -> jatuh (fall-through) ke Flask (app.py): halaman web,
                      login pegawai, upload/review/generate, dsb.

Struktur route pada instance FastAPI dibuat PERTAMA oleh api_fastapi.py
(route /api/v1/...), baru mount "/" (Flask) ditambahkan BELAKANGAN --
Starlette mencocokkan route secara berurutan, jadi API menang untuk
/api/v1/* dan sisanya diteruskan ke Flask. app.py TIDAK diubah sama sekali.

WsgiToAsgi dari library asgiref adalah jembatan standar WSGI -> ASGI
(Flask = WSGI, FastAPI = ASGI).

Jalankan (lokal):
    uvicorn main:app --host 0.0.0.0 --port 8000

Jalankan (produksi, lihat Procfile):
    uvicorn main:app --host 0.0.0.0 --port $PORT --workers 4
"""
from asgiref.wsgi import WsgiToAsgi

from api_fastapi import app as api_app
from app import app as flask_app

# Mount Flask sebagai child terakhir: semua request yang TIDAK cocok dengan
# route /api/v1/* di atas diteruskan ke aplikasi Flask.
api_app.mount("/", WsgiToAsgi(flask_app))

app = api_app
