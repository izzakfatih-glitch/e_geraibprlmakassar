# -*- coding: utf-8 -*-
"""
api_fastapi.py -- REST API (JSON) untuk e-GerAI KKPRL BPRL Makassar, versi FastAPI
====================================================================================

Reimplementasi endpoint yang sama persis dengan `api.py` (Flask Blueprint),
kali ini sebagai aplikasi FastAPI ASGI yang berdiri sendiri, terpisah dari
halaman web (app.py) yang sudah ada. Cocok dipakai langsung oleh aplikasi
lain (mobile app, sistem internal, Postman, dsb) dengan request/response
JSON murni -- tanpa perlu login sesi browser.

1) ASISTEN TANYA-JAWAB KKPRL
   POST /api/v1/asisten/chat        -> tanya jawab ke asisten (Claude API)
   GET  /api/v1/asisten/status      -> cek apakah asisten aktif (API key
                                        Anthropic sudah diset di server)

2) GENERATE DOKUMEN (Proposal Teknis PKKPRL)
   GET  /api/v1/dokumen/fields      -> daftar field yang bisa dikoreksi
   POST /api/v1/dokumen/ekstrak     -> upload PDF/Word Proposal + Laporan,
                                        server mengekstrak datanya dan
                                        mengembalikan job_id + data JSON
   POST /api/v1/dokumen/generate    -> kirim job_id + (opsional) koreksi
                                        data -> server membangun dokumen
                                        Word final dan mengembalikannya
                                        sebagai file untuk diunduh
   DELETE /api/v1/dokumen/job/<id>  -> batalkan/hapus job yang belum
                                        di-generate (opsional, buat
                                        housekeeping)

Cara menjalankan (terpisah dari app.py/Flask):
    pip install fastapi "uvicorn[standard]" python-multipart
    uvicorn api_fastapi:app --host 0.0.0.0 --port 8001

Autentikasi (opsional tapi disarankan untuk API publik):
    Set environment variable API_KEY di server. Kalau diset, semua
    endpoint /api/v1/* (kecuali /health) WAJIB menyertakan header:
        X-API-Key: <nilai API_KEY>
    Kalau API_KEY tidak diset, endpoint terbuka tanpa autentikasi
    (cocok untuk uji coba lokal saja -- JANGAN dipakai begitu saja di
    produksi publik tanpa API_KEY).
"""
import os
import uuid
import shutil
import traceback
from typing import Optional

from fastapi import FastAPI, File, Form, Header, UploadFile, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse, FileResponse
from pydantic import BaseModel, Field
from starlette.background import BackgroundTask

try:
    from dotenv import load_dotenv

    load_dotenv(os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env"))
except ImportError:
    pass  # no python-dotenv: ANTHROPIC_API_KEY must come from the process env

from extract import extract_proposal_with_fallback, extract_laporan_with_fallback
from generate_docx import build_document
from review_fields import FIELD_GROUPS, form_field_name, apply_form_values
import asisten_kkprl
import job_store

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
OUTPUT_DIR = os.path.join(BASE_DIR, "outputs")
JOBS_DIR = os.path.join(BASE_DIR, "jobs")
for _d in (UPLOAD_DIR, OUTPUT_DIR, JOBS_DIR):
    os.makedirs(_d, exist_ok=True)

ALLOWED_EXT = (".pdf", ".docx")
MAX_CONTENT_LENGTH = 30 * 1024 * 1024  # 30 MB, sama seperti batas di app.py

app = FastAPI(title="e-GerAI KKPRL BPRL Makassar API", version="1.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False,
    allow_methods=["GET", "POST", "DELETE", "OPTIONS"],
    allow_headers=["Content-Type", "X-API-Key"],
)


# ---------------------------------------------------------------------------
# Util: autentikasi API key (opsional) + helper error JSON, mengikuti bentuk
# response yang sama persis dengan api.py (Flask) supaya klien yang sudah
# terintegrasi tidak perlu berubah.
# ---------------------------------------------------------------------------
def _api_key_required():
    return bool(os.environ.get("API_KEY"))


def err(code, message, http_status=400, **extra):
    body = {"success": False, "error": {"code": code, "message": message}}
    body["error"].update(extra)
    return JSONResponse(body, status_code=http_status)


def ok(data=None, http_status=200, **extra):
    body = {"success": True}
    if data is not None:
        body["data"] = data
    body.update(extra)
    return JSONResponse(body, status_code=http_status)


def check_api_key(x_api_key: Optional[str]):
    """Return None kalau lolos, atau JSONResponse error 401 kalau ditolak."""
    if _api_key_required():
        expected = os.environ.get("API_KEY")
        if not x_api_key or x_api_key != expected:
            return err("unauthorized", "API key tidak valid atau tidak disertakan (header X-API-Key).", 401)
    return None


@app.middleware("http")
async def _limit_upload_size(request: Request, call_next):
    content_length = request.headers.get("content-length")
    if content_length and content_length.isdigit() and int(content_length) > MAX_CONTENT_LENGTH:
        return err("payload_too_large", "Ukuran file yang diunggah melebihi batas maksimum.", 413)
    return await call_next(request)


# ---------------------------------------------------------------------------
# 0) Health check
# ---------------------------------------------------------------------------
@app.get("/api/v1/health")
def health():
    return ok({"status": "ok"})


# ---------------------------------------------------------------------------
# 1) ASISTEN TANYA-JAWAB KKPRL
# ---------------------------------------------------------------------------
@app.get("/api/v1/asisten/status")
def asisten_status(x_api_key: Optional[str] = Header(default=None, alias="X-API-Key")):
    denied = check_api_key(x_api_key)
    if denied is not None:
        return denied
    aktif = bool(os.environ.get("ANTHROPIC_API_KEY"))
    return ok({"aktif": aktif})


class ChatMessage(BaseModel):
    role: str
    content: str


class ChatRequest(BaseModel):
    messages: list = Field(default_factory=list)


@app.post("/api/v1/asisten/chat")
async def asisten_chat(request: Request, x_api_key: Optional[str] = Header(default=None, alias="X-API-Key")):
    """
    Body JSON:
    {
      "messages": [
        {"role": "user", "content": "Apa itu KKPRL?"},
        {"role": "assistant", "content": "KKPRL adalah ..."},
        {"role": "user", "content": "Berapa lama prosesnya?"}
      ]
    }
    Urutan dari yang paling lama -> paling baru. Pesan terakhir harus
    role "user" (pertanyaan yang ingin dijawab).

    Response 200:
    { "success": true, "data": { "reply": "..." } }
    """
    denied = check_api_key(x_api_key)
    if denied is not None:
        return denied

    try:
        payload = await request.json()
    except Exception:
        return err("invalid_json", "Body request harus JSON valid dengan Content-Type: application/json.")

    if not isinstance(payload, dict):
        return err("invalid_json", "Body request harus JSON valid dengan Content-Type: application/json.")

    messages = payload.get("messages")
    if not isinstance(messages, list) or not messages:
        return err("invalid_messages", "Field 'messages' wajib berupa array berisi minimal 1 pesan.")

    clean_messages = []
    for m in messages:
        if not isinstance(m, dict):
            continue
        role = m.get("role")
        content = m.get("content")
        if role in ("user", "assistant") and isinstance(content, str) and content.strip():
            clean_messages.append({"role": role, "content": content[:4000]})

    if not clean_messages:
        return err("invalid_messages", "Tidak ada pesan valid (role harus 'user'/'assistant', content teks tidak kosong).")

    try:
        reply = asisten_kkprl.chat_reply(clean_messages)
    except Exception:
        traceback.print_exc()
        return err("internal_error", "Terjadi kesalahan saat memproses percakapan.", 500)

    return ok({"reply": reply})


# ---------------------------------------------------------------------------
# 2) GENERATE DOKUMEN
# ---------------------------------------------------------------------------
@app.get("/api/v1/dokumen/fields")
def dokumen_fields(x_api_key: Optional[str] = Header(default=None, alias="X-API-Key")):
    """Daftar field yang bisa dikoreksi sebelum generate, dikelompokkan,
    lengkap dengan nama field ('source__key') yang dipakai di endpoint
    /dokumen/generate."""
    denied = check_api_key(x_api_key)
    if denied is not None:
        return denied

    groups = []
    for group_name, fields in FIELD_GROUPS:
        groups.append({
            "group": group_name,
            "fields": [
                {
                    "field_name": form_field_name(source, key),
                    "source": source,
                    "key": key,
                    "label": label,
                }
                for source, key, label in fields
            ],
        })
    return ok({"groups": groups})


@app.post("/api/v1/dokumen/ekstrak")
async def dokumen_ekstrak(
    proposal: Optional[UploadFile] = File(default=None),
    laporan: Optional[UploadFile] = File(default=None),
    x_api_key: Optional[str] = Header(default=None, alias="X-API-Key"),
):
    """
    Multipart/form-data:
      - proposal: file PDF/.docx (Draft Proposal PKKPRL)
      - laporan:  file PDF/.docx (Laporan Kondisi/Hidro-Oseanografi)

    Response 200:
    {
      "success": true,
      "data": {
        "job_id": "abcd1234ef56",
        "prop_data": { ... field hasil ekstraksi proposal ... },
        "lap_data": { ... field hasil ekstraksi laporan ... },
        "expires_in_seconds": 7200
      }
    }

    job_id ini dipakai di POST /api/v1/dokumen/generate untuk membangun
    dokumen final (berlaku sekitar 2 jam sebelum dibersihkan otomatis).
    """
    denied = check_api_key(x_api_key)
    if denied is not None:
        return denied

    if not proposal or not laporan or not proposal.filename or not laporan.filename:
        return err("missing_files", "Kedua file wajib diunggah: field 'proposal' dan 'laporan'.")

    if not proposal.filename.lower().endswith(ALLOWED_EXT):
        return err("invalid_file_type", "File 'proposal' harus berformat PDF atau Word (.docx).")
    if not laporan.filename.lower().endswith(ALLOWED_EXT):
        return err("invalid_file_type", "File 'laporan' harus berformat PDF atau Word (.docx).")

    proposal_ext = ".docx" if proposal.filename.lower().endswith(".docx") else ".pdf"
    laporan_ext = ".docx" if laporan.filename.lower().endswith(".docx") else ".pdf"

    job_store.cleanup_old_jobs(JOBS_DIR)

    job_id = uuid.uuid4().hex[:12]
    tmp_dir = os.path.join(UPLOAD_DIR, job_id)
    os.makedirs(tmp_dir, exist_ok=True)
    proposal_path = os.path.join(tmp_dir, "proposal" + proposal_ext)
    laporan_path = os.path.join(tmp_dir, "laporan" + laporan_ext)

    try:
        with open(proposal_path, "wb") as f:
            f.write(await proposal.read())
        with open(laporan_path, "wb") as f:
            f.write(await laporan.read())

        prop_data, prop_images = extract_proposal_with_fallback(proposal_path, log=lambda *_: None)
        lap_data, lap_images = extract_laporan_with_fallback(laporan_path, log=lambda *_: None)
        job_store.save_job(JOBS_DIR, job_id, prop_data, prop_images, lap_data, lap_images)
    except Exception:
        traceback.print_exc()
        shutil.rmtree(tmp_dir, ignore_errors=True)
        job_store.delete_job(JOBS_DIR, job_id)
        return err(
            "extraction_failed",
            "Terjadi kesalahan saat memproses dokumen. Pastikan file Proposal dan Laporan adalah PDF/Word yang valid.",
            500,
        )
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)

    return ok({
        "job_id": job_id,
        "prop_data": prop_data,
        "lap_data": lap_data,
        "expires_in_seconds": job_store.JOB_MAX_AGE_SECONDS,
    })


@app.post("/api/v1/dokumen/generate")
async def dokumen_generate(request: Request, x_api_key: Optional[str] = Header(default=None, alias="X-API-Key")):
    """
    Body JSON:
    {
      "job_id": "abcd1234ef56",
      "koreksi": {
        "prop__Nama_Pemohon": "Budi Santoso",
        "lap__eko_total_ha": "12.5"
      }
    }

    "koreksi" bersifat OPSIONAL -- kalau tidak dikirim/kosong, dokumen
    dibangun langsung dari hasil ekstraksi otomatis tanpa koreksi.
    Nama-nama field yang valid untuk "koreksi" bisa dilihat lewat
    GET /api/v1/dokumen/fields (field "field_name").

    Response 200: file .docx (Content-Type Word) langsung sebagai body
    response, siap diunduh/disimpan oleh klien.

    Response error: JSON { "success": false, "error": {...} }
    """
    denied = check_api_key(x_api_key)
    if denied is not None:
        return denied

    try:
        payload = await request.json()
    except Exception:
        payload = {}
    if not isinstance(payload, dict):
        payload = {}

    job_id = (payload.get("job_id") or "").strip()
    koreksi = payload.get("koreksi") or {}

    if not job_id:
        return err("missing_job_id", "Field 'job_id' wajib diisi (didapat dari respons /dokumen/ekstrak).")
    if not isinstance(koreksi, dict):
        return err("invalid_koreksi", "Field 'koreksi' harus berupa object/dict (field_name -> nilai baru).")

    loaded = job_store.load_job(JOBS_DIR, job_id)
    if not loaded:
        return err(
            "job_not_found",
            "job_id tidak ditemukan atau sudah kedaluwarsa. Mohon panggil ulang /dokumen/ekstrak.",
            404,
        )

    prop_data, prop_images, lap_data, lap_images = loaded

    str_koreksi = {k: ("" if v is None else str(v)) for k, v in koreksi.items()}
    prop_data, lap_data = apply_form_values(str_koreksi, prop_data, lap_data)

    output_path = os.path.join(OUTPUT_DIR, f"Proposal_Final_{job_id}.docx")
    try:
        build_document(prop_data, prop_images, lap_data, lap_images, output_path)
    except Exception:
        traceback.print_exc()
        return err("generate_failed", "Terjadi kesalahan saat membuat dokumen final. Silakan coba lagi.", 500)
    finally:
        job_store.delete_job(JOBS_DIR, job_id)

    perusahaan = (prop_data.get("Nama Perusahaan/Instansi") or "PKKPRL").replace(" ", "_").replace(".", "")
    download_name = f"Proposal_Teknis_PKKPRL_{perusahaan}.docx"

    def _cleanup():
        try:
            os.remove(output_path)
        except OSError:
            pass

    return FileResponse(
        output_path,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename=download_name,
        background=BackgroundTask(_cleanup),
    )


@app.delete("/api/v1/dokumen/job/{job_id}")
def dokumen_hapus_job(job_id: str, x_api_key: Optional[str] = Header(default=None, alias="X-API-Key")):
    """Hapus job hasil ekstraksi yang belum jadi di-generate (opsional,
    untuk housekeeping / kalau pengguna batal melanjutkan)."""
    denied = check_api_key(x_api_key)
    if denied is not None:
        return denied
    job_store.delete_job(JOBS_DIR, job_id)
    return ok({"deleted": True, "job_id": job_id})


# ---------------------------------------------------------------------------
# Handler generik untuk error yang tidak ditangkap secara eksplisit di atas,
# supaya respons tetap berbentuk JSON konsisten { success, error } alih-alih
# halaman error HTML bawaan.
# ---------------------------------------------------------------------------
@app.exception_handler(404)
async def _not_found(request: Request, exc):
    return err("not_found", "Endpoint tidak ditemukan.", 404)


@app.exception_handler(Exception)
async def _unhandled_error(request: Request, exc):
    traceback.print_exc()
    return err("internal_error", "Terjadi kesalahan pada server.", 500)


if __name__ == "__main__":
    import uvicorn
    uvicorn.run("api_fastapi:app", host="0.0.0.0", port=int(os.environ.get("PORT", 8001)))
