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

3) ANALISIS & KOREKSI PROPOSAL
   POST   /api/v1/analisis/proposal        -> upload proposal (+ laporan
                                               pembanding, bisa banyak
                                               berkas), server mengekstrak
                                               teks lalu minta Claude audit
                                               konsistensinya -> markdown
   POST   /api/v1/analisis/unduh           -> ubah hasil analisis jadi
                                               file Word (.docx) untuk
                                               diunduh
   POST   /api/v1/analisis/simpan          -> simpan hasil analisis
                                               permanen, dapat entry_id
   GET    /api/v1/analisis/riwayat         -> daftar hasil tersimpan
   GET    /api/v1/analisis/riwayat/<id>    -> lihat 1 hasil tersimpan
   GET    /api/v1/analisis/riwayat/<id>/unduh -> unduh hasil tersimpan
                                                 sebagai .docx
   DELETE /api/v1/analisis/riwayat/<id>    -> hapus hasil tersimpan

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
import re
import uuid
import shutil
import traceback
from typing import List, Optional

from fastapi import FastAPI, File, Header, UploadFile, Request
from fastapi.exceptions import RequestValidationError
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse, FileResponse, Response
from pydantic import BaseModel, Field
from starlette.background import BackgroundTask
from starlette.concurrency import run_in_threadpool

try:
    from dotenv import load_dotenv

    load_dotenv(os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env"))
except ImportError:
    pass  # no python-dotenv: ANTHROPIC_API_KEY must come from the process env

from extract import extract_proposal_with_fallback, extract_laporan_with_fallback, extract_full_text_multi
from generate_docx import build_document, markdown_report_to_docx_bytes
from llm_fallback import analisis_konsistensi_proposal
from review_fields import FIELD_GROUPS, form_field_name, apply_form_values
import analisis_store
import asisten_kkprl
import job_store

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
OUTPUT_DIR = os.path.join(BASE_DIR, "outputs")
JOBS_DIR = os.path.join(BASE_DIR, "jobs")
# Ikut pola app.py: data dir bisa dioverride lewat env DATA_DIR (mis. Railway
# volume mount) supaya data analisis ikut tersimpan di tempat yang sama.
ANALISIS_STORE_DIR = os.path.join(os.environ.get("DATA_DIR") or BASE_DIR, "analisis_tersimpan")
for _d in (UPLOAD_DIR, OUTPUT_DIR, JOBS_DIR):
    os.makedirs(_d, exist_ok=True)

ALLOWED_EXT = (".pdf", ".docx")
# Batas & ekstensi yang sama dengan app.py (ALLOWED_ANALISIS_EXT, MAX_ANALISIS_FILES).
ALLOWED_ANALISIS_EXT = (".pdf", ".docx", ".xlsx", ".xlsm")
MAX_ANALISIS_FILES = 10
MAX_CONTENT_LENGTH = 30 * 1024 * 1024  # 30 MB, sama seperti batas di app.py

# job_id & analisis entry_id keduanya = uuid4().hex[:12]. Wajib divalidasi
# sebelum dipakai jadi nama direktori: nilai ".." saja akan membuat
# os.path.join melewati JOBS_DIR/ANALISIS_STORE_DIR (rmtree = hapus semua).
_ID_RE = re.compile(r"[0-9a-f]{12}")

DOCX_MEDIA = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

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
        # .strip() di kedua sisi: nilai tempelan di dashboard hosting sering
        # punya spasi/newline tersembunyi yang membuat perbandingan String
        # mentah selalu gagal (padahal key-nya benar).
        expected = (os.environ.get("API_KEY") or "").strip()
        if not x_api_key or x_api_key.strip() != expected:
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
        # run_in_threadpool: panggilan Claude bisa berjalan puluhan detik --
        # kalau dieksekusi langsung di event loop, seluruh worker ini (dan
        # semua request lain yang masuk ke worker itu) ikut tersendat.
        reply = await run_in_threadpool(asisten_kkprl.chat_reply, clean_messages)
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

        prop_data, prop_images = await run_in_threadpool(extract_proposal_with_fallback, proposal_path, log=lambda *_: None)
        lap_data, lap_images = await run_in_threadpool(extract_laporan_with_fallback, laporan_path, log=lambda *_: None)
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
    if not _ID_RE.fullmatch(job_id):
        return err("invalid_job_id", "Field 'job_id' tidak valid (harus 12 karakter hex dari respons /dokumen/ekstrak).")
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
        await run_in_threadpool(build_document, prop_data, prop_images, lap_data, lap_images, output_path)
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
    if not _ID_RE.fullmatch(job_id):
        # Tolak ".." dkk sebelum masuk job_store.delete_job: nilai itu akan
        # membentuk path JOBS_DIR/../.. dan rmtree-nya bisa menghapus isi
        # direktori project, bukan cuma satu job.
        return err("invalid_job_id", "job_id tidak valid.", 400)
    job_store.delete_job(JOBS_DIR, job_id)
    return ok({"deleted": True, "job_id": job_id})


# ---------------------------------------------------------------------------
# 3) ANALISIS & KOREKSI PROPOSAL
#    Padanan API dari halaman web /analisis-proposal + /analisis-riwayat.
#    Alurnya sama: unggah Proposal (wajib) + Laporan pembanding (opsional),
#    teks semua berkas diekstrak, lalu Claude audit konsistensinya dan
#    menghasilkan laporan Markdown.
# ---------------------------------------------------------------------------
def _analisis_docx_download_name(nama_proposal):
    """Nama file unduhan hasil analisis -- sama pola dengan app.py:
    ASCII-only supaya aman di header Content-Disposition."""
    base = re.sub(r"[^A-Za-z0-9_\-]", "_", (nama_proposal or "").strip()) or "proposal"
    return f"Analisis_Proposal_{base[:60]}.docx"


def _analisis_subjudul(nama_proposal, nama_laporan):
    sub = f"Proposal: {nama_proposal or '-'}"
    if nama_laporan:
        sub += f" · Laporan: {nama_laporan}"
    return sub


def _analisis_docx_response(hasil_markdown, nama_proposal, nama_laporan):
    """Build dokumen .docx dari hasil analisis, kirim sebagai response body."""
    try:
        docx_bytes = markdown_report_to_docx_bytes(
            hasil_markdown, subjudul=_analisis_subjudul(nama_proposal, nama_laporan)
        )
    except Exception:
        traceback.print_exc()
        return err("docx_failed", "Gagal membuat dokumen Word dari hasil analisis.", 500)
    return Response(
        content=docx_bytes,
        media_type=DOCX_MEDIA,
        headers={"Content-Disposition": f'attachment; filename="{_analisis_docx_download_name(nama_proposal)}"'},
    )


def _analisis_saved_name(filename):
    """Bersihkan nama berkas unggahan menjadi nama file penyimpanan yang
    aman (tanpa direktori), mirip secure_filename di app.py tapi tanpa
    dependensi werkzeug."""
    base = os.path.basename(filename or "")
    base = re.sub(r"[^\w.\- ()\[\]]+", "_", base).strip("._") or "berkas"
    return base


@app.post("/api/v1/analisis/proposal")
async def analisis_proposal(
    proposal: List[UploadFile] = File(default=[]),
    laporan: List[UploadFile] = File(default=[]),
    x_api_key: Optional[str] = Header(default=None, alias="X-API-Key"),
):
    """
    Multipart/form-data:
      - proposal: 1..10 file PDF/.docx/.xlsx/.xlsm (Proposal Teknis yang
                  akan diperiksa) -- field boleh diulang per berkas
      - laporan : 0..10 file pembanding (Laporan Kondisi Eksisting /
                  Hidro-Oseanografi); boleh kosong

    Response 200:
    {
      "success": true,
      "data": {
        "hasil_markdown": "## Laporan Analisis ...",
        "nama_proposal": "proposal_a.pdf, proposal_b.docx",
        "nama_laporan": "laporan.pdf"
      }
    }

    "hasil_markdown" bisa langsung dikirim ke POST /analisis/unduh (jadi
    file Word) atau POST /analisis/simpan (disimpan permanen).
    """
    denied = check_api_key(x_api_key)
    if denied is not None:
        return denied

    proposal_files = [f for f in (proposal or []) if f and f.filename]
    laporan_files = [f for f in (laporan or []) if f and f.filename]

    if not proposal_files:
        return err("missing_files", "Mohon unggah dokumen Proposal Teknis PKKPRL lewat field 'proposal' (bisa lebih dari satu berkas).")
    if len(proposal_files) > MAX_ANALISIS_FILES or len(laporan_files) > MAX_ANALISIS_FILES:
        return err("too_many_files", f"Maksimum {MAX_ANALISIS_FILES} berkas per field ('proposal'/'laporan').")

    for f in proposal_files + laporan_files:
        ext = os.path.splitext(f.filename)[1].lower()
        if ext not in ALLOWED_ANALISIS_EXT:
            return err("invalid_file_type", f"Format berkas '{f.filename}' tidak didukung. Mohon unggah file PDF, .docx, atau .xlsx.")

    tmp_dir = os.path.join(UPLOAD_DIR, f"analisis_{uuid.uuid4().hex[:12]}")
    os.makedirs(tmp_dir, exist_ok=True)
    try:
        proposal_paths, laporan_paths = [], []
        for i, f in enumerate(proposal_files):
            path = os.path.join(tmp_dir, f"proposal_{i}_{_analisis_saved_name(f.filename)}")
            with open(path, "wb") as out:
                out.write(await f.read())
            proposal_paths.append(path)
        for i, f in enumerate(laporan_files):
            path = os.path.join(tmp_dir, f"laporan_{i}_{_analisis_saved_name(f.filename)}")
            with open(path, "wb") as out:
                out.write(await f.read())
            laporan_paths.append(path)

        teks_proposal = await run_in_threadpool(extract_full_text_multi, proposal_paths)
        teks_laporan = await run_in_threadpool(extract_full_text_multi, laporan_paths) if laporan_paths else ""
        if not teks_proposal.strip():
            return err("no_text", "Tidak ada teks yang berhasil dibaca dari dokumen Proposal. Pastikan filenya valid dan bukan hasil scan gambar mentah.")

        hasil = await run_in_threadpool(analisis_konsistensi_proposal, teks_proposal, teks_laporan)
        if isinstance(hasil, dict) and hasil.get("error"):
            return err("analysis_failed", hasil["error"], 500)

        return ok({
            "hasil_markdown": hasil,
            "nama_proposal": ", ".join(f.filename for f in proposal_files),
            "nama_laporan": ", ".join(f.filename for f in laporan_files),
        })
    except Exception:
        traceback.print_exc()
        return err("analysis_failed", "Terjadi kesalahan saat menganalisis dokumen. Silakan coba lagi.", 500)
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


@app.post("/api/v1/analisis/unduh")
async def analisis_unduh(request: Request, x_api_key: Optional[str] = Header(default=None, alias="X-API-Key")):
    """
    Body JSON:
    {
      "hasil_markdown": "...",       (wajib, hasil dari /analisis/proposal)
      "nama_proposal": "...",        (opsional, untuk judul & nama file)
      "nama_laporan": "..."          (opsional)
    }

    Response 200: file .docx hasil analisis, siap diunduh.
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

    hasil_markdown = str(payload.get("hasil_markdown") or "")
    if not hasil_markdown.strip():
        return err("invalid_input", "Field 'hasil_markdown' wajib diisi (hasil dari POST /analisis/proposal).")
    nama_proposal = str(payload.get("nama_proposal") or "")
    nama_laporan = str(payload.get("nama_laporan") or "")

    return _analisis_docx_response(hasil_markdown, nama_proposal, nama_laporan)


@app.post("/api/v1/analisis/simpan")
async def analisis_simpan(request: Request, x_api_key: Optional[str] = Header(default=None, alias="X-API-Key")):
    """
    Body JSON:
    {
      "hasil_markdown": "...",       (wajib)
      "nama_proposal": "...",        (opsional)
      "nama_laporan": "...",         (opsional)
      "disimpan_oleh": "budi"        (opsional -- penanda petugas/pengguna,
                                     dipakai untuk filter riwayat; API tidak
                                     punya sesi login, jadi isi manual)
    }

    Response 200: { "entry_id": "abc123..." } -- dipakai di endpoint riwayat.
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

    hasil_markdown = str(payload.get("hasil_markdown") or "")
    if not hasil_markdown.strip():
        return err("invalid_input", "Field 'hasil_markdown' wajib diisi (hasil dari POST /analisis/proposal).")

    try:
        entry_id = await run_in_threadpool(
            analisis_store.simpan_hasil_analisis,
            ANALISIS_STORE_DIR,
            hasil_markdown,
            str(payload.get("nama_proposal") or ""),
            str(payload.get("nama_laporan") or ""),
            str(payload.get("disimpan_oleh") or ""),
        )
    except Exception:
        traceback.print_exc()
        return err("save_failed", "Gagal menyimpan hasil analisis di server.", 500)
    return ok({"entry_id": entry_id})


@app.get("/api/v1/analisis/riwayat")
async def analisis_riwayat_list(request: Request, x_api_key: Optional[str] = Header(default=None, alias="X-API-Key")):
    """
    Query string (semua opsional):
      - disimpan_oleh: hanya tampilkan hasil milik petugas tertentu
      - limit         : jumlah maksimum entri yang dikembalikan (default 200,
                        maks 500)

    Tanpa 'disimpan_oleh', semua hasil tersimpan ikut dikembalikan (karena
    API ini sudah dijaga API key, aksesnya dianggap setara admin).
    """
    denied = check_api_key(x_api_key)
    if denied is not None:
        return denied

    qp = request.query_params
    disimpan_oleh = qp.get("disimpan_oleh") or None
    try:
        limit = int(qp.get("limit") or 200)
    except (TypeError, ValueError):
        return err("invalid_limit", "Query 'limit' harus berupa angka.")
    limit = max(1, min(limit, 500))

    try:
        items = await run_in_threadpool(
            analisis_store.list_hasil_analisis, ANALISIS_STORE_DIR, disimpan_oleh=disimpan_oleh, limit=limit
        )
    except Exception:
        traceback.print_exc()
        return err("internal_error", "Gagal membaca daftar hasil analisis.", 500)
    return ok({"items": items, "total": len(items)})


@app.get("/api/v1/analisis/riwayat/{entry_id}")
async def analisis_riwayat_get(entry_id: str, x_api_key: Optional[str] = Header(default=None, alias="X-API-Key")):
    """Response 200: { "meta": {...}, "hasil_markdown": "..." }"""
    denied = check_api_key(x_api_key)
    if denied is not None:
        return denied
    if not _ID_RE.fullmatch(entry_id):
        return err("invalid_entry_id", "entry_id tidak valid.", 400)
    loaded = await run_in_threadpool(analisis_store.load_hasil_analisis, ANALISIS_STORE_DIR, entry_id)
    if not loaded:
        return err("entry_not_found", "Hasil analisis tidak ditemukan (mungkin sudah dihapus).", 404)
    meta, hasil_markdown = loaded
    return ok({"meta": meta, "hasil_markdown": hasil_markdown})


@app.get("/api/v1/analisis/riwayat/{entry_id}/unduh")
async def analisis_riwayat_unduh(entry_id: str, x_api_key: Optional[str] = Header(default=None, alias="X-API-Key")):
    """Response 200: file .docx dari hasil analisis tersimpan tersebut."""
    denied = check_api_key(x_api_key)
    if denied is not None:
        return denied
    if not _ID_RE.fullmatch(entry_id):
        return err("invalid_entry_id", "entry_id tidak valid.", 400)
    loaded = await run_in_threadpool(analisis_store.load_hasil_analisis, ANALISIS_STORE_DIR, entry_id)
    if not loaded:
        return err("entry_not_found", "Hasil analisis tidak ditemukan (mungkin sudah dihapus).", 404)
    meta, hasil_markdown = loaded
    return _analisis_docx_response(hasil_markdown, meta.get("nama_proposal", ""), meta.get("nama_laporan", ""))


@app.delete("/api/v1/analisis/riwayat/{entry_id}")
async def analisis_riwayat_hapus(entry_id: str, x_api_key: Optional[str] = Header(default=None, alias="X-API-Key")):
    """Hapus permanen satu hasil analisis tersimpan."""
    denied = check_api_key(x_api_key)
    if denied is not None:
        return denied
    if not _ID_RE.fullmatch(entry_id):
        return err("invalid_entry_id", "entry_id tidak valid.", 400)
    loaded = await run_in_threadpool(analisis_store.load_hasil_analisis, ANALISIS_STORE_DIR, entry_id)
    if not loaded:
        return err("entry_not_found", "Hasil analisis tidak ditemukan (mungkin sudah dihapus).", 404)
    await run_in_threadpool(analisis_store.delete_hasil_analisis, ANALISIS_STORE_DIR, entry_id)
    return ok({"deleted": True, "entry_id": entry_id})


# ---------------------------------------------------------------------------
# Route cadangan untuk path /api/v1/* yang tidak cocok route manapun di atas.
# Di mode gabungan (main.py), mount Flask di root akan menangkap SEMUA path --
# tanpa route ini, /api/v1/<salah> jatuh ke Flask dan membalas halaman HTML
# 404 alih-alih envelope JSON. Ditaruh setelah route asli & sebelum mount.
# ---------------------------------------------------------------------------
@app.api_route(
    "/api/v1/{unmatched_path:path}",
    methods=["GET", "POST", "PUT", "PATCH", "DELETE", "HEAD", "OPTIONS"],
)
async def api_v1_not_found(unmatched_path: str):
    return err("not_found", "Endpoint tidak ditemukan.", 404)


# ---------------------------------------------------------------------------
# Handler generik untuk error yang tidak ditangkap secara eksplisit di atas,
# supaya respons tetap berbentuk JSON konsisten { success, error } alih-alih
# halaman error HTML bawaan.
# ---------------------------------------------------------------------------
@app.exception_handler(RequestValidationError)
async def _validation_error(request: Request, exc: RequestValidationError):
    # Validasi bawaan FastAPI (mis. body multipart tidak sesuai skema) juga
    # dibungkus ke bentuk envelope yang sama supaya klien cukup cek 'success'.
    first = exc.errors()[0] if exc.errors() else {}
    lokasi = ".".join(str(p) for p in first.get("loc", []) if p != "body")
    detail = first.get("error") or first.get("msg") or "Data request tidak valid."
    return err("validation_error", f"{lokasi}: {detail}".strip(": "), 400)


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
