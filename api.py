# -*- coding: utf-8 -*-
"""
api.py -- REST API (JSON) untuk e-GerAI KKPRL BPRL Makassar
=============================================================

Menyediakan 2 kelompok endpoint sebagai Flask Blueprint terpisah dari
halaman web (app.py) yang sudah ada, supaya bisa dipanggil langsung oleh
aplikasi lain (mobile app, sistem internal, Postman, dsb) dengan request/
response JSON murni -- tanpa perlu login sesi browser.

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

Cara pasang ke app.py yang sudah ada (di dekat baris "app = Flask(__name__)"
setelah app dibuat, atau di baris paling bawah sebelum `if __name__ ==`):

    from api import api_bp
    app.register_blueprint(api_bp)

Autentikasi (opsional tapi disarankan untuk API publik):
    Set environment variable API_KEY di server. Kalau diset, semua
    endpoint /api/v1/* WAJIB menyertakan header:
        X-API-Key: <nilai API_KEY>
    Kalau API_KEY tidak diset, endpoint terbuka tanpa autentikasi
    (cocok untuk uji coba lokal saja -- JANGAN dipakai begitu saja di
    produksi publik tanpa API_KEY).
"""
import os
import uuid
import shutil
import traceback
from functools import wraps

from flask import Blueprint, request, jsonify, send_file, after_this_request

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

api_bp = Blueprint("api_v1", __name__, url_prefix="/api/v1")


# ---------------------------------------------------------------------------
# Util: autentikasi API key (opsional) + helper error JSON + CORS ringan
# ---------------------------------------------------------------------------
def _api_key_required():
    return bool(os.environ.get("API_KEY"))


def require_api_key(view):
    @wraps(view)
    def wrapped(*args, **kwargs):
        if _api_key_required():
            sent = request.headers.get("X-API-Key", "")
            if not sent or sent != os.environ.get("API_KEY"):
                return err("unauthorized", "API key tidak valid atau tidak disertakan (header X-API-Key).", 401)
        return view(*args, **kwargs)
    return wrapped


def err(code, message, http_status=400, **extra):
    body = {"success": False, "error": {"code": code, "message": message}}
    body["error"].update(extra)
    return jsonify(body), http_status


def ok(data=None, http_status=200, **extra):
    body = {"success": True}
    if data is not None:
        body["data"] = data
    body.update(extra)
    return jsonify(body), http_status


@api_bp.after_request
def _add_cors_headers(resp):
    resp.headers["Access-Control-Allow-Origin"] = "*"
    resp.headers["Access-Control-Allow-Headers"] = "Content-Type, X-API-Key"
    resp.headers["Access-Control-Allow-Methods"] = "GET, POST, DELETE, OPTIONS"
    return resp


@api_bp.route("/<path:_any>", methods=["OPTIONS"])
def _cors_preflight(_any):
    return ("", 204)


@api_bp.errorhandler(413)
def _too_large(_e):
    return err("payload_too_large", "Ukuran file yang diunggah melebihi batas maksimum.", 413)


# ---------------------------------------------------------------------------
# 0) Health check
# ---------------------------------------------------------------------------
@api_bp.route("/health", methods=["GET"])
def health():
    return ok({"status": "ok"})


# ---------------------------------------------------------------------------
# 1) ASISTEN TANYA-JAWAB KKPRL
# ---------------------------------------------------------------------------
@api_bp.route("/asisten/status", methods=["GET"])
@require_api_key
def asisten_status():
    aktif = bool(os.environ.get("ANTHROPIC_API_KEY"))
    return ok({"aktif": aktif})


@api_bp.route("/asisten/chat", methods=["POST"])
@require_api_key
def asisten_chat():
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
    payload = request.get_json(silent=True)
    if payload is None:
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
@api_bp.route("/dokumen/fields", methods=["GET"])
@require_api_key
def dokumen_fields():
    """Daftar field yang bisa dikoreksi sebelum generate, dikelompokkan,
    lengkap dengan nama field ('source__key') yang dipakai di endpoint
    /dokumen/generate."""
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


@api_bp.route("/dokumen/ekstrak", methods=["POST"])
@require_api_key
def dokumen_ekstrak():
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
    proposal_file = request.files.get("proposal")
    laporan_file = request.files.get("laporan")

    if not proposal_file or not laporan_file or proposal_file.filename == "" or laporan_file.filename == "":
        return err("missing_files", "Kedua file wajib diunggah: field 'proposal' dan 'laporan'.")

    if not proposal_file.filename.lower().endswith(ALLOWED_EXT):
        return err("invalid_file_type", "File 'proposal' harus berformat PDF atau Word (.docx).")
    if not laporan_file.filename.lower().endswith(ALLOWED_EXT):
        return err("invalid_file_type", "File 'laporan' harus berformat PDF atau Word (.docx).")

    proposal_ext = ".docx" if proposal_file.filename.lower().endswith(".docx") else ".pdf"
    laporan_ext = ".docx" if laporan_file.filename.lower().endswith(".docx") else ".pdf"

    job_store.cleanup_old_jobs(JOBS_DIR)

    job_id = uuid.uuid4().hex[:12]
    tmp_dir = os.path.join(UPLOAD_DIR, job_id)
    os.makedirs(tmp_dir, exist_ok=True)
    proposal_path = os.path.join(tmp_dir, "proposal" + proposal_ext)
    laporan_path = os.path.join(tmp_dir, "laporan" + laporan_ext)
    proposal_file.save(proposal_path)
    laporan_file.save(laporan_path)

    try:
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

    # _lokasi_parts, _sumber_gambar_asli dsb (kalau ada) ikut dikembalikan
    # apa adanya supaya klien bisa lihat semua data hasil ekstraksi.
    return ok({
        "job_id": job_id,
        "prop_data": prop_data,
        "lap_data": lap_data,
        "expires_in_seconds": job_store.JOB_MAX_AGE_SECONDS,
    })


@api_bp.route("/dokumen/generate", methods=["POST"])
@require_api_key
def dokumen_generate():
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
    payload = request.get_json(silent=True) or {}
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

    # apply_form_values menerima objek mirip dict (mendukung 'in' dan
    # '.get'); dict koreksi dari JSON sudah cocok dipakai langsung.
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

    @after_this_request
    def cleanup(response):
        try:
            os.remove(output_path)
        except OSError:
            pass
        return response

    perusahaan = (prop_data.get("Nama Perusahaan/Instansi") or "PKKPRL").replace(" ", "_").replace(".", "")
    download_name = f"Proposal_Teknis_PKKPRL_{perusahaan}.docx"

    return send_file(
        output_path,
        as_attachment=True,
        download_name=download_name,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )


@api_bp.route("/dokumen/job/<job_id>", methods=["DELETE"])
@require_api_key
def dokumen_hapus_job(job_id):
    """Hapus job hasil ekstraksi yang belum jadi di-generate (opsional,
    untuk housekeeping / kalau pengguna batal melanjutkan)."""
    job_store.delete_job(JOBS_DIR, job_id)
    return ok({"deleted": True, "job_id": job_id})
