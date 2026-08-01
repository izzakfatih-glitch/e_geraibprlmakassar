"""
Aplikasi Web Penggabung Proposal PKKPRL
=========================================
Versi web dari aplikasi CLI sebelumnya. Bisa diakses dari browser HP/laptop
apa saja, dan aman dipakai bersamaan oleh banyak pengguna (setiap request
diproses di folder sementara yang terpisah, tidak saling bercampur).

CARA MENJALANKAN (LOKAL / TES):
    pip install flask pymupdf python-docx anthropic --break-system-packages
    export ANTHROPIC_API_KEY="sk-ant-..."   # opsional, untuk fallback ekstraksi
    python3 app.py
    -> buka http://localhost:5000 di browser

CARA DEPLOY ONLINE (lihat README.md untuk panduan lengkap step-by-step):
    - Render.com / Railway.app (gratis, disarankan untuk pemula)
    - atau server sendiri dengan gunicorn: gunicorn -w 4 -b 0.0.0.0:8000 app:app
"""
import os
import uuid
import shutil
import traceback
from flask import Flask, request, render_template, send_file, after_this_request

from extract import extract_proposal_with_fallback, extract_laporan_with_fallback
from generate_docx import build_document

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
OUTPUT_DIR = os.path.join(BASE_DIR, "outputs")
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

MAX_CONTENT_LENGTH = 30 * 1024 * 1024  # 30 MB batas unggah per file gabungan

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = MAX_CONTENT_LENGTH


@app.route("/", methods=["GET"])
def index():
    return render_template("upload.html", error=None)


@app.route("/health", methods=["GET"])
def health():
    return {"status": "ok"}


@app.route("/generate", methods=["POST"])
def generate():
    proposal_file = request.files.get("proposal")
    laporan_file = request.files.get("laporan")

    if not proposal_file or not laporan_file or proposal_file.filename == "" or laporan_file.filename == "":
        return render_template("upload.html", error="Mohon unggah kedua file PDF (proposal & laporan)."), 400

    if not proposal_file.filename.lower().endswith(".pdf") or not laporan_file.filename.lower().endswith(".pdf"):
        return render_template("upload.html", error="Kedua file harus berformat PDF."), 400

    # Folder unik per-request supaya aman dipakai banyak pengguna bersamaan
    job_id = uuid.uuid4().hex[:12]
    job_dir = os.path.join(UPLOAD_DIR, job_id)
    os.makedirs(job_dir, exist_ok=True)

    proposal_path = os.path.join(job_dir, "proposal.pdf")
    laporan_path = os.path.join(job_dir, "laporan.pdf")
    proposal_file.save(proposal_path)
    laporan_file.save(laporan_path)

    output_path = os.path.join(OUTPUT_DIR, f"Proposal_Final_{job_id}.docx")

    try:
        prop_data, prop_images = extract_proposal_with_fallback(proposal_path, log=lambda *_: None)
        lap_data, lap_images = extract_laporan_with_fallback(laporan_path, log=lambda *_: None)
        build_document(prop_data, prop_images, lap_data, lap_images, output_path)
    except Exception:
        traceback.print_exc()
        shutil.rmtree(job_dir, ignore_errors=True)
        return render_template(
            "upload.html",
            error="Terjadi kesalahan saat memproses dokumen. Pastikan kedua file adalah PDF yang valid, lalu coba lagi.",
        ), 500
    finally:
        shutil.rmtree(job_dir, ignore_errors=True)  # hapus PDF asli setelah selesai diproses

    @after_this_request
    def cleanup(response):
        # hapus file hasil dari server setelah terkirim ke pengguna
        try:
            os.remove(output_path)
        except OSError:
            pass
        return response

    perusahaan = prop_data.get("Nama Perusahaan/Instansi", "PKKPRL").replace(" ", "_").replace(".", "")
    download_name = f"Proposal_Teknis_PKKPRL_{perusahaan}.docx"

    return send_file(
        output_path,
        as_attachment=True,
        download_name=download_name,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, threaded=True, debug=False)
