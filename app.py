"""
Aplikasi Web Penggabung Proposal PKKPRL
=========================================
Versi web dari aplikasi CLI sebelumnya. Bisa diakses dari browser HP/laptop
apa saja, dan aman dipakai bersamaan oleh banyak pengguna (setiap request
diproses di folder sementara yang terpisah, tidak saling bercampur).

CATATAN VERSI INI: HTML form ditanam langsung di dalam file ini (tidak lagi
memakai folder templates/ terpisah), supaya tidak ada masalah TemplateNotFound
akibat perbedaan cara resolusi path di berbagai platform hosting.

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
from flask import Flask, request, render_template_string, send_file, after_this_request

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

print(f"[startup] BASE_DIR = {BASE_DIR}")
print(f"[startup] Isi BASE_DIR = {os.listdir(BASE_DIR)}")

UPLOAD_HTML = """<!DOCTYPE html>
<html lang="id">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Penggabung Proposal PKKPRL</title>
<style>
  :root { --navy:#1F4E79; --bg:#f4f6f8; }
  * { box-sizing: border-box; }
  body {
    font-family: -apple-system, Segoe UI, Roboto, Arial, sans-serif;
    background: var(--bg);
    margin: 0; padding: 0;
    color: #222;
  }
  .wrap { max-width: 640px; margin: 0 auto; padding: 32px 20px 60px; }
  h1 { color: var(--navy); font-size: 22px; margin-bottom: 4px; }
  p.sub { color: #555; margin-top: 0; font-size: 14px; }
  .card {
    background: #fff; border-radius: 12px; padding: 24px;
    box-shadow: 0 2px 10px rgba(0,0,0,0.06); margin-top: 20px;
  }
  label { display:block; font-weight:600; margin-bottom:6px; margin-top:18px; font-size: 14px; }
  input[type=file] {
    display:block; width:100%; padding: 10px; border: 1px dashed #aaa;
    border-radius: 8px; background:#fafafa; font-size: 13px;
  }
  button {
    margin-top: 24px; width: 100%; background: var(--navy); color:#fff;
    border:none; padding: 14px; border-radius: 8px; font-size: 15px;
    font-weight: 600; cursor:pointer;
  }
  button:hover { background:#163a5c; }
  .note { font-size: 12px; color:#777; margin-top: 16px; line-height:1.5; }
  .flash { background:#fff3cd; border:1px solid #ffe08a; padding:12px 14px;
           border-radius:8px; margin-top:16px; font-size:13px; color:#7a5b00; }
  .spinner { display:none; text-align:center; margin-top:18px; font-size:13px; color:var(--navy); }
</style>
</head>
<body>
<div class="wrap">
  <h1>\U0001F4C4 Penggabung Proposal PKKPRL</h1>
  <p class="sub">Unggah Draft Proposal PKKPRL dan Laporan Kondisi/Hidro-Oseanografi (PDF) &mdash; dokumen Word final otomatis dibuat dan bisa langsung diunduh.</p>

  {% if error %}
    <div class="flash">\u26A0 {{ error }}</div>
  {% endif %}

  <div class="card">
    <form method="POST" action="/generate" enctype="multipart/form-data" id="genForm">
      <label>1. Draft Proposal PKKPRL (PDF)</label>
      <input type="file" name="proposal" accept="application/pdf" required>

      <label>2. Laporan Kondisi Eksisting / Hidro-Oseanografi (PDF)</label>
      <input type="file" name="laporan" accept="application/pdf" required>

      <button type="submit">Gabungkan &amp; Unduh Dokumen Word</button>
      <div class="spinner" id="spinner">\u23F3 Memproses dokumen, mohon tunggu...</div>
    </form>
  </div>

  <div class="note">
    Dokumen Anda diproses sementara di server hanya untuk pembuatan file ini,
    dan tidak disimpan permanen. Setiap pengguna diproses secara terpisah,
    sehingga aplikasi ini aman dipakai bersamaan oleh banyak orang.
  </div>
</div>
<script>
document.getElementById('genForm').addEventListener('submit', function() {
  document.getElementById('spinner').style.display = 'block';
});
</script>
</body>
</html>"""


@app.route("/", methods=["GET"])
def index():
    return render_template_string(UPLOAD_HTML, error=None)


@app.route("/health", methods=["GET"])
def health():
    return {"status": "ok"}


@app.route("/generate", methods=["POST"])
def generate():
    proposal_file = request.files.get("proposal")
    laporan_file = request.files.get("laporan")

    if not proposal_file or not laporan_file or proposal_file.filename == "" or laporan_file.filename == "":
        return render_template_string(UPLOAD_HTML, error="Mohon unggah kedua file PDF (proposal & laporan)."), 400

    if not proposal_file.filename.lower().endswith(".pdf") or not laporan_file.filename.lower().endswith(".pdf"):
        return render_template_string(UPLOAD_HTML, error="Kedua file harus berformat PDF."), 400

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
        return render_template_string(
            UPLOAD_HTML,
            error="Terjadi kesalahan saat memproses dokumen. Pastikan kedua file adalah PDF yang valid, lalu coba lagi.",
        ), 500
    finally:
        shutil.rmtree(job_dir, ignore_errors=True)

    @after_this_request
    def cleanup(response):
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
