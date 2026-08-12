"""
Penyimpanan sementara "job" antara tahap /review dan /finalize.

Setiap kali pengguna upload PDF, data hasil ekstraksi (teks + gambar)
disimpan sementara di folder unik (job_id) supaya bisa dipakai lagi saat
pengguna klik "Generate Dokumen Final" setelah meninjau/mengoreksi di
halaman review -- tanpa perlu upload ulang PDF-nya.

Folder job dibersihkan otomatis:
- segera setelah finalize() berhasil mengirim file, dan
- job lama (>2 jam, mis. ditinggal begitu saja oleh pengguna) dibersihkan
  setiap kali ada job baru dibuat.
"""
import os
import json
import time
import shutil

JOB_MAX_AGE_SECONDS = 2 * 60 * 60  # 2 jam


def cleanup_old_jobs(jobs_root):
    if not os.path.isdir(jobs_root):
        return
    now = time.time()
    for name in os.listdir(jobs_root):
        path = os.path.join(jobs_root, name)
        try:
            if os.path.isdir(path) and (now - os.path.getmtime(path)) > JOB_MAX_AGE_SECONDS:
                shutil.rmtree(path, ignore_errors=True)
        except OSError:
            pass


def save_job(jobs_root, job_id, prop_data, prop_images, lap_data, lap_images):
    job_dir = os.path.join(jobs_root, job_id)
    prop_img_dir = os.path.join(job_dir, "prop_images")
    lap_img_dir = os.path.join(job_dir, "lap_images")
    os.makedirs(prop_img_dir, exist_ok=True)
    os.makedirs(lap_img_dir, exist_ok=True)

    def save_images(images, folder):
        manifest = []
        for i, im in enumerate(images):
            fname = f"{i:03d}_{im['tag']}.{im['ext']}"
            with open(os.path.join(folder, fname), "wb") as f:
                f.write(im["bytes"])
            manifest.append({"tag": im["tag"], "file": fname, "ext": im["ext"]})
        return manifest

    prop_manifest = save_images(prop_images, prop_img_dir)
    lap_manifest = save_images(lap_images, lap_img_dir)

    with open(os.path.join(job_dir, "data.json"), "w", encoding="utf-8") as f:
        json.dump({
            "prop_data": prop_data,
            "lap_data": lap_data,
            "prop_images": prop_manifest,
            "lap_images": lap_manifest,
        }, f, ensure_ascii=False, indent=2)

    return job_dir


def load_job(jobs_root, job_id):
    job_dir = os.path.join(jobs_root, job_id)
    data_path = os.path.join(job_dir, "data.json")
    if not os.path.isfile(data_path):
        return None

    with open(data_path, "r", encoding="utf-8") as f:
        blob = json.load(f)

    def load_images(manifest, folder):
        images = []
        for item in manifest:
            path = os.path.join(folder, item["file"])
            if os.path.isfile(path):
                with open(path, "rb") as f:
                    images.append({"tag": item["tag"], "bytes": f.read(), "ext": item["ext"]})
        return images

    prop_images = load_images(blob["prop_images"], os.path.join(job_dir, "prop_images"))
    lap_images = load_images(blob["lap_images"], os.path.join(job_dir, "lap_images"))

    return blob["prop_data"], prop_images, blob["lap_data"], lap_images


def delete_job(jobs_root, job_id):
    job_dir = os.path.join(jobs_root, job_id)
    shutil.rmtree(job_dir, ignore_errors=True)
