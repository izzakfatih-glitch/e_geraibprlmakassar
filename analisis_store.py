"""
Penyimpanan permanen untuk hasil fitur "Analisis & Koreksi Proposal" yang
disimpan pengguna (tombol "Simpan Hasil Analisis") -- supaya bisa dibuka atau
diunduh lagi nanti tanpa perlu upload & analisis ulang dokumennya.

Beda dari job_store.py (yang isinya sementara/hanya untuk alur review->
finalize), data di sini disimpan permanen di disk sampai dihapus manual oleh
pengguna, mirip seperti history.jsonl untuk fitur penggabung proposal.

CATATAN: sama seperti history.jsonl, ini disimpan di disk lokal container --
kalau environment deploy memakai ephemeral filesystem (redeploy/restart bisa
menghapus data), pertimbangkan pindah ke storage eksternal di kemudian hari.
"""
import os
import json
import uuid
import time


def _index_path(store_root):
    return os.path.join(store_root, "index.jsonl")


def _entry_dir(store_root, entry_id):
    return os.path.join(store_root, entry_id)


def simpan_hasil_analisis(store_root, hasil_markdown, nama_proposal, nama_laporan, disimpan_oleh=""):
    """Simpan satu hasil analisis. Mengembalikan entry_id (str)."""
    os.makedirs(store_root, exist_ok=True)
    entry_id = uuid.uuid4().hex[:12]
    entry_dir = _entry_dir(store_root, entry_id)
    os.makedirs(entry_dir, exist_ok=True)

    with open(os.path.join(entry_dir, "hasil.md"), "w", encoding="utf-8") as f:
        f.write(hasil_markdown or "")

    meta = {
        "id": entry_id,
        "waktu": time.strftime("%Y-%m-%d %H:%M:%S"),
        "timestamp": time.time(),
        "nama_proposal": nama_proposal or "",
        "nama_laporan": nama_laporan or "",
        "disimpan_oleh": disimpan_oleh or "",
    }
    with open(os.path.join(entry_dir, "meta.json"), "w", encoding="utf-8") as f:
        json.dump(meta, f, ensure_ascii=False, indent=2)

    with open(_index_path(store_root), "a", encoding="utf-8") as f:
        f.write(json.dumps(meta, ensure_ascii=False) + "\n")

    return entry_id


def list_hasil_analisis(store_root, disimpan_oleh=None, limit=200):
    """Kembalikan daftar metadata hasil analisis tersimpan, terbaru dulu.
    Kalau disimpan_oleh diisi, hanya kembalikan milik pengguna itu."""
    path = _index_path(store_root)
    if not os.path.isfile(path):
        return []
    items = []
    with open(path, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            try:
                item = json.loads(line)
            except json.JSONDecodeError:
                continue
            if disimpan_oleh and item.get("disimpan_oleh") != disimpan_oleh:
                continue
            # pastikan entry-nya masih ada di disk (belum dihapus)
            if os.path.isfile(os.path.join(store_root, item["id"], "meta.json")):
                items.append(item)
    items.sort(key=lambda x: x.get("timestamp", 0), reverse=True)
    return items[:limit]


def load_hasil_analisis(store_root, entry_id):
    """Kembalikan (meta_dict, hasil_markdown) atau None kalau tidak ditemukan."""
    entry_dir = _entry_dir(store_root, entry_id)
    meta_path = os.path.join(entry_dir, "meta.json")
    md_path = os.path.join(entry_dir, "hasil.md")
    if not os.path.isfile(meta_path) or not os.path.isfile(md_path):
        return None
    with open(meta_path, "r", encoding="utf-8") as f:
        meta = json.load(f)
    with open(md_path, "r", encoding="utf-8") as f:
        hasil_markdown = f.read()
    return meta, hasil_markdown


def delete_hasil_analisis(store_root, entry_id):
    import shutil
    entry_dir = _entry_dir(store_root, entry_id)
    shutil.rmtree(entry_dir, ignore_errors=True)
    # Catatan: baris lama di index.jsonl tidak dihapus, tapi list_hasil_analisis()
    # sudah menyaring entry yang foldernya sudah tidak ada -- jadi tetap aman
    # tanpa perlu menulis ulang seluruh file index tiap kali hapus satu entri.
