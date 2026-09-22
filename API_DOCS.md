# REST API — e-GerAI KKPRL BPRL Makassar

API JSON terpisah dari halaman web yang sudah ada, untuk dua fitur:

1. **Asisten Tanya-Jawab KKPRL** (`/api/v1/asisten/*`)
2. **Generate Dokumen Proposal Teknis PKKPRL** (`/api/v1/dokumen/*`)

Base URL: `https://<domain-anda>/api/v1`

---

## Cara Pasang

1. Copy `api.py` ke folder root project (sejajar dengan `app.py`, `extract.py`, dst).
2. Terapkan `app.py.patch` ke `app.py`, atau tambahkan manual 2 potongan berikut:

   a) Tepat setelah `app = Flask(__name__)` dibuat:
   ```python
   from api import api_bp
   app.register_blueprint(api_bp)
   ```

   b) Ganti baris `PUBLIC_PREFIXES` (dekat `PUBLIC_PATHS`) menjadi:
   ```python
   PUBLIC_PREFIXES = ("/static/", "/api/v1/")
   ```
   (Ini wajib — kalau tidak, semua endpoint `/api/v1/*` akan ikut diblokir oleh halaman login pegawai yang berlaku untuk seluruh aplikasi.)

3. (Opsional tapi disarankan) Set environment variable `API_KEY` di server (Render/Railway/VPS) dengan nilai rahasia bebas Anda pilih. Kalau diset, semua endpoint (kecuali `/health`) wajib menyertakan header `X-API-Key`. Kalau tidak diset, API terbuka tanpa autentikasi — cocok untuk uji coba lokal saja.

Tidak ada perubahan lain pada `app.py` — semua route/halaman web yang sudah ada tetap jalan seperti biasa.

---

## Autentikasi

Kalau `API_KEY` diset di server, sertakan header di setiap request:
```
X-API-Key: <nilai API_KEY>
```
Tanpa/salah header → `401 Unauthorized`.

---

## 1) Asisten Tanya-Jawab KKPRL

### `GET /api/v1/asisten/status`
Cek apakah asisten aktif (`ANTHROPIC_API_KEY` sudah diset di server).

```bash
curl https://domain-anda/api/v1/asisten/status -H "X-API-Key: xxx"
```
```json
{ "success": true, "data": { "aktif": true } }
```

### `POST /api/v1/asisten/chat`
Kirim riwayat percakapan, dapat balasan asisten.

```bash
curl -X POST https://domain-anda/api/v1/asisten/chat \
  -H "Content-Type: application/json" -H "X-API-Key: xxx" \
  -d '{
    "messages": [
      {"role": "user", "content": "Apa itu KKPRL?"}
    ]
  }'
```
```json
{ "success": true, "data": { "reply": "KKPRL adalah ..." } }
```
- `messages`: array, urut lama → baru, pesan terakhir wajib `role: "user"`.
- Kalau `ANTHROPIC_API_KEY` belum diset di server, `reply` otomatis berisi pesan "asisten belum aktif" (tidak error).

---

## 2) Generate Dokumen (Proposal Teknis PKKPRL)

Alur: **ekstrak → (opsional koreksi) → generate**.

### `GET /api/v1/dokumen/fields`
Daftar field yang bisa dikoreksi sebelum generate (dipakai untuk tahu nama field yang valid).

```bash
curl https://domain-anda/api/v1/dokumen/fields -H "X-API-Key: xxx"
```
```json
{
  "success": true,
  "data": {
    "groups": [
      {
        "group": "Identitas Pemohon",
        "fields": [
          {"field_name": "prop__Nama_Pemohon", "source": "prop", "key": "Nama Pemohon", "label": "Nama Pemohon"},
          ...
        ]
      },
      ...
    ]
  }
}
```

### `POST /api/v1/dokumen/ekstrak`
Upload 2 file (`multipart/form-data`), server mengekstrak datanya.

```bash
curl -X POST https://domain-anda/api/v1/dokumen/ekstrak \
  -H "X-API-Key: xxx" \
  -F "proposal=@draft_proposal.pdf" \
  -F "laporan=@laporan_hidro.pdf"
```
```json
{
  "success": true,
  "data": {
    "job_id": "fdeb308be6c7",
    "prop_data": { "Nama Pemohon": "...", "...": "..." },
    "lap_data": { "eko_total_ha": "...", "...": "..." },
    "expires_in_seconds": 7200
  }
}
```
- Format file: PDF atau `.docx`.
- `job_id` berlaku ±2 jam, dipakai di langkah berikutnya.

### `POST /api/v1/dokumen/generate`
Bangun dokumen Word final dari `job_id`, dengan koreksi opsional.

```bash
curl -X POST https://domain-anda/api/v1/dokumen/generate \
  -H "Content-Type: application/json" -H "X-API-Key: xxx" \
  -d '{
    "job_id": "fdeb308be6c7",
    "koreksi": {
      "prop__Nama_Pemohon": "Budi Santoso",
      "lap__eko_total_ha": "12.5"
    }
  }' \
  -o Proposal_Final.docx
```
- Response sukses: **file `.docx` langsung** (bukan JSON) — simpan output curl langsung ke file (`-o`).
- Response gagal: JSON `{ "success": false, "error": {...} }`.
- `koreksi` opsional; nama field ikuti hasil `GET /dokumen/fields` (`field_name`).
- Job otomatis dihapus dari server setelah generate berhasil (sekali pakai).

### `DELETE /api/v1/dokumen/job/{job_id}`
Batalkan job yang belum di-generate (housekeeping opsional).

---

## Format Error

Semua error (kecuali sukses download file) berbentuk:
```json
{
  "success": false,
  "error": { "code": "invalid_messages", "message": "Penjelasan dalam Bahasa Indonesia" }
}
```

Kode umum: `unauthorized` (401), `missing_files` / `invalid_file_type` / `missing_job_id` / `invalid_messages` (400), `job_not_found` (404), `payload_too_large` (413), `extraction_failed` / `generate_failed` / `internal_error` (500).

---

## Status Pengujian

Sudah diuji end-to-end secara lokal:
- `GET /health`, `GET /asisten/status` ✔
- `POST /asisten/chat` (respons normal & fallback saat API key kosong, validasi input) ✔
- `GET /dokumen/fields` ✔
- `POST /dokumen/ekstrak` → `POST /dokumen/generate` (menghasilkan file `.docx` valid, koreksi field diterapkan) ✔
- Proteksi `X-API-Key` (401 tanpa/salah key, 200 dengan key benar) ✔
- Halaman web yang sudah ada (`/`, dsb) tidak terganggu ✔
