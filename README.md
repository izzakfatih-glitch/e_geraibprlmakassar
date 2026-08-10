# Aplikasi Web Penggabung Proposal PKKPRL

Versi web dari aplikasi penggabung proposal PKKPRL. Bisa diakses dari
**browser HP atau laptop mana saja** (tidak perlu install Python di
perangkat pengguna), dan **aman dipakai bersamaan oleh banyak orang**
sekaligus.

**Alur pemakaian:**
1. Unggah Draft Proposal PKKPRL + Laporan Kondisi/Hidro-Oseanografi (PDF)
2. Muncul halaman **Review**: pratinjau dokumen lengkap (semua teks &
   gambar) + form untuk mengoreksi data yang salah
3. Klik **"Generate Dokumen Final & Unduh"** → dokumen Word final terunduh

---

## Isi Folder

```
webapp/
├── app.py               <- server web (Flask) + HTML tertanam di dalamnya
├── extract.py           <- mesin pembaca/pengekstrak PDF (+ fallback Claude API)
├── generate_docx.py     <- mesin penyusun dokumen Word
├── review_fields.py     <- daftar field yang bisa dikoreksi di halaman review
├── job_store.py         <- penyimpanan sementara antara tahap review & finalize
├── llm_fallback.py      <- pemanggil Claude API untuk field yang gagal dibaca regex
├── requirements.txt     <- daftar library yang dibutuhkan
├── Procfile             <- konfigurasi untuk platform hosting (Render/Railway)
├── uploads/, outputs/, jobs/  <- folder sementara (otomatis terisi & dibersihkan)
```

---

## 1. Coba Dulu di Komputer Sendiri (opsional, sebelum online-kan)

```bash
pip install -r requirements.txt --break-system-packages
python3 app.py
```

Lalu buka `http://localhost:5000` di browser. Untuk mengakses dari HP di
jaringan WiFi yang sama, cari IP komputer Anda (mis. `192.168.1.5`) dan
buka `http://192.168.1.5:5000` dari HP.

---

## 2. Deploy Online — supaya bisa diakses dari mana saja & banyak orang

### Opsi A: Render.com (paling mudah, gratis, disarankan)

1. Buat akun di **render.com** (bisa daftar pakai GitHub).
2. Upload folder `webapp/` ini ke repository **GitHub** baru (lewat
   github.com, klik "Add file → Upload files", drag semua isi folder).
3. Di Render dashboard: **New +** → **Web Service** → hubungkan ke
   repository GitHub Anda tadi.
4. Isi pengaturan:
   - **Build Command**: `pip install -r requirements.txt`
   - **Start Command**: `gunicorn -w 4 -b 0.0.0.0:$PORT app:app --timeout 120`
5. (Opsional, untuk fallback Claude API) Di bagian **Environment
   Variables**, tambahkan:
   - Key: `ANTHROPIC_API_KEY`
   - Value: API key Anda dari console.anthropic.com
6. Klik **Create Web Service**. Tunggu beberapa menit sampai status
   "Live". Render akan memberi Anda URL seperti:
   `https://proposal-pkkprl.onrender.com`
7. Selesai — bagikan URL itu ke siapa saja, bisa dibuka dari HP/laptop
   mana pun, dan bisa dipakai banyak orang bersamaan.

> Catatan: paket gratis Render akan "tidur" jika tidak diakses 15 menit,
> dan perlu ~30 detik untuk "bangun" lagi saat diakses kembali. Ini wajar
> untuk paket gratis.

### Opsi B: Railway.app (alternatif, langkah serupa dengan Render)

1. Daftar di **railway.app** dengan GitHub.
2. **New Project** → **Deploy from GitHub repo** → pilih repo Anda.
3. Railway otomatis mendeteksi `Procfile` dan `requirements.txt`.
4. Tambahkan environment variable `ANTHROPIC_API_KEY` (opsional) di tab
   **Variables**.
5. Setelah deploy selesai, buka tab **Settings → Networking → Generate
   Domain** untuk mendapatkan URL publik.

### Opsi C: Server/VPS sendiri (untuk yang lebih teknis)

```bash
pip install -r requirements.txt
export ANTHROPIC_API_KEY="sk-ant-..."   # opsional
gunicorn -w 4 -b 0.0.0.0:8000 app:app --timeout 120
```
Lalu arahkan domain/Nginx ke port 8000.

---

## 3. Soal Claude API Key (opsional, untuk ekstraksi yang lebih tahan banting)

- **Tanpa API key**: aplikasi tetap berfungsi penuh, hanya mengandalkan
  regex (pola teks). Ini cukup untuk dokumen dengan format template yang
  sama seperti contoh yang sudah diuji.
- **Dengan API key**: jika suatu field gagal dibaca regex (misalnya
  karena format kalimat sedikit berbeda), aplikasi otomatis bertanya ke
  Claude API untuk mencari nilainya dari teks dokumen.
- API key **cukup diset SEKALI oleh pemilik aplikasi** (Anda) sebagai
  environment variable `ANTHROPIC_API_KEY` di platform hosting (lihat
  langkah 5 di Opsi A). **Pengguna lain yang memakai aplikasi via
  browser TIDAK perlu punya atau memasukkan API key sendiri** — mereka
  cukup unggah PDF dan unduh hasilnya.
- Dapatkan API key di **console.anthropic.com → API Keys → Create Key**.
  Pemakaian API dikenakan biaya per token (cek anthropic.com/pricing).

---

## 4. Keamanan & Privasi Data

- Setiap file yang diunggah diproses di folder sementara unik per
  permintaan, lalu **dihapus otomatis** begitu dokumen hasil selesai
  dikirim ke pengguna — tidak disimpan permanen di server.
- Karena setiap permintaan diproses secara terpisah (folder unik +
  Flask `threaded=True` / gunicorn multi-worker), aplikasi ini aman
  dipakai oleh banyak pengguna secara bersamaan tanpa data yang
  tercampur antar pengguna.

---

## 6. Fitur Baru: Asisten Tanya-Jawab KKPRL (`/asisten`)

Selain fitur generate dokumen, aplikasi ini kini punya halaman **Asisten
e-GeRAI — Tanya KKPRL** yang bisa diakses lewat menu **Bantuan** di
navbar, atau langsung ke `/asisten`.

- Halaman ini **publik** (tidak perlu login staf) — dirancang untuk
  masyarakat umum bertanya seputar persyaratan, alur OSS/e-SEA, biaya
  PNBP, reklamasi, dan tracking permohonan KKPRL.
- Jawaban dijawab otomatis oleh Claude API **di sisi server**
  (`asisten_kkprl.py`), memakai environment variable `ANTHROPIC_API_KEY`
  yang sama seperti fitur fallback ekstraksi (lihat bagian 3 di atas).
  API key **tidak pernah** dikirim/terlihat di browser pengguna.
- Kalau `ANTHROPIC_API_KEY` belum diset, halaman tetap tampil normal,
  tapi asisten akan membalas dengan pesan bahwa layanan belum aktif dan
  mengarahkan ke hotline/e-SEA.
- Basis pengetahuan asisten (system prompt di `asisten_kkprl.py`) berisi
  ringkasan materi sosialisasi KKPRL BPRL Makassar: landasan yuridis,
  tahapan & SLA permohonan, dokumen persyaratan, tarif PNBP (PP 85/2021),
  cek fakta, dan kewajiban pemegang KKPRL. Silakan sunting isi
  `SYSTEM_PROMPT` di file tersebut jika ada pembaruan peraturan.

---

## 7. Batasan

- Didesain untuk template dokumen yang formatnya konsisten (seperti
  dua contoh dokumen yang sudah diuji). Field yang gagal terbaca akan
  ditandai jelas di dokumen hasil: `[data tidak terdeteksi otomatis]`.
- Ukuran unggahan dibatasi 30 MB per file (bisa diubah di `app.py`,
  variabel `MAX_CONTENT_LENGTH`).
