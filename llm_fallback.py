"""
Fallback ekstraksi berbasis Claude API.

Dipakai HANYA untuk field yang gagal ditemukan oleh regex di extract.py --
supaya aplikasi tetap bisa mengekstrak data walau format PDF sumber sedikit
berbeda dari template yang sudah diuji (mis. urutan kata berbeda, label
field sedikit berbeda, dsb).

Membutuhkan environment variable ANTHROPIC_API_KEY. Jika tidak diset,
seluruh fungsi di sini otomatis dilewati (aplikasi tetap jalan mode
regex-only, hanya saja field yang gagal akan tetap kosong/tertandai).
"""
import os
import json
import re

MODEL = "claude-sonnet-4-6"
MAX_DOC_CHARS = 20000  # batas aman per dokumen tunggal, dipakai untuk laporan pembanding
MAX_PROPOSAL_CHARS = 50000  # proposal sekarang bisa berupa gabungan beberapa berkas sekaligus


def api_key_available():
    return bool(os.environ.get("ANTHROPIC_API_KEY"))


def analisis_konsistensi_proposal(teks_proposal, teks_laporan):
    """Bandingkan dokumen Proposal Teknis PKKPRL yang SUDAH JADI (ditulis
    manual oleh pengguna jasa, bisa format apa saja) dengan dokumen Laporan
    Kondisi Eksisting/Hidro-Oseanografi sebagai sumber data pembanding.
    Mengembalikan laporan analisis dalam format Markdown (str), atau dict
    {"error": pesan} kalau API tidak tersedia/gagal."""
    if not api_key_available():
        return {"error": "ANTHROPIC_API_KEY belum diset di server -- fitur analisis ini butuh akses Claude API."}

    proposal_trunc = teks_proposal[:MAX_PROPOSAL_CHARS]
    laporan_trunc = teks_laporan[:MAX_DOC_CHARS] if teks_laporan else ""

    prompt = f"""Anda adalah auditor teknis dokumen permohonan PKKPRL (Persetujuan Kesesuaian Kegiatan Pemanfaatan Ruang Laut) di Indonesia. Anda akan diberikan dua dokumen mentah (hasil ekstraksi teks apa adanya, mungkin ada noise format). Setiap dokumen bisa terdiri dari GABUNGAN BEBERAPA BERKAS (PDF/Word/Excel) yang merupakan satu kesatuan -- tiap berkas dipisahkan dengan penanda "===== BERKAS: nama_file =====". Perlakukan semua berkas dalam satu dokumen sebagai satu kesatuan informasi, tapi bila menemukan ketidaksesuaian atau data penting, sebutkan juga di berkas mana temuan itu berada (pakai nama berkas dari penanda tersebut) supaya mudah dirujuk.

=== DOKUMEN 1: PROPOSAL TEKNIS (yang akan diperiksa) ===
{proposal_trunc}

=== DOKUMEN 2: LAPORAN KONDISI EKSISTING / HIDRO-OSEANOGRAFI & EKOSISTEM (sumber data pembanding) ===
{laporan_trunc if laporan_trunc else "(tidak diunggah -- lewati pengecekan konsistensi data, fokus ke kelengkapan Dokumen 1 saja)"}

TUGAS ANDA:
1. Cek KONSISTENSI DATA: bandingkan setiap angka/nilai teknis yang disebutkan di Proposal (tinggi gelombang, kecepatan arus, pasang surut/HAT/LAT, kedalaman/batimetri, jenis & persentase tutupan ekosistem mangrove/lamun/terumbu karang, luas kebutuhan ruang, lokasi, dll) dengan nilai yang tercantum di Laporan sumber (kalau ada). Untuk SETIAP parameter yang disebutkan di SALAH SATU ATAU KEDUA dokumen, laporkan nilainya di masing-masing dan status kecocokannya.
2. Cek KELENGKAPAN: apakah ada bagian standar proposal PKKPRL yang tampak kosong/belum lengkap (misalnya ada teks placeholder semacam "[data tidak terdeteksi otomatis]" yang masih tertinggal, atau bagian penting yang seharusnya ada tapi tidak ditemukan)?
3. Cek KOORDINAT: cari titik koordinat (lintang/bujur atau UTM) yang disebutkan di Proposal. Periksa: (a) apakah formatnya valid (mis. lintang harus -11 s.d. 6, bujur 95 s.d. 141 untuk wilayah Indonesia), (b) apakah koordinat itu secara kasar berada di wilayah yang sesuai dengan Kabupaten/Provinsi yang disebutkan di dokumen (berdasarkan pengetahuan geografis umum Anda), (c) apakah ada koordinat yang tertukar lintang-bujurnya (format terbalik) atau formatnya tidak konsisten antar penyebutan.
4. Cek KONSISTENSI WILAYAH ADMINISTRATIF: periksa apakah hierarki Desa/Kelurahan -> Kecamatan -> Kabupaten/Kota -> Provinsi yang disebutkan di Proposal itu benar-benar cocok satu sama lain secara administratif (mis. apakah Kecamatan yang disebut benar-benar berada di Kabupaten yang disebut, dan Kabupaten itu benar-benar berada di Provinsi yang disebut), berdasarkan pengetahuan Anda tentang pembagian wilayah administratif Indonesia. Tandai juga kalau ada penyebutan lokasi yang TIDAK KONSISTEN antar bagian dokumen (mis. di satu bagian disebut Kabupaten A, di bagian lain Kabupaten B, untuk lokasi kegiatan yang sama).
5. Cek TYPO & KESALAHAN PENULISAN: cari kesalahan ketik, ejaan, atau tata bahasa yang jelas-jelas keliru dalam dokumen (nama tempat yang salah eja, istilah teknis yang salah tulis, dsb).
6. Beri REKOMENDASI PERBAIKAN yang konkret untuk tiap temuan (poin 1-5).

ATURAN KETAT:
- JANGAN mengarang data yang tidak benar-benar ada di salah satu dokumen.
- Kalau suatu parameter tidak disebutkan di salah satu/kedua dokumen, katakan "tidak disebutkan" apa adanya -- jangan menebak angka.
- Untuk pengecekan koordinat & wilayah administratif (poin 3-4): ini berdasarkan pengetahuan geografis umum Anda, BUKAN hasil pengecekan ke basis data resmi (mis. Kemendagri/BIG) -- selalu ingatkan bahwa hasil ini perlu diverifikasi ulang ke sumber resmi sebelum dijadikan dasar keputusan, terutama untuk wilayah yang baru dimekarkan.
- Kalau tidak ada koordinat yang disebutkan di dokumen sama sekali, katakan itu di bagian koordinat (jangan dikosongkan begitu saja).
- Balas HANYA dalam format Markdown berikut, tanpa teks pembuka/penutup apa pun di luar format ini:

## Ringkasan
(1 paragraf singkat kesimpulan keseluruhan: apakah dokumen proposal ini secara umum konsisten dengan data sumber atau ada banyak ketidaksesuaian)

## Ketidaksesuaian & Perbandingan Data
| Parameter | Nilai di Proposal | Nilai di Laporan Sumber | Status |
|---|---|---|---|
(satu baris per parameter yang berhasil dibandingkan; Status salah satu dari: Cocok / Tidak Cocok / Hanya di Proposal / Hanya di Laporan)

## Pengecekan Koordinat
(temuan soal validitas format, kecocokan dengan lokasi administratif yang disebutkan, indikasi lintang-bujur tertukar, dll -- atau nyatakan kalau tidak ada koordinat ditemukan di dokumen)

## Konsistensi Wilayah Administratif
(temuan soal kecocokan hierarki Desa/Kelurahan-Kecamatan-Kabupaten/Kota-Provinsi, dan penyebutan lokasi yang tidak konsisten antar bagian dokumen)

## Typo & Kesalahan Penulisan
(daftar typo/kesalahan ketik/ejaan yang ditemukan beserta perbaikannya; kalau tidak ada, katakan begitu)

## Kelengkapan Dokumen Proposal
(daftar poin -- apa saja yang tampak kosong/placeholder/belum lengkap; kalau semuanya lengkap, katakan begitu)

## Rekomendasi Perbaikan
(daftar bernomor, konkret, urut prioritas, mencakup semua temuan di atas)
"""

    try:
        client = _client()
        resp = client.messages.create(
            model=MODEL,
            max_tokens=6000,
            messages=[{"role": "user", "content": prompt}],
            timeout=130.0,
        )
        hasil = "".join(b.text for b in resp.content if b.type == "text").strip()
        if not hasil:
            return {"error": "Claude API mengembalikan respons kosong. Silakan coba lagi."}
        return hasil
    except Exception as e:
        return {"error": f"Gagal memproses analisis: {type(e).__name__}: {str(e)[:300]}"}


def _client():
    from anthropic import Anthropic
    # PENTING: batasi max_retries secara eksplisit. SDK Anthropic secara
    # default retry hingga 2x kalau request gagal/timeout -- kalau
    # dikombinasikan dengan timeout per-percobaan yang cukup besar, TOTAL
    # waktu tunggu worst-case bisa jauh melebihi timeout worker gunicorn
    # (300 detik), yang membuat gunicorn MEMATIKAN worker-nya secara paksa
    # di tengah request -- koneksi terputus mentah tanpa sempat mengirim
    # respons apa pun (proxy hosting biasanya menampilkan "upstream error"
    # ke pengguna, dan SEMUA data yang sudah benar diisi pengguna -- mis.
    # tabel koordinat, tabel jadwal kegiatan -- ikut tidak pernah sampai ke
    # dokumen yang dihasilkan, karena requestnya mati total sebelum sempat
    # selesai). Membatasi retry di sini mencegah efek domino itu di SEMUA
    # pemanggilan API.
    return Anthropic(max_retries=1)  # otomatis baca ANTHROPIC_API_KEY dari environment


def _client_fast():
    """Klien khusus untuk pemanggilan API yang sifatnya "pemanis" opsional
    (mis. perkuat_narasi_ilmiah, buat_analisis_ekosistem) -- dipanggil
    BERKALI-KALI (bisa sampai 8x) dalam satu proses generate dokumen, dan
    SUDAH punya fallback aman (kembalikan teks asli/kosong) kalau gagal.
    TANPA retry sama sekali supaya satu panggilan yang lambat langsung
    gagal cepat dan lanjut ke bagian berikutnya, bukan menunggu 2x lipat
    lebih lama -- penting karena delapan panggilan berurutan yang masing2
    retry bisa dengan mudah melebihi timeout worker gunicorn kalau terjadi
    bersamaan (mis. saat banyak pengguna memakai aplikasi bersamaan dan API
    sedang agak lambat merespons)."""
    from anthropic import Anthropic
    return Anthropic(max_retries=0)


def llm_fill_missing_fields(full_text, missing_keys, field_hints):
    """
    full_text: teks penuh dokumen (hasil ekstraksi PDF)
    missing_keys: list nama field yang gagal ditemukan regex, mis. ['investasi', 'hs_maks']
    field_hints: dict {key: deskripsi field dalam bahasa manusia} supaya prompt jelas

    Return: dict {key: value} (string). Field yang tetap tidak ditemukan akan
    diberi string kosong "".
    """
    if not missing_keys or not api_key_available():
        return {}

    hints_text = "\n".join(f"- {k}: {field_hints.get(k, k)}" for k in missing_keys)

    prompt = f"""Berikut adalah teks yang diekstrak dari sebuah dokumen PDF teknis (proposal PKKPRL
atau laporan hidro-oseanografi Indonesia). Beberapa nilai data GAGAL ditemukan
secara otomatis oleh sistem berbasis pola teks. Tolong baca teks dokumen di
bawah ini dan cari nilai untuk setiap field yang diminta.

FIELD YANG PERLU DICARI:
{hints_text}

ATURAN:
- Jawab HANYA dengan JSON murni (tanpa markdown, tanpa penjelasan, tanpa ```)
- Format: {{"nama_field": "nilai sebagai string"}}
- Jika suatu nilai benar-benar tidak ada di dalam teks, isi string kosong ""
- Untuk angka, tuliskan apa adanya seperti tertulis di dokumen (jangan diubah formatnya)

TEKS DOKUMEN:
---
{full_text[:15000]}
---

JSON:"""

    try:
        client = _client()
        resp = client.messages.create(
            model=MODEL,
            max_tokens=1500,
            messages=[{"role": "user", "content": prompt}],
            timeout=45.0,
        )
        text = "".join(block.text for block in resp.content if hasattr(block, "text"))
        text = re.sub(r"^```json\s*|\s*```$", "", text.strip())
        result = json.loads(text)
        return {k: str(v) for k, v in result.items() if k in missing_keys}
    except Exception as e:
        print(f"[llm_fallback] gagal memanggil Claude API: {e}")
        return {}


# Deskripsi field-field yang mungkin perlu fallback (dipakai untuk prompt)
PROPOSAL_FIELD_HINTS = {
    "Nama Pemohon": "nama lengkap pemohon/direktur perusahaan",
    "Jabatan Pemohon": "jabatan pemohon di perusahaan",
    "Nama Perusahaan/Instansi": "nama perusahaan/instansi pemohon",
    "NIB": "Nomor Induk Berusaha",
    "NPWP": "Nomor Pokok Wajib Pajak",
    "Nomor Telepon Selular": "nomor telepon pemohon",
    "Surat Elektronik": "alamat email pemohon",
    "Jenis Kegiatan": "jenis kegiatan pemanfaatan ruang laut yang dimohonkan",
    "Nama Perairan": "nama perairan/laut lokasi kegiatan",
    "Luas Kebutuhan Ruang": "luas ruang laut yang dimohonkan (dengan satuan)",
    "KBLI": "kode KBLI kegiatan",
    "investasi": "total nilai investasi/pendanaan kegiatan (angka saja, tanpa 'Rp')",
    "tenaga_kerja": "jumlah tenaga kerja per siklus (angka saja)",
    "mangrove_spesies": "nama spesies mangrove yang dominan di lokasi",
    "mangrove_persen": "persentase tutupan mangrove (angka saja)",
    "mangrove_kondisi": "kondisi/kepadatan tutupan mangrove",
    "desa_luas_ha": "luas wilayah desa dalam Hektar (angka saja)",
    "desa_penduduk": "jumlah penduduk desa (angka saja)",
}

LAPORAN_FIELD_HINTS = {
    "lokasi_studi": "nama lokasi/perairan studi",
    "batimetri_titik_pusat": "kedalaman pada titik pusat lokasi (meter, angka saja, boleh negatif)",
    "batimetri_panjang_lintasan": "panjang lintasan pemeruman batimetri (km, angka saja)",
    "batimetri_terdalam": "kedalaman terdalam pada profil batimetri (meter, angka saja, boleh negatif)",
    "hs_rata": "tinggi gelombang signifikan rata-rata (meter, angka saja)",
    "hs_maks": "tinggi gelombang signifikan maksimum ekstrem (meter, angka saja)",
    "hs_arah": "arah dominan gelombang ekstrem (derajat, angka saja)",
    "arus_rata": "kecepatan arus rata-rata (m/detik, angka saja)",
    "arus_maks": "kecepatan arus maksimum ekstrem (m/detik, angka saja)",
    "arus_arah": "arah dominan arus maksimum (derajat, angka saja)",
    "hat": "elevasi Highest Astronomical Tide/HAT (meter, angka saja)",
    "msl": "elevasi Mean Sea Level/MSL (meter, angka saja)",
    "lat": "elevasi Lowest Astronomical Tide/LAT (meter, angka saja, boleh negatif)",
    "tidal_range": "tunggang pasang surut/tidal range (meter, angka saja)",
    "formzahl": "bilangan Formzahl (angka saja)",
    "eko_total_ha": "total luas area kajian ekosistem (Ha, angka saja)",
    "eko_karang_ha": "luas tutupan terumbu karang (Ha, angka saja)",
    "eko_karang_pct": "persentase tutupan terumbu karang (angka saja)",
    "eko_lainnya_ha": "luas tutupan lainnya/non-terumbu (Ha, angka saja)",
    "eko_lainnya_pct": "persentase tutupan lainnya (angka saja)",
    "eko_terbuka_ha": "luas area laut terbuka tanpa ekosistem (Ha, angka saja)",
    "eko_terbuka_pct": "persentase area laut terbuka (angka saja)",
    "eko_jarak_terdekat_km": "jarak ekosistem terdekat dari titik pusat (km, angka saja)",
}


def buat_analisis_ekosistem(jenis, spesies, persentase, kondisi, konteks_lokasi=""):
    """Buat SATU paragraf analisis ilmiah singkat berdasarkan DATA PRIMER/MANUAL
    hasil isian form pengguna (jenis ekosistem, spesies dominan, persentase
    tutupan, kondisi/status) -- dipakai untuk memperkaya sub-bagian
    Mangrove/Lamun/Terumbu Karang ketika datanya berasal dari isian manual
    (bukan dari narasi dokumen sumber, yang sudah ditangani terpisah lewat
    perkuat_narasi_ilmiah()).

    Kalau ANTHROPIC_API_KEY tidak diset atau proses gagal, kembalikan string
    kosong (TIDAK mengarang analisis apa pun -- paragraf tambahan ini cukup
    dilewati, kalimat template dasar tetap tampil seperti biasa)."""
    if not api_key_available():
        return ""
    try:
        client = _client_fast()
        lokasi_txt = f" di {konteks_lokasi}" if konteks_lokasi else ""
        prompt = (
            f"Data primer/hasil pengamatan lapangan untuk ekosistem {jenis}{lokasi_txt}:\n"
            f"- Spesies/jenis dominan: {spesies or '(tidak disebutkan)'}\n"
            f"- Persentase tutupan: {persentase or '(tidak disebutkan)'}%\n"
            f"- Kondisi/status: {kondisi or '(tidak disebutkan)'}\n\n"
            "Tugas Anda: tulis SATU paragraf analisis ilmiah singkat (3-5 kalimat) "
            "untuk dokumen resmi permohonan PKKPRL, membahas signifikansi ekologis "
            "dari data di atas -- misalnya fungsi ekosistem tersebut bagi lingkungan "
            "pesisir, makna status/kondisi yang ditemukan, serta implikasinya "
            "terhadap rencana kegiatan (mis. kebutuhan mitigasi/pengelolaan) bila "
            "relevan.\n"
            "ATURAN KETAT:\n"
            "1. JANGAN mengarang data/angka baru -- hanya gunakan data yang diberikan di atas.\n"
            "2. JANGAN mengulang kalimat template yang sudah ada sebelumnya, cukup "
            "tambahkan analisis/konteks ilmiah baru.\n"
            "3. Balas HANYA dengan teks paragraf (tanpa pembuka/penutup/markdown apa pun)."
        )
        resp = client.messages.create(
            model=MODEL,
            max_tokens=600,
            messages=[{"role": "user", "content": prompt}],
            timeout=25.0,
        )
        hasil = "".join(b.text for b in resp.content if b.type == "text").strip()
        return hasil
    except Exception:
        return ""


def perkuat_narasi_ilmiah(teks_asli, konteks=""):
    """Perhalus & perkuat narasi asli (dari Laporan Hidro-Oseanografi/kondisi
    eksisting) supaya lebih ilmiah, TANPA menghilangkan informasi apa pun
    dari teks aslinya -- cuma boleh membetulkan tata bahasa dan menambahkan
    1-2 kalimat pendukung yang relevan kalau perlu.

    Kalau ANTHROPIC_API_KEY tidak diset atau proses gagal, teks ASLI
    dikembalikan apa adanya (tidak pernah menghilangkan konten pengguna
    hanya karena AI tidak tersedia)."""
    if not teks_asli or not teks_asli.strip():
        return teks_asli
    if not api_key_available():
        return teks_asli
    try:
        client = _client_fast()
        prompt = (
            "Berikut ini narasi/deskripsi asli dari sebuah laporan teknis kelautan "
            f"(konteks: {konteks}):\n\n\"\"\"\n{teks_asli}\n\"\"\"\n\n"
            "Tugas Anda: perhalus tata bahasa dan perkuat gaya penulisannya supaya "
            "lebih ilmiah dan formal, SESUAI untuk dokumen resmi pemerintah. "
            "ATURAN KETAT:\n"
            "1. JANGAN menghilangkan satu pun informasi/fakta dari teks asli.\n"
            "2. JANGAN mengubah angka atau data apa pun.\n"
            "3. Boleh menambahkan maksimal 1-2 kalimat pendukung yang relevan dan "
            "faktual untuk memperkuat konteks ilmiahnya, tapi jangan mengarang data baru.\n"
            "4. Balas HANYA dengan teks hasil revisi (paragraf biasa), tanpa "
            "pembuka/penutup/penjelasan/markdown apa pun."
        )
        resp = client.messages.create(
            model=MODEL,
            max_tokens=1500,
            messages=[{"role": "user", "content": prompt}],
            timeout=25.0,
        )
        hasil = "".join(b.text for b in resp.content if b.type == "text").strip()
        # Kalau hasilnya kosong/mencurigakan (jauh lebih pendek dari asli,
        # kemungkinan model malah meringkas bukan memperkuat), pakai teks asli.
        if not hasil or len(hasil) < len(teks_asli) * 0.6:
            return teks_asli
        return hasil
    except Exception:
        return teks_asli
