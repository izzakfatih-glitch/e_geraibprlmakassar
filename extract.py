"""
Modul ekstraksi data & gambar dari:
1. Draft Proposal PKKPRL (PDF)
2. Laporan Teknis Hidro-Oseanografi / Kondisi Eksisting Ekosistem (PDF)

Catatan teknis: PDF sumber mengekstrak teks dengan satu kata per baris pada
beberapa bagian, sehingga semua regex dijalankan terhadap teks yang sudah
dinormalisasi (semua whitespace/newline diubah menjadi satu spasi).
"""
import re
import hashlib
import fitz  # PyMuPDF


def norm(s):
    s = (s or "").replace("\u200b", " ").replace("\ufeff", " ").replace("\xa0", " ")
    return re.sub(r"\s+", " ", s).strip()


# ----------------------------------------------------------------------
# PROPOSAL PDF
# ----------------------------------------------------------------------

PROPOSAL_LABELS = [
    "Nama Pemohon", "Jabatan Pemohon", "Nama Perusahaan/Instansi", "NIB",
    "NPWP", "Nomor Telepon Selular", "Surat Elektronik", "Jenis Kegiatan",
    "Lokasi Kegiatan", "Nama Perairan", "Luas Kebutuhan Ruang", "KBLI",
    "Tanggal Penyusunan",
]

# Heading pattern -> section tag, dicek berurutan sesuai kemunculan teks
PROPOSAL_HEADINGS = [
    (r"2\.\s*Kegiatan Eksisting", "siteplan"),
    (r"E\.\s*RENCANA TAPAK", "siteplan"),
    (r"D\.\s*PETA LOKASI", "peta_lokasi"),
    (r"II\.\s*INFORMASI PEMANFAATAN", "foto_pantai"),
    (r"1\.\s*Mangrove", "foto_mangrove"),
    (r"2\.\s*Lamun", "foto_mangrove"),
    (r"3\.\s*Terumbu Karang", "terumbu_karang_section"),
    (r"E\.\s*AKSESIBILITAS", "peta_pola_ruang"),
]


def extract_proposal(pdf_path):
    doc = fitz.open(pdf_path)
    full_text_raw = "\n".join(page.get_text() for page in doc)
    full_text = norm(full_text_raw)

    data = {}

    for idx, label in enumerate(PROPOSAL_LABELS):
        next_label = PROPOSAL_LABELS[idx + 1] if idx + 1 < len(PROPOSAL_LABELS) else r"I\.\s*RENCANA"
        pattern = re.escape(label) + r"\s*(.*?)\s*" + next_label
        m = re.search(pattern, full_text, re.DOTALL)
        data[label] = norm(m.group(1)) if m else ""

    lokasi_raw = data.get("Lokasi Kegiatan", "")
    data["_lokasi_parts"] = [p for p in re.split(r"\s{1,}", lokasi_raw) if p.strip()]

    # Lokasi Kegiatan sebenarnya terdiri atas 4 baris terpisah (desa, kecamatan,
    # kabupaten, provinsi) pada dokumen sumber. Ambil ulang dari teks MENTAH
    # (belum dinormalisasi) supaya batas baris tidak hilang.
    m_loc = re.search(r"Lokasi Kegiatan\s*\n(.*?)\nNama Perairan", full_text_raw, re.DOTALL)
    if m_loc:
        lines = [norm(x) for x in m_loc.group(1).split("\n") if norm(x)]
        if len(lines) >= 4:
            data["_lokasi_parts"] = lines[:4]
        elif lines:
            data["_lokasi_parts"] = lines

    m = re.search(r"PT\.\s*[A-Z .]+?(?=\s+yang diwakili|\s+berencana)", full_text)
    if m:
        data["Nama Perusahaan/Instansi"] = norm(m.group(0))

    koord = re.findall(
        r"(\d)\s+(\d+°\s*\d+'\s*[\d,]+\"\s*BT)\s+(\d+°\s*\d+'\s*[\d,]+\"\s*LS)",
        full_text,
    )
    data["koordinat"] = koord

    m = re.search(r"komitmen\s*pendanaan\s*investasi\s*secara\s*keseluruhan\s*sebesar\s*([\d.,]+)", full_text)
    data["investasi"] = m.group(1).rstrip(".") if m else ""

    m = re.search(r"berjumlah\s*(\d+)\s*orang\s*per\s*siklus", full_text)
    data["tenaga_kerja"] = m.group(1) if m else ""
    m = re.search(r"tenaga\s*kerja\s*asing\s*berjumlah\s*(\d+)", full_text)
    data["tenaga_kerja_asing"] = m.group(1) if m else "0"

    m = re.search(
        r"didominasi\s*oleh\s*([A-Za-z][A-Za-z .]+?)\s*dengan\s*persentase\s*tutupan\s*mangrove\s*(\d+)%\s*kondisi\s*([A-Za-z ]+?)\s*serta",
        full_text,
    )
    if m:
        data["mangrove_spesies"] = norm(m.group(1))
        data["mangrove_persen"] = m.group(2)
        data["mangrove_kondisi"] = norm(m.group(3))

    m = re.search(r"seluas\s*(\d+)\s*Ha\s*dengan\s*jumlah\s*penduduk\s*sebanyak\s*([\d.]+)", full_text)
    if m:
        data["desa_luas_ha"] = m.group(1)
        data["desa_penduduk"] = m.group(2).rstrip(".")

    data["non_reklamasi"] = "tanpa reklamasi" in full_text.lower()
    data["kegiatan_berusaha"] = "kegiatan berusaha" in full_text.lower()
    data["non_strategis"] = "non-strategis nasional" in full_text.lower()

    # ---------------- IMAGES (sequential heading-tracking) ----------------
    seen_hash = set()
    images = []
    current_section = None
    section_img_count = 0

    for pnum in range(len(doc)):
        page = doc[pnum]
        page_text_norm = norm(page.get_text())

        matches = []
        for pattern, tag in PROPOSAL_HEADINGS:
            for mm in re.finditer(pattern, page_text_norm, re.IGNORECASE):
                matches.append((mm.start(), tag))
        if matches:
            matches.sort(key=lambda x: x[0])
            new_section = matches[-1][1]
            if new_section != current_section:
                current_section = new_section
                section_img_count = 0

        imglist = page.get_images(full=True)
        for i, img in enumerate(imglist):
            xref = img[0]
            base = doc.extract_image(xref)
            data_bytes = base["image"]
            h = hashlib.md5(data_bytes).hexdigest()
            w, ht = base.get("width", 0), base.get("height", 0)
            if w < 150 or ht < 150:
                continue

            if current_section == "terumbu_karang_section":
                tag = "foto_pantai" if section_img_count == 0 else "foto_karang_insitu"
            else:
                tag = current_section or "lainnya"
            section_img_count += 1

            if h in seen_hash:
                continue
            seen_hash.add(h)

            images.append({
                "tag": tag, "bytes": data_bytes, "ext": base["ext"],
                "width": w, "height": ht, "page": pnum + 1,
            })

    doc.close()
    return data, images


# ----------------------------------------------------------------------
# LAPORAN HIDRO-OSEANOGRAFI PDF
# ----------------------------------------------------------------------

# Urutan gambar sesuai template (dikonfirmasi manual): gelombang, arus,
# pasut, ekosistem, batimetri. Dipakai karena resource gambar pada PDF ini
# dibagi/shared di semua halaman sehingga pencocokan caption per-halaman
# tidak selalu reliable; urutan xref lebih konsisten untuk template ini.
LAPORAN_IMAGE_ORDER = ["mawar_gelombang", "mawar_arus", "siklus_pasut", "peta_ekosistem", "profil_batimetri"]


def extract_laporan(pdf_path):
    doc = fitz.open(pdf_path)
    full_text_raw = "\n".join(page.get_text() for page in doc)
    full_text = norm(full_text_raw)
    data = {}

    m = re.search(r"LOKASI TITIK PUSAT RENCANA KEGIATAN\s*(.+?)\s*I\.\s*PENDAHULUAN", full_text)
    data["lokasi_studi"] = norm(m.group(1)) if m else ""

    m = re.search(r"Kedalaman\s*pada\s*titik\s*pusat\s*tercatat\s*sebesar\s*(-?[\d.]+)\s*meter", full_text)
    data["batimetri_titik_pusat"] = m.group(1) if m else ""

    m = re.search(r"panjang\s*lintasan\s*([\d.]+)\s*kilometer", full_text)
    data["batimetri_panjang_lintasan"] = m.group(1) if m else ""

    m = re.search(r"nilai\s*terdalam\s*mencapai\s*(-?[\d.]+)\s*meter", full_text)
    data["batimetri_terdalam"] = m.group(1) if m else ""

    m = re.search(r"Tinggi\s*Gelombang\s*Signifikan\s*\(Hs\)\s*Rata-rata:\s*([\d.]+)\s*meter", full_text)
    data["hs_rata"] = m.group(1) if m else ""

    m = re.search(r"Maksimum\s*Ekstrem:\s*([\d.]+)\s*meter.*?arah\s*dominan\s*dari\s*([\d.]+)\s*derajat\s*\(sektor\s*barat", full_text)
    if m:
        data["hs_maks"], data["hs_arah"] = m.group(1), m.group(2)

    m = re.search(r"Kecepatan\s*Arus\s*Rata-rata:\s*([\d.]+)\s*meter\s*per\s*detik", full_text)
    data["arus_rata"] = m.group(1) if m else ""

    m = re.search(r"Kecepatan\s*Arus\s*Maksimum\s*Ekstrem:\s*([\d.]+)\s*meter\s*per\s*detik.*?arah\s*dominan\s*dari\s*([\d.]+)\s*derajat", full_text)
    if m:
        data["arus_maks"], data["arus_arah"] = m.group(1), m.group(2)

    m = re.search(r"Highest\s*Astronomical\s*Tide\s*\(HAT\):\s*\+?(-?[\d.]+)\s*meter", full_text)
    data["hat"] = m.group(1) if m else ""
    m = re.search(r"Mean\s*Sea\s*Level\s*\(MSL\):\s*(-?[\d.]+)\s*meter", full_text)
    data["msl"] = m.group(1) if m else ""
    m = re.search(r"Lowest\s*Astronomical\s*Tide\s*\(LAT\):\s*(-?[\d.]+)\s*meter", full_text)
    data["lat"] = m.group(1) if m else ""
    m = re.search(r"Tidal\s*Range:\s*([\d.]+)\s*meter", full_text)
    data["tidal_range"] = m.group(1) if m else ""
    m = re.search(r"Bilangan\s*Formzahl:\s*(\d+\.\d+)", full_text)
    data["formzahl"] = m.group(1) if m else ""
    tipe_m = re.search(r"diklasifikasikan.*?sebagai\s*([A-Za-z ]+?),", full_text)
    data["tipe_pasut"] = norm(tipe_m.group(1)) if tipe_m else "Mixed Diurnal"

    m = re.search(r"area\s*rencana\s*kegiatan\s*seluas\s*([\d.]+)\s*Hektar", full_text)
    data["eko_total_ha"] = m.group(1) if m else ""

    m = re.search(r"Terumbu\s*Karang:\s*([\d.]+)\s*Hektar\s*\(([\d.]+)\s*persen", full_text)
    if m:
        data["eko_karang_ha"], data["eko_karang_pct"] = m.group(1), m.group(2)

    m = re.search(r"Lainnya\s*\(termasuk\s*substrat\s*dasar\s*non-terumbu\):\s*([\d.]+)\s*Hektar\s*\(([\d.]+)\s*persen", full_text)
    if m:
        data["eko_lainnya_ha"], data["eko_lainnya_pct"] = m.group(1), m.group(2)

    m = re.search(r"Area\s*Laut\s*Terbuka\s*\(Tanpa\s*Ekosistem\):\s*([\d.]+)\s*Hektar\s*\(([\d.]+)\s*persen", full_text)
    if m:
        data["eko_terbuka_ha"], data["eko_terbuka_pct"] = m.group(1), m.group(2)

    m = re.search(r"jarak\s*ekosistem\s*terdekat\s*dari\s*titik\s*pusat\s*rencana\s*kegiatan\s*adalah\s*([\d.]+)\s*kilometer", full_text)
    data["eko_jarak_terdekat_km"] = m.group(1) if m else ""

    data["ada_lamun"] = "padang lamun teridentifikasi" in full_text.lower()

    # ---------------- IMAGES ----------------
    seen_hash = set()
    images = []
    for pnum in range(len(doc)):
        page = doc[pnum]
        imglist = page.get_images(full=True)
        for i, img in enumerate(imglist):
            xref = img[0]
            base = doc.extract_image(xref)
            data_bytes = base["image"]
            h = hashlib.md5(data_bytes).hexdigest()
            w, ht = base.get("width", 0), base.get("height", 0)
            if w < 150 or ht < 150 or h in seen_hash:
                continue
            seen_hash.add(h)
            order_idx = len(images)
            tag = LAPORAN_IMAGE_ORDER[order_idx] if order_idx < len(LAPORAN_IMAGE_ORDER) else "lainnya"
            images.append({
                "tag": tag, "bytes": data_bytes, "ext": base["ext"],
                "width": w, "height": ht, "page": pnum + 1,
            })
    doc.close()
    return data, images


if __name__ == "__main__":
    import sys
    import json

    prop_data, prop_imgs = extract_proposal(sys.argv[1])
    lap_data, lap_imgs = extract_laporan(sys.argv[2])
    print("=== PROPOSAL DATA ===")
    print(json.dumps({k: v for k, v in prop_data.items() if not k.startswith("_")}, indent=2, ensure_ascii=False))
    print("Images:", [(im["tag"], im["page"], im["width"], im["height"]) for im in prop_imgs])
    print("=== LAPORAN DATA ===")
    print(json.dumps(lap_data, indent=2, ensure_ascii=False))
    print("Images:", [(im["tag"], im["page"], im["width"], im["height"]) for im in lap_imgs])


# ----------------------------------------------------------------------
# WRAPPER DENGAN FALLBACK CLAUDE API
# ----------------------------------------------------------------------

def extract_proposal_with_fallback(pdf_path, use_llm=True, log=print):
    """Sama seperti extract_proposal(), tapi field yang gagal dibaca regex
    akan dicoba diisi ulang lewat Claude API (jika ANTHROPIC_API_KEY diset)."""
    import fitz
    from llm_fallback import llm_fill_missing_fields, PROPOSAL_FIELD_HINTS, api_key_available

    data, images = extract_proposal(pdf_path)

    missing = [k for k in PROPOSAL_FIELD_HINTS if not data.get(k)]
    if missing and use_llm:
        if api_key_available():
            log(f"      -> {len(missing)} field kosong, mencoba fallback Claude API: {missing}")
            doc = fitz.open(pdf_path)
            full_text = "\n".join(p.get_text() for p in doc)
            doc.close()
            filled = llm_fill_missing_fields(full_text, missing, PROPOSAL_FIELD_HINTS)
            for k, v in filled.items():
                if v:
                    data[k] = v
            log(f"      -> berhasil melengkapi {len(filled)} field via Claude API.")
        else:
            log(f"      -> {len(missing)} field kosong, ANTHROPIC_API_KEY tidak diset (lewati fallback).")

    return data, images


def extract_laporan_with_fallback(pdf_path, use_llm=True, log=print):
    import fitz
    from llm_fallback import llm_fill_missing_fields, LAPORAN_FIELD_HINTS, api_key_available

    data, images = extract_laporan(pdf_path)

    missing = [k for k in LAPORAN_FIELD_HINTS if not data.get(k)]
    if missing and use_llm:
        if api_key_available():
            log(f"      -> {len(missing)} field kosong, mencoba fallback Claude API: {missing}")
            doc = fitz.open(pdf_path)
            full_text = "\n".join(p.get_text() for p in doc)
            doc.close()
            filled = llm_fill_missing_fields(full_text, missing, LAPORAN_FIELD_HINTS)
            for k, v in filled.items():
                if v:
                    data[k] = v
            log(f"      -> berhasil melengkapi {len(filled)} field via Claude API.")
        else:
            log(f"      -> {len(missing)} field kosong, ANTHROPIC_API_KEY tidak diset (lewati fallback).")

    return data, images
