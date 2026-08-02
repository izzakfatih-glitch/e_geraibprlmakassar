"""
Definisi field yang bisa dikoreksi pengguna di halaman /review, dikelompokkan
supaya form-nya rapi. Dipakai untuk merender form (label + value saat ini)
dan untuk memetakan kembali input form ke dict prop_data / lap_data saat
/finalize dipanggil.

Setiap entri: (source, key, label)
  source: "prop" atau "lap" -> menunjuk ke prop_data atau lap_data
  key: nama key di dalam dict tersebut
  label: teks yang ditampilkan ke pengguna
"""

FIELD_GROUPS = [
    ("Identitas Pemohon", [
        ("prop", "Nama Pemohon", "Nama Pemohon"),
        ("prop", "Jabatan Pemohon", "Jabatan Pemohon"),
        ("prop", "Nama Perusahaan/Instansi", "Nama Perusahaan/Instansi"),
        ("prop", "NIB", "NIB"),
        ("prop", "NPWP", "NPWP"),
        ("prop", "Nomor Telepon Selular", "Nomor Telepon Selular"),
        ("prop", "Surat Elektronik", "Surat Elektronik"),
    ]),
    ("Kegiatan & Lokasi", [
        ("prop", "Jenis Kegiatan", "Jenis Kegiatan"),
        ("prop", "Nama Perairan", "Nama Perairan"),
        ("prop", "Luas Kebutuhan Ruang", "Luas Kebutuhan Ruang"),
        ("prop", "KBLI", "KBLI"),
        ("prop", "Tanggal Penyusunan", "Tanggal Penyusunan"),
        ("prop_loc", "3", "Provinsi"),
        ("prop_loc", "2", "Kabupaten"),
        ("prop_loc", "1", "Kecamatan"),
        ("prop_loc", "0", "Desa"),
    ]),
    ("Investasi & Tenaga Kerja", [
        ("prop", "investasi", "Nilai Investasi (Rp, angka saja)"),
        ("prop", "tenaga_kerja", "Jumlah Tenaga Kerja per Siklus"),
        ("prop", "tenaga_kerja_asing", "Jumlah Tenaga Kerja Asing"),
    ]),
    ("Ekosistem Mangrove", [
        ("prop", "mangrove_spesies", "Spesies Mangrove Dominan"),
        ("prop", "mangrove_persen", "Persentase Tutupan Mangrove (%)"),
        ("prop", "mangrove_kondisi", "Kondisi Tutupan Mangrove"),
    ]),
    ("Ekosistem Terumbu Karang & Lamun", [
        ("lap", "eko_total_ha", "Total Area Kajian (Ha)"),
        ("lap", "eko_karang_ha", "Luas Terumbu Karang (Ha)"),
        ("lap", "eko_karang_pct", "Persentase Terumbu Karang (%)"),
        ("lap", "eko_lainnya_ha", "Luas Substrat Non-Terumbu (Ha)"),
        ("lap", "eko_lainnya_pct", "Persentase Non-Terumbu (%)"),
        ("lap", "eko_terbuka_ha", "Luas Area Laut Terbuka (Ha)"),
        ("lap", "eko_terbuka_pct", "Persentase Area Terbuka (%)"),
        ("lap", "eko_jarak_terdekat_km", "Jarak Ekosistem Terdekat (km)"),
    ]),
    ("Batimetri", [
        ("lap", "batimetri_titik_pusat", "Kedalaman Titik Pusat (m)"),
        ("lap", "batimetri_panjang_lintasan", "Panjang Lintasan Pemeruman (km)"),
        ("lap", "batimetri_terdalam", "Kedalaman Terdalam (m)"),
    ]),
    ("Gelombang & Arus", [
        ("lap", "hs_rata", "Tinggi Gelombang Rata-rata (m)"),
        ("lap", "hs_maks", "Tinggi Gelombang Maksimum (m)"),
        ("lap", "hs_arah", "Arah Dominan Gelombang (\u00b0)"),
        ("lap", "arus_rata", "Kecepatan Arus Rata-rata (m/detik)"),
        ("lap", "arus_maks", "Kecepatan Arus Maksimum (m/detik)"),
        ("lap", "arus_arah", "Arah Dominan Arus (\u00b0)"),
    ]),
    ("Pasang Surut", [
        ("lap", "hat", "HAT (m)"),
        ("lap", "msl", "MSL (m)"),
        ("lap", "lat", "LAT (m)"),
        ("lap", "tidal_range", "Tidal Range (m)"),
        ("lap", "formzahl", "Bilangan Formzahl"),
        ("lap", "tipe_pasut", "Tipe Pasang Surut"),
    ]),
    ("Sosial Ekonomi", [
        ("prop", "desa_luas_ha", "Luas Desa (Ha)"),
        ("prop", "desa_penduduk", "Jumlah Penduduk Desa (jiwa)"),
    ]),
]


def form_field_name(source, key):
    """Nama unik untuk dipakai sebagai atribut 'name' pada <input> HTML."""
    safe_key = key.replace(" ", "_").replace("/", "_")
    return f"{source}__{safe_key}"


def get_value(source, key, prop_data, lap_data):
    if source == "prop":
        return prop_data.get(key, "") or ""
    if source == "lap":
        return lap_data.get(key, "") or ""
    if source == "prop_loc":
        parts = prop_data.get("_lokasi_parts") or []
        idx = int(key)
        return parts[idx] if idx < len(parts) else ""
    return ""


def apply_form_values(form, prop_data, lap_data):
    """Terapkan nilai dari request.form (hasil edit pengguna) kembali ke
    prop_data / lap_data. Dipanggil saat /finalize."""
    lokasi_parts = list(prop_data.get("_lokasi_parts") or ["", "", "", ""])
    while len(lokasi_parts) < 4:
        lokasi_parts.append("")

    for _, fields in FIELD_GROUPS:
        for source, key, _label in fields:
            fname = form_field_name(source, key)
            if fname not in form:
                continue
            val = form.get(fname, "").strip()
            if source == "prop":
                prop_data[key] = val
            elif source == "lap":
                lap_data[key] = val
            elif source == "prop_loc":
                idx = int(key)
                lokasi_parts[idx] = val

    prop_data["_lokasi_parts"] = lokasi_parts
    return prop_data, lap_data
