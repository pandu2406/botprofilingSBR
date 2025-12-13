import pandas as pd

# Baca file Excel
# df = pd.read_excel("/Users/hiden/Downloads/BPS Tanjung Jabung Barat/2025/koding/SBR/data/hasil_bersih_DPMPTSP.xlsx", sheet_name="Sheet1 (2)", dtype=str)
df = pd.read_excel("/Users/hiden/Downloads/BPS Tanjung Jabung Barat/2025/koding/SBR/data/template_upload_tambah_usaha_sbr.xlsx")

# # Cek ringkasan kolom 'Kelurahan'
# print("=== Ringkasan Kolom 'Kelurahan' ===")
# print(df["Kelurahan"].describe())

# # Lihat nilai unik
# print("\n=== Nilai unik dalam kolom 'Kelurahan' ===")
# print(df["Kelurahan"].unique())

# # Atau tampilkan dengan jumlah kemunculannya
# print("\n=== Jumlah tiap nilai (frekuensi) ===")
# print(df["Kelurahan"].value_counts())

# # === Mapping nama kecamatan ke format form ===
# mapping_kecamatan = {
#     "Tungkal Ulu": "[010] TUNGKAL ULU",
#     "Merlung": "[011] MERLUNG",
#     "Batang Asam": "[012] BATANG ASAM",
#     "Tebing Tinggi": "[013] TEBING TINGGI",
#     "Renah Mendaluh": "[014] RENAH MENDALUH",
#     "Muara Papalik": "[015] MUARA PAPALIK",
#     "Pengabuan": "[020] PENGABUAN",
#     "Senyerang": "[021] SENYERANG",
#     "Tungkal Ilir": "[030] TUNGKAL ILIR",
#     "Bram Itam": "[031] BRAM ITAM",
#     "Seberang Kota": "[032] SEBERANG KOTA",
#     "Betara": "[040] BETARA",
#     "Kuala Betara": "[041] KUALA BETARA"
# }

# # === Mapping nama kecamatan ke format form ===
# mapping_kecamatan = {
#     "Tungkal Ulu": "010",
#     "Merlung": "011",
#     "Batang Asam": "012",
#     "Tebing Tinggi": "013",
#     "Renah Mendaluh": "014",
#     "Muara Papalik": "015",
#     "Pengabuan": "020",
#     "Senyerang": "021",
#     "Tungkal Ilir": "030",
#     "Bram Itam": "031",
#     "Seberang Kota": "032",
#     "Betara": "040",
#     "Kuala Betara": "041"
# }

# # === Normalisasi dan ubah sesuai mapping ===
# df["Kecamatan"] = (
#     df["Kecamatan"]
#     .astype(str)
#     .str.strip()
#     .map(mapping_kecamatan)
#     .fillna(df["Kecamatan"])  # kalau tidak ada di mapping, biarkan
# )

# # === Mapping nama desa ke format ber-kode ===
# desa_mapping = {
#     "Purwodadi": "[001] PURWODADI",
#     "Sungai Kepayang": "[009] SUNGAI KEPAYANG",
#     "Adi Jaya": "[003] ADI JAYA",
#     "Taman Raja": "[026] TAMAN RAJA",
#     "Bram Itam Kiri": "[002] BRAM ITAM KIRI",
#     "Kuala Dasal": "[018] KUALA DASAL",
#     "Kampung Nelayan": "[018] KAMPUNG NELAYAN",
#     "Sungai Serindit": "[006] SUNGAI SERINDIT",
#     "Lubuk Kambing": "[001] LUBUK KAMBING",
#     "Rantau Benar": "[006] RANTAU BENAR",
#     "Teluk Sialang": "[014] TELUK SIALANG",
#     "Serdang Jaya": "[007] SERDANG JAYA",
#     "Sungai Gebar Barat": "[005] SUNGAI GEBAR BARAT",
#     "Kempas Jaya": "[007] KEMPAS JAYA",
#     "Terjun Gajah": "[010] TERJUN GAJAH",
#     "Lubuk Lawas": "[007] LUBUK LAWAS",
#     "Dataran Pinang": "[008] DATARAN PINANG",
#     "Kemuning": "[006] KEMUNING",
#     "Sungai Landak": "[010] SUNGAI LANDAK",
#     "Bram Itam Raya": "[008] BRAM ITAM RAYA",
#     "Tebing Tinggi": "[004] TEBING TINGGI",
#     "Lubuk Bernai": "[001] LUBUK BERNAI",
#     "Margo Rukun": "[001] MARGO RUKUN",
#     "Tanjung Pasir": "[007] TANJUNG PASIR",
#     "Bram Itam Kanan": "[001] BRAM ITAM KANAN",
#     "Lubuk Terap": "[016] LUBUK TERAP",
#     "Suak Labu": "[010] SUAK LABU",
#     "Tungkal III": "[009] TUNGKAL III",
#     "Talang Makmur": "[009] TALANG MAKMUR",
#     "Tanjung Tayas": "[017] TANJUNG TAYAS",
#     "Pematang Pauh": "[019] PEMATANG PAUH",
#     "Teluk Kulbi": "[014] TELUK KULBI",
#     "Patunas": "[017] PATUNAS",
#     "Merlung": "[014] MERLUNG",
#     "Tungkal II": "[010] TUNGKAL II",
#     "Muntialo": "[013] MUNTIALO",
#     "Sungai Jering": "[014] SUNGAI JERING",
#     "Sungai Terap": "[016] SUNGAI TERAP",
#     "Sungai Dualap": "[002] SUNGAI DUALAP",
#     "Sungai Baung": "[016] SUNGAI BAUNG",
#     "Karya Maju": "[017] KARYA MAJU",
#     "Sungai Pampang": "[013] SUNGAI PAMPANG",
#     "Sungai Muluk": "[007] SUNGAI MULUK",
#     "Pematang Buluh": "[012] PEMATANG BULUH",
#     "Mekar Tanjung": "[009] MEKAR TANJUNG",
#     "Mekar Jati": "[008] MEKAR JATI",
#     "Cinta Damai": "[003] CINTA DAMAI",
#     "Intan Jaya": "[001] INTAN JAYA",
#     "Pematang Lumut": "[001] PEMATANG LUMUT",
#     "Semau": "[010] SEMAU",
#     "Rantau Badak Lamo": "[010] RANTAU BADAK LAMO",
#     "Kemang Manis": "[003] KEMANG MANIS",
#     "Dusun Mudo": "[004] DUSUN MUDO",
#     "Bukit Indah": "[002] BUKIT INDAH",
#     "Sungai Dungun": "[006] SUNGAI DUNGUN",
#     "Tanjung Makmur": "[020] TANJUNG MAKMUR",
#     "Sungai Papauh": "[006] SUNGAI PAPAUH",
#     "Parit Pudin": "[007] PARIT PUDIN",
#     "Badang Sepakat": "[037] BADANG SEPAKAT",
#     "Pembengis": "[003] PEMBENGIS",
#     "Sungai Gebar": "[001] SUNGAI GEBAR",
#     "Jati Emas": "[005] JATI EMAS",
#     "Parit Sidang": "[012] PARIT SIDANG",
#     "Pantai Gading": "[007] PANTAI GADING",
#     "Suak Samin": "[011] SUAK SAMIN",
#     "Parit Bilal": "[010] PARIT BILAL",
#     "Sungai Keruh": "[006] SUNGAI KERUH",
#     "Sungsang": "[008] SUNGSANG",
#     "Lubuk Sebontan": "[009] LUBUK SEBONTAN",
#     "Sungai Raya": "[015] SUNGAI RAYA",
#     "Pasar Senin": "[018] PASAR SENIN",
#     "Pematang Tembesu": "[039] PEMATANG TEMBESU",
#     "Tanjung Senjulang": "[004] TANJUNG SENJULANG",
#     "Pematang Balam": "[008] PEMATANG BALAM",
#     "Tanjung Bojo": "[003] TANJUNG BOJO",
#     "Sungai Badar": "[008] SUNGAI BADAR",
#     "Mekar Jaya": "[009] MEKAR JAYA",
#     "Teluknilau": "[005] TELUK NILAU",
#     "Kampung Baru": "[002] KAMPUNG BARU",
#     "Bunga Tanjung": "[017] BUNGA TANJUNG",
#     "Dataran Kempas": "[007] DATARAN KEMPAS",
#     "Brasau": "[036] BRASAU",
#     "Betara Kanan": "[004] BETARA KANAN",
#     "Betara Kiri": "[003] BETARA KIRI",
#     "Sungainibung": "[015] SUNGAI NIBUNG",
#     "Suban": "[005] SUBAN",
#     "Dusun Kebun": "[004] DUSUN KEBON",
#     "Tanjung Benanak": "[005] TANJUNG BENANAK",
#     "Delima": "[008] DELIMA",
#     "Makmur Jaya": "[008] MAKMUR JAYA",
#     "Tungkal I": "[011] TUNGKAL I",
#     "Pulau Pauh": "[005] PULAU PAUH",
#     "Pinang Gading": "[010] PINANG GADING",
#     "Sriwijaya": "[016] SRIWIJAYA",
#     "Rantau Badak": "[005] RANTAU BADAK",
#     "Senyerang": "[005] SENYERANG",
#     "Pelabuhan Dagang": "[025] PELABUHAN DAGANG",
#     "Teluk Pengkah": "[010] TELUK PENGKAH",
#     "Tungkal V": "[002] TUNGKAL V",
#     "Lampisi": "[004] LAMPISI",
#     "Adi Purwa": "[007] ADI PURWA",
#     "Sungai Penoban": "[009] SUNGAI PENOBAN",
#     "Tungkal Empat Kota": "[008] TUNGKAL IV KOTA",
#     "Teluk Pulai Raya": "[001] TELUK PULAI RAYA",
#     "Tungkal Harapan": "[007] TUNGKAL HARAPAN",
#     "Penyabungan": "[017] PENYABUNGAN",
#     "Kelagian": "[005] KELAGIAN",
#     "Sungai Paur": "[009] SUNGAI PAUR",
#     "Lumahan": "[006] LUMAHAN",
#     "Teluk Ketapang": "[003] TELUK KETAPANG",
#     "Mandala Jaya": "[015] MANDALA JAYA",
#     "Suka Damai": "[002] SUKADAMAI",
#     "Bukit Harapan": "[006] BUKIT HARAPAN",
#     "Sri Agung": "[006] SRI AGUNG",
#     "Lubuk Terentang": "[011] LUBUK TERENTANG",
#     "Tanjung Paku": "[015] TANJUNG PAKU",
#     "Kuala Indah": "[009] KUALA INDAH",
#     "Tanah Tumbuh": "[008] TANAH TUMBUH",
#     "Kuala Baru": "[004] KUALA BARU",
#     "Sungai Rambai": "[002] SUNGAI RAMBAI",
#     "Kuala Kahar": "[005] KUALA KAHAR",
#     "Mekar Alam": "[007] MEKAR ALAM",
#     "Gemuruh": "[038] GEMURUH",
#     "Sungai Kayu Aro": "[004] SUNGAI KAYU ARO",
#     "Rawa Medang": "[011] RAWA MEDANG",
#     "Muara Danau": "[007] MUARA DANAU",
#     "Bukit Bakar": "[010] BUKIT BAKAR",
#     "Badang": "[016] BADANG",
#     "Sungai Rotan": "[002] SUNGAI ROTAN",
#     "Rawang Kempas": "[010] RAWANG KEMPAS",
#     "Tungkal IV Desa": "[003] TUNGKAL IV DESA",
#     "Harapan Jaya": "[006] HARAPAN JAYA",
#     "Muara Seberang": "[008] MUARA SEBERANG",
# }

# # === Mapping nama desa ke format ber-kode ===
# desa_mapping = {
#     "Purwodadi": "001",
#     "Sungai Kepayang": "009",
#     "Adi Jaya": "003",
#     "Taman Raja": "026",
#     "Bram Itam Kiri": "002",
#     "Kuala Dasal": "018",
#     "Kampung Nelayan": "018",
#     "Sungai Serindit": "006",
#     "Lubuk Kambing": "001",
#     "Rantau Benar": "006",
#     "Teluk Sialang": "014",
#     "Serdang Jaya": "007",
#     "Sungai Gebar Barat": "005",
#     "Kempas Jaya": "007",
#     "Terjun Gajah": "010",
#     "Lubuk Lawas": "007",
#     "Dataran Pinang": "008",
#     "Kemuning": "006",
#     "Sungai Landak": "010",
#     "Bram Itam Raya": "008",
#     "Tebing Tinggi": "004",
#     "Lubuk Bernai": "001",
#     "Margo Rukun": "001",
#     "Tanjung Pasir": "007",
#     "Bram Itam Kanan": "001",
#     "Lubuk Terap": "016",
#     "Suak Labu": "010",
#     "Tungkal III": "009",
#     "Talang Makmur": "009",
#     "Tanjung Tayas": "017",
#     "Pematang Pauh": "019",
#     "Teluk Kulbi": "014",
#     "Patunas": "017",
#     "Merlung": "014",
#     "Tungkal II": "010",
#     "Muntialo": "013",
#     "Sungai Jering": "014",
#     "Sungai Terap": "016",
#     "Sungai Dualap": "002",
#     "Sungai Baung": "016",
#     "Karya Maju": "017",
#     "Sungai Pampang": "013",
#     "Sungai Muluk": "007",
#     "Pematang Buluh": "012",
#     "Mekar Tanjung": "009",
#     "Mekar Jati": "008",
#     "Cinta Damai": "003",
#     "Intan Jaya": "001",
#     "Pematang Lumut": "001",
#     "Semau": "010",
#     "Rantau Badak Lamo": "010",
#     "Kemang Manis": "003",
#     "Dusun Mudo": "004",
#     "Bukit Indah": "002",
#     "Sungai Dungun": "006",
#     "Tanjung Makmur": "020",
#     "Sungai Papauh": "006",
#     "Parit Pudin": "007",
#     "Badang Sepakat": "037",
#     "Pembengis": "003",
#     "Sungai Gebar": "001",
#     "Jati Emas": "005",
#     "Parit Sidang": "012",
#     "Pantai Gading": "007",
#     "Suak Samin": "011",
#     "Parit Bilal": "010",
#     "Sungai Keruh": "006",
#     "Sungsang": "008",
#     "Lubuk Sebontan": "009",
#     "Sungai Raya": "015",
#     "Pasar Senin": "018",
#     "Pematang Tembesu": "039",
#     "Tanjung Senjulang": "004",
#     "Pematang Balam": "008",
#     "Tanjung Bojo": "003",
#     "Sungai Badar": "008",
#     "Mekar Jaya": "009",
#     "Teluknilau": "005",
#     "Kampung Baru": "002",
#     "Bunga Tanjung": "017",
#     "Dataran Kempas": "007",
#     "Brasau": "036",
#     "Betara Kanan": "004",
#     "Betara Kiri": "003",
#     "Sungainibung": "015",
#     "Suban": "005",
#     "Dusun Kebun": "004",
#     "Tanjung Benanak": "005",
#     "Delima": "008",
#     "Makmur Jaya": "008",
#     "Tungkal I": "011",
#     "Pulau Pauh": "005",
#     "Pinang Gading": "010",
#     "Sriwijaya": "016",
#     "Rantau Badak": "005",
#     "Senyerang": "005",
#     "Pelabuhan Dagang": "025",
#     "Teluk Pengkah": "010",
#     "Tungkal V": "002",
#     "Lampisi": "004",
#     "Adi Purwa": "007",
#     "Sungai Penoban": "009",
#     "Tungkal Empat Kota": "008",
#     "Teluk Pulai Raya": "001",
#     "Tungkal Harapan": "007",
#     "Penyabungan": "017",
#     "Kelagian": "005",
#     "Sungai Paur": "009",
#     "Lumahan": "006",
#     "Teluk Ketapang": "003",
#     "Mandala Jaya": "015",
#     "Suka Damai": "002",
#     "Bukit Harapan": "006",
#     "Sri Agung": "006",
#     "Lubuk Terentang": "011",
#     "Tanjung Paku": "015",
#     "Kuala Indah": "009",
#     "Tanah Tumbuh": "008",
#     "Kuala Baru": "004",
#     "Sungai Rambai": "002",
#     "Kuala Kahar": "005",
#     "Mekar Alam": "007",
#     "Gemuruh": "038",
#     "Sungai Kayu Aro": "004",
#     "Rawa Medang": "011",
#     "Muara Danau": "007",
#     "Bukit Bakar": "010",
#     "Badang": "016",
#     "Sungai Rotan": "002",
#     "Rawang Kempas": "010",
#     "Tungkal IV Desa": "003",
#     "Harapan Jaya": "006",
#     "Muara Seberang": "008",
# }

# # === Terapkan mapping pada kolom 'Desa' ===
# df["Kelurahan"] = df["Kelurahan"].map(desa_mapping).fillna(df["Kelurahan"])

df = df.astype(str)

# ====== Bersihkan kolom Nomor Telp ======
def bersihkan_nomor_telp(nomor):
    if pd.isna(nomor):
        return ""
    nomor = str(nomor).strip()
    if nomor.startswith("+620"):
        return nomor.replace("+620", "", 1).strip()
    elif nomor.startswith("+62"):
        return nomor.replace("+62", "", 1).strip()
    elif nomor.startswith("62"):
        return nomor.replace("62", "", 1).strip()
    elif nomor.startswith("+0"):
        return nomor.replace("+0", "", 1).strip()
    elif nomor.startswith("0"):
        return nomor.replace("0", "", 1).strip()
    elif nomor.startswith("NAKTIF_+62"):
        return nomor.replace("NAKTIF_+62", "", 1).strip()
    elif nomor.startswith("NONAKTIF_HAK_AKSES_+62"):
        return nomor.replace("NONAKTIF_HAK_AKSES_+62", "", 1).strip()
    elif nomor.startswith("NONAKTIF_HAK_AKSES_+620"):
        return nomor.replace("NONAKTIF_HAK_AKSES_+620", "", 1).strip()
    else:
        return nomor

df['nomor_whatsapp'] = df['nomor_whatsapp'].apply(bersihkan_nomor_telp)

# === Simpan kembali ke Excel ===
# df.to_excel("/Users/hiden/Downloads/BPS Tanjung Jabung Barat/2025/koding/SBR/data/hasil_bersih_DPMPTSP.xlsx", sheet_name="Sheet1 (2)", index=False)
df.to_excel("/Users/hiden/Downloads/BPS Tanjung Jabung Barat/2025/koding/SBR/data/template_upload_tambah_usaha_sbr.xlsx", index=False)
# print("berhasil disimpan ke 'hasil_bersih_DPMPTSP.xlsx'")
print("berhasil disimpan ke 'template_upload_tambah_usaha_sbr.xlsx'")
