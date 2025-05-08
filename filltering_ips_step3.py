import pandas as pd
import re
import os

# Load data dari scraping sebelumnya
df = pd.read_excel("excel/data_prodi_ptn_all.xlsx")

# Pastikan nama jurusan huruf kapital
df["NAMA"] = df["NAMA"].astype(str).str.upper()

# Daftar kata kunci jurusan IPS
keyword_ips = [
    "MANAJEMEN", "AKUNTANSI", "ADMINISTRASI", "EKONOMI", "HUKUM",
    "KOMUNIKASI", "SOSIOLOGI", "BISNIS", "POLITIK", "PARIWISATA",
    "HUBUNGAN INTERNASIONAL", "PSIKOLOGI", "SEJARAH", "ANTROPOLOGI",
    "PERPUSTAKAAN", "KESEJAHTERAAN SOSIAL", "PEMERINTAHAN", "FILSAFAT",
    "ARKEOLOGI", "JEPANG", "INDONESIA", "TATA BOGA", "SENI", "BAHASA",
    "SASTRA", "MUSIK", "TARI", "KEWIRAUSAHAAN", "BIMBINGAN", "KONSELING"
]

# Gabungkan kata kunci menjadi pola regex (misalnya: 'JEPANG|SASTRA|EKONOMI|...')
regex_pattern = "|".join([re.escape(k) for k in keyword_ips])

# Filter jurusan berdasarkan pencocokan kata kunci
df_ips = df[df["NAMA"].str.contains(regex_pattern, na=False)]

# Kelompokkan berdasarkan PTN
grouped = df_ips.groupby(["KODE_PTN", "NAMA_PTN"])

# Fungsi untuk membersihkan nama sheet Excel
def sanitize_sheet_name(name):
    name = re.sub(r"[:\\/?*\[\]]", "", name)
    return name.strip()[:31]

# Simpan hasil ke Excel multi-sheet

os.makedirs("excel", exist_ok=True)

# Simpan file ke dalam folder "excel"
output_file = "excel/daftar_jurusan_ips_per_ptn_keywords.xlsx"

with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
    sheet_names = set()
    for (kode, nama), group in grouped:
        safe_name = sanitize_sheet_name(nama)
        if safe_name in sheet_names:
            safe_name = f"{safe_name[:28]}_{kode}"
        sheet_names.add(safe_name)
        group.to_excel(writer, sheet_name=safe_name, index=False)

print(f"✅ Data jurusan IPS (berbasis kata kunci) berhasil disimpan ke: {output_file}")