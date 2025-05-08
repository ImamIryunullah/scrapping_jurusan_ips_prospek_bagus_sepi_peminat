import pandas as pd
import re
import os

# Pemetaan sektor BPS dan gaji rata-rata sesuai dengan data yang diberikan
keyword_to_sector_and_salary = {
    "MANAJEMEN": ("Aktivitas Keuangan dan Asuransi", 4878087),
    "AKUNTANSI": ("Aktivitas Keuangan dan Asuransi", 4878087),
    "ADMINISTRASI": ("Administrasi Pemerintahan dan Pertahanan", 3758174),
    "EKONOMI": ("Aktivitas Keuangan dan Asuransi", 4878087),
    "HUKUM": ("Administrasi Pemerintahan dan Pertahanan", 3758174),
    "KOMUNIKASI": ("Informasi dan Komunikasi", 4131648),
    "SOSIOLOGI": ("Aktivitas Kesehatan dan Kegiatan Sosial", 3415963),
    "BISNIS": ("Perdagangan Besar dan Eceran", 2667185),
    "POLITIK": ("Administrasi Pemerintahan dan Pertahanan", 3758174),
    "PARIWISATA": ("Penyediaan Akomodasi dan Makan Minum", 2424447),
    "HUBUNGAN INTERNASIONAL": ("Administrasi Pemerintahan dan Pertahanan", 3758174),
    "PSIKOLOGI": ("Aktivitas Kesehatan dan Kegiatan Sosial", 3415963),
    "SEJARAH": ("Pendidikan", 2794131),
    "ANTROPOLOGI": ("Pendidikan", 2794131),
    "PERPUSTAKAAN": ("Pendidikan", 2794131),
    "KESEJAHTERAAN SOSIAL": ("Aktivitas Kesehatan dan Kegiatan Sosial", 3415963),
    "PEMERINTAHAN": ("Administrasi Pemerintahan dan Pertahanan", 3758174),
    "FILSAFAT": ("Pendidikan", 2794131),
    "ARKEOLOGI": ("Pendidikan", 2794131),
    "JEPANG": ("Pendidikan", 2794131),
    "INDONESIA": ("Pendidikan", 2794131),
    "TATA BOGA": ("Penyediaan Akomodasi dan Makan Minum", 2424447),
    "SENI": ("Pendidikan", 2794131),
    "BAHASA": ("Pendidikan", 2794131),
    "SASTRA": ("Pendidikan", 2794131),
    "MUSIK": ("Pendidikan", 2794131),
    "TARI": ("Pendidikan", 2794131),
    "KEWIRAUSAHAAN": ("Perdagangan Besar dan Eceran", 2667185),
    "BIMBINGAN": ("Pendidikan", 2794131),
    "KONSELING": ("Pendidikan", 2794131)
}

# Fungsi untuk memetakan sektor dan gaji berdasarkan jurusan
def map_sector_and_salary(jurusan):
    for keyword, (sektor, gaji) in keyword_to_sector_and_salary.items():
        if keyword in jurusan:
            return sektor, gaji
    return "Sektor Tidak Diketahui", 0

# Load semua sheet dari file hasil sebelumnya
xls = pd.read_excel("excel/daftar_jurusan_ips_per_ptn_keywords.xlsx", sheet_name=None)  # None = semua sheet
df = pd.concat(xls.values(), ignore_index=True)  # Gabungkan semua sheet jadi satu DataFrame

# Pastikan kolom nama jurusan huruf kapital
df["NAMA"] = df["NAMA"].astype(str).str.upper()

# Pastikan kolom peminat numerik
df["PEMINAT 2024"] = pd.to_numeric(df["PEMINAT 2024"], errors="coerce")

# Filter hanya yang sepi peminat (misalnya < 100)
df_sepi = df[df["PEMINAT 2024"] < 100]

# Tambahkan kolom sektor BPS dan gaji
df_sepi[["SEKTOR_BPS", "GAJI_RATA_RATA"]] = df_sepi["NAMA"].apply(lambda x: pd.Series(map_sector_and_salary(x)))

# Menambahkan kolom "PROSPEK" berdasarkan gaji rata-rata
df_sepi["PROSPEK"] = df_sepi["GAJI_RATA_RATA"].apply(lambda x: "Prospek Bagus" if x > 3094818 else "Prospek Tidak Bagus")

# Kelompokkan lagi per PTN
grouped = df_sepi.groupby(["KODE_PTN", "NAMA_PTN"])

def sanitize_sheet_name(name):
    name = re.sub(r"[:\\/?*\[\]]", "", name)
    return name.strip()[:31]


import os

folder_path = "excel"
if not os.path.exists(folder_path):
    os.makedirs(folder_path)

# Gabungkan folder_path dengan nama file
output_file = os.path.join(folder_path, "jurusan_ips_sepi_peminat_per_ptn_dengan_sektor_bps_dan_gaji_dan_prospek.xlsx")

# Simpan hasil ke file baru
with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
    sheet_names = set()
    for (kode, nama), group in grouped:
        safe_name = sanitize_sheet_name(nama)
        if safe_name in sheet_names:
            safe_name = f"{safe_name[:28]}_{kode}"
        sheet_names.add(safe_name)
        group.to_excel(writer, sheet_name=safe_name, index=False)

print(f"✅ Data jurusan IPS sepi peminat dengan sektor BPS, gaji, dan prospek disimpan ke: {output_file}")

