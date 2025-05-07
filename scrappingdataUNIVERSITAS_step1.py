import pandas as pd
import os
from time import sleep

# URL utama
base_url = "https://sidatagrun-public-1076756628210.asia-southeast2.run.app/ptn_sn.php"
detail_base_url = "https://sidatagrun-public-1076756628210.asia-southeast2.run.app/ptn_sn.php?ptn="

# Ambil data PTN
ptn_tables = pd.read_html(base_url)
df_ptn = ptn_tables[0]

# List untuk menyimpan semua data prodi
all_prodi_data = []

# Iterasi setiap PTN
for idx, row in df_ptn.iterrows():
    kode = str(row['KODE'])
    
    if len(kode) < 3:
        continue  # Skip jika kode tidak valid

    kode_akhir_3 = kode[-3:]
    url_detail = detail_base_url + kode_akhir_3

    try:
        print(f"Mengambil data dari: {url_detail}")
        prodi_tables = pd.read_html(url_detail)
        if prodi_tables:
            df_prodi = prodi_tables[0]
            df_prodi['KODE_PTN'] = kode
            df_prodi['NAMA_PTN'] = row['NAMA']
            all_prodi_data.append(df_prodi)

        sleep(1)  # jeda 1 detik agar tidak terlalu cepat request-nya
    except Exception as e:
        print(f"❌ Gagal mengambil data dari {url_detail}: {e}")

# Gabungkan semua data prodi jadi satu DataFrame
folder_path = "excel"
if not os.path.exists(folder_path):
    os.makedirs(folder_path)

# Gabungkan folder dan nama file
file_name = "data_prodi_ptn.xlsx"
file_path = os.path.join(folder_path, file_name)

# Simpan data ke dalam folder "excel"
if all_prodi_data:
    final_df = pd.concat(all_prodi_data, ignore_index=True)
    final_df.to_excel(file_path, index=False)
    print(f"✅ Data berhasil disimpan ke {file_path}")
else:
    print("❗ Tidak ada data prodi yang berhasil diambil.")
