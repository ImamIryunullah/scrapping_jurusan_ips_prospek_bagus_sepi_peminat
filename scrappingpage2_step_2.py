import requests
from bs4 import BeautifulSoup
import pandas as pd
from time import sleep
import os

ptn_url = "https://sidatagrun-public-1076756628210.asia-southeast2.run.app/ptn_sn.php"
detail_url_base = "https://sidatagrun-public-1076756628210.asia-southeast2.run.app/ptn_sn.php?ptn="

print("📥 Mengambil daftar PTN...")
ptn_tables = pd.read_html(ptn_url)
df_ptn = ptn_tables[0]
all_prodi = []
for idx, row in df_ptn.iterrows():
    kode_ptn = str(row['KODE']).strip()
    nama_ptn = row['NAMA'].strip()
    
    if len(kode_ptn) < 3:
        continue

    kode_url = kode_ptn[-3:]  
    detail_url = detail_url_base + kode_url

    print(f"🔍 Memproses {nama_ptn} ({kode_ptn}) → {detail_url}")
    
    try:
        res = requests.get(detail_url)
        soup = BeautifulSoup(res.content, 'html.parser')
        table = soup.find('table')
        if not table:
            print("⚠️ Tidak ada tabel ditemukan.")
            continue

        rows = table.find_all('tr')
        headers = [th.get_text(strip=True) for th in rows[0].find_all('th')]

        for r in rows[1:]:
            cols = [td.get_text(strip=True) for td in r.find_all('td')]
            if len(cols) == len(headers):
                cols.append(kode_ptn)
                cols.append(nama_ptn)
                all_prodi.append(cols)

        sleep(1)  # Hindari request terlalu cepat
    except Exception as e:
        print(f"❌ Gagal mengambil {detail_url}: {e}")
# Tambah header PTN di akhir kolom
folder_path = "excel"
if not os.path.exists(folder_path):
    os.makedirs(folder_path)

# Tambah header PTN di akhir kolom dan simpan ke file dalam folder "excel"
if all_prodi:
    final_headers = headers + ['KODE_PTN', 'NAMA_PTN']
    df_final = pd.DataFrame(all_prodi, columns=final_headers)
    
    file_path = os.path.join(folder_path, "data_prodi_ptn_all.xlsx")
    df_final.to_excel(file_path, index=False)
    
    print(f"✅ Semua data berhasil disimpan ke {file_path}")
else:
    print("❗ Tidak ada data prodi yang berhasil diambil.")
