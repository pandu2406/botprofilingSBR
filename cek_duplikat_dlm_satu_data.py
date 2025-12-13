import pandas as pd

df = pd.read_excel("data/data DPMPTSP.xlsx")

df['Nama Perusahaan'] = df['Nama Perusahaan'].str.strip()

# Jika ingin hanya menghapus duplikat tapi tetap simpan index pertama
df = df.drop_duplicates(subset='Nama Perusahaan', keep='first')

df.to_excel("data/hasil_duplikat_DPMPTSP.xlsx", index=False)

print("✅ Selesai! Hasil clean duplikat tersimpan di 'hasil_duplikat_DPMPTSP.xlsx'")