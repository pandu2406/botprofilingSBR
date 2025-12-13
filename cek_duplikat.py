import pandas as pd
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity
from sklearn.ensemble import RandomForestClassifier
from sklearn.model_selection import train_test_split

# === 1. Baca file Excel ===
df_a = pd.read_excel("data/hasil_duplikat_DPMPTSP.xlsx", dtype=str)
df_b = pd.read_excel("data/direktori_usaha_2025-08-19_145350.xlsx", sheet_name="Sheet1 (2)", dtype=str)

kolom_a = "Nama Perusahaan"
kolom_b = "Nama Usaha"

# Hapus baris kosong
df_a = df_a.dropna(subset=[kolom_a])
df_b = df_b.dropna(subset=[kolom_b])

# === 2. Siapkan TF-IDF ===
vectorizer = TfidfVectorizer(analyzer="char_wb", ngram_range=(2,4))
vectorizer.fit(pd.concat([df_a[kolom_a], df_b[kolom_b]]))

# === 3. Fungsi bantu untuk hitung kemiripan antar batch ===
def hitung_kemiripan(batch_a, df_b, threshold=0.85):
    tfidf_a = vectorizer.transform(batch_a)
    tfidf_b = vectorizer.transform(df_b[kolom_b])
    sim = cosine_similarity(tfidf_a, tfidf_b)

    hasil = []
    for i, nama_a in enumerate(batch_a):
        for j, nama_b in enumerate(df_b[kolom_b]):
            skor = sim[i, j]
            if skor > threshold:  # hanya simpan pasangan mirip
                hasil.append({
                    "nama_a": nama_a,
                    "nama_b": nama_b,
                    "similarity": skor
                })
    return pd.DataFrame(hasil)

# === 4. Proses bertahap (hemat memori) ===
batch_size = 500  # bisa ubah jadi 1000 kalau RAM kuat
hasil_semua = []

for start in range(0, len(df_a), batch_size):
    end = min(start + batch_size, len(df_a))
    batch = df_a[kolom_a].iloc[start:end]
    print(f"🔹 Memproses batch {start}–{end} ...")
    hasil_batch = hitung_kemiripan(batch, df_b)
    hasil_semua.append(hasil_batch)

pairs = pd.concat(hasil_semua, ignore_index=True)

# === 5. Label otomatis untuk latihan Random Forest ===
pairs["label"] = (pairs["similarity"] > 0.9).astype(int)

# === 6. Train Random Forest ===
X = pairs[["similarity"]]
y = pairs["label"]

X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2, random_state=42)
rf = RandomForestClassifier(n_estimators=100, random_state=42)
rf.fit(X_train, y_train)

print("Akurasi:", rf.score(X_test, y_test))

# === 7. Prediksi akhir ===
pairs["prediksi"] = rf.predict(X)

# === 8. Simpan hasil ===
pairs.to_excel("data/hasil_duplikat_rf_batch2.xlsx", index=False)
print("✅ Selesai! Hasil tersimpan di 'hasil_duplikat_rf_batch2.xlsx'")

# === 7. Hapus data yang terduplikat dari df_a ===
# Ambil nama yang muncul di hasil duplikat
nama_duplikat = set(pairs["nama_a"].unique())

# Filter df_a untuk menyisakan yang tidak ada di daftar duplikat
df_a_bersih = df_a[~df_a[kolom_a].isin(nama_duplikat)]

# Simpan hasil bersih ke file baru
df_a_bersih.to_excel("data/hasil_bersih_DPMPTSP.xlsx", index=False)
print(f"🧹 File hasil bersih disimpan di 'data/hasil_bersih_DPMPTSP.xlsx' — total {len(df_a_bersih)} baris tersisa.")


# # === 2. Load model BERT ringan untuk similarity ===
# model = SentenceTransformer("all-MiniLM-L6-v2")

# # === 3. Buat embedding ===
# emb_a = model.encode(df_a[kolom_a].astype(str).tolist(), convert_to_tensor=True)
# emb_b = model.encode(df_b[kolom_b].astype(str).tolist(), convert_to_tensor=True)

# # === 4. Hitung cosine similarity matrix ===
# cosine_scores = util.cos_sim(emb_a, emb_b)

# # === 5. Ambil pasangan terbaik untuk setiap item di A ===
# results = []
# for i, nama_a in enumerate(df_a[kolom_a]):
#     best_idx = cosine_scores[i].argmax().item()
#     best_score = cosine_scores[i][best_idx].item()
    
#     results.append({
#         "nama_a": nama_a,
#         "match_b": df_b[kolom_b].iloc[best_idx],
#         "similarity": round(best_score, 4)
#     })

# hasil = pd.DataFrame(results)

# # === 6. Simpan hasil ke Excel ===
# hasil.to_excel("hasil_kemiripan.xlsx", index=False)

# print("✅ Selesai! Hasil kemiripan disimpan di 'hasil_kemiripan.xlsx'")
