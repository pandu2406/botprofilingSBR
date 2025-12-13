from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
import time
import pandas as pd

# ====== Load Excel Data ======
excel_file = 'direktori_usaha_2025-08-19_145350.xlsx'
df = pd.read_excel(excel_file)
df['IDSBR'] = df['IDSBR'].astype(str)

hasil = []

# ====== Setup Chrome Driver ======
chrome_options = Options()
chrome_options.add_argument("--disable-logging")
chrome_options.add_argument("--log-level=3")
chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'])

driver = webdriver.Chrome(service=Service(), options=chrome_options)
driver.get("https://matchapro.web.bps.go.id/direktori-usaha")
wait = WebDriverWait(driver, 30)

# ====== Tunggu login manual ======
input("🔐 Silakan login terlebih dahulu, lalu tekan Enter untuk melanjutkan...")

# ====== Fungsi bantu klik aman ======
def safe_click(by, value, timeout=10):
    element = WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((by, value)))
    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", element)
    time.sleep(0.5)
    element.click()

# ====== Fungsi tutup popup Locked By ======
def close_locked_popup():
    try:
        skip_btn = WebDriverWait(driver, 2).until(
            EC.element_to_be_clickable((By.XPATH, "//button[contains(text(),'Skip')]"))
        )
        skip_btn.click()
        print("ℹ️ Pop-up Locked By tertutup (klik Skip).")
    except:
        try:
            next_btn = WebDriverWait(driver, 1).until(
                EC.element_to_be_clickable((By.XPATH, "//button[contains(text(),'Next')]"))
            )
            next_btn.click()
            print("ℹ️ Pop-up Locked By tertutup (klik Next).")
        except:
            pass

# ====== Loop Tiap Baris ======
for index, row in df.iterrows():
    idsbr = row['IDSBR']
    username = ""
    status = ""
    last_update = ""

    close_locked_popup()

    try:
        close_locked_popup()
        # 1. Isi kolom IDSBR
        search_box = wait.until(EC.presence_of_element_located((By.NAME, "idsbr")))
        search_box.clear()
        search_box.send_keys(idsbr)
        time.sleep(2)

        # 2. Klik tombol filter
        safe_click(By.ID, "filter-data")
        time.sleep(3)

        # 3. Tutup popup Locked By (jika ada)
        close_locked_popup()

        # ==== CEK HASIL PENCARIAN ====
        try:
            WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.XPATH, f"//td[contains(text(), '{idsbr}')]"))
            )
        except:
            print(f"[!] {idsbr} tidak ditemukan, kemungkinan data baru.")
            username = "Baru"
            hasil.append({
                'idsbr': idsbr,
                'username': username,
                'status': status,
                'last_update': last_update
            })
            driver.get("https://matchapro.web.bps.go.id/direktori-usaha")
            time.sleep(2)
            continue

        # 4. Coba klik tombol history profiling
        try:
            history_button = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "a.btn-history-profiling"))
            )
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", history_button)
            time.sleep(1)
            history_button.click()

            # 🔒 Tutup popup kalau muncul setelah buka history
            close_locked_popup()

            # 5. Tunggu tabel history muncul
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "table-history-profiling"))
            )
            time.sleep(2)

            # 6. Ambil data dari baris pertama tabel
            try:
                row = WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "table#table-history-profiling tbody tr"))
                )
                tds = row.find_elements(By.TAG_NAME, "td")
                if len(tds) >= 4:
                    username = tds[0].text.strip()
                    status = tds[2].text.strip()
                    last_update = tds[3].text.strip()
                    print(f"[✓] {idsbr} diprofiling oleh {username}, status: {status}, update: {last_update}")
                else:
                    print(f"[!] {idsbr} diprofiling tapi kolom data tidak lengkap.")
            except:
                print(f"[!] {idsbr} diprofiling tapi tidak ditemukan baris dalam tabel.")

            # 7. Tutup dialog jika ada tombol Close
            try:
                close_button = driver.find_element(By.XPATH, "//button[contains(text(), 'Close')]")
                close_button.click()
                close_locked_popup()
            except:
                pass

        except:
            print(f"[!] {idsbr} belum diprofiling (tombol history tidak ditemukan atau gagal klik)")
            close_locked_popup()

    except Exception as e:
        print(f"[✗] Error pada {idsbr}: {e}")
        close_locked_popup()

    # 8. Simpan ke hasil
    hasil.append({
        'idsbr': idsbr,
        'username': username,
        'status': status,
        'last_update': last_update
    })

    # 9. Kembali ke halaman awal
    driver.get("https://matchapro.web.bps.go.id/direktori-usaha")
    time.sleep(2)

# ====== Simpan Hasil ke Excel ======
output_df = pd.DataFrame(hasil)
output_df.to_excel("hasil_profiling_1001-1500.xlsx", index=False)
print("✅ Semua data berhasil disimpan ke hasil_profiling.xlsx")

# ====== Tutup Browser ======
driver.quit()
