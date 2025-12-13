

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
import time
import pandas as pd

# ====== Setup Chrome Driver ======
chrome_options = Options()
chrome_options.add_argument("--disable-logging")
chrome_options.add_argument("--log-level=3")
chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'])

driver = webdriver.Chrome(service=Service(), options=chrome_options)
driver.get("https://matchapro.web.bps.go.id")
wait = WebDriverWait(driver, 30)

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

# === Klik tombol "Sign in with SSO BPS" ===
try:
    sso_button = wait.until(EC.element_to_be_clickable(
        (By.XPATH, "//a[contains(., 'Sign in with SSO BPS')]")
    ))
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", sso_button)
    time.sleep(0.5)
    sso_button.click()
    print("➡️ Klik tombol 'Sign in with SSO BPS'")
except Exception as e:
    print(f"❌ Gagal klik tombol SSO: {e}")
    driver.quit()
    raise

# === Tunggu halaman SSO terbuka dan isi kredensial ===
try:
    username_input = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.ID, "username"))
    )
    password_input = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.ID, "password"))
    )
    print("🧩 Halaman SSO siap, mengisi kredensial...")

    username_input.clear()
    username_input.send_keys("") #isi username sso
    password_input.clear()
    password_input.send_keys("") #isi password sso

    login_button = wait.until(EC.element_to_be_clickable((By.ID, "kc-login")))
    login_button.click()
    print("🚀 Login SSO dikirim.")
except Exception as e:
    print(f"❌ Gagal login SSO: {e}")
    driver.quit()
    raise

# === SETELAH LOGIN SSO ===

wait = WebDriverWait(driver, 30)

# 1. Klik tombol navbar (garis tiga) agar sidebar muncul
try:
    print("📂 Membuka sidebar...")

    navbar_btn = wait.until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, "a.nav-link.menu-toggle"))
    )
    
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", navbar_btn)
    time.sleep(0.5)
    driver.execute_script("arguments[0].click();", navbar_btn)

    print("✅ Sidebar terbuka.")
    time.sleep(1)

except Exception as e:
    print(f"❌ Gagal membuka sidebar: {e}")

# 2. Klik menu “Direktori Usaha”
try:
    print("📁 Mencari menu Direktori Usaha...")

    direktori_menu = wait.until(
        EC.element_to_be_clickable((
            By.XPATH,
            "//a[contains(@href,'direktori-usaha')]"
        ))
    )

    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", direktori_menu)
    time.sleep(0.5)
    driver.execute_script("arguments[0].click();", direktori_menu)

    print("➡️ Menu Direktori Usaha diklik.")

except Exception as e:
    print(f"❌ Gagal klik menu Direktori Usaha: {e}")

# 3. Tunggu halaman selesai dimuat
try:
    wait.until(
        EC.presence_of_element_located(
            (By.XPATH, "//span[contains(., 'Direktori Usaha')]")
        )
    )
    print("✅ Berhasil membuka halaman Direktori Usaha.")

except Exception as e:
    print(f"❌ Halaman Direktori Usaha gagal dimuat: {e}")

close_locked_popup()

# ====== Set filter HISTORY PROFILING = Belum Pernah Profiling ======
locked_select = Select(wait.until(EC.presence_of_element_located((By.NAME, "filter_pernah_profiling"))))
locked_select.select_by_value("BELUM")  # value "BELUM" = Belum Pernah Profiling
time.sleep(1)

# Klik tombol filter
filter_button = wait.until(EC.element_to_be_clickable((By.ID, "filter-data")))
filter_button.click()
time.sleep(3)

# ====== Atur jumlah baris jadi 50 per halaman ======
try:
    select_length = Select(wait.until(EC.presence_of_element_located((By.NAME, "table_direktori_usaha_length"))))
    select_length.select_by_value("50")
    time.sleep(2)
    print("ℹ️ Tabel diatur menjadi 50 baris per halaman.")
except Exception as e:
    print("⚠️ Gagal set tabel ke 50 baris:", e)

hasil = []

# ====== Scrape 1 halaman ======
def scrape_current_page():
    rows = driver.find_elements(By.CSS_SELECTOR, "table#table_direktori_usaha tbody tr")
    for r in rows:
        cols = r.find_elements(By.TAG_NAME, "td")
        if len(cols) > 0:
            idsbr = cols[0].text.strip()  # kolom pertama
            nama = cols[1].text.strip()   # kolom kedua
            alamat = cols[2].text.strip() # kolom ketiga
            status = cols[4].text.strip()

            try:
                history_button = r.find_element(By.CSS_SELECTOR, "a.btn-history-profiling")
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", history_button)
                time.sleep(1)
                history_button.click()

                # kalau muncul popup locked by → tutup dulu
                close_locked_popup()

                # tunggu tabel history muncul
                WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.ID, "table-history-profiling"))
                )
                time.sleep(1)

                # ambil baris pertama history
                try:
                    row_hist = driver.find_element(By.CSS_SELECTOR, "table#table-history-profiling tbody tr")
                    tds_hist = row_hist.find_elements(By.TAG_NAME, "td")
                    if len(tds_hist) >= 4:
                        nama = tds_hist[1].text.strip()
                        alamat = tds_hist[2].text.strip()
                        status = tds_hist[4].text.strip()
                except:
                    pass

                # tutup modal history
                try:
                    close_button = driver.find_element(By.XPATH, "//button[contains(text(), 'Close')]")
                    close_button.click()
                except:
                    pass

            except:
                pass

            hasil.append({
                "IDSBR": idsbr,
                "Nama": nama,
                "Alamat": alamat,
                "Status": status
            })
            print(f"[✓] {idsbr} | {nama} | {alamat} | {status}")

# ====== LOOPING PAGINATION ======
MAX_HALAMAN = 14  # batas scraping
halaman = 1

while True:
    print(f"\n📄 Sedang scrape halaman {halaman} ...")
    scrape_current_page()

    if halaman >= MAX_HALAMAN:
        print(f"✅ Sudah sampai halaman {MAX_HALAMAN}. Scraping selesai.")
        break

    try:
        # cari tombol Next
        next_btn = wait.until(EC.element_to_be_clickable((By.ID, "table_direktori_usaha_next")))
        parent_li = next_btn.find_element(By.XPATH, "./..")

        # ambil nomor halaman aktif sekarang
        current_page = driver.find_element(By.CSS_SELECTOR, "#table_direktori_usaha_paginate li.active").text.strip()
        
        if "disabled" in parent_li.get_attribute("class"):
            print("✅ Tombol Next sudah disabled, scraping berhenti.")
            break
        else:
            driver.execute_script("arguments[0].click();", next_btn)
            print("➡️ Klik tombol Next, tunggu halaman baru...")

            # tunggu sampai halaman aktif berubah
            WebDriverWait(driver, 10).until_not(
                EC.text_to_be_present_in_element(
                    (By.CSS_SELECTOR, "#table_direktori_usaha_paginate li.active"), current_page
                )
            )
            halaman += 1
    except Exception as e:
        print("❌ Tidak bisa menemukan tombol Next:", e)
        break

# ====== Simpan hasil ======
output_df = pd.DataFrame(hasil)
output_df.to_excel("data/hasil_profiling_loop_17_Nov.xlsx", index=False)
print("✅ Semua data tersimpan ke hasil_profiling_loop_17_Nov.xlsx")
