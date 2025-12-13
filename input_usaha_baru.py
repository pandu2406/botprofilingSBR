from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
import time
import pandas as pd

# ====== Setup Chrome Driver ======
chrome_options = Options()
chrome_options.add_argument("--disable-logging")
chrome_options.add_argument("--log-level=3")
chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'])

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
driver.get("https://matchapro.web.bps.go.id")
wait = WebDriverWait(driver, 30)

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

# === Tunggu sampai masuk ke halaman utama MatchaPro ===
try:
    wait.until(EC.presence_of_element_located((By.XPATH, "//span[normalize-space()='Tambah Usaha']")))
    print("✅ Login SSO berhasil, sudah masuk ke halaman utama.")
except Exception as e:
    print(f"❌ Gagal memuat halaman utama setelah login: {e}")
    driver.quit()
    raise

# ====== Klik tombol navbar (garis tiga) agar menu kiri muncul ======
try:
    print("📂 Mencoba membuka sidebar navbar...")
    
    # Tunggu tombol muncul
    navbar_toggle = WebDriverWait(driver, 15).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, "a.nav-link.menu-toggle"))
    )
    
    # Scroll biar tombol kelihatan di viewport
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", navbar_toggle)
    time.sleep(0.5)
    
    # Kadang klik biasa gagal karena overlap → gunakan JavaScript click
    driver.execute_script("arguments[0].click();", navbar_toggle)
    
    print("✅ Sidebar navbar berhasil dibuka.")
    time.sleep(2)
except Exception as e:
    print(f"⚠️ Gagal klik tombol navbar: {e}")

# ====== Baca file Excel ======
file_path = "H:/Other computers/Mac Saya/BPS Tanjung Jabung Barat/2025/koding/SBR/data/hasil_bersih_DPMPTSP.xlsx"
df = pd.read_excel(file_path, sheet_name="Sheet1", dtype=str)
# df['Nomor Telp'] = df['Nomor Telp'].astype(str).str[1:]

# ====== Bersihkan kolom Nomor Telp ======
def bersihkan_nomor_telp(nomor):
    if pd.isna(nomor):
        return ""
    nomor = str(nomor).strip()
    if nomor.startswith("+62"):
        return nomor.replace("+62", "", 1).strip()
    elif nomor.startswith("62"):
        return nomor.replace("62", "", 1).strip()
    elif nomor.startswith("0"):
        return nomor.replace("0", "", 1).strip()
    else:
        ""

df['Nomor Telp'] = df['Nomor Telp'].apply(bersihkan_nomor_telp)

print(f"📘 Ditemukan {len(df)} baris data dari file Excel")

# ====== Fungsi bantu klik dropdown berdasarkan teks ======
from selenium.common.exceptions import ElementClickInterceptedException, TimeoutException
from selenium.webdriver import ActionChains

def pilih_dropdown_select2(driver, select_id, teks_dicari, timeout=12, max_retries=2):
    """
    Lebih robust: klik tampilan Select2, tunggu opsi muncul, cari opsi yang mengandung teks_dicari (case-insensitive).
    Mengatasi ElementClickInterceptedException dengan beberapa fallback (move_to_element, JS click, close overlays).
    """
    teks_dicari = str(teks_dicari).strip().upper()
    xpath_span = f"//select[@id='{select_id}']/ancestor::div[contains(@class,'position-relative')][1]//span[contains(@class,'select2-selection')]"

    for attempt in range(1, max_retries+1):
        try:
            dropdown = WebDriverWait(driver, timeout).until(
                EC.element_to_be_clickable((By.XPATH, xpath_span))
            )
            # scroll ke tombol
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", dropdown)
            time.sleep(0.15)

            # 1) coba klik biasa
            try:
                dropdown.click()
            except ElementClickInterceptedException:
                # 2) coba ActionChains move + click
                try:
                    ActionChains(driver).move_to_element(dropdown).pause(0.1).click().perform()
                except Exception:
                    # 3) fallback: klik via JavaScript
                    driver.execute_script("arguments[0].click();", dropdown)

            time.sleep(0.35)

            # tunggu opsi muncul (list Select2)
            opsi = WebDriverWait(driver, timeout).until(
                EC.presence_of_all_elements_located(
                    (By.XPATH, "//li[contains(@class,'select2-results__option') and not(contains(@class,'loading-results'))]")
                )
            )

            # cari match
            for o in opsi:
                teks = o.text.strip().upper()
                if teks_dicari in teks:
                    try:
                        driver.execute_script("arguments[0].scrollIntoView(true);", o)
                        o.click()
                    except ElementClickInterceptedException:
                        driver.execute_script("arguments[0].click();", o)
                    print(f"✅ {select_id} dipilih: {teks}")
                    return True

            # jika tidak ditemukan → log dan return False
            print(f"⚠️ '{teks_dicari}' tidak ditemukan dalam dropdown {select_id}.")
            return False

        except (ElementClickInterceptedException, TimeoutException) as e:
            print(f"⚠️ Attempt {attempt}/{max_retries} — masalah klik atau timeout saat membuka {select_id}: {e}")

            # coba tutup potensi overlay/toast yang menutupi (umum di UI bootstrap)
            try:
                # contoh selector overlay umum; sesuaikan bila perlu
                overlays = driver.find_elements(By.CSS_SELECTOR, ".swal2-container, .modal-backdrop, .toast, .iziToast, .loader, .spinner")
                for ov in overlays:
                    try:
                        driver.execute_script("arguments[0].style.display='none';", ov)
                    except:
                        pass
                time.sleep(0.3)
            except:
                pass

            # beri jeda dan ulangi
            time.sleep(0.8)
            continue

    # semua retry gagal
    print(f"❌ Gagal buka/klik dropdown {select_id} setelah {max_retries} percobaan.")
    return False

# ====== Fungsi bantu klik aman ======
def safe_click(by, value, timeout=10):
    element = WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((by, value)))
    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", element)
    time.sleep(0.5)
    element.click()

# ====== Loop tiap baris Excel ======
for idx, row in df.iterrows():
    # ====== Klik menu Tambah Usaha > Input Form ======
    try:
        # pastikan menu sidebar terlihat
        driver.execute_script("window.scrollTo(0, 0);")
        time.sleep(1)
                
        # klik menu induk "Tambah Usaha" (kalau belum terbuka)
        tambah_usaha_menu = wait.until(EC.presence_of_element_located(
            (By.XPATH, "//span[contains(@class,'menu-title') and normalize-space()='Tambah Usaha']")
        ))
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", tambah_usaha_menu)
        time.sleep(0.5)
        tambah_usaha_menu.click()
        print("📂 Menu 'Tambah Usaha' diklik.")
        time.sleep(1)

        # klik submenu "Input Form"
        input_form_link = wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//a[@href='https://matchapro.web.bps.go.id/profiling/create/usaha/input-form']")
        ))
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", input_form_link)
        time.sleep(0.5)
        input_form_link.click()
        print("🟢 Submenu 'Input Form' diklik.")
        time.sleep(4)

    except Exception as e:
        print(f"❌ Gagal membuka menu Input Form: {e}")
        
    try:
        nama_usaha = str(row["Nama Perusahaan"]).strip()
        alamat_usaha = str(row["Alamat Usaha"]).strip()
        kecamatan = str(row["Kecamatan"]).strip()
        kelurahan = str(row["Kelurahan"]).strip()
        whatsapp = str(row["Nomor Telp"]).strip()
        email = str(row["Email"]).strip()
        sumber_profiling = str(row['sumber_profiling']).strip()
        catatan_profiling = str(row['catatan_profiling']).strip()

        print(f"\n🟢 Memproses baris {idx+1}: {nama_usaha}")

        # ====== Isi Form ======
        try:
            print("📝 Mengisi form input usaha...")

            # === Nama Usaha ===
            nama_usaha = row["Nama Perusahaan"]
            nama_field = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "nama_usaha"))
            )
            nama_field.clear()
            nama_field.send_keys(nama_usaha)
            print(f"✅ Nama usaha diisi: {nama_usaha}")

            # Jika karakter kurang dari 10, tambahkan teks pelengkap
            if len(alamat_usaha) < 10:
                alamat_usaha += " (alamat lengkap belum tersedia)"
                print(f"⚠️ Alamat terlalu pendek, ditambahkan teks pelengkap: {alamat_usaha}")

            alamat_field = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "alamat"))
            )
            alamat_field.clear()
            alamat_field.send_keys(alamat_usaha)
            print(f"✅ Alamat usaha diisi: {alamat_usaha}")

            pilih_dropdown_select2(driver, "select2-provinsi", "[15] JAMBI")
            # tunggu kabupaten terload
            WebDriverWait(driver, 12).until(
                EC.presence_of_all_elements_located((By.XPATH, "//select[@id='select2-kabupaten_kota']/option"))
            )
            pilih_dropdown_select2(driver, "select2-kabupaten_kota", "TANJUNG JABUNG BARAT")
            time.sleep(0.8)
            pilih_dropdown_select2(driver, "select2-kecamatan", row["Kecamatan"])
            time.sleep(0.8)
            pilih_dropdown_select2(driver, "select2-kelurahan_desa", row["Kelurahan"])

            # Klik Next
            next_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[normalize-space()='Next']")))
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", next_btn)
            time.sleep(0.5)
            next_btn.click()

            # Centang checkbox
            checkbox = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@type='checkbox']")))
            driver.execute_script("arguments[0].click();", checkbox)
            print("☑️ Checklist konfirmasi ditandai.")

            # Klik Submit
            submit_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[normalize-space()='Submit']")))
            driver.execute_script("arguments[0].click();", submit_btn)
            print("🚀 Klik Submit berhasil.")

            # Tunggu transisi/animasi halaman setelah submit
            time.sleep(2)

            # === Nomor WhatsApp ===
            for attempt in range(3):  # coba sampai 3 kali
                try:
                    whatsapp_value = str(row["Nomor Whatsapp"]).strip() if pd.notna(row["Nomor Whatsapp"]) else ""
                    
                    # Tunggu sampai input benar-benar bisa diklik dan diketik
                    whatsapp_input = WebDriverWait(driver, 20).until(
                        EC.element_to_be_clickable((By.NAME, "whatsapp"))
                    )
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", whatsapp_input)
                    time.sleep(0.5)
                    whatsapp_input.clear()

                    if len(whatsapp_value) > 0:
                        whatsapp_input.send_keys(whatsapp_value)
                        print(f"✅ Nomor WhatsApp diisi: {whatsapp_value}")
                    else:
                        print("⚠️ Kolom WhatsApp kosong, dilewati.")
                    break  # keluar dari loop jika berhasil
                except Exception as e:
                    print(f"⏳ Gagal isi WhatsApp (percobaan {attempt+1}/3): {e}")
                    time.sleep(3)  # beri jeda sebelum mencoba ulang
            else:
                print("❌ Gagal mengisi WhatsApp setelah 3 percobaan.")

            # === Email ===
            try:
                email_value = str(row["Email"]).strip() if pd.notna(row["Email"]) else ""

                # Tunggu elemen checkbox dan email muncul
                checkbox = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.ID, "check-email"))
                )
                email_input = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.ID, "email"))
                )

                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", email_input)
                time.sleep(0.5)

                # Bersihkan dulu field email
                email_input.clear()

                try:
                    # Tunggu popup SweetAlert2 muncul
                    ok_button = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, "button.swal2-confirm"))
                    )

                    # Scroll ke tombol OK jika perlu
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", ok_button)
                    time.sleep(0.5)

                    # Klik tombol OK
                    ok_button.click()
                    print("✅ Tombol 'OK' pada popup berhasil diklik.")

                    # Tunggu popup benar-benar hilang sebelum lanjut
                    WebDriverWait(driver, 10).until(
                        EC.invisibility_of_element_located((By.CLASS_NAME, "swal2-popup"))
                    )

                except Exception as e:
                    print(f"⚠️ Tidak ada popup atau gagal klik tombol 'OK': {e}")

                # Jika email ada dan valid, pastikan checkbox dalam keadaan dicentang
                if len(email_value) > 5 and "@" in email_value:
                    if not checkbox.is_selected():
                        checkbox.click()
                    email_input.send_keys(email_value)
                    print(f"✅ Email diisi: {email_value}")

                # Jika email kosong → uncentang checkbox supaya tidak muncul error
                else:
                    if checkbox.is_selected():
                        checkbox.click()
                    print("⚠️ Email kosong atau tidak valid → checkbox di-uncheck.")

            except Exception as e:
                print(f"❌ Gagal mengisi email: {e}")

            sumber_input = driver.find_element(By.NAME, "sumber_profiling")
            sumber_input.clear()
            sumber_input.send_keys(sumber_profiling)

            catatan_input = driver.find_element(By.ID, "catatan_profiling")
            catatan_input.clear()
            catatan_input.send_keys(catatan_profiling)

            # 🔄 Cek submit-final lagi sebelum submit
            try:
                driver.find_element(By.ID, "submit-final")
            except:
                print(f"⚠ Submit-final tidak tersedia saat submit, IDSBR: {idsbr}")
                gagal_list.append(idsbr)
                driver.get("https://matchapro.web.bps.go.id/direktori-usaha")
                wait.until(EC.presence_of_element_located((By.NAME, "idsbr")))
                close_locked_popup()
                continue

            # 10. Klik Submit Final
            safe_click(By.ID, "submit-final")

            # 11. Jika muncul dialog "Ignore", klik ignore
            try:
                ignore_btn = WebDriverWait(driver, 3).until(
                    EC.element_to_be_clickable((By.ID, "confirm-consistency"))
                )
                ignore_btn.click()
            except:
                pass

            # 12. Klik "Ya, Submit!"
            safe_click(By.CSS_SELECTOR, 'button.swal2-confirm.btn.btn-primary')

            # 13. Klik tombol OK setelah Submit
            try:
                ok_btn = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, 'button.swal2-confirm.btn-success'))
                )
                ok_btn.click()
            except:
                pass

            print ("✅ Berhasil update informasi terkait usaha/perusahaan")

            # 14. Kembali ke halaman Input Form
            try:
                print("↩️ Kembali ke menu Tambah Usaha untuk input berikutnya...")
                driver.get("https://matchapro.web.bps.go.id/profiling/create/usaha/input-form")
                WebDriverWait(driver, 20).until(
                    EC.presence_of_element_located((By.ID, "nama_usaha"))
                )
                print("✅ Halaman Input Form siap untuk baris berikutnya.")
                time.sleep(2)
            except Exception as e:
                print(f"⚠️ Gagal kembali ke halaman Input Form: {e}")

        except Exception as e:
            print(f"❌ Gagal mengisi form: {e}")

        except Exception as e:
            print(f"❌ Gagal membuka menu Input Form: {e}")

    except Exception as e:
        print(f"❌ Gagal input baris {idx+1}: {e}")
        continue

print("🎯 Semua data telah diproses.")