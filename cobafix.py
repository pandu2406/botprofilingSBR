from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import TimeoutException, NoSuchElementException

import time
import pandas as pd

# ====== Load Excel Data ======
excel_file = 'data/data.xlsx'
df = pd.read_excel(excel_file, sheet_name='Sheet4')
df['idsbr'] = df['idsbr'].astype(str).str.replace(r'\.0$', '', regex=True)
df['kode_kbli'] = df['kode_kbli'].astype(str)
df['kategori_kbli'] = df['kategori_kbli'].astype(str)
df['kegiatan_profiling'] = df['kegiatan_profiling'].astype(str)
df['kepemilikan_usaha'] = df['kepemilikan_usaha'].astype(str)
df['badan_hukum'] = df['badan_hukum'].astype(str).str.replace(r'\.0$', '', regex=True)
df['sektor_institusi'] = df['sektor_institusi'].astype(str)
df['sumber_profiling'] = df['sumber_profiling'].astype(str)
df['catatan_profiling'] = df['catatan_profiling'].astype(str)

# Daftar pencatatan hasil
sukses_list = []
gagal_list = []

# ====== Setup Selenium ======
options = webdriver.ChromeOptions()

# Fullscreen langsung
options.add_argument("--start-maximized")

# (Opsional) membuka dalam fullscreen mode penuh (seperti F11)
def smooth_scroll(driver, element, step=150, delay=0.02):
    try:
        element_location = element.location_once_scrolled_into_view
        y_target = int(element_location['y'])

        # posisi sekarang (float → int)
        y_current = int(driver.execute_script("return window.pageYOffset;"))

        direction = 1 if y_target > y_current else -1
        
        # scroll bertahap
        for y in range(y_current, y_target, direction * step):
            driver.execute_script(f"window.scrollTo(0, {y});")
            time.sleep(delay)

        driver.execute_script(f"window.scrollTo(0, {y_target});")
        time.sleep(0.3)

    except Exception as e:
        print(f"⚠ Gagal scroll ke elemen: {e}")


# options.add_argument("--start-fullscreen")

driver = webdriver.Chrome(options=options)
driver.get("https://matchapro.web.bps.go.id")

# Set zoom 80%
driver.execute_script("document.body.style.zoom='80%'")

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
    username_input.send_keys("...") #isi username sso
    password_input.clear()
    password_input.send_keys("...") #isi password sso

    login_button = wait.until(EC.element_to_be_clickable((By.ID, "kc-login")))
    login_button.click()
    print("🚀 Login SSO dikirim.")
except Exception as e:
    print(f"❌ Gagal login SSO: {e}")
    driver.quit()
    raise

# === SETELAH LOGIN SSO ===

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

# ====== Fungsi bantu klik aman ======
def safe_click(by, value, timeout=10):
    element = WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((by, value)))
    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", element)
    time.sleep(0.5)
    driver.execute_script("arguments[0].click();", element)

def select2_choose_by_text(driver, select2_container_id, value):
    if value in ["", "NA", "nan", "None"]:
        return  # jangan isi jika kosong

    try:
        # klik select2
        box = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, f"#{select2_container_id} .select2-selection"))
        )
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", box)
        time.sleep(0.3)
        box.click()

        # ketik pencarian
        search_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "input.select2-search__field"))
        )
        search_input.send_keys(value)
        time.sleep(0.5)

        # pilih hasil pertama
        first_option = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "li.select2-results__option--highlighted"))
        )
        first_option.click()
        time.sleep(0.3)

        print(f"✔ Select2 '{select2_container_id}' diisi: {value}")

    except Exception as e:
        print(f"⚠ Gagal memilih Select2 {select2_container_id} dengan value '{value}': {e}")

# ====== Loop Tiap Baris ======
for index, row in df.iterrows():
    idsbr = row['idsbr']
    try:
        kode_kbli= str(row['kode_kbli']).strip()
        kategori_kbli = str(row['kategori_kbli']).strip()
        kegiatan_profiling = str(row['kegiatan_profiling']).strip()
        kepemilikan_usaha = str(row['kepemilikan_usaha']).strip()
        badan_hukum = str(row['badan_hukum']).strip()
        sektor_institusi = str(row['sektor_institusi']).strip()
        sumber_profiling = str(row['sumber_profiling']).strip()
        catatan_profiling = str(row['catatan_profiling']).strip()

        # 1. Isi kolom IDSBR
        search_box = wait.until(EC.presence_of_element_located((By.NAME, "idsbr")))
        search_box.clear()
        search_box.send_keys(idsbr)
        time.sleep(1)

        # 2. Klik tombol filter
        safe_click(By.ID, "filter-data")
        time.sleep(2)

        # 3. Klik tombol edit
        elements = driver.find_elements(By.CSS_SELECTOR, 'svg.feather.feather-edit')
        if not elements:
            print(f"❌ IDSBR tidak ditemukan: {idsbr}")
            gagal_list.append(idsbr)
            driver.get("https://matchapro.web.bps.go.id/direktori-usaha")
            wait.until(EC.presence_of_element_located((By.NAME, "idsbr")))
            close_locked_popup()
            continue

        elements[0].click()
        time.sleep(0.5)

        # 4. Klik "Ya, edit!"
        safe_click(By.CSS_SELECTOR, 'button.swal2-confirm')

        # === Berpindah ke jendela baru ===
        main_window = driver.current_window_handle
        WebDriverWait(driver, 10).until(lambda d: len(d.window_handles) > 1)
        new_windows = driver.window_handles

        for handle in new_windows:
            if handle != main_window:
                driver.close()  # tutup jendela utama
                driver.switch_to.window(handle)  # pindah ke jendela form
                break

        # 5. Jika muncul kotak dialog lain (misal warning), klik OK lalu ulangi
        try:
            ok_dialog = WebDriverWait(driver, 3).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, 'button.swal2-confirm.btn-success'))
            )
            ok_dialog.click()
            print(f"⚠ Kotak dialog muncul setelah edit, mengulang IDSBR: {idsbr}")
            driver.get("https://matchapro.web.bps.go.id/direktori-usaha")
            wait.until(EC.presence_of_element_located((By.NAME, "idsbr")))
            close_locked_popup()
            continue
        except:
            pass

        # ✅ Cek apakah form berhasil dimuat (submit-final ada)
        try:
            wait.until(EC.presence_of_element_located((By.ID, "submit-final")))
        except:
            print(f"⚠ Form tidak lengkap untuk IDSBR: {idsbr}, skip.")
            gagal_list.append(idsbr)
            driver.get("https://matchapro.web.bps.go.id/direktori-usaha")
            wait.until(EC.presence_of_element_located((By.NAME, "idsbr")))
            close_locked_popup()
            continue
        time.sleep(3)

        # 🔍 Cek email field kosong atau tidak → uncheck jika kosong
        try:
            email_field = wait.until(EC.presence_of_element_located((By.ID, "email")))
            email_value = email_field.get_attribute("value").strip()

            if email_value == "":
                print(f"📭 Email kosong untuk IDSBR: {idsbr}, pastikan checkbox email di-uncheck.")
                checkbox = wait.until(EC.element_to_be_clickable((By.ID, "check-email")))
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", checkbox)
                time.sleep(0.5)

                checked_attr = checkbox.get_attribute("checked")
                aria_checked = checkbox.get_attribute("aria-checked")

                if checked_attr or aria_checked == "true" or checkbox.is_selected():
                    checkbox.click()
                    time.sleep(0.5)
                    print(f"☑ Checkbox email berhasil di-uncheck untuk IDSBR: {idsbr}")
                else:
                    print(f"ℹ Checkbox email sudah uncheck.")
        except Exception as e:
            print(f"⚠ Gagal periksa/uncheck email untuk IDSBR {idsbr}: {e}")

        # ================================
        # 6-10. Input Data Profiling
        # ================================

                # --- Ambil semua baris kegiatan usaha (repeater) ---
        rows = driver.find_elements(By.CSS_SELECTOR, "div[data-repeater-item]")

        # --- 6. Input Kategori KBLI (Select2 per baris) ---
        try:
            kategori_list = [k.strip() for k in kategori_kbli.split(",") if k.strip() != ""]

            for i, row in enumerate(rows):
                if i >= len(kategori_list):
                    break

                kategori_value = kategori_list[i]

                # Ambil elemen Select2 tampilan (bukan <select>)
                select2_display = row.find_element(
                    By.CSS_SELECTOR, "span.select2-selection--single"
                )

                # Scroll pelan ke elemen
                smooth_scroll(driver, select2_display)
                time.sleep(0.2)

                # Klik tampilan Select2
                select2_display.click()
                time.sleep(0.3)

                # Input pencarian Select2
                search_input = wait.until(
                    EC.presence_of_element_located(
                        (By.CSS_SELECTOR, "input.select2-search__field")
                    )
                )

                # Ketik kategori → ENTER
                search_input.send_keys(kategori_value)
                time.sleep(1)
                search_input.send_keys(" - ")   # Tambah spasi agar hasil Select2 muncul
                time.sleep(1)
                search_input.send_keys(Keys.ENTER)
                time.sleep(0.5)

                print(f"✔ Kategori KBLI baris {i+1} terisi: {kategori_value}")

        except Exception as e:
            print(f"⚠ Gagal memilih kategori_kbli untuk IDSBR {idsbr}: {e}")
       
        # --- 7. Input KODE KBLI (Select2 per baris) ---        
        try:
            # Ambil list kode KBLI dari Excel
            kbli_list = [k.strip() for k in kode_kbli.split(",") if k.strip()]

            for i, row in enumerate(rows):
                if i >= len(kbli_list):
                    break

                kbli_value = kbli_list[i]

                # Ambil elemen Select2 tampilan (bukan <select>)
                try:
                    select2_display = row.find_element(
                        By.CSS_SELECTOR, "span.select2-selection[aria-labelledby^='select2-l_kbli']"
                    )
                except NoSuchElementException:
                    print(f"⚠ Select2 KBLI tidak ditemukan di row {i+1}")
                    continue

                # Scroll pelan ke elemen
                smooth_scroll(driver, select2_display)
                time.sleep(0.2)

                # Klik tampilan Select2
                select2_display.click()
                time.sleep(0.3)

                # Input pencarian Select2
                try:
                    search_input = WebDriverWait(driver, 10).until(
                        EC.visibility_of_element_located((By.CSS_SELECTOR, "input.select2-search__field"))
                    )
                except TimeoutException:
                    print(f"⚠ Input Select2 KBLI tidak muncul di row {i+1}")
                    continue

                # Ketik kategori → ENTER
                search_input.send_keys(kbli_value)
                # Tunggu hasil muncul
                # WebDriverWait(driver, 10).until(
                #     EC.visibility_of_element_located((By.CSS_SELECTOR, "li.select2-results__option"))
                # )
                time.sleep(2)
                # search_input.send_keys(" ")   # Tambah spasi agar hasil Select2 muncul
                # time.sleep(0.2)
                search_input.send_keys(Keys.ENTER)
                time.sleep(2)

                print(f"✔ KBLI baris {i+1} terisi: {kbli_value}")

        except Exception as e:
            print(f"⚠ Gagal mengisi kode_kbli untuk IDSBR {idsbr}: {e}")


        # --- 8. Input Kegiatan Profiling (per baris sesuai urutan Excel) ---
        try:
            kegiatan_list = [k.strip() for k in kegiatan_profiling.split(" | ") if k.strip()]

            for i, row in enumerate(rows):
                if i >= len(kegiatan_list):
                    break  # berhenti jika baris lebih banyak dari kegiatan

                kegiatan_value = kegiatan_list[i]

                kegiatan_input = row.find_element(
                    By.CSS_SELECTOR, "input[name='l_kegiatan_usaha']"
                )

                smooth_scroll(driver, kegiatan_input)

                kegiatan_input.clear()
                kegiatan_input.send_keys(kegiatan_value)

                print(f"✔ Kegiatan baris {i+1} terisi: {kegiatan_value}")

        except Exception as e:
            print(f"⚠ Gagal mengisi kegiatan_profiling: {e}")


        # =========================================================
        # 8A. Input Kepemilikan Usaha (Select2) — STABLE VERSION
        # =========================================================
        try:
            # 1️⃣ Pastikan tidak ada dropdown Select2 lama yang masih terbuka
            driver.find_element(By.TAG_NAME, "body").click()
            time.sleep(0.3)

            # 2️⃣ Ambil elemen Select2 tampilan (kliknya di sini)
            kepemilikan_display = wait.until(
                EC.element_to_be_clickable(
                    (
                        By.CSS_SELECTOR,
                        "span.select2-selection--single[aria-labelledby='select2-jenis_kepemilikan_usaha-container']"
                    )
                )
            )

            # 3️⃣ Scroll pelan (pakai fungsi yang sudah terbukti stabil)
            smooth_scroll(driver, kepemilikan_display)
            time.sleep(0.3)

            # 4️⃣ Klik Select2 (NORMAL click dulu)
            kepemilikan_display.click()
            time.sleep(0.5)

            # 5️⃣ Tunggu input search Select2 benar-benar ada
            search_input = wait.until(
                EC.presence_of_element_located(
                    (By.CSS_SELECTOR, "input.select2-search__field")
                )
            )

            # 6️⃣ Isi nilai
            search_input.clear()
            search_input.send_keys(kepemilikan_usaha)
            time.sleep(0.6)

            # 8️⃣ Pilih
            search_input.send_keys(Keys.ENTER)
            time.sleep(0.6)

            print(f"✔ Kepemilikan usaha terisi: {kepemilikan_usaha}")

        except Exception as e:
            print(f"⚠ Gagal mengisi kepemilikan_usaha IDSBR {idsbr}: {e}")





        # # =========================================================
        # # 8B. Input Bentuk Badan Hukum / Usaha (Select2)
        # # =========================================================
        # try:
        #     badan_hukum_display = wait.until(
        #         EC.presence_of_element_located(
        #             (
        #                 By.CSS_SELECTOR,
        #                 "span.select2-selection[aria-labelledby='select2-badan_usaha-container']"
        #             )
        #         )
        #     )

        #     driver.execute_script(
        #         "arguments[0].scrollIntoView({block:'center'});",
        #         badan_hukum_display
        #     )
        #     time.sleep(0.3)

        #     driver.execute_script("arguments[0].click();", badan_hukum_display)
        #     time.sleep(0.5)

        #     search_input = wait.until(
        #         EC.visibility_of_element_located(
        #             (By.CSS_SELECTOR, "input.select2-search__field")
        #         )
        #     )

        #     search_input.clear()
        #     search_input.send_keys(badan_hukum)
        #     time.sleep(0.4)
        #     search_input.send_keys(Keys.ENTER)

        #     print(f"✔ Badan hukum terisi: {badan_hukum}")

        # except Exception as e:
        #     print(f"⚠ Gagal mengisi badan_hukum IDSBR {idsbr}: {e}")

        # =========================================================
        # 8C. Klik Radio Button Jaringan Usaha → Tunggal
        # =========================================================
        try:
            radio_tunggal = wait.until(
                EC.presence_of_element_located(
                    (By.XPATH, "//label[contains(., 'Tunggal')]")
                )
            )

            driver.execute_script(
                "arguments[0].scrollIntoView({block:'center'});",
                radio_tunggal
            )
            time.sleep(0.3)

            driver.execute_script("arguments[0].click();", radio_tunggal)
            print("☑ Jaringan usaha: Tunggal dipilih")

        except Exception as e:
            print(f"⚠ Gagal klik radio Tunggal IDSBR {idsbr}: {e}")

        # =========================================================
        # 8D. Input Sektor Institusi (Select2) — STABLE
        # =========================================================
        try:
            # 1️⃣ Reset fokus (hindari dropdown / overlay sisa)
            driver.find_element(By.TAG_NAME, "body").click()
            time.sleep(0.3)

            # 2️⃣ Ambil elemen Select2 tampilan (parent clickable)
            sektor_display = wait.until(
                EC.element_to_be_clickable(
                    (
                        By.CSS_SELECTOR,
                        "span.select2-selection--single[aria-labelledby='select2-sektor_institusi_usaha-container']"
                    )
                )
            )

            # 3️⃣ Scroll pelan ke elemen
            smooth_scroll(driver, sektor_display)
            time.sleep(0.3)

            # 4️⃣ Klik Select2 (NORMAL click)
            sektor_display.click()
            time.sleep(0.5)

            # 5️⃣ Tunggu input search Select2 muncul
            search_input = wait.until(
                EC.presence_of_element_located(
                    (By.CSS_SELECTOR, "input.select2-search__field")
                )
            )

            # 6️⃣ Isi nilai (contoh: "S14")
            search_input.clear()
            search_input.send_keys(sektor_institusi)
            time.sleep(0.6)

            # 8️⃣ Pilih
            search_input.send_keys(Keys.ENTER)
            time.sleep(0.6)

            print(f"✔ Sektor institusi terisi: {sektor_institusi}")

        except Exception as e:
            print(f"⚠ Gagal mengisi sektor_institusi IDSBR {idsbr}: {e}")


        # --- 9. Input Sumber Profiling (input biasa) ---
        try:
            sumber_input = driver.find_element(By.NAME, "sumber_profiling")
            sumber_input.clear()
            sumber_input.send_keys(sumber_profiling)
        except Exception as e:
            print(f"⚠ Gagal mengisi sumber_profiling untuk IDSBR {idsbr}: {e}")


        # --- 10. Input Catatan Profiling ---
        try:
            catatan_input = driver.find_element(By.ID, "catatan_profiling")
            catatan_input.clear()
            catatan_input.send_keys(catatan_profiling)
        except Exception as e:
            print(f"⚠ Gagal mengisi catatan_profiling untuk IDSBR {idsbr}: {e}")


        # ✅ Klik radio button "Aktif" (id=kondisi_aktif)
        try:
            radio_aktif = wait.until(EC.presence_of_element_located((By.ID, "kondisi_aktif")))
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", radio_aktif)
            time.sleep(0.5)

            if not radio_aktif.is_selected():
                try:
                    radio_aktif.click()
                except:
                    driver.execute_script("arguments[0].click();", radio_aktif)

            # verifikasi ulang
            if radio_aktif.is_selected():
                print(f"☑ Radio 'Aktif' dipilih untuk IDSBR: {idsbr}")
            else:
                # fallback: coba klik label
                try:
                    label_aktif = driver.find_element(By.CSS_SELECTOR, "label[for='kondisi_aktif']")
                    driver.execute_script("arguments[0].click();", label_aktif)
                    print(f"☑ Radio 'Aktif' dipilih via label untuk IDSBR: {idsbr}")
                except:
                    print(f"⚠ Radio 'Aktif' masih gagal dipilih untuk IDSBR: {idsbr}")
        except Exception as e:
            print(f"⚠ Gagal klik radio 'Aktif' untuk IDSBR {idsbr}: {e}")

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

        # 14. Kembali ke halaman awal
        driver.get("https://matchapro.web.bps.go.id/direktori-usaha")
        wait.until(EC.presence_of_element_located((By.NAME, "idsbr")))
        close_locked_popup()

        print(f"✅ Sukses input data IDSBR: {idsbr}")
        sukses_list.append(idsbr)

    except Exception as e:
        print(f"❌ Gagal memproses IDSBR {idsbr}: {e}")
        gagal_list.append(idsbr)
        try:
            driver.get("https://matchapro.web.bps.go.id/direktori-usaha")
            wait.until(EC.presence_of_element_located((By.NAME, "idsbr")))
            close_locked_popup()
        except:
            pass
        continue

# ====== Simpan Log ======
result_df = pd.DataFrame({
    'idsbr': df['idsbr'],
    'status': df['idsbr'].apply(lambda x: 'SUKSES' if x in sukses_list else ('GAGAL' if x in gagal_list else 'TIDAK DIPROSES'))
})
result_df.to_excel("hasil_update_idsbr.xlsx", index=False)

print("\n🎉 Selesai memproses semua data.")
print(f"Total sukses: {len(sukses_list)} | Total gagal: {len(gagal_list)}")
print("list yang berhasil :")
print(sukses_list)
print("list yang gagal :")
print(gagal_list)
