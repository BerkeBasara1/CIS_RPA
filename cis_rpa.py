import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from selenium.webdriver import ActionChains
import win32com.client
import win32gui
import win32con
import pyperclip
import pyautogui

URL_DASH = "https://diseadmin.skoda-auto.com/sms/dashboard"
URL_CARS = "https://diseadmin.skoda-auto.com/sms/cars"

# ---- Genel 2 sn bekleme helper'ı ----
def pause():
    time.sleep(5)


def click_excel_icon_with_image(icon_path=r"C:\Users\berkeb\Downloads\excel_gorsel.png", timeout=10):
    """
    Ekranda verilen ikon görselini (excel_gorsel.png) arar,
    bulursa merkezine tıklar.
    """
    print("[30] Excel ikonu görsel benzerlik ile aranıyor...")
    pause()

    pyautogui.FAILSAFE = True  # mouse'u köşeye götürürsen script durur

    start = time.time()
    location = None

    while time.time() - start < timeout and location is None:
        try:
            location = pyautogui.locateCenterOnScreen(
                icon_path,
                confidence=0.8  # bulamazsa 0.7/0.6 yapabilirsin
            )
        except Exception as e:
            print(f"    locateOnScreen sırasında hata: {e}")
            break

        if location is None:
            print("    Excel ikonu henüz bulunamadı, tekrar deniyorum...")
            time.sleep(0.5)

    if location is None:
        print("⛔ Excel ikonu belirtilen süre içinde bulunamadı.")
        return False

    x, y = location
    x, y = location
    print(f"    ✓ Excel ikonu bulundu: ({x}, {y}) → tıklanıyor...")
    pause()

    try:
        # Burada fail-safe devreye girebilir
        pyautogui.moveTo(x, y, duration=0.4)
        pyautogui.click()
        print("    ✓ PyAutoGUI ile Excel ikonuna tıklandı.")
        pause()
        return True

    except FailSafeException:
        print("⛔ PyAutoGUI FAIL-SAFE tetiklendi! Mouse ekranın köşesine gitmiş.")
        print("   → İşlem güvenlik sebebiyle iptal edildi. Mouse'u köşeden uzak tutup tekrar deneyin.")
        return False


# ⚠ BURAYI KENDİN DOLDUR (ve mümkünse şifreyi sonra değiştir)
SKODA_USER = "XTR01374AHX"
SKODA_PASS = "Roadyuce190768!"


def copy_chassis_from_excel():
    """
    Ön plandaki Excel'den aktif sayfanın E2 hücresini okur
    ve string olarak geri döner, panoya da kopyalar.
    (Görsel tıklama yok, sadece COM ile okuyor.)
    """
    print("[20] Excel COM nesnesine bağlanılıyor...")
    pause()

    try:
        excel = win32com.client.GetActiveObject("Excel.Application")
    except Exception:
        excel = win32com.client.Dispatch("Excel.Application")

    excel.Visible = True
    pause()

    wb = excel.ActiveWorkbook
    ws = wb.ActiveSheet

    print("[21] E2 hücresi okunuyor...")
    pause()
    chassis = ws.Range("E2").Value

    if chassis is None:
        print("    ⚠ E3 hücresinde değer yok (None).")
        return None

    chassis_str = str(chassis)
    pyperclip.copy(chassis_str)
    print(f"    ✓ Şasi değeri kopyalandı: {chassis_str}")
    pause()

    return chassis_str




def create_driver():
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    # options.add_argument("--headless=new")
    driver = webdriver.Chrome(
        service=Service(ChromeDriverManager().install()),
        options=options
    )
    pause()
    return driver


def find_element_with_fallback(driver, wait, selectors, desc):
    last_exc = None
    for by, value in selectors:
        try:
            print(f"  → {desc} için denenen locator: {by} = {value}")
            el = wait.until(EC.visibility_of_element_located((by, value)))
            print(f"  ✓ {desc} bulundu: {by} = {value}")
            pause()
            return el
        except Exception as e:
            last_exc = e
    raise last_exc


def switch_to_login_window(driver):
    pause()
    handles = driver.window_handles
    print(f"[2] Bulunan pencere sayısı: {len(handles)}")
    pause()

    for h in handles:
        driver.switch_to.window(h)
        url = driver.current_url
        print(f"    → Handle: {h}, URL: {url}")
        if "identity" in url or "login" in url.lower():
            print("    ✓ Login penceresi bulundu, bu pencereye geçildi.")
            pause()
            return

    print("    ⚠ 'identity' içeren pencere bulunamadı, mevcut pencerede devam ediliyor.")
    pause()
def get_next_chassis_from_excel(current_chassis):
    print("[NEXT] Hatalı şasi → bir alt satıra geçiliyor...")

    excel = win32com.client.GetActiveObject("Excel.Application")
    wb = excel.ActiveWorkbook
    ws = wb.ActiveSheet

    used_rows = ws.UsedRange.Rows.Count

    row_found = None
    for r in range(2, used_rows + 1):
        val = ws.Cells(r, 5).Value
        if val and str(val).strip() == str(current_chassis).strip():
            row_found = r
            break

    if not row_found:
        print("⚠ Şasi bulunamadı, döngü duracak.")
        return None

    next_row = row_found + 1
    new_val = ws.Cells(next_row, 5).Value

    if not new_val:
        print("⚠ Bir alt satır boş → iş bitti.")
        return None

    new_val = str(new_val).strip()
    pyperclip.copy(new_val)
    print(f"    ✓ Yeni şasi bulundu: {new_val}")

    return new_val

def open_cars_and_click_new_button(driver, chassis_value):
    print("[10] Cars sayfası açılıyor...")
    driver.get(URL_CARS)
    pause()

    wait = WebDriverWait(driver, 30)

    try:
        wait.until(EC.url_contains("/sms/cars"))
        print(f"    Cars URL: {driver.current_url}")
    except TimeoutException:
        print("⚠ /sms/cars URL'si zamanında yüklenmedi, ama devam ediyorum.")
    pause()

    # --- New car butonu ---
    print("[11] New car butonu (class='sc-hyBbbR hOXGnK') aranıyor...")
    new_car_btn = wait.until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, "button.sc-hyBbbR.hOXGnK"))
    )
    print("    ✓ New car butonu bulundu.")
    pause()
    new_car_btn.click()
    pause()

    # --- Model select ---
    print("[13] Model select aranıyor...")
    model_select_el = wait.until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, "select.sc-jhlPcU.iTPTrg"))
    )
    print("    ✓ Model select bulundu.")
    pause()

    print("[14] 'Elroq' modeli seçiliyor...")
    select = Select(model_select_el)
    select.select_by_value("elroq")
    pause()

    # --- Excel ikonuna tıkla (aç) ---
    print("[15] Excel ikonu açılıyor...")
    click_excel_icon_with_image(r"C:\Users\berkeb\Downloads\excel_gorsel.png")
    pause()

 
    # --- Excel ikonuna tıkla (küçült) ---
    print("[17] Excel küçültülüyor...")
    click_excel_icon_with_image(r"C:\Users\berkeb\Downloads\excel_gorsel.png")
    pause()

    # --- VIN input ---
    print("[19] VIN input aranıyor...")
    vin_input = wait.until(
        EC.element_to_be_clickable((By.XPATH, "//input[@placeholder='VIN']"))
    )
    pause()

    print("[20] VIN yazılıyor...")
    vin_input.clear()
    vin_input.send_keys(chassis_value)
    pause()

    # --- CREATE tıkla ---
    print("[21] Create aranıyor...")
    create_btn = wait.until(
        EC.element_to_be_clickable((
            By.XPATH,
            "//button[contains(@class,'sc-hyBbbR') and contains(@class,'hOXGnK') and normalize-space()='Create']"
        ))
    )
    pause()
    create_btn.click()
    pause()
        # --- CREATE sonrası hata kontrolü ---
    print("[22.1] Create sonrası hata kontrol ediliyor...")

    try:
        error_box = driver.find_element(By.XPATH, "//*[contains(text(), 'Failed to get car preview')]")
        print("⛔ Sistem hatası tespit edildi! Bu şasi zaten kayıtlı.")
        # ALTSATIR ŞASİYİ AL VE RETURN ET
        return get_next_chassis_from_excel(chassis_value)

    except:
        print("    ✓ Hata bulunamadı, işleme devam ediliyor...")

    # ============================================================
    # === 1) DIŞ TASARIM SLIDER İLK 3 GÖRSEL ======================
    # ============================================================

    print("[23] DIŞ tasarım slider ilk 3 görsel alınıyor...")

    img_elements = wait.until(
        EC.presence_of_all_elements_located(
            (By.CSS_SELECTOR, "ul.slider.animated li.slide img")
        )
    )

    if len(img_elements) < 3:
        print("⚠ Dış tasarım slider’da 3 görsel yok!")
        return None

    exterior_urls = [
        img_elements[0].get_attribute("src"),
        img_elements[1].get_attribute("src"),
        img_elements[2].get_attribute("src")
    ]

    for idx, url in enumerate(exterior_urls):
        print(f"    ✓ Dış {idx+1}. görsel: {url}")

    write_multiple_urls_for_chassis(chassis_value, exterior_urls, mode="exterior")
    print("✓ Dış tasarım görselleri Excel'e yazıldı.")
    pause()

    # ============================================================
    # === 2) İÇ TASARIM SEKME TIKLAMA =============================
    # ============================================================

    print("[30] İç tasarım sekmesi aranıyor...")

    interior_tab = wait.until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, "div.sc-ejfMa-d.eoRTQh"))
    )
    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", interior_tab)
    interior_tab.click()
    pause()

    # ============================================================
    # === 3) İÇ TASARIM SLIDER İLK 3 GÖRSEL ========================
    # ============================================================

    print("[31] İç tasarım slider ilk 3 görsel alınıyor...")

    interior_imgs = wait.until(
        EC.presence_of_all_elements_located(
            (By.CSS_SELECTOR, "div.slider-wrapper.axis-horizontal ul.slider.animated li.slide img")
        )
    )

    if len(interior_imgs) < 3:
        print("⚠ İç tasarım slider’da 3 görsel yok!")
        return None

    interior_urls = [
        interior_imgs[0].get_attribute("src"),
        interior_imgs[1].get_attribute("src"),
        interior_imgs[2].get_attribute("src")
    ]

    for idx, url in enumerate(interior_urls):
        print(f"    ✓ İç {idx+1}. görsel: {url}")

    write_multiple_urls_for_chassis(chassis_value, interior_urls, mode="interior")
    print("✓ İç tasarım görselleri Excel'e yazıldı.")
    pause()

    print("🎉 TÜM GÖRSELLER ALINDI VE EXCEL'E YAZILDI.")
        # --- CANCEL butonuna tıkla ---
        # --- CANCEL butonuna tıkla ---
    print("[40] CANCEL butonu aranıyor...")

    cancel_btn = wait.until(
        EC.element_to_be_clickable(
            (By.CSS_SELECTOR, "button.sc-hyBbbR.cIGnGg")
        )
    )

    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", cancel_btn)
    driver.execute_script("arguments[0].click();", cancel_btn)
    print("    ✓ CANCEL tıklandı.")
    pause()


    # --- OK butonuna tıkla (JS FORCE CLICK) ---
    print("[41] OK butonu aranıyor...")

    ok_btn = wait.until(
        EC.visibility_of_element_located(
            (By.CSS_SELECTOR, "button.sc-hyBbbR.hOXGnK")
        )
    )

    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", ok_btn)
    driver.execute_script("arguments[0].click();", ok_btn)   # ← FORCE CLICK
    print("    ✓ OK tıklandı (JS Force Click).")
    pause()


    # --- EXCEL: mevcut şasi satırının bir altındaki şasiyi kopyala ---
    print("[42] Excel'de bir alt satırdaki şasi alınıyor...")

    try:
        excel = win32com.client.GetActiveObject("Excel.Application")
    except Exception:
        excel = win32com.client.Dispatch("Excel.Application")

    excel.Visible = True
    pause()

    wb = excel.ActiveWorkbook
    ws = wb.ActiveSheet

    used_rows = ws.UsedRange.Rows.Count
    row_found = None

    for r in range(2, used_rows + 1):
        val = ws.Cells(r, 5).Value
        if val and str(val).strip() == str(chassis_value).strip():
            row_found = r
            break

    if not row_found:
        print("⚠ Şasi satırı bulunamadı! Varsayılan satır 2 kullanılacak.")
        row_found = 2

    next_row = row_found + 1
    new_chassis_value = ws.Cells(next_row, 5).Value

    if new_chassis_value:
        new_chassis_value = str(new_chassis_value).strip()
        pyperclip.copy(new_chassis_value)
        print(f"    ✓ Bir alt satırdaki şasi alındı: {new_chassis_value}")
    else:
        print("⚠ Bir alt satırda şasi yok! Döngü duracak.")
        return None

    pause()

    # →→→ DÖNGÜ İÇİN BURASI ŞART! ←←←
    return new_chassis_value



def write_multiple_urls_for_chassis(chassis_value, img_urls, mode="exterior"):
    """
    mode = "exterior" → dış tasarım için (kolonlar F, G, H)
    mode = "interior" → iç tasarım için (kolonlar I, J, K)
    """
    print(f"[Excel] {mode} görselleri yazılıyor...")
    pause()

    try:
        excel = win32com.client.GetActiveObject("Excel.Application")
    except Exception:
        excel = win32com.client.Dispatch("Excel.Application")

    excel.Visible = True
    pause()

    wb = excel.ActiveWorkbook
    ws = wb.ActiveSheet

    used_rows = ws.UsedRange.Rows.Count
    used_cols = ws.UsedRange.Columns.Count

    # 1) Şasi satırını bul (E sütunu)
    row_found = None
    for r in range(2, used_rows + 1):
        val = ws.Cells(r, 5).Value
        if val and str(val).strip() == str(chassis_value).strip():
            row_found = r
            break

    if not row_found:
        print("⚠ Şasi bulunamadı, 2. satır kullanılacak.")
        row_found = 2

    # 2) Mode'a göre başlık seç
    if mode == "exterior":
        headers = [
            "Araç Dış Arka Görsel Yolu",
            "Araç Dış Ön Görsel Yolu",
            "Araç Dış Yan Görsel Yolu"
        ]
    else:  # interior
        headers = [
            "Araç İç Arka Görsel Yolu",
            "Araç İç Ön Görsel Yolu",
            "Araç İç Yan Görsel Yolu"
        ]

    # 3) Başlıklara göre kolon bul
    header_columns = []
    for header in headers:
        col_found = None
        for c in range(1, used_cols + 1):
            val = ws.Cells(1, c).Value
            if isinstance(val, str) and val.strip() == header:
                col_found = c
                break
        header_columns.append(col_found)

    # 4) URL’leri yaz
    for idx, img_url in enumerate(img_urls):
        col = header_columns[idx]
        if col is None:
            print(f"⚠ '{headers[idx]}' bulunamadı!")
            continue
        ws.Cells(row_found, col).Value = img_url
        print(f"    ✓ {mode} {idx+1}. görsel → {ws.Cells(row_found, col).Address}")

    pause()








def login_and_open_dashboard():
    if not SKODA_USER or not SKODA_PASS:
        print("⚠ SKODA_USER / SKODA_PASS doldurulmamış!")
        return

    driver = create_driver()
    print("[1] Dashboard sayfası açılıyor...")
    driver.get(URL_DASH)
    pause()

    wait = WebDriverWait(driver, 30)

    try:
        print(f"    İlk URL: {driver.current_url}")
        pause()
        switch_to_login_window(driver)

        print(f"[3] Aktif URL: {driver.current_url}")
        pause()

        # ============================
        # 1) USERNAME INPUT
        # ============================
        username_selectors = [
            (By.NAME, "userName"),
            (By.CSS_SELECTOR, "input[name='userName']"),
            (By.XPATH, "//input[@placeholder='Kullanıcı Adı']"),
        ]

        print("[4] Username aranıyor...")
        username_box = find_element_with_fallback(driver, wait, username_selectors, "Username input")

        # ============================
        # 2) PASSWORD INPUT
        # ============================
        password_selectors = [
            (By.NAME, "password"),
            (By.CSS_SELECTOR, "input[name='password']"),
            (By.XPATH, "//input[@placeholder='Şifre']"),
            (By.CSS_SELECTOR, "input.input-password"),
        ]

        print("[5] Password aranıyor...")
        password_box = find_element_with_fallback(driver, wait, password_selectors, "Password input")

        # ============================
        # 3) DEĞERLERİN YAZILMASI
        # ============================
        print("[6] Kullanıcı adı ve şifre yazılıyor...")
        username_box.clear()
        username_box.send_keys(SKODA_USER)
        pause()

        password_box.clear()
        password_box.send_keys(SKODA_PASS)
        pause()

        # ============================
        # 4) LOGIN (‘DEVAM’) BUTONU
        # ============================
        login_btn_selectors = [
            (By.CSS_SELECTOR, "button[type='submit']"),
            (By.CSS_SELECTOR, "button.btn.btn-primary"),
            (By.XPATH, "//button[contains(text(), 'Devam')]"),
        ]

        print("[7] 'Devam' butonu aranıyor...")
        login_button = find_element_with_fallback(driver, wait, login_btn_selectors, "Login button")

        print("[8] 'Devam' tıklanıyor...")
        login_button.click()
        pause()

        # ============================
        # 5) DASHBOARD YÜKLENİYOR MU?
        # ============================
        print("[9] Dashboard yükleniyor...")
        time.sleep(5)

        for h in driver.window_handles:
            driver.switch_to.window(h)
            if "diseadmin.skoda-auto.com" in driver.current_url:
                print("✅ Dashboard'a giriş başarılı!")
                break

        pause()

        # =============================================================
        #  OTOMATİK RPA ŞASİ DÖNGÜSÜ BAŞLIYOR
        # =============================================================
        print("🔁 OTOMATİK RPA BAŞLATILIYOR...")

        next_chassis = copy_chassis_from_excel()

        if not next_chassis:
            print("⛔ İlk şasi bulunamadı!")
            return

        while next_chassis:
            print(f"\n\n🚗 Yeni şasi işleniyor: {next_chassis}\n")

            new_chassis = open_cars_and_click_new_button(driver, next_chassis)

            if not new_chassis:
                print("⛔ Yeni şasi yok → döngü durduruldu.")
                break

            print(f"➡ Bir sonraki şasi: {new_chassis}")

            next_chassis = new_chassis

        print("🎉 Tüm şasi satırları işlendi! Döngü tamamlandı.")

    except Exception as e:
        print(f"⛔ Beklenmeyen hata: {e}")




if __name__ == "__main__":
    login_and_open_dashboard()
