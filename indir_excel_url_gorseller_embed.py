
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel'deki URL'leri indirip her satır için Şasi numarasına göre klasör oluşturan araç.
Bu sürümde Excel ve çıktı yolları KOD İÇİNDEN ayarlanır (komut satırı kullanılmaz).

Gereksinimler:
    pip install pandas openpyxl requests
"""

import concurrent.futures as futures
import csv
import os
import re
import sys
import time
from pathlib import Path
from urllib.parse import urlsplit, unquote

# İsteğe bağlı: requests yoksa urllib ile deneyelim
try:
    import requests
except Exception:  # pragma: no cover
    requests = None

import pandas as pd

# ========================
# CONFIG (burayı düzenleyin)
# ========================
EXCEL_PATH   = r"C:\Users\berkeb\Desktop\CIS_RPA\CIS için Tekil Şasi Listesi.xlsx"  # .xlsx dosya yolu
SHEET_NAME   = 0                                            # None => ilk sayfa
OUTPUT_DIR   = r"C:\Users\berkeb\Desktop\indirilen_gorseller3"  # çıktı klasörü
VIN_COL      = "Şasi"                                          # şasi/vin sütun adı
WORKERS      = 8                                               # eşzamanlı indirme iş parçacığı sayısı
TIMEOUT      = 30.0                                            # saniye
NO_SKIP_EXISTING = False                                       # True => mevcut dosyanın ÜZERİNE yaz
INSECURE_SSL = False                                           # True => SSL doğrulamasını kapatır (önerilmez)
OPEN_LOG_WHEN_DONE = False                                     # Windows'ta bittiğinde log'u aç
PAUSE_WHEN_DONE    = True                                      # Çift tıklamada konsol kapanmasın

# Sütun -> Dosya numarası eşlemesi (istenen)
COLUMN_TO_INDEX = {
    "Araç Dış Arka Görsel Yolu": "Araç Dış Arka Görsel Yolu",
    "Araç Dış Ön Görsel Yolu": "Araç Dış Ön Görsel Yolu",
    "Araç Dış Yan Görsel Yolu": "Araç Dış Yan Görsel Yolu",
    "Araç İç Arka Görsel Yolu": "Araç İç Arka Görsel Yolu",
    "Araç İç Ön Görsel Yolu": "Araç İç Ön Görsel Yolu",
    "Araç İç Yan Görsel Yolu": "Araç İç Yan Görsel Yolu",
}

# Bazı sunucular 403 dönmesin diye basit bir User-Agent
DEFAULT_HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                  "AppleWebKit/537.36 (KHTML, like Gecko) "
                  "Chrome/124.0 Safari/537.36",
    "Accept": "*/*",
    "Accept-Language": "tr-TR,tr;q=0.9,en-US;q=0.8,en;q=0.7",
}

VALID_EXTS = {".jpg",".jpeg",".png",".webp",".bmp",".gif",".tiff",".tif"}

def normalize(s: str) -> str:
    if s is None:
        return ""
    tr_map = str.maketrans("şŞıİçÇöÖüÜğĞ", "sSiIcCoOuUgG")
    s = str(s).translate(tr_map).strip().lower()
    s = re.sub(r"\s+", " ", s)
    return s

def slugify(text: str) -> str:
    # Klasör ismini güvenli hale getir
    text = str(text).strip()
    text = re.sub(r"[\\/:*?\"<>|]", "_", text)  # Windows yasaklı karakterler
    text = re.sub(r"\s+", "_", text)
    return text or "bos"

def find_column(df_cols, target_name):
    # Aynı/benzer sütunu bulmaya çalış (birebir, lower/normalize)
    if target_name in df_cols:
        return target_name
    low_map = {str(c).lower(): c for c in df_cols}
    if target_name and target_name.lower() in low_map:
        return low_map[target_name.lower()]
    n_target = normalize(target_name or "")
    for c in df_cols:
        if normalize(str(c)) == n_target:
            return c
    return None

def detect_vin_column(df_cols, user_hint=None):
    if user_hint:
        c = find_column(df_cols, user_hint)
        if c:
            return c
    # Heuristik: 'Şasi', 'VIN', 'Chassis' gibi geçen sütunlar
    candidates = []
    for c in df_cols:
        n = normalize(str(c))
        if any(k in n for k in ["vin", "sasi", "sasino", "sasinumarasi", "chassis"]):
            candidates.append(c)
    return candidates[0] if candidates else None

def extract_urls(cell_value):
    if cell_value is None:
        return []
    text = str(cell_value).strip()
    if not text:
        return []
    # URL ayıklama: http/https ile başlayanları al
    pattern = r"(https?://[^\s,;\"\)]+)"
    urls = re.findall(pattern, text)
    return [u.strip() for u in urls if u.strip()]

def pick_extension(url_path, content_type=None):
    path = urlsplit(url_path).path
    path = unquote(path)
    ext = os.path.splitext(path)[1].lower()
    # Sorgu parametresi ile gelen .jpg?auth=... gibi durumlar için temizle
    ext = ext.split("?")[0]
    if ext in VALID_EXTS:
        return ext
    if content_type:
        ct = content_type.lower()
        if "image/jpeg" in ct or "image/jpg" in ct:
            return ".jpg"
        if "image/png" in ct:
            return ".png"
        if "image/webp" in ct:
            return ".webp"
        if "image/gif" in ct:
            return ".gif"
        if "image/bmp" in ct:
            return ".bmp"
        if "image/tiff" in ct:
            return ".tiff"
    return ".jpg"  # varsayılan

def ensure_unique(path: Path) -> Path:
    if not path.exists():
        return path
    stem, ext = path.stem, path.suffix
    i = 2
    while True:
        candidate = path.with_name(f"{stem} ({i}){ext}")
        if not candidate.exists():
            return candidate
        i += 1

def download_one(session, url, dest: Path, timeout, verify_ssl=True):
    try:
        if requests is None:
            # urllib ile indirme (uzantıyı Content-Type'tan güncelle)
            import urllib.request
            req = urllib.request.Request(url, headers=DEFAULT_HEADERS)
            with urllib.request.urlopen(req, timeout=timeout) as r:
                ct = r.getheader("Content-Type")
                ext = pick_extension(url, ct)
                if dest.suffix.lower() != ext.lower():
                    dest = dest.with_suffix(ext)
                dest.parent.mkdir(parents=True, exist_ok=True)
                final_dest = ensure_unique(dest) if dest.exists() else dest
                with open(final_dest, "wb") as f:
                    f.write(r.read())
            return True, "indirildi"
        # requests
        with session.get(url, stream=True, timeout=timeout, verify=verify_ssl, allow_redirects=True) as r:
            if r.status_code != 200:
                return False, f"HTTP {r.status_code}"
            ext = pick_extension(url, r.headers.get("Content-Type"))
            if dest.suffix.lower() != ext.lower():
                dest = dest.with_suffix(ext)
            dest.parent.mkdir(parents=True, exist_ok=True)
            final_dest = ensure_unique(dest) if dest.exists() else dest
            with open(final_dest, "wb") as f:
                for chunk in r.iter_content(chunk_size=8192):
                    if chunk:
                        f.write(chunk)
        return True, "indirildi"
    except Exception as e:
        return False, f"hata: {e}"

def run():
    excel_path = Path(EXCEL_PATH)
    out_dir = Path(OUTPUT_DIR)
    out_dir.mkdir(parents=True, exist_ok=True)

    # Excel'i oku
    try:
        df = pd.read_excel(excel_path, sheet_name=SHEET_NAME)
    except Exception as e:
        print("Excel okunamadı:", e)
        if PAUSE_WHEN_DONE:
            try:
                input("\nÇıkmak için Enter'a basın...")
            except Exception:
                pass
        sys.exit(1)

    df_cols = list(df.columns)
    # VIN/Şasi sütunu tespit et
    vin_col = detect_vin_column(df_cols, user_hint=VIN_COL)
    if not vin_col:
        print("Şasi/VIN sütunu bulunamadı. CONFIG bölümünde VIN_COL'u doğru sütun adıyla değiştirin.")
        print("Mevcut sütunlar:", df_cols)
        if PAUSE_WHEN_DONE:
            try:
                input("\nÇıkmak için Enter'a basın...")
            except Exception:
                pass
        sys.exit(2)

    # Görsel sütunlarını bul
    col_map = {}
    for canon_name, idx in COLUMN_TO_INDEX.items():
        found = find_column(df_cols, canon_name)
        if found is None:
            print(f"UYARI: '{canon_name}' sütunu Excel'de bulunamadı. (Atlanacak)")
        else:
            col_map[found] = idx

    if not col_map:
        print("Hiçbir görsel yolu sütunu bulunamadı. Lütfen sütun adlarını kontrol edin.")
        if PAUSE_WHEN_DONE:
            try:
                input("\nÇıkmak için Enter'a basın...")
            except Exception:
                pass
        sys.exit(3)

    # Log dosyası
    log_path = out_dir / "indirilenler_log.csv"
    log_f = open(log_path, "w", newline="", encoding="utf-8")
    log = csv.writer(log_f)
    log.writerow(["satir_index","sasi","sutun","dosya_index","url","dosya_yolu","durum","not"])

    tasks = []

    # Requests session
    sess = requests.Session() if requests else None
    if sess:
        sess.headers.update(DEFAULT_HEADERS)

    verify_ssl = not INSECURE_SSL

    def schedule_download(vin_value, url, file_idx, seq_num):
        # Dosya adı: 1.jpg, birden fazla URL varsa 1_2.jpg
        base_name = f"{file_idx}"
        if seq_num > 1:
            base_name += f"_{seq_num}"
        dest = out_dir / slugify(vin_value) / base_name
        # Uzantı, indirme sırasında Content-Type'a göre düzeltilecek
        dest = dest.with_suffix(".jpg")
        return (url, dest)

    for i, row in df.iterrows():
        vin_value = row.get(vin_col, None)
        if pd.isna(vin_value) or str(vin_value).strip() == "":
            # Şasi boşsa atla
            continue

        for col_name, file_idx in col_map.items():
            cell = row.get(col_name, None)
            urls = extract_urls(cell)
            if not urls:
                # log boş
                log.writerow([i, vin_value, col_name, file_idx, "", "", "bos", ""])
                continue
            for k, url in enumerate(urls, start=1):
                url = url.strip()
                if not (url.startswith("http://") or url.startswith("https://")):
                    log.writerow([i, vin_value, col_name, file_idx, url, "", "gecersiz_url", "http/https ile baslamiyor"])
                    continue
                url, dest = schedule_download(vin_value, url, file_idx, k)
                tasks.append((i, vin_value, col_name, file_idx, url, dest))

    # İndirmeleri yürüt
    total = len(tasks)
    print(f"Toplam {total} dosya indirilecek.")
    start = time.time()

    def worker(t):
        i, vin_value, col_name, file_idx, url, dest = t
        if (not NO_SKIP_EXISTING) and dest.exists():
            return (i, vin_value, col_name, file_idx, url, str(dest), "atlandi", "dosya zaten var")
        ok, note = download_one(sess, url, dest, timeout=TIMEOUT, verify_ssl=verify_ssl)
        status = "indirildi" if ok else "hata"
        return (i, vin_value, col_name, file_idx, url, str(dest), status, note)

    if WORKERS and WORKERS > 1 and total > 1:
        with futures.ThreadPoolExecutor(max_workers=WORKERS) as ex:
            for res in ex.map(worker, tasks):
                log.writerow(list(res))
    else:
        for t in tasks:
            res = worker(t)
            log.writerow(list(res))

    log_f.close()
    elapsed = time.time() - start
    print("Bitti. Süre: {:.1f} sn".format(elapsed))
    print("Log:", log_path)

    if os.name == "nt" and OPEN_LOG_WHEN_DONE:
        try:
            os.startfile(str(log_path))
        except Exception:
            pass

    if PAUSE_WHEN_DONE:
        try:
            input("\nKapatmak için Enter'a basın...")
        except Exception:
            pass

if __name__ == "__main__":
    run()
