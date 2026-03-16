import shutil
from pathlib import Path
import pandas as pd

# =======================
# AYARLAR (SENİN YOLLARIN)
# =======================
EXCEL_PATH = r"C:\Users\berkeb\Downloads\Pipeline All 09.02.2026 1 (1).xlsx"

UNIQUE_SHEET = "Unique"
PIPELINE_SHEET = "Pipeline All"

# Görsellerin kök klasörü (içinde TMB... klasörleri var)
IMAGES_ROOT = r"C:\Users\berkeb\Downloads\indirilen_gorseller2\indirilen_gorseller2"

# Çıktı: Pipeline All'daki her şasi için klasör burada oluşacak
OUT_ROOT = r"C:\Users\berkeb\Downloads\pipeline_out"

# Kolon adayları (otomatik bulma)
MATCH_ID_CANDIDATES = ["Eşleşme ID", "Eslesme ID", "Match ID", "Matching ID", "EslesmeId", "EşlesmeID"]
CHASSIS_CANDIDATES = ["Şasi", "Sasi", "Chassis", "VIN", "ŞASİ", "SASI"]

# Klasörler root altında değilse derin tarama (genelde gerek yok)
MAX_SCAN_DEPTH = 2  # senin ekranda direkt root altında olduğu için 1-2 yeter

# =======================
# HELPERS
# =======================
def norm_str(x) -> str:
    s = str(x).strip()
    return "" if s.lower() == "nan" else s

def norm_key(s: str) -> str:
    return norm_str(s).strip().upper()

def find_col(df, candidates):
    cols = {c.strip().lower(): c for c in df.columns}
    for cand in candidates:
        key = cand.strip().lower()
        if key in cols:
            return cols[key]
    # contains fallback
    for lc, orig in cols.items():
        for cand in candidates:
            if cand.strip().lower() in lc:
                return orig
    raise ValueError(f"Kolon bulunamadı. Adaylar: {candidates}\nMevcut kolonlar: {list(df.columns)}")

def list_all_files(folder: Path):
    # Uzantı filtresi yok: klasörde ne varsa kopyalar
    if not folder or not folder.exists():
        return []
    return [p for p in folder.rglob("*") if p.is_file()]

def safe_copy(src: Path, dst_dir: Path, prefix: str = ""):
    dst_dir.mkdir(parents=True, exist_ok=True)
    name = src.name
    if prefix:
        name = f"{prefix}__{name}"
    dst = dst_dir / name

    # çakışma olursa numaralandır
    if dst.exists():
        stem, suf = dst.stem, dst.suffix
        i = 2
        while True:
            cand = dst_dir / f"{stem}__{i}{suf}"
            if not cand.exists():
                dst = cand
                break
            i += 1

    shutil.copy2(src, dst)

def iter_dirs(root: Path, max_depth: int):
    root = root.resolve()
    stack = [(root, 0)]
    while stack:
        p, d = stack.pop()
        if not p.is_dir():
            continue
        yield p, d
        if d >= max_depth:
            continue
        try:
            for child in p.iterdir():
                if child.is_dir():
                    stack.append((child, d + 1))
        except PermissionError:
            pass

def build_folder_index(root: Path, max_depth: int):
    """
    returns: dict[normalized_folder_name] -> Path
    """
    idx = {}
    for folder, _depth in iter_dirs(root, max_depth):
        key = norm_key(folder.name)
        if key and key not in idx:
            idx[key] = folder
    return idx

# =======================
# MAIN
# =======================
def main():
    excel_path = Path(EXCEL_PATH)
    images_root = Path(IMAGES_ROOT)
    out_root = Path(OUT_ROOT)

    if not excel_path.exists():
        raise FileNotFoundError(f"Excel bulunamadı: {excel_path}")
    if not images_root.exists():
        raise FileNotFoundError(f"Görsel root bulunamadı: {images_root}")

    unique_df = pd.read_excel(excel_path, sheet_name=UNIQUE_SHEET)
    pipeline_df = pd.read_excel(excel_path, sheet_name=PIPELINE_SHEET)

    u_mid = find_col(unique_df, MATCH_ID_CANDIDATES)
    u_chs = find_col(unique_df, CHASSIS_CANDIDATES)

    p_mid = find_col(pipeline_df, MATCH_ID_CANDIDATES)
    p_chs = find_col(pipeline_df, CHASSIS_CANDIDATES)

    print("✅ Kolonlar bulundu:")
    print(f"  Unique:   match_id='{u_mid}', chassis='{u_chs}'")
    print(f"  Pipeline: match_id='{p_mid}', chassis='{p_chs}'")

    # ✅ Görsel klasör index'i: "TMB... -> path"
    folder_idx = build_folder_index(images_root, MAX_SCAN_DEPTH)

    # Unique: match_id -> [unique_chassis...]
    unique_map = {}
    unique_chassis_set = set()

    for _, r in unique_df.iterrows():
        mid = norm_str(r[u_mid])
        chs = norm_str(r[u_chs])
        if not mid or not chs:
            continue
        unique_map.setdefault(mid, []).append(chs)
        unique_chassis_set.add(norm_key(chs))

    # Pipeline: match_id -> [pipeline_chassis...]
    pipeline_map = {}
    for _, r in pipeline_df.iterrows():
        mid = norm_str(r[p_mid])
        chs = norm_str(r[p_chs])
        if not mid or not chs:
            continue
        pipeline_map.setdefault(mid, []).append(chs)

    # ✅ Unique şasi klasör eşleşmesi say
    hits = sum(1 for k in unique_chassis_set if k in folder_idx)
    print(f"🟢 Images Root: {images_root}")
    print(f"🟢 Unique şasi klasör eşleşmesi: {hits} / {len(unique_chassis_set)} (scan depth={MAX_SCAN_DEPTH})")

    # =======================
    # EK RAPORLAR: eksik klasörler
    # =======================
    root_folders = {norm_key(p.name) for p in images_root.iterdir() if p.is_dir()}
    missing_unique_chassis = sorted(unique_chassis_set - root_folders)
    extra_folders = sorted(root_folders - unique_chassis_set)

    out_root.mkdir(parents=True, exist_ok=True)
    (out_root / "missing_unique_chassis.txt").write_text("\n".join(missing_unique_chassis), encoding="utf-8")
    (out_root / "folders_not_in_unique.txt").write_text("\n".join(extra_folders), encoding="utf-8")

    # =======================
    # KOPYALAMA (match_id grup bazlı)
    # =======================
    report_rows = []
    total_copied = 0
    groups_copied = 0
    groups_no_source_folder = 0
    groups_no_files = 0
    groups_not_in_unique = 0

    for mid, pipeline_chassis_list in pipeline_map.items():
        unique_chassis_list = unique_map.get(mid)
        if not unique_chassis_list:
            groups_not_in_unique += 1
            report_rows.append({
                "match_id": mid,
                "status": "NO_MATCH_ID_IN_UNIQUE_SHEET",
                "source_chassis": "",
                "source_folder": "",
                "files_found": 0,
                "pipeline_chassis_count": len(pipeline_chassis_list),
                "files_copied_total": 0
            })
            continue

        # Bu match_id için, klasörü gerçekten bulunan ilk Unique şasiyi kaynak seç
        source_folder = None
        source_chassis = None

        for uch in unique_chassis_list:
            k = norm_key(uch)
            if k in folder_idx:
                source_folder = folder_idx[k]
                source_chassis = uch
                break

        if not source_folder:
            groups_no_source_folder += 1
            report_rows.append({
                "match_id": mid,
                "status": "MATCH_ID_FOUND_BUT_NO_SOURCE_FOLDER_FOR_ANY_UNIQUE_CHASSIS",
                "source_chassis": "",
                "source_folder": "",
                "files_found": 0,
                "pipeline_chassis_count": len(pipeline_chassis_list),
                "files_copied_total": 0
            })
            continue

        files = list_all_files(source_folder)
        if not files:
            groups_no_files += 1
            report_rows.append({
                "match_id": mid,
                "status": "SOURCE_FOLDER_EXISTS_BUT_EMPTY",
                "source_chassis": source_chassis,
                "source_folder": str(source_folder),
                "files_found": 0,
                "pipeline_chassis_count": len(pipeline_chassis_list),
                "files_copied_total": 0
            })
            continue

        # Aynı match_id grubundaki TÜM pipeline şasilerine aynı dosyaları kopyala
        copied_this_group = 0
        prefix = norm_key(source_chassis)

        for pchs in pipeline_chassis_list:
            dest_dir = out_root / norm_str(pchs)
            dest_dir.mkdir(parents=True, exist_ok=True)

            for f in files:
                safe_copy(f, dest_dir, prefix=prefix)
                copied_this_group += 1
                total_copied += 1

        groups_copied += 1
        report_rows.append({
            "match_id": mid,
            "status": "COPIED_TO_ALL_PIPELINE_CHASSIS_IN_GROUP",
            "source_chassis": source_chassis,
            "source_folder": str(source_folder),
            "files_found": len(files),
            "pipeline_chassis_count": len(pipeline_chassis_list),
            "files_copied_total": copied_this_group
        })

    # rapor kaydet
    report_path = out_root / "copy_report.xlsx"
    pd.DataFrame(report_rows).to_excel(report_path, index=False)

    print("\n✅ Bitti")
    print(f"📦 Toplam kopyalanan dosya: {total_copied}")
    print(f"✅ Kopyalanan match_id grubu: {groups_copied}")
    print(f"⚠️ Unique sheet'te olmayan match_id grubu: {groups_not_in_unique}")
    print(f"⚠️ Kaynak klasörü bulunamayan match_id grubu: {groups_no_source_folder}")
    print(f"⚠️ Kaynak klasörü boş olan match_id grubu: {groups_no_files}")
    print(f"📄 Rapor: {report_path}")
    print(f"🧾 Eksik klasör listesi: {out_root / 'missing_unique_chassis.txt'}")
    print(f"🧾 Fazla klasör listesi: {out_root / 'folders_not_in_unique.txt'}")

if __name__ == "__main__":
    main()
