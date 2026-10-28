"""
update_data.py
--------------
Baca Datamart_Risiko_Operasional.xlsx (sheet DATAMART + LAPORAN)
→ generate public/data.js untuk portal web.

Jalankan setiap kali selesai update Excel, lalu git push ke GitHub.
Vercel otomatis re-deploy dalam ~30 detik.

Usage:
    python update_data.py            # jalan normal, warning ditampilkan tapi tetap lanjut
    python update_data.py --strict   # berhenti (exit code 1) kalau ada warning data
"""

import json, sys, re, argparse
from pathlib import Path
from datetime import datetime
from openpyxl import load_workbook

BASE       = Path(__file__).parent
EXCEL_FILE = BASE / "Datamart_Risiko_Operasional.xlsx"
OUT_FILE   = BASE / "public" / "data.js"

# Kolom yang isinya WAJIB numerik atau persen (selain teks bebas).
# "jen" boleh cuma salah satu dari dua nilai ini — typo di sini bikin
# filter "Hasil"/"Pembentuk" di web diam-diam gagal.
JENIS_VALID = {"Hasil", "Pembentuk"}

# Pola nilai yang dianggap "bersih": angka biasa, atau persen format Indonesia (1.234,56%)
RE_NUMERIC = re.compile(r"^-?\d{1,3}(\.\d{3})*(,\d+)?$|^-?\d+(,\d+)?$")
RE_PERCENT = re.compile(r"^-?\d{1,3}(\.\d{3})*(,\d+)?%$|^-?\d+(,\d+)?%$")

WARNINGS = []


def warn(msg):
    WARNINGS.append(msg)
    print(f"  ⚠ {msg}")


RE_RAW_PERCENT_CANDIDATE = re.compile(r"^-?0?\.\d+$|^-?1\.0+$")  # 0.xxxx atau 1.0 (gaya US, antara 0 dan 1)
RE_US_DECIMAL = re.compile(r"^-?\d+\.\d+$")
RE_URL = re.compile(r"^https?://", re.IGNORECASE)

# Arah "baik" tiap indikator Hasil, dikonfirmasi pemilik portal 27 Okt 2026.
# Dipakai HANYA kalau sheet DATAMART belum punya kolom "Arah" atau selnya kosong.
# naik  = nilai naik itu membaik (hijau ▲), turun = nilai turun itu membaik (hijau ▼),
# target = mendekati pita hijau tua di sheet THRESHOLD itu membaik (ambang dua sisi).
ARAH_DEFAULT = {
    "OPS-00000001": "target", # NOC/AO ≥ 1 Bulan
    "OPS-00000004": "target", # Kelompok/AO ≥ 1 Bulan
    "OPS-00000007": "turun",  # Turn Over AO Mekaar 12 Bulan Terakhir
    "OPS-00000010": "turun",  # AO Mekaar Masa Kerja ≤ 3 Bulan
    "OPS-00000013": "naik",   # PKM Full Payment
    "OPS-00000014": "naik",   # PKM Face Detect
    "OPS-00000015": "naik",   # Kelengkapan KTP & FP4
    "OPS-00000018": "naik",   # Identitas DRTD Yang Tervalidasi
    "OPS-00000021": "turun",  # Anggota Kelompok Tidak Sesuai
    "OPS-00000024": "naik",   # Daily Closing Mekaar H+0
    "OPS-00000027": "turun",  # Temuan Mekaar Belum Tindak Lanjut
    "OPS-00000030": "turun",  # % Total Klaim CIS CIT
    "OPS-00000034": "turun",  # % Pencairan diluar hari PKM
}
ARAH_VALID = ("naik", "turun", "target", "netral")

# Ambang status (RAG) per indikator Hasil, dari dashboard KRI yang diberikan pemilik portal 28 Okt 2026.
# Dipakai HANYA kalau sheet THRESHOLD belum ada atau tidak punya baris untuk Primary Key tersebut.
# Notasi sama dengan dashboard: "x ≤ 10%", "300 < x ≤ 340", "60% ≥ x > 50%", "x = 0%".
# Beberapa rentang untuk satu warna dipisah titik koma (ambang dua sisi).
LEVELS = ("hijau_tua", "hijau", "kuning", "oranye", "merah")
LEVEL_HEADERS = {"hijau tua": "hijau_tua", "hijau": "hijau", "kuning": "kuning", "oranye": "oranye", "merah": "merah"}
_TURN_OVER = {"hijau_tua": "x ≤ 10%", "hijau": "10% < x ≤ 20%", "kuning": "20% < x ≤ 35%",
              "oranye": "35% < x ≤ 50%", "merah": "x > 50%"}
_HADIR_PKM = {"hijau_tua": "x > 60%", "hijau": "60% ≥ x > 50%", "kuning": "50% ≥ x > 40%",
              "oranye": "40% ≥ x > 30%", "merah": "x ≤ 30%"}
THRESHOLD_DEFAULT = {
    "OPS-00000001": {"hijau_tua": "340 < x ≤ 380", "kuning": "300 < x ≤ 340; 380 < x ≤ 420",
                     "merah": "x ≤ 300; x > 420"},
    "OPS-00000004": {"hijau_tua": "17 ≤ x ≤ 23", "kuning": "12 < x < 17; 23 < x ≤ 28",
                     "merah": "x ≤ 12; x > 28"},
    "OPS-00000007": _TURN_OVER,
    "OPS-00000010": _TURN_OVER,
    "OPS-00000013": _HADIR_PKM,
    "OPS-00000014": _HADIR_PKM,
    "OPS-00000021": {"hijau_tua": "x = 0%", "hijau": "0% < x ≤ 3%", "kuning": "3% < x ≤ 6%",
                     "oranye": "6% < x ≤ 9%", "merah": "x > 9%"},
    "OPS-00000024": {"hijau_tua": "x = 100%", "kuning": "100% > x ≥ 70%", "merah": "x < 70%"},
    # = "% Belum TLHP (Mekaar)" di dashboard
    "OPS-00000027": {"hijau_tua": "x = 0%", "hijau": "0% < x ≤ 5%", "kuning": "5% < x ≤ 8%",
                     "oranye": "8% < x ≤ 10%", "merah": "x > 10%"},
}

_NUM = r"-?\d+(?:[.,]\d+)?%?"
_OP = r"<=|>=|≤|≥|<|>|="
RE_COND_RANGE = re.compile(rf"^({_NUM})\s*({_OP})\s*x\s*({_OP})\s*({_NUM})$")
RE_COND_X_LEFT = re.compile(rf"^x\s*({_OP})\s*({_NUM})$")
RE_COND_X_RIGHT = re.compile(rf"^({_NUM})\s*({_OP})\s*x$")
_FLIP = {"<": ">", ">": "<", "≤": "≥", "≥": "≤", "=": "="}
EPS = 1e-9  # toleransi float, mis. 0.03 dari formula vs ambang 3%


def _norm_op(op):
    return {"<=": "≤", ">=": "≥"}.get(op, op)


def _bound(op, num):
    """Interval untuk "x <op> num"."""
    if op == "<":
        return {"lo": None, "loInc": False, "hi": num, "hiInc": False}
    if op == "≤":
        return {"lo": None, "loInc": False, "hi": num, "hiInc": True}
    if op == ">":
        return {"lo": num, "loInc": False, "hi": None, "hiInc": False}
    if op == "≥":
        return {"lo": num, "loInc": True, "hi": None, "hiInc": False}
    return {"lo": num, "loInc": True, "hi": num, "hiInc": True}


def parse_condition(text):
    """
    "300 < x ≤ 340" → ({"lo":300,"loInc":False,"hi":340,"hiInc":True}, False, label).
    Angka ber-% dibagi 100 supaya sebanding dengan nilai pecahan di DATAMART.
    Mengembalikan None kalau notasi tidak dikenali.
    """
    label = text.strip().replace("X", "x").replace("<=", "≤").replace(">=", "≥")
    label = re.sub(r"\s+", " ", re.sub(r"\s*([≤≥<>=])\s*", r" \1 ", label)).strip()
    s = label.replace(" ", "")

    def num(tok):
        v = float(tok.rstrip("%").replace(",", "."))
        return (v / 100, True) if tok.endswith("%") else (v, False)

    m = RE_COND_RANGE.match(s)
    if m:
        (a, pa), (b, pb) = num(m.group(1)), num(m.group(4))
        left = _bound(_FLIP[_norm_op(m.group(2))], a)   # "a < x" ≡ "x > a"
        right = _bound(_norm_op(m.group(3)), b)
        iv = {"lo": left["lo"] if left["lo"] is not None else right["lo"],
              "loInc": left["loInc"] if left["lo"] is not None else right["loInc"],
              "hi": right["hi"] if right["hi"] is not None else left["hi"],
              "hiInc": right["hiInc"] if right["hi"] is not None else left["hiInc"]}
        return iv, pa or pb, label
    m = RE_COND_X_LEFT.match(s)
    if m:
        a, pa = num(m.group(2))
        return _bound(_norm_op(m.group(1)), a), pa, label
    m = RE_COND_X_RIGHT.match(s)
    if m:
        a, pa = num(m.group(1))
        return _bound(_FLIP[_norm_op(m.group(2))], a), pa, label
    return None


def in_interval(v, iv):
    if iv["lo"] is not None and (v < iv["lo"] - EPS if iv["loInc"] else v <= iv["lo"] + EPS):
        return False
    if iv["hi"] is not None and (v > iv["hi"] + EPS if iv["hiInc"] else v >= iv["hi"] - EPS):
        return False
    return True


def parse_bands(spec, sat, row_label):
    """spec: {level: "kondisi; kondisi"} → list pita {lvl, label, lo, loInc, hi, hiInc}."""
    bands = []
    for lvl in LEVELS:
        for part in str(spec.get(lvl) or "").split(";"):
            if not part.strip():
                continue
            parsed = parse_condition(part)
            if not parsed:
                warn(f"'{row_label}': ambang {lvl.replace('_', ' ')} \"{part.strip()}\" tidak dikenali — "
                     f"pakai notasi seperti \"x ≤ 10%\" atau \"300 < x ≤ 340\".")
                continue
            iv, has_pct, label = parsed
            if has_pct != (sat == "%"):
                warn(f"'{row_label}': ambang \"{label}\" {'pakai' if has_pct else 'tanpa'} tanda %, "
                     f"padahal satuan indikator ini {sat}.")
            bands.append({"lvl": lvl, "label": label, **iv})
    return bands


def classify(v, bands):
    """Level pita yang memuat nilai v. Mengembalikan (level|None, daftar level yang cocok)."""
    hits = [b["lvl"] for b in bands if in_interval(v, b)]
    return (hits[0] if hits else None), hits


def target_band(bands):
    """Gabungan rentang level terbaik (biasanya hijau tua) — acuan arah 'target'."""
    best = [b for b in bands if b["lvl"] == bands[0]["lvl"]]
    los = [b["lo"] for b in best]
    his = [b["hi"] for b in best]
    return {"lo": None if None in los else min(los), "hi": None if None in his else max(his)}


def legend_order(bands, arah):
    """Urutan chip legenda mengikuti nilai: naik dari kecil, 'naik = baik' dibalik (terbaik di kiri)."""
    key = lambda b: (float("-inf") if b["lo"] is None else b["lo"], 0 if b["loInc"] else 1)
    return sorted(bands, key=key, reverse=(arah == "naik"))


def looks_like_unformatted_percent(s):
    """Angka desimal titik antara -1 dan 1 (mis. 0.4344) — kandidat persen yang belum diformat."""
    return bool(RE_RAW_PERCENT_CANDIDATE.match(s))


def parse_cell(val):
    """
    Ubah nilai sel jadi angka (float), atau None kalau bukan angka.
    - Sel numerik Excel dipakai apa adanya (presisi penuh, mis. 0.4434818141530953).
    - Teks persen Indonesia "56,77%" → 0.5677.
    - Teks angka Indonesia "1.234,56" atau desimal gaya US "405.99" → float.
    - Teks lain (kualitatif) → None, ditampilkan apa adanya.
    """
    if isinstance(val, bool):
        return None
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip()
    if RE_PERCENT.match(s):
        return float(s[:-1].replace(".", "").replace(",", ".")) / 100
    if RE_NUMERIC.match(s):
        return float(s.replace(".", "").replace(",", "."))
    if RE_US_DECIMAL.match(s):
        return float(s)
    return None


def fmt_id(x, decimals):
    """Format angka gaya Indonesia: 15710335 → "15.710.335", 407.7217 → "407,72"."""
    s = f"{x:,.{decimals}f}"
    return s.replace(",", "\x00").replace(".", ",").replace("\x00", ".")


def format_display(num, sat):
    if sat == "%":
        return fmt_id(num * 100, 2) + "%"
    if sat == "rasio" or num != int(num):
        return fmt_id(num, 2)
    return fmt_id(num, 0)


def detect_satuan(jen, cells):
    """
    cells: list (val, number_format) untuk semua periode di satu baris.
    "%"     → ada sel berformat persen di Excel (mis. "0.00%") atau teks "xx,xx%";
              sel numerik di baris ini disimpan sebagai pecahan 0–1.
    "rasio" → indikator Hasil non-persen (mis. NOC/AO), selalu 2 desimal.
    "angka" → selain itu (Pembentuk).
    """
    for val, nf in cells:
        if val is None:
            continue
        if isinstance(val, (int, float)) and "%" in (nf or ""):
            return "%"
        if isinstance(val, str) and RE_PERCENT.match(val.strip()):
            return "%"
    return "rasio" if jen == "Hasil" else "angka"


def find_header_col(ws, name):
    """Cari kolom berdasarkan teks header (baris 2 atau 3), tidak peka huruf besar/kecil."""
    for r in (2, 3):
        for c in range(1, ws.max_column + 1):
            v = ws.cell(row=r, column=c).value
            if v is not None and str(v).strip().lower() == name.lower():
                return c
    return None


def read_source(cell):
    """
    Pisahkan nama sumber dan URL sumber dari kolom Link to Source.
    URL hanya diisi kalau benar-benar absolut (http/https) — baik dari teks sel
    maupun dari hyperlink yang tertanam di sel. Hyperlink relatif (mis. path
    SharePoint "../../:x:/r/sites/...") tidak bisa dibuka dari portal, jadi diabaikan.
    """
    text = str(cell.value or "").strip()
    target = (cell.hyperlink.target or "").strip() if cell.hyperlink else ""
    if RE_URL.match(text):
        return "", text
    return text, target if RE_URL.match(target) else ""


def check_value_format(val_str, row_label, period):
    """
    Cek apakah sebuah nilai sel terlihat seperti angka/persen yang valid.
    Tidak menolak teks lain (beberapa parameter mungkin sengaja kualitatif),
    tapi nilai yang KELIHATAN seperti angka rusak (mis. "Rp 1.000.000",
    "1.234.567" ribuan salah format dengan banyak titik, sel error formula
    "#REF!") akan di-flag.

    Catatan: angka dengan SATU titik desimal (gaya US) — baik desimalnya
    panjang seperti "405.9992105263158" (hasil formula) maupun pendek
    seperti "41921.5" atau "15.95" — TIDAK di-warning. Keduanya valid
    secara matematis; satu titik desimal satu-satunya jarang merupakan
    typo. Yang tetap di-warning hanya pola dengan BANYAK titik dalam satu
    angka (mis. "1.234.567"), yang merupakan ciri khas typo pemisah ribuan
    gaya Indonesia yang salah ketik di luar konteksnya.

    Desimal antara -1 dan 1 (mis. "0.4344") yang merupakan kandidat
    persentase belum diformat juga tidak di-warning di sini — satuannya
    ditentukan di detect_satuan().
    """
    if val_str is None:
        return
    s = val_str.strip()
    if s == "":
        return
    if RE_NUMERIC.match(s) or RE_PERCENT.match(s):
        return
    if looks_like_unformatted_percent(s):
        return
    if re.match(r"^[A-Z#].*$", s) and ("#" in s or "REF" in s or "DIV" in s or "N/A" in s.upper()):
        warn(f"'{row_label}' periode {period}: sel berisi kemungkinan ERROR FORMULA Excel → \"{s}\"")
        return
    if re.search(r"rp\s*\d|\d.*\.\d{3}\.\d{3}", s, re.IGNORECASE):
        warn(f"'{row_label}' periode {period}: format mencurigakan (mirip Rupiah/ribuan salah format) → \"{s}\"")
        return
    # Satu titik desimal (gaya US), berapapun panjang desimalnya, dianggap valid — tidak di-warning.
    # Selain itu dianggap teks kualitatif yang sah, tidak di-warn.


def read_thresholds(wb):
    """
    Sheet THRESHOLD (opsional): header "Primary Key", "Hijau Tua", "Hijau", "Kuning", "Oranye", "Merah"
    di salah satu dari 5 baris pertama; satu baris per indikator. Sel kosong = warna itu tidak dipakai.
    """
    if "THRESHOLD" not in wb.sheetnames:
        return {}
    ws = wb["THRESHOLD"]
    for hr in range(1, 6):
        heads = {str(ws.cell(row=hr, column=c).value or "").strip().lower(): c
                 for c in range(1, ws.max_column + 1)}
        if "primary key" in heads:
            break
    else:
        warn("Sheet THRESHOLD ada tapi header 'Primary Key' tidak ditemukan di 5 baris pertama — sheet diabaikan.")
        return {}
    level_cols = {LEVEL_HEADERS[h]: c for h, c in heads.items() if h in LEVEL_HEADERS}
    out = {}
    for er in range(hr + 1, ws.max_row + 1):
        key = str(ws.cell(row=er, column=heads["primary key"]).value or "").strip()
        if key:
            out[key] = {lvl: str(ws.cell(row=er, column=c).value or "").strip() for lvl, c in level_cols.items()}
    print(f"  THRESHOLD: {len(out)} indikator dari sheet THRESHOLD")
    return out


def read_datamart(wb, thr_xl):
    if "DATAMART" not in wb.sheetnames:
        print("ERROR: Sheet 'DATAMART' tidak ditemukan di file Excel.")
        sys.exit(1)
    ws = wb["DATAMART"]
    row3 = [ws.cell(row=3, column=c).value for c in range(1, ws.max_column + 1)]
    periods, period_cols = [], {}
    for ci, val in enumerate(row3):
        s = str(val).strip() if val is not None else ""
        if re.match(r"^\d{6}$", s):
            periods.append(s)
            period_cols[s] = ci + 1
    if not periods:
        print("ERROR: Tidak ada kolom periode (format YYYYMM, contoh 202604) di baris 3 sheet DATAMART.")
        sys.exit(1)
    periods.sort()
    arah_col = find_header_col(ws, "Arah")
    if arah_col:
        print(f"  DATAMART : kolom Arah ditemukan di kolom {arah_col}")

    seen_keys = set()
    rows = []
    for er in range(5, ws.max_row + 1):
        par = ws.cell(row=er, column=3).value
        if not par or str(par).strip() == "":
            continue
        key = str(ws.cell(row=er, column=1).value or "").strip()
        jen = str(ws.cell(row=er, column=4).value or "").strip()
        par_clean = str(par).strip()

        if not key:
            warn(f"Baris {er} ('{par_clean}'): kolom KEY kosong — baris ini tidak punya ID unik.")
        elif key in seen_keys:
            warn(f"Baris {er} ('{par_clean}'): KEY '{key}' duplikat — sudah dipakai parameter lain.")
        else:
            seen_keys.add(key)

        if jen and jen not in JENIS_VALID:
            warn(f"Baris {er} ('{par_clean}'): kolom JENIS berisi '{jen}', diharapkan 'Hasil' atau 'Pembentuk' — cek typo.")

        arah = ""
        if jen == "Hasil":
            arah_xl = str(ws.cell(row=er, column=arah_col).value or "").strip().lower() if arah_col else ""
            if arah_xl and arah_xl not in ARAH_VALID:
                warn(f"Baris {er} ('{par_clean}'): kolom ARAH berisi '{arah_xl}', diharapkan Naik, Turun, Target, atau Netral.")
            arah = arah_xl if arah_xl in ARAH_VALID else ARAH_DEFAULT.get(key, "")
            if not arah:
                warn(f"Baris {er} ('{par_clean}'): indikator Hasil tanpa ARAH — isi kolom Arah dengan Naik, Turun, Target, atau Netral.")
                arah = "netral"

        src, url = read_source(ws.cell(row=er, column=5))
        cells = []
        for p in periods:
            cell = ws.cell(row=er, column=period_cols[p])
            val = cell.value
            if isinstance(val, str) and val.strip() == "":
                val = None
            check_value_format(None if val is None else str(val).strip(), par_clean, p)
            cells.append((val, cell.number_format))
        sat = detect_satuan(jen, cells)

        row = {
            "key": key,
            "bag": str(ws.cell(row=er, column=2).value or "").strip(),
            "par": par_clean,
            "jen": jen,
            "arah": arah,
            "sat": sat,
            "src": src,
            "url": url,
        }
        nums = {}
        for p, (val, _) in zip(periods, cells):
            if val is None:
                row[p], nums[p] = None, None
                continue
            # Di baris persen, angka dari Excel (atau teks "0.4344") adalah pecahan 0–1.
            num = parse_cell(val)
            if num is None:
                row[p], nums[p] = str(val).strip(), None
                continue
            row[p], nums[p] = format_display(num, sat), num
        row["num"] = nums

        spec = (thr_xl.get(key) or THRESHOLD_DEFAULT.get(key)) if jen == "Hasil" else None
        bands = parse_bands(spec, sat, par_clean) if spec else []
        if row["arah"] == "target" and not bands:
            warn(f"Baris {er} ('{par_clean}'): ARAH Target butuh ambang di sheet THRESHOLD — sementara dianggap Netral.")
            row["arah"] = "netral"
        if bands:
            row["thr"] = legend_order(bands, row["arah"])
            row["tgt"] = target_band(bands)
            row["st"] = {}
            for p in periods:
                if nums[p] is None:
                    row["st"][p] = None
                    continue
                lvl, hits = classify(nums[p], bands)
                if not hits:
                    warn(f"'{par_clean}' periode {p}: nilai {row[p]} tidak masuk pita ambang mana pun — cek celah di ambang.")
                elif len(set(hits)) > 1:
                    warn(f"'{par_clean}' periode {p}: nilai {row[p]} masuk lebih dari satu warna ({', '.join(hits)}) — ambang tumpang tindih.")
                row["st"][p] = lvl
        rows.append(row)

    print(f"  DATAMART : {len(rows)} parameter, {len(periods)} periode")
    return {"periods": periods, "rows": rows}


def read_laporan(wb):
    if "LAPORAN" not in wb.sheetnames:
        print("  LAPORAN  : sheet tidak ditemukan, dilewati.")
        return []
    ws = wb["LAPORAN"]
    files = []
    for er in range(5, ws.max_row + 1):
        nama = ws.cell(row=er, column=1).value
        if not nama or str(nama).strip() == "":
            continue
        aktif = str(ws.cell(row=er, column=8).value or "Y").strip().upper()
        if aktif != "Y":
            continue
        link = str(ws.cell(row=er, column=5).value or "").strip()
        tipe = str(ws.cell(row=er, column=4).value or "pdf").strip().lower()
        if not link:
            warn(f"Baris {er} LAPORAN ('{str(nama).strip()}'): kolom LINK kosong — kartu dokumen ini tidak akan bisa dibuka.")
        if tipe != "link" and link.startswith("http"):
            warn(f"Baris {er} LAPORAN ('{str(nama).strip()}'): tipe '{tipe}' tapi link berupa URL ('{link}') — mungkin maksudnya tipe 'link'?")
        files.append({
            "nama":  str(nama).strip(),
            "desc":  str(ws.cell(row=er, column=2).value or "").strip(),
            "kat":   str(ws.cell(row=er, column=3).value or "Dokumen").strip(),
            "tipe":  tipe,
            "link":  link,
            "ukuran":str(ws.cell(row=er, column=6).value or "").strip(),
            "tgl":   str(ws.cell(row=er, column=7).value or "").strip(),
        })
    print(f"  LAPORAN  : {len(files)} dokumen aktif")
    return files


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--strict", action="store_true",
                         help="Berhenti dengan exit code 1 jika ada warning data (cocok untuk CI/GitHub Actions).")
    args = parser.parse_args()

    if not EXCEL_FILE.exists():
        print(f"ERROR: File tidak ditemukan → {EXCEL_FILE}")
        sys.exit(1)

    print(f"Membaca {EXCEL_FILE.name}...")
    wb = load_workbook(EXCEL_FILE, data_only=True)
    dm      = read_datamart(wb, read_thresholds(wb))
    laporan = read_laporan(wb)
    ts      = datetime.now().strftime("%Y-%m-%d %H:%M")

    payload = {
        "periods": dm["periods"],
        "rows":    dm["rows"],
        "laporan": laporan,
        "updated": ts,
        "warnings": WARNINGS,   # ikut disimpan supaya bisa ditampilkan sebagai indikator di portal
    }
    OUT_FILE.parent.mkdir(parents=True, exist_ok=True)
    js = (
        f"// Auto-generated by update_data.py — {ts}\n"
        f"// Jangan edit manual. Edit Excel lalu jalankan script ini.\n"
        f"window.DATAMART = {json.dumps(payload, ensure_ascii=False, indent=2)};\n"
    )
    OUT_FILE.write_text(js, encoding="utf-8")

    print(f"\n✓ Berhasil → {OUT_FILE}")
    print(f"  {len(dm['rows'])} parameter · {len(dm['periods'])} periode · {len(laporan)} dokumen")

    if WARNINGS:
        print(f"\n⚠ {len(WARNINGS)} WARNING ditemukan (lihat detail di atas).")
        if args.strict:
            print("✗ Mode --strict aktif → berhenti. Perbaiki Excel lalu jalankan ulang.")
            sys.exit(1)
        else:
            print("  Tetap lanjut (mode normal). Jalankan dengan --strict untuk memblokir deploy jika ada warning.")
    else:
        print("  Tidak ada warning data. ✓")

    print(f"\nSelanjutnya:")
    print(f"  git add public/data.js Datamart_Risiko_Operasional.xlsx")
    print(f"  git commit -m 'update data {dm['periods'][-1]}'")
    print(f"  git push")


if __name__ == "__main__":
    main()
