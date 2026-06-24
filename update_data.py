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


def check_value_format(val_str, row_label, period):
    """
    Cek apakah sebuah nilai sel terlihat seperti angka/persen yang valid.
    Tidak menolak teks lain (beberapa parameter mungkin sengaja kualitatif),
    tapi nilai yang KELIHATAN seperti angka rusak (mis. "Rp 1.000.000",
    "1,234.56", sel berisi formula error "#REF!") akan di-flag.
    """
    if val_str is None:
        return
    s = val_str.strip()
    if s == "":
        return
    if RE_NUMERIC.match(s) or RE_PERCENT.match(s):
        return
    # Pola yang sering jadi jebakan: format Rupiah, koma desimal gaya US,
    # atau error formula Excel yang ke-paste sebagai teks.
    if re.match(r"^[A-Z#].*$", s) and ("#" in s or "REF" in s or "DIV" in s or "N/A" in s.upper()):
        warn(f"'{row_label}' periode {period}: sel berisi kemungkinan ERROR FORMULA Excel → \"{s}\"")
    elif re.search(r"rp\s*\d|\d.*\.\d{3}\.\d{3}", s, re.IGNORECASE):
        warn(f"'{row_label}' periode {period}: format mencurigakan (mirip Rupiah/ribuan salah format) → \"{s}\"")
    elif re.match(r"^-?\d+\.\d+$", s):
        # angka dengan titik desimal gaya US di tengah data yang lain pakai koma
        warn(f"'{row_label}' periode {period}: pakai titik desimal (gaya US) sedangkan parameter lain pakai koma → \"{s}\"")
    # Selain itu dianggap teks kualitatif yang sah, tidak di-warn.


def read_datamart(wb):
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

        row = {
            "key": key,
            "bag": str(ws.cell(row=er, column=2).value or "").strip(),
            "par": par_clean,
            "jen": jen,
            "lnk": str(ws.cell(row=er, column=5).value or "").strip(),
        }
        for p in periods:
            val = ws.cell(row=er, column=period_cols[p]).value
            val_str = None if (val is None or str(val).strip() == "") else str(val).strip()
            check_value_format(val_str, par_clean, p)
            row[p] = val_str
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
    dm      = read_datamart(wb)
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
