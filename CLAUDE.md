# CLAUDE.md — Portal Risiko Operasional (Datamart-Riskops)

Konteks handoff dari sesi review sebelumnya (27 Okt 2026). Baca ini dulu sebelum ngubah kode.

## Apa ini

Portal self-service KRI Risiko Operasional bisnis **Mekaar** (AO, NOC, PKM, kelompok).
Pipeline statis, tanpa backend:

```
Datamart_Risiko_Operasional.xlsx  (sheet PANDUAN, DATAMART, LAPORAN)
        │  push ke GitHub
        ▼
.github/workflows/update-data.yml → python update_data.py --strict
        │  validasi + normalisasi → commit public/data.js
        ▼
Vercel (vercel.json, static) → public/index.html membaca window.DATAMART dari data.js
```

- **DATAMART**: header periode YYYYMM di baris 3, data mulai baris 5. Kolom: A=Primary Key, B=Bagian, C=Parameter, D=Jenis (Hasil/Pembentuk), E=Link Sumber, F..=periode. Saat ini 36 parameter × 25 periode (202405–202605).
- **LAPORAN**: data mulai baris 5. Kolom: Nama, Deskripsi, Kategori, Tipe, Link, Ukuran, Tanggal, Aktif (Y/N).
- Tiap indikator **Hasil** = rasio dari beberapa baris **Pembentuk** (pembilang/penyebut) di bawahnya. Sudah dicek konsisten, mis. NOC/AO 202605 = 15.710.335 / 38.532 = 407,72.
- `public/data.js` AUTO-GENERATED — jangan edit manual.
- Portal punya 3 tab: Pencarian Cepat (keyword matching, BUKAN AI), Datamart (kartu ringkasan + tabel + filter), Laporan & Dokumen.

## ⚠️ Keamanan (prioritas nomor satu)

- README bilang repo **private**, faktanya repo **PUBLIC** (bisa di-clone tanpa login). Excel, riwayat git, dan `data.js` terbuka untuk umum.
- `GATE_PASSWORD` di-hardcode plaintext di `public/index.html` (~baris 614) → gate tidak ada nilainya.
- Langkah: set repo private → ganti password → aktifkan Vercel Deployment Protection / auth di edge → lapor ke atasan/IT security kalau data sensitif. Privat-in repo saja TIDAK menutup akses `data.js` via URL Vercel.

## Bug yang sudah terverifikasi (urut prioritas)

1. **Tab Laporan & Dokumen mati total.** 9 dokumen menunjuk `files/...` tapi folder `public/files/` tidak ada. Dokumen ke-10 menunjuk `sharepoint.perusahaan.com` (placeholder). Isi LAPORAN masih dummy.
2. **Data basi.** Update terakhir 24 Jun 2026, periode terbaru 202605. Juni–September belum masuk. 3 dari 12 indikator Hasil kosong di periode terakhir: Identitas DRTD Yang Tervalidasi, % Total Klaim CIS CIT, % Pencairan diluar hari PKM.
3. **Warna tren terbalik untuk indikator "naik = baik".** `renderCards()` (~baris 425) dan `query()` (~baris 563) selalu memakai naik = merah ▲, turun = hijau ▼. Benar untuk Turn Over, tapi salah untuk PKM Face Detect, Kelengkapan KTP & FP4, Daily Closing H+0, dan sejenisnya. Perlu metadata **Arah** per indikator.
4. **Rasio tampil sebagai desimal mentah + presisi hilang.** Nilai seperti `0.44`, `0.43`, `0.99`. Konversi persen di `normalize_row_values()` hanya aktif kalau baris itu sudah punya riwayat format `"xx,xx%"`, dan di data saat ini tidak ada satu baris pun → fitur mati. `normalize_value()` membulatkan ke 2 desimal → 0,4344 jadi 0,43. Perlu metadata **Satuan** (%, rasio, angka) dan presisi yang cukup untuk perbandingan dengan threshold.
5. **Tombol saran Pencarian Cepat rusak.**
   - `kwds()` (~baris 588) membuang karakter non-alfanumerik: "NOC/AO" → `nocao` → 0 hasil.
   - Kata "total" masuk stopword: "Total AO Mekaar" skornya seri dengan "Turn Over AO Mekaar".
   - Token pendek (`ao`) dicocokkan sebagai substring → banyak false match.
6. **Kolom Sumber jadi link mati.** 26 baris punya `lnk` berisi nama sistem (mis. "Indikator Dashboard Register LI"), bukan URL, tapi tetap dirender sebagai `<a href>` (~baris 451). Pisahkan "nama sumber" dan "URL sumber", render link hanya kalau `^https?://`.
7. **Kemungkinan salah label di Excel.** "AO Mekaar Masa Kerja ≤ 3 Bulan" = 9.446/43.215 = 22%, tapi pembentuknya berlabel "Jml AO Masa Kerja **≥** 3 Bulan". Hampir pasti seharusnya ≤ — konfirmasi ke user/sumber. "Jml AO Masa Kerja ≥ 1 Bulan" juga muncul dua kali (dipakai dua rasio; pertimbangkan dedup/relasi).
8. **Minor:**
   - Timestamp memakai UTC: `datetime.now()` (`update_data.py` ~baris 255) di GitHub Actions → portal salah 7 jam. Pakai `zoneinfo("Asia/Jakarta")`.
   - Validasi tidak menangkap indikator Hasil kosong di periode terbaru atau kolom periode baru yang belum diisi → CI tetap hijau.
   - JS parse angka dengan `replace(/[%,]/g,'')` → akan salah 100× untuk format Indonesia "43,44%". Wajib diperbaiki bersamaan dengan poin 4.
   - `bag`, `lnk`, dan nilai sel masuk ke HTML tanpa `esc()` (`fillFilters`, `renderTbl`, `tblR`, `renderLap` onclick).
   - Vercel deploy 2× per update (push Excel + push data.js). Pertimbangkan ignored build step atau deploy hanya dari commit bot.
   - Kartu ringkasan hanya mengambil 4 indikator Hasil pertama sesuai urutan Excel, bukan berdasarkan prioritas/status.
   - Badge peringatan di portal praktis tidak pernah tampil karena CI selalu `--strict`.

## Gap fungsi (kandidat upgrade)

Portal saat ini = katalog angka + pencarian, belum alat pemantauan risiko. Urutan nilai per effort:

1. **Threshold + status RAG** (hijau/kuning/merah) per indikator Hasil — inti dari KRI.
2. **Grafik tren** (sekarang hanya tabel).
3. **Drill-down Hasil → Pembentuk** (apakah perubahan karena pembilang atau penyebut).
4. **Ekspor** CSV/Excel.
5. Pecahan per regional/cabang (butuh perubahan struktur data, effort besar).

Saran implementasi: tambah kolom metadata di sheet DATAMART (mis. `Arah`, `Satuan`, `Threshold Kuning`, `Threshold Merah`) supaya tetap bisa dikelola user non-teknis via Excel. Update validasi di `update_data.py` agar kolom baru tidak terbaca sebagai periode (deteksi periode hanya regex `^\d{6}$`, jadi aman selama header metadata bukan 6 digit).

## Status perbaikan

- **Bug 3–6 sudah dikerjakan (27 Okt 2026, belum di-commit/deploy saat ditulis).** `data.js` sekarang punya per baris: `arah` (naik/turun/netral, hanya Hasil), `sat` (%/rasio/angka, dibaca dari number format sel Excel), `num` (nilai presisi penuh per periode), `src` + `url` (menggantikan `lnk`). Nilai tampilan diformat gaya Indonesia ("44,35%", "15.710.335").
- Arah dibaca dari kolom header `Arah` di sheet DATAMART kalau ada; kalau belum ada atau selnya kosong, pakai `ARAH_DEFAULT` di `update_data.py` (jawaban pemilik 27–28 Okt 2026: NOC/AO & Kelompok/AO = target (ambang dua sisi, mendekati pita hijau tua = baik); Turn Over, AO ≤3 Bln, Anggota Tidak Sesuai, Temuan, Klaim CIS CIT, Pencairan diluar PKM = turun baik; Full Payment, Face Detect, KTP & FP4, DRTD, DC H+0 = naik baik).
- **Threshold + status RAG sudah dikerjakan (28 Okt 2026).** Ambang dari dashboard KRI pemilik (5 warna: hijau tua, hijau, kuning, oranye, merah). Sumber: sheet opsional `THRESHOLD` (header `Primary Key`, `Hijau Tua`, `Hijau`, `Kuning`, `Oranye`, `Merah`; isi notasi seperti `300 < x ≤ 340`, beberapa rentang dipisah `;`), fallback `THRESHOLD_DEFAULT`. `data.js` Hasil punya `thr` (pita urut legenda), `tgt` (pita target), `st` (status per periode). Ambang tumpang tindih / nilai jatuh di celah → warning (CI gagal). Celah hanya terdeteksi kalau ada nilai yang jatuh di sana.
- Belum ada ambang: Kelengkapan KTP & FP4, Identitas DRTD, % Klaim CIS CIT, % Pencairan diluar hari PKM (tampil abu-abu). "Temuan Mekaar Belum Tindak Lanjut" = "% Belum TLHP (Mekaar)" di dashboard (dikonfirmasi pemilik).
- Di dashboard tapi belum ada baris di DATAMART: Turn Over AO Monthly, % Belum TLHP ULaMM, % Pendingan Pekerjaan Notaris, dan 7 KRI IT (Pentest, Storage, Lisensi, Respond & Resolution, Downtime Jaringan, Downtime MEKDI BO).
- Kartu ringkasan sekarang grid semua indikator Hasil (minor poin "hanya 4 kartu" selesai).
- Tren satuan % ditampilkan sebagai selisih poin persen (pp).
- Hyperlink di kolom E Excel berupa path SharePoint relatif → tidak dirender sebagai link.
- **Bug 2 sebagian (28 Okt 2026):** periode 202606–202609 diisi dari Google Sheet pemilik "Datamart OPS_Update Sept" (header baris 1, layout beda dari Excel) lewat Excel COM, bukan openpyxl. Excel tetap sumber resmi; pemilik yang update ke depannya. KTP & FP4, DRTD, Klaim CIS CIT, Pencairan diluar PKM belum punya data baru di sheet itu. NOC/AO dan Kelompok/AO 202606+ di sheet berupa bilangan bulat (bukan hasil bagi presisi).
- Jangan simpan ulang Excel lewat openpyxl: ada 80 sel formula, nilai cache-nya hilang dan CI membaca sel kosong.

## Pertanyaan terbuka (jangan ditebak, tanya user)

- Ambang untuk 4 indikator yang belum punya (lihat Status perbaikan).
- Siapa user utama portal dan keputusan apa yang mereka ambil dari sini?
- Dokumen asli LAPORAN ada di mana (upload ke `public/files/` vs link SharePoint asli)?

## Cara kerja lokal

```
pip install -r requirements.txt
python update_data.py            # normal, warning tetap lanjut
python update_data.py --strict   # sama seperti CI
# preview: buka public/index.html di browser (data.js dimuat via <script>)
```

Jangan commit perubahan ke `public/data.js` secara manual — selalu regenerate dari Excel.
