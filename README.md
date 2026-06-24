# Portal Risiko Operasional

Self-service portal untuk seluruh divisi internal.

---

## 🔒 Soal Keamanan — Baca Ini Dulu

- Repo ini **private** — kode & riwayat Excel tidak bisa diakses orang luar
- Portal punya **kode akses sederhana** (password gate) di halaman pertama
- **PENTING — pahami batasannya:** kode akses ini cuma penyaring ringan di sisi browser. Siapapun yang tahu URL Vercel-nya dan cukup paham teknis bisa membuka `data.js` langsung tanpa lewat gate, karena file data tetap di-serve secara publik oleh Vercel. Ini **bukan** pengganti keamanan sesungguhnya.
- **Kalau data di portal ini sensitif**, aktifkan **Vercel Deployment Protection** (Project Settings → Deployment Protection → Password Protection, perlu plan Pro) atau pasang autentikasi di level server/edge. Itu satu-satunya cara menutup akses ke `data.js` secara nyata.
- Ganti kode akses default di `public/index.html`, cari variabel `GATE_PASSWORD`.

---

## 🚀 Setup Pertama Kali (hanya sekali)

### 1. Push repo ini ke GitHub (private)
```
Buat repo baru di github.com → Private → upload semua file ini
```

### 2. Connect ke Vercel
```
vercel.com → Add New Project → Import repo ini → Deploy
```
> Vercel otomatis baca `vercel.json` dan deploy folder `public/`

### 3. Aktifkan GitHub Actions
```
Di repo GitHub → tab Actions → klik "I understand my workflows, enable them"
```

Selesai. Sekarang setiap kali lo upload Excel baru, portal update otomatis.

---

## 📅 Cara Update Data (TANPA CODING)

### Langkah 1 — Update Excel
Buka `Datamart_Risiko_Operasional.xlsx` di komputer:
- Sheet **DATAMART** → isi kolom kuning (periode baru)
- Sheet **LAPORAN** → tambah/edit baris dokumen kalau ada yang baru
- Simpan file

### Langkah 2 — Upload ke GitHub (drag & drop)
1. Buka repo ini di **github.com**
2. Klik file `Datamart_Risiko_Operasional.xlsx`
3. Klik ikon **pensil (edit)** atau tombol **"..."** → **Replace file**
   > Atau: di halaman utama repo → drag & drop file Excel langsung ke browser
4. Scroll ke bawah → **Commit changes** → klik **Commit changes**

### Langkah 3 — Tunggu ~1 menit
GitHub Actions otomatis:
- Baca Excel
- **Validasi format** tiap sel (angka/persen rusak, KEY duplikat, JENIS typo, dll)
- Kalau ada masalah format → proses **berhenti, portal tidak ter-update** (mode `--strict`), supaya data rusak tidak pernah sampai ke pengguna
- Kalau bersih → generate ulang `public/data.js`, push ke repo, Vercel re-deploy

**Portal langsung update — tidak perlu buka terminal sama sekali.**

### Cek status update
Tab **Actions** di GitHub → lihat workflow "Update Portal Data":
- ✅ hijau → berhasil update
- ❌ merah → ada masalah format di Excel, klik untuk lihat detail pesan error, perbaiki, lalu upload ulang

### Menjalankan manual (opsional, kalau punya Python di komputer)
```
python update_data.py            # jalan normal, warning ditampilkan tapi tetap lanjut
python update_data.py --strict   # mode sama seperti yang dipakai GitHub Actions
```

---

## 📁 Cara Tambah File Dokumen Baru

1. **Upload file** ke folder `public/files/` di GitHub
   - Buka folder `public/files/` di repo → **Add file → Upload files**
2. **Tambah baris** di sheet `LAPORAN` di Excel:
   - Kolom Link: `files/nama_file.pdf`
   - Kolom Aktif: `Y`
3. Upload ulang Excel → otomatis update seperti biasa

---

## 📁 Struktur Repo

```
├── .github/workflows/
│   └── update-data.yml         ← GitHub Actions, jalankan update_data.py --strict
├── public/
│   ├── index.html              ← Portal web (termasuk gate akses)
│   ├── data.js                 ← AUTO-GENERATED, jangan edit manual
│   └── files/                  ← Taruh PDF/XLSX/dll di sini
├── Datamart_Risiko_Operasional.xlsx   ← ← ← EDIT INI
├── update_data.py              ← Dijalankan otomatis oleh GitHub Actions
├── requirements.txt
└── vercel.json
```

---

Internal — Manajemen Risiko Operasional
