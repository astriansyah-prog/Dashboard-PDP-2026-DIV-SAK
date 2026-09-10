# Monitoring PDP sampai Juni 2026 — DIV SAK, PT PLN (Persero)

Dashboard satu halaman (single-page) untuk memantau **PDP (Pekerjaan Dalam
Pelaksanaan)**: infografis sebaran proyek per UIP, performance monitoring,
early warning system (EWS), rincian kategori D1–D4, 9 cluster proyek
berkendala, dan rekonsiliasi nilai PDP DIV PMO vs DIV AKT (SAP). Tampilan &
logikanya mengikuti dashboard asli yang diupload, namun datanya sekarang
diambil langsung dari Google Spreadsheet sehingga selalu bisa diperbarui
tanpa mengedit HTML.

## Struktur folder

```
index.html               -> tampilan dashboard (HTML+CSS+JS, tidak perlu diedit)
build_data.py             -> script Python yang mengambil data dari Google Spreadsheet
                             dan mengubahnya menjadi data/pdp-data.json
map_path.txt              -> data path SVG peta Indonesia (statis, untuk peta bubble)
data/pdp-data.json        -> data hasil olahan (dibaca oleh index.html lewat fetch())
requirements.txt          -> dependency Python (openpyxl)
.github/workflows/
  update-data.yml          -> GitHub Action: ambil data terbaru tiap hari + deploy ke Pages
```

## Cara kerja singkat

1. `build_data.py` membaca spreadsheet sumber (sheet: `SAP-DIVAKT`,
   `Rekap Settlement`, `D1`, `D2`, `D3`, `D4`, `Cluster Proyek 1-9`,
   `Biaya Ditangguhkan`) dan menghitungnya menjadi satu file
   `data/pdp-data.json` — persis skema data yang dipakai dashboard.
2. `index.html` memuat `data/pdp-data.json` lewat `fetch()` saat halaman
   dibuka, lalu menjalankan seluruh logika dashboard (peta, grafik Chart.js,
   tabel, filter, deteksi proyek kritis) di sisi browser.
3. GitHub Action (`update-data.yml`) menjalankan `build_data.py` setiap
   hari secara otomatis, meng-commit `data/pdp-data.json` jika ada
   perubahan, lalu mem-publish ulang ke GitHub Pages — jadi dashboard
   selalu menampilkan data terbaru dari spreadsheet tanpa Anda perlu
   melakukan apa pun secara manual.

## Deploy ke GitHub Pages (langkah demi langkah)

1. Buat repository baru di GitHub (bisa publik atau privat — privat pun
   tetap bisa memakai GitHub Pages di paket berbayar; kalau ingin gratis,
   pakai repo publik).
2. Upload semua file di folder ini ke repository tersebut (lewat
   "Add file → Upload files" di web GitHub, atau lewat `git push` biasa).
3. Buka **Settings → Pages** di repo, pada "Build and deployment" pilih
   source **GitHub Actions**.
4. Buka **Settings → Actions → General**, pastikan "Workflow permissions"
   diset ke **Read and write permissions** (supaya Action bisa melakukan
   commit otomatis untuk `data/pdp-data.json`).
5. Buka tab **Actions**, jalankan workflow **"Update data & deploy
   dashboard"** sekali secara manual ("Run workflow") untuk build pertama.
6. Setelah selesai, link dashboard bisa dilihat di
   **Settings → Pages** (format umumnya
   `https://<username>.github.io/<nama-repo>/`).

Setelah langkah di atas, dashboard akan otomatis memperbarui datanya
**setiap hari jam 06:00 WIB** mengikuti isi spreadsheet terbaru. Anda juga
bisa memicu update kapan saja lewat tombol "Run workflow" di tab Actions.

## Sumber data (Google Spreadsheet)

Spreadsheet sumber:
`https://docs.google.com/spreadsheets/d/15RlWr9D1QsWxqBKnzFlz0bcFhzkxxM_xLvxYjpR-SzQ`

Spreadsheet ini **harus** tetap di-share sebagai **"Anyone with the link
can view"** (siapa pun yang punya link bisa melihat) agar `build_data.py`
bisa mengambil datanya secara otomatis lewat GitHub Action. Jika Anda
membuat salinan spreadsheet sendiri, atur ID spreadsheet baru lewat
**Settings → Secrets and variables → Actions → Variables**, buat variable
bernama `SHEET_ID` berisi ID spreadsheet Anda (bagian di antara `/d/` dan
`/edit` pada URL spreadsheet).

Struktur sheet yang dibutuhkan (nama sheet harus persis sama):

| Sheet                     | Isi                                                        |
|---------------------------|-------------------------------------------------------------|
| `SAP-DIVAKT`               | Saldo PDP per entitas (tarikan SAP dari DIV AKT)             |
| `Rekap Settlement`         | Target & realisasi settlement PDP/ATBM ke Aset Tetap per UIP |
| `D1`, `D2`, `D3`           | Daftar proyek per kategori (settlement, penerbitan SLO, konstruksi berjalan) |
| `D4`                       | Material PDP yang belum termanfaatkan                        |
| `Cluster Proyek 1-9`       | Pengelompokan proyek berkendala non-konstruksi (izin, lahan, kontrak, dll.) |
| `Biaya Ditangguhkan`       | Biaya pra-konstruksi yang belum masuk PDP                    |

## Menjalankan / mengetes secara lokal

```bash
pip install -r requirements.txt

# opsi A: pakai spreadsheet online langsung (harus sudah publik/"anyone with link")
python3 build_data.py

# opsi B: pakai file Excel yang sudah didownload manual (File > Download > .xlsx)
#   letakkan sebagai sheet.xlsx di folder ini, lalu:
LOCAL_XLSX=sheet.xlsx python3 build_data.py

# lalu buka dashboard lewat web server lokal (WAJIB pakai server, tidak bisa
# dibuka langsung dari file lokal karena fetch() butuh http:// atau https://)
python3 -m http.server 8000
# buka http://localhost:8000/index.html di browser
```

## Menambah bulan baru di spreadsheet

Struktur spreadsheet saat ini mendukung 10 bulan tetap: September 2025
sampai Juni 2026, dengan **Juni 2026 sebagai bulan snapshot** yang
ditampilkan di seluruh dashboard (KPI, peta, tabel D1-D4, dsb). Kalau Anda
menambah kolom bulan baru (misal Juli 2026) di sheet `D1`/`D2`/`D3`/`D4`,
ikuti nama kolom yang sama persis dengan pola sebelumnya (`Saldo PDP
<Bulan> <Tahun>`, `Progress Fisik (%) ... <Bulan> <Tahun>` khusus `D3`),
lalu:

1. Tambahkan entri bulan baru ke daftar `BULAN` di bagian atas
   `build_data.py`.
2. Ubah `PERIODE_LABEL` di `build_data.py` ke bulan snapshot yang baru
   (mis. `"Juli 2026"`), dan sesuaikan kolom yang dicari untuk saldo/
   progress "bulan terakhir" (`find_col(h, "saldo pdp", "juni")` dst.)
   menjadi nama bulan yang baru.

Script ini mencari kolom berdasarkan **teks header**, bukan posisi/urutan
kolom, jadi cukup aman walau urutan kolom di spreadsheet berantakan — asal
teks headernya konsisten dengan bulan-bulan sebelumnya.

## Menggunakan dashboard ini di aplikasi lain / AI app builder

`index.html` + folder `data/` adalah situs statis biasa (tidak butuh
backend/database server) — jadi bisa langsung dipakai/di-embed di:
- **GitHub Pages** (lihat langkah di atas) — paling sederhana.
- Platform hosting statis lain (Netlify, Vercel, Cloudflare Pages, dsb) —
  cukup upload folder ini, atur `build_data.py` untuk jalan terjadwal
  lewat fitur scheduled build masing-masing platform.
- Di-embed via `<iframe>` ke dalam website/aplikasi lain.
- Dibaca ulang oleh AI app builder (mis. untuk dimodifikasi tampilannya)
  karena seluruh logika ada di satu file `index.html` yang mudah dibaca.

## Catatan tentang keakuratan data

`build_data.py` sudah diuji dengan membandingkan hasil parsing terhadap
data pada dashboard versi awal yang diupload (posisi Juni 2026, digenerate
10 September 2026) — seluruh data (SAP, rekap settlement, proyek D1-D3,
material D4, cluster 1-9, biaya ditangguhkan) cocok **hingga ke detail
baris dan karakter demi karakter**, termasuk kasus-kasus khusus seperti sel
kosong, nilai yang salah format, dan baris pengecualian (unit yang tidak
dikenal, baris "NIHIL"/"TOTAL", dsb).

Peta sebaran (path SVG Indonesia dan posisi bubble tiap UIP di
`map_path.txt`/`build_data.py`) bersifat statis/presentasional dan tidak
berasal dari spreadsheet — hanya nilai (d1/d2/d3/d4/total) di tiap unit
yang dihitung dari data live.
