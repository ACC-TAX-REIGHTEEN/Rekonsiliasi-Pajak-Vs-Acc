# 🧾 Rekonsiliasi Pajak vs Accurate (`Rekonsiliasi-Pajak-Vs-Acc`)

![Python](https://img.shields.io/badge/Python-3.9%2B-3776AB?style=for-the-badge&logo=python&logoColor=white)
![Pandas](https://img.shields.io/badge/Pandas-Data%20Processing-150458?style=for-the-badge&logo=pandas&logoColor=white)
![Calamine](https://img.shields.io/badge/Engine-Calamine%20%7C%20OpenPyXL%20%7C%20XLRD-2E8B57?style=for-the-badge)
![XlsxWriter](https://img.shields.io/badge/Output-XlsxWriter-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)

Pipeline otomasi berbasis Python untuk membersihkan, menormalisasi, dan merekonsiliasi data Faktur Pajak antara laporan ekspor **Accurate** (`Acc.xls`) dan data ekspor **Coretax** (`data_export*.xlsx`). Dirancang dengan mesin pembaca Excel adaptif (*adaptive engine fallback*) agar berjalan cepat di lingkungan Pandas terbaru sekaligus tetap kompatibel pada PC kantor lawas.

---

## 📑 Daftar Isi

- [Fitur Utama](#-fitur-utama)
- [Arsitektur & Alur Kerja](#-arsitektur--alur-kerja)
- [Teknologi yang Digunakan](#-teknologi-yang-digunakan)
- [Struktur Direktori](#-struktur-direktori)
- [Prasyarat Sistem](#-prasyarat-sistem)
- [Panduan Instalasi & Konfigurasi](#-panduan-instalasi--konfigurasi)
- [Cara Penggunaan](#-cara-penggunaan)
- [Referensi Konfigurasi](#-referensi-konfigurasi)
- [Penanganan Masalah Umum (Troubleshooting)](#-penanganan-masalah-umum-troubleshooting)
- [Kontribusi & Lisensi](#-kontribusi--lisensi)

---

## ✨ Fitur Utama

### 1. Adaptive Excel Reader (Kompatibel Lintas Versi Pandas)
- **Prioritas `calamine`**: Memanfaatkan parser berbasis Rust (`python-calamine`) untuk pembacaan file `.xls` maupun `.xlsx` secara cepat tanpa memicu peringatan depresiasi Pandas 3/4.
- **Fallback Otomatis**: Secara otomatis beralih ke `openpyxl` (untuk `.xlsx`) atau `xlrd` (khusus format biner `.xls` lawas pada tahap pembersihan) sambil menekan `FutureWarning`/`Pandas4Warning` jika `calamine` belum terpasang di mesin target.

### 2. Pembersihan & Normalisasi Data Accurate (`Dapur/1_AccCleaner.py`)
- **Deteksi Header Dinamis**: Memindai 50 baris awal terlebih dahulu (dan melakukan *full scan* otomatis jika header berada di bawah baris ke-50) untuk menemukan baris header yang mencocokkan minimal 4 dari 8 kolom target.
- **Sanitasi Nomor ID & NPWP**: Membersihkan trailing `.0`, `.00`, `,00`, titik, dan koma pada `No. Referensi`, `No. Faktur Pajak`, dan `Nomor Pajak Pelanggan`, lalu menyimpannya dengan format teks Excel (`@`).
- **Normalisasi Mata Uang & Tanggal Indonesia**: Mengonversi format angka lokal ke tipe numerik (`float`) dan memetakan singkatan bulan Bahasa Indonesia (`Mei`, `Agu`, `Okt`, `Nop`, `Des`) menjadi tanggal standar (`dd/mm/yyyy`).

### 3. Rekonsiliasi & Pelaporan Otomatis (`Dapur/2_CompareAccVsCtx.py`)
- **Outer Merge Komprehensif**: Membandingkan `Acc_temp.xlsx` dan `data_export*.xlsx` berdasarkan kunci `No. Referensi` vs `Referensi` untuk mendeteksi transaksi yang hilang di salah satu sistem.
- **Validasi Nominal Berbasis Toleransi (`calc_gap`)**: Membandingkan `Jumlah Pajak` (Accurate) dengan `PPN` (Coretax) menggunakan ambang batas selisih yang dapat diatur melalui `config.conf`.
- **Validasi Tanggal Pajak**: Mendeteksi ketidakcocokan antara `Tgl. Pajak` (Accurate) dan `Tanggal Faktur Pajak` (Coretax).
- **Laporan Excel Multi-Sheet Terformat**: Menghasilkan `Komparasi Acc vs Ctx.xlsx` dengan 3 lembar kerja (`Hasil Analisis`, `Data Asli Acc_temp`, `Data Asli Export`), ringkasan selisih total, pengurutan prioritas anomali di baris teratas, penandaan warna merah untuk status non-`Cocok`, serta penyesuaian lebar kolom otomatis (*auto-fit*).

### 4. Orkestrasi Pipeline Terpusat (`Proses Rekon.py`)
- Menjalankan seluruh tahapan pemrosesan di dalam folder `Dapur/` secara berurutan dari satu titik masuk (*single entry point*) sehingga operasional harian tidak perlu mengeksekusi skrip pekerja satu per satu.

---

## 🏗️ Arsitektur & Alur Kerja

```text
+-----------------------+       +---------------------------+       +-----------------------+
|   Input Accurate      |       |     Input Coretax         |       |     Konfigurasi       |
|      (Acc.xls)        |       |   (data_export*.xlsx)     |       |   (Dapur/config.conf) |
+-----------+-----------+       +-------------+-------------+       +-----------+-----------+
            |                                 |                                 |
            v                                 |                                 |
+-----------------------+                     |                                 |
|  1_AccCleaner.py      |                     |                                 |
|  - Adaptive reader    |                     |                                 |
|  - Dynamic header     |                     |                                 |
|  - Clean ID/Date/Curr |                     |                                 |
+-----------+-----------+                     |                                 |
            |                                 |                                 |
            v                                 |                                 |
+-----------------------+                     |                                 |
|  Intermediate File    |                     |                                 |
|   (Acc_temp.xlsx)     |                     |                                 |
+-----------+-----------+                     |                                 |
            |                                 |                                 |
            +----------------+----------------+----------------+----------------+
                             |
                             v
              +-----------------------------+
              |    2_CompareAccVsCtx.py     |
              |  - Outer Merge Referensi    |
              |  - Cek Nominal (> calc_gap) |
              |  - Cek Beda Tanggal Pajak   |
              |  - Sort Anomali di Atas     |
              +--------------+--------------+
                             |
                             v
              +-----------------------------+
              |  Komparasi Acc vs Ctx.xlsx  |
              |  1. Hasil Analisis          |
              |  2. Data Asli Acc_temp      |
              |  3. Data Asli Export        |
              +-----------------------------+
```

---

## 🛠️ Teknologi yang Digunakan

| Komponen | Teknologi / Pustaka | Peran dalam Sistem |
| :--- | :--- | :--- |
| **Bahasa Utama** | Python 3.9+ | Runtime utama orkestrator dan modul ETL |
| **Olah Data** | `pandas` | Parsing tabular, transformasi kolom, *outer merge*, dan agregasi |
| **Excel Reader (Utama)** | `python-calamine` | Mesin baca berkecepatan tinggi untuk `.xls` dan `.xlsx` |
| **Excel Reader (Fallback)** | `openpyxl`, `xlrd` | Mesin cadangan untuk file `.xlsx` (`openpyxl`) dan `.xls` biner (`xlrd`) |
| **Excel Writer** | `xlsxwriter` | Generasi workbook terformat (warna header, border, format angka/teks, highlight anomali) |
| **Konfigurasi & Sistem** | `configparser`, `glob`, `os`, `sys`, `warnings` | Pustaka standar Python untuk manajemen parameter, pencarian berkas, dan supresi warning |

---

## 📂 Struktur Direktori

```text
Rekonsiliasi-Pajak-Vs-Acc-test-pandas/
├── Proses Rekon.py              # Entry point / orkestrator utama pipeline rekonsiliasi
├── README.md                    # Dokumentasi proyek
└── Dapur/                       # Direktori modul pekerja dan konfigurasi
    ├── __init__.py              # Penanda paket Python untuk direktori Dapur
    ├── 1_AccCleaner.py          # Tahap 1: Pembersihan & normalisasi Acc.xls -> Acc_temp.xlsx
    ├── 2_CompareAccVsCtx.py     # Tahap 2: Rekonsiliasi Acc_temp.xlsx vs data_export*.xlsx
    └── config.conf              # File konfigurasi toleransi selisih (calc_gap)
```

---

## 📋 Prasyarat Sistem

### Lingkungan Perangkat Lunak
- **Python 3.9 atau lebih baru** (direkomendasikan menggunakan *virtual environment*).
- Sistem operasi Windows 10/11, Linux, atau macOS.

### Berkas Masukan yang Wajib Tersedia
1. **File Ekspor Accurate (`Acc.xls`)**
   Harus memuat baris header yang mencakup kolom-kolom target berikut:
   - `Tanggal`
   - `Tgl. Pajak`
   - `No. Referensi`
   - `No. Faktur Pajak`
   - `Nama Pelanggan`
   - `Negara Pelanggan`
   - `Jumlah Pajak`
   - `Nomor Pajak Pelanggan`

2. **File Ekspor Coretax (`data_export*.xlsx`)**
   Penamaan file diawali `data_export` dengan ekstensi `.xlsx` dan wajib memuat kolom berikut:
   - `Tanggal Faktur Pajak`
   - `Referensi`
   - `PPN`
   - `Status Faktur`

---

## ⚙️ Panduan Instalasi & Konfigurasi

### 1. Clone Repositori
```bash
git clone <url-repositori-github-anda>
cd Rekonsiliasi-Pajak-Vs-Acc-test-pandas
```

### 2. Buat dan Aktifkan Virtual Environment (Opsional, Direkomendasikan)
```bash
# Windows (PowerShell)
python -m venv .venv
.\.venv\Scripts\Activate.ps1

# Linux / macOS
python3 -m venv .venv
source .venv/bin/activate
```

### 3. Instal Dependensi Python
Instal seluruh pustaka yang diperlukan termasuk engine utama (`python-calamine`) dan engine cadangan (`openpyxl`, `xlrd`):
```bash
pip install pandas python-calamine openpyxl xlrd xlsxwriter
```

### 4. Konfigurasi `Dapur/config.conf`
Buka atau buat file `config.conf` di dalam direktori `Dapur/` (atau pada direktori kerja aktif tempat skrip dijalankan). Gunakan format INI standar dengan seksi `[CALC]`:

```ini
[CALC]
calc_gap = 0
```

> **Catatan Konfigurasi:**
> - Nilai `calc_gap` menentukan batas toleransi selisih nominal (dalam Rupiah) antara `Jumlah Pajak` (Accurate) dan `PPN` (Coretax).
> - Contoh: jika `calc_gap = 100`, maka selisih pembulatan $\le \text{Rp }100$ tetap dianggap **Cocok** (tidak ditandai `Selisih Nominal`).
> - Jika `config.conf` tidak ditemukan, rusak, atau tidak memiliki kunci `calc_gap`, sistem otomatis menggunakan nilai bawaan `0.0`.

---

## 🚀 Cara Penggunaan

### Opsi A: Eksekusi Penuh via Orkestrator (Direkomendasikan)
1. Letakkan file masukan `Acc.xls` dan `data_export*.xlsx` pada direktori kerja yang sesuai dengan konfigurasi proyek.
2. Pastikan parameter toleransi di `Dapur/config.conf` sudah sesuai kebutuhan.
3. Jalankan orkestrator utama:
   ```bash
   python "Proses Rekon.py"
   ```
4. Tunggu hingga seluruh tahapan selesai dan periksa file hasil **`Komparasi Acc vs Ctx.xlsx`**.

### Opsi B: Eksekusi Manual per Tahap (Debugging / Pengujian)
Jika ingin menjalankan modul secara terpisah dari direktori yang berisi `Acc.xls`, `data_export*.xlsx`, dan `config.conf`:

```bash
# Tahap 1: Bersihkan laporan Accurate menjadi Acc_temp.xlsx
python Dapur/1_AccCleaner.py

# Tahap 2: Bandingkan Acc_temp.xlsx dengan data_export*.xlsx
python Dapur/2_CompareAccVsCtx.py
```

### Membaca Hasil Output (`Komparasi Acc vs Ctx.xlsx`)
Workbook keluaran terdiri atas 3 sheet:
1. **`Hasil Analisis`**:
   - **Bagian A (`Ringkasan Total`)**: Menampilkan Total Jumlah Pajak Accurate, Total PPN Coretax, dan Selisih Total.
   - **Bagian B (`Detail Data Komparasi`)**: Menampilkan perbandingan per nomor referensi. Semua baris anomali ditempatkan di bagian atas dengan sorotan merah pada kolom `Keterangan Analisis`:
     - `Cocok`: Referensi ditemukan di kedua sumber, selisih nominal $\le$ `calc_gap`, dan tanggal pajak sama.
     - `Selisih Nominal`: Selisih absolut `|Jumlah Pajak - PPN|` melebihi `calc_gap`.
     - `Beda Tanggal`: `Tgl. Pajak` (Accurate) berbeda hari dengan `Tanggal Faktur Pajak` (Coretax).
     - `Hanya ada di Acc_temp`: Faktur tercatat di Accurate tetapi tidak ditemukan di ekspor Coretax.
     - `Hanya ada di Data Export`: Faktur tercatat di Coretax tetapi tidak ditemukan di ekspor Accurate.
2. **`Data Asli Acc_temp`**: Salinan penuh data Accurate yang telah dibersihkan untuk keperluan audit cepat.
3. **`Data Asli Export`**: Salinan penuh data ekspor Coretax sebagai rujukan silang.

---

## 📘 Referensi Konfigurasi

File konfigurasi: `Dapur/config.conf`

| Seksi | Parameter | Tipe Data | Default | Contoh Nilai | Deskripsi |
| :--- | :--- | :--- | :--- | :--- | :--- |
| `[CALC]` | `calc_gap` | `float` | `0.0` | `0`, `10`, `100.5` | Batas maksimal toleransi selisih absolut antara `Jumlah Pajak` (Accurate) dan `PPN` (Coretax) sebelum ditandai sebagai `Selisih Nominal`. |

---

## 🔧 Penanganan Masalah Umum (Troubleshooting)

| Gejala / Pesan Error | Penyebab Umum | Solusi |
| :--- | :--- | :--- |
| `Header kolom tidak ditemukan.` | Struktur kolom di `Acc.xls` berubah atau kurang dari 4 kolom target yang cocok. | Pastikan ekspor Accurate menggunakan template laporan pajak standar yang berisi 8 kolom wajib (`Tanggal`, `Tgl. Pajak`, `No. Referensi`, `No. Faktur Pajak`, `Nama Pelanggan`, `Negara Pelanggan`, `Jumlah Pajak`, `Nomor Pajak Pelanggan`). |
| `File Acc_temp.xlsx tidak ditemukan.` | Tahap 2 dijalankan sebelum Tahap 1 berhasil membuat `Acc_temp.xlsx`, atau *working directory* tidak sesuai. | Jalankan `Proses Rekon.py` atau pastikan `1_AccCleaner.py` selesai tanpa error di direktori kerja yang sama. |
| `File data_export tidak ditemukan.` | Tidak ada file berawalan `data_export` dengan ekstensi `.xlsx` di direktori kerja. | Pastikan file ekspor Coretax berformat `.xlsx` dan namanya diawali `data_export` (misalnya `data_export_nov2026.xlsx`). |
| `Calamine gagal (...), beralih ke fallback engine...` | Pustaka `python-calamine` bermasalah atau format berkas memiliki anomali struktur. | Periksa apakah `openpyxl` dan `xlrd` sudah terinstal. Jika ingin kecepatan optimal di PC lama/baru, jalankan `pip install -U python-calamine`. |
| `File config.conf tidak ditemukan. Menggunakan default gap = 0.0` | `config.conf` tidak berada di direktori kerja aktif saat `2_CompareAccVsCtx.py` dieksekusi. | Pastikan `config.conf` tersedia di direktori kerja eksekusi atau jalankan melalui `Proses Rekon.py`. |
| `PermissionError: [Errno 13] Permission denied: 'Komparasi Acc vs Ctx.xlsx'` | File hasil komparasi sebelumnya masih terbuka di Microsoft Excel. | Tutup file `Komparasi Acc vs Ctx.xlsx` (atau `Acc_temp.xlsx`) di Excel terlebih dahulu, lalu jalankan ulang skrip. |
| Banyak baris muncul sebagai `Hanya ada di Acc_temp` & `Hanya ada di Data Export` padahal nomor sama | Perbedaan format spasi/karakter khusus pada kolom `No. Referensi` vs `Referensi` (misalnya pemisah titik/tanda hubung yang berbeda di salah satu sumber). | Periksa tab `Data Asli Acc_temp` dan `Data Asli Export` untuk memastikan pola nomor referensi di Accurate konsisten dengan kolom `Referensi` di Coretax. |

---

## 🤝 Kontribusi & Lisensi

### Alur Kontribusi
1. *Fork* repositori ini.
2. Buat branch fitur baru (`git checkout -b fitur/nama-fitur`).
3. *Commit* perubahan Anda dengan pesan yang deskriptif (`git commit -m "Menambahkan fitur X"`).
4. *Push* ke branch tersebut (`git push origin fitur/nama-fitur`).
5. Ajukan *Pull Request* untuk ditinjau.

## 📜 Lisensi

Proyek ini dikembangkan untuk keperluan internal internal perusahaan. Silakan sesuaikan dengan kebutuhan organisasi Anda.

---

* Dikembangkan oleh [ACC-TAX-REIGHTEEN](https://github.com/ACC-TAX-REIGHTEEN)
