# POP SIPGN Autofill — SPPG Kalisapu

[![Python Version](https://img.shields.io/badge/python-3.10%2B-blue.svg)](https://www.python.org/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![Platform](https://img.shields.io/badge/platform-Windows-lightgrey.svg)]()

Aplikasi desktop cerdas berbasis **Python + Tkinter** untuk mengotomatisasi pengisian formulir **Jumlah Porsi per KPM (Porsi Besar & Porsi Kecil)** pada sistem **POP SIPGN (Langkah 2 dari 4: Tambah Penugasan Produksi)** berdasarkan data dari file Excel (`.xlsx`).

Aplikasi ini dirancang khusus untuk mempermudah operasional kerja PKL di **SPPG Kalisapu**, dengan fokus pada kemudahan penggunaan, kecepatan eksekusi, akurasi data, dan kepatuhan penuh terhadap standar keselamatan web resmi.

---

## Daftar Isi

1. [Pengenalan & Tujuan](#pengenalan--tujuan)
2. [Fitur Utama](#fitur-utama)
3. [Teknologi yang Digunakan](#teknologi-yang-digunakan)
4. [Persyaratan Sistem](#persyaratan-sistem)
5. [Instalasi & Pengaturan](#instalasi--pengaturan)
6. [Cara Menjalankan & Menggunakan](#cara-menjalankan--menggunakan)
7. [Struktur Folder & File](#struktur-folder--file)
8. [Konfigurasi Environment & Keamanan](#konfigurasi-environment--keamanan)
9. [Pengujian Otomatis (Unit Test)](#pengujian-otomatis-unit-test)
10. [Cara Berkontribusi](#cara-berkontribusi)
11. [Lisensi & Kontak Pengembang](#lisensi--kontak-pengembang)

---

## Pengenalan & Tujuan

Pada alur operasional POP SIPGN, operator SPPG perlu memasukkan alokasi porsi harian (Porsi Besar dan Porsi Kecil) untuk puluhan sekolah dan posyandu penerima manfaat (KPM). Memasukkan data ini secara manual satu per satu memakan waktu, rentan salah ketik (*human error*), dan melelahkan.

**Tujuan Project Ini:**
- Mempercepat proses entri porsi dari puluhan menit menjadi hitungan detik.
- Menjamin akurasi nilai porsi sesuai rekap file Excel resmi.
- Menghilangkan repotnya login ulang browser dengan pendekatan *in-browser script execution*.
- Mempertahankan kontrol manusia (*human-in-the-loop*): tombol **Lanjut** dan **Simpan** tetap ditekan secara manual oleh pengguna setelah memeriksa data.

---

## Fitur Utama

- 📂 **Pembacaan Excel Dinamis:** Mendukung pemilihan file `.xlsx`, deteksi sheet otomatis (misalnya Sheet *Oktober*), dan deteksi otomatis kolom sekolah serta porsi.
- 🔢 **Preservasi Nilai Nol & Validasi Integer:** Memvalidasi bahwa porsi adalah bilangan bulat nonnegatif ($\ge 0$). Nilai `0` dipertahankan sebagai angka `0` (bukan dianggap kosong).
- 🧠 **Normalisasi Nama Cerdas:** Menyamakan variasi singkatan sekolah Indonesia (`SDN` $\leftrightarrow$ `SD NEGERI`, `TKN` $\leftrightarrow$ `TK NEGERI`, `KEC.` $\leftrightarrow$ `KECAMATAN`), spasi ganda, tanda baca, serta perbedaan kapitalisasi.
- ⚡ **Eksekusi Seketika Tanpa Login Ulang:** Menghasilkan skrip aman yang dieksekusi langsung di tab browser pengguna yang sudah login, sehingga terhindar dari pemblokiran sesi, login berulang, atau CAPTCHA.
- 🎯 **Kompatibilitas React / Modern Framework:** Nilai formulir disuntikkan menggunakan event setter internal (`_valueTracker` + event `input`, `change`, `blur`), memastikan nilai tersimpan dengan benar di state sistem web.
- 🟢 **Visual Feedback di Browser:** Kartu sekolah yang berhasil terisi otomatis diberi bingkai (*border*) warna hijau di layar browser untuk memudahkan verifikasi visual.
- 🛡️ **Aturan Keselamatan Ketat:** Tidak pernah menekan tombol finalisasi (Lanjut/Kirim/Simpan) secara otomatis, tidak menyimpan kredensial/password, dan tidak mengubah file Excel sumber.
- 💾 **Audit Log & Ekspor Data:** Riwayat proses dicatat dalam folder `logs/` dan data tervalidasi dapat diekspor ke format `.csv` di folder `reports/`.

---

## Teknologi yang Digunakan

- **Bahasa Pemrograman:** Python 3 (kompatibel dengan Python 3.10 hingga 3.15+)
- **Antarmuka Pengguna (GUI):** Tkinter & ttk (bawaan standar Python)
- **Pengolahan Data:**
  - `pandas` (manipulasi tabel data)
  - `openpyxl` (pembacaan berkas Excel `.xlsx`)
- **Eksekusi Otomatisasi:** JavaScript DOM Injection via browser DevTools Console (kompatibel dengan Chromium: Brave, Google Chrome, Microsoft Edge)
- **Peluncur:** Batch Script Windows (`run.bat`) dengan pendeteksi dependensi otomatis

---

## Persyaratan Sistem

- **Sistem Operasi:** Windows 10 atau Windows 11 (64-bit disarankan)
- **Python:** Versi 3.10 atau yang lebih baru ([Unduh di python.org](https://www.python.org/downloads/))
- **Browser:** Browser berbasis Chromium yang sudah terpasang (Brave Browser, Google Chrome, atau Microsoft Edge)
- **Koneksi Internet:** Untuk mengakses portal resmi POP SIPGN

---

## Instalasi & Pengaturan

### 1. Unduh / Clone Repository
Buka terminal (PowerShell atau Command Prompt), lalu jalankan:

```bash
git clone https://github.com/SanjayaPutra23/pop-sipgn-autofill.git
cd pop-sipgn-autofill
```

Atau unduh file ZIP dari tombol **Code ➔ Download ZIP** di GitHub, lalu ekstrak ke folder komputer Anda.

### 2. Pasang Dependensi Python
Pastikan Python sudah terpasang di sistem (ingat centang **"Add Python to PATH"** saat instalasi). Buka terminal di folder project, lalu ketik:

```cmd
pip install -r requirements.txt
```

*Hanya ada 2 library ringan yang diinstal: `pandas` dan `openpyxl`.*

---

## Cara Menjalankan & Menggunakan

### Cara 1: Menggunakan `run.bat` (Rekomendasi - Cukup 1 Klik)
Klik ganda (*double-click*) file **`run.bat`**. Skrip akan otomatis memeriksa apakah Python dan dependensi sudah terpasang, lalu langsung membuka aplikasi.

### Cara 2: Menjalankan via Terminal
```cmd
python main.py
```

---

### Alur Kerja Harian (3 Langkah Cepat)

```text
┌────────────────────────┐      ┌────────────────────────┐      ┌────────────────────────┐
│  1. Muat Data Excel    │ ──➔  │  2. Salin Kode         │ ──➔  │  3. Paste di Browser   │
│  (Pilih file & sheet)  │      │  (Klik 1 tombol)       │      │  (F12 ➔ Console ➔ Done)│
└────────────────────────┘      └────────────────────────┘      └────────────────────────┘
```

1. **Langkah 1: Muat & Validasi Data di Aplikasi**
   - Klik **Pilih File...** dan tentukan file Excel data porsi Anda (atau gunakan `contoh_data_sppg.xlsx`).
   - Pilih Sheet yang ingin diproses (misal: *Oktober*).
   - Kolom nama sekolah dan porsi akan terdeteksi otomatis. Sesuaikan jika diperlukan.
   - Klik tombol **📥 Muat Data**.
2. **Langkah 2: Salin Kode Autofill**
   - Klik tombol: **`📋 SALIN KODE AUTOFILL (PASTE DI CONSOLE BROWSER)`**.
   - Kode pengisian otomatis instan akan langsung tersalin ke Clipboard komputer Anda.
3. **Langkah 3: Jalankan di Tab Browser Anda**
   - Buka browser Anda yang **sudah login ke POP SIPGN dan berada di Langkah 2 (Jumlah Porsi per KPM)**.
   - Tekan tombol **F12** pada keyboard (atau Klik Kanan ➔ **Inspect / Periksa**).
   - Masuk ke tab **Console**.
   - *(Jika baru pertama kali di browser tersebut)*: Ketik manual `allow pasting` lalu tekan **Enter**.
   - Tekan **Ctrl + V** (Paste), lalu tekan **Enter**.
   - Seluruh kartu formulir akan langsung terisi dalam 1 detik dengan bingkai hijau sebagai penanda sukses.
   - Periksa kembali isian formulir, lalu klik tombol **Lanjut** secara manual di website.

---

## Struktur Folder & File

```text
pop-sipgn-autofill/
├── main.py                 # Antarmuka grafis desktop (Tkinter GUI) & generator skrip
├── excel_reader.py         # Modul pembaca berkas Excel (.xlsx) secara read-only
├── validator.py            # Modul validasi format porsi (integer >= 0 & baris kosong)
├── test_core.py            # Pengujian unit otomatis (unit testing)
├── contoh_data_sppg.xlsx   # Berkas data Excel sampel SPPG Kalisapu untuk uji coba
├── requirements.txt        # Daftar dependensi library (pandas, openpyxl)
├── run.bat                 # Peluncur Windows otomatis (auto check & launch)
├── .gitignore              # Daftar berkas/folder yang dikecualikan dari Git
├── LICENSE                 # Lisensi open-source MIT
├── README.md               # Dokumentasi lengkap project
├── logs/                   # Folder penyimpanan berkas log aktivitas audit
│   └── .gitkeep
└── reports/                # Folder penyimpanan hasil ekspor data laporan
    └── .gitkeep
```

---

## Konfigurasi Environment & Keamanan

- **Zero Secret Exposure:** Project ini **tidak memerlukan** file `.env`, API key, token rahasia, ataupun password.
- **Tanpa Penyimpanan Kredensial:** Aplikasi tidak pernah meminta, membaca, atau menyimpan kata sandi akun POP SIPGN pengguna.
- **Read-Only Source:** File Excel sumber dibuka secara *read-only* sehingga data asli Anda tidak akan pernah tertimpa atau rusak.

---

## Pengujian Otomatis (Unit Test)

Project ini dilengkapi skrip pengujian unit otomatis untuk memverifikasi keakuratan pembacaan Excel dan validasi angka porsi. Jalankan tes kapan saja dengan perintah:

```cmd
python test_core.py
```

---

## Cara Berkontribusi

Kontribusi untuk pengembangan project ini sangat terbuka:

1. **Fork** repository ini ke akun GitHub Anda.
2. Buat branch fitur baru (`git checkout -b fitur-baru`).
3. Lakukan commit perubahan Anda dengan pesan yang jelas (`git commit -m 'Menambahkan validasi format porsi baru'`).
4. Push branch Anda (`git push origin fitur-baru`).
5. Buat **Pull Request** baru di GitHub.

---

## Lisensi & Kontak Pengembang

Project ini didistribusikan di bawah lisensi open-source [MIT License](LICENSE). Bebas digunakan, dipelajari, dan dimodifikasi untuk kebutuhan operasional maupun edukasi.

- **Pengembang:** Sanjaya Putra
- **Institusi/Afiliasi:** SPPG Kalisapu
