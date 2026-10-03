# 🗺️ Google Maps Review Scraper (Kecepatan Ultra)

Google Maps Review Scraper adalah alat otomasi (bot) berbasis Python dan Playwright yang dirancang khusus untuk mengekstrak ratusan hingga ribuan ulasan dari suatu lokasi/bisnis di Google Maps secara **otomatis, cepat, dan tanpa batas**. 

Berbeda dengan scraper tradisional, alat ini menggunakan metode injeksi JavaScript langsung ke inti *browser* untuk melakukan ekstraksi dengan **kecepatan ultra** tanpa mengalami lag atau kemacetan (O(N^2) complexity fix). Alat ini mampu mem-bypass batasan pemuatan lambat Google Maps dan menyimpan hasilnya ke dalam format yang rapi.

## ✨ Fitur Unggulan
- ⚡ **Kecepatan Ultra**: Ekstraksi menggunakan native JavaScript browser, memproses ratusan ulasan dalam waktu kurang dari sedetik per gulir.
- 🔓 **Bypass Anti-Bot**: Menggunakan profil Chrome asli sehingga aman dari pemblokiran login Google.
- 🔄 **Otomasi Cerdas**: Men-scroll halaman secara mandiri, mengeklik tombol "Selengkapnya", dan melewati ulasan yang tidak relevan secara presisi.
- 📊 **Ekspor Rapi**: Hasil ulasan (Nama, Rating, Komentar, dan Tanggal) diekspor langsung ke **Microsoft Excel (.xlsx)** berserta styling rapi, dan format CSV standar.

---

## 🛠️ Prasyarat (Requirements)
Sebelum memulai, pastikan komputer Anda telah terinstal:
1. **[Git](https://git-scm.com/downloads)** (Untuk mengunduh repositori)
2. **[Python 3.9+](https://www.python.org/downloads/)** (Centang opsi *"Add Python to PATH"* saat instalasi!)

---

## 🚀 Cara Instalasi

**1. Clone Repositori**
Buka terminal/CMD dan jalankan perintah ini untuk mengunduh kode:
```bash
git clone https://github.com/USERNAME/GoogleMapsScraper.git
cd GoogleMapsScraper
```

**2. Instalasi Library yang Dibutuhkan**
Instal library dasar Python yang dibutuhkan scraper:
```bash
pip install -r requirements.txt
```

**3. Instalasi Engine Browser Playwright**
Instal engine Chromium asli dari Playwright:
```bash
playwright install chromium
```

---

## ⚙️ Pengaturan Konfigurasi (`config.json`)
Semua pengaturan utama ada di dalam file `config.json`. Anda bisa mengubahnya menggunakan Notepad atau text editor lainnya.

```json
{
  "google_maps_url": "https://maps.app.goo.gl/URL_SINGKAT_ATAU_PANJANG_ANDA",
  "output_filename": "google_maps_reviews.xlsx",
  "max_reviews": 0,
  "headless": false,
  "scroll_pause": 0.5,
  "max_scroll_attempts_without_new_reviews": 10
}
```
*Catatan: Ubah `max_reviews` menjadi angka tertentu (misal: 100) jika ingin membatasi. Biarkan `0` untuk "Tanpa Batas" (ambil semua).*

---

## 🔐 Langkah Wajib: Login Manual (Anti-Blokir)

Google Maps menyembunyikan ulasan penuh jika Anda tidak *login*. Karena Google mendeteksi bot saat login otomatis, Anda **DIWAJIBKAN** untuk login secara manual **satu kali saja** sebelum menjalankan scraper.

1. Buka PowerShell/CMD Anda, dan jalankan perintah ajaib ini untuk membuka Chrome dengan profil khusus milik scraper:
   ```powershell
   Start-Process "chrome.exe" -ArgumentList "--user-data-dir=`"$(Get-Location)\output\chrome_profile`""
   ```
2. Jendela Google Chrome akan terbuka.
3. Buka situs `https://google.com` dan lakukan **Login ke Akun Google Anda** seperti biasa.
4. Jika sudah berhasil masuk, **Tutup Jendela Chrome Tersebut secara total.**

---

## 🏃‍♂️ Cara Menjalankan Scraper

Setelah instalasi dan proses login selesai, langkah terakhir sangat mudah. Cukup ketik perintah ini di terminal root direktori aplikasi:

```bash
python main.py
```

Scraper akan membuka Google Maps, menavigasi ke halaman target, men-scroll otomatis, dan mengunduh ulasan dengan kecepatan tinggi. 

### 📁 Dimana File Hasilnya?
Setelah proses mencapai tanda *[INFO] Selesai*, Anda dapat menemukan hasilnya di dalam folder `output/`:
- 🟢 **`output/google_maps_reviews.xlsx`** (Untuk dibaca di Excel dengan tabel berwarna rapi)
- 📝 **`output/google_maps_reviews.csv`** (Data mentah ringan)

Selamat menggunakan! 🥳
