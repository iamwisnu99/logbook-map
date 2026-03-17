# Logbook MAP Downloader

Program otomatisasi berbasis Python dan Selenium untuk mengunduh, memproses, dan merapikan file Logbook Bulanan secara otomatis melalui integrasi Bot Telegram.

## Fitur Utama

- **Automated Scraping**: Login dan navigasi otomatis ke portal laporan menggunakan Selenium.
- **Multi-Account Support**: Mendukung pemrosesan banyak akun sekaligus dari file Excel.
- **Telegram Integration**: Kontrol penuh program (mulai, berhenti, konfigurasi) via Bot Telegram.
- **PDF Processing**: Otomatis menghapus password PDF dan melakukan penamaan ulang (renaming) sesuai standar.
- **Auto-Organization**: Folder output tersusun rapi berdasarkan Tahun dan Bulan.
- **Notification System**: Laporan progress real-time dikirimkan langsung ke Telegram.
- **ZIP Export**: Pilihan untuk mengirim hasil unduhan dalam bentuk file ZIP melalui Telegram.

## Prasyarat

Sebelum menjalankan program, pastikan sistem Anda telah terpasang:

1. **Python 3.10+**
2. **Google Chrome** (Versi terbaru)
3. **ChromeDriver** (Akan dikelola otomatis oleh `webdriver-manager`)

## Instalasi

1. Clone atau salin folder proyek ini ke komputer Anda.
2. Buka terminal/command prompt di direktori proyek.
3. Instal semua dependensi yang dibutuhkan:
   ```bash
   pip install -r requirements.txt
   ```

## Konfigurasi

### 1. File `.env`

Buat file bernama `.env` di direktori utama dan isi dengan token bot Telegram Anda:

```env
TELEGRAM_BOT_TOKEN=your_bot_token_here
TELEGRAM_CHAT_ID=your_chat_id_here
```

### 2. File `akun.xlsx`

Siapkan data akun dalam format Excel dengan kolom berikut:

- **Nama**: Nama pangkalan/identitas akun.
- **Email**: Email atau nomor ponsel login.
- **PIN**: PIN login akun.
- **MID/Password:** MID Pangkalan.

### 3. Path Penyimpanan

Buka file `main.py` dan sesuaikan variabel `BASE_LOGBOOK_PATH` pada baris 30:

```python
BASE_LOGBOOK_PATH = r"D:\Documents\LOGBOOK MAP" # === Sesuaikan lokasi anda ===
```

## Cara Penggunaan

1. Jalankan program dengan perintah:
   ```bash
   python app.py
   ```
2. Buka Bot Telegram Anda dan ketik `/start`.
3. Klik tombol **"Start Program"**.
4. Ikuti instruksi bot untuk konfigurasi:
   - **Pilih Bulan**: Masukkan nama bulan (misal: `Januari, Februari`).
   - **Pilih Tahun**: Masukkan tahun (misal: `2025`) atau ketik `Saat Ini`.
   - **Pilih Akun**: Masukkan nomor urut akun (misal: `1,2,5`), `all` untuk semua, atau `#1` untuk kecuali nomor 1.
   - **Kirim Telegram**: Pilih `Ya` jika ingin file ZIP dikirim ke Telegram setelah selesai.
5. Program akan berjalan otomatis di latar belakang. Anda dapat memantau progress melalui pesan yang dikirim bot.
6. Untuk menghentikan program di tengah jalan, ketik **"Stop"** di Telegram.

## 📂 Struktur Output

Hasil download akan disimpan dengan struktur:

```text
BASE_LOGBOOK_PATH/
└── [Tahun]/
    └── [Bulan Tahun]/
        └── [Nama] - [Bulan Tahun].pdf
```

## 📄 Lisensi dan Kebijakan

Program ini diproduksi dan dimiliki oleh **PT PRIMADEV DIGITAL TECHNOLOGY**.
Untuk rincian lebih lanjut, silakan baca dokumen berikut:

- [LICENSE](LICENSE)
- [SYARAT_DAN_KETENTUAN.md](SYARAT_DAN_KETENTUAN.md)

---

© 2025 PT PRIMADEV DIGITAL TECHNOLOGY. All Rights Reserved.
