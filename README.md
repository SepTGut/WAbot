# Bot WhatsApp Laporan Progress Pekerjaan (Baileys + Spreadsheet)

Bot ini memungkinkan pekerja melapor progress pekerjaan lewat WhatsApp cukup dengan
mengikuti alur menu step-by-step, mengikuti struktur form yang sudah ada.
Hasil laporan otomatis tersimpan ke `laporan-progress.xlsx` (file Excel asli, bisa
langsung dibuka/diformat/di-filter seperti spreadsheet biasa), dan opsional juga
tersinkron ke Google Sheets secara real-time.

## Alur penggunaan (dari sisi pekerja)

1. Kirim pesan apa saja ke bot → muncul daftar gedung, ketik nomornya
2. Pilih sub pekerjaan: **1** Mekanikal / **2** Elektrikal
3. Pilih Titik Lokasi dari daftar yang muncul (sesuai gedung + sub pekerjaan)
4. Ketik progres pekerjaan (teks bebas, contoh: "Pemasangan kipas (3/10) titik")
5. Kirim foto (opsional, ketik **-** untuk lewati)
6. Pilih status: **1** Open / **2** Close
7. Ketik kendala (atau **-** kalau tidak ada)
8. Laporan otomatis tersimpan ✅

**Shortcut:**
- Ketik **menu** kapan saja → kembali ke daftar gedung
- Ketik **batal** → batalkan laporan yang sedang diisi
- Ketik **ulang** → lanjut lapor di lokasi yang sama tanpa pilih gedung dari awal

Tanggal dan jam pengerjaan otomatis tercatat dari waktu pelaporan (WIB).

## 1. Persiapan

Pastikan sudah terinstal **Node.js versi 18 ke atas**. Cek dengan:
```bash
node -v
```

## 2. Siapkan data Titik Lokasi

Isi file **`data-titik-lokasi.xlsx`** (template terlampir) dengan kolom:
- **Nama Gedung** — nama gedung (akan otomatis muncul di menu bot)
- **Sub Pekerjaan** — ketik persis **"Mekanikal"** atau **"Elektrikal"**
- **Titik Lokasi** — nama titik spesifik di gedung & sub pekerjaan itu
- **Keterangan** — opsional

Satu gedung bisa punya banyak baris (titik lokasi berbeda-beda, untuk Mekanikal maupun
Elektrikal). Taruh file ini di folder proyek yang sama dengan `index.js`.

> **Catatan:** Daftar gedung di menu bot dibuat otomatis dari file ini. Kalau ada gedung
> baru, cukup tambahkan di spreadsheet lalu restart bot — tidak perlu edit kode.

## 3. Konfigurasi

Salin file `.env.example` menjadi `.env`, lalu isi sesuai kebutuhan:

```bash
cp .env.example .env
```

| Variabel | Keterangan |
|----------|-----------|
| `WA_PHONE_NUMBER` | Nomor WhatsApp bot (format internasional, contoh: `6281234567890`) |
| `APPS_SCRIPT_URL` | URL Web App Google Apps Script (lihat langkah 6 di bawah) |
| `APPS_SCRIPT_TOKEN` | Token rahasia yang sama dengan di Apps Script |
| `GEMINI_API_KEY` | (Belum dipakai, untuk pengembangan fitur AI nanti) |

## 4. Instalasi

```bash
npm install
```

## 5. Jalankan bot

### Opsi A: Mode Manual (Terminal biasa)
```bash
npm start
```
Akan muncul **QR code** di terminal. Scan menggunakan WhatsApp di HP:
`Setelan > Perangkat Tertaut > Tautkan Perangkat`

### Opsi B: Mode Background / Production dengan PM2
PM2 membuat bot tetap berjalan di background secara otomatis, me-restart otomatis jika crash, dan bisa jalan saat PC/server menyala.

1. **Install PM2 secara global (cukup sekali):**
   ```bash
   npm install -g pm2
   ```
2. **Jalankan bot via PM2:**
   ```bash
   npm run pm2:start
   ```
3. **Perintah penting PM2:**
   ```bash
   npm run pm2:logs     # Lihat log real-time & scan QR code (jika sesi baru)
   npm run pm2:status   # Cek status bot (online/offline)
   npm run pm2:restart  # Restart bot
   npm run pm2:stop     # Matikan bot
   ```
4. *(Opsional)* **Simpan agar otomatis start saat komputer/server reboot:**
   ```bash
   pm2 save
   pm2 startup
   ```

### Opsi C: Mode Container dengan Docker
Docker mengisolasi bot dan memudahkan deployment ke VPS/Server Linux maupun Windows.

1. **Jalankan container:**
   ```bash
   docker compose up -d
   ```
2. **Lihat log & scan QR code (jika belum login):**
   ```bash
   docker compose logs -f
   ```
3. **Perintah penting Docker:**
   ```bash
   docker compose ps        # Cek status container
   docker compose restart   # Restart bot
   docker compose down      # Hentikan bot
   docker compose build     # Rebuild image setelah ada perubahan kode
   ```

> **Catatan Volume Docker:** Semua data penting (`auth_info/`, `laporan-progress.xlsx`, `foto-laporan/`, `bot.log`) dimount langsung ke folder lokal di luar container, jadi sesi login dan data laporan tetap aman dan tidak akan hilang saat container di-restart/di-update.

Gunakan nomor WhatsApp khusus untuk CS (**jangan nomor pribadi utama**), karena ini
memakai library tidak resmi (Baileys) yang berisiko kecil kena pembatasan dari WhatsApp
jika dipakai dengan volume pesan sangat tinggi.

## 6. Cek hasil laporan

Secara default, laporan tersimpan ke file lokal `laporan-progress.xlsx` (kolom:
Waktu Input, Nomor Pengirim, Nama Pengirim, Gedung, Tanggal Pengerjaan, Jam Selesai,
Sub Pekerjaan, Titik Lokasi, Progres, Status, Kendala). Buka langsung dengan Excel
kapan saja.

Foto yang dikirim pekerja **otomatis tersisip langsung sebagai gambar** di kolom
"Foto" pada laporan Excel (bukan cuma nama file) — jadi begitu buka file laporannya,
foto langsung terlihat di baris yang bersangkutan. File aslinya tetap tersimpan juga
di folder `foto-laporan/` sebagai cadangan.

## 7. Log

Bot otomatis menyimpan log ke file `bot.log` di folder proyek. Log ini mencatat semua
pesan masuk, laporan tersimpan, error, dan aktivitas koneksi. Berguna untuk diagnosa
kalau ada masalah.

## 8. (Opsional) Sinkronisasi ke Google Sheets

File Excel lokal cuma ada di komputermu, jadi orang lain tidak bisa lihat update
secara real-time. Untuk itu, laporan bisa juga otomatis ditulis ke Google Sheets
(yang bisa dibuka bareng-bareng dari HP/laptop siapa saja yang diberi akses),
pakai **Google Apps Script** — jauh lebih simpel dari cara Service Account/API key,
tidak perlu Google Cloud Console sama sekali.

1. Buka sheets.google.com, buat spreadsheet baru.
2. Klik **Extensions > Apps Script**.
3. Hapus kode default, tempel isi file **`apps-script.gs`** (disertakan bersama file ini). Script ini otomatis upload foto ke folder Google Drive bernama "Foto Laporan Bot" (dibuat otomatis) dan menampilkannya langsung sebagai gambar di kolom "Foto" pada sheet.
4. Di baris `const TOKEN_RAHASIA = "GANTI_DENGAN_KODE_RAHASIA_BEBAS";`, ganti dengan kode rahasia bebas buatanmu sendiri (bebas, asal diingat).
5. Simpan (Ctrl+S). Klik **Deploy > New deployment**. Pilih tipe **Web app**. Execute as: **Me**. Who has access: **Anyone**. Klik **Deploy**.
6. Google akan minta izin — klik **Authorize access**, pilih akunmu, klik **Advanced > Go to (nama project) (unsafe) > Allow**. (Ini normal karena scriptnya buatan sendiri, belum diverifikasi Google — aman.)
7. Setelah deploy, salin **URL Web App** yang muncul (diawali `https://script.google.com/macros/s/.../exec`).
8. Buka file `.env`, isi `APPS_SCRIPT_URL` dengan URL itu, dan `APPS_SCRIPT_TOKEN` dengan kode rahasia yang sama persis seperti di langkah 4.
9. `npm start` lagi, coba isi 1 laporan — cek apakah muncul otomatis di Google Sheets.

**Kalau langkah ini dilewati/gagal** (belum sempat setup, internet putus, dll), bot otomatis tetap simpan ke Excel lokal seperti biasa — tidak ada laporan yang hilang.

## Catatan penting

- Ini adalah **prototipe**, bukan solusi produksi. Untuk pemakaian resmi jangka panjang
  oleh perusahaan, sebaiknya upgrade ke **WhatsApp Business API** resmi.
- File folder `auth_info/` menyimpan sesi login WhatsApp — jangan dibagikan ke siapa pun.
- Saat ini bot hanya membalas pesan pribadi (bukan pesan di grup) dan hanya pesan teks + foto.
- Bot otomatis mengirim pengingat harian jam 16:00 WIB ke nomor terdaftar yang belum lapor hari itu.
- Untuk mematikan bot dengan aman, tekan **Ctrl+C** di terminal (bot akan menyimpan sesi sebelum keluar).
