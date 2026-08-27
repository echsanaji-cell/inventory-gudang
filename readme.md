# Inventory Gudang — CV. Laras Sejati

Aplikasi pendataan keluar-masuk buku gudang: stok, distribusi ke tujuan, mapping lokasi rak, dan simulasi packing. Dibangun dengan Flask + Supabase (Postgres).

## Tech Stack
- **Backend**: Flask (Python), Gunicorn (production server)
- **Database**: Supabase (PostgreSQL), diakses lewat psycopg2
- **Frontend**: Jinja2 templates + Bootstrap 5, tema custom "label laci katalog"
- **Autentikasi**: session-based, password di-hash (Werkzeug), 2FA (TOTP) opsional per akun + backup code
- **Integrasi eksternal**: Google Books API (lookup ISBN), Google Sheets API (sync), Gmail SMTP (notifikasi email), python-barcode (cetak label), html5-qrcode (scan kamera)
- **Deploy**: Render (lihat env var `RENDER`)

## Cara Jalankan Lokal
```bash
python3 -m venv venv
source venv/bin/activate          # Mac/Linux
pip install -r requirements.txt --break-system-packages
cp .env.example .env               # lalu isi semua nilainya, lihat bagian Environment Variables
python3 app.py
```
Buka `http://localhost:5000`.

## Environment Variables
Lihat `.env.example` untuk daftar lengkap variabel yang dibutuhkan beserta penjelasannya. **Jangan pernah commit file `.env` ke git** — sudah ada di `.gitignore`.

## Struktur Folder
app.py # semua route Flask
db.py # koneksi database (get_db_connection)
templates/
base.html # layout utama + navbar
login.html, login_verifikasi_2fa.html
buku/ # form, list, detail, mapping area, distribusi
transaksi/ # buku masuk, buku keluar
tujuan/ # daftar tujuan, detail, rencana distribusi, simulasi packing
admin/ # kelola penerbit, kelola kardus
akun/ # pengaturan 2FA
scan_mobile.html # halaman scan cepat untuk staf gudang (HP)
static/
css/style.css
js/scanner.js # helper scan barcode (kamera + scanner fisik)
docs/
DATABASE_SCHEMA.md
OPERASIONAL.md


## Role Pengguna
- **admin** — akses penuh, termasuk Kelola User, Kelola Penerbit, Kelola Kardus, Activity Log, Hapus Paksa
- **(role staf/non-viewer)** — bisa input transaksi masuk/keluar, tapi tidak akses menu admin
- **viewer** — hanya lihat data, tidak bisa input transaksi (`viewer_blocked` decorator)

## Fitur yang Sudah Ada
- Pendataan buku + lokasi rak + ukuran fisik (opsional, untuk simulasi packing)
- Transaksi masuk/keluar dengan scan barcode (kamera & scanner fisik)
- Deteksi ISBN duplikat real-time
- Tanggal masuk terkunci ke kedatangan pertama (untuk kasus kedatangan sebagian/bertahap)
- Distribusi rencana per Tujuan, verifikasi pengiriman, export Excel/PDF
- Simulasi Packing (estimasi jumlah kardus berdasar volume buku — perlu data ukuran buku diisi dulu)
- Scan Mobile — halaman ringkas untuk staf gudang catat transaksi dari HP
- Notifikasi email stok kritis, sync ke Google Sheets, backup manual (Excel/JSON)
- 2FA (TOTP) opsional per akun + backup code recovery + lockout percobaan gagal
- Activity Log — jejak audit semua aksi penting
- Hapus Paksa — admin only, hapus buku beserta riwayat transaksinya (dengan konfirmasi ketik ulang judul)

## Belum Dikerjakan / Di-skip
- Perbandingan ISBN otomatis dengan Perpustakaan Nasional — di-skip, belum ada akses API resmi
- Retur/write-off buku, dashboard analitik, reminder verifikasi tertunda — ide untuk dikembangkan lagi nanti

## Insiden Keamanan (untuk konteks sejarah)
Agustus 2026: `.env` sempat ter-commit ke git (beberapa commit awal). Semua kredensial sudah dirotate dan riwayat git sudah dibersihkan pakai `git-filter-repo`. Sejak itu `.env` sudah benar di `.gitignore` dan tidak pernah di-commit lagi.