# SOP Operasional — Inventory Gudang

## Peran & Akses
- **Admin**: kelola user, penerbit, ukuran kardus, hapus buku (termasuk hapus paksa), lihat activity log
- **Staf gudang**: input transaksi masuk/keluar, tambah/edit buku, kelola tujuan & distribusi
- **Viewer**: lihat data saja, tidak bisa input apapun

## Alur Kerja Harian
1. **Buku baru datang** → menu Data Buku → Tambah Buku (isi ISBN, judul, lokasi rak; ukuran fisik opsional untuk simulasi packing nanti)
2. **Catat stok masuk** → menu Transaksi → Buku Masuk (scan barcode atau ketik ISBN manual), atau pakai **Scan Cepat (Mobile)** dari HP kalau sedang di lantai gudang
3. **Rencana distribusi ke tujuan** → menu Distribusi → Tujuan → pilih tujuan → Import Rencana (upload Excel) atau input manual
4. **Kirim buku ke tujuan** → Transaksi → Buku Keluar, pilih tujuan tujuan pengiriman
5. **Verifikasi pengiriman selesai** → menu Distribusi → Verifikasi

## Kalau HP Admin Hilang (2FA)
Pakai salah satu **backup code** yang sudah disimpan (dari saat aktivasi 2FA) di halaman login. Setelah berhasil masuk, segera buat ulang backup code baru dan setup ulang authenticator app di HP baru.

## Kalau Butuh Hapus Buku Data Testing/Salah Input
Kalau muncul error "masih ada transaksi terkait", buku itu sudah pernah ada transaksi. Kalau memang cuma data testing, admin bisa pakai **Hapus Paksa** di halaman Detail Buku (akan minta ketik ulang judul buku untuk konfirmasi karena ini menghapus riwayat transaksi juga).

## Kontak & Tanggung Jawab
| Nama | Peran | Kontak |
|---|---|---|
| admin | Admin utama | Bayu |
| krisman | Staf  | krisman |
| mella | Staf  | mella |
| kemal | Staf  | kemal |
| mella2 | Admin  | mella |
| mella2 | Admin  | Adam |

## Backup Data
- Backup manual (Excel/JSON) tersedia di menu Admin — jalankan minimal sebulan sekali, simpan di Google Drive terpisah dari Supabase
- Cek juga pengaturan backup otomatis Supabase di dashboard (Settings → Database → Backups)