# Skema Database — Inventory Gudang

> ⚠️ Dokumen ini disusun berdasarkan query yang dipakai di `app.py`, bukan hasil export langsung dari database. Untuk skema yang 100% akurat, jalankan `pg_dump --schema-only` di Supabase atau screenshot struktur tabel dari dashboard, lalu lampirkan di sini.

## `buku`
Data induk buku.
| Kolom | Keterangan |
|---|---|
| id | PK |
| isbn | unik |
| judul, penulis, penerbit | |
| stok | jumlah fisik saat ini |
| stok_minimum | ambang batas notifikasi stok kritis |
| jumlah_rencana | rencana total yang akan didistribusikan |
| tanggal_masuk | **terkunci** ke kedatangan pertama kali stok terisi (lihat logika di `transaksi_masuk`) |
| catatan | catatan bebas |
| lokasi_rak | lokasi fisik di gudang |
| panjang_cm, lebar_cm, tinggi_cm | opsional, untuk fitur Simulasi Packing |
| updated_at | |

## `transaksi`
Riwayat mutasi masuk/keluar per buku.
| Kolom | Keterangan |
|---|---|
| id | PK |
| buku_id | FK → buku |
| tipe | `'masuk'` atau `'keluar'` |
| jumlah | |
| tanggal | tanggal transaksi (bisa beda dari tanggal input) |
| pihak_terkait, keterangan | |
| tujuan_id | FK → tujuan (khusus tipe keluar, opsional) |
| user_id | FK → users, siapa yang input |
| created_at | |

## `tujuan`
Daftar tujuan distribusi.
| Kolom | Keterangan |
|---|---|
| id | PK |
| nama | |
| desa_kelurahan, kecamatan, kabupaten_kota, provinsi | |

## `distribusi_rencana`
Rencana pengiriman per buku per tujuan (many-to-many buku↔tujuan).
| Kolom | Keterangan |
|---|---|
| id | PK |
| tujuan_id | FK → tujuan |
| buku_id | FK → buku |
| jumlah_rencana | target jumlah yang akan dikirim |

## `users`
| Kolom | Keterangan |
|---|---|
| id | PK |
| username | unik |
| password_hash | Werkzeug hash |
| nama_lengkap | |
| role | `admin` / role staf non-viewer / `viewer` |
| is_active | untuk nonaktifkan akun tanpa hapus |
| failed_attempts, locked_until | lockout percobaan login gagal (dipakai juga untuk lockout 2FA) |
| totp_secret, totp_enabled | 2FA |

## `user_backup_codes`
Backup code 2FA, sekali pakai.
| Kolom | Keterangan |
|---|---|
| id | PK |
| user_id | FK → users (ON DELETE CASCADE) |
| kode_hash | hash dari kode (bukan plaintext) |
| digunakan | boolean, ditandai setelah dipakai |

## `activity_log`
Jejak audit aksi penting (tambah/edit/hapus buku, hapus paksa, aktivasi 2FA, dll).
| Kolom | Keterangan |
|---|---|
| id | PK |
| aksi, detail | deskripsi aksi |
| buku_id | FK → buku (opsional, nullable) |
| created_at | |

## `ukuran_kardus`
Ukuran kardus untuk fitur Simulasi Packing.
| Kolom | Keterangan |
|---|---|
| id | PK |
| nama | misal "Kardus L" |
| panjang_cm, lebar_cm, tinggi_cm | |
| aktif | boolean, kardus nonaktif tidak dipakai dalam simulasi |

## `import_staging`
Tabel sementara untuk proses import massal Excel (Buku Masuk/rencana distribusi) sebelum dikonfirmasi.