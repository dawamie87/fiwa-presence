# Daftar uji singkat FIWA Presence (perlu Supabase + browser)

Siapkan: salin `.env.example` menjadi `.env.local` dan isi nilai proyek Anda; `npm install && npm run build`.
Uji dengan 3 akun: `admin_sistem`, `manajer_hrd`, `staf_hrd`, plus 1 pegawai biasa dan 1 akun Google tak terdaftar.

1. **Login Google**: akun terdaftar masuk; tak terdaftar ditolak (`/login?denied=1`); pegawai biasa tidak bisa membuka halaman admin (langsung ketik URL).
2. **Role & RLS** (uji lewat aplikasi, bukan hanya UI): `staf_hrd` tidak bisa menyimpan Pengaturan, mengunci/membuka periode, atau membaca Log Audit; pegawai biasa tidak bisa membaca/menulis tabel admin.
3. **RPC**: `adm_create_employees`, `adm_decide_final`, `adm_decide_cancel`, `adm_mark_applied`, `adm_swap_decide`, `adm_lock_period`, `adm_reopen_period` ada, berparameter sesuai kode, dan menolak role tak berhak. (Tiga yang terakhir belum ada sumbernya di `supabase/`.)
4. **CRUD**: pegawai, departemen, akun admin, jadwal normal, shift, libur, jenis pengajuan, tahun ajaran: tambah, ubah, hapus; pastikan baris benar-benar tersimpan setelah muat ulang.
5. **Persetujuan**: alur L1 -> HRD, penolakan, pembatalan, override; cek saldo cuti dan rekap ikut berubah.
6. **Impor fingerprint**: unggah contoh file, duplikat dilewati, ID tak dikenal tercatat, batalkan batch; log pada periode terkunci.
7. **Penutupan periode**: kunci bulan uji -> status berubah di database; koreksi manual, impor, proses ulang rekap, dan edit shift pada bulan itu ditolak/dilewati; buka kembali dengan alasan; riwayat tampil.
8. **Pengaturan**: simpan, muat ulang, nilai sama; ubah zona waktu / periode saldo / batas lembur lalu cek Rekap dan Saldo Cuti; role non-HRD hanya bisa melihat.
9. **Log Audit**: hanya `manajer_hrd`/`admin_sistem`; catatan dari langkah 4-8 muncul; pencarian dan dialog Detail berfungsi.
