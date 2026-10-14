# FIWA Presence

Sistem presensi pegawai Fitrah Islamic World Academy, berjalan di Next.js (Vercel) dengan database Supabase.

Seluruh aplikasi dari template Blogger (semua menu admin dan pegawai, tampilan, mode uji coba) dipindahkan utuh,
ditambah hal-hal yang tidak bisa dilakukan di Blogger:

| Peningkatan | Keterangan |
|---|---|
| Aplikasi HP sungguhan (PWA) | Manifest & service worker asli: bisa dipasang di layar utama, terbuka lebih cepat |
| Konfigurasi di Vercel | URL & kunci Supabase diatur lewat Environment Variables, tidak ditulis di kode |
| Perlindungan kunci rahasia | Bila kunci `service_role` tidak sengaja diisi, aplikasi menolak memakainya |
| Header keamanan | Anti clickjacking, HSTS, izin lokasi hanya untuk situs sendiri |
| Cache cerdas | Aset ber-hash di-cache lama; pembaruan langsung terlihat tanpa hapus cache |
| Domain sendiri | Bisa memakai `presensi.fiwa.sch.id` dengan HTTPS otomatis |

## Struktur

```
public/
  fiwa/index.html        Halaman aplikasi (disajikan di alamat utama "/")
  fiwa/app.<hash>.js     Kode aplikasi (dari template Blogger)
  fiwa/app.<hash>.css    Tampilan aplikasi
  sw.js                  Service worker (PWA)
  manifest.webmanifest   Manifest PWA
  icons/                 Ikon aplikasi (ganti dengan logo lembaga, ukuran sama)
src/app/config.js/       Mengirim konfigurasi dari Environment Variables ke aplikasi
next.config.mjs          Alamat utama + header keamanan + cache
tools/extract-blogger.py Alat pindah ulang bila template Blogger diubah
```

## Environment Variables (Vercel > Settings > Environment Variables)

| Nama | Isi |
|---|---|
| `NEXT_PUBLIC_SUPABASE_URL` | Project URL Supabase |
| `NEXT_PUBLIC_SUPABASE_ANON_KEY` | Kunci anon / publishable (bukan service_role) |
| `NEXT_PUBLIC_PREVIEW` | Opsional: `true` = mode uji coba dengan data contoh |
| `NEXT_PUBLIC_LOGO_URL` | Opsional: alamat gambar logo |

## Menjalankan di komputer

```bash
npm install
cp .env.example .env.local   # isi nilainya
npm run dev                  # http://localhost:3000
```

## Mengganti ikon dengan logo lembaga

Ganti file di `public/icons/` dengan PNG logo berukuran sama (192x192, 512x512, 180x180;
`icon-maskable-512.png` beri ruang kosong ±20% di tepi). Nama file harus tetap.

## Mengubah aplikasi

Edit `public/fiwa/app.<hash>.js` atau `.css`. Setiap kali mengubah file ini, **naikkan VERSION di `public/sw.js`**
(mis. `fiwa-v2`) agar HP karyawan mengambil versi baru.
