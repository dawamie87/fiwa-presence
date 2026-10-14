/* Aplikasi utama disajikan dari public/fiwa/index.html (lihat next.config.mjs).
   Layout ini hanya dipakai halaman bawaan Next.js seperti 404. */
export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (<html lang="id"><body style={{ fontFamily: 'system-ui, sans-serif' }}>{children}</body></html>);
}
