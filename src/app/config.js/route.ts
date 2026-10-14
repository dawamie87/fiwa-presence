/**
 * Konfigurasi aplikasi dari Environment Variables (Vercel > Settings > Environment Variables).
 * Dulu URL & kunci Supabase ditulis langsung di template Blogger; sekarang cukup diubah di Vercel.
 * Kunci yang dipakai adalah kunci PUBLIK (anon / publishable). Jangan pernah mengisi kunci service_role.
 */
export const dynamic = 'force-dynamic';

export function GET() {
  const url = (process.env.NEXT_PUBLIC_SUPABASE_URL || '').trim().replace(/\/+$/, '');
  const key = (process.env.NEXT_PUBLIC_SUPABASE_ANON_KEY || '').trim();
  if (/service_role|sb_secret_/i.test(key)) {
    return new Response('console.error("FIWA: kunci Supabase yang diisi adalah kunci RAHASIA. Ganti dengan anon/publishable key.");window.FIWA_CONFIG={};',
      { headers: { 'Content-Type': 'application/javascript; charset=utf-8', 'Cache-Control': 'no-store' } });
  }
  const cfg = {
    url, key,
    preview: process.env.NEXT_PUBLIC_PREVIEW === 'true',
    logo: process.env.NEXT_PUBLIC_LOGO_URL || undefined,
  };
  return new Response(`window.FIWA_CONFIG=${JSON.stringify(cfg)};`, {
    headers: { 'Content-Type': 'application/javascript; charset=utf-8', 'Cache-Control': 'public, max-age=60, must-revalidate' },
  });
}
