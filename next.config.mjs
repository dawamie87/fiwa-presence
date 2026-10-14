/** @type {import('next').NextConfig} */
const security = [
  { key: 'X-Content-Type-Options', value: 'nosniff' },
  { key: 'X-Frame-Options', value: 'DENY' },
  { key: 'Referrer-Policy', value: 'strict-origin-when-cross-origin' },
  { key: 'Permissions-Policy', value: 'geolocation=(self), camera=(self), microphone=(), payment=()' },
  { key: 'Strict-Transport-Security', value: 'max-age=31536000; includeSubDomains' },
];

const nextConfig = {
  poweredByHeader: false,
  // Halaman utama = aplikasi FIWA Presence (dipindahkan utuh dari template Blogger)
  async rewrites() {
    return { beforeFiles: [{ source: '/', destination: '/fiwa/index.html' }] };
  },
  async headers() {
    return [
      { source: '/:path*', headers: security },
      // Aset bernama hash (app.<hash>.js/css): aman di-cache lama, nama berubah tiap ada perubahan
      { source: '/fiwa/app.:hash.:ext', headers: [{ key: 'Cache-Control', value: 'public, max-age=31536000, immutable' }] },
      { source: '/', headers: [{ key: 'Cache-Control', value: 'no-cache' }] },
      { source: '/fiwa/index.html', headers: [{ key: 'Cache-Control', value: 'no-cache' }] },
      { source: '/sw.js', headers: [{ key: 'Cache-Control', value: 'no-cache' }, { key: 'Service-Worker-Allowed', value: '/' }] },
    ];
  },
};
export default nextConfig;
