"""Ekstrak aplikasi FIWA Presence dari template Blogger menjadi aset statis untuk Next.js."""
import re, sys, hashlib, pathlib
src, out = pathlib.Path(sys.argv[1]), pathlib.Path(sys.argv[2])
L = src.read_text(encoding='utf-8').split('\n')
seg = lambda a, b: '\n'.join(L[a-1:b])          # baris 1-based, inklusif

css = seg(40, 1837)                              # isi <b:skin><![CDATA[ ... ]]>
markup = seg(1842, 1876)                         # splash + noscript + #sk-root
app = seg(2747, 6398)                            # IIFE aplikasi
splash = seg(6404, 6414)

assert css.lstrip().startswith('#sk-root,#sk-root *'), css[:60]
assert app.lstrip().startswith('(function(){') and app.rstrip().endswith('})();'), (app[:40], app[-40:])
assert 'sk-splash' in splash

# XML self-closing pada elemen non-void (<div/>, <span/>, <i/>...) tidak sah di HTML -> tulis pasangan tag
VOID = {'br','img','input','meta','link','hr'}
def fix(m):
    tag, attrs = m.group(1), m.group(2)
    return m.group(0) if tag.lower() in VOID else f'<{tag}{attrs.rstrip()}></{tag}>'
markup = re.sub(r'<([a-zA-Z][\w-]*)((?:\s+[^<>]*?)?)\s*/>', fix, markup)
assert '/>' not in re.sub(r'<(br|img)[^>]*/>', '', markup), 'masih ada self-closing'

# Peningkatan: service worker asli (Blogger tidak bisa menyimpan sw.js)
n = app.count("const PWA_SW_URL='';"); assert n == 1
app = app.replace("const PWA_SW_URL='';", "const PWA_SW_URL='/sw.js';")
# Manifest & ikon statis sudah disediakan; manifest blob hanya dipakai bila manifest statis tidak ada
old = "  pwaIcons().then(pwaApply).catch("
assert app.count(old) == 1
app = app.replace(old, "  if(!document.querySelector('link[rel=manifest][data-static]'))pwaIcons().then(pwaApply).catch(")

h = lambda s: hashlib.sha256(s.encode()).hexdigest()[:10]
d = out / 'public' / 'fiwa'; d.mkdir(parents=True, exist_ok=True)
for f in d.glob('app.*'): f.unlink()
cssf, jsf = f'app.{h(css)}.css', f'app.{h(app)}.js'
(d / cssf).write_text(css.strip() + '\n', encoding='utf-8')
(d / jsf).write_text("/* FIWA Presence - kode aplikasi (dipindahkan dari template Blogger). */\n" + app.strip() + '\n', encoding='utf-8')

html = f"""<!doctype html>
<html lang="id">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
  <title>FIWA Presence - Sistem Presensi Pegawai</title>
  <meta name="description" content="Sistem presensi pegawai Fitrah Islamic World Academy">
  <meta name="theme-color" content="#0f766e">
  <meta name="mobile-web-app-capable" content="yes">
  <meta name="apple-mobile-web-app-capable" content="yes">
  <meta name="apple-mobile-web-app-title" content="FIWA Presence">
  <meta name="apple-mobile-web-app-status-bar-style" content="default">
  <meta name="application-name" content="FIWA Presence">
  <link rel="manifest" href="/manifest.webmanifest" data-static>
  <link rel="icon" href="/icons/icon-192.png" type="image/png">
  <link rel="apple-touch-icon" href="/icons/apple-touch-icon.png">
  <link rel="preconnect" href="https://i.ibb.co.com">
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap">
  <link rel="stylesheet" href="/fiwa/{cssf}">
</head>
<body>
{markup}

  <!-- Konfigurasi dari environment variable Vercel (dibuat oleh /config.js) -->
  <script src="/config.js"></script>
  <script>
const FIWA_CFG = window.FIWA_CONFIG || {{}};
const IS_PREVIEW = !!FIWA_CFG.preview; // true = mode uji coba dengan data contoh, tanpa Supabase
const SUPABASE_URL = FIWA_CFG.url || "";
const SUPABASE_ANON_KEY = FIWA_CFG.key || "ANON_KEY_ANDA";
const SUPABASE_ESM = "https://cdn.jsdelivr.net/npm/@supabase/supabase-js@2.45.4/+esm";
const XLSX_URL = "https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js";
const APP_NAME = "FIWA Presence";
const APP_TAGLINE = "Sistem Presensi Pegawai";
const LOGO_URL = FIWA_CFG.logo || "https://i.ibb.co.com/hxTc3H6D/Logo-80x80-01.webp";
function logoImg(){{return '<img src="'+LOGO_URL+'" alt="Logo Fitrah Islamic World Academy" width="80" height="80" decoding="async" referrerpolicy="no-referrer">'}}
  </script>
  <script src="/fiwa/{jsf}"></script>
  <script>
{splash}
  </script>
</body>
</html>
"""
(d / 'index.html').write_text(html, encoding='utf-8')
print(cssf, len(css)); print(jsf, len(app)); print('index.html', len(html))
