/* 轻转 LiteConvert — Service Worker:离线缓存应用外壳 */
const VERSION = 'lite-v1';
const ASSETS = [
  './',
  'index.html',
  'manifest.webmanifest',
  'assets/app.css',
  'assets/app.js',
  'assets/encoders.js',
  'assets/docxrender.js',
  'assets/pptxrender.js',
  'assets/converters.js',
  'vendor/jszip.min.js',
  'vendor/pdf-lib.min.js',
  'vendor/pdf.min.mjs',
  'vendor/pdf.worker.min.mjs',
  'vendor/docx.iife.js',
  'vendor/pptxgen.bundle.js',
  'icons/icon-192.png',
  'icons/icon-512.png',
  'icons/apple-touch-icon.png',
  'icons/favicon-32.png'
];

self.addEventListener('install', (e) => {
  e.waitUntil(caches.open(VERSION).then((c) => c.addAll(ASSETS)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', (e) => {
  e.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(keys.filter((k) => k !== VERSION).map((k) => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', (e) => {
  if (e.request.method !== 'GET') return;
  e.respondWith(
    caches.match(e.request, { ignoreSearch: true }).then((hit) => hit ||
      fetch(e.request).then((res) => {
        if (res.ok && new URL(e.request.url).origin === location.origin) {
          const clone = res.clone();
          caches.open(VERSION).then((c) => c.put(e.request, clone));
        }
        return res;
      })
    )
  );
});
