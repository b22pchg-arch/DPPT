const CACHE_NAME = 'scada-load-forecast-pwa-lv9_8_1-20260612-01';
const APP_ASSETS = [
  './',
  './index.html',
  './app.js',
  './workflow_lv9.js',
  './manifest.webmanifest',
  './favicon.ico',
  './favicon.svg',
  './sample_load_data.csv',
  './sample_load_data.xlsx',
  './sample_load_data_lv6_chidanh.csv',
  './sample_load_data_lv6_chidanh.xlsx',
  './thresholds_lv6.csv',
  './expected_operation_events_lv6.csv',
  './README_sample_lv6.txt',
  './libs/pako.min.js',
  './libs/xlsx.full.min.js',
  './libs/sheetjs-xlsx-lite.js',
  './icons/icon-72.png',
  './icons/icon-96.png',
  './icons/icon-128.png',
  './icons/icon-144.png',
  './icons/icon-152.png',
  './icons/icon-180.png',
  './icons/icon-192.png',
  './icons/icon-384.png',
  './icons/icon-512.png',
  './icons/icon-512-maskable.png'
];

self.addEventListener('install', event => {
  event.waitUntil(caches.open(CACHE_NAME).then(cache => cache.addAll(APP_ASSETS)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(keys => Promise.all(keys.filter(k => k !== CACHE_NAME).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', event => {
  if (event.request.method !== 'GET') return;
  event.respondWith(
    caches.match(event.request).then(cached => {
      if (cached) return cached;
      return fetch(event.request).then(response => {
        const copy = response.clone();
        caches.open(CACHE_NAME).then(cache => cache.put(event.request, copy));
        return response;
      }).catch(() => caches.match('./index.html'));
    })
  );
});
