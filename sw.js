/* ============================================================
   SERVICE WORKER — Flores El Trigal · Tablero de Ejecución Semanal
   Estrategia: cache-first para el shell de la app y las librerías
   externas (Bootstrap, SheetJS, fuentes), con actualización en
   segundo plano ("stale-while-revalidate") para que la app siga
   funcionando sin internet una vez visitada al menos una vez.
   ============================================================ */

const CACHE_VERSION = 'v1';
const CACHE_NAME = `trigal-tablero-${CACHE_VERSION}`;

/* Archivos propios de la app — se precargan al instalar el SW */
const APP_SHELL = [
  './',
  './index.html',
  './style.css',
  './app.js',
  './manifest.json',
  './icons/icon-72.png',
  './icons/icon-96.png',
  './icons/icon-128.png',
  './icons/icon-144.png',
  './icons/icon-152.png',
  './icons/icon-192.png',
  './icons/icon-384.png',
  './icons/icon-512.png',
  './icons/favicon.png',
  './icons/apple-touch-icon.png',
];

/* Librerías externas (CDN) usadas por la app */
const VENDOR_ASSETS = [
  'https://cdnjs.cloudflare.com/ajax/libs/bootstrap/5.3.3/css/bootstrap.min.css',
  'https://cdnjs.cloudflare.com/ajax/libs/bootstrap-icons/1.11.3/font/bootstrap-icons.min.css',
  'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js',
  'https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js',
];

self.addEventListener('install', (event) => {
  self.skipWaiting();
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => {
      // Los recursos propios deben poder precargarse sin fallar la instalación;
      // los de CDN se agregan "best effort" (si falla uno no rompe el SW).
      return cache.addAll(APP_SHELL).then(() =>
        Promise.allSettled(
          VENDOR_ASSETS.map((url) =>
            fetch(url, { mode: 'cors' })
              .then((res) => (res.ok ? cache.put(url, res) : null))
              .catch(() => null)
          )
        )
      );
    })
  );
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(
        keys
          .filter((key) => key !== CACHE_NAME)
          .map((key) => caches.delete(key))
      )
    ).then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', (event) => {
  const req = event.request;

  // Solo interceptamos GET; todo lo demás pasa directo a la red.
  if (req.method !== 'GET') return;

  event.respondWith(
    caches.match(req).then((cached) => {
      const networkFetch = fetch(req)
        .then((res) => {
          // Solo cacheamos respuestas válidas (evita opaque/basic errores)
          if (res && res.ok) {
            const resClone = res.clone();
            caches.open(CACHE_NAME).then((cache) => cache.put(req, resClone));
          }
          return res;
        })
        .catch(() => cached); // sin red: usa lo cacheado si existe

      // Stale-while-revalidate: responde con cache al toque si existe,
      // y de fondo actualiza el cache con la red.
      return cached || networkFetch;
    })
  );
});
