/* CTPL · Service worker: guarda los archivos del sitio para que la app abra rápido.
   Las consultas al servidor (Apps Script) nunca se guardan. */
const VERSION = 'ctpl-v5';
const ARCHIVOS = ['./', './index.html', './config.js', './css/app.css', './js/comun.js', './js/carnet.js', './js/llavero.js',
  './js/jsQR.js', './img/plantilla_carnet.jpg', './img/icon-192.png', './img/icon-512.png', './manifest.webmanifest'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(VERSION).then(c => c.addAll(ARCHIVOS)).then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== VERSION).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});
self.addEventListener('fetch', e => {
  const u = new URL(e.request.url);
  if (e.request.method !== 'GET' || u.origin !== location.origin) return;   // API y fuentes: directo a la red
  // Primero la red (para recibir actualizaciones); si no hay conexión, la copia guardada
  e.respondWith(fetch(e.request, { cache: 'no-cache' }).then(r => {
    const copia = r.clone(); caches.open(VERSION).then(c => c.put(e.request, copia)); return r;
  }).catch(() => caches.match(e.request, { ignoreSearch: true })));
});
