// Service worker mínimo: solo habilita "agregar a inicio" en el celular.
// La app necesita al servidor (misma red Wi-Fi) para funcionar, así que no
// cacheamos datos — solo dejamos pasar las peticiones a la red.
self.addEventListener('install', (e) => self.skipWaiting());
self.addEventListener('activate', (e) => self.clients.claim());
self.addEventListener('fetch', (e) => {
  e.respondWith(fetch(e.request).catch(() => caches.match(e.request)));
});
