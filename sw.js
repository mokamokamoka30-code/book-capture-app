const CACHE_NAME = 'book-capture-v1';
const CORE = ['./', './index.html', './style.css', './app.js', './manifest.webmanifest'];
self.addEventListener('install', event => event.waitUntil(caches.open(CACHE_NAME).then(cache => cache.addAll(CORE))));
self.addEventListener('activate', event => event.waitUntil(self.clients.claim()));
self.addEventListener('fetch', event => {
  if (event.request.method !== 'GET' || new URL(event.request.url).origin !== self.location.origin) return;
  event.respondWith(caches.match(event.request).then(cached => cached || fetch(event.request)));
});
