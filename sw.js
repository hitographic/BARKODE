const CACHE_NAME = 'scanner-pwa-v1';

// Install Service Worker
self.addEventListener('install', event => {
    event.waitUntil(
        caches.open(CACHE_NAME).then(cache => {
            return cache.addAll(['./', './index.html', './manifest.json']);
        })
    );
});

// Fetch Intercept (Network First Strategy)
// Selalu usahakan ambil data dari internet, kalau offline baru ambil dari cache
self.addEventListener('fetch', event => {
    event.respondWith(
        fetch(event.request).catch(() => caches.match(event.request))
    );
});