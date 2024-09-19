const CACHE_NAME = 'coaching-app-cache-v1';
const urlsToCache = [
  'index.html',
  'styles.css',
  'app.js',
  'manifest.json',
  'emails.json' // Ensure this file is available
];

// Install the service worker
self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then(cache => {
        return cache.addAll(urlsToCache);
      })
  );
});

// Fetch resources
self.addEventListener('fetch', event => {
  event.respondWith(
    caches.match(event.request)
      .then(response => {
        // Return cached resource or fetch from network
        return response || fetch(event.request);
      })
  );
});
