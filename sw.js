const CACHE_NAME = 'tsc80-pwa-v1.1';
const PRECACHE_ASSETS = [
  './',
  './index.html',
  './pins.json',
  './manifest.json',
  './icon-192.png',
  './icon-512.png',
  './apple-touch-icon.png',
  './favicon.png',
  './favicon.ico'
];

// 1. Install Service Worker & Pre-cache Core Files
self.addEventListener('install', event => {
  self.skipWaiting();
  event.waitUntil(
    caches.open(CACHE_NAME).then(cache => {
      return Promise.allSettled(
        PRECACHE_ASSETS.map(url => cache.add(new Request(url, { cache: 'reload' })))
      );
    })
  );
});

// 2. Activate & Remove Outdated Caches
self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(keys => {
      return Promise.all(
        keys.map(key => {
          if (key !== CACHE_NAME) {
            return caches.delete(key);
          }
        })
      );
    }).then(() => self.clients.claim())
  );
});

// 3. Fetch Handler: Offline Resilience
self.addEventListener('fetch', event => {
  const req = event.request;
  const url = new URL(req.url);

  // Ignore non-GET requests and external chrome-extension / analytics requests
  if (req.method !== 'GET' || !url.protocol.startsWith('http')) return;

  // Specific handler for pins.json (Network First, Cache Fallback)
  if (url.pathname.endsWith('/pins.json')) {
    event.respondWith(
      fetch(req)
        .then(res => {
          if (res && res.ok) {
            const copy = res.clone();
            caches.open(CACHE_NAME).then(cache => cache.put('./pins.json', copy));
          }
          return res;
        })
        .catch(async () => {
          const cached = await caches.match('./pins.json') || await caches.match('pins.json');
          if (cached) return cached;
          return new Response('[]', { headers: { 'Content-Type': 'application/json' } });
        })
    );
    return;
  }

  // Specific handler for page navigation (Network First with Cache Fallback)
  if (req.mode === 'navigate') {
    event.respondWith(
      fetch(req)
        .then(res => {
          if (res && res.ok) {
            const copy = res.clone();
            caches.open(CACHE_NAME).then(cache => cache.put('./index.html', copy));
          }
          return res;
        })
        .catch(async () => {
          return (await caches.match('./index.html')) || (await caches.match('./'));
        })
    );
    return;
  }

  // Cache First for static assets (icons, manifest, etc.)
  event.respondWith(
    caches.match(req).then(cached => {
      if (cached) return cached;
      return fetch(req).then(res => {
        if (!res || res.status !== 200 || res.type !== 'basic') return res;
        const copy = res.clone();
        caches.open(CACHE_NAME).then(cache => cache.put(req, copy));
        return res;
      });
    })
  );
});
