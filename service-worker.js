// WordWise PH Service Worker v1
const CACHE = 'wordwiseph-v1';
const STATIC = [
  '/',
  '/index.html',
  '/app.html',
  '/questions.json',
  '/favicon.png',
  '/og-image.png',
  '/icons/icon-192.png',
  '/icons/icon-512.png',
];

// Install: cache static assets
self.addEventListener('install', e => {
  self.skipWaiting();
  e.waitUntil(
    caches.open(CACHE).then(c => c.addAll(STATIC))
  );
});

// Activate: clean old caches
self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)))
    )
  );
  self.clients.claim();
});

// Fetch: network first, fallback to cache
self.addEventListener('fetch', e => {
  // Skip non-GET and chrome-extension requests
  if (e.request.method !== 'GET') return;
  if (!e.request.url.startsWith('http')) return;
  // Skip Supabase API calls — always need fresh data
  if (e.request.url.includes('supabase.co')) return;

  e.respondWith(
    fetch(e.request)
      .then(res => {
        // Cache successful responses for static assets
        if (res.ok && (
          e.request.url.endsWith('.html') ||
          e.request.url.endsWith('.json') ||
          e.request.url.endsWith('.png') ||
          e.request.url.endsWith('.jpg') ||
          e.request.url.endsWith('.css') ||
          e.request.url.endsWith('.js')
        )) {
          const clone = res.clone();
          caches.open(CACHE).then(c => c.put(e.request, clone));
        }
        return res;
      })
      .catch(() => caches.match(e.request))
  );
});
