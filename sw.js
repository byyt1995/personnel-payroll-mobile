const CACHE_VERSION = 'payroll-mobile-shell-v1.5.0';
const SHELL_ASSETS = [
  './',
  './index.html',
  './styles.css',
  './cloud-snapshot.js',
  './vendor/msal-browser.min.js',
  './microsoft-auth.js',
  './onedrive-snapshot.js',
  './app.js',
  './manifest.webmanifest',
  './icons/icon-192.png',
  './icons/icon-512.png',
];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_VERSION)
      .then((cache) => cache.addAll(SHELL_ASSETS))
      .then(() => self.skipWaiting()),
  );
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(keys
        .filter((key) => key.startsWith('payroll-mobile-shell-') && key !== CACHE_VERSION)
        .map((key) => caches.delete(key))))
      .then(() => self.clients.claim()),
  );
});

self.addEventListener('fetch', (event) => {
  const request = event.request;
  const url = new URL(request.url);

  if (url.origin !== self.location.origin) return;

  // API responses can contain personal data. They are always fetched from the
  // server and are never read from or written to Cache Storage.
  if (url.pathname.startsWith('/api/')) {
    event.respondWith(fetch(request, { cache: 'no-store' }));
    return;
  }

  if (request.method !== 'GET') return;

  const shellUrls = new Set(SHELL_ASSETS.map((asset) => new URL(asset, self.registration.scope).href));
  const isNavigation = request.mode === 'navigate';
  const isShellAsset = shellUrls.has(url.href);
  if (!isNavigation && !isShellAsset) return;

  event.respondWith(
    caches.match(isNavigation ? new URL('./index.html', self.registration.scope).href : request)
      .then((cached) => cached || fetch(request)),
  );
});
