const CACHE_PREFIX = 'so-thu-hoc-sinh-hieugiang-so2-shell-';
const CACHE_NAME = CACHE_PREFIX + 'v15-bidv-export';
const SHELL = ['./', './index.html', './styles.css', './app.js', './insurance.js', './manifest.json', './icon.svg', './vendor/qrcode.min.js', './vendor/jszip.min.js'];
self.addEventListener('install', event => {
  event.waitUntil(caches.open(CACHE_NAME).then(cache => cache.addAll(SHELL)));
  self.skipWaiting();
});
self.addEventListener('activate', event => {
  event.waitUntil(caches.keys().then(keys => Promise.all(keys.filter(key => key.startsWith(CACHE_PREFIX) && key !== CACHE_NAME).map(key => caches.delete(key)))));
  self.clients.claim();
});
self.addEventListener('fetch', event => {
  const request = event.request;
  if (request.method !== 'GET' || new URL(request.url).origin !== self.location.origin) return;
  const url = new URL(request.url);
  const isCritical = request.mode === 'navigate' || /\/(index\.html|app\.js|styles\.css|sw\.js)$/.test(url.pathname);
  if (isCritical) {
    event.respondWith(fetch(request).then(response => {
      if (response.ok) { const copy=response.clone(); caches.open(CACHE_NAME).then(cache=>cache.put(request,copy)); }
      return response;
    }).catch(() => caches.match(request).then(cached => cached || caches.match('./index.html'))));
    return;
  }
  event.respondWith(caches.match(request).then(cached => cached || fetch(request).then(response => {
    if (response.ok) { const copy=response.clone(); caches.open(CACHE_NAME).then(cache=>cache.put(request,copy)); }
    return response;
  }).catch(() => caches.match('./index.html'))));
});
