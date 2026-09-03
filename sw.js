/* Service worker SNAF Cotes — réseau d'abord, autonettoyant.
   Le cache ne sert que de secours hors-ligne. Les /api/* ne sont jamais interceptés. */
const CACHE = 'snaf-cotes-v5';
const STATIC = ['./', './index.html'];

self.addEventListener('install', function (e) {
  e.waitUntil(
    Promise.all([
      caches.open(CACHE).then(function (c) { return c.addAll(STATIC); }),
      self.skipWaiting()
    ])
  );
});

self.addEventListener('activate', function (e) {
  e.waitUntil(
    caches.keys().then(function (keys) {
      return Promise.all(
        keys.filter(function (k) { return k !== CACHE; }).map(function (k) { return caches.delete(k); })
      );
    }).then(function () { return self.clients.claim(); })
  );
});

self.addEventListener('message', function (e) {
  if (e.data && e.data.type === 'SKIP_WAITING') self.skipWaiting();
});

self.addEventListener('fetch', function (e) {
  var url = new URL(e.request.url);
  if (url.origin !== self.location.origin) return;
  if (url.pathname.indexOf('/api/') === 0) return; // jamais l'API
  e.respondWith(
    fetch(e.request).then(function (resp) {
      if (resp && resp.status === 200) {
        var copy = resp.clone();
        caches.open(CACHE).then(function (c) { c.put(e.request, copy); });
      }
      return resp;
    }).catch(function () {
      return caches.match(e.request).then(function (hit) { return hit || Response.error(); });
    })
  );
});
