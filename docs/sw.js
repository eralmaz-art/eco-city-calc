/*
 * Service worker офлайн-калькулятора «Эко Сити Айни».
 * После первой загрузки все файлы лежат в кэше — приложение работает без сети.
 *
 * ПРИ ЛЮБОМ ОБНОВЛЕНИИ ФАЙЛОВ (apartments.js, index.html…) поднимите номер
 * CACHE_VERSION на единицу — телефоны заметят новую версию и обновят кэш.
 */
var CACHE_VERSION = 2;
var CACHE_NAME = 'eco-city-aini-calc-v' + CACHE_VERSION;

var ASSETS = [
  './',
  'index.html',
  'calc.js',
  'apartments.js',
  'manifest.json',
  'icon-192.png',
  'icon-512.png',
  'icon-180.png'
];

self.addEventListener('install', function (e) {
  e.waitUntil(
    caches.open(CACHE_NAME)
      .then(function (cache) { return cache.addAll(ASSETS); })
      .then(function () { return self.skipWaiting(); })
  );
});

self.addEventListener('activate', function (e) {
  e.waitUntil(
    caches.keys().then(function (keys) {
      return Promise.all(keys.map(function (k) {
        if (k !== CACHE_NAME) return caches.delete(k);
      }));
    }).then(function () { return self.clients.claim(); })
  );
});

/* Кэш — первым (мгновенно и офлайн), сеть — фоновым обновлением кэша. */
self.addEventListener('fetch', function (e) {
  if (e.request.method !== 'GET') return;
  e.respondWith(
    caches.match(e.request, { ignoreSearch: true }).then(function (cached) {
      var fetched = fetch(e.request).then(function (resp) {
        if (resp && resp.ok) {
          var copy = resp.clone();
          caches.open(CACHE_NAME).then(function (cache) { cache.put(e.request, copy); });
        }
        return resp;
      }).catch(function () { return cached; });
      return cached || fetched;
    })
  );
});
