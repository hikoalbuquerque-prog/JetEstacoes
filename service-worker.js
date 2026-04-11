/**
 * service-worker.js — App Estações Campo PWA
 * SW_VERSION: campo-v5
 *
 * GAS backend usa redirect 302 -> script.googleusercontent.com
 * O SW NAO intercepta chamadas ao GAS -- deixa o browser lidar
 * diretamente para evitar problemas com redirects opacos.
 */

var SW_VERSION   = 'campo-v5';
var CACHE_STATIC = SW_VERSION + '-static';

// Assets do app shell (servidos pelo GitHub Pages)
var STATIC_ASSETS = [
  './',
  './campo.html',
  './manifest.json',
];

// ── Install ────────────────────────────────────────────────────
self.addEventListener('install', function(event) {
  event.waitUntil(
    caches.open(CACHE_STATIC).then(function(cache) {
      // addAll com ignoreSearch para evitar erro em assets opcionais
      return Promise.allSettled(
        STATIC_ASSETS.map(function(url) { return cache.add(url); })
      );
    }).then(function() {
      return self.skipWaiting();
    })
  );
});

// ── Activate: limpar caches antigos ───────────────────────────
self.addEventListener('activate', function(event) {
  event.waitUntil(
    caches.keys().then(function(keys) {
      return Promise.all(
        keys.filter(function(key) {
          return key !== CACHE_STATIC;
        }).map(function(key) {
          return caches.delete(key);
        })
      );
    }).then(function() {
      return self.clients.claim();
    })
  );
});

// ── Fetch ──────────────────────────────────────────────────────
self.addEventListener('fetch', function(event) {
  var url = event.request.url;

  // NAO interceptar chamadas ao GAS nem ao Google APIs
  // Deixar o browser lidar diretamente -- evita problemas com
  // redirects 302 opacos do script.google.com
  if (url.indexOf('script.google.com') >= 0 ||
      url.indexOf('script.googleusercontent.com') >= 0 ||
      url.indexOf('googleapis.com') >= 0 ||
      url.indexOf('gstatic.com') >= 0 ||
      url.indexOf('overpass') >= 0 ||
      url.indexOf('maps.google') >= 0) {
    // Nao chamar event.respondWith -- browser trata normalmente
    return;
  }

  // App shell: cache-first para assets estaticos do GitHub Pages
  if (event.request.method !== 'GET') return;

  event.respondWith(
    caches.match(event.request).then(function(cached) {
      if (cached) return cached;
      return fetch(event.request).then(function(response) {
        if (response && response.status === 200 && response.type === 'basic') {
          var clone = response.clone();
          caches.open(CACHE_STATIC).then(function(cache) {
            cache.put(event.request, clone);
          });
        }
        return response;
      }).catch(function() {
        // Offline e nao tem cache -- retorna 503
        return new Response('Offline', { status: 503 });
      });
    })
  );
});

// Aceitar comando SKIP_WAITING para atualizar imediatamente
self.addEventListener('message', function(event) {
  if (event.data && event.data.type === 'SKIP_WAITING') {
    self.skipWaiting();
  }
});
