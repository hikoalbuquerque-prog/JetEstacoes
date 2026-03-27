/**
 * service-worker.js — App Estações Campo PWA
 *
 * Estratégia: Cache-first para assets estáticos (shell),
 * Network-first para chamadas ao GAS backend.
 *
 * Cache names são versionados — bumpar SW_VERSION para forçar
 * atualização em todos os clientes.
 */

var SW_VERSION   = 'campo-v3';
var CACHE_STATIC = SW_VERSION + '-static';
var CACHE_DATA   = SW_VERSION + '-data';

// Assets do app shell (servidos pelo GitHub Pages)
var STATIC_ASSETS = [
  './',
  './index.html',
  './manifest.json',
  './icon-192.png',
  './icon-512.png',
];

// ── Install: pre-cache o shell ────────────────────────────────
self.addEventListener('install', function(event) {
  event.waitUntil(
    caches.open(CACHE_STATIC).then(function(cache) {
      return cache.addAll(STATIC_ASSETS);
    }).then(function() {
      return self.skipWaiting();
    })
  );
});

// ── Activate: limpar caches antigos ──────────────────────────
self.addEventListener('activate', function(event) {
  event.waitUntil(
    caches.keys().then(function(keys) {
      return Promise.all(
        keys.filter(function(key) {
          return key !== CACHE_STATIC && key !== CACHE_DATA;
        }).map(function(key) {
          return caches.delete(key);
        })
      );
    }).then(function() {
      return self.clients.claim();
    })
  );
});

// ── Fetch: estratégia por tipo de request ────────────────────
self.addEventListener('fetch', function(event) {
  var url = new URL(event.request.url);

  // Chamadas ao GAS backend: network-first, sem cache
  if (url.hostname === 'script.google.com' ||
      url.hostname.endsWith('.googleusercontent.com')) {
    event.respondWith(
      fetch(event.request).catch(function() {
        return new Response(
          JSON.stringify({ ok: false, error: 'Sem conexao. Verifique a internet.' }),
          { headers: { 'Content-Type': 'application/json' } }
        );
      })
    );
    return;
  }

  // Google Maps API: network-first com fallback
  if (url.hostname === 'maps.googleapis.com' ||
      url.hostname === 'maps.gstatic.com') {
    event.respondWith(fetch(event.request));
    return;
  }

  // App shell (GitHub Pages): cache-first
  event.respondWith(
    caches.match(event.request).then(function(cached) {
      if (cached) return cached;
      return fetch(event.request).then(function(response) {
        // Cache apenas respostas válidas de assets estáticos
        if (response && response.status === 200 &&
            response.type === 'basic') {
          var clone = response.clone();
          caches.open(CACHE_STATIC).then(function(cache) {
            cache.put(event.request, clone);
          });
        }
        return response;
      });
    })
  );
});
