/* Service worker de « Rapports des piliers MVE » : l'application s'ouvre et se remplit sans Internet.
   Réseau d'abord pour les fichiers de l'application (nouvelle version prise dès qu'il y a Internet),
   cache seulement hors ligne. Les envois vers Google Sheets ne passent pas par ici. */
var VERSION = 'piliers-mve-v30'; // v2 : en-tête officiel SGI MVE 17
var SHELL = ['./', 'index.html', 'app-piliers.js', 'piliers-def.js', 'agregats.js', 'qualite-piliers.js', 'exports.js', 'diapo-merci.jpg', 'jszip.min.js', 'pptxgen.bundle.js', 'xlsx.mini.min.js',
  'manifest.webmanifest', 'icon-192.png', 'icon-512.png', 'maskable-512.png', 'apple-touch-icon.png', 'embleme.png', 'entete-officiel.png', 'embleme-petit.png'];
self.addEventListener('install', function (e) {
  e.waitUntil(caches.open(VERSION).then(function (c) {
    // Fichier par fichier : un fichier manquant n'empêche pas le fonctionnement hors ligne
    return Promise.all(SHELL.map(function (u) {
      return fetch(new Request(u, { cache: 'reload' })).then(function (r) { if (r.ok) return c.put(u, r); }).catch(function () { /* ignoré */ });
    }));
  }).then(function () { return self.skipWaiting(); }));
});
self.addEventListener('activate', function (e) {
  e.waitUntil(caches.keys().then(function (keys) {
    return Promise.all(keys.filter(function (k) { return k.indexOf('piliers-mve-') === 0 && k !== VERSION; }).map(function (k) { return caches.delete(k); }));
  }).then(function () { return self.clients.claim(); }));
});
self.addEventListener('fetch', function (e) {
  var url = new URL(e.request.url);
  if (e.request.method !== 'GET' || url.origin !== self.location.origin) return;
  e.respondWith(caches.open(VERSION).then(function (c) {
    return fetch(e.request, { cache: 'no-cache' }).then(function (res) {
      if (res && res.ok) c.put(e.request, res.clone());
      return res;
    }).catch(function () {
      return c.match(e.request, { ignoreSearch: true }).then(function (hit) { return hit || c.match('index.html'); });
    });
  }));
});
