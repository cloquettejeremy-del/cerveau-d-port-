const CACHE_NAME = "decharge-mentale-v3";
const APP_SHELL = [
  "./",
  "./index.html",
  "./style.css",
  "./app.js",
  "./manifest.json",
  "./icons/icon-192.png",
  "./icons/icon-512.png",
];

self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(APP_SHELL)).catch(() => {})
  );
  self.skipWaiting();
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((k) => k !== CACHE_NAME).map((k) => caches.delete(k)))
    )
  );
  self.clients.claim();
});

// Réseau d'abord pour le shell de l'app (HTML/JS/CSS/manifest) : ça garantit que les
// mises à jour du site sont prises en compte tout de suite, sans devoir vider le cache
// manuellement. Si le réseau est indisponible (hors-ligne), on retombe sur le cache.
// Les appels vers Microsoft (Graph/login/MSAL) ne passent jamais par le cache.
self.addEventListener("fetch", (event) => {
  const url = event.request.url;
  if (url.includes("graph.microsoft.com") || url.includes("login.microsoftonline.com") || url.includes("msauth.net")) {
    return; // laisser passer normalement, pas de cache
  }
  event.respondWith(
    fetch(event.request)
      .then((response) => {
        const copy = response.clone();
        caches.open(CACHE_NAME).then((cache) => cache.put(event.request, copy)).catch(() => {});
        return response;
      })
      .catch(() => caches.match(event.request))
  );
});
