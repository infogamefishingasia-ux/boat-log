// Aqua Creed Boat Log - Service Worker
// Caches the app shell so the form loads with zero connectivity.
// Bump CACHE_NAME whenever index.html/manifest/icons change to force an update.
const CACHE_NAME = "boatlog-shell-v9";
const SHELL_FILES = [
  "./",
  "./index.html",
  "./admin.html",
  "./manifest.json",
  "./icon-192.png",
  "./icon-512.png",
  "./logo.png"
];

self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(SHELL_FILES))
  );
  self.skipWaiting();
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(
        keys.filter((k) => k !== CACHE_NAME).map((k) => caches.delete(k))
      )
    )
  );
  self.clients.claim();
});

// Cache-first for app shell files; network passthrough for everything else
// (sync requests to Apps Script always go to the network, never cached).
self.addEventListener("fetch", (event) => {
  const url = new URL(event.request.url);

  // Only handle GET requests for same-origin shell files
  if (event.request.method !== "GET" || url.origin !== self.location.origin) {
    return;
  }

  event.respondWith(
    caches.match(event.request).then((cached) => {
      if (cached) return cached;
      return fetch(event.request).then((response) => {
        // Cache new same-origin shell assets as they're fetched
        const copy = response.clone();
        caches.open(CACHE_NAME).then((cache) => cache.put(event.request, copy));
        return response;
      }).catch(() => cached);
    })
  );
});
