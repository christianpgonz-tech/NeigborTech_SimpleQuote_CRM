/**
 * Service worker for SimpleQuote — caches the app SHELL only (the
 * static HTML/CSS/JS files) so the app can open even with no
 * internet connection. It never caches or intercepts anything else.
 *
 * CRITICAL SAFETY RULE: any request that isn't for one of our own
 * static files — most importantly, every JSONP call this app makes
 * to the Google Apps Script backend — must always go straight to the
 * network, untouched. Quotes/invoices are live business data; they
 * must never be served from a stale cache. See the origin check
 * below, which is what guarantees this.
 */

const CACHE_NAME = 'neighbortech-simplequote-shell-v1';
const SHELL_FILES = [
  './app.html',
  './manifest.json',
  './icons/icon-192.png',
  './icons/icon-512.png'
];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(SHELL_FILES))
  );
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((names) =>
      Promise.all(names.filter((n) => n !== CACHE_NAME).map((n) => caches.delete(n)))
    )
  );
  self.clients.claim();
});

self.addEventListener('fetch', (event) => {
  const url = new URL(event.request.url);

  // Any cross-origin request — the Apps Script backend, Google Fonts,
  // anything — is left completely alone. Not calling respondWith()
  // means the browser handles it exactly as if this service worker
  // didn't exist at all.
  if (url.origin !== self.location.origin) {
    return;
  }

  // Our own static files: serve from cache if we have it (so the
  // shell opens instantly, even offline), otherwise fetch normally.
  event.respondWith(
    fetch(event.request).catch(() => caches.match(event.request))
  );
});
