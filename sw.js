// Cache-first service worker for the app shell so the PWA opens offline once installed.
// IMPORTANT: only same-origin GETs are cached — Microsoft Graph / login traffic is never touched.
const CACHE = "tmg-tt-v3";
const ASSETS = [
  "./", "index.html", "style.css",
  "app.js", "config.js", "models.js", "db.js", "auth.js", "graph.js", "sync.js",
  "vendor/msal-browser.min.js",
  "manifest.webmanifest", "icon-180.png", "icon-512.png",
];

self.addEventListener("install", (e) => {
  e.waitUntil(caches.open(CACHE).then((c) => c.addAll(ASSETS)).then(() => self.skipWaiting()));
});

self.addEventListener("activate", (e) => {
  e.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener("fetch", (e) => {
  const req = e.request;
  if (req.method !== "GET") return;
  if (new URL(req.url).origin !== location.origin) return; // leave Graph / MSAL alone
  e.respondWith(
    caches.match(req).then((hit) =>
      hit || fetch(req).then((res) => {
        const copy = res.clone();
        caches.open(CACHE).then((c) => c.put(req, copy)).catch(() => {});
        return res;
      }).catch(() => caches.match("index.html"))
    )
  );
});
