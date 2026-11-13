// Keeps the app's own files on the phone so it opens without internet. Data lives in IndexedDB, not here.
const V = "ssi-d673147a18";
const FILES = ["./", "index.html", "config.js", "onedrive.js", "msal-browser.min.js", "sql-asm.js", "xlsx.mini.min.js", "jspdf.umd.min.js", "jspdf.plugin.autotable.min.js", "appdb.txt", "manifest.webmanifest", "icon-192.png", "icon-512.png", "icon-maskable-512.png", "apple-180.png"];
self.addEventListener("install", (e) => { e.waitUntil(caches.open(V).then((c) => c.addAll(FILES))); });
self.addEventListener("message", (e) => { if (e.data === "skip") self.skipWaiting(); });
self.addEventListener("activate", (e) => { e.waitUntil(caches.keys().then((ks) => Promise.all(ks.filter((k) => k !== V).map((k) => caches.delete(k)))).then(() => self.clients.claim())); });
self.addEventListener("fetch", (e) => {
  const u = new URL(e.request.url);
  if (e.request.method !== "GET" || u.origin !== location.origin) return;
  e.respondWith(caches.open(V).then((c) => c.match(e.request, { ignoreSearch: true }).then((hit) => hit || fetch(e.request))));
});
