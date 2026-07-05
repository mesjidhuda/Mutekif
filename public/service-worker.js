const CACHE_NAME = "app-cache-v1";

const FILES = [
    "/",
    "/index.html",
    "/generator.html",
    "/validator.html",
    "/css/premium.css"
];

self.addEventListener("install", event => {
    event.waitUntil(caches.open(CACHE_NAME).then(cache => cache.addAll(FILES)));
});

self.addEventListener("fetch", event => {
    event.respondWith(
        caches.match(event.request).then(res => res || fetch(event.request))
    );
});
