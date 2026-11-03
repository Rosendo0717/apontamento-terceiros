// Guarda o formulário no aparelho para abrir sem internet.
// Página: tenta a internet (até 4 s) e, se não der, usa a cópia guardada.
const CACHE = "apont-v1";
const ARQS = ["./", "./index.html", "./manifest.webmanifest", "./icon-192.png", "./icon-512.png"];

self.addEventListener("install", e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(ARQS)));
  self.skipWaiting();
});
self.addEventListener("activate", e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});
const comTempo = (p, ms) => Promise.race([p, new Promise((_, rej) => setTimeout(() => rej(new Error("tempo")), ms))]);

self.addEventListener("fetch", e => {
  const req = e.request; if (req.method !== "GET") return;
  const u = new URL(req.url);
  if (u.origin === location.origin) {
    e.respondWith(
      comTempo(fetch(req), 4000)
        .then(r => { if (r.ok) { const cp = r.clone(); caches.open(CACHE).then(c => c.put(req, cp)); } return r; })
        .catch(() => caches.match(req, { ignoreSearch: true }).then(r => r || caches.match("./index.html")))
    );
  } else if (u.hostname === "fonts.googleapis.com" || u.hostname === "fonts.gstatic.com") {
    e.respondWith(caches.match(req).then(r => r || fetch(req).then(res => { const cp = res.clone(); caches.open(CACHE).then(c => c.put(req, cp)); return res; })));
  }
  // o envio para o Google (script.google.com) não passa por aqui: quem cuida é a fila do formulário
});
