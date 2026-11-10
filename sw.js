// Guarda o formulário no aparelho para abrir sem internet e envia a fila com o app fechado.
const CACHE = "apont-v12";
const ARQS = ["./", "./index.html", "./manifest.webmanifest", "./icon-192.png", "./icon-512.png", "./rdo-assets.js"];

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
});

/* ---------- envio em segundo plano (Android/Chrome): roda quando a internet volta, mesmo com o app fechado ---------- */
function abrirDb() {
  return new Promise((ok, err) => {
    const r = indexedDB.open("apontTerceiros", 2);
    r.onupgradeneeded = () => { const db = r.result;
      ["fila", "entregues", "anexos"].forEach(n => { if (!db.objectStoreNames.contains(n)) db.createObjectStore(n, { keyPath: "id" }); }); };
    r.onsuccess = () => ok(r.result); r.onerror = () => err(r.error);
  });
}
const tx = (db, store, mode, fn) => new Promise((ok, err) => { const t = db.transaction(store, mode); const q = fn(t.objectStore(store)); t.oncomplete = () => ok(q && q.result); t.onerror = () => err(t.error); });

async function enviarFila() {
  const db = await abrirDb();
  const fila = await tx(db, "fila", "readonly", s => s.getAll());
  let entregou = 0;
  for (const item of fila) {
    const r = await fetch(item.url, { method: "POST", headers: { "Content-Type": "text/plain;charset=utf-8" }, body: JSON.stringify(item.payload) }); // sem internet: lança erro e o navegador tenta de novo depois
    const j = await r.json().catch(() => ({ ok: false }));
    if (j.ok) {
      await tx(db, "fila", "readwrite", s => s.delete(item.id));
      await tx(db, "entregues", "readwrite", s => s.put({ id: item.id, recebido: new Date().toISOString() }));
      entregou++;
    }
  }
  if (entregou) {
    const abertos = await self.clients.matchAll({ includeUncontrolled: true });
    abertos.forEach(c => c.postMessage({ tipo: "fila" }));
  }
}
self.addEventListener("sync", e => { if (e.tag === "enviar-fila") e.waitUntil(enviarFila()); });
