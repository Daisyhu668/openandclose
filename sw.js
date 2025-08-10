const CACHE = 'openandclose-v1';
const ASSETS = [
  '/openandclose/',
  '/openandclose/index.html',
  '/openandclose/closegen.html',
  '/openandclose/pizzip.min.js',
  '/openandclose/close_template.docx' // 若需要预置模板就加上；你也可以只用本地缓存模板
];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(ASSETS)));
  self.skipWaiting();
});
self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k))))
  );
  self.clients.claim();
});
self.addEventListener('fetch', e => {
  const req = e.request;
  e.respondWith(
    caches.match(req).then(hit => {
      if (hit) return hit;
      return fetch(req).then(resp => {
        if (req.method === 'GET' && resp.ok && new URL(req.url).origin === location.origin) {
          const copy = resp.clone();
          caches.open(CACHE).then(c => c.put(req, copy));
        }
        return resp;
      }).catch(() => caches.match('/openandclose/index.html'));
    })
  );
});
