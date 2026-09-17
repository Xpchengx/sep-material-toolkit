/* Service Worker —— 工作台 PWA
   策略：
     · 应用外壳（页面/图标/清单）走 cache-first，保证离线可打开
     · 数据文件（workbench-data.js）不缓存，永远取新的
   ⚠️ 改动页面后请把 VERSION 加一，否则用户拿到的是旧缓存。
*/
const VERSION = 'v1';
const CACHE = 'material-workbench-' + VERSION;

const SHELL = [
  './',
  './index.html',
  './manifest.webmanifest',
  './icons/icon.svg',
  './icons/icon-192.png',
  './icons/icon-512.png',
  './icons/apple-touch-icon.png',
];

self.addEventListener('install', (e) => {
  e.waitUntil(
    caches.open(CACHE)
      .then((c) => c.addAll(SHELL).catch(() => undefined))
      .then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', (e) => {
  e.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(
        keys.filter((k) => k !== CACHE).map((k) => caches.delete(k))
      ))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', (e) => {
  const req = e.request;
  if (req.method !== 'GET') return;

  const url = new URL(req.url);
  // 只处理同源
  if (url.origin !== self.location.origin) return;
  // 数据文件不缓存
  if (url.pathname.endsWith('workbench-data.js')) return;

  // 页面导航：网络优先，断网回落缓存
  if (req.mode === 'navigate') {
    e.respondWith(
      fetch(req)
        .then((res) => {
          const copy = res.clone();
          caches.open(CACHE).then((c) => c.put('./index.html', copy));
          return res;
        })
        .catch(() => caches.match('./index.html').then((r) => r || caches.match('./')))
    );
    return;
  }

  // 其它静态资源：缓存优先
  e.respondWith(
    caches.match(req).then((hit) => {
      if (hit) return hit;
      return fetch(req).then((res) => {
        if (res && res.status === 200 && res.type === 'basic') {
          const copy = res.clone();
          caches.open(CACHE).then((c) => c.put(req, copy));
        }
        return res;
      }).catch(() => hit);
    })
  );
});
