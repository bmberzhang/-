// 水闸设计系统 - Service Worker（多用户登录版）
const CACHE = 'sluice-app-v2';
const PRECACHE = [
  '/static/manifest.json',
  '/static/icons/icon-192.png',
  '/static/icons/icon-512.png',
  '/static/icons/icon.svg',
  '/static/sw.js',
];

// 安装：预缓存（只缓存静态资源，不缓存需要登录的页面）
self.addEventListener('install', (e) => {
  e.waitUntil(caches.open(CACHE).then((c) => c.addAll(PRECACHE).catch(() => {})));
  self.skipWaiting();
});

// 激活：清理旧缓存
self.addEventListener('activate', (e) => {
  e.waitUntil(caches.keys().then((keys) =>
    Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k)))
  ));
  self.clients.claim();
});

// 拦截请求
self.addEventListener('fetch', (e) => {
  const url = new URL(e.request.url);
  // API 永远走网络（不缓存）
  if (url.pathname.startsWith('/drawing/generate') ||
      url.pathname.startsWith('/generate') ||
      url.pathname.startsWith('/api/')) {
    return;
  }
  // 页面导航请求（/、/drawing、/login 等）：网络优先，离线时回退缓存
  if (e.request.mode === 'navigate') {
    e.respondWith(
      fetch(e.request).then((resp) => {
        // 只缓存 200 的页面（避免缓存登录页/错误页）
        if (resp && resp.status === 200) {
          const copy = resp.clone();
          caches.open(CACHE).then((c) => c.put(e.request, copy));
        }
        return resp;
      }).catch(() => caches.match(e.request).then((cached) => cached || caches.match('/login')))
    );
    return;
  }
  // 其他静态资源：缓存优先
  e.respondWith(
    caches.match(e.request).then((cached) => {
      if (cached) return cached;
      return fetch(e.request).then((resp) => {
        if (e.request.method === 'GET' && resp.status === 200) {
          const copy = resp.clone();
          caches.open(CACHE).then((c) => c.put(e.request, copy));
        }
        return resp;
      }).catch(() => cached);
    })
  );
});