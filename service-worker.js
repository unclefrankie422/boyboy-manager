// Frankie 工作台 Service Worker - 离线缓存
const CACHE_NAME = 'frankie-workbuddy-v1';
const FILES_TO_CACHE = [
  './',
  './index.html',
  './catalog.html',
  './library.html',
  './library-en.html',
  './business-plan.html',
  './wecom-guide.html',
  './frankie-card-en.html',
  './english-course-market-research.html',
  './english-day1-view.html',
  './english-day2-view.html',
  './english-day3-view.html',
  './why-four-seed-oil-view.html',
  './why-four-seed-oil-dl.html',
  './personal-ip-strategy.html',
  './why-four-seed-oil.pptx',
  './manifest.json'
];

// 安装时缓存所有文件
self.addEventListener('install', (event) => {
  console.log('[SW] Installing and caching files');
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then((cache) => {
        return cache.addAll(FILES_TO_CACHE);
      })
      .then(() => self.skipWaiting())
  );
});

// 激活时清理旧缓存
self.addEventListener('activate', (event) => {
  console.log('[SW] Activating and cleaning old caches');
  event.waitUntil(
    caches.keys().then((cacheNames) => {
      return Promise.all(
        cacheNames.map((cacheName) => {
          if (cacheName !== CACHE_NAME) {
            console.log('[SW] Deleting old cache:', cacheName);
            return caches.delete(cacheName);
          }
        })
      );
    }).then(() => self.clients.claim())
  );
});

// 拦截请求 - 优先使用缓存
self.addEventListener('fetch', (event) => {
  if (event.request.method !== 'GET') return;

  event.respondWith(
    caches.match(event.request)
      .then((cachedResponse) => {
        if (cachedResponse) {
          return cachedResponse;
        }
        return fetch(event.request).then((networkResponse) => {
          // 缓存新的请求
          if (networkResponse && networkResponse.status === 200 && networkResponse.type === 'basic') {
            const responseToCache = networkResponse.clone();
            caches.open(CACHE_NAME).then((cache) => {
              cache.put(event.request, responseToCache);
            });
          }
          return networkResponse;
        });
      })
      .catch(() => {
        // 离线兜底
        if (event.request.destination === 'document') {
          return caches.match('./index.html');
        }
      })
  );
});