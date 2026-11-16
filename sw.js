/* 이종대왕 달력 — 서비스 워커
   GitHub Pages 하위 경로(/leejong-calendar/)에서도 동작하도록 모든 경로를 상대 경로로 둡니다.
   전략
     · 첫 설치: 앱 껍데기(HTML·매니페스트·아이콘)를 미리 받아 캐시 → 오프라인에서도 열림
     · 화면 이동(navigation): Network-first (새 버전이 있으면 바로 반영, 오프라인이면 캐시)
     · 그 밖의 정적 파일: Cache-first + 뒤에서 갱신(stale-while-revalidate)
   기록 데이터는 여기에 오지 않습니다. 기록은 브라우저 localStorage에만 있습니다. */
const CACHE = 'ljc-shell-v3';
const ASSETS = [
  './',
  './index.html',
  './manifest.json',
  './icon-192.png',
  './icon-512.png'
];

self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE)
      .then(cache => cache.addAll(ASSETS).catch(() => {}))
      .then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys()
      .then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', event => {
  const req = event.request;
  if (req.method !== 'GET') return;

  const url = new URL(req.url);
  if (url.origin !== self.location.origin) return;

  if (req.mode === 'navigate') {
    event.respondWith(
      fetch(req)
        .then(res => {
          if (res && res.status === 200) {
            const copy = res.clone();
            caches.open(CACHE).then(c => c.put(req, copy)).catch(() => {});
          }
          return res;
        })
        .catch(() => caches.match(req, { ignoreSearch: true })
          .then(hit => hit || caches.match('./index.html'))
          .then(page => page || Response.error()))
    );
    return;
  }

  event.respondWith(
    caches.match(req, { ignoreSearch: true }).then(hit => {
      const network = fetch(req).then(res => {
        if (res && res.status === 200 && res.type === 'basic') {
          const copy = res.clone();
          caches.open(CACHE).then(c => c.put(req, copy)).catch(() => {});
        }
        return res;
      }).catch(() => hit || Response.error());
      return hit || network;
    })
  );
});
