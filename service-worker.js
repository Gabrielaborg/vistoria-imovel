// Service Worker do app de Vistoria (G&G Engenharia)
// Objetivo: deixar o app instalável (PWA) e permitir que a TELA do app
// (não os dados enviados) continue abrindo mesmo com internet ruim/instável.
// Envios de laudo (POST) NUNCA passam por aqui — vão sempre direto pra rede,
// e se falharem, quem cuida disso é a fila de reenvio automático do próprio app.

const CACHE_NAME = 'vistoria-ggeng-v1';
const APP_SHELL = ['/', '/logo.png', '/manifest.json', '/icon-192.png', '/icon-512.png'];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then((cache) => cache.addAll(APP_SHELL))
      .catch((e) => console.error('Falha ao preparar cache do app shell:', e))
  );
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((chaves) =>
      Promise.all(chaves.filter((k) => k !== CACHE_NAME).map((k) => caches.delete(k)))
    )
  );
  self.clients.claim();
});

self.addEventListener('fetch', (event) => {
  const { request } = event;

  // Só cuida de pedidos GET (a tela do app, imagens, etc).
  // Envio de laudo, geração de conclusão por IA, download de histórico
  // e qualquer outro POST/GET dinâmico passam direto pra rede, sem cache.
  if (request.method !== 'GET') return;

  const url = new URL(request.url);
  const rotasDinamicas = ['/historico', '/download', '/download-pdf', '/gerar-laudo', '/gerar-conclusao'];
  if (rotasDinamicas.some((r) => url.pathname.startsWith(r))) return;

  event.respondWith(
    fetch(request)
      .then((res) => {
        const resClone = res.clone();
        caches.open(CACHE_NAME).then((cache) => cache.put(request, resClone));
        return res;
      })
      .catch(() => caches.match(request))
  );
});
