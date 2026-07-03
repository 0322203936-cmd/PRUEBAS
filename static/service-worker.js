/**
 * service-worker.js — Service Worker con soporte offline completo
 *
 * Estrategia: Cache-First para assets estáticos, Network-First para páginas HTML
 * - Pre-cachea las páginas principales al instalar
 * - Si hay red → sirve desde red y actualiza caché
 * - Si no hay red → sirve desde caché sin error
 */

const CACHE_NAME    = 'inventario-v7';
const STATIC_ASSETS = [
    '/',
    '/inventario',
    '/gastos',
    '/reporte',
    '/static/offline.js',
    '/static/manifest.json',
    '/static/icon-192.png',
    '/static/icon-512.png',
    'https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css',
    'https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap'
];

// ── Install: pre-cachear assets ────────────────────────────────────────────────

self.addEventListener('install', (event) => {
    console.log('[SW] Instalando v6…');
    event.waitUntil(
        caches.open(CACHE_NAME).then(async (cache) => {
            for (const url of STATIC_ASSETS) {
                try {
                    await cache.add(url);
                    console.log('[SW] Cacheado:', url);
                } catch (err) {
                    console.warn('[SW] No se pudo cachear:', url, err.message);
                }
            }
        })
    );
    // Activar inmediatamente sin esperar a que se cierre la pestaña anterior
    self.skipWaiting();
});

// ── Activate: limpiar cachés viejos ───────────────────────────────────────────

self.addEventListener('activate', (event) => {
    console.log('[SW] Activado v6');
    event.waitUntil(
        caches.keys().then((keys) =>
            Promise.all(
                keys
                    .filter((k) => k !== CACHE_NAME)
                    .map((k) => {
                        console.log('[SW] Eliminando caché viejo:', k);
                        return caches.delete(k);
                    })
            )
        ).then(() => clients.claim())
    );
});

// ── Fetch: Network-First con fallback a caché ─────────────────────────────────

self.addEventListener('fetch', (event) => {
    const req = event.request;
    const url = new URL(req.url);

    // Solo manejamos GET
    if (req.method !== 'GET') return;

    // /ping y /sync van siempre a la red (no cachear)
    if (url.pathname === '/ping' || url.pathname === '/sync' || url.pathname === '/gastos/sync') return;

    // Para recursos externos (CDN), usar Cache-First
    if (url.origin !== self.location.origin) {
        event.respondWith(cacheFirstExternal(req));
        return;
    }

    // Para todo lo demás: Network-First con fallback a caché
    event.respondWith(networkFirstWithCache(req));
});

// Cache-First para recursos externos (Bootstrap CDN, etc.)
async function cacheFirstExternal(req) {
    const cache = await caches.open(CACHE_NAME);
    const cached = await cache.match(req);
    if (cached) return cached;
    try {
        const networkRes = await fetch(req);
        if (networkRes.ok) cache.put(req, networkRes.clone());
        return networkRes;
    } catch {
        return new Response('', { status: 503 });
    }
}

// Network-First para páginas propias
async function networkFirstWithCache(req) {
    const cache = await caches.open(CACHE_NAME);
    try {
        // Intentar red con timeout de 5 segundos
        const controller = new AbortController();
        const timeoutId = setTimeout(() => controller.abort(), 5000);
        const networkRes = await fetch(req, { signal: controller.signal });
        clearTimeout(timeoutId);

        if (networkRes.ok) {
            // Guardar copia fresca en caché
            cache.put(req, networkRes.clone());
        }
        return networkRes;
    } catch {
        // Sin red (o timeout) → buscar en caché
        const cached = await cache.match(req);
        if (cached) {
            console.log('[SW] Sirviendo desde caché:', req.url);
            return cached;
        }

        // Si es navegación y no hay caché exacta, intentar la raíz
        if (req.mode === 'navigate') {
            const root = await cache.match('/');
            if (root) {
                console.log('[SW] Fallback a raíz cacheada');
                return root;
            }
        }

        // Último recurso: página de error offline
        return new Response(
            `<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Sin conexión</title>
  <style>
    body { font-family: 'Segoe UI', sans-serif; background: #f0f2f5; display: flex;
           align-items: center; justify-content: center; min-height: 100vh; margin: 0; }
    .card { background: #fff; border-radius: 16px; padding: 40px 32px; text-align: center;
            max-width: 340px; box-shadow: 0 4px 24px rgba(0,0,0,0.10); }
    h2 { color: #202124; margin-bottom: 10px; }
    p  { color: #5f6368; margin-bottom: 24px; }
    button { background: #1a73e8; color: #fff; border: none; border-radius: 8px;
             padding: 12px 28px; font-size: 16px; font-weight: 600; cursor: pointer; }
  </style>
</head>
<body>
  <div class="card">
    <div style="font-size:48px">📵</div>
    <h2>Sin conexión</h2>
    <p>La app no pudo cargar. Asegúrate de haber abierto la app con internet al menos una vez para activar el modo offline.</p>
    <button onclick="location.reload()">Reintentar</button>
  </div>
</body>
</html>`,
            {
                status: 503,
                headers: { 'Content-Type': 'text/html; charset=utf-8' }
            }
        );
    }
}
