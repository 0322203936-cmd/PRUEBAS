/**
 * offline.js — Módulo de soporte offline para Inventario App
 *
 * - Almacena registros pendientes en IndexedDB cuando no hay conexión
 * - Sincroniza automáticamente con el servidor al recuperar internet
 * - Expone funciones globales usadas por index.html y registros.html
 */

const DB_NAME    = 'inventario-offline';
const DB_VERSION = 2;          // subimos version para agregar el nuevo store
const STORE_NAME = 'pendientes';
const STORE_GASTOS = 'gastos-pendientes';

// ── IndexedDB helpers ──────────────────────────────────────────────────────────

function abrirDB() {
    return new Promise((resolve, reject) => {
        const req = indexedDB.open(DB_NAME, DB_VERSION);
        req.onupgradeneeded = (e) => {
            const db = e.target.result;
            if (!db.objectStoreNames.contains(STORE_NAME)) {
                db.createObjectStore(STORE_NAME, { keyPath: 'id', autoIncrement: true });
            }
            // Nuevo store para gastos (fotos)
            if (!db.objectStoreNames.contains(STORE_GASTOS)) {
                db.createObjectStore(STORE_GASTOS, { keyPath: 'id', autoIncrement: true });
            }
        };
        req.onsuccess = (e) => resolve(e.target.result);
        req.onerror   = (e) => reject(e.target.error);
    });
}

/**
 * Guarda un registro pendiente en IndexedDB.
 * @param {Object} datos - El objeto con los datos del formulario
 */
async function guardarPendiente(datos) {
    const db    = await abrirDB();
    const tx    = db.transaction(STORE_NAME, 'readwrite');
    const store = tx.objectStore(STORE_NAME);
    return new Promise((resolve, reject) => {
        const req = store.add({ ...datos, _guardadoEn: new Date().toISOString() });
        req.onsuccess = () => resolve(req.result);
        req.onerror   = () => reject(req.error);
    });
}

/**
 * Lee todos los registros pendientes de IndexedDB.
 * @returns {Promise<Array>}
 */
async function leerPendientes() {
    const db    = await abrirDB();
    const tx    = db.transaction(STORE_NAME, 'readonly');
    const store = tx.objectStore(STORE_NAME);
    return new Promise((resolve, reject) => {
        const req = store.getAll();
        req.onsuccess = () => resolve(req.result || []);
        req.onerror   = () => reject(req.error);
    });
}

/**
 * Cuenta los registros pendientes.
 * @returns {Promise<number>}
 */
async function contarPendientes() {
    const db    = await abrirDB();
    const tx    = db.transaction(STORE_NAME, 'readonly');
    const store = tx.objectStore(STORE_NAME);
    return new Promise((resolve, reject) => {
        const req = store.count();
        req.onsuccess = () => resolve(req.result);
        req.onerror   = () => reject(req.error);
    });
}

/**
 * Borra un registro pendiente por su id de IndexedDB.
 * @param {number} id
 */
async function borrarPendiente(id) {
    const db    = await abrirDB();
    const tx    = db.transaction(STORE_NAME, 'readwrite');
    const store = tx.objectStore(STORE_NAME);
    return new Promise((resolve, reject) => {
        const req = store.delete(id);
        req.onsuccess = () => resolve();
        req.onerror   = () => reject(req.error);
    });
}

/**
 * Limpia todos los registros pendientes de IndexedDB.
 */
async function limpiarPendientes() {
    const db    = await abrirDB();
    const tx    = db.transaction(STORE_NAME, 'readwrite');
    const store = tx.objectStore(STORE_NAME);
    return new Promise((resolve, reject) => {
        const req = store.clear();
        req.onsuccess = () => resolve();
        req.onerror   = () => reject(req.error);
    });
}

// ── Conectividad ───────────────────────────────────────────────────────────────

/**
 * Verifica conectividad real con el SERVIDOR haciendo un fetch a /ping.
 * Distingue entre:
 *   - Servidor accesible (online=true, serverUp=true)
 *   - Internet pero servidor dormido/caído (online=true, serverUp=false) — puede tardar en despertar
 *   - Sin internet en absoluto (online=false, serverUp=false)
 *
 * NOTA: En Render.com (plan gratuito) el servidor puede tardar ~15-30s en despertar.
 * Por eso usamos navigator.onLine para saber si hay internet, y /ping para saber
 * si el servidor ya está despierto.
 *
 * @returns {Promise<{internet: boolean, serverUp: boolean}>}
 */
async function checkConectividad() {
    // Paso 1: verificar si el navegador tiene internet (rápido, sin red real)
    const tieneInternet = navigator.onLine;

    if (!tieneInternet) {
        return { internet: false, serverUp: false };
    }

    // Paso 2: intentar llegar al servidor (puede estar dormido en Render)
    try {
        const res = await fetch('/ping', {
            method: 'GET',
            cache: 'no-store',
            signal: AbortSignal.timeout(6000)  // 6 segundos de espera
        });
        return { internet: true, serverUp: res.ok };
    } catch {
        // Hay internet pero el servidor no responde (dormido o sin red real)
        return { internet: tieneInternet, serverUp: false };
    }
}

/**
 * Versión simplificada: ¿puede el servidor procesar solicitudes ahora mismo?
 * @returns {Promise<boolean>}
 */
async function isOnline() {
    const { serverUp } = await checkConectividad();
    return serverUp;
}

// ── Sincronización ─────────────────────────────────────────────────────────────

let _sincronizando = false;

/**
 * Envía todos los registros pendientes al servidor (/sync).
 * Se llama automáticamente al detectar internet.
 * @returns {Promise<number>} Cantidad de registros sincronizados
 */
async function syncPendientes() {
    if (_sincronizando) return 0;
    _sincronizando = true;

    try {
        const pendientes = await leerPendientes();
        if (!pendientes.length) return 0;

        const online = await isOnline();
        if (!online) return 0;

        const res = await fetch('/sync', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ pendientes }),
            signal: AbortSignal.timeout(30000)  // 30s para subir a SharePoint
        });

        if (res.ok) {
            const data = await res.json();
            if (data.ok) {
                await limpiarPendientes();
                actualizarBannerOffline();
                return pendientes.length;
            }
        }
        return 0;
    } catch (err) {
        console.warn('[Offline] Error al sincronizar:', err);
        return 0;
    } finally {
        _sincronizando = false;
    }
}

// ── Gastos (fotos de tickets) ───────────────────────────────────────────────

/**
 * Guarda un gasto pendiente (con fotos) en IndexedDB.
 */
async function guardarGastoPendiente(datos) {
    const db    = await abrirDB();
    const tx    = db.transaction(STORE_GASTOS, 'readwrite');
    const store = tx.objectStore(STORE_GASTOS);
    return new Promise((resolve, reject) => {
        const req = store.add({ ...datos, _guardadoEn: new Date().toISOString() });
        req.onsuccess = () => resolve(req.result);
        req.onerror   = () => reject(req.error);
    });
}

/**
 * Lee todos los gastos pendientes de IndexedDB.
 */
async function leerGastosPendientes() {
    const db    = await abrirDB();
    const tx    = db.transaction(STORE_GASTOS, 'readonly');
    const store = tx.objectStore(STORE_GASTOS);
    return new Promise((resolve, reject) => {
        const req = store.getAll();
        req.onsuccess = () => resolve(req.result || []);
        req.onerror   = () => reject(req.error);
    });
}

/**
 * Cuenta los gastos pendientes.
 */
async function contarGastosPendientes() {
    const db    = await abrirDB();
    const tx    = db.transaction(STORE_GASTOS, 'readonly');
    const store = tx.objectStore(STORE_GASTOS);
    return new Promise((resolve, reject) => {
        const req = store.count();
        req.onsuccess = () => resolve(req.result);
        req.onerror   = () => reject(req.error);
    });
}

/**
 * Limpia todos los gastos pendientes.
 */
async function limpiarGastosPendientes() {
    const db    = await abrirDB();
    const tx    = db.transaction(STORE_GASTOS, 'readwrite');
    const store = tx.objectStore(STORE_GASTOS);
    return new Promise((resolve, reject) => {
        const req = store.clear();
        req.onsuccess = () => resolve();
        req.onerror   = () => reject(req.error);
    });
}

let _sincronizandoGastos = false;

/**
 * Envia los gastos pendientes al servidor (/gastos/sync).
 */
async function syncGastosPendientes() {
    if (_sincronizandoGastos) return 0;
    _sincronizandoGastos = true;

    try {
        const pendientes = await leerGastosPendientes();
        if (!pendientes.length) return 0;

        const online = await isOnline();
        if (!online) return 0;

        const res = await fetch('/gastos/sync', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ pendientes }),
            signal: AbortSignal.timeout(60000)
        });

        if (res.ok) {
            const data = await res.json();
            if (data.ok) {
                await limpiarGastosPendientes();
                return pendientes.length;
            }
        }
        return 0;
    } catch (err) {
        console.warn('[Gastos Offline] Error al sincronizar:', err);
        return 0;
    } finally {
        _sincronizandoGastos = false;
    }
}

// ── Banner de estado ───────────────────────────────────────────────────────────

/**
 * Actualiza el banner visual de estado online/offline.
 */
async function actualizarBannerOffline() {
    const banner = document.getElementById('offlineBanner');
    if (!banner) return;

    const { internet, serverUp } = await checkConectividad();
    const pendientes = await contarPendientes();

    if (serverUp) {
        // Servidor accesible
        if (pendientes > 0) {
            banner.className = 'offline-banner syncing';
            banner.innerHTML = `<span>🔄 Sincronizando ${pendientes} registro${pendientes !== 1 ? 's' : ''}…</span>`;
            banner.style.display = 'flex';
            syncPendientes().then(n => {
                if (n > 0) mostrarToast(`✅ ${n} registro${n !== 1 ? 's' : ''} sincronizado${n !== 1 ? 's' : ''} con SharePoint`);
                actualizarBannerOffline();
            });
        } else {
            banner.style.display = 'none';
        }
    } else if (internet && !serverUp) {
        // Hay internet pero el servidor está despertando (Render)
        banner.className = 'offline-banner waking';
        if (pendientes > 0) {
            banner.innerHTML = `<span>⏳ Esperando servidor… <b>${pendientes}</b> registro${pendientes !== 1 ? 's' : ''} pendiente${pendientes !== 1 ? 's' : ''}</span>`;
        } else {
            banner.innerHTML = `<span>⏳ Conectando con el servidor…</span>`;
        }
        banner.style.display = 'flex';
    } else {
        // Sin internet
        if (pendientes > 0) {
            banner.className = 'offline-banner offline';
            banner.innerHTML = `<span>📵 Sin conexión — <b>${pendientes}</b> registro${pendientes !== 1 ? 's' : ''} pendiente${pendientes !== 1 ? 's' : ''}</span>`;
        } else {
            banner.className = 'offline-banner offline';
            banner.innerHTML = `<span>📵 Sin conexión — guardando localmente</span>`;
        }
        banner.style.display = 'flex';
    }
}

// ── Toast ──────────────────────────────────────────────────────────────────────

/**
 * Muestra un mensaje toast temporal en la parte superior.
 * @param {string} msg
 * @param {string} tipo - 'success' | 'error' | 'info'
 */
function mostrarToast(msg, tipo = 'success') {
    let toast = document.getElementById('offlineToast');
    if (!toast) {
        toast = document.createElement('div');
        toast.id = 'offlineToast';
        document.body.appendChild(toast);
    }
    toast.className = `offline-toast toast-${tipo}`;
    toast.textContent = msg;
    toast.style.display = 'block';
    toast.style.opacity = '1';

    clearTimeout(toast._timer);
    toast._timer = setTimeout(() => {
        toast.style.opacity = '0';
        setTimeout(() => { toast.style.display = 'none'; }, 400);
    }, 3500);
}

// ── Estilos del banner y toast ─────────────────────────────────────────────────

(function inyectarEstilos() {
    if (document.getElementById('offline-styles')) return;
    const style = document.createElement('style');
    style.id = 'offline-styles';
    style.textContent = `
        .offline-banner {
            display: none;
            align-items: center;
            justify-content: center;
            padding: 8px 16px;
            font-size: 13px;
            font-weight: 600;
            text-align: center;
            gap: 8px;
            transition: background 0.3s;
        }
        .offline-banner.offline {
            background: #fff3cd;
            color: #856404;
            border-bottom: 1px solid #ffc107;
        }
        .offline-banner.syncing {
            background: #d1ecf1;
            color: #0c5460;
            border-bottom: 1px solid #bee5eb;
        }
        .offline-banner.waking {
            background: #e8eaf6;
            color: #3949ab;
            border-bottom: 1px solid #9fa8da;
        }

        #offlineToast {
            display: none;
            position: fixed;
            top: 16px;
            left: 50%;
            transform: translateX(-50%);
            z-index: 9999;
            padding: 10px 20px;
            border-radius: 24px;
            font-size: 14px;
            font-weight: 600;
            box-shadow: 0 4px 16px rgba(0,0,0,0.18);
            transition: opacity 0.4s;
            white-space: nowrap;
            max-width: 90vw;
            text-align: center;
        }
        .toast-success { background: #1e7e34; color: #fff; }
        .toast-error   { background: #c82333; color: #fff; }
        .toast-info    { background: #117a8b; color: #fff; }

        .badge-pendiente {
            display: inline-block;
            background: #fd7e14;
            color: #fff;
            border-radius: 10px;
            padding: 1px 8px;
            font-size: 11px;
            font-weight: 700;
            margin-left: 6px;
        }
    `;
    document.head.appendChild(style);
})();

// ── Inicialización automática ──────────────────────────────────────────────────

window.addEventListener('online', async () => {
    console.log('[Offline] Conexión recuperada, esperando servidor…');
    await actualizarBannerOffline();
    // Intentar sincronizar — puede fallar si el servidor aún está despertando en Render
    // Se reintenta cada 15 segundos hasta lograrlo
    let intentos = 0;
    const MAX_INTENTOS = 4;
    const reintentar = setInterval(async () => {
        intentos++;
        const n = await syncPendientes();
        const g = await syncGastosPendientes();
        if (n > 0 || g > 0) {
            if (n > 0) mostrarToast(`✅ ${n} inventario${n !== 1 ? 's' : ''} sincronizado${n !== 1 ? 's' : ''} con SharePoint`);
            if (g > 0) mostrarToast(`✅ ${g} gasto${g !== 1 ? 's' : ''} sincronizado${g !== 1 ? 's' : ''} con SharePoint`);
            actualizarBannerOffline();
            clearInterval(reintentar);
        } else if (intentos >= MAX_INTENTOS) {
            clearInterval(reintentar);
            actualizarBannerOffline();
        }
    }, 15000);
});

window.addEventListener('offline', () => {
    console.log('[Offline] Sin conexión.');
    actualizarBannerOffline();
});

// Exportar funciones al scope global para uso en templates
window.offlineApp = {
    guardarPendiente,
    leerPendientes,
    contarPendientes,
    borrarPendiente,
    limpiarPendientes,
    isOnline,
    checkConectividad,
    syncPendientes,
    actualizarBannerOffline,
    mostrarToast,
    // Gastos
    guardarGastoPendiente,
    leerGastosPendientes,
    contarGastosPendientes,
    limpiarGastosPendientes,
    syncGastosPendientes
};
