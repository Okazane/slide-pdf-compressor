const CACHE_NAME = 'slide-pdf-compressor-pwa-v2';

const PRECACHE_ASSETS = [
    './',
    './index.html',
    './manifest.json',
    './favicon.svg',
    './fonts.css',
    './style.css',
    './db.js',
    './offline-engine.js',
    './app.js',
    './icons/icon-192.png',
    './icons/icon-512.png',
    './vendor/dexie.min.js',
    './vendor/jszip.min.js',
    './vendor/pdf-lib.min.js',
    './vendor/pdf.min.mjs',
    './vendor/pdf.worker.min.mjs',
    './vendor/standard_fonts/FoxitDingbats.pfb',
    './vendor/standard_fonts/FoxitFixed.pfb',
    './vendor/standard_fonts/FoxitFixedBold.pfb',
    './vendor/standard_fonts/FoxitFixedBoldItalic.pfb',
    './vendor/standard_fonts/FoxitFixedItalic.pfb',
    './vendor/standard_fonts/FoxitSerif.pfb',
    './vendor/standard_fonts/FoxitSerifBold.pfb',
    './vendor/standard_fonts/FoxitSerifBoldItalic.pfb',
    './vendor/standard_fonts/FoxitSerifItalic.pfb',
    './vendor/standard_fonts/FoxitSymbol.pfb',
    './vendor/standard_fonts/LiberationSans-Bold.ttf',
    './vendor/standard_fonts/LiberationSans-BoldItalic.ttf',
    './vendor/standard_fonts/LiberationSans-Italic.ttf',
    './vendor/standard_fonts/LiberationSans-Regular.ttf',
    './fonts/be-vietnam-pro-latin-400-normal.woff2',
    './fonts/be-vietnam-pro-latin-500-normal.woff2',
    './fonts/be-vietnam-pro-latin-600-normal.woff2',
    './fonts/be-vietnam-pro-latin-700-normal.woff2',
    './fonts/be-vietnam-pro-latin-800-normal.woff2',
    './fonts/be-vietnam-pro-latin-ext-400-normal.woff2',
    './fonts/be-vietnam-pro-latin-ext-500-normal.woff2',
    './fonts/be-vietnam-pro-latin-ext-600-normal.woff2',
    './fonts/be-vietnam-pro-latin-ext-700-normal.woff2',
    './fonts/be-vietnam-pro-latin-ext-800-normal.woff2',
    './fonts/be-vietnam-pro-vietnamese-400-normal.woff2',
    './fonts/be-vietnam-pro-vietnamese-500-normal.woff2',
    './fonts/be-vietnam-pro-vietnamese-600-normal.woff2',
    './fonts/be-vietnam-pro-vietnamese-700-normal.woff2',
    './fonts/be-vietnam-pro-vietnamese-800-normal.woff2'
];

self.addEventListener('install', (event) => {
    event.waitUntil(
        (async () => {
            const cache = await caches.open(CACHE_NAME);
            await cache.addAll(PRECACHE_ASSETS);
            await self.skipWaiting();
        })()
    );
});

self.addEventListener('activate', (event) => {
    event.waitUntil(
        (async () => {
            const keys = await caches.keys();
            await Promise.all(
                keys
                    .filter((key) => key !== CACHE_NAME)
                    .map((key) => caches.delete(key))
            );
            await self.clients.claim();
        })()
    );
});

function isNavigationOrHtml(request, url) {
    if (request.mode === 'navigate') return true;
    const accept = request.headers.get('accept') || '';
    if (accept.includes('text/html')) return true;
    return url.pathname.endsWith('/') || url.pathname.endsWith('.html');
}

function isImmutableVendorOrFont(url) {
    return (
        url.pathname.includes('/fonts/') ||
        url.pathname.includes('/vendor/') ||
        url.pathname.includes('/icons/')
    );
}

self.addEventListener('fetch', (event) => {
    const { request } = event;
    if (request.method !== 'GET') return;

    const url = new URL(request.url);
    if (url.origin !== self.location.origin) return;
    if (url.pathname.startsWith('/api/')) return;

    // 1. HTML / Document Navigation: Network-First with synchronous res.clone()
    if (isNavigationOrHtml(request, url)) {
        event.respondWith(
            (async () => {
                try {
                    const networkRes = await fetch(request);
                    if (networkRes && networkRes.ok) {
                        const cloneForReq = networkRes.clone();
                        const cloneForShell = networkRes.clone();
                        try {
                            const cache = await caches.open(CACHE_NAME);
                            await Promise.all([
                                cache.put(request, cloneForReq),
                                cache.put('./index.html', cloneForShell),
                            ]);
                        } catch (_cacheErr) {
                            // Never fail a successful network response if Cache Storage quota/permissions reject
                        }
                    }
                    return networkRes;
                } catch (_err) {
                    const cached =
                        (await caches.match(request)) ||
                        (await caches.match('./index.html')) ||
                        (await caches.match('./'));
                    return cached || Response.error();
                }
            })()
        );
        return;
    }

    // 2. Immutable Fonts / Vendor / Icons: Cache-First with Response.error() fallback
    if (isImmutableVendorOrFont(url)) {
        event.respondWith(
            (async () => {
                const cached = await caches.match(request);
                if (cached) return cached;
                try {
                    const networkRes = await fetch(request);
                    if (networkRes && networkRes.ok) {
                        const copy = networkRes.clone();
                        try {
                            const cache = await caches.open(CACHE_NAME);
                            await cache.put(request, copy);
                        } catch (_cacheErr) {
                            // Ignore cache write error and still return valid networkRes
                        }
                    }
                    return networkRes;
                } catch (_err) {
                    return Response.error();
                }
            })()
        );
        return;
    }

    // 3. App Shell JS / CSS / Manifest: Stale-While-Revalidate with synchronous clone & Response.error() fallback
    event.respondWith(
        (async () => {
            const cached = await caches.match(request);
            const revalidatePromise = (async () => {
                try {
                    const networkRes = await fetch(request);
                    if (networkRes && networkRes.ok) {
                        const copy = networkRes.clone();
                        const cache = await caches.open(CACHE_NAME);
                        await cache.put(request, copy);
                    }
                    return networkRes;
                } catch (_err) {
                    return null;
                }
            })();

            if (cached) {
                event.waitUntil(revalidatePromise);
                return cached;
            }

            const fresh = await revalidatePromise;
            return fresh || Response.error();
        })()
    );
});
