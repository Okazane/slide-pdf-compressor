/**
 * Automated Offline PWA, Security & Local Database Verification Suite
 * Run via: node tests/verify_offline_pwa.mjs
 */

import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const ROOT_DIR = path.resolve(__dirname, '..');
const STATIC_DIR = path.join(ROOT_DIR, 'static');

const dbModuleUrl = pathToFileURL(path.join(STATIC_DIR, 'db.js')).href;
const {
    MAX_BACKUP_BYTES,
    MAX_HISTORY_ITEMS,
    addHistoryEntry,
    clearLocalDatabase,
    exportDatabaseToJson,
    importDatabaseFromJson,
    listHistoryEntries,
    loadPreferences,
    savePreferences,
    validateBackupJsonText,
    validateHistoryRecord,
    validatePreferences,
    validateSafeJpegDataUri,
} = await import(dbModuleUrl);

const engineModuleUrl = pathToFileURL(path.join(STATIC_DIR, 'offline-engine.js')).href;
const {
    MAX_CANVAS_PIXELS,
    MAX_FILE_BYTES,
    MAX_ZIP_ENTRIES,
    MAX_ZIP_UNCOMPRESSED_BYTES,
    parseSafeXmlString,
    throwIfAborted,
    validateInputFileSafety,
    validateZipSafety,
} = await import(engineModuleUrl);

let passed = 0;
async function check(name, fn) {
    await fn();
    passed++;
    console.log(`  [PASS] ${name}`);
}

console.log('=== 1. Verifying JS Syntax Across Static PWA Modules ===');
for (const relJs of ['db.js', 'offline-engine.js', 'app.js', 'sw.js']) {
    check(`Syntax check static/${relJs}`, () => {
        execFileSync(process.execPath, ['--check', path.join(STATIC_DIR, relJs)], {
            stdio: 'pipe',
        });
    });
}

console.log('\n=== 2. Verifying Strict .JSON Backup Schema Validator (Dexie DB) ===');

// Minimal valid JPEG SOI header (FF D8 FF E0 ...) encoded as base64
const validJpegBytes = Buffer.from([
    0xff, 0xd8, 0xff, 0xe0, 0x00, 0x10, 0x4a, 0x46,
    0x49, 0x46, 0x00, 0x01, 0xff, 0xd9, 0x00, 0x00,
]);
const validJpegDataUri = `data:image/jpeg;base64,${validJpegBytes.toString('base64')}`;

check('Accepts valid preferences + history .json backup', () => {
    const validPayload = JSON.stringify({
        app: 'slide-pdf-compressor',
        version: 1,
        exportedAt: '2026-11-20T12:00:00.000Z',
        preferences: {
            theme: 'dark',
            targetMb: 10,
            mode: 'flatten',
            outputFormat: 'pdf',
        },
        history: [
            {
                filename: 'Bài giảng Tiếng Việt.pptx',
                fileType: 'pptx',
                jobType: 'compress',
                mode: 'pptx_direct',
                outputFormat: 'pptx',
                targetMb: 10,
                targetMet: true,
                originalSizeMb: 24.5,
                compressedSizeMb: 6.2,
                reductionPercent: 74.7,
                itemCount: 18,
                previewBase64: validJpegDataUri,
                createdAt: 1760000000000,
            },
        ],
    });
    const res = validateBackupJsonText(validPayload);
    assert.equal(res.preferences.theme, 'dark');
    assert.equal(res.history.length, 1);
    assert.equal(res.history[0].filename, 'Bài giảng Tiếng Việt.pptx');
});

check('Blocks Prototype Pollution (__proto__, constructor, prototype)', () => {
    const protoRoot = '{"version":1,"__proto__":{"polluted":true},"preferences":{"theme":"dark"}}';
    assert.throws(() => validateBackupJsonText(protoRoot), /Prototype Pollution|bị cấm/i);

    const ctorNested = '{"version":1,"preferences":{"theme":"dark","constructor":{"prototype":{"x":1}}}}';
    assert.throws(() => validateBackupJsonText(ctorNested), /Prototype Pollution|bị cấm/i);

    const protoHistory = JSON.stringify({
        version: 1,
        history: [],
    }).replace('"history":[]', '"history":[{"prototype":{"evil":1}}]');
    assert.throws(() => validateBackupJsonText(protoHistory), /Prototype Pollution|bị cấm/i);

    assert.equal(({}).polluted, undefined);
});

check('Rejects SVG / HTML / javascript: / Spoofed non-JPEG previewBase64 (XSS Defense)', () => {
    const svgUri = `data:image/svg+xml;base64,${Buffer.from('<svg onload="alert(1)"/>').toString('base64')}`;
    assert.throws(() => validateSafeJpegDataUri(svgUri), /jpeg/i);

    assert.throws(() => validateSafeJpegDataUri('javascript:alert(1)'), /jpeg/i);
    assert.throws(() => validateSafeJpegDataUri('data:text/html;base64,PHNjcmlwdD5hbGVydCgxKTwvc2NyaXB0Pg=='), /jpeg/i);

    // Spoofed prefix with SVG payload bytes instead of 0xFF 0xD8 0xFF JPEG magic
    const spoofedJpeg = `data:image/jpeg;base64,${Buffer.from('<svg onload="alert(1)"></svg>').toString('base64')}`;
    assert.throws(() => validateSafeJpegDataUri(spoofedJpeg), /JPEG|base64/i);
});

check('Enforces numeric bounds, history cap (<= 30), and <= 5 MB backup limit', () => {
    assert.throws(() => validatePreferences({ theme: 'neon', targetMb: 10, mode: 'flatten', outputFormat: 'pdf' }));
    assert.throws(() => validatePreferences({ theme: 'dark', targetMb: -5, mode: 'flatten', outputFormat: 'pdf' }));
    assert.throws(() => validatePreferences({ theme: 'dark', targetMb: 999, mode: 'flatten', outputFormat: 'pdf' }));

    assert.throws(() =>
        validateHistoryRecord({
            filename: 'bad\x00file.pdf',
            fileType: 'pdf',
            jobType: 'compress',
            mode: 'flatten',
            outputFormat: 'pdf',
            originalSizeMb: 5,
            compressedSizeMb: 2,
            reductionPercent: 60,
            itemCount: 3,
            createdAt: 1760000000000,
        })
    );

    const tooMany = Array.from({ length: MAX_HISTORY_ITEMS + 1 }, (_, i) => ({
        filename: `doc_${i}.pdf`,
        fileType: 'pdf',
        jobType: 'compress',
        mode: 'flatten',
        outputFormat: 'pdf',
        originalSizeMb: 2,
        compressedSizeMb: 1,
        reductionPercent: 50,
        itemCount: 1,
        previewBase64: null,
        createdAt: 1760000000000 + i,
    }));
    assert.throws(() => validateBackupJsonText(JSON.stringify({ version: 1, history: tooMany })));

    const hugeStr = 'x'.repeat(MAX_BACKUP_BYTES + 64);
    assert.throws(() => validateBackupJsonText(hugeStr), /5 MB/i);
});

console.log('\n=== 3. Verifying Zero External CDNs, Zero Inline Scripts & Safe DOM ===');

const indexHtml = fs.readFileSync(path.join(STATIC_DIR, 'index.html'), 'utf8');

check('static/index.html has zero external Google Fonts / CDN URLs and zero inline scripts', () => {
    assert.equal(/https?:\/\//i.test(indexHtml), false, 'Found external http(s) URL in index.html');
    assert.equal(/fonts\.googleapis\.com|fonts\.gstatic\.com/i.test(indexHtml), false);

    // Every <script> tag must have a relative src="..." and empty body
    const scriptTags = [...indexHtml.matchAll(/<script\b([^>]*)>([\s\S]*?)<\/script>/gi)];
    assert.ok(scriptTags.length >= 4, 'Expected vendored + module script tags');
    for (const [, attrs, inner] of scriptTags) {
        assert.ok(/\bsrc\s*=\s*"[^"]+"/i.test(attrs), `<script> tag missing src attribute: ${attrs}`);
        assert.equal(inner.trim(), '', 'Inline script code found inside <script> tag');
        assert.equal(/\bsrc\s*=\s*"\//i.test(attrs), false, `Use relative asset path instead of leading /: ${attrs}`);
    }

    // Zero inline event handler attributes (onclick=, onload=, onerror=, etc.)
    assert.equal(/\son[a-z]+\s*=/i.test(indexHtml), false, 'Inline event handler attribute found in index.html');
});

check('App JS & CSS contain zero external CDN/font dependencies, zero LAN IPs, and zero innerHTML sinks', () => {
    for (const rel of ['fonts.css', 'style.css', 'db.js', 'offline-engine.js', 'app.js', 'sw.js']) {
        const content = fs.readFileSync(path.join(STATIC_DIR, rel), 'utf8');
        assert.equal(
            /fonts\.googleapis\.com|fonts\.gstatic\.com|cdn\.jsdelivr\.net|unpkg\.com|cdnjs\.cloudflare\.com/i.test(content),
            false,
            `External CDN/font URL found in static/${rel}`
        );
        assert.equal(
            /\b(?:192\.168\.\d+\.\d+|10\.\d+\.\d+\.\d+)\b/.test(content),
            false,
            `Hardcoded LAN IP found in static/${rel}`
        );
        if (rel.endsWith('.js')) {
            assert.equal(/\.innerHTML\b/.test(content), false, `Unsafe innerHTML usage found in static/${rel}`);
        }
    }
});

console.log('\n=== 4. Verifying PWA Manifest, Service Worker Precache & vercel.json ===');

check('All icons in static/manifest.json exist on disk', () => {
    const manifest = JSON.parse(fs.readFileSync(path.join(STATIC_DIR, 'manifest.json'), 'utf8'));
    assert.equal(manifest.display, 'standalone');
    assert.ok(Array.isArray(manifest.icons) && manifest.icons.length >= 2);
    for (const icon of manifest.icons) {
        const iconPath = path.join(STATIC_DIR, icon.src);
        assert.ok(fs.existsSync(iconPath), `Missing manifest icon on disk: ${icon.src}`);
        assert.ok(fs.statSync(iconPath).size > 100, `Manifest icon is empty: ${icon.src}`);
    }
});

check('Every asset in static/sw.js PRECACHE_ASSETS exists on disk', () => {
    const swText = fs.readFileSync(path.join(STATIC_DIR, 'sw.js'), 'utf8');
    const match = swText.match(/const\s+PRECACHE_ASSETS\s*=\s*\[([\s\S]*?)\];/);
    assert.ok(match, 'PRECACHE_ASSETS array not found in static/sw.js');
    const assetList = [...match[1].matchAll(/'([^']+)'/g)].map((m) => m[1]);
    assert.ok(assetList.length >= 35, `Expected full precache list, got ${assetList.length}`);

    for (const relAsset of assetList) {
        const normalized = relAsset === './' ? 'index.html' : relAsset.replace(/^\.\//, '');
        const fullPath = path.join(STATIC_DIR, normalized);
        assert.ok(fs.existsSync(fullPath), `Missing precached asset on disk: ${relAsset} -> ${fullPath}`);
        assert.ok(fs.statSync(fullPath).size > 0, `Empty precached asset: ${relAsset}`);
    }

    assert.ok(swText.includes('.clone()'), 'Service Worker must clone responses synchronously');
    assert.ok(swText.includes('Response.error()'), 'Service Worker must fallback to Response.error()');
});

check('vercel.json configures static outputDirectory and strict CSP headers', () => {
    const vercelCfg = JSON.parse(fs.readFileSync(path.join(ROOT_DIR, 'vercel.json'), 'utf8'));
    assert.equal(vercelCfg.buildCommand, '');
    assert.equal(vercelCfg.outputDirectory, 'static');
    const allHeaders = vercelCfg.headers.flatMap((h) => h.headers);
    const csp = allHeaders.find((h) => h.key.toLowerCase() === 'content-security-policy');
    assert.ok(csp, 'Missing Content-Security-Policy in vercel.json');
    assert.ok(csp.value.includes("script-src 'self'"), 'CSP must restrict script-src to self');
    assert.equal(csp.value.includes("script-src 'self' 'unsafe-inline'"), false, 'CSP must forbid inline scripts');
    assert.ok(allHeaders.some((h) => h.key === 'X-Content-Type-Options' && h.value === 'nosniff'));
    assert.ok(allHeaders.some((h) => h.key === 'Referrer-Policy'));
    const swHeaderRule = vercelCfg.headers.find((h) => h.source === '/sw.js');
    assert.ok(swHeaderRule, 'Missing /sw.js cache-control rule in vercel.json');
    assert.ok(swHeaderRule.headers.some((h) => h.value.includes('no-cache')));
});

console.log('\n=== 5. Verifying Client-Side Offline Engine Safety Bounds & AbortController ===');

check('Enforces <= 500 MB file bound, <= 64 MP pixel cap, ZIP bomb & XML DTD/ENTITY guards, and AbortController', () => {
    assert.equal(MAX_FILE_BYTES, 500 * 1024 * 1024);
    assert.equal(MAX_ZIP_ENTRIES, 5000);
    assert.equal(MAX_ZIP_UNCOMPRESSED_BYTES, 1_500_000_000);
    assert.equal(MAX_CANVAS_PIXELS, 64_000_000);

    assert.throws(() => validateInputFileSafety({ name: 'huge.pdf', size: MAX_FILE_BYTES + 1 }), /500 MB/);
    assert.throws(() => validateInputFileSafety({ name: 'empty.pdf', size: 0 }), /0 byte/);
    assert.throws(() => validateInputFileSafety({ name: 'script.exe', size: 100 }), /không được hỗ trợ/);

    // ZIP path traversal & uncompressed size bomb
    assert.throws(() => validateZipSafety({ files: { '../evil.xml': { _data: { uncompressedSize: 10 } } } }), /traversal/i);
    assert.throws(
        () => validateZipSafety({ files: { 'ppt/media/bomb.bmp': { _data: { uncompressedSize: 2_000_000_000 } } } }),
        /ZIP bomb/i
    );

    // XML <!ENTITY / <!DOCTYPE bomb rejection before DOMParser
    assert.throws(() => parseSafeXmlString('<?xml version="1.0"?><!DOCTYPE x [<!ENTITY a "b">]><r/>'), /DOCTYPE|ENTITY/i);

    // AbortController cancellation
    const ctrl = new AbortController();
    throwIfAborted(ctrl.signal);
    ctrl.abort();
    assert.throws(() => throwIfAborted(ctrl.signal), { name: 'AbortError' });
});

await check('Resilient in-memory fallback when IndexedDB is blocked or unavailable (Private Browsing)', async () => {
    const saved = await savePreferences({ theme: 'light', targetMb: 15, mode: 'vector', outputFormat: 'pptx' });
    assert.equal(saved.theme, 'light');
    assert.equal(saved.targetMb, 15);

    await addHistoryEntry({
        filename: 'test_private.pdf',
        fileType: 'pdf',
        jobType: 'compress',
        mode: 'flatten',
        outputFormat: 'pdf',
        targetMb: 10,
        targetMet: true,
        originalSizeMb: 8,
        compressedSizeMb: 3,
        reductionPercent: 62.5,
        itemCount: 4,
        previewBase64: validJpegDataUri,
        createdAt: 1760000000000,
    });
    const list = await listHistoryEntries();
    assert.equal(list.length, 1);
    assert.equal(list[0].filename, 'test_private.pdf');

    const exported = await exportDatabaseToJson();
    await clearLocalDatabase();
    assert.equal((await listHistoryEntries()).length, 0);

    const imported = await importDatabaseFromJson(exported);
    assert.equal(imported.importedCount, 1);
    assert.equal((await loadPreferences()).theme, 'light');
    await clearLocalDatabase();
});

console.log(`\nALL ${passed} OFFLINE PWA & SECURITY VERIFICATION CHECKS PASSED!`);
