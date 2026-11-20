/**
 * Local IndexedDB Storage & Strict JSON Backup Validator for Slide & PDF Smart Compressor.
 * Powered by Dexie.js (offline-first, zero server dependency).
 */

export const MAX_BACKUP_BYTES = 5 * 1024 * 1024; // 5 MB
export const MAX_HISTORY_ITEMS = 30;
export const MAX_PREVIEW_DATA_URI_LEN = 450_000;

const FORBIDDEN_KEYS = new Set(['__proto__', 'constructor', 'prototype']);
const ALLOWED_TOP_KEYS = new Set(['app', 'version', 'exportedAt', 'preferences', 'history']);
const ALLOWED_THEMES = new Set(['dark', 'light']);
const ALLOWED_MODES = new Set(['flatten', 'vector']);
const ALLOWED_PREF_FORMATS = new Set(['pdf', 'pptx']);
const ALLOWED_FILE_TYPES = new Set(['pdf', 'pptx', 'ppt']);
const ALLOWED_JOB_TYPES = new Set(['compress', 'extract']);
const ALLOWED_HISTORY_MODES = new Set(['flatten', 'vector', 'pptx_direct', 'images', 'slides']);
const ALLOWED_OUTPUT_FORMATS = new Set(['pdf', 'pptx', 'zip']);

const DEFAULT_PREFERENCES = Object.freeze({
    id: 'user_prefs',
    theme: 'dark',
    targetMb: 10,
    mode: 'flatten',
    outputFormat: 'pdf',
});

function isPlainObject(val) {
    if (val === null || typeof val !== 'object' || Array.isArray(val)) {
        return false;
    }
    const proto = Object.getPrototypeOf(val);
    return proto === Object.prototype || proto === null;
}

function assertNoForbiddenKeysDeep(node, depth = 0) {
    if (depth > 8) {
        throw new Error('Cấu trúc tệp sao lưu lồng quá sâu không hợp lệ.');
    }
    if (node === null || typeof node !== 'object') {
        return;
    }
    if (!Array.isArray(node) && !isPlainObject(node)) {
        throw new Error('Đối tượng JSON không hợp lệ.');
    }
    for (const key of Object.getOwnPropertyNames(node)) {
        if (FORBIDDEN_KEYS.has(key)) {
            throw new Error(`Phát hiện khóa bị cấm trong tệp sao lưu: "${key}".`);
        }
        assertNoForbiddenKeysDeep(node[key], depth + 1);
    }
}

function decodeBase64Prefix(b64Slice) {
    if (typeof atob === 'function') {
        const bin = atob(b64Slice);
        const out = new Uint8Array(bin.length);
        for (let i = 0; i < bin.length; i++) {
            out[i] = bin.charCodeAt(i);
        }
        return out;
    }
    if (typeof Buffer !== 'undefined') {
        return new Uint8Array(Buffer.from(b64Slice, 'base64'));
    }
    return new Uint8Array(0);
}

export function validateSafeJpegDataUri(value) {
    if (value === null || value === undefined || value === '') {
        return null;
    }
    if (typeof value !== 'string') {
        throw new Error('Ảnh xem trước (previewBase64) phải là chuỗi.');
    }
    if (value.length > MAX_PREVIEW_DATA_URI_LEN) {
        throw new Error('Ảnh xem trước vượt quá kích thước an toàn cho phép.');
    }
    const prefix = 'data:image/jpeg;base64,';
    if (!value.startsWith(prefix)) {
        throw new Error('Chỉ chấp nhận ảnh xem trước định dạng raster data:image/jpeg;base64,.');
    }
    const payload = value.slice(prefix.length);
    if (payload.length < 16 || payload.length % 4 !== 0 || !/^[A-Za-z0-9+/]+={0,2}$/.test(payload)) {
        throw new Error('Dữ liệu base64 của ảnh xem trước không hợp lệ.');
    }
    const headerBytes = decodeBase64Prefix(payload.slice(0, 16));
    if (
        headerBytes.length < 3 ||
        headerBytes[0] !== 0xff ||
        headerBytes[1] !== 0xd8 ||
        headerBytes[2] !== 0xff
    ) {
        throw new Error('Ảnh xem trước không đúng chữ ký nhị phân JPEG hợp lệ.');
    }
    return value;
}

export function validatePreferences(raw) {
    if (!isPlainObject(raw)) {
        throw new Error('Mục preferences phải là một đối tượng hợp lệ.');
    }
    assertNoForbiddenKeysDeep(raw);

    const theme = raw.theme ?? DEFAULT_PREFERENCES.theme;
    if (typeof theme !== 'string' || !ALLOWED_THEMES.has(theme)) {
        throw new Error('Giá trị giao diện (theme) không hợp lệ.');
    }

    const targetMb = Number(raw.targetMb ?? DEFAULT_PREFERENCES.targetMb);
    if (!Number.isFinite(targetMb) || targetMb < 0.5 || targetMb > 200) {
        throw new Error('Giá trị dung lượng mục tiêu (targetMb) phải nằm trong khoảng 0.5 - 200 MB.');
    }

    const mode = raw.mode ?? DEFAULT_PREFERENCES.mode;
    if (typeof mode !== 'string' || !ALLOWED_MODES.has(mode)) {
        throw new Error('Chế độ nén (mode) không hợp lệ.');
    }

    const outputFormat = raw.outputFormat ?? DEFAULT_PREFERENCES.outputFormat;
    if (typeof outputFormat !== 'string' || !ALLOWED_PREF_FORMATS.has(outputFormat)) {
        throw new Error('Định dạng xuất mặc định (outputFormat) không hợp lệ.');
    }

    return {
        id: 'user_prefs',
        theme,
        targetMb: Math.round(targetMb * 100) / 100,
        mode,
        outputFormat,
    };
}

export function validateHistoryRecord(raw) {
    if (!isPlainObject(raw)) {
        throw new Error('Bản ghi lịch sử không hợp lệ.');
    }
    assertNoForbiddenKeysDeep(raw);

    if (typeof raw.filename !== 'string') {
        throw new Error('Tên tệp trong lịch sử phải là chuỗi ký tự.');
    }
    const cleanFilename = raw.filename.trim();
    if (
        cleanFilename.length < 1 ||
        cleanFilename.length > 255 ||
        /[\x00-\x1f\x7f]/.test(cleanFilename)
    ) {
        throw new Error('Tên tệp trong lịch sử có độ dài hoặc ký tự điều khiển không hợp lệ.');
    }

    if (typeof raw.fileType !== 'string' || !ALLOWED_FILE_TYPES.has(raw.fileType)) {
        throw new Error('Loại tệp (fileType) trong lịch sử không hợp lệ.');
    }

    if (typeof raw.jobType !== 'string' || !ALLOWED_JOB_TYPES.has(raw.jobType)) {
        throw new Error('Loại tác vụ (jobType) trong lịch sử không hợp lệ.');
    }

    if (typeof raw.mode !== 'string' || !ALLOWED_HISTORY_MODES.has(raw.mode)) {
        throw new Error('Chế độ xử lý (mode) trong lịch sử không hợp lệ.');
    }

    if (typeof raw.outputFormat !== 'string' || !ALLOWED_OUTPUT_FORMATS.has(raw.outputFormat)) {
        throw new Error('Định dạng đầu ra (outputFormat) trong lịch sử không hợp lệ.');
    }

    let targetMb = null;
    if (raw.targetMb !== null && raw.targetMb !== undefined) {
        if (typeof raw.targetMb !== 'number' || !Number.isFinite(raw.targetMb) || raw.targetMb <= 0 || raw.targetMb > 500) {
            throw new Error('Dung lượng mục tiêu (targetMb) trong lịch sử vượt ngưỡng hợp lệ.');
        }
        targetMb = Math.round(raw.targetMb * 1000) / 1000;
    }

    let targetMet = null;
    if (raw.targetMet !== null && raw.targetMet !== undefined) {
        if (typeof raw.targetMet !== 'boolean') {
            throw new Error('Trường targetMet phải là kiểu boolean hoặc null.');
        }
        targetMet = raw.targetMet;
    }

    const originalSizeMb = raw.originalSizeMb;
    if (typeof originalSizeMb !== 'number' || !Number.isFinite(originalSizeMb) || originalSizeMb < 0 || originalSizeMb > 500) {
        throw new Error('Dung lượng gốc (originalSizeMb) vượt ngưỡng hợp lệ (0 - 500 MB).');
    }

    const compressedSizeMb = raw.compressedSizeMb;
    if (typeof compressedSizeMb !== 'number' || !Number.isFinite(compressedSizeMb) || compressedSizeMb < 0 || compressedSizeMb > 2000) {
        throw new Error('Dung lượng sau xử lý (compressedSizeMb) vượt ngưỡng hợp lệ.');
    }

    const reductionPercent = raw.reductionPercent ?? 0;
    if (typeof reductionPercent !== 'number' || !Number.isFinite(reductionPercent) || reductionPercent < -1000 || reductionPercent > 100) {
        throw new Error('Phần trăm giảm dung lượng (reductionPercent) không hợp lệ.');
    }

    const itemCount = raw.itemCount ?? 0;
    if (typeof itemCount !== 'number' || !Number.isInteger(itemCount) || itemCount < 0 || itemCount > 50000) {
        throw new Error('Số lượng trang/mục (itemCount) không hợp lệ.');
    }

    let createdAt = raw.createdAt;
    if (typeof createdAt === 'string') {
        const parsedTs = Date.parse(createdAt);
        if (!Number.isFinite(parsedTs)) {
            throw new Error('Thời gian tạo bản ghi (createdAt) không hợp lệ.');
        }
        createdAt = parsedTs;
    }
    if (
        typeof createdAt !== 'number' ||
        !Number.isFinite(createdAt) ||
        createdAt < 946684800000 ||
        createdAt > 4102444800000
    ) {
        throw new Error('Mốc thời gian (createdAt) nằm ngoài phạm vi hợp lệ.');
    }

    const previewBase64 = validateSafeJpegDataUri(raw.previewBase64);

    return {
        filename: cleanFilename,
        fileType: raw.fileType,
        jobType: raw.jobType,
        mode: raw.mode,
        outputFormat: raw.outputFormat,
        targetMb,
        targetMet,
        originalSizeMb: Math.round(originalSizeMb * 100) / 100,
        compressedSizeMb: Math.round(compressedSizeMb * 100) / 100,
        reductionPercent: Math.round(reductionPercent * 10) / 10,
        itemCount,
        previewBase64,
        createdAt: Math.trunc(createdAt),
    };
}

export function validateBackupJsonText(rawText) {
    if (typeof rawText !== 'string' || rawText.trim().length === 0) {
        throw new Error('Tệp sao lưu .json trống hoặc không đúng định dạng văn bản.');
    }
    const byteLen = new TextEncoder().encode(rawText).byteLength;
    if (byteLen > MAX_BACKUP_BYTES) {
        throw new Error('Dung lượng tệp sao lưu vượt giới hạn an toàn tối đa 5 MB.');
    }

    let parsed;
    try {
        parsed = JSON.parse(rawText, function reviver(key, value) {
            if (FORBIDDEN_KEYS.has(key)) {
                throw new Error(`Phát hiện khóa nguy hiểm "${key}" (ngăn chặn Prototype Pollution).`);
            }
            return value;
        });
    } catch (err) {
        if (err instanceof Error && err.message.includes('Prototype Pollution')) {
            throw err;
        }
        throw new Error('Nội dung tệp không phải JSON hợp lệ.');
    }

    if (!isPlainObject(parsed)) {
        throw new Error('Gốc dữ liệu sao lưu phải là một đối tượng JSON.');
    }
    assertNoForbiddenKeysDeep(parsed);

    for (const k of Object.keys(parsed)) {
        if (!ALLOWED_TOP_KEYS.has(k)) {
            throw new Error(`Trường dữ liệu không được phép trong bản sao lưu: "${k}".`);
        }
    }

    if (parsed.version !== undefined && parsed.version !== 1) {
        throw new Error('Phiên bản cấu trúc tệp sao lưu không được hỗ trợ.');
    }

    if (parsed.preferences === undefined && parsed.history === undefined) {
        throw new Error('Tệp sao lưu không chứa preferences hoặc history.');
    }

    const cleanPreferences = parsed.preferences !== undefined
        ? validatePreferences(parsed.preferences)
        : null;

    let cleanHistory = [];
    if (parsed.history !== undefined) {
        if (!Array.isArray(parsed.history)) {
            throw new Error('Danh sách lịch sử (history) phải là một mảng.');
        }
        if (parsed.history.length > MAX_HISTORY_ITEMS) {
            throw new Error(`Số lượng bản ghi lịch sử vượt quá giới hạn tối đa (${MAX_HISTORY_ITEMS} mục).`);
        }
        cleanHistory = parsed.history.map((item) => validateHistoryRecord(item));
    }

    return {
        version: 1,
        preferences: cleanPreferences,
        history: cleanHistory,
    };
}

let dbInstance = null;
let memoryPreferences = { ...DEFAULT_PREFERENCES };
let memoryHistory = [];

export function getSlideCompressorDB() {
    if (dbInstance) {
        return dbInstance;
    }
    const hasIndexedDB =
        typeof globalThis !== 'undefined' &&
        typeof globalThis.indexedDB !== 'undefined' &&
        globalThis.indexedDB !== null;
    const DexieCtor = typeof globalThis !== 'undefined' ? globalThis.Dexie : undefined;
    if (!hasIndexedDB || !DexieCtor) {
        return null;
    }
    try {
        const db = new DexieCtor('SlideCompressorDB');
        db.version(1).stores({
            preferences: 'id',
            history: '++id, createdAt, jobType, fileType',
        });
        dbInstance = db;
        return db;
    } catch (_err) {
        return null;
    }
}

export async function requestPersistentStorage() {
    try {
        if (
            typeof navigator !== 'undefined' &&
            navigator.storage &&
            typeof navigator.storage.persist === 'function'
        ) {
            return await navigator.storage.persist();
        }
    } catch (_err) {
        // Ignore storage persistence errors in restricted contexts
    }
    return false;
}

export async function loadPreferences() {
    const db = getSlideCompressorDB();
    if (!db) {
        return { ...memoryPreferences };
    }
    try {
        const stored = await db.table('preferences').get('user_prefs');
        if (!stored) {
            return { ...memoryPreferences };
        }
        const valid = validatePreferences(stored);
        memoryPreferences = { ...valid };
        return valid;
    } catch (_err) {
        return { ...memoryPreferences };
    }
}

export async function savePreferences(partialPrefs) {
    const current = await loadPreferences();
    const merged = validatePreferences({
        ...current,
        ...partialPrefs,
    });
    memoryPreferences = { ...merged };
    const db = getSlideCompressorDB();
    if (db) {
        try {
            await db.table('preferences').put(merged);
        } catch (_err) {
            // Gracefully fall back to in-memory preferences if IndexedDB is blocked
        }
    }
    return merged;
}

export async function addHistoryEntry(entryInput) {
    const record = validateHistoryRecord({
        ...entryInput,
        createdAt: entryInput.createdAt ?? Date.now(),
    });
    memoryHistory = [record, ...memoryHistory]
        .sort((a, b) => b.createdAt - a.createdAt)
        .slice(0, MAX_HISTORY_ITEMS);

    const db = getSlideCompressorDB();
    if (!db) {
        return record;
    }
    try {
        const table = db.table('history');
        await db.transaction('rw', table, async () => {
            await table.add(record);
            const allDesc = await table.orderBy('createdAt').reverse().toArray();
            if (allDesc.length > MAX_HISTORY_ITEMS) {
                const staleIds = allDesc
                    .slice(MAX_HISTORY_ITEMS)
                    .map((item) => item.id)
                    .filter((id) => id !== undefined);
                if (staleIds.length > 0) {
                    await table.bulkDelete(staleIds);
                }
            }
        });
    } catch (_err) {
        // Gracefully fall back to in-memory history if IndexedDB is blocked
    }
    return record;
}

export async function listHistoryEntries() {
    const db = getSlideCompressorDB();
    if (!db) {
        return [...memoryHistory];
    }
    try {
        const rows = await db
            .table('history')
            .orderBy('createdAt')
            .reverse()
            .limit(MAX_HISTORY_ITEMS)
            .toArray();
        memoryHistory = rows;
        return rows;
    } catch (_err) {
        return [...memoryHistory];
    }
}

export async function exportDatabaseToJson() {
    const preferences = await loadPreferences();
    const rawHistory = await listHistoryEntries();
    const cleanHistory = rawHistory.map((item) => {
        const clean = validateHistoryRecord(item);
        return clean;
    });
    const payload = {
        app: 'slide-pdf-compressor',
        version: 1,
        exportedAt: new Date().toISOString(),
        preferences: {
            theme: preferences.theme,
            targetMb: preferences.targetMb,
            mode: preferences.mode,
            outputFormat: preferences.outputFormat,
        },
        history: cleanHistory,
    };
    return JSON.stringify(payload, null, 2);
}

export async function importDatabaseFromJson(rawText) {
    const validated = validateBackupJsonText(rawText);
    if (validated.preferences) {
        memoryPreferences = { ...validated.preferences };
    }
    const sorted = [...validated.history]
        .sort((a, b) => b.createdAt - a.createdAt)
        .slice(0, MAX_HISTORY_ITEMS);
    memoryHistory = sorted;

    const db = getSlideCompressorDB();
    if (db) {
        try {
            const prefTable = db.table('preferences');
            const histTable = db.table('history');

            await db.transaction('rw', prefTable, histTable, async () => {
                if (validated.preferences) {
                    await prefTable.put(validated.preferences);
                }
                await histTable.clear();
                if (sorted.length > 0) {
                    await histTable.bulkAdd(sorted);
                }
            });
        } catch (_err) {
            // Keep imported state in memory if IndexedDB throws in private browsing
        }
    }

    return {
        preferences: validated.preferences || (await loadPreferences()),
        importedCount: validated.history.length,
    };
}

export async function clearLocalDatabase() {
    memoryPreferences = { ...DEFAULT_PREFERENCES };
    memoryHistory = [];
    const db = getSlideCompressorDB();
    if (!db) {
        return;
    }
    try {
        const prefTable = db.table('preferences');
        const histTable = db.table('history');
        await db.transaction('rw', prefTable, histTable, async () => {
            await prefTable.clear();
            await histTable.clear();
        });
    } catch (_err) {
        // Ignore IndexedDB error after clearing in-memory store
    }
}
