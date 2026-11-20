/**
 * 100% Standalone Client-Side Compression & Extraction Engine (PDF & PPTX).
 * Runs entirely in the browser offline via vendored pdf.js, pdf-lib, and JSZip.
 */

import * as pdfjsLib from './vendor/pdf.min.mjs';

pdfjsLib.GlobalWorkerOptions.workerSrc = new URL(
    './vendor/pdf.worker.min.mjs',
    import.meta.url
).href;

const STANDARD_FONTS_URL = new URL(
    './vendor/standard_fonts/',
    import.meta.url
).href;

export const MAX_FILE_BYTES = 500 * 1024 * 1024; // 500 MB
export const MAX_ZIP_ENTRIES = 5000;
export const MAX_ZIP_UNCOMPRESSED_BYTES = 1_500_000_000; // 1.5 GB
export const MAX_CANVAS_PIXELS = 64_000_000; // 64 MP
const MAX_SINGLE_XML_BYTES = 50 * 1024 * 1024; // 50 MB

const IMAGE_EXTENSIONS = new Set([
    '.jpg',
    '.jpeg',
    '.png',
    '.bmp',
    '.gif',
    '.webp',
    '.tif',
    '.tiff',
]);

const STORED_ZIP_EXTENSIONS = new Set([
    '.jpg',
    '.jpeg',
    '.png',
    '.webp',
    '.mp4',
    '.m4a',
    '.mov',
    '.avi',
    '.wmv',
    '.mp3',
    '.aac',
]);

const MIME_BY_EXT = {
    '.jpg': 'image/jpeg',
    '.jpeg': 'image/jpeg',
    '.png': 'image/png',
    '.webp': 'image/webp',
    '.gif': 'image/gif',
    '.bmp': 'image/bmp',
    '.tif': 'image/tiff',
    '.tiff': 'image/tiff',
};

function getJSZip() {
    const ctor = globalThis.JSZip;
    if (!ctor) {
        throw new Error('Thư viện JSZip chưa sẵn sàng.');
    }
    return ctor;
}

function getPDFLib() {
    const lib = globalThis.PDFLib;
    if (!lib) {
        throw new Error('Thư viện PDFLib chưa sẵn sàng.');
    }
    return lib;
}

export function createAbortError() {
    const err = new Error('Đã huỷ tác vụ theo yêu cầu người dùng.');
    err.name = 'AbortError';
    return err;
}

export function throwIfAborted(signal) {
    if (signal && signal.aborted) {
        throw createAbortError();
    }
}

function notifyProgress(cb, current, total, message) {
    if (typeof cb === 'function') {
        try {
            cb(current, Math.max(total, 1), message || '');
        } catch (_err) {
            // Ignore UI callback errors
        }
    }
}

function getExt(name) {
    const idx = name.lastIndexOf('.');
    return idx >= 0 ? name.slice(idx).toLowerCase() : '';
}

function getStem(name) {
    const base = name.split(/[\\/]/).pop() || 'document';
    const idx = base.lastIndexOf('.');
    return idx > 0 ? base.slice(0, idx) : base;
}

function posixNormalize(pathStr) {
    const parts = pathStr.replace(/\\/g, '/').split('/');
    const stack = [];
    for (const part of parts) {
        if (!part || part === '.') continue;
        if (part === '..') {
            stack.pop();
        } else {
            stack.push(part);
        }
    }
    return stack.join('/');
}

function posixDirname(pathStr) {
    const norm = posixNormalize(pathStr);
    const idx = norm.lastIndexOf('/');
    return idx >= 0 ? norm.slice(0, idx) : '';
}

function posixRelative(fromDir, toPath) {
    const fromParts = posixNormalize(fromDir).split('/').filter(Boolean);
    const toParts = posixNormalize(toPath).split('/').filter(Boolean);
    let common = 0;
    while (
        common < fromParts.length &&
        common < toParts.length &&
        fromParts[common] === toParts[common]
    ) {
        common++;
    }
    const up = new Array(fromParts.length - common).fill('..');
    return [...up, ...toParts.slice(common)].join('/');
}

function hasZipMagic(bytes) {
    return (
        bytes.byteLength >= 4 &&
        bytes[0] === 0x50 &&
        bytes[1] === 0x4b &&
        (bytes[2] === 0x03 || bytes[2] === 0x05 || bytes[2] === 0x07)
    );
}

function hasPdfMagic(bytes) {
    const limit = Math.min(bytes.byteLength, 1024);
    for (let i = 0; i <= limit - 4; i++) {
        if (
            bytes[i] === 0x25 && // %
            bytes[i + 1] === 0x50 && // P
            bytes[i + 2] === 0x44 && // D
            bytes[i + 3] === 0x46 // F
        ) {
            return true;
        }
    }
    return false;
}

export function validateInputFileSafety(file) {
    if (!file) {
        throw new Error('Vui lòng chọn tệp để xử lý.');
    }
    if (file.size === 0) {
        throw new Error('Tệp tải lên trống (0 byte).');
    }
    if (file.size > MAX_FILE_BYTES) {
        throw new Error('Dung lượng tệp vượt quá giới hạn an toàn tối đa 500 MB.');
    }
    const ext = getExt(file.name);
    if (!['.pdf', '.pptx', '.ppt'].includes(ext)) {
        throw new Error(`Định dạng "${ext || 'không rõ'}" không được hỗ trợ. Chỉ hỗ trợ .pdf, .pptx, .ppt.`);
    }
    return ext;
}

export function validateZipSafety(zip) {
    const names = Object.keys(zip.files);
    if (names.length > MAX_ZIP_ENTRIES) {
        throw new Error(
            `Tệp nén chứa quá nhiều mục (${names.length} > ${MAX_ZIP_ENTRIES}), từ chối xử lý để chống ZIP bomb.`
        );
    }

    let totalUncompressed = 0;
    for (const name of names) {
        if (
            name.includes('\x00') ||
            name.startsWith('/') ||
            name.startsWith('\\') ||
            name.includes('\\')
        ) {
            throw new Error('Phát hiện đường dẫn tệp không hợp lệ bên trong gói ZIP/PPTX.');
        }
        const segments = name.split('/');
        if (segments.includes('..') || (segments[0] && segments[0].includes(':'))) {
            throw new Error('Phát hiện đường dẫn leo thang (path traversal) bên trong gói ZIP/PPTX.');
        }
        const entry = zip.files[name];
        const uncomp = Number(entry?._data?.uncompressedSize ?? 0);
        if (uncomp < 0) {
            throw new Error('Kích thước giải nén bên trong gói ZIP/PPTX không hợp lệ.');
        }
        totalUncompressed += uncomp;
        if (totalUncompressed > MAX_ZIP_UNCOMPRESSED_BYTES) {
            throw new Error(
                'Tổng dung lượng giải nén vượt ngưỡng an toàn (1.5 GB), từ chối xử lý để chống ZIP bomb.'
            );
        }
    }
}

export function parseSafeXmlString(xmlText) {
    if (typeof xmlText !== 'string') {
        throw new Error('Nội dung XML không hợp lệ.');
    }
    const upper = xmlText.toUpperCase();
    if (upper.includes('<!DOCTYPE') || upper.includes('<!ENTITY')) {
        throw new Error('Phát hiện khai báo XML không an toàn (<!DOCTYPE / <!ENTITY), từ chối xử lý.');
    }
    const parser = new DOMParser();
    const xmlDoc = parser.parseFromString(xmlText, 'application/xml');
    if (
        xmlDoc.getElementsByTagName('parsererror').length > 0 ||
        xmlDoc.querySelector?.('parsererror')
    ) {
        throw new Error('Cấu trúc XML bên trong tệp PowerPoint bị lỗi (parsererror).');
    }
    return xmlDoc;
}

async function readSafeXmlFromZip(zip, entryName) {
    const entry = zip.file(entryName);
    if (!entry) return null;
    const uncomp = Number(entry?._data?.uncompressedSize ?? 0);
    if (uncomp > MAX_SINGLE_XML_BYTES) {
        throw new Error('Tệp XML bên trong slide vượt quá giới hạn an toàn 50 MB.');
    }
    const xmlText = await entry.async('string');
    return parseSafeXmlString(xmlText);
}

async function validateAllZipXmlDtdSafety(zip, signal) {
    for (const [name, entry] of Object.entries(zip.files)) {
        if (entry.dir) continue;
        const lower = name.toLowerCase();
        if (lower.endsWith('.xml') || lower.endsWith('.rels')) {
            throwIfAborted(signal);
            const uncomp = Number(entry?._data?.uncompressedSize ?? 0);
            if (uncomp > MAX_SINGLE_XML_BYTES) {
                throw new Error('Tệp XML cấu hình bên trong slide vượt ngưỡng an toàn 50 MB.');
            }
            const rawStr = await entry.async('string');
            const upper = rawStr.toUpperCase();
            if (upper.includes('<!DOCTYPE') || upper.includes('<!ENTITY')) {
                throw new Error('Phát hiện khai báo XML không an toàn (<!DOCTYPE / <!ENTITY), từ chối xử lý.');
            }
        }
    }
}

function clampCanvasDimensions(width, height, maxPixels = MAX_CANVAS_PIXELS) {
    let w = Math.max(1, Math.round(width));
    let h = Math.max(1, Math.round(height));
    if (w * h > maxPixels) {
        const ratio = Math.sqrt(maxPixels / (w * h));
        w = Math.max(1, Math.floor(w * ratio));
        h = Math.max(1, Math.floor(h * ratio));
    }
    return { width: w, height: h };
}

function releaseCanvas(canvas) {
    if (canvas) {
        canvas.width = 0;
        canvas.height = 0;
    }
}

async function canvasToBlob(canvas, mimeType = 'image/jpeg', quality = 0.82) {
    return await new Promise((resolve, reject) => {
        canvas.toBlob(
            (blob) => {
                if (!blob) {
                    reject(new Error('Không thể xuất dữ liệu ảnh từ Canvas.'));
                } else {
                    resolve(blob);
                }
            },
            mimeType,
            quality
        );
    });
}

async function canvasToBytes(canvas, mimeType = 'image/jpeg', quality = 0.82) {
    const blob = await canvasToBlob(canvas, mimeType, quality);
    const ab = await blob.arrayBuffer();
    return new Uint8Array(ab);
}

function bytesToBase64DataUri(bytes, mime = 'image/jpeg') {
    let binary = '';
    const chunk = 0x8000;
    for (let i = 0; i < bytes.length; i += chunk) {
        binary += String.fromCharCode(...bytes.subarray(i, i + chunk));
    }
    return `data:${mime};base64,${btoa(binary)}`;
}

async function canvasToJpegDataUri(canvas, quality = 0.80) {
    const jpgBytes = await canvasToBytes(canvas, 'image/jpeg', quality);
    return bytesToBase64DataUri(jpgBytes, 'image/jpeg');
}

async function decodeImageBytes(bytes, mimeType = 'image/png') {
    const blob = new Blob([bytes], { type: mimeType });
    if (typeof createImageBitmap === 'function') {
        const bmp = await createImageBitmap(blob);
        if (bmp.width * bmp.height > MAX_CANVAS_PIXELS) {
            bmp.close();
            throw new Error('Kích thước điểm ảnh vượt ngưỡng an toàn 64 MP.');
        }
        return {
            source: bmp,
            width: bmp.width,
            height: bmp.height,
            close: () => bmp.close(),
        };
    }

    const url = URL.createObjectURL(blob);
    try {
        const img = await new Promise((resolve, reject) => {
            const el = new Image();
            el.onload = () => resolve(el);
            el.onerror = () => reject(new Error('Không thể giải mã hình ảnh.'));
            el.src = url;
        });
        const w = img.naturalWidth || img.width;
        const h = img.naturalHeight || img.height;
        if (w * h > MAX_CANVAS_PIXELS) {
            throw new Error('Kích thước điểm ảnh vượt ngưỡng an toàn 64 MP.');
        }
        return {
            source: img,
            width: w,
            height: h,
            close: () => {},
        };
    } finally {
        URL.revokeObjectURL(url);
    }
}

async function sha256Hex(bytes) {
    if (globalThis.crypto && globalThis.crypto.subtle) {
        const digest = await globalThis.crypto.subtle.digest('SHA-256', bytes);
        const arr = new Uint8Array(digest);
        let hex = '';
        for (let i = 0; i < arr.length; i++) {
            hex += arr[i].toString(16).padStart(2, '0');
        }
        return hex;
    }
    // Pure-JS fallback hash (double 32-bit FNV-1a + byteLength) for non-secure local contexts
    let h1 = 0x811c9dc5;
    let h2 = 0x01000193;
    for (let i = 0; i < bytes.length; i++) {
        h1 ^= bytes[i];
        h1 = Math.imul(h1, 0x01000193) >>> 0;
        h2 ^= bytes[i] ^ (i & 0xff);
        h2 = Math.imul(h2, 0x85ebca6b) >>> 0;
    }
    return `${h1.toString(16).padStart(8, '0')}${h2.toString(16).padStart(8, '0')}_${bytes.length}`;
}

async function openPdfJsDoc(pdfBytes) {
    const loadingTask = pdfjsLib.getDocument({
        data: pdfBytes.slice(),
        standardFontDataUrl: STANDARD_FONTS_URL,
        useSystemFonts: true,
        isEvalSupported: false,
        maxImageSize: MAX_CANVAS_PIXELS,
    });
    return await loadingTask.promise;
}

function pageBaseFactor(wPt, hPt) {
    const maxDim = Math.max(wPt, hPt, 1.0);
    let factor = 1.0;
    if (maxDim < 1000.0) {
        factor = Math.min(1.5, 1440.0 / maxDim);
    } else if (maxDim > 2000.0) {
        factor = 1600.0 / maxDim;
    }
    const estPixels = (wPt * factor) * (hPt * factor);
    if (estPixels > MAX_CANVAS_PIXELS) {
        factor *= Math.sqrt(MAX_CANVAS_PIXELS / estPixels);
    }
    return factor;
}

async function renderPdfPageToCanvas(pdfDoc, pageNumber, renderScale, targetCanvas) {
    const page = await pdfDoc.getPage(pageNumber);
    try {
        const baseVp = page.getViewport({ scale: 1.0 });
        const rawVp = page.getViewport({ scale: renderScale });
        const safeDims = clampCanvasDimensions(rawVp.width, rawVp.height);
        const effScale = renderScale * Math.min(
            safeDims.width / Math.max(rawVp.width, 1),
            safeDims.height / Math.max(rawVp.height, 1)
        );
        const viewport = page.getViewport({ scale: effScale });
        const canvas = targetCanvas || document.createElement('canvas');
        canvas.width = Math.max(1, Math.round(viewport.width));
        canvas.height = Math.max(1, Math.round(viewport.height));
        const ctx = canvas.getContext('2d', { alpha: false });
        ctx.fillStyle = '#ffffff';
        ctx.fillRect(0, 0, canvas.width, canvas.height);
        await page.render({ canvasContext: ctx, viewport }).promise;
        return {
            canvas,
            wPt: baseVp.width,
            hPt: baseVp.height,
            pxW: canvas.width,
            pxH: canvas.height,
        };
    } finally {
        page.cleanup();
    }
}

async function countUniquePdfEmbeddedImages(pdfBytes) {
    try {
        const { PDFDocument, PDFRawStream, PDFName } = getPDFLib();
        const doc = await PDFDocument.load(pdfBytes, {
            ignoreEncryption: true,
            updateMetadata: false,
        });
        const subtypeKey = PDFName.of('Subtype');
        const smaskKey = PDFName.of('SMask');
        const smaskRefs = new Set();

        for (const [, obj] of doc.context.enumerateIndirectObjects()) {
            if (obj instanceof PDFRawStream) {
                const sm = obj.dict.get(smaskKey);
                if (sm && typeof sm.toString === 'function') {
                    smaskRefs.add(sm.toString());
                }
            }
        }

        let count = 0;
        for (const [ref, obj] of doc.context.enumerateIndirectObjects()) {
            if (obj instanceof PDFRawStream) {
                if (smaskRefs.has(ref.toString())) continue;
                const st = obj.dict.get(subtypeKey);
                if (st && st.toString() === '/Image') {
                    count++;
                }
            }
        }
        return count;
    } catch (_err) {
        return 0;
    }
}

// ---------------------------------------------------------------------------
// PPTX XML Inspection & Slide Canvas Renderer
// ---------------------------------------------------------------------------

function childByLocalName(parent, localName) {
    if (!parent || !parent.children) return null;
    for (let i = 0; i < parent.children.length; i++) {
        if (parent.children[i].localName === localName) {
            return parent.children[i];
        }
    }
    return null;
}

function childrenByLocalName(parent, localName) {
    const out = [];
    if (!parent || !parent.children) return out;
    for (let i = 0; i < parent.children.length; i++) {
        if (parent.children[i].localName === localName) {
            out.push(parent.children[i]);
        }
    }
    return out;
}

function descendantsByLocalName(root, localName) {
    const out = [];
    if (!root) return out;
    const all = root.getElementsByTagName('*');
    for (let i = 0; i < all.length; i++) {
        if (all[i].localName === localName) {
            out.push(all[i]);
        }
    }
    return out;
}

function parseHexColorNode(solidFillNode) {
    if (!solidFillNode) return null;
    const srgb = descendantsByLocalName(solidFillNode, 'srgbClr')[0];
    if (srgb) {
        const val = (srgb.getAttribute('val') || '').trim();
        if (/^[0-9A-Fa-f]{6}$/.test(val)) {
            return {
                hex: `#${val}`,
                r: parseInt(val.slice(0, 2), 16),
                g: parseInt(val.slice(2, 4), 16),
                b: parseInt(val.slice(4, 6), 16),
            };
        }
    }
    return null;
}

function getEmbedAttr(el) {
    if (!el || !el.attributes) return null;
    for (let i = 0; i < el.attributes.length; i++) {
        const attr = el.attributes[i];
        if (attr.localName === 'embed' || attr.name === 'r:embed') {
            return attr.value;
        }
    }
    return null;
}

async function parseRelsMap(zip, partZipPath) {
    const dir = posixDirname(partZipPath);
    const base = partZipPath.split('/').pop();
    const relsPath = dir ? `${dir}/_rels/${base}.rels` : `_rels/${base}.rels`;
    const map = new Map();
    const xmlDoc = await readSafeXmlFromZip(zip, relsPath);
    if (!xmlDoc) return map;

    const rels = descendantsByLocalName(xmlDoc, 'Relationship');
    for (const rel of rels) {
        const id = rel.getAttribute('Id');
        const target = rel.getAttribute('Target');
        const mode = (rel.getAttribute('TargetMode') || '').toLowerCase();
        if (!id || !target || mode === 'external') continue;
        const resolved = target.startsWith('/')
            ? posixNormalize(target.slice(1))
            : posixNormalize(dir ? `${dir}/${target}` : target);
        map.set(id, resolved);
    }
    return map;
}

async function getOrderedPptxSlidePaths(zip) {
    let emuW = 12192000;
    let emuH = 6858000;
    const orderedSlides = [];

    const presDoc = await readSafeXmlFromZip(zip, 'ppt/presentation.xml');
    if (presDoc) {
        const sldSz = descendantsByLocalName(presDoc, 'sldSz')[0];
        if (sldSz) {
            const cx = Number(sldSz.getAttribute('cx'));
            const cy = Number(sldSz.getAttribute('cy'));
            if (Number.isFinite(cx) && cx > 0) emuW = cx;
            if (Number.isFinite(cy) && cy > 0) emuH = cy;
        }
        const presRels = await parseRelsMap(zip, 'ppt/presentation.xml');
        const sldIds = descendantsByLocalName(presDoc, 'sldId');
        for (const sldId of sldIds) {
            let rId = null;
            for (let i = 0; i < sldId.attributes.length; i++) {
                const a = sldId.attributes[i];
                if (a.localName === 'id' && a.name !== 'id') {
                    rId = a.value;
                    break;
                }
            }
            if (rId && presRels.has(rId)) {
                const target = presRels.get(rId);
                if (zip.file(target)) {
                    orderedSlides.push(target);
                }
            }
        }
    }

    if (orderedSlides.length === 0) {
        const fallback = Object.keys(zip.files)
            .filter((n) => /^ppt\/slides\/slide\d+\.xml$/i.test(n))
            .sort((a, b) => {
                const na = parseInt((a.match(/slide(\d+)\.xml/i) || [])[1] || '0', 10);
                const nb = parseInt((b.match(/slide(\d+)\.xml/i) || [])[1] || '0', 10);
                return na - nb;
            });
        orderedSlides.push(...fallback);
    }

    return { slidePaths: orderedSlides, emuW, emuH };
}

function wrapTextLines(ctx, text, maxWidth) {
    if (!text) return [];
    const words = text.split(/\s+/).filter(Boolean);
    if (words.length === 0) return [];
    const lines = [];
    let current = '';
    for (const word of words) {
        const cand = current ? `${current} ${word}` : word;
        if (ctx.measureText(cand).width <= maxWidth || !current) {
            current = cand;
        } else {
            lines.push(current);
            current = word;
        }
    }
    if (current) lines.push(current);
    return lines;
}

function readShapeXfrm(spPrOrGrpSpPr) {
    if (!spPrOrGrpSpPr) return null;
    const xfrm = childByLocalName(spPrOrGrpSpPr, 'xfrm');
    if (!xfrm) return null;
    const off = childByLocalName(xfrm, 'off');
    const ext = childByLocalName(xfrm, 'ext');
    const chOff = childByLocalName(xfrm, 'chOff');
    const chExt = childByLocalName(xfrm, 'chExt');
    return {
        x: Number(off?.getAttribute('x') ?? 0),
        y: Number(off?.getAttribute('y') ?? 0),
        cx: Number(ext?.getAttribute('cx') ?? 0),
        cy: Number(ext?.getAttribute('cy') ?? 0),
        chX: Number(chOff?.getAttribute('x') ?? 0),
        chY: Number(chOff?.getAttribute('y') ?? 0),
        chCx: Number(chExt?.getAttribute('cx') ?? 0),
        chCy: Number(chExt?.getAttribute('cy') ?? 0),
    };
}

async function drawMediaFromZipOnCanvas(zip, mediaPath, ctx, leftPx, topPx, wPx, hPx) {
    const entry = zip.file(mediaPath);
    if (!entry) return;
    const ext = getExt(mediaPath);
    if (!IMAGE_EXTENSIONS.has(ext)) return;
    try {
        const bytes = await entry.async('uint8array');
        const decoded = await decodeImageBytes(bytes, MIME_BY_EXT[ext] || 'image/png');
        try {
            ctx.drawImage(decoded.source, leftPx, topPx, Math.max(1, wPx), Math.max(1, hPx));
        } finally {
            decoded.close();
        }
    } catch (_err) {
        // Skip unsupported or corrupt embedded decorative shapes gracefully
    }
}

async function renderPptxNodeRecursive(
    node,
    zip,
    relsMap,
    ctx,
    emuW,
    emuH,
    canvasW,
    canvasH,
    slideBgColor,
    transform = { offX: 0, offY: 0, scaleX: 1, scaleY: 1, chOffX: 0, chOffY: 0 }
) {
    const tag = node.localName;

    if (tag === 'grpSp') {
        const grpSpPr = childByLocalName(node, 'grpSpPr');
        const xf = readShapeXfrm(grpSpPr);
        let nextTf = transform;
        if (xf) {
            const sLeft = transform.offX + (xf.x - transform.chOffX) * transform.scaleX;
            const sTop = transform.offY + (xf.y - transform.chOffY) * transform.scaleY;
            const sW = xf.cx * transform.scaleX;
            const sH = xf.cy * transform.scaleY;
            if (xf.chCx > 0 && xf.chCy > 0 && sW > 0 && sH > 0) {
                nextTf = {
                    offX: sLeft,
                    offY: sTop,
                    scaleX: sW / xf.chCx,
                    scaleY: sH / xf.chCy,
                    chOffX: xf.chX,
                    chOffY: xf.chY,
                };
            } else {
                nextTf = {
                    offX: sLeft,
                    offY: sTop,
                    scaleX: transform.scaleX,
                    scaleY: transform.scaleY,
                    chOffX: 0,
                    chOffY: 0,
                };
            }
        }
        for (let i = 0; i < node.children.length; i++) {
            const ch = node.children[i];
            if (['sp', 'pic', 'grpSp'].includes(ch.localName)) {
                await renderPptxNodeRecursive(
                    ch,
                    zip,
                    relsMap,
                    ctx,
                    emuW,
                    emuH,
                    canvasW,
                    canvasH,
                    slideBgColor,
                    nextTf
                );
            }
        }
        return;
    }

    if (tag !== 'sp' && tag !== 'pic') return;

    const spPr = childByLocalName(node, 'spPr');
    const xf = readShapeXfrm(spPr);
    if (!xf || (xf.cx <= 0 && xf.cy <= 0)) return;

    const sLeft = transform.offX + (xf.x - transform.chOffX) * transform.scaleX;
    const sTop = transform.offY + (xf.y - transform.chOffY) * transform.scaleY;
    const sW = xf.cx * transform.scaleX;
    const sH = xf.cy * transform.scaleY;

    const leftPx = Math.round((sLeft / emuW) * canvasW);
    const topPx = Math.round((sTop / emuH) * canvasH);
    const wPx = Math.max(1, Math.round((sW / emuW) * canvasW));
    const hPx = Math.max(1, Math.round((sH / emuH) * canvasH));

    let shapeFill = null;
    if (spPr) {
        const solid = childByLocalName(spPr, 'solidFill');
        shapeFill = parseHexColorNode(solid);
        if (shapeFill) {
            ctx.fillStyle = shapeFill.hex;
            ctx.fillRect(leftPx, topPx, wPx, hPx);
        }
    }

    const blipFill = childByLocalName(node, 'blipFill') || (spPr ? childByLocalName(spPr, 'blipFill') : null);
    if (blipFill) {
        const blip = descendantsByLocalName(blipFill, 'blip')[0];
        const rId = getEmbedAttr(blip);
        if (rId && relsMap.has(rId)) {
            await drawMediaFromZipOnCanvas(zip, relsMap.get(rId), ctx, leftPx, topPx, wPx, hPx);
        }
    }

    const txBody = childByLocalName(node, 'txBody');
    if (txBody) {
        const slideWPt = Math.max(emuW / 12700.0, 1.0);
        const pxPerPt = canvasW / slideWPt;
        const effBg = shapeFill || slideBgColor;
        const bgLum = 0.299 * effBg.r + 0.587 * effBg.g + 0.114 * effBg.b;
        const defaultFg = bgLum < 110 ? '#f8fafc' : '#1e293b';

        const padX = Math.max(4, Math.round(canvasW * 0.006));
        const padY = Math.max(4, Math.round(canvasH * 0.006));
        const maxTextW = Math.max(36, wPx - 2 * padX);
        let currY = Math.max(0, topPx + padY);
        const maxY = Math.min(canvasH - 4, Math.max(topPx + hPx + Math.round(canvasH * 0.08), currY + 40));

        const paragraphs = childrenByLocalName(txBody, 'p');
        for (const p of paragraphs) {
            const pPr = childByLocalName(p, 'pPr');
            const algn = (pPr?.getAttribute('algn') || 'l').toLowerCase();

            let ptSize = 18;
            let isBold = false;
            let fgHex = defaultFg;

            const defRPr = pPr ? childByLocalName(pPr, 'defRPr') : null;
            if (defRPr) {
                const szAttr = Number(defRPr.getAttribute('sz'));
                if (Number.isFinite(szAttr) && szAttr > 0) ptSize = szAttr / 100;
                if (defRPr.getAttribute('b') === '1') isBold = true;
                const c = parseHexColorNode(childByLocalName(defRPr, 'solidFill'));
                if (c) fgHex = c.hex;
            }

            let rawText = '';
            for (let i = 0; i < p.children.length; i++) {
                const ch = p.children[i];
                if (ch.localName === 'r' || ch.localName === 'fld') {
                    const rPr = childByLocalName(ch, 'rPr');
                    if (rPr) {
                        const szAttr = Number(rPr.getAttribute('sz'));
                        if (ptSize === 18 && Number.isFinite(szAttr) && szAttr > 0) {
                            ptSize = szAttr / 100;
                        }
                        if (rPr.getAttribute('b') === '1') isBold = true;
                        if (fgHex === defaultFg) {
                            const c = parseHexColorNode(childByLocalName(rPr, 'solidFill'));
                            if (c) fgHex = c.hex;
                        }
                    }
                    const tNode = childByLocalName(ch, 't');
                    if (tNode && tNode.textContent) {
                        rawText += tNode.textContent;
                    }
                } else if (ch.localName === 'br') {
                    rawText += '\n';
                }
            }

            if (!rawText.trim()) {
                currY += Math.max(6, Math.round(10 * pxPerPt));
                continue;
            }

            const fontPx = Math.max(11, Math.min(220, Math.round(ptSize * pxPerPt)));
            ctx.font = `${isBold ? 700 : 400} ${fontPx}px "Be Vietnam Pro", sans-serif`;
            ctx.fillStyle = fgHex;
            ctx.textBaseline = 'top';
            const lineStep = Math.max(fontPx + 4, Math.round(fontPx * 1.28));

            for (const rawLine of rawText.split('\n')) {
                const wrapped = wrapTextLines(ctx, rawLine.trim(), maxTextW);
                for (const lineStr of wrapped) {
                    if (currY > maxY) break;
                    let drawX = Math.max(0, leftPx + padX);
                    if (algn === 'ctr') {
                        const lw = ctx.measureText(lineStr).width;
                        drawX = Math.max(0, leftPx + Math.round((wPx - lw) / 2));
                    } else if (algn === 'r') {
                        const lw = ctx.measureText(lineStr).width;
                        drawX = Math.max(0, leftPx + wPx - padX - Math.round(lw));
                    }
                    ctx.fillText(lineStr, drawX, currY);
                    currY += lineStep;
                }
            }
        }
    }
}

async function resolveSlideBackground(zip, slideDoc, slideRels) {
    const checkDocBg = async (doc, rels) => {
        if (!doc) return null;
        const bg = descendantsByLocalName(doc, 'bg')[0];
        if (!bg) return null;
        const solid = descendantsByLocalName(bg, 'solidFill')[0];
        const color = parseHexColorNode(solid);
        if (color) return { type: 'color', color };
        const blipFill = descendantsByLocalName(bg, 'blipFill')[0];
        if (blipFill) {
            const blip = descendantsByLocalName(blipFill, 'blip')[0];
            const rId = getEmbedAttr(blip);
            if (rId && rels.has(rId)) {
                return { type: 'image', mediaPath: rels.get(rId), color: { hex: '#ffffff', r: 255, g: 255, b: 255 } };
            }
        }
        return null;
    };

    const direct = await checkDocBg(slideDoc, slideRels);
    if (direct) return direct;

    for (const target of slideRels.values()) {
        if (target.includes('slideLayouts/')) {
            const layoutDoc = await readSafeXmlFromZip(zip, target);
            const layoutRels = await parseRelsMap(zip, target);
            const layoutBg = await checkDocBg(layoutDoc, layoutRels);
            if (layoutBg) return layoutBg;
            for (const mTarget of layoutRels.values()) {
                if (mTarget.includes('slideMasters/')) {
                    const masterDoc = await readSafeXmlFromZip(zip, mTarget);
                    const masterRels = await parseRelsMap(zip, mTarget);
                    const masterBg = await checkDocBg(masterDoc, masterRels);
                    if (masterBg) return masterBg;
                }
            }
        }
    }

    return { type: 'color', color: { hex: '#ffffff', r: 255, g: 255, b: 255 } };
}

async function renderPptxSlideToCanvas(zip, slidePath, emuW, emuH, targetWidth = 1440, reusableCanvas = null) {
    const aspect = emuW > 0 ? emuH / emuW : 9 / 16;
    const safeDims = clampCanvasDimensions(targetWidth, Math.round(targetWidth * aspect));
    const canvas = reusableCanvas || document.createElement('canvas');
    canvas.width = safeDims.width;
    canvas.height = safeDims.height;

    const ctx = canvas.getContext('2d', { alpha: false });
    const slideDoc = await readSafeXmlFromZip(zip, slidePath);
    const slideRels = await parseRelsMap(zip, slidePath);
    const bgInfo = await resolveSlideBackground(zip, slideDoc, slideRels);

    ctx.fillStyle = bgInfo.color.hex;
    ctx.fillRect(0, 0, canvas.width, canvas.height);
    if (bgInfo.type === 'image' && bgInfo.mediaPath) {
        await drawMediaFromZipOnCanvas(zip, bgInfo.mediaPath, ctx, 0, 0, canvas.width, canvas.height);
    }

    if (slideDoc) {
        const spTree = descendantsByLocalName(slideDoc, 'spTree')[0];
        if (spTree) {
            for (let i = 0; i < spTree.children.length; i++) {
                const child = spTree.children[i];
                if (['sp', 'pic', 'grpSp'].includes(child.localName)) {
                    await renderPptxNodeRecursive(
                        child,
                        zip,
                        slideRels,
                        ctx,
                        emuW,
                        emuH,
                        canvas.width,
                        canvas.height,
                        bgInfo.color
                    );
                }
            }
        }
    }

    return canvas;
}

// ---------------------------------------------------------------------------
// 1. Inspect File (PDF / PPTX / Legacy PPT)
// ---------------------------------------------------------------------------

export async function inspectFileOffline(file, signal) {
    const ext = validateInputFileSafety(file);
    throwIfAborted(signal);

    const ab = await file.arrayBuffer();
    throwIfAborted(signal);
    const rawBytes = new Uint8Array(ab);
    const originalSizeBytes = rawBytes.byteLength;
    const originalSizeMb = Math.round((originalSizeBytes / (1024 * 1024)) * 100) / 100;

    if (ext === '.pdf') {
        if (!hasPdfMagic(rawBytes)) {
            throw new Error('Tệp .pdf không hợp lệ hoặc bị hỏng.');
        }
        const pdfDoc = await openPdfJsDoc(rawBytes);
        const canvas = document.createElement('canvas');
        try {
            const pageCount = pdfDoc.numPages;
            let width = 0;
            let height = 0;
            let previewBase64 = null;
            if (pageCount > 0) {
                throwIfAborted(signal);
                const page1 = await pdfDoc.getPage(1);
                const vp1 = page1.getViewport({ scale: 1.0 });
                width = Math.round(vp1.width * 100) / 100;
                height = Math.round(vp1.height * 100) / 100;
                page1.cleanup();

                const maxPt = Math.max(width, height, 1.0);
                const thumbScale = Math.min(1.0, 1280.0 / maxPt);
                await renderPdfPageToCanvas(pdfDoc, 1, thumbScale, canvas);
                previewBase64 = await canvasToJpegDataUri(canvas, 0.80);
            }
            const mediaCount = await countUniquePdfEmbeddedImages(rawBytes);
            return {
                filename: file.name,
                file_type: 'pdf',
                is_legacy_ppt: false,
                original_size_bytes: originalSizeBytes,
                original_size_mb: originalSizeMb,
                page_count: pageCount,
                slide_count: pageCount,
                media_count: mediaCount,
                width,
                height,
                preview_base64: previewBase64,
            };
        } finally {
            releaseCanvas(canvas);
            await pdfDoc.destroy();
        }
    }

    const isZip = hasZipMagic(rawBytes);
    if (ext === '.pptx' && !isZip) {
        throw new Error('Tệp .pptx bị hỏng hoặc không đúng cấu trúc gói ZIP.');
    }
    if (!isZip) {
        // Legacy binary .ppt (97-2003 OLE2)
        return {
            filename: file.name,
            file_type: 'ppt',
            is_legacy_ppt: true,
            original_size_bytes: originalSizeBytes,
            original_size_mb: originalSizeMb,
            page_count: 0,
            slide_count: 0,
            media_count: 0,
            preview_base64: null,
        };
    }

    const JSZip = getJSZip();
    const zip = await JSZip.loadAsync(rawBytes);
    validateZipSafety(zip);
    await validateAllZipXmlDtdSafety(zip, signal);

    const { slidePaths, emuW, emuH } = await getOrderedPptxSlidePaths(zip);
    const mediaCount = Object.values(zip.files).filter(
        (entry) => !entry.dir && entry.name.startsWith('ppt/media/')
    ).length;

    let previewBase64 = null;
    if (slidePaths.length > 0) {
        const canvas = document.createElement('canvas');
        try {
            await renderPptxSlideToCanvas(zip, slidePaths[0], emuW, emuH, 960, canvas);
            previewBase64 = await canvasToJpegDataUri(canvas, 0.80);
        } catch (_err) {
            previewBase64 = null;
        } finally {
            releaseCanvas(canvas);
        }
    }

    return {
        filename: file.name,
        file_type: 'pptx',
        is_legacy_ppt: false,
        original_size_bytes: originalSizeBytes,
        original_size_mb: originalSizeMb,
        page_count: slidePaths.length,
        slide_count: slidePaths.length,
        media_count: mediaCount,
        width: Math.round((emuW / 12700.0) * 100) / 100,
        height: Math.round((emuH / 12700.0) * 100) / 100,
        preview_base64: previewBase64,
    };
}

// ---------------------------------------------------------------------------
// 2. PDF Target-MB Compression (Flatten & Vector)
// ---------------------------------------------------------------------------

async function buildPdfFromJpegBuffers(pageBuffers) {
    const { PDFDocument } = getPDFLib();
    const outDoc = await PDFDocument.create();
    for (const buf of pageBuffers) {
        const embedded = await outDoc.embedJpg(buf.jpegBytes);
        const page = outDoc.addPage([buf.wPt, buf.hPt]);
        page.drawImage(embedded, {
            x: 0,
            y: 0,
            width: buf.wPt,
            height: buf.hPt,
        });
    }
    return await outDoc.save({ useObjectStreams: true });
}

async function reencodeJpegBuffersStepDown(sourceBuffers, scaleRatio, quality, signal) {
    const out = [];
    const canvas = document.createElement('canvas');
    try {
        for (const item of sourceBuffers) {
            throwIfAborted(signal);
            const decoded = await decodeImageBytes(item.jpegBytes, 'image/jpeg');
            try {
                const newDims = clampCanvasDimensions(
                    Math.max(1, Math.round(item.pxW * scaleRatio)),
                    Math.max(1, Math.round(item.pxH * scaleRatio))
                );
                canvas.width = newDims.width;
                canvas.height = newDims.height;
                const ctx = canvas.getContext('2d', { alpha: false });
                ctx.fillStyle = '#ffffff';
                ctx.fillRect(0, 0, canvas.width, canvas.height);
                ctx.imageSmoothingEnabled = true;
                ctx.imageSmoothingQuality = 'high';
                ctx.drawImage(decoded.source, 0, 0, canvas.width, canvas.height);
                const newJpeg = await canvasToBytes(canvas, 'image/jpeg', quality);
                out.push({
                    jpegBytes: newJpeg,
                    pxW: canvas.width,
                    pxH: canvas.height,
                    wPt: item.wPt,
                    hPt: item.hPt,
                });
            } finally {
                decoded.close();
            }
        }
    } finally {
        releaseCanvas(canvas);
    }
    return out;
}

async function compressPageBuffersToTargetPdf(initialBuffers, chosenScale, chosenQuality, targetBytes, signal) {
    let scaleUsed = chosenScale;
    let qualityUsed = chosenQuality;
    let currentBuffers = initialBuffers;
    let pdfBytes = await buildPdfFromJpegBuffers(currentBuffers);

    const strictLimit = Math.max(1024, Math.floor(targetBytes * 0.985));
    while (pdfBytes.byteLength >= strictLimit) {
        throwIfAborted(signal);
        if (qualityUsed > 0.405) {
            qualityUsed = Math.max(0.40, Math.round((qualityUsed - 0.08) * 100) / 100);
        } else if (scaleUsed > 0.505) {
            scaleUsed = Math.max(0.50, Math.round((scaleUsed - 0.10) * 100) / 100);
            qualityUsed = 0.45;
        } else {
            if (pdfBytes.byteLength > targetBytes && scaleUsed > 0.305) {
                scaleUsed = Math.max(0.30, Math.round(scaleUsed * 0.75 * 100) / 100);
                qualityUsed = Math.max(0.25, Math.round((qualityUsed - 0.10) * 100) / 100);
                const ratio = scaleUsed / chosenScale;
                currentBuffers = await reencodeJpegBuffersStepDown(initialBuffers, ratio, qualityUsed, signal);
                pdfBytes = await buildPdfFromJpegBuffers(currentBuffers);
                continue;
            }
            break;
        }
        const ratio = scaleUsed / chosenScale;
        currentBuffers = await reencodeJpegBuffersStepDown(initialBuffers, ratio, qualityUsed, signal);
        pdfBytes = await buildPdfFromJpegBuffers(currentBuffers);
    }

    return { pdfBytes, scaleUsed, qualityUsed };
}

async function compressPdfFlattenOffline(rawPdfBytes, targetMb, onProgress, signal) {
    const targetBytes = Math.floor(targetMb * 1024 * 1024);
    const budgetBytes = Math.floor(targetBytes * 0.90);

    const candidates = [
        [1.0, 0.85],
        [1.0, 0.78],
        [1.0, 0.72],
        [1.0, 0.65],
        [0.9, 0.75],
        [0.85, 0.70],
        [0.8, 0.65],
        [0.75, 0.62],
        [0.7, 0.58],
        [0.65, 0.55],
        [0.6, 0.50],
        [0.55, 0.45],
        [0.5, 0.40],
    ];

    const pdfDoc = await openPdfJsDoc(rawPdfBytes);
    const workCanvas = document.createElement('canvas');
    const scaleCanvas = document.createElement('canvas');
    const sampleCanvases = [];

    try {
        const totalPages = pdfDoc.numPages;
        if (totalPages === 0) {
            throw new Error('Tệp PDF không có trang nào.');
        }

        const rawSampleNums = [
            1,
            Math.floor(totalPages / 3) + 1,
            Math.floor((2 * totalPages) / 3) + 1,
            totalPages,
        ];
        const sampleNums = [...new Set(rawSampleNums.filter((n) => n >= 1 && n <= totalPages))].sort(
            (a, b) => a - b
        );

        notifyProgress(onProgress, 0, totalPages, 'Đang lấy mẫu 4 trang để tính thông số tối ưu...');

        for (const pNum of sampleNums) {
            throwIfAborted(signal);
            const page = await pdfDoc.getPage(pNum);
            const vp1 = page.getViewport({ scale: 1.0 });
            page.cleanup();
            const bf = pageBaseFactor(vp1.width, vp1.height);
            const sc = document.createElement('canvas');
            await renderPdfPageToCanvas(pdfDoc, pNum, 1.0 * bf, sc);
            sampleCanvases.push(sc);
        }

        let chosenScale = 0.5;
        let chosenQuality = 0.40;

        for (const [candScale, candQuality] of candidates) {
            throwIfAborted(signal);
            let sampleTotalBytes = 0;
            for (const baseCanvas of sampleCanvases) {
                if (Math.abs(candScale - 1.0) < 1e-3) {
                    const jb = await canvasToBytes(baseCanvas, 'image/jpeg', candQuality);
                    sampleTotalBytes += jb.byteLength;
                } else {
                    scaleCanvas.width = Math.max(1, Math.round(baseCanvas.width * candScale));
                    scaleCanvas.height = Math.max(1, Math.round(baseCanvas.height * candScale));
                    const sctx = scaleCanvas.getContext('2d', { alpha: false });
                    sctx.imageSmoothingEnabled = true;
                    sctx.imageSmoothingQuality = 'high';
                    sctx.drawImage(baseCanvas, 0, 0, scaleCanvas.width, scaleCanvas.height);
                    const jb = await canvasToBytes(scaleCanvas, 'image/jpeg', candQuality);
                    sampleTotalBytes += jb.byteLength;
                }
            }
            const avgPage = sampleTotalBytes / sampleCanvases.length;
            const estTotal = Math.floor((avgPage + 1536) * totalPages);
            if (estTotal <= budgetBytes) {
                chosenScale = candScale;
                chosenQuality = candQuality;
                break;
            }
        }

        for (const sc of sampleCanvases) releaseCanvas(sc);
        sampleCanvases.length = 0;
        releaseCanvas(scaleCanvas);

        const initialBuffers = [];
        for (let pNum = 1; pNum <= totalPages; pNum++) {
            throwIfAborted(signal);
            const page = await pdfDoc.getPage(pNum);
            const vp1 = page.getViewport({ scale: 1.0 });
            page.cleanup();

            const renderScale = chosenScale * pageBaseFactor(vp1.width, vp1.height);
            const rendered = await renderPdfPageToCanvas(pdfDoc, pNum, renderScale, workCanvas);
            const jpegBytes = await canvasToBytes(workCanvas, 'image/jpeg', chosenQuality);
            initialBuffers.push({
                jpegBytes,
                pxW: rendered.pxW,
                pxH: rendered.pxH,
                wPt: rendered.wPt,
                hPt: rendered.hPt,
            });
            notifyProgress(
                onProgress,
                pNum,
                totalPages,
                `Đang gộp phẳng & tối ưu trang ${pNum}/${totalPages}...`
            );
        }

        releaseCanvas(workCanvas);
        const packed = await compressPageBuffersToTargetPdf(
            initialBuffers,
            chosenScale,
            chosenQuality,
            targetBytes,
            signal
        );
        return {
            outBytes: packed.pdfBytes,
            pageCount: totalPages,
            scaleUsed: packed.scaleUsed,
            qualityUsed: Math.round(packed.qualityUsed * 100),
        };
    } finally {
        for (const sc of sampleCanvases) releaseCanvas(sc);
        releaseCanvas(workCanvas);
        releaseCanvas(scaleCanvas);
        await pdfDoc.destroy();
    }
}

async function compressPdfVectorOffline(rawPdfBytes, targetMb, onProgress, signal) {
    const { PDFDocument, PDFRawStream, PDFName, PDFNumber, decodePDFRawStream } = getPDFLib();
    const originalSizeBytes = rawPdfBytes.byteLength;
    const targetBytes = Math.floor(targetMb * 1024 * 1024);

    let passes;
    if (originalSizeBytes <= targetBytes) {
        passes = [[1440, 0.80, false]];
    } else {
        const ratio = originalSizeBytes > 0 ? targetBytes / originalSizeBytes : 0.5;
        let initDim = 800;
        let initQ = 0.58;
        if (ratio >= 0.75) {
            initDim = 1280;
            initQ = 0.75;
        } else if (ratio >= 0.40) {
            initDim = 1024;
            initQ = 0.68;
        }
        passes = [
            [initDim, initQ, false],
            [720, 0.50, true],
            [480, 0.40, true],
        ];
    }

    let bestBytes = null;
    let bestQuality = Math.round(passes[0][1] * 100);
    let pageCount = 0;

    const canvas = document.createElement('canvas');
    try {
        for (let passIdx = 0; passIdx < passes.length; passIdx++) {
            throwIfAborted(signal);
            const [maxDim, jpgQuality] = passes[passIdx];
            const pdfDoc = await PDFDocument.load(rawPdfBytes, {
                updateMetadata: false,
            });
            pageCount = pdfDoc.getPageCount();

            const subtypeKey = PDFName.of('Subtype');
            const filterKey = PDFName.of('Filter');
            const colorSpaceKey = PDFName.of('ColorSpace');
            const widthKey = PDFName.of('Width');
            const heightKey = PDFName.of('Height');
            const bpcKey = PDFName.of('BitsPerComponent');
            const smaskKey = PDFName.of('SMask');
            const maskKey = PDFName.of('Mask');
            const decodeParmsKey = PDFName.of('DecodeParms');
            const lengthKey = PDFName.of('Length');

            const imageEntries = [];
            for (const [ref, obj] of pdfDoc.context.enumerateIndirectObjects()) {
                if (obj instanceof PDFRawStream) {
                    const st = obj.dict.get(subtypeKey);
                    if (st && st.toString() === '/Image') {
                        imageEntries.push([ref, obj]);
                    }
                }
            }

            const totalImgs = Math.max(imageEntries.length, 1);
            for (let i = 0; i < imageEntries.length; i++) {
                throwIfAborted(signal);
                const [ref, streamObj] = imageEntries[i];
                const dict = streamObj.dict;

                if (dict.get(smaskKey) || dict.get(maskKey)) {
                    continue;
                }
                const cs = dict.get(colorSpaceKey)?.toString() || '';
                if (cs !== '/DeviceRGB' && cs !== '/DeviceGray') {
                    continue;
                }
                const filterStr = dict.get(filterKey)?.toString() || '';
                const origBytes = streamObj.contents;
                if (!origBytes || origBytes.byteLength < 4096) {
                    continue;
                }

                try {
                    if (filterStr === '/DCTDecode' || filterStr === '[/DCTDecode]') {
                        const decoded = await decodeImageBytes(origBytes, 'image/jpeg');
                        try {
                            const scale = Math.max(decoded.width, decoded.height) > maxDim
                                ? maxDim / Math.max(decoded.width, decoded.height)
                                : 1.0;
                            const targetDims = clampCanvasDimensions(
                                Math.round(decoded.width * scale),
                                Math.round(decoded.height * scale)
                            );
                            canvas.width = targetDims.width;
                            canvas.height = targetDims.height;
                            const ctx = canvas.getContext('2d', { alpha: false });
                            ctx.fillStyle = '#ffffff';
                            ctx.fillRect(0, 0, canvas.width, canvas.height);
                            ctx.imageSmoothingEnabled = true;
                            ctx.imageSmoothingQuality = 'high';
                            ctx.drawImage(decoded.source, 0, 0, canvas.width, canvas.height);
                            const newJpeg = await canvasToBytes(canvas, 'image/jpeg', jpgQuality);
                            if (newJpeg.byteLength < origBytes.byteLength) {
                                dict.set(widthKey, PDFNumber.of(canvas.width));
                                dict.set(heightKey, PDFNumber.of(canvas.height));
                                dict.set(colorSpaceKey, PDFName.of('DeviceRGB'));
                                dict.set(bpcKey, PDFNumber.of(8));
                                dict.set(filterKey, PDFName.of('DCTDecode'));
                                dict.set(lengthKey, PDFNumber.of(newJpeg.byteLength));
                                pdfDoc.context.assign(ref, PDFRawStream.of(dict, newJpeg));
                            }
                        } finally {
                            decoded.close();
                        }
                    } else if (
                        (filterStr === '/FlateDecode' || filterStr === '[/FlateDecode]') &&
                        !dict.get(decodeParmsKey) &&
                        dict.get(bpcKey)?.asNumber?.() === 8
                    ) {
                        const w = dict.get(widthKey)?.asNumber?.() || 0;
                        const h = dict.get(heightKey)?.asNumber?.() || 0;
                        if (w > 0 && h > 0 && w * h <= MAX_CANVAS_PIXELS) {
                            const rawPixels = decodePDFRawStream(streamObj).decode();
                            const channels = cs === '/DeviceRGB' ? 3 : 1;
                            if (rawPixels.byteLength === w * h * channels) {
                                const srcCanvas = document.createElement('canvas');
                                srcCanvas.width = w;
                                srcCanvas.height = h;
                                const sctx = srcCanvas.getContext('2d');
                                const imgData = sctx.createImageData(w, h);
                                const dst = imgData.data;
                                if (channels === 3) {
                                    for (let p = 0, q = 0; p < rawPixels.length; p += 3, q += 4) {
                                        dst[q] = rawPixels[p];
                                        dst[q + 1] = rawPixels[p + 1];
                                        dst[q + 2] = rawPixels[p + 2];
                                        dst[q + 3] = 255;
                                    }
                                } else {
                                    for (let p = 0, q = 0; p < rawPixels.length; p++, q += 4) {
                                        dst[q] = rawPixels[p];
                                        dst[q + 1] = rawPixels[p];
                                        dst[q + 2] = rawPixels[p];
                                        dst[q + 3] = 255;
                                    }
                                }
                                sctx.putImageData(imgData, 0, 0);

                                const scale = Math.max(w, h) > maxDim ? maxDim / Math.max(w, h) : 1.0;
                                const targetDims = clampCanvasDimensions(Math.round(w * scale), Math.round(h * scale));
                                canvas.width = targetDims.width;
                                canvas.height = targetDims.height;
                                const ctx = canvas.getContext('2d', { alpha: false });
                                ctx.imageSmoothingEnabled = true;
                                ctx.imageSmoothingQuality = 'high';
                                ctx.drawImage(srcCanvas, 0, 0, canvas.width, canvas.height);
                                releaseCanvas(srcCanvas);

                                const newJpeg = await canvasToBytes(canvas, 'image/jpeg', jpgQuality);
                                if (newJpeg.byteLength < origBytes.byteLength) {
                                    dict.set(widthKey, PDFNumber.of(canvas.width));
                                    dict.set(heightKey, PDFNumber.of(canvas.height));
                                    dict.set(colorSpaceKey, PDFName.of('DeviceRGB'));
                                    dict.set(bpcKey, PDFNumber.of(8));
                                    dict.set(filterKey, PDFName.of('DCTDecode'));
                                    dict.set(lengthKey, PDFNumber.of(newJpeg.byteLength));
                                    pdfDoc.context.assign(ref, PDFRawStream.of(dict, newJpeg));
                                }
                            }
                        }
                    }
                } catch (_imgErr) {
                    // Leave complex/unsupported XObject untouched to preserve PDF integrity
                }

                notifyProgress(
                    onProgress,
                    i + 1,
                    totalImgs,
                    `Đang tối ưu ảnh nhúng giữ chữ Vector (${i + 1}/${totalImgs})...`
                );
            }

            const candBytes = await pdfDoc.save({ useObjectStreams: true });
            if (!bestBytes || candBytes.byteLength <= bestBytes.byteLength) {
                bestBytes = candBytes;
                bestQuality = Math.round(jpgQuality * 100);
            }
            if (bestBytes.byteLength <= targetBytes && bestBytes.byteLength <= originalSizeBytes) {
                break;
            }
        }
    } finally {
        releaseCanvas(canvas);
    }

    return {
        outBytes: bestBytes || rawPdfBytes,
        pageCount,
        scaleUsed: 1.0,
        qualityUsed: bestQuality,
    };
}

export async function compressPdfOffline(file, options, onProgress, signal) {
    validateInputFileSafety(file);
    throwIfAborted(signal);

    const targetMb = Math.max(0.001, Number(options?.targetMb ?? 10));
    const mode = options?.mode === 'vector' ? 'vector' : 'flatten';

    const rawPdfBytes = new Uint8Array(await file.arrayBuffer());
    const originalSizeBytes = rawPdfBytes.byteLength;
    const originalSizeMb = Math.round((originalSizeBytes / (1024 * 1024)) * 100) / 100;
    const targetBytes = Math.floor(targetMb * 1024 * 1024);

    let beforePreview = options?.beforePreview || null;
    if (!beforePreview) {
        const doc = await openPdfJsDoc(rawPdfBytes);
        const c = document.createElement('canvas');
        try {
            if (doc.numPages > 0) {
                await renderPdfPageToCanvas(doc, 1, 1.0, c);
                beforePreview = await canvasToJpegDataUri(c, 0.82);
            }
        } finally {
            releaseCanvas(c);
            await doc.destroy();
        }
    }

    const comp = mode === 'vector'
        ? await compressPdfVectorOffline(rawPdfBytes, targetMb, onProgress, signal)
        : await compressPdfFlattenOffline(rawPdfBytes, targetMb, onProgress, signal);

    let finalBytes = comp.outBytes;
    let keptOriginal = false;
    if (finalBytes.byteLength >= originalSizeBytes && originalSizeBytes <= targetBytes) {
        finalBytes = rawPdfBytes;
        keptOriginal = true;
    }

    const compressedSizeBytes = finalBytes.byteLength;
    const compressedSizeMb = Math.round((compressedSizeBytes / (1024 * 1024)) * 100) / 100;
    const reductionPercent = originalSizeBytes > 0
        ? Math.round((1 - compressedSizeBytes / originalSizeBytes) * 1000) / 10
        : 0;
    const targetMet = compressedSizeBytes <= targetBytes;

    let afterPreview = beforePreview;
    if (!keptOriginal) {
        const outDoc = await openPdfJsDoc(finalBytes);
        const c = document.createElement('canvas');
        try {
            if (outDoc.numPages > 0) {
                await renderPdfPageToCanvas(outDoc, 1, 1.0, c);
                afterPreview = await canvasToJpegDataUri(c, 0.82);
            }
        } catch (_err) {
            afterPreview = beforePreview;
        } finally {
            releaseCanvas(c);
            await outDoc.destroy();
        }
    }

    const outputFilename = `${getStem(file.name)}_compressed.pdf`;
    const blob = new Blob([finalBytes], { type: 'application/pdf' });

    return {
        status: 'done',
        job_type: 'compress',
        filename: file.name,
        output_filename: outputFilename,
        output_format: 'pdf',
        output_blob: blob,
        mode_used: mode,
        page_count: comp.pageCount,
        target_mb: targetMb,
        target_met: targetMet,
        original_size_bytes: originalSizeBytes,
        compressed_size_bytes: compressedSizeBytes,
        original_size_mb: originalSizeMb,
        compressed_size_mb: compressedSizeMb,
        reduction_percent: reductionPercent,
        before_preview_base64: beforePreview,
        after_preview_base64: afterPreview,
    };
}

// ---------------------------------------------------------------------------
// 3. PPTX Direct Compression & PPTX -> PDF Conversion
// ---------------------------------------------------------------------------

function checkCanvasOpaque(ctx, w, h) {
    const data = ctx.getImageData(0, 0, w, h).data;
    for (let i = 3; i < data.length; i += 4) {
        if (data[i] < 255) {
            return false;
        }
    }
    return true;
}

async function optimizePptxMediaBytes(rawBytes, ext, maxDim, jpgQuality, allowJpegRename, workCanvas) {
    if (!IMAGE_EXTENSIONS.has(ext) || ext === '.gif' || ext === '.tif' || ext === '.tiff') {
        return { bytes: rawBytes, finalExt: ext };
    }

    let decoded;
    try {
        decoded = await decodeImageBytes(rawBytes, MIME_BY_EXT[ext] || 'image/png');
    } catch (_err) {
        return { bytes: rawBytes, finalExt: ext };
    }

    try {
        const origW = decoded.width;
        const origH = decoded.height;
        const isLarge = origW * origH >= 100_000 || rawBytes.byteLength >= 60_000;
        const scale = Math.max(origW, origH) > maxDim ? maxDim / Math.max(origW, origH) : 1.0;
        const dims = clampCanvasDimensions(Math.round(origW * scale), Math.round(origH * scale));

        workCanvas.width = dims.width;
        workCanvas.height = dims.height;
        const ctx = workCanvas.getContext('2d', { willReadFrequently: true });
        ctx.clearRect(0, 0, workCanvas.width, workCanvas.height);
        ctx.imageSmoothingEnabled = true;
        ctx.imageSmoothingQuality = 'high';
        ctx.drawImage(decoded.source, 0, 0, workCanvas.width, workCanvas.height);

        const isOpaque = (ext === '.jpg' || ext === '.jpeg' || ext === '.bmp')
            ? true
            : checkCanvasOpaque(ctx, workCanvas.width, workCanvas.height);

        let bestInplace = rawBytes;
        if (ext === '.jpg' || ext === '.jpeg') {
            const cand = await canvasToBytes(workCanvas, 'image/jpeg', jpgQuality);
            if (cand.byteLength < rawBytes.byteLength) {
                bestInplace = cand;
            }
        } else if (ext === '.webp') {
            const cand = await canvasToBytes(workCanvas, 'image/webp', jpgQuality);
            if (cand.byteLength < rawBytes.byteLength) {
                bestInplace = cand;
            }
        } else if (ext === '.png' && scale < 0.999) {
            const cand = await canvasToBytes(workCanvas, 'image/png');
            if (cand.byteLength < rawBytes.byteLength) {
                bestInplace = cand;
            }
        }

        if (allowJpegRename && (ext === '.png' || ext === '.bmp') && isOpaque && isLarge) {
            // Composite over white and encode as JPEG
            const rgbCanvas = document.createElement('canvas');
            rgbCanvas.width = workCanvas.width;
            rgbCanvas.height = workCanvas.height;
            try {
                const rctx = rgbCanvas.getContext('2d', { alpha: false });
                rctx.fillStyle = '#ffffff';
                rctx.fillRect(0, 0, rgbCanvas.width, rgbCanvas.height);
                rctx.drawImage(workCanvas, 0, 0);
                const jpgBytes = await canvasToBytes(rgbCanvas, 'image/jpeg', jpgQuality);
                if (jpgBytes.byteLength < Math.floor(bestInplace.byteLength * 0.75)) {
                    return { bytes: jpgBytes, finalExt: '.jpg' };
                }
            } finally {
                releaseCanvas(rgbCanvas);
            }
        }

        return { bytes: bestInplace, finalExt: ext };
    } finally {
        decoded.close();
    }
}

async function findHardcodedMediaBasenamesInZip(zip, candidateBasenames, signal) {
    const hardcoded = new Set();
    if (candidateBasenames.size === 0) return hardcoded;

    for (const [fname, entry] of Object.entries(zip.files)) {
        if (entry.dir || !fname.toLowerCase().endsWith('.xml') || fname === '[Content_Types].xml') {
            continue;
        }
        throwIfAborted(signal);
        const xmlText = await entry.async('string');
        if (!xmlText.includes('media') && !xmlText.includes('.')) continue;
        for (const bname of candidateBasenames) {
            if (xmlText.includes(bname)) {
                hardcoded.add(bname);
            }
        }
    }
    return hardcoded;
}

function rewriteRelsXmlDom(rawXmlText, relsZipPath, renameMap) {
    if (renameMap.size === 0) return rawXmlText;
    const xmlDoc = parseSafeXmlString(rawXmlText);
    const relsDir = posixDirname(relsZipPath);
    const sourceDir = relsDir.toLowerCase().endsWith('_rels') ? posixDirname(relsDir) : relsDir;

    let changed = false;
    const rels = descendantsByLocalName(xmlDoc, 'Relationship');
    for (const rel of rels) {
        const mode = (rel.getAttribute('TargetMode') || '').trim().toLowerCase();
        if (mode === 'external') continue;
        const target = rel.getAttribute('Target');
        if (!target) continue;

        const resolved = target.startsWith('/')
            ? posixNormalize(target.slice(1))
            : posixNormalize(sourceDir ? `${sourceDir}/${target}` : target);

        if (renameMap.has(resolved)) {
            const newZipPath = renameMap.get(resolved);
            const newTarget = target.startsWith('/')
                ? `/${newZipPath}`
                : sourceDir
                    ? posixRelative(sourceDir, newZipPath)
                    : newZipPath;
            rel.setAttribute('Target', newTarget);
            changed = true;
        }
    }

    if (!changed) return rawXmlText;
    return new XMLSerializer().serializeToString(xmlDoc);
}

function rewriteContentTypesXmlDom(rawXmlText, renameMap) {
    if (renameMap.size === 0) return rawXmlText;
    const xmlDoc = parseSafeXmlString(rawXmlText);
    const root = xmlDoc.documentElement;
    const ns = root.namespaceURI || 'http://schemas.openxmlformats.org/package/2006/content-types';

    let hasJpgDefault = false;
    const toRemove = [];
    for (let i = 0; i < root.children.length; i++) {
        const ch = root.children[i];
        if (ch.localName === 'Default') {
            if ((ch.getAttribute('Extension') || '').trim().toLowerCase() === 'jpg') {
                hasJpgDefault = true;
            }
        } else if (ch.localName === 'Override') {
            const partName = posixNormalize((ch.getAttribute('PartName') || '').replace(/^\/+/, ''));
            if (renameMap.has(partName)) {
                toRemove.push(ch);
            }
        }
    }

    for (const node of toRemove) {
        root.removeChild(node);
    }

    if (!hasJpgDefault) {
        const defEl = xmlDoc.createElementNS(ns, 'Default');
        defEl.setAttribute('Extension', 'jpg');
        defEl.setAttribute('ContentType', 'image/jpeg');
        root.insertBefore(defEl, root.firstChild);
    }

    return new XMLSerializer().serializeToString(xmlDoc);
}

export async function compressPptxDirectOffline(file, options, onProgress, signal) {
    validateInputFileSafety(file);
    throwIfAborted(signal);

    const targetMb = Math.max(0.001, Number(options?.targetMb ?? 10));
    const targetBytes = Math.floor(targetMb * 1024 * 1024);

    const rawBytes = new Uint8Array(await file.arrayBuffer());
    if (!hasZipMagic(rawBytes)) {
        throw new Error('Định dạng .ppt cũ hoặc tệp hỏng không hỗ trợ nén trực tiếp sang .pptx.');
    }

    const originalSizeBytes = rawBytes.byteLength;
    const originalSizeMb = Math.round((originalSizeBytes / (1024 * 1024)) * 100) / 100;

    const JSZip = getJSZip();
    const zip = await JSZip.loadAsync(rawBytes);
    validateZipSafety(zip);
    await validateAllZipXmlDtdSafety(zip, signal);

    const { slidePaths, emuW, emuH } = await getOrderedPptxSlidePaths(zip);
    let beforePreview = options?.beforePreview || null;
    if (!beforePreview && slidePaths.length > 0) {
        const c = document.createElement('canvas');
        try {
            await renderPptxSlideToCanvas(zip, slidePaths[0], emuW, emuH, 960, c);
            beforePreview = await canvasToJpegDataUri(c, 0.80);
        } finally {
            releaseCanvas(c);
        }
    }

    const ratio = originalSizeBytes > 0 ? targetBytes / originalSizeBytes : 0.5;
    let initialProfile = [900, 0.58];
    if (ratio >= 0.70) initialProfile = [1440, 0.78];
    else if (ratio >= 0.40) initialProfile = [1152, 0.68];

    const profiles = [
        initialProfile,
        [800, 0.52],
        [600, 0.42],
        [480, 0.36],
    ];

    const allEntries = Object.values(zip.files).filter((e) => !e.dir);
    const mediaEntries = allEntries.filter(
        (e) => e.name.startsWith('ppt/media/') && IMAGE_EXTENSIONS.has(getExt(e.name))
    );
    const mediaSet = new Set(mediaEntries.map((e) => e.name));
    const baseNamesLower = new Set(allEntries.map((e) => e.name.toLowerCase()));

    const candidateBasenames = new Set(
        mediaEntries
            .filter((e) => ['.png', '.bmp'].includes(getExt(e.name)))
            .map((e) => e.name.split('/').pop())
    );
    const hardcodedBasenames = await findHardcodedMediaBasenamesInZip(zip, candidateBasenames, signal);

    let bestOutBytes = rawBytes;
    const workCanvas = document.createElement('canvas');

    try {
        for (const [maxDim, jpgQ] of profiles) {
            throwIfAborted(signal);
            const outZip = new JSZip();
            const usedNamesLower = new Set(baseNamesLower);
            const renameMap = new Map();
            const totalMedia = Math.max(mediaEntries.length, 1);

            for (let i = 0; i < mediaEntries.length; i++) {
                throwIfAborted(signal);
                const mEntry = mediaEntries[i];
                const origName = posixNormalize(mEntry.name);
                const origExt = getExt(origName);
                const bname = origName.split('/').pop();
                const mediaRaw = await mEntry.async('uint8array');

                const allowRename = !hardcodedBasenames.has(bname);
                const { bytes: optBytes, finalExt } = await optimizePptxMediaBytes(
                    mediaRaw,
                    origExt,
                    maxDim,
                    jpgQ,
                    allowRename,
                    workCanvas
                );

                let targetName = mEntry.name;
                if (finalExt !== origExt && finalExt === '.jpg') {
                    usedNamesLower.delete(origName.toLowerCase());
                    const dir = posixDirname(origName);
                    const stem = getStem(origName);
                    let candName = `${dir}/${stem}.jpg`;
                    let idx = 1;
                    while (usedNamesLower.has(candName.toLowerCase())) {
                        candName = `${dir}/${stem}_${idx}.jpg`;
                        idx++;
                    }
                    usedNamesLower.add(candName.toLowerCase());
                    renameMap.set(origName, candName);
                    targetName = candName;
                }

                const targetExt = getExt(targetName);
                outZip.file(targetName, optBytes, {
                    compression: STORED_ZIP_EXTENSIONS.has(targetExt) ? 'STORE' : 'DEFLATE',
                });

                notifyProgress(
                    onProgress,
                    i + 1,
                    totalMedia,
                    `Đang nén tài nguyên ảnh trong Slide (${i + 1}/${mediaEntries.length})...`
                );
            }

            for (const entry of allEntries) {
                if (mediaSet.has(entry.name)) continue;
                throwIfAborted(signal);
                const fname = entry.name;
                if (fname === '[Content_Types].xml' && renameMap.size > 0) {
                    const rawXml = await entry.async('string');
                    const updatedXml = rewriteContentTypesXmlDom(rawXml, renameMap);
                    outZip.file(fname, updatedXml, { compression: 'DEFLATE' });
                } else if (fname.toLowerCase().endsWith('.rels') && renameMap.size > 0) {
                    const rawXml = await entry.async('string');
                    const updatedXml = rewriteRelsXmlDom(rawXml, posixNormalize(fname), renameMap);
                    outZip.file(fname, updatedXml, { compression: 'DEFLATE' });
                } else {
                    const rawData = await entry.async('uint8array');
                    const ext = getExt(fname);
                    outZip.file(fname, rawData, {
                        compression: STORED_ZIP_EXTENSIONS.has(ext) ? 'STORE' : 'DEFLATE',
                    });
                }
            }

            const candZipBytes = await outZip.generateAsync({
                type: 'uint8array',
                compression: 'DEFLATE',
                compressionOptions: { level: 6 },
            });
            bestOutBytes = candZipBytes;
            if (candZipBytes.byteLength <= targetBytes) {
                break;
            }
        }
    } finally {
        releaseCanvas(workCanvas);
    }

    let keptOriginal = false;
    if (bestOutBytes.byteLength >= originalSizeBytes && originalSizeBytes <= targetBytes) {
        bestOutBytes = rawBytes;
        keptOriginal = true;
    }

    let afterPreview = beforePreview;
    if (!keptOriginal && slidePaths.length > 0) {
        const c = document.createElement('canvas');
        try {
            const afterZip = await JSZip.loadAsync(bestOutBytes);
            await renderPptxSlideToCanvas(afterZip, slidePaths[0], emuW, emuH, 960, c);
            afterPreview = await canvasToJpegDataUri(c, 0.80);
        } catch (_err) {
            afterPreview = beforePreview;
        } finally {
            releaseCanvas(c);
        }
    }

    const compressedSizeBytes = bestOutBytes.byteLength;
    const compressedSizeMb = Math.round((compressedSizeBytes / (1024 * 1024)) * 100) / 100;
    const reductionPercent = originalSizeBytes > 0
        ? Math.round((1 - compressedSizeBytes / originalSizeBytes) * 1000) / 10
        : 0;

    return {
        status: 'done',
        job_type: 'compress',
        filename: file.name,
        output_filename: `${getStem(file.name)}_compressed.pptx`,
        output_format: 'pptx',
        output_blob: new Blob([bestOutBytes], {
            type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
        }),
        mode_used: 'pptx_direct',
        page_count: slidePaths.length,
        slide_count: slidePaths.length,
        target_mb: targetMb,
        target_met: compressedSizeBytes <= targetBytes,
        original_size_bytes: originalSizeBytes,
        compressed_size_bytes: compressedSizeBytes,
        original_size_mb: originalSizeMb,
        compressed_size_mb: compressedSizeMb,
        reduction_percent: reductionPercent,
        before_preview_base64: beforePreview,
        after_preview_base64: afterPreview,
    };
}

export async function convertPptxToPdfOffline(file, options, onProgress, signal) {
    validateInputFileSafety(file);
    throwIfAborted(signal);

    const targetMb = Math.max(0.001, Number(options?.targetMb ?? 10));
    const targetBytes = Math.floor(targetMb * 1024 * 1024);

    const rawBytes = new Uint8Array(await file.arrayBuffer());
    if (!hasZipMagic(rawBytes)) {
        throw new Error(
            'Định dạng .PPT nhị phân cũ (97–2003) cần mở bằng PowerPoint và lưu sang .PPTX hoặc .PDF để xử lý trực tiếp trên trình duyệt Offline.'
        );
    }

    const originalSizeBytes = rawBytes.byteLength;
    const originalSizeMb = Math.round((originalSizeBytes / (1024 * 1024)) * 100) / 100;

    const JSZip = getJSZip();
    const zip = await JSZip.loadAsync(rawBytes);
    validateZipSafety(zip);
    await validateAllZipXmlDtdSafety(zip, signal);

    const { slidePaths, emuW, emuH } = await getOrderedPptxSlidePaths(zip);
    if (slidePaths.length === 0) {
        throw new Error('Không tìm thấy trang slide nào bên trong tệp .pptx.');
    }

    const wPt = Math.max(72, emuW / 12700.0);
    const hPt = Math.max(72, emuH / 12700.0);

    const initialBuffers = [];
    let beforePreview = options?.beforePreview || null;
    const canvas = document.createElement('canvas');

    try {
        for (let i = 0; i < slidePaths.length; i++) {
            throwIfAborted(signal);
            await renderPptxSlideToCanvas(zip, slidePaths[i], emuW, emuH, 1440, canvas);
            if (i === 0 && !beforePreview) {
                beforePreview = await canvasToJpegDataUri(canvas, 0.82);
            }
            const jpegBytes = await canvasToBytes(canvas, 'image/jpeg', 0.85);
            initialBuffers.push({
                jpegBytes,
                pxW: canvas.width,
                pxH: canvas.height,
                wPt,
                hPt,
            });
            notifyProgress(
                onProgress,
                i + 1,
                slidePaths.length,
                `Đang dựng Slide ${i + 1}/${slidePaths.length} sang PDF...`
            );
        }
    } finally {
        releaseCanvas(canvas);
    }

    const packed = await compressPageBuffersToTargetPdf(
        initialBuffers,
        1.0,
        0.85,
        targetBytes,
        signal
    );

    const finalBytes = packed.pdfBytes;
    const compressedSizeBytes = finalBytes.byteLength;
    const compressedSizeMb = Math.round((compressedSizeBytes / (1024 * 1024)) * 100) / 100;
    const reductionPercent = originalSizeBytes > 0
        ? Math.round((1 - compressedSizeBytes / originalSizeBytes) * 1000) / 10
        : 0;

    let afterPreview = beforePreview;
    const outDoc = await openPdfJsDoc(finalBytes);
    const prevCanvas = document.createElement('canvas');
    try {
        if (outDoc.numPages > 0) {
            await renderPdfPageToCanvas(outDoc, 1, 1.0, prevCanvas);
            afterPreview = await canvasToJpegDataUri(prevCanvas, 0.82);
        }
    } catch (_err) {
        afterPreview = beforePreview;
    } finally {
        releaseCanvas(prevCanvas);
        await outDoc.destroy();
    }

    return {
        status: 'done',
        job_type: 'compress',
        filename: file.name,
        output_filename: `${getStem(file.name)}_compressed.pdf`,
        output_format: 'pdf',
        output_blob: new Blob([finalBytes], { type: 'application/pdf' }),
        mode_used: options?.mode || 'flatten',
        page_count: slidePaths.length,
        slide_count: slidePaths.length,
        target_mb: targetMb,
        target_met: compressedSizeBytes <= targetBytes,
        original_size_bytes: originalSizeBytes,
        compressed_size_bytes: compressedSizeBytes,
        original_size_mb: originalSizeMb,
        compressed_size_mb: compressedSizeMb,
        reduction_percent: reductionPercent,
        before_preview_base64: beforePreview,
        after_preview_base64: afterPreview,
    };
}

// ---------------------------------------------------------------------------
// 4. Resource Extraction (.ZIP of Embedded Images or Full HD Slides)
// ---------------------------------------------------------------------------

async function extractPdfEmbeddedImagesToZip(rawPdfBytes, outZip, onProgress, signal) {
    const { PDFDocument, PDFRawStream, PDFName, decodePDFRawStream } = getPDFLib();
    const pdfDoc = await PDFDocument.load(rawPdfBytes, {
        ignoreEncryption: true,
        updateMetadata: false,
    });

    const subtypeKey = PDFName.of('Subtype');
    const filterKey = PDFName.of('Filter');
    const colorSpaceKey = PDFName.of('ColorSpace');
    const widthKey = PDFName.of('Width');
    const heightKey = PDFName.of('Height');
    const bpcKey = PDFName.of('BitsPerComponent');
    const smaskKey = PDFName.of('SMask');
    const decodeParmsKey = PDFName.of('DecodeParms');

    const smaskRefStrings = new Set();
    const allImageStreams = [];
    for (const [ref, obj] of pdfDoc.context.enumerateIndirectObjects()) {
        if (obj instanceof PDFRawStream) {
            const sm = obj.dict.get(smaskKey);
            if (sm && typeof sm.toString === 'function') {
                smaskRefStrings.add(sm.toString());
            }
            if (obj.dict.get(subtypeKey)?.toString() === '/Image') {
                allImageStreams.push([ref, obj]);
            }
        }
    }

    const seenHashes = new Set();
    let extractedCount = 0;
    let previewBase64 = null;
    const canvas = document.createElement('canvas');

    try {
        for (let i = 0; i < allImageStreams.length; i++) {
            throwIfAborted(signal);
            const [ref, streamObj] = allImageStreams[i];
            if (smaskRefStrings.has(ref.toString())) continue;

            const dict = streamObj.dict;
            const filterStr = dict.get(filterKey)?.toString() || '';
            const cs = dict.get(colorSpaceKey)?.toString() || '';
            let imgBytes = null;
            let imgExt = '.jpg';

            try {
                if (
                    (filterStr === '/DCTDecode' || filterStr === '[/DCTDecode]') &&
                    !dict.get(smaskKey)
                ) {
                    imgBytes = streamObj.contents;
                    imgExt = '.jpg';
                } else if (
                    (filterStr === '/FlateDecode' || filterStr === '[/FlateDecode]') &&
                    !dict.get(decodeParmsKey) &&
                    dict.get(bpcKey)?.asNumber?.() === 8 &&
                    (cs === '/DeviceRGB' || cs === '/DeviceGray')
                ) {
                    const w = dict.get(widthKey)?.asNumber?.() || 0;
                    const h = dict.get(heightKey)?.asNumber?.() || 0;
                    if (w > 0 && h > 0 && w * h <= MAX_CANVAS_PIXELS) {
                        const rawPixels = decodePDFRawStream(streamObj).decode();
                        const channels = cs === '/DeviceRGB' ? 3 : 1;
                        if (rawPixels.byteLength === w * h * channels) {
                            let alphaPixels = null;
                            const smRef = dict.get(smaskKey);
                            if (smRef) {
                                const smObj = pdfDoc.context.lookup(smRef);
                                if (smObj instanceof PDFRawStream) {
                                    const decodedAlpha = decodePDFRawStream(smObj).decode();
                                    if (decodedAlpha.byteLength === w * h) {
                                        alphaPixels = decodedAlpha;
                                    }
                                }
                            }
                            canvas.width = w;
                            canvas.height = h;
                            const ctx = canvas.getContext('2d');
                            const imgData = ctx.createImageData(w, h);
                            const dst = imgData.data;
                            if (channels === 3) {
                                for (let p = 0, q = 0, a = 0; p < rawPixels.length; p += 3, q += 4, a++) {
                                    dst[q] = rawPixels[p];
                                    dst[q + 1] = rawPixels[p + 1];
                                    dst[q + 2] = rawPixels[p + 2];
                                    dst[q + 3] = alphaPixels ? alphaPixels[a] : 255;
                                }
                            } else {
                                for (let p = 0, q = 0; p < rawPixels.length; p++, q += 4) {
                                    dst[q] = rawPixels[p];
                                    dst[q + 1] = rawPixels[p];
                                    dst[q + 2] = rawPixels[p];
                                    dst[q + 3] = alphaPixels ? alphaPixels[p] : 255;
                                }
                            }
                            ctx.putImageData(imgData, 0, 0);
                            if (alphaPixels) {
                                imgBytes = await canvasToBytes(canvas, 'image/png');
                                imgExt = '.png';
                            } else {
                                imgBytes = await canvasToBytes(canvas, 'image/jpeg', 0.92);
                                imgExt = '.jpg';
                            }
                        }
                    }
                }
            } catch (_err) {
                imgBytes = null;
            }

            if (imgBytes && imgBytes.byteLength > 0) {
                const hash = await sha256Hex(imgBytes);
                if (!seenHashes.has(hash)) {
                    seenHashes.add(hash);
                    extractedCount++;
                    outZip.file(`images/anh_${String(extractedCount).padStart(3, '0')}${imgExt}`, imgBytes, {
                        compression: 'STORE',
                    });

                    if (!previewBase64) {
                        try {
                            const dec = await decodeImageBytes(imgBytes, MIME_BY_EXT[imgExt] || 'image/jpeg');
                            try {
                                const dims = clampCanvasDimensions(dec.width, dec.height, 1440 * 1440);
                                canvas.width = dims.width;
                                canvas.height = dims.height;
                                const pctx = canvas.getContext('2d', { alpha: false });
                                pctx.fillStyle = '#ffffff';
                                pctx.fillRect(0, 0, canvas.width, canvas.height);
                                pctx.drawImage(dec.source, 0, 0, canvas.width, canvas.height);
                                previewBase64 = await canvasToJpegDataUri(canvas, 0.82);
                            } finally {
                                dec.close();
                            }
                        } catch (_pErr) {
                            // Ignore preview decode error
                        }
                    }
                }
            }

            notifyProgress(
                onProgress,
                i + 1,
                Math.max(allImageStreams.length, 1),
                `Đang quét ảnh gốc nhúng trong PDF (${i + 1}/${allImageStreams.length})...`
            );
        }
    } finally {
        releaseCanvas(canvas);
    }

    return { extractedCount, previewBase64 };
}

async function renderPdfPagesToZipOffline(rawPdfBytes, outZip, folder, prefix, onProgress, signal) {
    const pdfDoc = await openPdfJsDoc(rawPdfBytes);
    const canvas = document.createElement('canvas');
    let extractedCount = 0;
    let previewBase64 = null;

    try {
        const totalPages = pdfDoc.numPages;
        for (let pNum = 1; pNum <= totalPages; pNum++) {
            throwIfAborted(signal);
            const page = await pdfDoc.getPage(pNum);
            const vp1 = page.getViewport({ scale: 1.0 });
            page.cleanup();
            const maxPt = Math.max(vp1.width, vp1.height, 1.0);
            const scale = Math.max(0.5, Math.min(4.0, 1920.0 / maxPt));

            await renderPdfPageToCanvas(pdfDoc, pNum, scale, canvas);
            if (!previewBase64) {
                previewBase64 = await canvasToJpegDataUri(canvas, 0.82);
            }
            const jpgBytes = await canvasToBytes(canvas, 'image/jpeg', 0.90);
            outZip.file(
                `${folder}/${prefix}_${String(pNum).padStart(3, '0')}.jpg`,
                jpgBytes,
                { compression: 'STORE' }
            );
            extractedCount++;

            notifyProgress(
                onProgress,
                pNum,
                totalPages,
                `Đang kết xuất trang HD ${pNum}/${totalPages}...`
            );
        }
    } finally {
        releaseCanvas(canvas);
        await pdfDoc.destroy();
    }

    return { extractedCount, previewBase64 };
}

export async function extractArchiveOffline(file, extractMode, onProgress, signal) {
    const ext = validateInputFileSafety(file);
    throwIfAborted(signal);

    const mode = extractMode === 'slides' ? 'slides' : 'images';
    const rawBytes = new Uint8Array(await file.arrayBuffer());
    const originalSizeMb = Math.round((rawBytes.byteLength / (1024 * 1024)) * 100) / 100;

    const JSZip = getJSZip();
    const outZip = new JSZip();
    let extractedCount = 0;
    let previewBase64 = null;

    if (ext === '.pdf') {
        if (!hasPdfMagic(rawBytes)) {
            throw new Error('Tệp .pdf không hợp lệ hoặc bị hỏng.');
        }
        if (mode === 'images') {
            const embedded = await extractPdfEmbeddedImagesToZip(rawBytes, outZip, onProgress, signal);
            extractedCount = embedded.extractedCount;
            previewBase64 = embedded.previewBase64;

            if (extractedCount === 0) {
                const fallback = await renderPdfPagesToZipOffline(
                    rawPdfBytesOrSlice(rawBytes),
                    outZip,
                    'pages',
                    'trang',
                    onProgress,
                    signal
                );
                extractedCount = fallback.extractedCount;
                previewBase64 = fallback.previewBase64;
            }
        } else {
            const rendered = await renderPdfPagesToZipOffline(
                rawBytes,
                outZip,
                'slides',
                'trang',
                onProgress,
                signal
            );
            extractedCount = rendered.extractedCount;
            previewBase64 = rendered.previewBase64;
        }
    } else {
        const isZip = hasZipMagic(rawBytes);
        if (!isZip) {
            throw new Error(
                'Định dạng .PPT nhị phân cũ (97–2003) cần lưu sang .PPTX hoặc .PDF để trích xuất tài nguyên trực tiếp trên trình duyệt Offline.'
            );
        }

        const inZip = await JSZip.loadAsync(rawBytes);
        validateZipSafety(inZip);
        await validateAllZipXmlDtdSafety(inZip, signal);

        if (mode === 'images') {
            const mediaEntries = Object.values(inZip.files).filter(
                (e) => !e.dir && e.name.startsWith('ppt/media/') && IMAGE_EXTENSIONS.has(getExt(e.name))
            );
            const seenHashes = new Set();
            const usedNames = new Set();
            const previewCanvas = document.createElement('canvas');

            try {
                for (let i = 0; i < mediaEntries.length; i++) {
                    throwIfAborted(signal);
                    const entry = mediaEntries[i];
                    const data = await entry.async('uint8array');
                    const hash = await sha256Hex(data);

                    if (!seenHashes.has(hash)) {
                        seenHashes.add(hash);
                        let baseName = entry.name.split('/').pop();
                        if (usedNames.has(baseName.toLowerCase())) {
                            const sfx = getExt(baseName);
                            baseName = `${getStem(baseName)}_${String(extractedCount + 1).padStart(3, '0')}${sfx}`;
                        }
                        usedNames.add(baseName.toLowerCase());
                        extractedCount++;

                        const mExt = getExt(baseName);
                        outZip.file(`images/${baseName}`, data, {
                            compression: STORED_ZIP_EXTENSIONS.has(mExt) ? 'STORE' : 'DEFLATE',
                        });

                        if (!previewBase64) {
                            try {
                                const dec = await decodeImageBytes(data, MIME_BY_EXT[mExt] || 'image/jpeg');
                                try {
                                    const dims = clampCanvasDimensions(dec.width, dec.height, 1440 * 1440);
                                    previewCanvas.width = dims.width;
                                    previewCanvas.height = dims.height;
                                    const pctx = previewCanvas.getContext('2d', { alpha: false });
                                    pctx.fillStyle = '#ffffff';
                                    pctx.fillRect(0, 0, previewCanvas.width, previewCanvas.height);
                                    pctx.drawImage(dec.source, 0, 0, previewCanvas.width, previewCanvas.height);
                                    previewBase64 = await canvasToJpegDataUri(previewCanvas, 0.82);
                                } finally {
                                    dec.close();
                                }
                            } catch (_err) {
                                // Ignore preview decode error for exotic image
                            }
                        }
                    }

                    notifyProgress(
                        onProgress,
                        i + 1,
                        Math.max(mediaEntries.length, 1),
                        `Đang trích xuất ảnh gốc ${i + 1}/${mediaEntries.length}...`
                    );
                }
            } finally {
                releaseCanvas(previewCanvas);
            }
        } else {
            const { slidePaths, emuW, emuH } = await getOrderedPptxSlidePaths(inZip);
            const canvas = document.createElement('canvas');
            try {
                for (let i = 0; i < slidePaths.length; i++) {
                    throwIfAborted(signal);
                    await renderPptxSlideToCanvas(inZip, slidePaths[i], emuW, emuH, 1920, canvas);
                    if (!previewBase64) {
                        previewBase64 = await canvasToJpegDataUri(canvas, 0.82);
                    }
                    const jpgBytes = await canvasToBytes(canvas, 'image/jpeg', 0.90);
                    outZip.file(`slides/slide_${String(i + 1).padStart(3, '0')}.jpg`, jpgBytes, {
                        compression: 'STORE',
                    });
                    extractedCount++;

                    notifyProgress(
                        onProgress,
                        i + 1,
                        slidePaths.length,
                        `Đang tách Slide Full HD ${i + 1}/${slidePaths.length}...`
                    );
                }
            } finally {
                releaseCanvas(canvas);
            }
        }
    }

    throwIfAborted(signal);
    const zipBytes = await outZip.generateAsync({
        type: 'uint8array',
        compression: 'DEFLATE',
        compressionOptions: { level: 6 },
    });
    const zipSizeBytes = zipBytes.byteLength;
    const zipSizeMb = Math.round((zipSizeBytes / (1024 * 1024)) * 10000) / 10000;

    return {
        status: 'done',
        job_type: 'extract',
        extract_mode: mode,
        extracted_count: extractedCount,
        zip_size_bytes: zipSizeBytes,
        zip_size_mb: zipSizeMb,
        original_size_mb: originalSizeMb,
        compressed_size_mb: Math.round(zipSizeMb * 100) / 100,
        output_filename: `${getStem(file.name)}_${mode}.zip`,
        output_format: 'zip',
        output_blob: new Blob([zipBytes], { type: 'application/zip' }),
        preview_base64: previewBase64,
    };
}

function rawPdfBytesOrSlice(bytes) {
    return bytes;
}
