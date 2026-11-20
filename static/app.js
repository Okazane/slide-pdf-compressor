import {
    MAX_BACKUP_BYTES,
    addHistoryEntry,
    clearLocalDatabase,
    exportDatabaseToJson,
    importDatabaseFromJson,
    listHistoryEntries,
    loadPreferences,
    requestPersistentStorage,
    savePreferences,
    validateSafeJpegDataUri,
} from './db.js';

import {
    compressPdfOffline,
    compressPptxDirectOffline,
    convertPptxToPdfOffline,
    extractArchiveOffline,
    inspectFileOffline,
} from './offline-engine.js';

function initApp() {
    // State
    let currentFileInfo = null;
    let currentRawFile = null;
    let currentBlobUrl = null;
    let activeActionTab = 'compress';
    let pollTimer = null;
    let activeAbortController = null;
    let localPythonBackendAvailable = false;
    let stickyPrefs = {
        theme: 'dark',
        targetMb: 10,
        mode: 'flatten',
        outputFormat: 'pdf',
    };

    // Elements
    const htmlRoot = document.documentElement;
    const offlineStatusPill = document.getElementById('offlineStatusPill');
    const themeToggleBtn = document.getElementById('themeToggleBtn');
    const themeIcon = document.getElementById('themeIcon');
    const themeLabel = document.getElementById('themeLabel');

    const openHistoryBtn = document.getElementById('openHistoryBtn');
    const historyCountBadge = document.getElementById('historyCountBadge');
    const historyModal = document.getElementById('historyModal');
    const closeHistoryModalBtn = document.getElementById('closeHistoryModalBtn');
    const exportDataBtn = document.getElementById('exportDataBtn');
    const importDataBtn = document.getElementById('importDataBtn');
    const importDataInput = document.getElementById('importDataInput');
    const clearDataBtn = document.getElementById('clearDataBtn');
    const confirmClearBar = document.getElementById('confirmClearBar');
    const confirmClearYesBtn = document.getElementById('confirmClearYesBtn');
    const confirmClearNoBtn = document.getElementById('confirmClearNoBtn');
    const historyFeedbackBanner = document.getElementById('historyFeedbackBanner');
    const historyEmptyState = document.getElementById('historyEmptyState');
    const historyListContainer = document.getElementById('historyListContainer');

    const uploadSection = document.getElementById('uploadSection');
    const dropZone = document.getElementById('dropZone');
    const fileInput = document.getElementById('fileInput');
    const analyzeLoading = document.getElementById('analyzeLoading');
    const analyzeStatusText = document.getElementById('analyzeStatusText');
    const cancelAnalyzeBtn = document.getElementById('cancelAnalyzeBtn');

    const configSection = document.getElementById('configSection');
    const initialPreviewImg = document.getElementById('initialPreviewImg');
    const initialThumbPlaceholder = document.getElementById('initialThumbPlaceholder');
    const fileTypeBadge = document.getElementById('fileTypeBadge');
    const fileNameDisplay = document.getElementById('fileNameDisplay');
    const origSizeDisplay = document.getElementById('origSizeDisplay');
    const pageCountDisplay = document.getElementById('pageCountDisplay');
    const mediaCountPill = document.getElementById('mediaCountPill');
    const mediaCountDisplay = document.getElementById('mediaCountDisplay');
    const changeFileBtn = document.getElementById('changeFileBtn');

    const tabCompressBtn = document.getElementById('tabCompressBtn');
    const tabExtractBtn = document.getElementById('tabExtractBtn');

    const pdfOutputOptions = document.getElementById('pdfOutputOptions');
    const pptxOutputOptions = document.getElementById('pptxOutputOptions');
    const pptxPdfOptionCard = document.getElementById('pptxPdfOptionCard');
    const pptxDirectOptionCard = document.getElementById('pptxDirectOptionCard');
    const pptLegacyCompressHint = document.getElementById('pptLegacyCompressHint');
    const modeConfigGroup = document.getElementById('modeConfigGroup');
    const modeFlattenCard = document.getElementById('modeFlattenCard');
    const modeVectorCard = document.getElementById('modeVectorCard');
    const targetPills = document.getElementById('targetPills');
    const targetMbInput = document.getElementById('targetMbInput');
    const compressForm = document.getElementById('compressForm');
    const startCompressBtn = document.getElementById('startCompressBtn');

    const extractForm = document.getElementById('extractForm');
    const extractImagesCard = document.getElementById('extractImagesCard');
    const extractSlidesCard = document.getElementById('extractSlidesCard');
    const pptLegacyExtractHint = document.getElementById('pptLegacyExtractHint');
    const startExtractBtn = document.getElementById('startExtractBtn');

    const progressSection = document.getElementById('progressSection');
    const progressStatusTitle = document.getElementById('progressStatusTitle');
    const progressDetailText = document.getElementById('progressDetailText');
    const progressPercentBadge = document.getElementById('progressPercentBadge');
    const progressBarFill = document.getElementById('progressBarFill');
    const cancelJobBtn = document.getElementById('cancelJobBtn');

    const resultSection = document.getElementById('resultSection');
    const resultStatusIcon = document.getElementById('resultStatusIcon');
    const resultMainTitle = document.getElementById('resultMainTitle');
    const resultSubSummary = document.getElementById('resultSubSummary');
    const savedPercentBadge = document.getElementById('savedPercentBadge');
    const resBeforeLabel = document.getElementById('resBeforeLabel');
    const resBeforeMb = document.getElementById('resBeforeMb');
    const resAfterLabel = document.getElementById('resAfterLabel');
    const resAfterMb = document.getElementById('resAfterMb');
    const downloadFileBtn = document.getElementById('downloadFileBtn');
    const downloadBtnText = document.getElementById('downloadBtnText');
    const downloadFilenameLabel = document.getElementById('downloadFilenameLabel');
    const compressAnotherBtn = document.getElementById('compressAnotherBtn');

    const previewComparisonPanel = document.getElementById('previewComparisonPanel');
    const previewBeforeSizeTag = document.getElementById('previewBeforeSizeTag');
    const previewAfterSizeTag = document.getElementById('previewAfterSizeTag');
    const beforePreviewImg = document.getElementById('beforePreviewImg');
    const afterPreviewImg = document.getElementById('afterPreviewImg');

    const extractPreviewPanel = document.getElementById('extractPreviewPanel');
    const extractPreviewCountTag = document.getElementById('extractPreviewCountTag');
    const extractSamplePreviewImg = document.getElementById('extractSamplePreviewImg');

    const errorAlert = document.getElementById('errorAlert');
    const errorMessageText = document.getElementById('errorMessageText');
    const closeErrorBtn = document.getElementById('closeErrorBtn');

    const imageModal = document.getElementById('imageModal');
    const modalTitle = document.getElementById('modalTitle');
    const modalZoomImg = document.getElementById('modalZoomImg');
    const closeModalBtn = document.getElementById('closeModalBtn');

    // -----------------------------------------------------------------------
    // ObjectURL Hygiene & Connectivity Status
    // -----------------------------------------------------------------------

    function revokeCurrentBlobUrl() {
        if (currentBlobUrl) {
            URL.revokeObjectURL(currentBlobUrl);
            currentBlobUrl = null;
        }
    }

    window.addEventListener('pagehide', revokeCurrentBlobUrl);

    function updateOfflineStatusBadge() {
        if (!offlineStatusPill) return;
        if (navigator.onLine === false) {
            offlineStatusPill.textContent = '📴 Đang Ngoại Tuyến • Hoạt động 100% Offline';
        } else {
            offlineStatusPill.textContent = '⚡ 100% Offline • Xử lý trực tiếp trên trình duyệt';
        }
    }

    window.addEventListener('online', updateOfflineStatusBadge);
    window.addEventListener('offline', updateOfflineStatusBadge);
    updateOfflineStatusBadge();

    // Check same-origin local Python engine ONLY when served on localhost (for legacy .ppt 97-2003)
    async function probeSameOriginLocalPython() {
        const host = window.location.hostname;
        if (host !== 'localhost' && host !== '127.0.0.1') {
            localPythonBackendAvailable = false;
            return false;
        }
        const ctrl = new AbortController();
        const timer = setTimeout(() => ctrl.abort(), 1200);
        try {
            const resp = await fetch('./api/health', {
                method: 'GET',
                cache: 'no-store',
                signal: ctrl.signal,
            });
            if (!resp.ok) {
                localPythonBackendAvailable = false;
                return false;
            }
            const data = await resp.json();
            localPythonBackendAvailable = Boolean(data && data.status === 'ok');
            return localPythonBackendAvailable;
        } catch (_err) {
            localPythonBackendAvailable = false;
            return false;
        } finally {
            clearTimeout(timer);
        }
    }

    // -----------------------------------------------------------------------
    // Preferences & Theme Management (Dexie IndexedDB)
    // -----------------------------------------------------------------------

    function applyThemeVisual(theme) {
        const clean = theme === 'light' ? 'light' : 'dark';
        htmlRoot.setAttribute('data-theme', clean);
        if (clean === 'light') {
            themeIcon.textContent = '🌙';
            themeLabel.textContent = 'Giao diện Tối';
        } else {
            themeIcon.textContent = '☀️';
            themeLabel.textContent = 'Giao diện Sáng';
        }
    }

    function applyPreferencesToControls(prefs) {
        stickyPrefs = { ...stickyPrefs, ...prefs };
        applyThemeVisual(stickyPrefs.theme);

        const mbVal = Number(stickyPrefs.targetMb) || 10;
        targetMbInput.value = String(mbVal);
        targetPills.querySelectorAll('.pill-btn').forEach((btn) => {
            btn.classList.toggle('active', Number(btn.getAttribute('data-mb')) === mbVal);
        });

        if (stickyPrefs.mode === 'vector') {
            selectRadioCard(modeVectorCard);
        } else {
            selectRadioCard(modeFlattenCard);
        }

        if (currentFileInfo && currentFileInfo.file_type === 'pptx' && !currentFileInfo.is_legacy_ppt) {
            if (stickyPrefs.outputFormat === 'pptx') {
                selectRadioCard(pptxDirectOptionCard);
            } else {
                selectRadioCard(pptxPdfOptionCard);
            }
            updateModeVisibilityForPptx();
        }
    }

    async function persistPrefUpdate(partial) {
        try {
            stickyPrefs = await savePreferences(partial);
        } catch (_err) {
            stickyPrefs = { ...stickyPrefs, ...partial };
        }
    }

    themeToggleBtn.addEventListener('click', () => {
        const current = htmlRoot.getAttribute('data-theme') || 'dark';
        const next = current === 'dark' ? 'light' : 'dark';
        applyThemeVisual(next);
        persistPrefUpdate({ theme: next });
    });

    // -----------------------------------------------------------------------
    // Error Banner & Cancel Controls
    // -----------------------------------------------------------------------

    function showError(msg) {
        errorMessageText.textContent = msg;
        errorAlert.classList.remove('hidden');
    }

    function hideError() {
        errorAlert.classList.add('hidden');
    }

    closeErrorBtn.addEventListener('click', hideError);

    function cancelActiveWork() {
        if (activeAbortController) {
            activeAbortController.abort();
            activeAbortController = null;
        }
        if (pollTimer) {
            clearInterval(pollTimer);
            pollTimer = null;
        }
        analyzeLoading.classList.add('hidden');
        progressSection.classList.add('hidden');
        reenableActionButtons();
    }

    if (cancelJobBtn) {
        cancelJobBtn.addEventListener('click', () => {
            cancelActiveWork();
            showError('Đã huỷ tác vụ xử lý theo yêu cầu của bạn.');
        });
    }

    if (cancelAnalyzeBtn) {
        cancelAnalyzeBtn.addEventListener('click', () => {
            cancelActiveWork();
        });
    }

    // -----------------------------------------------------------------------
    // Action Mode Tabs & Radio Card Helpers
    // -----------------------------------------------------------------------

    function switchActionTab(tab) {
        activeActionTab = tab === 'extract' ? 'extract' : 'compress';
        hideError();

        if (activeActionTab === 'compress') {
            tabCompressBtn.classList.add('active');
            tabCompressBtn.setAttribute('aria-selected', 'true');
            tabExtractBtn.classList.remove('active');
            tabExtractBtn.setAttribute('aria-selected', 'false');

            compressForm.classList.remove('hidden');
            extractForm.classList.add('hidden');
        } else {
            tabExtractBtn.classList.add('active');
            tabExtractBtn.setAttribute('aria-selected', 'true');
            tabCompressBtn.classList.remove('active');
            tabCompressBtn.setAttribute('aria-selected', 'false');

            extractForm.classList.remove('hidden');
            compressForm.classList.add('hidden');
        }
    }

    tabCompressBtn.addEventListener('click', () => switchActionTab('compress'));
    tabExtractBtn.addEventListener('click', () => switchActionTab('extract'));

    function setRadioOptionState(cardEl, disabled) {
        if (!cardEl) return;
        const radio = cardEl.querySelector('input[type="radio"]');
        if (radio) {
            radio.disabled = Boolean(disabled);
        }
        cardEl.classList.toggle('disabled', Boolean(disabled));
    }

    function selectRadioCard(cardEl) {
        if (!cardEl) return;
        const radio = cardEl.querySelector('input[type="radio"]');
        const group = cardEl.closest('.radio-card-group');
        if (group) {
            group.querySelectorAll('.radio-card').forEach((c) => c.classList.remove('active'));
        }
        if (radio) {
            radio.checked = true;
        }
        cardEl.classList.add('active');
    }

    function updateModeVisibilityForPptx() {
        const selectedPptxOut = document.querySelector('input[name="output_format_pptx"]:checked');
        if (selectedPptxOut && selectedPptxOut.value === 'pptx') {
            modeConfigGroup.classList.add('hidden');
        } else {
            modeConfigGroup.classList.remove('hidden');
        }
    }

    document.querySelectorAll('.radio-card input[type="radio"]').forEach((radio) => {
        radio.addEventListener('change', (e) => {
            if (e.target.disabled) return;
            const group = e.target.closest('.radio-card-group');
            if (group) {
                group.querySelectorAll('.radio-card').forEach((c) => c.classList.remove('active'));
            }
            const card = e.target.closest('.radio-card');
            if (card) {
                card.classList.add('active');
            }
            if (e.target.name === 'output_format_pptx') {
                updateModeVisibilityForPptx();
                persistPrefUpdate({ outputFormat: e.target.value });
            } else if (e.target.name === 'compression_mode') {
                persistPrefUpdate({ mode: e.target.value });
            }
        });
    });

    targetPills.querySelectorAll('.pill-btn').forEach((btn) => {
        btn.addEventListener('click', () => {
            targetPills.querySelectorAll('.pill-btn').forEach((b) => b.classList.remove('active'));
            btn.classList.add('active');
            const mb = parseFloat(btn.getAttribute('data-mb')) || 10;
            targetMbInput.value = String(mb);
            persistPrefUpdate({ targetMb: mb });
        });
    });

    targetMbInput.addEventListener('input', () => {
        const valStr = String(targetMbInput.value).trim();
        targetPills.querySelectorAll('.pill-btn').forEach((btn) => {
            btn.classList.toggle('active', btn.getAttribute('data-mb') === valStr);
        });
        const num = parseFloat(valStr);
        if (Number.isFinite(num) && num >= 0.5 && num <= 200) {
            persistPrefUpdate({ targetMb: num });
        }
    });

    function reenableActionButtons() {
        const isOfflineUnsupportedLegacyPpt = Boolean(
            currentFileInfo &&
            (currentFileInfo.file_type === 'ppt' || currentFileInfo.is_legacy_ppt) &&
            !localPythonBackendAvailable
        );
        startCompressBtn.disabled = isOfflineUnsupportedLegacyPpt;
        if (startExtractBtn) {
            startExtractBtn.disabled = isOfflineUnsupportedLegacyPpt;
        }
    }

    // -----------------------------------------------------------------------
    // File Upload & Inspection (100% Browser Offline Primary Engine)
    // -----------------------------------------------------------------------

    dropZone.addEventListener('click', () => fileInput.click());
    dropZone.addEventListener('keydown', (e) => {
        if (e.key === 'Enter' || e.key === ' ') {
            e.preventDefault();
            fileInput.click();
        }
    });

    dropZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        dropZone.classList.add('drag-over');
    });

    dropZone.addEventListener('dragleave', () => {
        dropZone.classList.remove('drag-over');
    });

    dropZone.addEventListener('drop', (e) => {
        e.preventDefault();
        dropZone.classList.remove('drag-over');
        if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
            handleFileUpload(e.dataTransfer.files[0]);
        }
    });

    fileInput.addEventListener('change', () => {
        if (fileInput.files && fileInput.files.length > 0) {
            handleFileUpload(fileInput.files[0]);
        }
    });

    async function safeParseResponse(response) {
        if (response.status === 413) {
            throw new Error('Dung lượng tệp vượt giới hạn tối đa cho phép (500 MB).');
        }
        const text = await response.text();
        try {
            return JSON.parse(text);
        } catch (_err) {
            throw new Error(text ? text.slice(0, 180) : `Máy chủ phản hồi lỗi HTTP ${response.status}.`);
        }
    }

    async function handleFileUpload(file) {
        hideError();
        const dotIdx = file.name.lastIndexOf('.');
        const ext = dotIdx >= 0 ? file.name.slice(dotIdx).toLowerCase() : '';
        if (!['.pdf', '.pptx', '.ppt'].includes(ext)) {
            showError('Định dạng tệp không hợp lệ. Vui lòng chọn tệp .pdf, .pptx hoặc .ppt.');
            return;
        }
        if (file.size > 500 * 1024 * 1024) {
            showError('Dung lượng tệp vượt quá giới hạn tối đa 500 MB.');
            return;
        }

        if (activeAbortController) {
            activeAbortController.abort();
        }
        activeAbortController = new AbortController();
        const { signal } = activeAbortController;

        analyzeLoading.classList.remove('hidden');
        analyzeStatusText.textContent = `Đang phân tích cục bộ "${file.name}" trên trình duyệt...`;

        try {
            let info = await inspectFileOffline(file, signal);

            // Only if binary .ppt (97-2003) AND local Python server is running on localhost, enrich via /api/analyze
            if (info.is_legacy_ppt && (await probeSameOriginLocalPython())) {
                const formData = new FormData();
                formData.append('file', file);
                const resp = await fetch('./api/analyze', {
                    method: 'POST',
                    body: formData,
                    signal,
                });
                const apiData = await safeParseResponse(resp);
                if (resp.ok) {
                    info = apiData;
                }
            }

            currentRawFile = file;
            currentFileInfo = info;
            populateConfigPanel(info);
        } catch (err) {
            if (err && err.name === 'AbortError') return;
            showError(err.message || 'Không thể phân tích cấu trúc tệp.');
        } finally {
            analyzeLoading.classList.add('hidden');
            activeAbortController = null;
        }
    }

    function populateConfigPanel(info) {
        uploadSection.classList.add('hidden');
        resultSection.classList.add('hidden');
        progressSection.classList.add('hidden');
        configSection.classList.remove('hidden');

        fileNameDisplay.textContent = info.filename;
        origSizeDisplay.textContent = `${Number(info.original_size_mb).toFixed(2)} MB`;

        const isLegacyPpt = Boolean(info.file_type === 'ppt' || info.is_legacy_ppt);
        const count = info.page_count || info.slide_count || 0;
        const mediaCount = Number(info.media_count || 0);

        setRadioOptionState(pptxDirectOptionCard, false);
        setRadioOptionState(extractImagesCard, false);
        pptLegacyCompressHint.classList.add('hidden');
        pptLegacyExtractHint.classList.add('hidden');

        if (info.file_type === 'pdf') {
            fileTypeBadge.textContent = 'PDF';
            fileTypeBadge.className = 'format-badge pdf-badge';
            pageCountDisplay.textContent = `${count} trang`;

            if (mediaCount > 0) {
                mediaCountDisplay.textContent = `${mediaCount} ảnh`;
                mediaCountPill.classList.remove('hidden');
            } else {
                mediaCountPill.classList.add('hidden');
            }

            pdfOutputOptions.classList.remove('hidden');
            pptxOutputOptions.classList.add('hidden');
            modeConfigGroup.classList.remove('hidden');
            selectRadioCard(extractImagesCard);
        } else if (isLegacyPpt) {
            fileTypeBadge.textContent = 'PPT';
            fileTypeBadge.className = 'format-badge ppt-badge';
            pageCountDisplay.textContent = count > 0 ? `${count} slide` : 'Tệp .PPT 97–2003';
            mediaCountPill.classList.add('hidden');

            pdfOutputOptions.classList.add('hidden');
            pptxOutputOptions.classList.remove('hidden');

            setRadioOptionState(pptxDirectOptionCard, true);
            selectRadioCard(pptxPdfOptionCard);
            pptLegacyCompressHint.classList.remove('hidden');
            updateModeVisibilityForPptx();

            setRadioOptionState(extractImagesCard, true);
            selectRadioCard(extractSlidesCard);
            pptLegacyExtractHint.classList.remove('hidden');

            if (localPythonBackendAvailable) {
                pptLegacyCompressHint.textContent =
                    'ℹ️ Định dạng .PPT cũ: Đã kết nối Engine Python cục bộ (127.0.0.1) để hỗ trợ chuyển đổi sang .PDF nhẹ.';
                pptLegacyExtractHint.textContent =
                    'ℹ️ Định dạng .PPT cũ: Hỗ trợ tách từng slide thành bộ ảnh HD (.JPG) qua Engine Python cục bộ.';
            } else {
                pptLegacyCompressHint.textContent =
                    '⚠️ Định dạng .PPT nhị phân cũ (PowerPoint 97–2003) cần mở bằng PowerPoint và lưu sang .PPTX hoặc .PDF để xử lý trực tiếp trên trình duyệt Offline.';
                pptLegacyExtractHint.textContent =
                    '⚠️ Định dạng .PPT nhị phân cũ (PowerPoint 97–2003) cần lưu sang .PPTX hoặc .PDF để trích xuất tài nguyên trực tiếp trên trình duyệt Offline.';
            }
        } else {
            fileTypeBadge.textContent = 'PPTX';
            fileTypeBadge.className = 'format-badge pptx-badge';
            pageCountDisplay.textContent = `${count} slide`;

            mediaCountDisplay.textContent = `${mediaCount} ảnh`;
            mediaCountPill.classList.remove('hidden');

            pdfOutputOptions.classList.add('hidden');
            pptxOutputOptions.classList.remove('hidden');

            if (stickyPrefs.outputFormat === 'pptx') {
                selectRadioCard(pptxDirectOptionCard);
            } else {
                selectRadioCard(pptxPdfOptionCard);
            }
            updateModeVisibilityForPptx();
            selectRadioCard(extractImagesCard);
        }

        const safePreview = (() => {
            try {
                return validateSafeJpegDataUri(info.preview_base64);
            } catch (_err) {
                return null;
            }
        })();

        if (safePreview) {
            initialPreviewImg.src = safePreview;
            initialPreviewImg.classList.remove('hidden');
            initialThumbPlaceholder.classList.add('hidden');
        } else {
            initialPreviewImg.classList.add('hidden');
            initialThumbPlaceholder.classList.remove('hidden');
            initialThumbPlaceholder.textContent = (info.file_type || 'SLIDE').toUpperCase();
        }

        reenableActionButtons();
    }

    function resetToUpload() {
        cancelActiveWork();
        revokeCurrentBlobUrl();
        currentFileInfo = null;
        currentRawFile = null;
        fileInput.value = '';
        hideError();
        configSection.classList.add('hidden');
        progressSection.classList.add('hidden');
        resultSection.classList.add('hidden');
        uploadSection.classList.remove('hidden');
    }

    changeFileBtn.addEventListener('click', resetToUpload);
    compressAnotherBtn.addEventListener('click', resetToUpload);

    function updateProgressUi(current, total, message, unitLabel = 'mục') {
        const safeTotal = Math.max(Number(total) || 1, 1);
        const safeCur = Math.max(0, Number(current) || 0);
        const pct = Math.min(Math.round((safeCur / safeTotal) * 100), 100);
        progressStatusTitle.textContent = message || `Đang xử lý ${safeCur}/${safeTotal} ${unitLabel}...`;
        progressDetailText.textContent = `Tiến độ xử lý cục bộ: ${safeCur}/${safeTotal} ${unitLabel}`;
        progressPercentBadge.textContent = `${pct}%`;
        progressBarFill.style.width = `${Math.max(pct, 6)}%`;
    }

    // -----------------------------------------------------------------------
    // Smart Compression Submit (100% Client-Side Offline + Local .PPT Fallback)
    // -----------------------------------------------------------------------

    compressForm.addEventListener('submit', async (e) => {
        e.preventDefault();
        if (!currentFileInfo || !currentRawFile) return;
        hideError();

        const isLegacyPpt = Boolean(currentFileInfo.file_type === 'ppt' || currentFileInfo.is_legacy_ppt);
        let outputFormat = 'pdf';
        if (currentFileInfo.file_type === 'pptx' && !isLegacyPpt) {
            const checkedOut = document.querySelector('input[name="output_format_pptx"]:checked');
            outputFormat = checkedOut ? checkedOut.value : 'pdf';
        }

        const checkedMode = document.querySelector('input[name="compression_mode"]:checked');
        const mode = checkedMode ? checkedMode.value : 'flatten';
        const targetMb = Math.min(200, Math.max(0.5, parseFloat(targetMbInput.value) || 10.0));

        await persistPrefUpdate({ targetMb, mode, outputFormat });

        if (isLegacyPpt && !localPythonBackendAvailable) {
            showError(
                'Tệp .PPT nhị phân cũ (97–2003) cần mở bằng PowerPoint và lưu dưới dạng .PPTX hoặc .PDF để nén trực tiếp 100% Offline trên trình duyệt.'
            );
            return;
        }

        startCompressBtn.disabled = true;
        if (startExtractBtn) startExtractBtn.disabled = true;
        progressSection.classList.remove('hidden');
        resultSection.classList.add('hidden');

        const totalUnits = outputFormat === 'pptx'
            ? (currentFileInfo.media_count || currentFileInfo.slide_count || 1)
            : (currentFileInfo.page_count || currentFileInfo.slide_count || 1);
        const unitLabel = outputFormat === 'pptx' ? 'ảnh' : 'trang';
        updateProgressUi(0, totalUnits, `Đang khởi tạo nén thông minh (mục tiêu < ${targetMb} MB)...`, unitLabel);

        if (activeAbortController) activeAbortController.abort();
        activeAbortController = new AbortController();
        const { signal } = activeAbortController;

        try {
            if (isLegacyPpt && localPythonBackendAvailable && currentFileInfo.file_id) {
                const resp = await fetch('./api/compress', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({
                        file_id: currentFileInfo.file_id,
                        target_mb: targetMb,
                        mode,
                        output_format: 'pdf',
                    }),
                    signal,
                });
                const data = await safeParseResponse(resp);
                if (!resp.ok) {
                    throw new Error(data.detail || 'Không thể bắt đầu nén tệp .ppt.');
                }
                pollLocalJobStatus(data.job_id, 'compress');
                return;
            }

            let result;
            const onProg = (cur, tot, msg) => updateProgressUi(cur, tot, msg, unitLabel);
            if (currentFileInfo.file_type === 'pdf') {
                result = await compressPdfOffline(
                    currentRawFile,
                    { targetMb, mode, beforePreview: currentFileInfo.preview_base64 },
                    onProg,
                    signal
                );
            } else if (outputFormat === 'pptx') {
                result = await compressPptxDirectOffline(
                    currentRawFile,
                    { targetMb, beforePreview: currentFileInfo.preview_base64 },
                    onProg,
                    signal
                );
            } else {
                result = await convertPptxToPdfOffline(
                    currentRawFile,
                    { targetMb, mode, beforePreview: currentFileInfo.preview_base64 },
                    onProg,
                    signal
                );
            }

            activeAbortController = null;
            reenableActionButtons();
            await recordJobInHistory(result, 'compress');
            renderResultDashboard(result);
        } catch (err) {
            activeAbortController = null;
            reenableActionButtons();
            progressSection.classList.add('hidden');
            if (err && err.name === 'AbortError') return;
            showError(err.message || 'Quá trình nén tệp gặp sự cố.');
        }
    });

    // -----------------------------------------------------------------------
    // Resource Extraction Submit (.ZIP)
    // -----------------------------------------------------------------------

    extractForm.addEventListener('submit', async (e) => {
        e.preventDefault();
        if (!currentFileInfo || !currentRawFile) return;
        hideError();

        const isLegacyPpt = Boolean(currentFileInfo.file_type === 'ppt' || currentFileInfo.is_legacy_ppt);
        const checkedExtract = document.querySelector('input[name="extract_mode"]:checked');
        const extractMode = checkedExtract ? checkedExtract.value : 'images';

        if (isLegacyPpt && !localPythonBackendAvailable) {
            showError(
                'Tệp .PPT nhị phân cũ (97–2003) cần lưu sang .PPTX hoặc .PDF để trích xuất gói .ZIP trực tiếp 100% Offline trên trình duyệt.'
            );
            return;
        }

        startExtractBtn.disabled = true;
        startCompressBtn.disabled = true;
        progressSection.classList.remove('hidden');
        resultSection.classList.add('hidden');

        const estTotal = extractMode === 'images' && currentFileInfo.media_count
            ? currentFileInfo.media_count
            : (currentFileInfo.page_count || currentFileInfo.slide_count || 1);
        updateProgressUi(0, estTotal, 'Đang quét cấu trúc tệp và trích xuất gói .ZIP...', 'mục');

        if (activeAbortController) activeAbortController.abort();
        activeAbortController = new AbortController();
        const { signal } = activeAbortController;

        try {
            if (isLegacyPpt && localPythonBackendAvailable && currentFileInfo.file_id) {
                const resp = await fetch('/api/extract', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({
                        file_id: currentFileInfo.file_id,
                        extract_mode: extractMode,
                    }),
                    signal,
                });
                const data = await safeParseResponse(resp);
                if (!resp.ok) {
                    throw new Error(data.detail || 'Không thể giải nén tệp .ppt.');
                }
                pollLocalJobStatus(data.job_id, 'extract');
                return;
            }

            const result = await extractArchiveOffline(
                currentRawFile,
                extractMode,
                (cur, tot, msg) => updateProgressUi(cur, tot, msg, 'mục'),
                signal
            );

            activeAbortController = null;
            reenableActionButtons();
            await recordJobInHistory(result, 'extract');
            renderResultDashboard(result);
        } catch (err) {
            activeAbortController = null;
            reenableActionButtons();
            progressSection.classList.add('hidden');
            if (err && err.name === 'AbortError') return;
            showError(err.message || 'Quá trình giải nén tài nguyên gặp sự cố.');
        }
    });

    function pollLocalJobStatus(jobId, expectedJobType = 'compress') {
        if (pollTimer) clearInterval(pollTimer);

        pollTimer = setInterval(async () => {
            try {
                const res = await fetch(`./api/job/${encodeURIComponent(jobId)}`);
                const job = await res.json();
                if (!res.ok) {
                    throw new Error(job.detail || 'Không thể lấy trạng thái tiến trình.');
                }

                const cur = job.progress_current || 0;
                const tot = Math.max(job.progress_total || 1, 1);
                updateProgressUi(cur, tot, job.message, 'mục');

                if (job.status === 'done') {
                    clearInterval(pollTimer);
                    pollTimer = null;
                    activeAbortController = null;
                    reenableActionButtons();
                    await recordJobInHistory(job.result || job, expectedJobType);
                    renderResultDashboard(job);
                } else if (job.status === 'error') {
                    clearInterval(pollTimer);
                    pollTimer = null;
                    activeAbortController = null;
                    reenableActionButtons();
                    progressSection.classList.add('hidden');
                    showError(job.error || job.message || 'Quá trình xử lý gặp sự cố.');
                }
            } catch (err) {
                clearInterval(pollTimer);
                pollTimer = null;
                activeAbortController = null;
                reenableActionButtons();
                progressSection.classList.add('hidden');
                showError(err.message);
            }
        }, 400);
    }

    // -----------------------------------------------------------------------
    // Results Dashboard & History Persistence
    // -----------------------------------------------------------------------

    function formatZipSizeMb(res) {
        const rawMb = Number(res.zip_size_mb ?? res.compressed_size_mb ?? 0);
        if (rawMb > 0 && rawMb < 0.01) {
            return '< 0.01';
        }
        return rawMb.toFixed(2);
    }

    async function recordJobInHistory(res, fallbackJobType) {
        if (!currentFileInfo) return;
        try {
            const isExtract = (res.job_type || fallbackJobType) === 'extract' || res.output_format === 'zip';
            const previewCandidate = isExtract
                ? (res.preview_base64 || currentFileInfo.preview_base64)
                : (res.after_preview_base64 || res.before_preview_base64 || currentFileInfo.preview_base64);

            let safePreview = null;
            try {
                safePreview = validateSafeJpegDataUri(previewCandidate);
            } catch (_err) {
                safePreview = null;
            }

            await addHistoryEntry({
                filename: currentFileInfo.filename,
                fileType: currentFileInfo.file_type || 'pdf',
                jobType: isExtract ? 'extract' : 'compress',
                mode: isExtract ? (res.extract_mode || 'images') : (res.mode_used || stickyPrefs.mode || 'flatten'),
                outputFormat: isExtract ? 'zip' : (res.output_format || 'pdf'),
                targetMb: isExtract ? null : Number(res.target_mb ?? stickyPrefs.targetMb ?? 10),
                targetMet: isExtract ? null : Boolean(res.target_met !== false),
                originalSizeMb: Number(res.original_size_mb ?? currentFileInfo.original_size_mb ?? 0),
                compressedSizeMb: Number(res.compressed_size_mb ?? res.zip_size_mb ?? 0),
                reductionPercent: isExtract ? 0 : Number(res.reduction_percent ?? 0),
                itemCount: isExtract
                    ? Number(res.extracted_count ?? 0)
                    : Number(res.page_count ?? res.slide_count ?? currentFileInfo.page_count ?? 0),
                previewBase64: safePreview,
                createdAt: Date.now(),
            });
            await refreshHistoryBadgeAndList();
        } catch (_err) {
            // Non-blocking if browser is in strict private mode without IndexedDB
        }
    }

    function renderResultDashboard(job) {
        progressSection.classList.add('hidden');
        resultSection.classList.remove('hidden');

        const res = job.result || job;
        const jobType = res.job_type || job.job_type || (res.output_format === 'zip' ? 'extract' : 'compress');
        revokeCurrentBlobUrl();

        if (res.output_blob instanceof Blob) {
            currentBlobUrl = URL.createObjectURL(res.output_blob);
            downloadFileBtn.href = currentBlobUrl;
        } else if (res.download_url) {
            downloadFileBtn.href = res.download_url;
        } else if (job.job_id) {
            downloadFileBtn.href = `./api/download/${encodeURIComponent(job.job_id)}`;
        }

        if (res.output_filename) {
            downloadFileBtn.setAttribute('download', res.output_filename);
        }
        downloadFilenameLabel.textContent = res.output_filename ? `(${res.output_filename})` : '';

        if (jobType === 'extract' || res.output_format === 'zip') {
            const extractedCount = Number(res.extracted_count ?? 0);
            const zipMbStr = formatZipSizeMb(res);
            const isSlides = res.extract_mode === 'slides';

            resultStatusIcon.textContent = '📦';
            resultMainTitle.textContent = 'Giải Nén Tài Nguyên Hoàn Tất!';
            resultSubSummary.textContent = isSlides
                ? 'Toàn bộ các trang / slide đã được tách thành bộ ảnh HD (.JPG) đóng gói trong tệp .ZIP ngay trên trình duyệt.'
                : 'Toàn bộ hình ảnh nhúng đã được trích xuất nguyên bản (khử trùng lặp SHA-256) trong tệp .ZIP.';

            savedPercentBadge.className = 'saved-badge';
            savedPercentBadge.textContent = `Đã trích xuất ${extractedCount} mục`;

            resBeforeLabel.textContent = 'Số lượng trích xuất';
            resBeforeMb.textContent = `${extractedCount} ảnh/slide`;
            resAfterLabel.textContent = 'Dung lượng gói .ZIP';
            resAfterMb.className = 'metric-value highlight-green';
            resAfterMb.textContent = `${zipMbStr} MB`;

            downloadBtnText.textContent = 'Tải Gói Tài Nguyên (.ZIP)';
            compressAnotherBtn.textContent = 'Xử lý tệp khác';

            previewComparisonPanel.classList.add('hidden');
            const rawSample = res.preview_base64 || (currentFileInfo && currentFileInfo.preview_base64);
            let sampleImg = null;
            try {
                sampleImg = validateSafeJpegDataUri(rawSample);
            } catch (_err) {
                sampleImg = null;
            }

            if (sampleImg && extractPreviewPanel) {
                extractPreviewPanel.classList.remove('hidden');
                extractPreviewCountTag.textContent = `${extractedCount} mục (${zipMbStr} MB)`;
                extractSamplePreviewImg.src = sampleImg;
            } else if (extractPreviewPanel) {
                extractPreviewPanel.classList.add('hidden');
            }
        } else {
            if (extractPreviewPanel) {
                extractPreviewPanel.classList.add('hidden');
            }

            const origMb = Number(
                res.original_size_mb ?? (currentFileInfo && currentFileInfo.original_size_mb) ?? 0
            ).toFixed(2);
            const compMb = Number(res.compressed_size_mb ?? 0).toFixed(2);
            const rawReduction = Number(res.reduction_percent ?? 0);
            const targetMet = res.target_met !== false;
            const targetMbVal = Number(res.target_mb ?? job.target_mb ?? targetMbInput.value ?? 10).toFixed(1);

            resBeforeLabel.textContent = 'Dung lượng Trước khi nén';
            resBeforeMb.textContent = `${origMb} MB`;
            resAfterLabel.textContent = 'Dung lượng Sau khi nén';
            resAfterMb.textContent = `${compMb} MB`;

            downloadBtnText.textContent = 'Tải File Đã Nén';
            compressAnotherBtn.textContent = 'Nén tệp khác';

            const reductionLabel = rawReduction <= 0
                ? 'Giữ nguyên 0%'
                : `Giảm ${rawReduction.toFixed(1)}%`;

            if (targetMet) {
                resultStatusIcon.textContent = '✨';
                resultMainTitle.textContent = 'Nén Tệp Hoàn Tất!';
                resultSubSummary.textContent = 'Tệp đã đạt chuẩn dung lượng mục tiêu với độ sắc nét tối ưu.';
                savedPercentBadge.className = 'saved-badge';
                savedPercentBadge.textContent = reductionLabel;
                resAfterMb.className = 'metric-value highlight-green';
            } else {
                resultStatusIcon.textContent = '⚠️';
                resultMainTitle.textContent = 'Nén Tệp Hoàn Tất (Chạm Ngưỡng An Toàn)';
                resultSubSummary.textContent = `Đã nén tối đa trong ngưỡng an toàn giữ độ nét (${compMb} MB), chưa thể ép xuống dưới mức mục tiêu ${targetMbVal} MB mà không làm vỡ chữ/ảnh.`;
                savedPercentBadge.className = 'saved-badge warning-badge';
                savedPercentBadge.textContent = `${reductionLabel} • Trên mức ${targetMbVal} MB`;
                resAfterMb.className = 'metric-value highlight-warning';
            }

            let beforeImg = null;
            let afterImg = null;
            try {
                beforeImg = validateSafeJpegDataUri(
                    res.before_preview_base64 || (currentFileInfo && currentFileInfo.preview_base64)
                );
            } catch (_err) {
                beforeImg = null;
            }
            try {
                afterImg = validateSafeJpegDataUri(res.after_preview_base64) || beforeImg;
            } catch (_err) {
                afterImg = beforeImg;
            }

            if (beforeImg || afterImg) {
                previewComparisonPanel.classList.remove('hidden');
                previewBeforeSizeTag.textContent = `${origMb} MB`;
                previewAfterSizeTag.textContent = `${compMb} MB`;
                if (beforeImg) beforePreviewImg.src = beforeImg;
                if (afterImg) afterPreviewImg.src = afterImg;
            } else {
                previewComparisonPanel.classList.add('hidden');
            }
        }

        resultSection.scrollIntoView({ behavior: 'smooth', block: 'start' });
    }

    // -----------------------------------------------------------------------
    // Image Zoom Modal
    // -----------------------------------------------------------------------

    function openZoomModal(src, title) {
        if (!src) return;
        modalTitle.textContent = title;
        modalZoomImg.src = src;
        imageModal.classList.remove('hidden');
    }

    beforePreviewImg.addEventListener('click', () => {
        openZoomModal(beforePreviewImg.src, 'Bản gốc (Trước khi nén) - Trang 1');
    });

    afterPreviewImg.addEventListener('click', () => {
        openZoomModal(afterPreviewImg.src, 'Bản đã nén (Sau khi nén) - Trang 1');
    });

    if (extractSamplePreviewImg) {
        extractSamplePreviewImg.addEventListener('click', () => {
            openZoomModal(extractSamplePreviewImg.src, 'Ảnh mẫu trích xuất trong gói .ZIP');
        });
    }

    closeModalBtn.addEventListener('click', () => {
        imageModal.classList.add('hidden');
    });

    imageModal.addEventListener('click', (e) => {
        if (e.target === imageModal) {
            imageModal.classList.add('hidden');
        }
    });

    // -----------------------------------------------------------------------
    // Local History & JSON Backup Drawer (100% Safe DOM Rendering, Zero XSS)
    // -----------------------------------------------------------------------

    function showHistoryFeedback(msg, isError = false) {
        if (!historyFeedbackBanner) return;
        historyFeedbackBanner.textContent = msg;
        historyFeedbackBanner.classList.toggle('is-error', Boolean(isError));
        historyFeedbackBanner.classList.remove('hidden');
    }

    function hideHistoryFeedback() {
        if (historyFeedbackBanner) {
            historyFeedbackBanner.classList.add('hidden');
        }
    }

    function formatTimestampVi(ts) {
        try {
            return new Date(ts).toLocaleString('vi-VN', {
                day: '2-digit',
                month: '2-digit',
                year: 'numeric',
                hour: '2-digit',
                minute: '2-digit',
            });
        } catch (_err) {
            return '';
        }
    }

    function buildHistoryItemDom(item) {
        const row = document.createElement('div');
        row.className = 'history-item';

        const thumb = document.createElement('div');
        thumb.className = 'history-thumb';
        let safeThumb = null;
        try {
            safeThumb = validateSafeJpegDataUri(item.previewBase64);
        } catch (_err) {
            safeThumb = null;
        }
        if (safeThumb) {
            const img = document.createElement('img');
            img.src = safeThumb;
            img.alt = 'Ảnh xem trước';
            img.addEventListener('click', () => {
                openZoomModal(safeThumb, item.filename);
            });
            thumb.appendChild(img);
        } else {
            thumb.textContent = String(item.fileType || 'FILE').toUpperCase();
        }
        row.appendChild(thumb);

        const info = document.createElement('div');
        info.className = 'history-info';

        const titleLine = document.createElement('div');
        titleLine.className = 'history-title-line';

        const ftBadge = document.createElement('span');
        const ft = item.fileType === 'pptx' ? 'pptx' : item.fileType === 'ppt' ? 'ppt' : 'pdf';
        ftBadge.className = `format-badge ${ft}-badge`;
        ftBadge.textContent = ft.toUpperCase();
        titleLine.appendChild(ftBadge);

        const fnameEl = document.createElement('span');
        fnameEl.className = 'history-filename';
        fnameEl.textContent = item.filename;
        titleLine.appendChild(fnameEl);

        info.appendChild(titleLine);

        const metaLine = document.createElement('div');
        metaLine.className = 'history-meta-line';

        const actionSpan = document.createElement('span');
        if (item.jobType === 'extract') {
            actionSpan.textContent = `📦 Giải nén .ZIP (${item.itemCount} mục • ${Number(item.compressedSizeMb).toFixed(2)} MB)`;
        } else {
            actionSpan.textContent = `🗜️ ${Number(item.originalSizeMb).toFixed(2)} MB ➔ ${Number(item.compressedSizeMb).toFixed(2)} MB ( xuất .${String(item.outputFormat || 'pdf').toUpperCase()} )`;
        }
        metaLine.appendChild(actionSpan);

        const dateSpan = document.createElement('span');
        dateSpan.textContent = `🕒 ${formatTimestampVi(item.createdAt)}`;
        metaLine.appendChild(dateSpan);

        info.appendChild(metaLine);
        row.appendChild(info);

        const badge = document.createElement('span');
        if (item.jobType === 'extract') {
            badge.className = 'history-badge';
            badge.textContent = `${item.itemCount} mục`;
        } else if (item.targetMet === false) {
            badge.className = 'history-badge warn';
            badge.textContent = `Giảm ${Number(item.reductionPercent).toFixed(1)}%`;
        } else {
            badge.className = 'history-badge';
            const red = Number(item.reductionPercent);
            badge.textContent = red > 0 ? `-${red.toFixed(1)}%` : '0%';
        }
        row.appendChild(badge);

        return row;
    }

    async function refreshHistoryBadgeAndList() {
        const items = await listHistoryEntries();
        if (historyCountBadge) {
            historyCountBadge.textContent = String(items.length);
        }
        if (!historyListContainer || !historyEmptyState) return;

        historyListContainer.replaceChildren();
        if (items.length === 0) {
            historyEmptyState.classList.remove('hidden');
        } else {
            historyEmptyState.classList.add('hidden');
            const frag = document.createDocumentFragment();
            for (const item of items) {
                frag.appendChild(buildHistoryItemDom(item));
            }
            historyListContainer.appendChild(frag);
        }
    }

    if (openHistoryBtn) {
        openHistoryBtn.addEventListener('click', async () => {
            hideHistoryFeedback();
            confirmClearBar.classList.add('hidden');
            await refreshHistoryBadgeAndList();
            historyModal.classList.remove('hidden');
        });
    }

    if (closeHistoryModalBtn) {
        closeHistoryModalBtn.addEventListener('click', () => {
            historyModal.classList.add('hidden');
        });
    }

    if (historyModal) {
        historyModal.addEventListener('click', (e) => {
            if (e.target === historyModal) {
                historyModal.classList.add('hidden');
            }
        });
    }

    if (exportDataBtn) {
        exportDataBtn.addEventListener('click', async () => {
            hideHistoryFeedback();
            try {
                const jsonStr = await exportDatabaseToJson();
                const blob = new Blob([jsonStr], { type: 'application/json;charset=utf-8' });
                const tempUrl = URL.createObjectURL(blob);
                const link = document.createElement('a');
                const dateTag = new Date().toISOString().slice(0, 10);
                link.href = tempUrl;
                link.download = `slide-pdf-compressor-backup-${dateTag}.json`;
                document.body.appendChild(link);
                link.click();
                link.remove();
                setTimeout(() => URL.revokeObjectURL(tempUrl), 1000);
                showHistoryFeedback('Đã xuất tệp sao lưu .json thành công!');
            } catch (err) {
                showHistoryFeedback(err.message || 'Không thể xuất dữ liệu sao lưu.', true);
            }
        });
    }

    if (importDataBtn && importDataInput) {
        importDataBtn.addEventListener('click', () => {
            hideHistoryFeedback();
            importDataInput.value = '';
            importDataInput.click();
        });

        importDataInput.addEventListener('change', async () => {
            const file = importDataInput.files && importDataInput.files[0];
            if (!file) return;
            hideHistoryFeedback();

            if (file.size > MAX_BACKUP_BYTES) {
                showHistoryFeedback('Tệp sao lưu vượt giới hạn an toàn tối đa 5 MB.', true);
                return;
            }

            try {
                const rawText = await file.text();
                const res = await importDatabaseFromJson(rawText);
                applyPreferencesToControls(res.preferences);
                await refreshHistoryBadgeAndList();
                showHistoryFeedback(`Đã khôi phục thành công ${res.importedCount} bản ghi lịch sử & tùy chọn!`);
            } catch (err) {
                showHistoryFeedback(err.message || 'Tệp sao lưu không hợp lệ.', true);
            } finally {
                importDataInput.value = '';
            }
        });
    }

    if (clearDataBtn && confirmClearBar) {
        clearDataBtn.addEventListener('click', () => {
            hideHistoryFeedback();
            confirmClearBar.classList.remove('hidden');
        });

        confirmClearNoBtn.addEventListener('click', () => {
            confirmClearBar.classList.add('hidden');
        });

        confirmClearYesBtn.addEventListener('click', async () => {
            confirmClearBar.classList.add('hidden');
            try {
                await clearLocalDatabase();
                const defaults = await loadPreferences();
                applyPreferencesToControls(defaults);
                await refreshHistoryBadgeAndList();
                showHistoryFeedback('Đã xoá sạch toàn bộ lịch sử & dữ liệu cục bộ trên thiết bị.');
            } catch (err) {
                showHistoryFeedback(err.message || 'Không thể xoá dữ liệu cục bộ.', true);
            }
        });
    }

    // -----------------------------------------------------------------------
    // Bootstrap IndexedDB & Register Service Worker
    // -----------------------------------------------------------------------

    (async () => {
        await requestPersistentStorage();
        const prefs = await loadPreferences();
        applyPreferencesToControls(prefs);
        await refreshHistoryBadgeAndList();
    })();

    if ('serviceWorker' in navigator) {
        window.addEventListener('load', () => {
            navigator.serviceWorker.register('./sw.js').catch(() => {
                // Ignore SW registration failure on file:// protocol
            });
        });
    }
}

if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initApp);
} else {
    initApp();
}
