document.addEventListener('DOMContentLoaded', () => {
    // State
    let currentFileInfo = null;
    let currentRawFile = null;
    let currentBlobUrl = null;
    let activeActionTab = 'compress';
    let pollTimer = null;

    // Elements
    const htmlRoot = document.documentElement;
    const themeToggleBtn = document.getElementById('themeToggleBtn');
    const themeIcon = document.getElementById('themeIcon');
    const themeLabel = document.getElementById('themeLabel');

    const uploadSection = document.getElementById('uploadSection');
    const dropZone = document.getElementById('dropZone');
    const fileInput = document.getElementById('fileInput');
    const analyzeLoading = document.getElementById('analyzeLoading');
    const analyzeStatusText = document.getElementById('analyzeStatusText');

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

    // Action Mode Tabs
    const tabCompressBtn = document.getElementById('tabCompressBtn');
    const tabExtractBtn = document.getElementById('tabExtractBtn');

    // Compress Form Elements
    const pdfOutputOptions = document.getElementById('pdfOutputOptions');
    const pptxOutputOptions = document.getElementById('pptxOutputOptions');
    const pptxPdfOptionCard = document.getElementById('pptxPdfOptionCard');
    const pptxDirectOptionCard = document.getElementById('pptxDirectOptionCard');
    const pptLegacyCompressHint = document.getElementById('pptLegacyCompressHint');
    const modeConfigGroup = document.getElementById('modeConfigGroup');
    const targetPills = document.getElementById('targetPills');
    const targetMbInput = document.getElementById('targetMbInput');
    const compressForm = document.getElementById('compressForm');
    const startCompressBtn = document.getElementById('startCompressBtn');

    // Extract Form Elements
    const extractForm = document.getElementById('extractForm');
    const extractImagesCard = document.getElementById('extractImagesCard');
    const extractSlidesCard = document.getElementById('extractSlidesCard');
    const pptLegacyExtractHint = document.getElementById('pptLegacyExtractHint');
    const startExtractBtn = document.getElementById('startExtractBtn');

    // Progress Elements
    const progressSection = document.getElementById('progressSection');
    const progressStatusTitle = document.getElementById('progressStatusTitle');
    const progressDetailText = document.getElementById('progressDetailText');
    const progressPercentBadge = document.getElementById('progressPercentBadge');
    const progressBarFill = document.getElementById('progressBarFill');

    // Result Elements
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

    // Preview Elements
    const previewComparisonPanel = document.getElementById('previewComparisonPanel');
    const previewBeforeSizeTag = document.getElementById('previewBeforeSizeTag');
    const previewAfterSizeTag = document.getElementById('previewAfterSizeTag');
    const beforePreviewImg = document.getElementById('beforePreviewImg');
    const afterPreviewImg = document.getElementById('afterPreviewImg');

    const extractPreviewPanel = document.getElementById('extractPreviewPanel');
    const extractPreviewCountTag = document.getElementById('extractPreviewCountTag');
    const extractSamplePreviewImg = document.getElementById('extractSamplePreviewImg');

    // Error & Modal Elements
    const errorAlert = document.getElementById('errorAlert');
    const errorMessageText = document.getElementById('errorMessageText');
    const closeErrorBtn = document.getElementById('closeErrorBtn');

    const imageModal = document.getElementById('imageModal');
    const modalTitle = document.getElementById('modalTitle');
    const modalZoomImg = document.getElementById('modalZoomImg');
    const closeModalBtn = document.getElementById('closeModalBtn');

    // Theme Toggle
    themeToggleBtn.addEventListener('click', () => {
        const current = htmlRoot.getAttribute('data-theme') || 'dark';
        const next = current === 'dark' ? 'light' : 'dark';
        htmlRoot.setAttribute('data-theme', next);
        if (next === 'light') {
            themeIcon.textContent = '🌙';
            themeLabel.textContent = 'Giao diện Tối';
        } else {
            themeIcon.textContent = '☀️';
            themeLabel.textContent = 'Giao diện Sáng';
        }
    });

    function showError(msg) {
        errorMessageText.textContent = msg;
        errorAlert.classList.remove('hidden');
    }

    function hideError() {
        errorAlert.classList.add('hidden');
    }

    closeErrorBtn.addEventListener('click', hideError);

    // Switch between Compress and Extract Tabs
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

    // Drag & Drop Upload
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
            throw new Error('Dung lượng tệp vượt giới hạn tải trực tiếp trên máy chủ đám mây (~4.5 MB với gói Vercel miễn phí). Với tệp lớn tới 500 MB, vui lòng chạy bản máy cục bộ (run.bat).');
        }
        const text = await response.text();
        try {
            return JSON.parse(text);
        } catch (_err) {
            throw new Error(text ? text.slice(0, 180) : `Máy chủ phản hồi lỗi HTTP ${response.status}.`);
        }
    }

    async function reuploadCurrentFileIfNeeded() {
        if (!currentRawFile) return false;
        const formData = new FormData();
        formData.append('file', currentRawFile);
        const resp = await fetch('/api/analyze', { method: 'POST', body: formData });
        const data = await safeParseResponse(resp);
        if (!resp.ok) return false;
        currentFileInfo = data;
        return true;
    }

    async function handleFileUpload(file) {
        hideError();
        const dotIdx = file.name.lastIndexOf('.');
        const ext = dotIdx >= 0 ? file.name.slice(dotIdx).toLowerCase() : '';
        if (!['.pdf', '.pptx', '.ppt'].includes(ext)) {
            showError('Định dạng tệp không hợp lệ. Vui lòng chọn tệp .pdf, .pptx hoặc .ppt.');
            return;
        }

        analyzeLoading.classList.remove('hidden');
        analyzeStatusText.textContent = `Đang tải lên và phân tích "${file.name}"...`;

        const formData = new FormData();
        formData.append('file', file);

        try {
            const response = await fetch('/api/analyze', {
                method: 'POST',
                body: formData
            });
            const data = await safeParseResponse(response);
            if (!response.ok) {
                throw new Error(data.detail || 'Không thể phân tích tệp.');
            }
            currentRawFile = file;
            currentFileInfo = data;
            populateConfigPanel(data);
        } catch (err) {
            showError(err.message || 'Lỗi kết nối máy chủ khi phân tích tệp.');
        } finally {
            analyzeLoading.classList.add('hidden');
        }
    }

    function setRadioOptionState(cardEl, disabled) {
        if (!cardEl) return;
        const radio = cardEl.querySelector('input[type="radio"]');
        if (radio) {
            radio.disabled = disabled;
        }
        cardEl.classList.toggle('disabled', Boolean(disabled));
    }

    function selectRadioCard(cardEl) {
        if (!cardEl) return;
        const radio = cardEl.querySelector('input[type="radio"]');
        const group = cardEl.closest('.radio-card-group');
        if (group) {
            group.querySelectorAll('.radio-card').forEach(c => c.classList.remove('active'));
        }
        if (radio) {
            radio.checked = true;
        }
        cardEl.classList.add('active');
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

        // Reset legacy locks before applying file-type rules
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
            pageCountDisplay.textContent = count > 0 ? `${count} slide` : 'Tệp .PPT cũ';

            if (mediaCount > 0) {
                mediaCountDisplay.textContent = `${mediaCount} ảnh`;
                mediaCountPill.classList.remove('hidden');
            } else {
                mediaCountPill.classList.add('hidden');
            }

            pdfOutputOptions.classList.add('hidden');
            pptxOutputOptions.classList.remove('hidden');

            // Disable direct .PPTX output for legacy .ppt and auto-select PDF
            setRadioOptionState(pptxDirectOptionCard, true);
            selectRadioCard(pptxPdfOptionCard);
            pptLegacyCompressHint.classList.remove('hidden');
            updateModeVisibilityForPptx();

            // Disable direct embedded image extraction for legacy .ppt and auto-select slides
            setRadioOptionState(extractImagesCard, true);
            selectRadioCard(extractSlidesCard);
            pptLegacyExtractHint.classList.remove('hidden');
        } else {
            fileTypeBadge.textContent = 'PPTX';
            fileTypeBadge.className = 'format-badge pptx-badge';
            pageCountDisplay.textContent = `${count} slide`;

            if (info.media_count !== undefined) {
                mediaCountDisplay.textContent = `${mediaCount} ảnh`;
                mediaCountPill.classList.remove('hidden');
            } else {
                mediaCountPill.classList.add('hidden');
            }

            pdfOutputOptions.classList.add('hidden');
            pptxOutputOptions.classList.remove('hidden');
            updateModeVisibilityForPptx();

            selectRadioCard(extractImagesCard);
        }

        if (info.preview_base64) {
            initialPreviewImg.src = info.preview_base64;
            initialPreviewImg.classList.remove('hidden');
            initialThumbPlaceholder.classList.add('hidden');
        } else {
            initialPreviewImg.classList.add('hidden');
            initialThumbPlaceholder.classList.remove('hidden');
            initialThumbPlaceholder.textContent = (info.file_type || 'SLIDE').toUpperCase();
        }
    }

    function updateModeVisibilityForPptx() {
        const selectedPptxOut = document.querySelector('input[name="output_format_pptx"]:checked');
        if (selectedPptxOut && selectedPptxOut.value === 'pptx') {
            modeConfigGroup.classList.add('hidden');
        } else {
            modeConfigGroup.classList.remove('hidden');
        }
    }

    // Radio card visual state
    document.querySelectorAll('.radio-card input[type="radio"]').forEach((radio) => {
        radio.addEventListener('change', (e) => {
            if (e.target.disabled) return;
            const group = e.target.closest('.radio-card-group');
            if (group) {
                group.querySelectorAll('.radio-card').forEach(c => c.classList.remove('active'));
            }
            const card = e.target.closest('.radio-card');
            if (card) {
                card.classList.add('active');
            }
            if (e.target.name === 'output_format_pptx') {
                updateModeVisibilityForPptx();
            }
        });
    });

    // Target MB quick pills
    targetPills.querySelectorAll('.pill-btn').forEach((btn) => {
        btn.addEventListener('click', () => {
            targetPills.querySelectorAll('.pill-btn').forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            targetMbInput.value = btn.getAttribute('data-mb');
        });
    });

    targetMbInput.addEventListener('input', () => {
        const val = String(targetMbInput.value).trim();
        targetPills.querySelectorAll('.pill-btn').forEach((btn) => {
            if (btn.getAttribute('data-mb') === val) {
                btn.classList.add('active');
            } else {
                btn.classList.remove('active');
            }
        });
    });

    function reenableActionButtons() {
        startCompressBtn.disabled = false;
        if (startExtractBtn) {
            startExtractBtn.disabled = false;
        }
    }

    // Change file / Reset
    function resetToUpload() {
        if (pollTimer) {
            clearInterval(pollTimer);
            pollTimer = null;
        }
        if (currentBlobUrl) {
            URL.revokeObjectURL(currentBlobUrl);
            currentBlobUrl = null;
        }
        currentFileInfo = null;
        currentRawFile = null;
        fileInput.value = '';
        reenableActionButtons();
        hideError();
        configSection.classList.add('hidden');
        progressSection.classList.add('hidden');
        resultSection.classList.add('hidden');
        uploadSection.classList.remove('hidden');
    }

    changeFileBtn.addEventListener('click', resetToUpload);
    compressAnotherBtn.addEventListener('click', resetToUpload);

    // Start Compression
    compressForm.addEventListener('submit', async (e) => {
        e.preventDefault();
        if (!currentFileInfo) return;
        hideError();

        let outputFormat = 'pdf';
        const isLegacyPpt = Boolean(currentFileInfo.file_type === 'ppt' || currentFileInfo.is_legacy_ppt);
        if (currentFileInfo.file_type === 'pptx' && !isLegacyPpt) {
            const checkedOut = document.querySelector('input[name="output_format_pptx"]:checked');
            outputFormat = checkedOut ? checkedOut.value : 'pdf';
        }

        const checkedMode = document.querySelector('input[name="compression_mode"]:checked');
        const mode = checkedMode ? checkedMode.value : 'flatten';
        const targetMb = parseFloat(targetMbInput.value) || 10.0;

        startCompressBtn.disabled = true;
        if (startExtractBtn) startExtractBtn.disabled = true;
        progressSection.classList.remove('hidden');
        resultSection.classList.add('hidden');

        const totalUnits = outputFormat === 'pptx'
            ? (currentFileInfo.media_count || currentFileInfo.slide_count || 1)
            : (currentFileInfo.page_count || currentFileInfo.slide_count || 1);
        const initUnitLabel = outputFormat === 'pptx' ? 'mục' : 'trang';

        progressStatusTitle.textContent = `Đang khởi tạo xử lý 0/${totalUnits} ${initUnitLabel}...`;
        progressDetailText.textContent = 'Đang khởi tạo tiến trình nén thông minh...';
        progressPercentBadge.textContent = '0%';
        progressBarFill.style.width = '5%';

        try {
            let response = await fetch('/api/compress', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    file_id: currentFileInfo.file_id,
                    target_mb: targetMb,
                    mode: mode,
                    output_format: outputFormat
                })
            });
            if (response.status === 404 && await reuploadCurrentFileIfNeeded()) {
                response = await fetch('/api/compress', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({
                        file_id: currentFileInfo.file_id,
                        target_mb: targetMb,
                        mode: mode,
                        output_format: outputFormat
                    })
                });
            }
            const data = await safeParseResponse(response);
            if (!response.ok) {
                throw new Error(data.detail || 'Không thể bắt đầu nén tệp.');
            }
            if (data.status === 'done') {
                progressPercentBadge.textContent = '100%';
                progressBarFill.style.width = '100%';
                reenableActionButtons();
                renderResultDashboard(data);
            } else if (data.status === 'error') {
                throw new Error(data.error || data.message || 'Quá trình nén gặp sự cố.');
            } else {
                pollJobStatus(data.job_id, 'compress');
            }
        } catch (err) {
            showError(err.message || 'Lỗi khi gửi yêu cầu nén.');
            reenableActionButtons();
            progressSection.classList.add('hidden');
        }
    });

    // Start Resource Extraction (.ZIP)
    extractForm.addEventListener('submit', async (e) => {
        e.preventDefault();
        if (!currentFileInfo) return;
        hideError();

        const checkedExtract = document.querySelector('input[name="extract_mode"]:checked');
        const extractMode = checkedExtract ? checkedExtract.value : 'images';

        startExtractBtn.disabled = true;
        startCompressBtn.disabled = true;
        progressSection.classList.remove('hidden');
        resultSection.classList.add('hidden');

        const estTotal = extractMode === 'images' && currentFileInfo.media_count
            ? currentFileInfo.media_count
            : (currentFileInfo.page_count || currentFileInfo.slide_count || 1);

        progressStatusTitle.textContent = `Đang chuẩn bị trích xuất 0/${estTotal} mục...`;
        progressDetailText.textContent = 'Đang quét cấu trúc tệp và đóng gói tài nguyên .ZIP...';
        progressPercentBadge.textContent = '0%';
        progressBarFill.style.width = '5%';

        try {
            let response = await fetch('/api/extract', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    file_id: currentFileInfo.file_id,
                    extract_mode: extractMode
                })
            });
            if (response.status === 404 && await reuploadCurrentFileIfNeeded()) {
                response = await fetch('/api/extract', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({
                        file_id: currentFileInfo.file_id,
                        extract_mode: extractMode
                    })
                });
            }
            const data = await safeParseResponse(response);
            if (!response.ok) {
                throw new Error(data.detail || 'Không thể bắt đầu giải nén tài nguyên.');
            }
            if (data.status === 'done') {
                progressPercentBadge.textContent = '100%';
                progressBarFill.style.width = '100%';
                reenableActionButtons();
                renderResultDashboard(data);
            } else if (data.status === 'error') {
                throw new Error(data.error || data.message || 'Quá trình giải nén gặp sự cố.');
            } else {
                pollJobStatus(data.job_id, 'extract');
            }
        } catch (err) {
            showError(err.message || 'Lỗi khi gửi yêu cầu giải nén.');
            reenableActionButtons();
            progressSection.classList.add('hidden');
        }
    });

    function pollJobStatus(jobId, expectedJobType = 'compress') {
        if (pollTimer) clearInterval(pollTimer);

        pollTimer = setInterval(async () => {
            try {
                const res = await fetch(`/api/job/${jobId}`);
                const job = await res.json();
                if (!res.ok) {
                    throw new Error(job.detail || 'Không thể lấy trạng thái tiến trình.');
                }

                const jobType = job.job_type || expectedJobType;
                const cur = job.progress_current || 0;
                const tot = Math.max(job.progress_total || 1, 1);
                const pct = Math.min(Math.round((cur / tot) * 100), 100);

                const msg = job.message || '';
                const isItemBased = jobType === 'extract'
                    || job.output_format === 'zip'
                    || job.output_format === 'pptx'
                    || (/ảnh|media|tài nguyên|mục|slide/i.test(msg) && !/trang/i.test(msg));
                const unitLabel = isItemBased ? 'mục' : 'trang';

                progressStatusTitle.textContent = msg || `Đang xử lý ${cur}/${tot} ${unitLabel}...`;
                progressDetailText.textContent = `Tiến độ xử lý: ${cur}/${tot} ${unitLabel}`;
                progressPercentBadge.textContent = `${pct}%`;
                progressBarFill.style.width = `${Math.max(pct, 6)}%`;

                if (job.status === 'done') {
                    clearInterval(pollTimer);
                    pollTimer = null;
                    reenableActionButtons();
                    renderResultDashboard(job);
                } else if (job.status === 'error') {
                    clearInterval(pollTimer);
                    pollTimer = null;
                    reenableActionButtons();
                    progressSection.classList.add('hidden');
                    showError(job.error || job.message || 'Quá trình xử lý gặp sự cố.');
                }
            } catch (err) {
                clearInterval(pollTimer);
                pollTimer = null;
                reenableActionButtons();
                showError(err.message);
            }
        }, 350);
    }

    function formatZipSizeMb(res) {
        const rawMb = Number(res.zip_size_mb ?? res.compressed_size_mb ?? 0);
        if (rawMb > 0 && rawMb < 0.01) {
            return '< 0.01';
        }
        return rawMb.toFixed(2);
    }

    function renderResultDashboard(job) {
        progressSection.classList.add('hidden');
        resultSection.classList.remove('hidden');

        const res = job.result || job;
        const jobType = res.job_type || job.job_type || (res.output_format === 'zip' ? 'extract' : 'compress');
        const downloadUrl = res.download_url || `/api/download/${job.job_id}`;
        if (currentBlobUrl) {
            URL.revokeObjectURL(currentBlobUrl);
            currentBlobUrl = null;
        }
        if (res.download_data_b64) {
            try {
                const binStr = atob(res.download_data_b64);
                const bytes = new Uint8Array(binStr.length);
                for (let i = 0; i < binStr.length; i++) {
                    bytes[i] = binStr.charCodeAt(i);
                }
                const blob = new Blob([bytes], { type: res.download_mime || 'application/octet-stream' });
                currentBlobUrl = URL.createObjectURL(blob);
                downloadFileBtn.href = currentBlobUrl;
            } catch (_e) {
                downloadFileBtn.href = downloadUrl;
            }
        } else {
            downloadFileBtn.href = downloadUrl;
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
                ? 'Toàn bộ các trang / slide đã được tách thành bộ ảnh HD (.JPG) đóng gói trong tệp .ZIP.'
                : 'Toàn bộ hình ảnh nhúng đã được trích xuất nguyên bản và đóng gói trong tệp .ZIP.';

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
            const sampleImg = res.preview_base64 || (currentFileInfo && currentFileInfo.preview_base64);
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

            const origMb = Number(res.original_size_mb ?? (currentFileInfo && currentFileInfo.original_size_mb) ?? 0).toFixed(2);
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
                resultSubSummary.textContent = 'Tệp đã đạt chuẩn dung lượng mục tiêu với độ nét cao nhất.';
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

            const beforeImg = res.before_preview_base64 || (currentFileInfo && currentFileInfo.preview_base64);
            const afterImg = res.after_preview_base64 || beforeImg;

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

    // Image Zoom Modal
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
});
