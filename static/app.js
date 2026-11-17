document.addEventListener('DOMContentLoaded', () => {
    // State
    let currentFileInfo = null;
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

    const pdfOutputOptions = document.getElementById('pdfOutputOptions');
    const pptxOutputOptions = document.getElementById('pptxOutputOptions');
    const modeConfigGroup = document.getElementById('modeConfigGroup');
    const targetPills = document.getElementById('targetPills');
    const targetMbInput = document.getElementById('targetMbInput');
    const compressForm = document.getElementById('compressForm');
    const startCompressBtn = document.getElementById('startCompressBtn');

    const progressSection = document.getElementById('progressSection');
    const progressStatusTitle = document.getElementById('progressStatusTitle');
    const progressDetailText = document.getElementById('progressDetailText');
    const progressPercentBadge = document.getElementById('progressPercentBadge');
    const progressBarFill = document.getElementById('progressBarFill');

    const resultSection = document.getElementById('resultSection');
    const savedPercentBadge = document.getElementById('savedPercentBadge');
    const resBeforeMb = document.getElementById('resBeforeMb');
    const resAfterMb = document.getElementById('resAfterMb');
    const downloadFileBtn = document.getElementById('downloadFileBtn');
    const downloadFilenameLabel = document.getElementById('downloadFilenameLabel');
    const compressAnotherBtn = document.getElementById('compressAnotherBtn');

    const previewComparisonPanel = document.getElementById('previewComparisonPanel');
    const previewBeforeSizeTag = document.getElementById('previewBeforeSizeTag');
    const previewAfterSizeTag = document.getElementById('previewAfterSizeTag');
    const beforePreviewImg = document.getElementById('beforePreviewImg');
    const afterPreviewImg = document.getElementById('afterPreviewImg');

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

    async function handleFileUpload(file) {
        hideError();
        const ext = file.name.slice(file.name.lastIndexOf('.')).toLowerCase();
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
            const data = await response.json();
            if (!response.ok) {
                throw new Error(data.detail || 'Không thể phân tích tệp.');
            }
            currentFileInfo = data;
            populateConfigPanel(data);
        } catch (err) {
            showError(err.message || 'Lỗi kết nối máy chủ khi phân tích tệp.');
        } finally {
            analyzeLoading.classList.add('hidden');
        }
    }

    function populateConfigPanel(info) {
        uploadSection.classList.add('hidden');
        resultSection.classList.add('hidden');
        progressSection.classList.add('hidden');
        configSection.classList.remove('hidden');

        fileNameDisplay.textContent = info.filename;
        origSizeDisplay.textContent = `${Number(info.original_size_mb).toFixed(2)} MB`;

        const count = info.page_count || info.slide_count || 0;
        if (info.file_type === 'pdf') {
            fileTypeBadge.textContent = 'PDF';
            fileTypeBadge.className = 'format-badge pdf-badge';
            pageCountDisplay.textContent = `${count} trang`;
            mediaCountPill.classList.add('hidden');

            pdfOutputOptions.classList.remove('hidden');
            pptxOutputOptions.classList.add('hidden');
            modeConfigGroup.classList.remove('hidden');
        } else {
            fileTypeBadge.textContent = 'PPTX';
            fileTypeBadge.className = 'format-badge pptx-badge';
            pageCountDisplay.textContent = `${count} slide`;
            if (info.media_count !== undefined) {
                mediaCountDisplay.textContent = `${info.media_count} ảnh`;
                mediaCountPill.classList.remove('hidden');
            }

            pdfOutputOptions.classList.add('hidden');
            pptxOutputOptions.classList.remove('hidden');
            updateModeVisibilityForPptx();
        }

        if (info.preview_base64) {
            initialPreviewImg.src = info.preview_base64;
            initialPreviewImg.classList.remove('hidden');
            initialThumbPlaceholder.classList.add('hidden');
        } else {
            initialPreviewImg.classList.add('hidden');
            initialThumbPlaceholder.classList.remove('hidden');
            initialThumbPlaceholder.textContent = info.file_type.toUpperCase();
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

    // Change file / Reset
    function resetToUpload() {
        if (pollTimer) {
            clearInterval(pollTimer);
            pollTimer = null;
        }
        currentFileInfo = null;
        fileInput.value = '';
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
        if (currentFileInfo.file_type === 'pptx') {
            const checkedOut = document.querySelector('input[name="output_format_pptx"]:checked');
            outputFormat = checkedOut ? checkedOut.value : 'pdf';
        }

        const checkedMode = document.querySelector('input[name="compression_mode"]:checked');
        const mode = checkedMode ? checkedMode.value : 'flatten';
        const targetMb = parseFloat(targetMbInput.value) || 10.0;

        startCompressBtn.disabled = true;
        progressSection.classList.remove('hidden');
        resultSection.classList.add('hidden');

        const totalPages = currentFileInfo.page_count || currentFileInfo.slide_count || 1;
        progressStatusTitle.textContent = `Đang xử lý trang 0/${totalPages}...`;
        progressDetailText.textContent = 'Đang khởi tạo tiến trình nén thông minh...';
        progressPercentBadge.textContent = '0%';
        progressBarFill.style.width = '5%';

        try {
            const response = await fetch('/api/compress', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    file_id: currentFileInfo.file_id,
                    target_mb: targetMb,
                    mode: mode,
                    output_format: outputFormat
                })
            });
            const data = await response.json();
            if (!response.ok) {
                throw new Error(data.detail || 'Không thể bắt đầu nén tệp.');
            }
            pollJobStatus(data.job_id);
        } catch (err) {
            showError(err.message || 'Lỗi khi gửi yêu cầu nén.');
            startCompressBtn.disabled = false;
            progressSection.classList.add('hidden');
        }
    });

    function pollJobStatus(jobId) {
        if (pollTimer) clearInterval(pollTimer);

        pollTimer = setInterval(async () => {
            try {
                const res = await fetch(`/api/job/${jobId}`);
                const job = await res.json();
                if (!res.ok) {
                    throw new Error(job.detail || 'Không thể lấy trạng thái tiến trình.');
                }

                const cur = job.progress_current || 0;
                const tot = Math.max(job.progress_total || 1, 1);
                const pct = Math.min(Math.round((cur / tot) * 100), 100);

                progressStatusTitle.textContent = job.message || `Đang xử lý trang ${cur}/${tot}...`;
                progressDetailText.textContent = `Tiến độ xử lý: ${cur}/${tot} trang`;
                progressPercentBadge.textContent = `${pct}%`;
                progressBarFill.style.width = `${Math.max(pct, 6)}%`;

                if (job.status === 'done') {
                    clearInterval(pollTimer);
                    pollTimer = null;
                    startCompressBtn.disabled = false;
                    renderResultDashboard(job);
                } else if (job.status === 'error') {
                    clearInterval(pollTimer);
                    pollTimer = null;
                    startCompressBtn.disabled = false;
                    progressSection.classList.add('hidden');
                    showError(job.error || job.message || 'Quá trình nén gặp sự cố.');
                }
            } catch (err) {
                clearInterval(pollTimer);
                pollTimer = null;
                startCompressBtn.disabled = false;
                showError(err.message);
            }
        }, 350);
    }

    function renderResultDashboard(job) {
        progressSection.classList.add('hidden');
        resultSection.classList.remove('hidden');

        const res = job.result || job;
        const origMb = Number(res.original_size_mb ?? currentFileInfo.original_size_mb ?? 0).toFixed(2);
        const compMb = Number(res.compressed_size_mb ?? 0).toFixed(2);
        const savedPct = Number(res.reduction_percent ?? 0).toFixed(1);

        resBeforeMb.textContent = `${origMb} MB`;
        resAfterMb.textContent = `${compMb} MB`;
        savedPercentBadge.textContent = `Giảm ${savedPct}%`;

        const downloadUrl = res.download_url || `/api/download/${job.job_id}`;
        downloadFileBtn.href = downloadUrl;
        downloadFilenameLabel.textContent = res.output_filename ? `(${res.output_filename})` : '';

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

    closeModalBtn.addEventListener('click', () => {
        imageModal.classList.add('hidden');
    });

    imageModal.addEventListener('click', (e) => {
        if (e.target === imageModal) {
            imageModal.classList.add('hidden');
        }
    });
});
