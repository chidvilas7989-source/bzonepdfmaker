document.addEventListener('DOMContentLoaded', () => {
    // Core DOM Elements
    const uploadStatus = document.getElementById('upload-status');
    const statusText = document.getElementById('status-text');
    const resultSection = document.getElementById('result-section');
    const previewText = document.getElementById('preview-text');
    const resultEditor = document.getElementById('result-editor');
    const resultTabsContainer = document.getElementById('result-tabs-container');
    const editorInstructionsMsg = document.getElementById('editor-instructions-msg');
    const inputContainer = document.getElementById('input-container');
    const btnResetWorkflow = document.getElementById('btn-reset-workflow');
    
    // Download Buttons
    const btnDownloadPrimary = document.getElementById('btn-download-primary');
    const primaryDlText = document.getElementById('primary-dl-text');
    const primaryDlIcon = document.getElementById('primary-dl-icon');
    const btnDownloadPdf = document.getElementById('btn-download-pdf');
    const btnDownloadDocx = document.getElementById('btn-download-docx');
    const btnDownloadMd = document.getElementById('btn-download-md');

    // Global states
    let currentResult = null;
    let selectedOrgFiles = [];
    let activeTab = 'ocr';

    // --- Tab Navigation Logic ---
    const tabButtons = document.querySelectorAll('.tab-btn');
    const workspacePanes = document.querySelectorAll('.workspace-pane');

    tabButtons.forEach(button => {
        button.addEventListener('click', () => {
            tabButtons.forEach(btn => btn.classList.remove('active'));
            workspacePanes.forEach(pane => pane.classList.add('hidden'));

            button.classList.add('active');
            activeTab = button.getAttribute('data-tab');
            document.getElementById(`${activeTab}-workspace`).classList.remove('hidden');

            // Show input container if it was hidden by loading/results
            inputContainer.classList.remove('hidden');

            // Reset result pane on tab switch
            resultSection.classList.add('hidden');
            currentResult = null;
            
            // Clean up organizer states
            selectedOrgFiles = [];
            renderSelectedFiles();
        });
    });

    // --- Tabbed Result Panel Logic (Edit vs Preview) ---
    const btnTabEdit = document.getElementById('btn-tab-edit');
    const btnTabPreview = document.getElementById('btn-tab-preview');
    const editorWrapper = document.getElementById('editor-wrapper');

    btnTabEdit.addEventListener('click', () => {
        btnTabEdit.classList.add('active');
        btnTabPreview.classList.remove('active');
        editorWrapper.classList.remove('hidden');
        previewText.classList.add('hidden');
    });

    btnTabPreview.addEventListener('click', () => {
        btnTabPreview.classList.add('active');
        btnTabEdit.classList.remove('active');
        editorWrapper.classList.add('hidden');
        previewText.classList.remove('hidden');

        // Render current edited markdown to HTML
        const mdText = resultEditor.value;
        if (mdText) {
            previewText.innerHTML = DOMPurify.sanitize(marked.parse(mdText));
        } else {
            previewText.innerHTML = '<em>No content available.</em>';
        }
    });

    // Get selected output format from the active workspace dropdown
    function getSelectedOutputFormat() {
        const selectEl = document.getElementById(`${activeTab}-format-select`);
        return selectEl ? selectEl.value : 'pdf';
    }

    // Configure the large primary download button based on user selection
    function configurePrimaryButton(format) {
        if (format === 'docx') {
            primaryDlText.innerText = 'Download Word Document (.docx)';
            primaryDlIcon.setAttribute('data-lucide', 'file-text');
        } else if (format === 'md') {
            primaryDlText.innerText = 'Download Text Document (.md)';
            primaryDlIcon.setAttribute('data-lucide', 'file-code');
        } else {
            primaryDlText.innerText = 'Download PDF Document (.pdf)';
            primaryDlIcon.setAttribute('data-lucide', 'file');
        }
        lucide.createIcons();
    }

    // --- Workspace 1: Handwriting OCR ---
    const ocrDropZone = document.getElementById('drop-zone');
    const ocrFileInput = document.getElementById('file-input');

    setupDragAndDrop(ocrDropZone, file => {
        const ext = file.name.split('.').pop().toLowerCase();
        const allowedExts = ['jpg', 'jpeg', 'png', 'webp', 'pdf', 'doc', 'docx', 'txt'];
        if (!file.type.match('image.*') && file.type !== 'application/pdf' && !allowedExts.includes(ext)) {
            alert('Please upload a valid file type: Images, PDFs, Word documents, or Text files.');
            return;
        }
        processOCR(file);
    });

    ocrFileInput.addEventListener('change', function() {
        if (this.files.length) {
            processOCR(this.files[0]);
        }
    });

    async function processOCR(file) {
        showLoading(`Uploading and analyzing ${file.name}...`);
        
        const docTypeSelect = document.getElementById('doc-type-select');
        const customInstructions = document.getElementById('custom-instructions');

        const formData = new FormData();
        formData.append('file', file);
        formData.append('doc_type', docTypeSelect.value);
        formData.append('custom_prompt', customInstructions.value);

        try {
            statusText.innerText = 'Transcribing content with AI...';
            const response = await fetch('/api/process', {
                method: 'POST',
                body: formData
            });

            if (!response.ok) {
                const errData = await response.json();
                throw new Error(errData.detail || 'OCR failed');
            }
            const data = await response.json();
            showResults(data);
        } catch (error) {
            handleError(error, 'Document processing failed.');
        }
    }

    // --- Workspace 2: PDF Organizer ---
    const orgDropZone = document.getElementById('org-drop-zone');
    const orgFileInput = document.getElementById('org-file-input');
    const orgActionSelect = document.getElementById('org-action-select');
    const splitSettingsGroup = document.getElementById('split-settings-group');
    const rotateSettingsGroup = document.getElementById('rotate-settings-group');
    const orgActionContainer = document.getElementById('org-action-container');
    const btnRunOrganizer = document.getElementById('btn-run-organizer');
    const selectedFilesList = document.getElementById('selected-files-list');

    orgActionSelect.addEventListener('change', () => {
        const action = orgActionSelect.value;
        splitSettingsGroup.style.display = action === 'split' ? 'block' : 'none';
        rotateSettingsGroup.style.display = action === 'rotate' ? 'block' : 'none';

        selectedOrgFiles = [];
        renderSelectedFiles();

        const uploadTitle = document.getElementById('org-upload-title');
        if (action === 'merge') {
            uploadTitle.innerText = "Drop PDF files here to combine";
            orgFileInput.multiple = true;
        } else {
            uploadTitle.innerText = `Drop a PDF here to ${action}`;
            orgFileInput.removeAttribute('multiple');
        }
    });

    setupDragAndDrop(orgDropZone, files => {
        const action = orgActionSelect.value;
        const fileList = files instanceof FileList ? Array.from(files) : [files];
        const pdfFiles = fileList.filter(f => f.name.toLowerCase().endsWith('.pdf'));
        
        if (pdfFiles.length === 0) {
            alert('Please select valid PDF files.');
            return;
        }

        if (action === 'merge') {
            selectedOrgFiles = selectedOrgFiles.concat(pdfFiles);
            renderSelectedFiles();
        } else {
            processOrganizer(pdfFiles[0]);
        }
    });

    orgFileInput.addEventListener('change', function() {
        if (this.files.length) {
            const action = orgActionSelect.value;
            const pdfFiles = Array.from(this.files).filter(f => f.name.toLowerCase().endsWith('.pdf'));
            
            if (action === 'merge') {
                selectedOrgFiles = selectedOrgFiles.concat(pdfFiles);
                renderSelectedFiles();
            } else {
                processOrganizer(pdfFiles[0]);
            }
        }
    });

    btnRunOrganizer.addEventListener('click', () => {
        if (orgActionSelect.value === 'merge') {
            if (selectedOrgFiles.length < 2) {
                alert('Please upload at least 2 PDFs to combine.');
                return;
            }
            processMerge();
        }
    });

    function renderSelectedFiles() {
        selectedFilesList.innerHTML = '';
        if (selectedOrgFiles.length > 0 && orgActionSelect.value === 'merge') {
            selectedFilesList.classList.remove('hidden');
            orgActionContainer.classList.remove('hidden');

            selectedOrgFiles.forEach((file, index) => {
                const item = document.createElement('div');
                item.className = 'selected-file-item';
                item.innerHTML = `
                    <span class="selected-file-name">
                        <i data-lucide="file" style="width:16px;height:16px;color:var(--secondary)"></i>
                        ${file.name}
                    </span>
                    <button class="selected-file-remove" data-index="${index}">
                        <i data-lucide="trash-2" style="width:14px;height:14px;"></i>
                    </button>
                `;
                selectedFilesList.appendChild(item);
            });

            document.querySelectorAll('.selected-file-remove').forEach(btn => {
                btn.addEventListener('click', (e) => {
                    e.stopPropagation();
                    const index = parseInt(btn.getAttribute('data-index'));
                    selectedOrgFiles.splice(index, 1);
                    renderSelectedFiles();
                });
            });

            lucide.createIcons();
        } else {
            selectedFilesList.classList.add('hidden');
            orgActionContainer.classList.add('hidden');
        }
    }

    async function processMerge() {
        showLoading('Combining PDF files...');
        const formData = new FormData();
        selectedOrgFiles.forEach(file => {
            formData.append('files', file);
        });

        try {
            const response = await fetch('/api/merge', {
                method: 'POST',
                body: formData
            });
            if (!response.ok) throw new Error('PDF merge failed');
            const data = await response.json();
            showResults({
                pdf_url: data.pdf_url,
                docx_url: data.docx_url,
                html_summary: `<h3>PDF Merge Successful!</h3><p>Successfully combined ${selectedOrgFiles.length} files into a single document.</p>`
            });
        } catch (error) {
            handleError(error, 'PDF merge failed.');
        }
    }

    async function processOrganizer(file) {
        const action = orgActionSelect.value;
        const formData = new FormData();
        formData.append('file', file);

        let endpoint = '';
        if (action === 'split') {
            const pagesInput = document.getElementById('split-pages-input').value;
            if (!pagesInput) {
                alert('Please enter page numbers or ranges (e.g. 1, 2-5).');
                return;
            }
            formData.append('pages_str', pagesInput);
            endpoint = '/api/split';
            showLoading(`Extracting pages from ${file.name}...`);
        } else if (action === 'rotate') {
            const deg = document.getElementById('rotate-deg-select').value;
            formData.append('degrees', deg);
            endpoint = '/api/rotate';
            showLoading(`Rotating pages of ${file.name} by ${deg}°...`);
        }

        try {
            const response = await fetch(endpoint, {
                method: 'POST',
                body: formData
            });
            if (!response.ok) {
                const errData = await response.json();
                throw new Error(errData.detail || 'Operation failed');
            }
            const data = await response.json();
            showResults({
                pdf_url: data.pdf_url,
                docx_url: data.docx_url,
                html_summary: `<h3>PDF operation complete!</h3><p>Your modified document is ready.</p>`
            });
        } catch (error) {
            handleError(error, `PDF ${action} failed.`);
        }
    }

    // --- Workspace 3: AI Document Suite ---
    const aiDropZone = document.getElementById('ai-drop-zone');
    const aiFileInput = document.getElementById('ai-file-input');
    const aiActionSelect = document.getElementById('ai-action-select');
    const translateLangGroup = document.getElementById('translate-lang-group');

    aiActionSelect.addEventListener('change', () => {
        translateLangGroup.style.display = aiActionSelect.value === 'translate' ? 'block' : 'none';
    });

    setupDragAndDrop(aiDropZone, file => {
        if (!file.name.toLowerCase().endsWith('.pdf')) {
            alert('Please select a valid PDF file.');
            return;
        }
        processAI(file);
    });

    aiFileInput.addEventListener('change', function() {
        if (this.files.length) {
            processAI(this.files[0]);
        }
    });

    async function processAI(file) {
        const action = aiActionSelect.value;
        const instructions = document.getElementById('ai-instructions').value;
        const formData = new FormData();
        formData.append('file', file);
        formData.append('custom_prompt', instructions);

        let endpoint = '';
        if (action === 'summarize') {
            endpoint = '/api/summarize';
            showLoading(`Generating summary for ${file.name}...`);
        } else if (action === 'translate') {
            const lang = document.getElementById('translate-lang-select').value;
            formData.append('target_language', lang);
            endpoint = '/api/translate';
            showLoading(`Translating ${file.name} to ${lang}...`);
        }

        try {
            const response = await fetch(endpoint, {
                method: 'POST',
                body: formData
            });
            if (!response.ok) {
                const errData = await response.json();
                throw new Error(errData.detail || 'AI Suite request failed');
            }
            const data = await response.json();
            showResults(data);
        } catch (error) {
            handleError(error, 'AI processing failed.');
        }
    }

    // --- Workspace 4: PDF Security ---
    const secDropZone = document.getElementById('sec-drop-zone');
    const secFileInput = document.getElementById('sec-file-input');
    const secActionSelect = document.getElementById('sec-action-select');

    setupDragAndDrop(secDropZone, file => {
        if (!file.name.toLowerCase().endsWith('.pdf')) {
            alert('Please select a valid PDF file.');
            return;
        }
        processSecurity(file);
    });

    secFileInput.addEventListener('change', function() {
        if (this.files.length) {
            processSecurity(this.files[0]);
        }
    });

    async function processSecurity(file) {
        const action = secActionSelect.value;
        const password = document.getElementById('sec-password-input').value;
        if (!password) {
            alert('Please enter a password.');
            return;
        }

        const formData = new FormData();
        formData.append('file', file);
        formData.append('password', password);

        const endpoint = action === 'protect' ? '/api/protect' : '/api/unlock';
        showLoading(`${action === 'protect' ? 'Encrypting' : 'Decrypting'} ${file.name}...`);

        try {
            const response = await fetch(endpoint, {
                method: 'POST',
                body: formData
            });
            if (!response.ok) {
                const errData = await response.json();
                throw new Error(errData.detail || 'Security action failed');
            }
            const data = await response.json();
            showResults({
                pdf_url: data.pdf_url,
                docx_url: data.docx_url,
                html_summary: `<h3>Security Configuration Success!</h3><p>Your PDF has been successfully ${action === 'protect' ? 'password-protected' : 'unlocked'}.</p>`
            });
        } catch (error) {
            handleError(error, `Security operation failed.`);
        }
    }

    // --- Drag & Drop Utilities ---
    function setupDragAndDrop(element, callback) {
        ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
            element.addEventListener(eventName, e => {
                e.preventDefault();
                e.stopPropagation();
            }, false);
        });

        ['dragenter', 'dragover'].forEach(eventName => {
            element.addEventListener(eventName, () => {
                element.classList.add('dragover');
            }, false);
        });

        ['dragleave', 'drop'].forEach(eventName => {
            element.addEventListener(eventName, () => {
                element.classList.remove('dragover');
            }, false);
        });

        element.addEventListener('drop', e => {
            const dt = e.dataTransfer;
            const files = dt.files;
            if (files.length) {
                callback(orgActionSelect && orgActionSelect.value === 'merge' && element === orgDropZone ? files : files[0]);
            }
        });
    }

    // --- UI Helper: Show Spinner ---
    function showLoading(text) {
        resultSection.classList.add('hidden');
        uploadStatus.classList.remove('hidden');
        statusText.innerText = text;
        uploadStatus.scrollIntoView({ behavior: 'smooth' });
    }

    // --- UI Helper: Display Results panel ---
    function showResults(data) {
        uploadStatus.classList.add('hidden');
        resultSection.classList.remove('hidden');
        currentResult = data;

        // Clear all input values
        ocrFileInput.value = '';
        orgFileInput.value = '';
        aiFileInput.value = '';
        secFileInput.value = '';
        document.getElementById('sec-password-input').value = '';

        // Configure based on selected format preference
        const preferredFormat = getSelectedOutputFormat();
        configurePrimaryButton(preferredFormat);

        // Show/hide editing panel based on output content type
        if (data.markdown) {
            resultTabsContainer.classList.remove('hidden');
            editorInstructionsMsg.classList.remove('hidden');
            editorWrapper.classList.remove('hidden');
            previewText.classList.add('hidden');
            
            btnTabEdit.classList.add('active');
            btnTabPreview.classList.remove('active');

            resultEditor.value = data.markdown;
            
            btnDownloadMd.style.display = 'inline-flex';
            btnDownloadPdf.style.display = 'inline-flex';
            btnDownloadDocx.style.display = 'inline-flex';
        } else {
            resultTabsContainer.classList.add('hidden');
            editorInstructionsMsg.classList.add('hidden');
            editorWrapper.classList.add('hidden');
            previewText.classList.remove('hidden');

            previewText.innerHTML = data.html_summary || '<em>Document processed successfully.</em>';
            
            btnDownloadMd.style.display = 'none';
            btnDownloadPdf.style.display = data.pdf_url ? 'inline-flex' : 'none';
            btnDownloadDocx.style.display = data.docx_url ? 'inline-flex' : 'none';
        }

        resultSection.scrollIntoView({ behavior: 'smooth' });
    }

    function handleError(error, message) {
        console.error('Error:', error);
        alert(message + ' ' + (error.message ? `\nDetails: ${error.message}` : ''));
        uploadStatus.classList.add('hidden');
        inputContainer.classList.remove('hidden');
    }

    // --- Document Generation & Exports (Handles Edited Text) ---
    async function triggerExport(format) {
        if (!currentResult) return;

        // If markdown is available, send the current text state to get the updated document
        if (currentResult.markdown) {
            const currentText = resultEditor.value;
            showLoading(`Generating your updated ${format.toUpperCase()}...`);
            
            const formData = new FormData();
            formData.append('markdown_text', currentText);
            formData.append('export_format', format);
            formData.append('document_title', 'Bzone Document');

            try {
                const response = await fetch('/api/export', {
                    method: 'POST',
                    body: formData
                });
                if (!response.ok) throw new Error('Export compilation failed.');
                const exportData = await response.json();
                
                uploadStatus.classList.add('hidden');
                resultSection.classList.remove('hidden');
                
                if (exportData.file_url) {
                    window.open(exportData.file_url, '_blank');
                }
            } catch (error) {
                handleError(error, `Failed to download ${format.toUpperCase()}.`);
            }
        } else {
            // For static formats (no markdown editable content), download the URL directly
            if (format === 'docx' && currentResult.docx_url) {
                window.open(currentResult.docx_url, '_blank');
            } else if (format === 'pdf' && currentResult.pdf_url) {
                window.open(currentResult.pdf_url, '_blank');
            } else {
                alert(`${format.toUpperCase()} file is not available for this operation.`);
            }
        }
    }

    // Handle large primary button action
    btnDownloadPrimary.addEventListener('click', () => {
        const preferredFormat = getSelectedOutputFormat();
        triggerExport(preferredFormat);
    });

    // Handle individual secondary buttons action
    btnDownloadPdf.addEventListener('click', () => triggerExport('pdf'));
    btnDownloadDocx.addEventListener('click', () => triggerExport('docx'));
    btnDownloadMd.addEventListener('click', () => {
        if (currentResult && currentResult.markdown) {
            // Client-side quick download for plain text markdown
            const textContent = resultEditor.value;
            const blob = new Blob([textContent], { type: 'text/markdown' });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = 'bzone_document.md';
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);
        } else {
            triggerExport('md');
        }
    });

    // Reset/Start Over workflow logic
    function resetWorkflow() {
        resultSection.classList.add('hidden');
        uploadStatus.classList.add('hidden');
        inputContainer.classList.remove('hidden');
        currentResult = null;
        
        // Clear all inputs
        ocrFileInput.value = '';
        orgFileInput.value = '';
        aiFileInput.value = '';
        secFileInput.value = '';
        document.getElementById('sec-password-input').value = '';
        selectedOrgFiles = [];
        renderSelectedFiles();
        
        // Scroll back to top workspace options
        document.querySelector('.app-tabs').scrollIntoView({ behavior: 'smooth' });
    }

    if (btnResetWorkflow) {
        btnResetWorkflow.addEventListener('click', resetWorkflow);
    }
});
