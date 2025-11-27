document.addEventListener('DOMContentLoaded', () => {
    const dropZone = document.getElementById('drop-zone');
    const fileInput = document.getElementById('file-input');
    const fileInfo = document.getElementById('file-info');
    const filenameDisplay = document.getElementById('filename');
    const removeFileBtn = document.getElementById('remove-file');
    const convertBtn = document.getElementById('convert-btn');
    const radioCards = document.querySelectorAll('.radio-card');
    const dpiSlider = document.getElementById('dpi-slider');
    const dpiValue = document.getElementById('dpi-value');
    const dpiGroup = document.getElementById('dpi-group');
    const loadingOverlay = document.getElementById('loading-overlay');

    let selectedFile = null;

    // Drag & Drop
    dropZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        dropZone.classList.add('dragover');
    });

    dropZone.addEventListener('dragleave', () => {
        dropZone.classList.remove('dragover');
    });

    dropZone.addEventListener('drop', (e) => {
        e.preventDefault();
        dropZone.classList.remove('dragover');
        if (e.dataTransfer.files.length) {
            handleFile(e.dataTransfer.files[0]);
        }
    });

    dropZone.addEventListener('click', () => fileInput.click());

    fileInput.addEventListener('change', (e) => {
        if (e.target.files.length) {
            handleFile(e.target.files[0]);
        }
    });

    function handleFile(file) {
        if (file.type !== 'application/pdf') {
            alert('Please upload a PDF file.');
            return;
        }
        selectedFile = file;
        filenameDisplay.textContent = file.name;
        dropZone.classList.add('hidden');
        fileInfo.classList.remove('hidden');
        convertBtn.disabled = false;
        document.getElementById('extract-text-btn').disabled = false;
    }

    removeFileBtn.addEventListener('click', () => {
        selectedFile = null;
        fileInput.value = '';
        dropZone.classList.remove('hidden');
        fileInfo.classList.add('hidden');
        convertBtn.disabled = true;
        document.getElementById('extract-text-btn').disabled = true;
    });

    // Settings
    radioCards.forEach(card => {
        card.addEventListener('click', () => {
            radioCards.forEach(c => c.classList.remove('selected'));
            card.classList.add('selected');
            const radio = card.querySelector('input');
            radio.checked = true;

            // Toggle DPI visibility
            if (radio.value === 'image') {
                dpiGroup.classList.remove('hidden');
            } else {
                dpiGroup.classList.add('hidden');
            }
        });
    });

    dpiSlider.addEventListener('input', (e) => {
        dpiValue.textContent = e.target.value;
    });

    // Conversion
    convertBtn.addEventListener('click', async () => {
        if (!selectedFile) return;

        const mode = document.querySelector('input[name="mode"]:checked').value;
        const dpi = dpiSlider.value;

        const formData = new FormData();
        formData.append('file', selectedFile);
        formData.append('mode', mode);
        formData.append('dpi', dpi);

        loadingOverlay.classList.remove('hidden');

        try {
            const response = await fetch('/convert', {
                method: 'POST',
                body: formData
            });

            if (!response.ok) {
                let errorMessage = 'Conversion failed';
                try {
                    const errorData = await response.json();
                    errorMessage = errorData.detail || errorMessage;
                } catch (e) {
                    // If response is not JSON, use status text
                    errorMessage = response.statusText || errorMessage;
                }
                throw new Error(errorMessage);
            }

            const blob = await response.blob();
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = 'converted_presentation.pptx';
            document.body.appendChild(a);
            a.click();
            window.URL.revokeObjectURL(url);
            a.remove();
        } catch (error) {
            alert('An error occurred during conversion: ' + error.message);
        } finally {
            loadingOverlay.classList.add('hidden');
        }
    });

    // Text Extraction
    const extractTextBtn = document.getElementById('extract-text-btn');

    extractTextBtn.addEventListener('click', async () => {
        if (!selectedFile) return;

        const formData = new FormData();
        formData.append('file', selectedFile);

        loadingOverlay.classList.remove('hidden');

        try {
            const response = await fetch('/extract_text', {
                method: 'POST',
                body: formData
            });

            if (!response.ok) {
                let errorMessage = 'Extraction failed';
                try {
                    const errorData = await response.json();
                    errorMessage = errorData.detail || errorMessage;
                } catch (e) {
                    // If response is not JSON, use status text
                    errorMessage = response.statusText || errorMessage;
                }
                throw new Error(errorMessage);
            }

            const blob = await response.blob();
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = 'extracted_text.txt';
            document.body.appendChild(a);
            a.click();
            window.URL.revokeObjectURL(url);
            a.remove();
        } catch (error) {
            alert('An error occurred during extraction: ' + error.message);
        } finally {
            loadingOverlay.classList.add('hidden');
        }
    });
});
