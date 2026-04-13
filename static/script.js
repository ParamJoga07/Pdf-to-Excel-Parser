const dropZone = document.getElementById('drop-zone');
const fileInput = document.getElementById('file-input');
const fileList = document.getElementById('file-list');
const convertBtn = document.getElementById('convert-btn');
const statusMessage = document.getElementById('status-message');

let selectedFiles = [];

// Handle click on drop zone
dropZone.addEventListener('click', () => fileInput.click());

// Handle file selection
fileInput.addEventListener('change', (e) => {
    handleFiles(e.target.files);
});

// Drag and drop handlers
dropZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropZone.classList.add('active');
});

dropZone.addEventListener('dragleave', () => {
    dropZone.classList.remove('active');
});

dropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropZone.classList.remove('active');
    handleFiles(e.dataTransfer.files);
});

function handleFiles(files) {
    const newFiles = Array.from(files).filter(file => file.type === 'application/pdf');
    
    if (newFiles.length === 0 && files.length > 0) {
        showStatus('Please upload valid PDF files.', 'error');
        return;
    }

    selectedFiles = [...selectedFiles, ...newFiles].slice(0, 100); // Limit to 100 files
    updateFileList();
    
    convertBtn.disabled = selectedFiles.length === 0;
    statusMessage.textContent = '';
}

function updateFileList() {
    fileList.innerHTML = '';
    selectedFiles.forEach((file, index) => {
        const item = document.createElement('div');
        item.className = 'file-item';
        item.innerHTML = `
            <span>${file.name} (${(file.size / 1024).toFixed(1)} KB)</span>
            <button onclick="removeFile(${index})" style="background:none; border:none; color:#ef4444; cursor:pointer; font-size:1.2rem;">&times;</button>
        `;
        fileList.appendChild(item);
    });
}

window.removeFile = (index) => {
    selectedFiles.splice(index, 1);
    updateFileList();
    convertBtn.disabled = selectedFiles.length === 0;
};

convertBtn.addEventListener('click', async () => {
    if (selectedFiles.length === 0) return;

    const formData = new FormData();
    selectedFiles.forEach(file => {
        formData.append('files', file);
    });

    setLoading(true);
    showStatus('Processing files... Please wait.', 'success');

    try {
        const response = await fetch('/convert', {
            method: 'POST',
            body: formData
        });

        if (!response.ok) {
            const errorData = await response.json();
            throw new Error(errorData.detail || 'Conversion failed');
        }

        const blob = await response.blob();
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = 'extracted_data.xlsx';
        document.body.appendChild(a);
        a.click();
        a.remove();
        window.URL.revokeObjectURL(url);

        showStatus('Successfully converted! Download started.', 'success');
        // Reset after success
        selectedFiles = [];
        updateFileList();
        convertBtn.disabled = true;
        fileInput.value = '';
    } catch (error) {
        showStatus(`Error: ${error.message}`, 'error');
    } finally {
        setLoading(false);
    }
});

function setLoading(isLoading) {
    if (isLoading) {
        convertBtn.classList.add('loading');
        convertBtn.disabled = true;
    } else {
        convertBtn.classList.remove('loading');
        convertBtn.disabled = selectedFiles.length === 0;
    }
}

function showStatus(message, type) {
    statusMessage.textContent = message;
    statusMessage.className = `status-message status-${type}`;
}
