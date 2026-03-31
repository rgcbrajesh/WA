let excelData = [];
let columns = [];
let batchIndex = -1;
let isBatchActive = false;

// DOM Elements
const dropZone = document.getElementById('drop-zone');
const fileInput = document.getElementById('file-input');
const fileStatus = document.getElementById('file-status');
const filenameDisplay = document.getElementById('filename-display');
const phoneSelect = document.getElementById('phone-column');
const variableSelect = document.getElementById('variable-select');
const messageTemplate = document.getElementById('message-template');
const tableHeader = document.getElementById('table-header');
const tableBody = document.getElementById('table-body');
const generateBtn = document.getElementById('generate-btn');

// --- File Handling ---

dropZone.addEventListener('click', () => fileInput.click());

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
    const files = e.dataTransfer.files;
    if (files.length) handleFile(files[0]);
});

fileInput.addEventListener('change', (e) => {
    if (e.target.files.length) handleFile(e.target.files[0]);
});

function handleFile(file) {
    const reader = new FileReader();
    reader.onload = (e) => {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        
        // Convert to JSON
        excelData = XLSX.utils.sheet_to_json(worksheet);
        
        if (excelData.length > 0) {
            columns = Object.keys(excelData[0]);
            updateUIAfterUpload(file.name);
            renderTable();
            populateColumnSelects();
            generateBtn.disabled = false;
        } else {
            alert('The excel file seems to be empty!');
        }
    };
    reader.readAsArrayBuffer(file);
}

function updateUIAfterUpload(filename) {
    fileStatus.classList.remove('hidden');
    filenameDisplay.textContent = filename;
    dropZone.querySelector('p').innerHTML = `File loaded: <span>${filename}</span>`;
}

// --- UI Rendering ---

function populateColumnSelects() {
    phoneSelect.innerHTML = '<option value="">Select phone column...</option>';
    variableSelect.innerHTML = '<option value="">Select variable...</option>';
    
    columns.forEach(col => {
        // Phone Select
        const phoneOpt = document.createElement('option');
        phoneOpt.value = col;
        phoneOpt.textContent = col;
        if (col.toLowerCase().includes('phone') || col.toLowerCase().includes('mobile') || col.toLowerCase().includes('whatsapp')) {
            phoneOpt.selected = true;
        }
        phoneSelect.appendChild(phoneOpt);

        // Variable Select
        const varOpt = document.createElement('option');
        varOpt.value = col;
        varOpt.textContent = col;
        variableSelect.appendChild(varOpt);
    });
}

function renderTable() {
    // Header
    tableHeader.innerHTML = '<th>#</th>';
    columns.forEach(col => {
        const th = document.createElement('th');
        th.textContent = col;
        tableHeader.appendChild(th);
    });
    tableHeader.innerHTML += '<th>Action</th>';

    // Body
    updateTableBody();
}

function updateTableBody() {
    tableBody.innerHTML = '';
    const template = messageTemplate.value;
    const phoneCol = phoneSelect.value;

    excelData.forEach((row, index) => {
        const tr = document.createElement('tr');
        
        // Row Number
        const tdIdx = document.createElement('td');
        tdIdx.textContent = index + 1;
        tr.appendChild(tdIdx);

        // Data Columns
        columns.forEach(col => {
            const td = document.createElement('td');
            td.textContent = row[col] || '';
            tr.appendChild(td);
        });

        // WhatsApp Action
        const tdAction = document.createElement('td');
        const waLink = generateWhatsAppLink(row, template, phoneCol);
        
        if (waLink) {
            const a = document.createElement('a');
            a.href = waLink;
            a.target = '_blank';
            a.className = 'btn-send';
            a.innerHTML = '<i data-lucide="send"></i> Send';
            tdAction.appendChild(a);
        } else {
            tdAction.innerHTML = '<span style="color: var(--text-secondary); font-size: 0.8rem;">No Phone</span>';
        }
        
        tr.appendChild(tdAction);
        tableBody.appendChild(tr);
    });
    
    // Re-initialize icons in the table
    lucide.createIcons();
}

// --- Logic ---

function generateWhatsAppLink(row, template, phoneCol) {
    let phone = phoneCol ? row[phoneCol] : null;
    if (!phone) return null;

    // Clean phone number (remove +, spaces, dashes)
    phone = String(phone).replace(/\D/g, '');
    
    // Simple template replacement
    let message = template || "Hello!";
    columns.forEach(col => {
        const placeholder = `{${col}}`;
        message = message.replace(new RegExp(placeholder, 'g'), row[col] || '');
    });

    return `https://wa.me/${phone}?text=${encodeURIComponent(message)}`;
}

function addVariableToTemplate() {
    const col = variableSelect.value;
    if (!col) return;

    const textarea = messageTemplate;
    const start = textarea.selectionStart;
    const end = textarea.selectionEnd;
    const text = textarea.value;
    const before = text.substring(0, start);
    const after = text.substring(end, text.length);
    const variable = `{${col}}`;

    textarea.value = before + variable + after;
    textarea.selectionStart = textarea.selectionEnd = start + variable.length;
    textarea.focus();

    // Trigger update
    updateTableBody();
}

// Real-time updates when template or phone column changes
messageTemplate.addEventListener('input', updateTableBody);
phoneSelect.addEventListener('change', updateTableBody);

function resetApp() {
    if (confirm('Are you sure you want to reset everything?')) {
        location.reload();
    }
}

// --- Batch Sending Logic ---

function startBatch() {
    if (!phoneSelect.value) {
        alert('Please select the column that contains phone numbers first!');
        return;
    }
    
    if (excelData.length === 0) {
        alert('Please upload an Excel file first.');
        return;
    }

    const count = excelData.length;
    if (confirm(`This will start a batch for ${count} contacts. Only ONE new tab will open, and you can clicked "Next" to send to the next person. Ready?`)) {
        batchIndex = 0;
        isBatchActive = true;
        document.getElementById('batch-overlay').classList.remove('hidden');
        processBatchStep();
    }
}

function processBatchStep() {
    if (batchIndex >= excelData.length) {
        alert('Campaign Complete! All messages have been drafted.');
        stopBatch();
        return;
    }

    const row = excelData[batchIndex];
    const template = messageTemplate.value;
    const phoneCol = phoneSelect.value;
    const url = generateWhatsAppLink(row, template, phoneCol);

    // Update UI
    document.getElementById('batch-progress').textContent = `Contact ${batchIndex + 1} of ${excelData.length}`;
    document.getElementById('batch-progress-inner').style.width = `${((batchIndex + 1) / excelData.length) * 100}%`;
    document.getElementById('batch-current-name').innerHTML = `Sending to: <strong>${row[columns[0]] || 'Select Contact'}</strong>`;

    if (url) {
        // OPEN or REFRESH the named window
        window.open(url, 'WA_BATCH_TAB');
    } else {
        alert(`No phone number found for row ${batchIndex + 1}. Skipping...`);
        nextInBatch();
    }
}

function nextInBatch() {
    batchIndex++;
    processBatchStep();
}

function stopBatch() {
    isBatchActive = false;
    batchIndex = -1;
    document.getElementById('batch-overlay').classList.add('hidden');
}

generateBtn.addEventListener('click', startBatch);
