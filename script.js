let excelData = [];
let columns = [];
let batchIndex = -1;
let isBatchActive = false;
let isAutoAdvance = false;
let isPaused = false;
let advanceTimer = null;
let countdownInterval = null;

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
        excelData = XLSX.utils.sheet_to_json(worksheet).map(row => ({
            ...row,
            status: 'Pending'
        }));

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
    tableHeader.innerHTML = '<th>#</th><th>Status</th>';
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
        if (isBatchActive && index === batchIndex) {
            tr.classList.add('row-active');
        }

        // Row Number
        const tdIdx = document.createElement('td');
        tdIdx.setAttribute('data-label', '#');
        tdIdx.textContent = index + 1;
        tr.appendChild(tdIdx);

        // Status Badge
        const tdStatus = document.createElement('td');
        tdStatus.setAttribute('data-label', 'Status');
        const statusClass = `badge-${row.status.toLowerCase()}`;
        tdStatus.innerHTML = `<span class="status-badge ${statusClass}">${row.status}</span>`;
        tr.appendChild(tdStatus);

        // Data Columns
        columns.forEach(col => {
            const td = document.createElement('td');
            td.setAttribute('data-label', col);
            td.textContent = row[col] || '';
            tr.appendChild(td);
        });

        // WhatsApp Action
        const tdAction = document.createElement('td');
        tdAction.setAttribute('data-label', 'Action');
        const waLink = generateWhatsAppLink(row, template, phoneCol);

        if (waLink) {
            const a = document.createElement('a');
            a.href = waLink;
            a.target = '_blank';
            a.className = 'btn-send';
            a.innerHTML = '<i data-lucide="send"></i> Send';
            tdAction.appendChild(a);
        } else {
            tdAction.innerHTML = '<span style="color: var(--text-secondary); font-size: 0.8rem;">Invalid / No Phone</span>';
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

    // Basic Validation: must be at least 7 digits
    if (phone.length < 7) return null;

    // Simple template replacement
    let message = template || "Hello!";
    columns.forEach(col => {
        const placeholder = `{${col}}`;
        message = message.replace(new RegExp(placeholder, 'g'), row[col] || '');
    });

    // Use web.whatsapp.com to attempt forcing the browser version
    return `https://web.whatsapp.com/send?phone=${phone}&text=${encodeURIComponent(message)}`;
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

function toggleAutoAdvance() {
    isAutoAdvance = !isAutoAdvance;
    const ctrl = document.getElementById('auto-advance-ctrl');
    const timerDisplay = document.getElementById('timer-display');
    const helperText = document.getElementById('batch-helper-text');
    const pauseBtn = document.getElementById('pause-btn');

    if (isAutoAdvance) {
        ctrl.classList.add('auto-advance-active');
        timerDisplay.classList.remove('hidden');
        pauseBtn.classList.remove('hidden');
        helperText.textContent = "Will automatically open the next contact after the delay.";
        if (isBatchActive && !isPaused) startAutoAdvanceTimer();
    } else {
        ctrl.classList.remove('auto-advance-active');
        timerDisplay.classList.add('hidden');
        pauseBtn.classList.add('hidden');
        helperText.textContent = "Clicks 'Send Next' after completing the WhatsApp message.";
        clearTimers();
    }
}

function togglePauseBatch() {
    if (!isBatchActive) return;
    isPaused = !isPaused;

    const overlay = document.getElementById('batch-overlay');
    const pauseBtn = document.getElementById('pause-btn');

    if (isPaused) {
        overlay.classList.add('paused');
        pauseBtn.innerHTML = '<i data-lucide="play"></i> Resume';
        clearTimers();
    } else {
        overlay.classList.remove('paused');
        pauseBtn.innerHTML = '<i data-lucide="pause"></i> Pause';
        if (isAutoAdvance) startAutoAdvanceTimer();
    }
    lucide.createIcons();
}

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
    let message = `This will start a batch for ${count} contacts. Ready?`;
    if (isAutoAdvance) {
        message = `This will start an AUTOMATIC campaign for ${count} contacts. It will open tabs every few seconds. Ready?`;
    }

    if (confirm(message)) {
        batchIndex = 0;
        isBatchActive = true;
        isPaused = false;
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
    document.getElementById('batch-current-name').innerHTML = `Processing: <strong>${row[columns[0]] || 'Select Contact'}</strong>`;

    if (url) {
        row.status = 'Drafted';
        updateTableBody(); // Highlight row & update badge
        window.open(url, 'WA_BATCH_TAB');

        if (isAutoAdvance && !isPaused) {
            startAutoAdvanceTimer();
        }
    } else {
        console.warn(`Invalid or no phone found for row ${batchIndex + 1}. Skipping...`);
        row.status = 'Invalid';
        updateTableBody();

        // Brief delay before skipping automatically if in auto-advance mode
        if (isAutoAdvance && !isPaused) {
            setTimeout(nextInBatch, 500);
        } else if (!isAutoAdvance) {
            // If manual, just sit there so the user sees the 'Invalid' status
        }
    }
}

function startAutoAdvanceTimer() {
    clearTimers();
    let delay = parseInt(document.getElementById('advance-delay').value) || 5;
    let remaining = delay;

    document.getElementById('timer-seconds').textContent = remaining;

    countdownInterval = setInterval(() => {
        remaining--;
        document.getElementById('timer-seconds').textContent = remaining;
        if (remaining <= 0) {
            clearInterval(countdownInterval);
            nextInBatch();
        }
    }, 1000);
}

function clearTimers() {
    if (countdownInterval) clearInterval(countdownInterval);
    if (advanceTimer) clearTimeout(advanceTimer);
}

function nextInBatch() {
    clearTimers();
    batchIndex++;
    processBatchStep();
}

function stopBatch() {
    isBatchActive = false;
    isPaused = false;
    batchIndex = -1;
    clearTimers();
    document.getElementById('batch-overlay').classList.add('hidden');
    document.getElementById('batch-overlay').classList.remove('paused');
    updateTableBody();
}

generateBtn.addEventListener('click', startBatch);
