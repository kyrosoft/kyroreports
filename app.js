// Template Creator

// Status bar functions
function showStatus(message, type = 'info') {
    const statusBar = document.getElementById('statusBar');
    const statusMessage = document.getElementById('statusMessage');

    statusBar.className = 'status-bar';
    if (type === 'error') {
        statusBar.classList.add('error');
    } else if (type === 'success') {
        statusBar.classList.add('success');
    }

    statusMessage.textContent = message;

    // Auto-clear after 5 seconds for info/success
    if (type !== 'error') {
        setTimeout(() => {
            showStatus('Ready');
        }, 5000);
    }
}

function parseColumns() {
    const input = document.getElementById('columnName');
    const value = input.value.trim();

    if (!value) return [];

    return value.split(',')
        .map(col => col.trim())
        .filter(col => col.length > 0);
}

function updatePreview() {
    const columns = parseColumns();
    const container = document.getElementById('columnsPreview');

    if (columns.length === 0) {
        container.innerHTML = '<p class="info-text-small">Type column names separated by commas to see preview</p>';
        return;
    }

    let html = '<table><thead><tr>';
    columns.forEach(col => {
        html += `<th>${col}</th>`;
    });
    html += '</tr></thead><tbody><tr>';
    columns.forEach(() => {
        html += '<td>[data]</td>';
    });
    html += '</tr></tbody></table>';

    // Show column count
    html += `<p class="info-text-small">${columns.length} column(s)</p>`;

    container.innerHTML = html;
}

// Initialize preview and add event listener
document.addEventListener('DOMContentLoaded', () => {
    const columnNameInput = document.getElementById('columnName');
    if (columnNameInput) {
        columnNameInput.addEventListener('input', updatePreview);
        // Initial render
        updatePreview();
    }
});

function createTemplate() {
    const columns = parseColumns();

    if (columns.length === 0) {
        showStatus('Add at least one column', 'error');
        return;
    }

    // Create workbook
    const wb = XLSX.utils.book_new();

    // Create data sheet with header row
    const data = [columns]; // Header row with column names

    const wsData = XLSX.utils.aoa_to_sheet(data);
    XLSX.utils.book_append_sheet(wb, wsData, 'KyroReports');

    // Download
    XLSX.writeFile(wb, 'Template - KyroReports.xlsx');
    showStatus('Template downloaded successfully', 'success');
}

// Utility to extract xlsx files from zip
async function extractXlsxFromZip(file) {
    const zip = new JSZip();
    const arrayBuffer = await file.arrayBuffer();
    const zipContent = await zip.loadAsync(arrayBuffer);
    const xlsxFiles = [];

    for (const [filename, zipEntry] of Object.entries(zipContent.files)) {
        if (!zipEntry.dir && filename.toLowerCase().endsWith('.xlsx')) {
            const content = await zipEntry.async('arraybuffer');
            xlsxFiles.push({
                name: filename,
                content: content,
                originalFile: file
            });
        }
    }

    return xlsxFiles;
}

// Validate Excel file
function validateExcelFile(arrayBuffer) {
    try {
        const wb = XLSX.read(arrayBuffer, { type: 'array' });
        return {
            valid: wb.SheetNames.length > 0,
            sheets: wb.SheetNames.length,
            sheetNames: wb.SheetNames
        };
    } catch {
        return { valid: false, sheets: 0, sheetNames: [] };
    }
}

// Validate and Compile Section
let templateFile = null;
let reportFiles = [];
let validationComplete = false;
let compiledWorkbook = null;

const validateBtn = document.getElementById('validateBtn');
const compileBtn = document.getElementById('compileBtn');
const downloadBtn = document.getElementById('downloadBtn');

document.getElementById('templateInput').onchange = async (e) => setTemplate(e.target.files[0]);

async function setTemplate(file) {
    if (!file || !file.name.toLowerCase().endsWith('.xlsx')) {
        showCompileError('Template must be a .xlsx file');
        return;
    }

    try {
        const content = await file.arrayBuffer();
        const validation = validateExcelFile(content);

        if (!validation.valid) {
            showCompileError('Invalid Excel file');
            return;
        }

        templateFile = {
            name: file.name,
            content: content
        };
    } catch (err) {
        showCompileError('Failed to read file');
        return;
    }

    document.getElementById('templateInfo').innerHTML =
        `<b>${templateFile.name}</b> (${size(templateFile.content.byteLength)})`;
    document.getElementById('templateInfo').classList.remove('hidden');
    clearCompileErrors();
    updateButtons();
}

function showCompileError(msg) {
    showStatus(msg, 'error');
    const el = document.getElementById('validationErrors');
    el.textContent = msg;
    el.classList.remove('hidden');
}

function clearCompileErrors() {
    showStatus('Ready');
    document.getElementById('validationErrors').classList.add('hidden');
}

// Reports upload
document.getElementById('reportsInput').onchange = async (e) => {
    const file = e.target.files[0];
    if (!file) return;

    if (!file.name.toLowerCase().endsWith('.zip')) {
        showCompileError(`${file.name} is not a .zip file`);
        return;
    }

    // Clear previous reports
    reportFiles = [];

    try {
        const xlsxFiles = await extractXlsxFromZip(file);
        if (xlsxFiles.length === 0) {
            showCompileError(`No .xlsx files found in ${file.name}`);
            return;
        }

        for (const xlsxFile of xlsxFiles) {
            const validation = validateExcelFile(xlsxFile.content);
            reportFiles.push({
                file: xlsxFile,
                valid: validation.valid,
                sheets: validation.sheets,
                sheetNames: validation.sheetNames,
                originalName: file.name
            });
        }
    } catch (err) {
        showCompileError(`Failed to read ${file.name}`);
        return;
    }

    renderReports();
    clearCompileErrors();

    // Reset validation state when new files are added
    validationComplete = false;
    document.getElementById('validationSummary').classList.add('hidden');
    updateButtons();
};

function renderReports() {
    const list = document.getElementById('reportsList');
    if (reportFiles.length === 0) {
        list.classList.add('hidden');
        return;
    }
    list.innerHTML = reportFiles.map((item, i) => `
        <div class="file-item ${item.valid ? 'valid' : 'invalid'}">
            <span>${item.file.name} (${size(item.file.content.byteLength)})
                <span class="file-status ${item.valid ? 'valid' : 'invalid'}">
                    ${item.valid ? '✓ Valid' : '✗ Invalid'}
                </span>
                ${item.originalName ? `<small>from ${item.originalName}</small>` : ''}
            </span>
            <button class="file-remove" onclick="removeReport(${i})">×</button>
        </div>
    `).join('');
    list.classList.remove('hidden');
}

function removeReport(i) {
    reportFiles.splice(i, 1);
    renderReports();

    // Reset validation state when files are removed
    validationComplete = false;
    document.getElementById('validationSummary').classList.add('hidden');
    updateButtons();
}

// Validate button
validateBtn.onclick = () => {
    const validCount = reportFiles.filter(r => r.valid).length;
    const invalidCount = reportFiles.length - validCount;

    let summary = `<div class="validation-summary"><h4>Validation Results</h4>`;
    summary += `<p><strong>Total Files:</strong> ${reportFiles.length}</p>`;
    summary += `<p style="color: #10b981;"><strong>Valid:</strong> ${validCount}</p>`;
    summary += `<p style="color: #ef4444;"><strong>Invalid:</strong> ${invalidCount}</p>`;

    if (invalidCount > 0) {
        summary += '<p><strong>Invalid Files:</strong></p><ul>';
        reportFiles.filter(f => !f.valid).forEach(f => {
            summary += `<li>${f.file.name}</li>`;
        });
        summary += '</ul>';
    }

    summary += '<p><strong>All Files:</strong></p><ul>';
    reportFiles.forEach(f => {
        summary += `<li>${f.file.name} - ${f.valid ? '✓' : '✗'} (${f.sheets} sheets)</li>`;
    });
    summary += '</ul></div>';

    document.getElementById('validationSummary').innerHTML = summary;
    document.getElementById('validationSummary').classList.remove('hidden');

    validationComplete = true;
    updateButtons();

    if (invalidCount > 0) {
        showStatus(`Validation complete: ${validCount} valid, ${invalidCount} invalid files`, 'error');
    } else {
        showStatus(`Validation complete: All ${validCount} files are valid`, 'success');
    }
};

// Compile button
compileBtn.onclick = async () => {
    if (!validationComplete) return;

    compileBtn.disabled = true;
    document.getElementById('progressSection').classList.remove('hidden');
    showStatus('Compiling reports...');

    const bar = document.getElementById('progressFill');
    const validReports = reportFiles.filter(r => r.valid);
    const total = validReports.length;

    // Read template
    const templateWb = XLSX.read(templateFile.content, { type: 'array' });
    const newWb = XLSX.utils.book_new();

    // Copy template sheets
    templateWb.SheetNames.forEach(name => {
        const ws = templateWb.Sheets[name];
        XLSX.utils.book_append_sheet(newWb, ws, name);
    });

    // Append data from each report
    for (let i = 0; i < total; i++) {
        const { file } = validReports[i];
        const wb = XLSX.read(file.content, { type: 'array' });

        // Add each sheet with report name prefix
        wb.SheetNames.forEach(name => {
            const ws = wb.Sheets[name];
            const baseName = file.name.replace('.xlsx', '');
            const newSheetName = `${baseName} - ${name}`;
            XLSX.utils.book_append_sheet(newWb, ws, newSheetName);
        });

        bar.style.width = ((i + 1) / total * 100) + '%';
    }

    compiledWorkbook = newWb;

    document.getElementById('progressSection').classList.add('hidden');
    compileBtn.disabled = false;
    updateButtons();
    showStatus(`Compilation complete: ${total} files compiled`, 'success');
};

// Download button
downloadBtn.onclick = () => {
    if (compiledWorkbook) {
        XLSX.writeFile(compiledWorkbook, 'compiled_reports.xlsx');
        showStatus('Report downloaded successfully', 'success');
    }
};

function size(bytes) {
    const k = 1024;
    const units = ['B', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return (bytes / Math.pow(k, i)).toFixed(1) + ' ' + units[i];
}

function updateButtons() {
    const hasTemplate = !!templateFile;
    const hasReports = reportFiles.length > 0;

    // Enable Validate when template and reports are loaded
    validateBtn.disabled = !(hasTemplate && hasReports);

    // Enable Compile when validation is complete
    compileBtn.disabled = !validationComplete;

    // Enable Download when compilation is complete
    downloadBtn.disabled = !compiledWorkbook;
}
