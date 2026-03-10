// Template Creator

// Toast notification system
function showToast(message, type = 'error') {
    const container = document.getElementById('toastContainer');
    const toast = document.createElement('div');
    toast.className = `toast toast-${type}`;
    toast.textContent = message;

    container.appendChild(toast);

    // Auto-remove after 3 seconds
    setTimeout(() => {
        toast.classList.add('toast-fade-out');
        setTimeout(() => {
            toast.remove();
        }, 300);
    }, 3000);
}

// Status bar functions (no-op - status bar removed)
function showStatus(message, type = 'info') {
    // Status bar removed - do nothing
}

// Show compile errors/warnings as toast
function showCompileErrors(errors) {
    if (errors.length === 0) return;

    errors.forEach(err => {
        showToast(err, 'error');
    });
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
        showTemplateErrors(['Please define at least one column']);
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
    showToast('Template downloaded successfully', 'success');
}

function showTemplateErrors(errors) {
    if (errors.length === 0) return;

    errors.forEach(err => {
        showToast(err, 'error');
    });
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
            sheetNames: wb.SheetNames,
            workbook: wb
        };
    } catch {
        return { valid: false, sheets: 0, sheetNames: [], workbook: null };
    }
}

// Validate and Compile Section
let templateFile = null;
let reportFiles = [];
let compiledWorkbook = null;

const downloadBtn = document.getElementById('downloadBtn');

document.getElementById('templateInput').onchange = async (e) => setTemplate(e.target.files[0]);

async function setTemplate(file) {
    if (!file || !file.name.toLowerCase().endsWith('.xlsx')) {
        showToast('Template must be a .xlsx file', 'error');
        return;
    }

    try {
        const content = await file.arrayBuffer();
        const validation = validateExcelFile(content);

        if (!validation.valid) {
            showToast('Invalid Excel file', 'error');
            return;
        }

        templateFile = {
            name: file.name,
            content: content
        };
    } catch (err) {
        showToast('Failed to read file', 'error');
        return;
    }

    document.getElementById('templateInfo').innerHTML =
        `<b>${templateFile.name}</b> (${size(templateFile.content.byteLength)})`;
    document.getElementById('templateInfo').classList.remove('hidden');

    // Re-validate existing reports against the new template
    if (reportFiles.length > 0) {
        validateReportsAgainstTemplate();
    }
}


// Reports upload
document.getElementById('reportsInput').onchange = async (e) => {
    const file = e.target.files[0];
    if (!file) return;

    if (!file.name.toLowerCase().endsWith('.zip')) {
        showToast(`${file.name} is not a .zip file`, 'error');
        return;
    }

    // Clear previous reports
    reportFiles = [];

    try {
        const xlsxFiles = await extractXlsxFromZip(file);
        if (xlsxFiles.length === 0) {
            showToast(`No .xlsx files found in ${file.name}`, 'error');
            return;
        }

        for (const xlsxFile of xlsxFiles) {
            const validation = validateExcelFile(xlsxFile.content);
            reportFiles.push({
                file: xlsxFile,
                valid: validation.valid,
                sheets: validation.sheets,
                sheetNames: validation.sheetNames,
                workbook: validation.workbook,
                originalName: file.name
            });
        }
    } catch (err) {
        showToast(`Failed to read ${file.name}`, 'error');
        return;
    }

    renderReports();

    // If template exists, validate reports against it
    if (templateFile) {
        validateReportsAgainstTemplate();
    }
}

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
                    ${item.valid ? '✓ Valid' : `✗ ${item.validationError || 'Invalid'}`}
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
}

// Validate reports against template and update their status
function validateReportsAgainstTemplate() {
    if (!templateFile || reportFiles.length === 0) return;

    const templateInfo = getTemplateInfo();
    if (!templateInfo) return;

    reportFiles.forEach(report => {
        const validation = {
            valid: report.valid,
            sheets: report.sheets,
            sheetNames: report.sheetNames,
            workbook: report.workbook
        };

        // Check if template sheet name exists in report
        if (!validation.sheetNames.includes(templateInfo.sheetName)) {
            report.valid = false;
            report.validationError = `Sheet "${templateInfo.sheetName}" not found`;
            return;
        }

        // Check columns match
        const wb = validation.workbook;
        const ws = wb.Sheets[templateInfo.sheetName];
        const reportData = XLSX.utils.sheet_to_json(ws, { header: 1 });

        if (reportData.length === 0) {
            report.valid = false;
            report.validationError = 'Sheet is empty';
            return;
        }

        const reportColumns = reportData[0];
        const templateCols = templateInfo.columns;

        // Compare columns
        const templateColsStr = templateCols.map(c => String(c).trim()).join(',');
        const reportColsStr = reportColumns.map(c => String(c).trim()).join(',');

        if (templateColsStr !== reportColsStr) {
            report.valid = false;
            report.validationError = `Columns don't match template`;
            return;
        }

        // All checks passed
        report.valid = true;
        delete report.validationError;
    });

    renderReports();
}

// Get template sheet name and columns
function getTemplateInfo() {
    if (!templateFile) return null;

    const wb = XLSX.read(templateFile.content, { type: 'array' });
    const sheetName = wb.SheetNames[0];
    const ws = wb.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(ws, { header: 1 });

    if (data.length === 0) return null;

    return {
        sheetName: sheetName,
        columns: data[0], // First row as columns
        data: data
    };
}

// Compile & Download button
downloadBtn.onclick = async () => {
    const errors = [];

    // Check if template is selected
    if (!templateFile) {
        errors.push('Please upload a template file');
    }

    // Check if reports ZIP is selected
    if (reportFiles.length === 0) {
        errors.push('Please upload a reports ZIP file');
    }

    // If basic checks pass, validate file contents
    if (errors.length === 0 && templateFile && reportFiles.length > 0) {
        const templateInfo = getTemplateInfo();
        if (!templateInfo) {
            errors.push('Template file is empty or invalid');
        } else {
            // Check each report file
            for (const report of reportFiles) {
                const validation = {
                    valid: report.valid,
                    sheets: report.sheets,
                    sheetNames: report.sheetNames,
                    workbook: report.workbook
                };

                if (!validation.valid) {
                    errors.push(`${report.file.name}: Invalid Excel file`);
                    continue;
                }

                // Check if template sheet name exists in report
                if (!validation.sheetNames.includes(templateInfo.sheetName)) {
                    errors.push(`${report.file.name}: Sheet "${templateInfo.sheetName}" not found. Available sheets: ${validation.sheetNames.join(', ')}`);
                    continue;
                }

                // Check columns match
                const wb = validation.workbook;
                const ws = wb.Sheets[templateInfo.sheetName];
                const reportData = XLSX.utils.sheet_to_json(ws, { header: 1 });

                if (reportData.length === 0) {
                    errors.push(`${report.file.name}: Sheet is empty`);
                    continue;
                }

                const reportColumns = reportData[0];
                const templateCols = templateInfo.columns;

                // Compare columns
                const templateColsStr = templateCols.map(c => String(c).trim()).join(',');
                const reportColsStr = reportColumns.map(c => String(c).trim()).join(',');

                if (templateColsStr !== reportColsStr) {
                    errors.push(`${report.file.name}: Columns do not match template. Expected: "${templateColsStr}", Found: "${reportColsStr}"`);
                }
            }
        }
    }

    // Show errors if any
    if (errors.length > 0) {
        showCompileErrors(errors);
        return;
    }

    // Clear errors and proceed with compilation
    downloadBtn.disabled = true;
    document.getElementById('progressSection').classList.remove('hidden');
    showStatus('Compiling reports...');

    const bar = document.getElementById('progressFill');
    const templateInfo = getTemplateInfo();
    const templateWb = XLSX.read(templateFile.content, { type: 'array' });
    const newWb = XLSX.utils.book_new();

    // Copy template sheets (header only)
    const templateWs = templateWb.Sheets[templateInfo.sheetName];
    const allData = [templateInfo.columns]; // Start with header row

    // Append data from each report
    for (let i = 0; i < reportFiles.length; i++) {
        const report = reportFiles[i];
        const ws = report.workbook.Sheets[templateInfo.sheetName];
        const data = XLSX.utils.sheet_to_json(ws, { header: 1 });

        // Skip header row and append data rows
        for (let j = 1; j < data.length; j++) {
            allData.push(data[j]);
        }

        bar.style.width = ((i + 1) / reportFiles.length * 100) + '%';
    }

    // Create compiled sheet with all data
    const compiledWs = XLSX.utils.aoa_to_sheet(allData);
    XLSX.utils.book_append_sheet(newWb, compiledWs, templateInfo.sheetName);

    document.getElementById('progressSection').classList.add('hidden');

    // Download the compiled workbook
    XLSX.writeFile(newWb, 'compiled_reports.xlsx');
    showToast('Report compiled and downloaded successfully', 'success');
    downloadBtn.disabled = false;
};

function size(bytes) {
    const k = 1024;
    const units = ['B', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return (bytes / Math.pow(k, i)).toFixed(1) + ' ' + units[i];
}
