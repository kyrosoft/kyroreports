// Template Creator
let columns = [];

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


// Save pending input values before re-rendering
function savePendingEdits() {
    const inputs = document.querySelectorAll('.column-input');
    inputs.forEach((input, index) => {
        if (input && index < columns.length) {
            const value = input.value.trim();
            if (value) {
                columns[index] = value;
            }
        }
    });
}

function renderTemplate() {
    const container = document.getElementById('templatePreview');

    if (columns.length === 0) {
        // Always show at least one empty input
        columns.push('');
    }

    let html = '<div class="columns-container">';

    columns.forEach((col, index) => {
        // Calculate width based on text length (approx 8px per character + padding)
        const textLength = col.length;
        const calculatedWidth = Math.max(80, Math.min(200, textLength * 8 + 40));
        const widthStyle = `style="width: ${calculatedWidth}px;"`;

        html += `
            <div class="column-input-wrapper"
                 draggable="true"
                 data-index="${index}"
                 ondragstart="handleDragStart(event, ${index})"
                 ondragover="handleDragOver(event)"
                 ondragenter="handleDragEnter(event)"
                 ondragleave="handleDragLeave(event)"
                 ondrop="handleDrop(event, ${index})"
                 ondragend="handleDragEnd(event)">
                <input type="text"
                       value="${col}"
                       data-index="${index}"
                       placeholder="Column ${index + 1}..."
                       onchange="updateColumn(${index}, this.value)"
                       onblur="updateColumn(${index}, this.value)"
                       oninput="autoResizeInput(this)"
                       class="column-input-field"
                       ${widthStyle}>
                <button onclick="removeColumn(${index})" class="column-close-btn" title="Remove column">×</button>
            </div>`;
    });

    // Add the "+" button
    html += `
        <button onclick="addColumn()" class="add-column-small-btn" title="Add column">+</button>
    `;

    html += '</div>';

    // Show column count
    html += `<p class="info-text-small">${columns.length} column(s)</p>`;

    container.innerHTML = html;
}

// Auto-resize input based on content
function autoResizeInput(input) {
    const textLength = input.value.length;
    const calculatedWidth = Math.max(80, Math.min(200, textLength * 8 + 40));
    input.style.width = calculatedWidth + 'px';
}

// Drag and drop handlers
let draggedIndex = null;

function handleDragStart(event, index) {
    draggedIndex = index;
    event.target.classList.add('dragging');
    event.dataTransfer.effectAllowed = 'move';
    event.dataTransfer.setData('text/html', event.target.innerHTML);
}

function handleDragOver(event) {
    if (event.preventDefault) {
        event.preventDefault();
    }
    event.dataTransfer.dropEffect = 'move';
    return false;
}

function handleDragEnter(event) {
    const target = event.target.closest('.column-input-wrapper');
    if (target && draggedIndex !== null) {
        target.classList.add('drag-over');
    }
}

function handleDragLeave(event) {
    const target = event.target.closest('.column-input-wrapper');
    if (target) {
        target.classList.remove('drag-over');
    }
}

function handleDrop(event, dropIndex) {
    if (event.stopPropagation) {
        event.stopPropagation();
    }

    if (draggedIndex !== null && draggedIndex !== dropIndex) {
        // Save any pending edits before rearranging
        savePendingEdits();

        // Rearrange the columns array
        const draggedColumn = columns[draggedIndex];
        columns.splice(draggedIndex, 1);
        columns.splice(dropIndex, 0, draggedColumn);

        // Re-render
        renderTemplate();
        showToast('Column moved', 'success');
    }

    return false;
}

function handleDragEnd(event) {
    const wrappers = document.querySelectorAll('.column-input-wrapper');
    wrappers.forEach(wrapper => {
        wrapper.classList.remove('dragging');
        wrapper.classList.remove('drag-over');
    });
    draggedIndex = null;
}

function addColumn() {
    // Save any pending edits before re-rendering
    savePendingEdits();

    // Add an empty column
    columns.push('');
    renderTemplate();

    // Focus the new input
    const inputs = document.querySelectorAll('.column-input-field');
    if (inputs.length > 0) {
        inputs[inputs.length - 1].focus();
    }
}

function removeColumn(index) {
    if (columns.length <= 1) {
        showToast('You must have at least one column', 'error');
        return;
    }

    // Save any pending edits before re-rendering
    savePendingEdits();

    columns.splice(index, 1);
    renderTemplate();
}

function updateColumn(index, value) {
    // Allow empty values during editing, just trim and save
    columns[index] = value.trim();
}

// Initialize template on page load
document.addEventListener('DOMContentLoaded', () => {
    renderTemplate();
});

function createTemplate() {
    if (columns.length === 0) {
        showTemplateErrors(['Please add at least one column']);
        return;
    }

    // Create workbook
    const wb = XLSX.utils.book_new();

    // Create data sheet with header row
    const data = [columns]; // Header row with column names

    const wsData = XLSX.utils.aoa_to_sheet(data);

    // Apply styling to header row (bold, centered)
    const range = XLSX.utils.decode_range(wsData['!ref']);
    for (let col = range.s.c; col <= range.e.c; col++) {
        const cellAddress = XLSX.utils.encode_cell({ r: 0, c: col });
        if (!wsData[cellAddress]) continue;
        wsData[cellAddress].s = {
            font: { bold: true },
            alignment: { horizontal: 'center', vertical: 'center' }
        };
    }

    // Auto-fit column widths based on header content
    const colWidths = columns.map(col => ({
        wch: Math.max(col.length + 2, 10) // Minimum width of 10
    }));
    wsData['!cols'] = colWidths;

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

    // Copy template sheets (header only) with "No." as first column
    const headerRow = ['No.', ...templateInfo.columns];
    const allData = [headerRow]; // Start with header row
    let rowNum = 1;

    // Append data from each report
    for (let i = 0; i < reportFiles.length; i++) {
        const report = reportFiles[i];
        const ws = report.workbook.Sheets[templateInfo.sheetName];
        const data = XLSX.utils.sheet_to_json(ws, { header: 1 });

        // Skip header row and append data rows with row number
        for (let j = 1; j < data.length; j++) {
            allData.push([rowNum++, ...data[j]]);
        }

        bar.style.width = ((i + 1) / reportFiles.length * 100) + '%';
    }

    // Create compiled sheet with all data
    const compiledWs = XLSX.utils.aoa_to_sheet(allData);

    // Apply styling to header row (bold, centered)
    const range = XLSX.utils.decode_range(compiledWs['!ref']);
    for (let col = range.s.c; col <= range.e.c; col++) {
        const cellAddress = XLSX.utils.encode_cell({ r: 0, c: col });
        if (!compiledWs[cellAddress]) continue;
        compiledWs[cellAddress].s = {
            font: { bold: true },
            alignment: { horizontal: 'center', vertical: 'center' }
        };
    }

    // Auto-fit column widths
    const colWidths = headerRow.map(col => ({
        wch: Math.max(String(col).length + 2, 10)
    }));
    compiledWs['!cols'] = colWidths;

    XLSX.utils.book_append_sheet(newWb, compiledWs, templateInfo.sheetName);

    document.getElementById('progressSection').classList.add('hidden');

    // Download the compiled workbook
    XLSX.writeFile(newWb, 'Compiled - KyroReports.xlsx');
    showToast('Report compiled and downloaded successfully', 'success');
    downloadBtn.disabled = false;
};

function size(bytes) {
    const k = 1024;
    const units = ['B', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return (bytes / Math.pow(k, i)).toFixed(1) + ' ' + units[i];
}

// Single Report Validation Section
let validateTemplateFile = null;
let validateReportFile = null;

document.getElementById('validateTemplateInput').onchange = async (e) => setValidateTemplate(e.target.files[0]);
document.getElementById('validateReportInput').onchange = async (e) => setValidateReport(e.target.files[0]);

async function setValidateTemplate(file) {
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

        validateTemplateFile = {
            name: file.name,
            content: content
        };

        document.getElementById('validateTemplateInfo').innerHTML =
            `<b>${validateTemplateFile.name}</b> (${size(validateTemplateFile.content.byteLength)})`;
        document.getElementById('validateTemplateInfo').classList.remove('hidden');

        // Re-validate if report exists
        if (validateReportFile) {
            validateSingleReport();
        }
    } catch (err) {
        showToast('Failed to read file', 'error');
    }
}

async function setValidateReport(file) {
    if (!file || !file.name.toLowerCase().endsWith('.xlsx')) {
        showToast('Report must be a .xlsx file', 'error');
        return;
    }

    try {
        const content = await file.arrayBuffer();
        const validation = validateExcelFile(content);

        if (!validation.valid) {
            showToast('Invalid Excel file', 'error');
            return;
        }

        validateReportFile = {
            name: file.name,
            content: content,
            validation: validation
        };

        // Validate against template if exists
        if (validateTemplateFile) {
            validateSingleReport();
        } else {
            showValidationResult('Please upload a template first', false);
        }
    } catch (err) {
        showToast('Failed to read file', 'error');
    }
}

function validateSingleReport() {
    if (!validateTemplateFile) {
        showToast('Please upload a template file first', 'error');
        return;
    }

    if (!validateReportFile) {
        showToast('Please upload a report file first', 'error');
        return;
    }

    const templateInfo = getValidateTemplateInfo();
    if (!templateInfo) {
        showValidationResult('Template file is empty or invalid', false);
        return;
    }

    const validation = validateReportFile.validation;
    const errors = [];

    // Check if template sheet name exists in report
    if (!validation.sheetNames.includes(templateInfo.sheetName)) {
        errors.push(`Sheet "${templateInfo.sheetName}" not found. Available sheets: ${validation.sheetNames.join(', ')}`);
        showValidationResult(errors.join('<br>'), false);
        return;
    }

    // Check columns match
    const wb = validation.workbook;
    const ws = wb.Sheets[templateInfo.sheetName];
    const reportData = XLSX.utils.sheet_to_json(ws, { header: 1 });

    if (reportData.length === 0) {
        errors.push('Sheet is empty');
        showValidationResult(errors.join('<br>'), false);
        return;
    }

    const reportColumns = reportData[0];
    const templateCols = templateInfo.columns;

    // Compare columns
    const templateColsStr = templateCols.map(c => String(c).trim()).join(',');
    const reportColsStr = reportColumns.map(c => String(c).trim()).join(',');

    if (templateColsStr !== reportColsStr) {
        errors.push(`Columns don't match template.<br>Expected: "${templateColsStr}"<br>Found: "${reportColsStr}"`);
        showValidationResult(errors.join('<br>'), false);
        return;
    }

    // All checks passed
    showValidationResult(`<b>✓ Valid</b><br>Report matches template structure`, true);
}

function getValidateTemplateInfo() {
    if (!validateTemplateFile) return null;

    const wb = XLSX.read(validateTemplateFile.content, { type: 'array' });
    const sheetName = wb.SheetNames[0];
    const ws = wb.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(ws, { header: 1 });

    if (data.length === 0) return null;

    return {
        sheetName: sheetName,
        columns: data[0],
        data: data
    };
}

function showValidationResult(message, isValid) {
    const resultDiv = document.getElementById('validateReportResult');
    resultDiv.innerHTML = message;
    resultDiv.className = `file-item ${isValid ? 'valid' : 'invalid'}`;
    resultDiv.classList.remove('hidden');
}