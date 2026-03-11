import * as XLSX from 'xlsx-js-style';
import JSZip from 'jszip';
import type { ExcelValidation, TemplateInfo, TemplateFile } from '$lib/types';

export async function extractXlsxFromZip(file: File): Promise<any[]> {
	const zip = new JSZip();
	const arrayBuffer = await file.arrayBuffer();
	const zipContent = await zip.loadAsync(arrayBuffer);
	const xlsxFiles: any[] = [];

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

export function validateExcelFile(arrayBuffer: ArrayBuffer): ExcelValidation {
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

export function getTemplateInfo(templateFile: TemplateFile): TemplateInfo | null {
	if (!templateFile) return null;

	const wb = XLSX.read(templateFile.content, { type: 'array' });
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

export function formatFileSize(bytes: number): string {
	const k = 1024;
	const units = ['B', 'KB', 'MB', 'GB'];
	const i = Math.floor(Math.log(bytes) / Math.log(k));
	return (bytes / Math.pow(k, i)).toFixed(1) + ' ' + units[i];
}

export function createExcelWorkbook(
	columns: string[],
	sheetName: string = 'KyroReports',
	filename: string = 'Template - KyroReports.xlsx'
): void {
	const wb = XLSX.utils.book_new();
	const data = [columns];

	const wsData = XLSX.utils.aoa_to_sheet(data);

	// Apply styling to header row
	const range = XLSX.utils.decode_range(wsData['!ref']);
	for (let col = range.s.c; col <= range.e.c; col++) {
		const cellAddress = XLSX.utils.encode_cell({ r: 0, c: col });
		if (!wsData[cellAddress]) continue;
		wsData[cellAddress].s = {
			font: { bold: true },
			alignment: { horizontal: 'center', vertical: 'center' }
		};
	}

	const colWidths = columns.map((col) => ({
		wch: Math.max(String(col).length + 2, 10)
	}));
	wsData['!cols'] = colWidths;

	XLSX.utils.book_append_sheet(wb, wsData, sheetName);

	XLSX.writeFile(wb, filename);
}

export function compileReports(
	templateFile: TemplateFile,
	reportFiles: any[],
	templateInfo: TemplateInfo,
	onProgress: (progress: number) => void
): void {
	const newWb = XLSX.utils.book_new();

	const headerRow = ['No.', ...templateInfo.columns];
	const allData = [headerRow];
	let rowNum = 1;

	for (let i = 0; i < reportFiles.length; i++) {
		const report = reportFiles[i];
		const ws = report.workbook.Sheets[templateInfo.sheetName];
		const data = XLSX.utils.sheet_to_json(ws, { header: 1 });

		for (let j = 1; j < data.length; j++) {
			allData.push([rowNum++, ...data[j]]);
		}

		onProgress(((i + 1) / reportFiles.length) * 100);
	}

	const compiledWs = XLSX.utils.aoa_to_sheet(allData);

	// Apply styling to header row
	const range = XLSX.utils.decode_range(compiledWs['!ref']);
	for (let col = range.s.c; col <= range.e.c; col++) {
		const cellAddress = XLSX.utils.encode_cell({ r: 0, c: col });
		if (!compiledWs[cellAddress]) continue;
		compiledWs[cellAddress].s = {
			font: { bold: true },
			alignment: { horizontal: 'center', vertical: 'center' }
		};
	}

	const colWidths = headerRow.map((col: any) => ({
		wch: Math.max(String(col).length + 2, 10)
	}));
	compiledWs['!cols'] = colWidths;

	XLSX.utils.book_append_sheet(newWb, compiledWs, templateInfo.sheetName);

	XLSX.writeFile(newWb, 'Compiled - KyroReports.xlsx');
}
