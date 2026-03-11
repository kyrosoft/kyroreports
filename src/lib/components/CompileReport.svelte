<script lang="ts">
	import { toasts } from '$lib/utils/toast';
	import {
		extractXlsxFromZip,
		validateExcelFile,
		getTemplateInfo,
		formatFileSize,
		compileReports
	} from '$lib/utils/excel';
	import * as XLSX from 'xlsx-js-style';
	import type { TemplateFile, ReportFile, TemplateInfo } from '$lib/types';

	export let templateFile: TemplateFile | null = null;
	export let reportFiles: ReportFile[] = [];

	let progress = 0;
	let compiling = false;

	async function setTemplate(file: any) {
		if (!file || !file.name?.toLowerCase()?.endsWith('.xlsx')) {
			toasts.show('Template must be a .xlsx file', 'error');
			return;
		}

		try {
			const content = await file.arrayBuffer();
			const validation = validateExcelFile(content);

			if (!validation.valid) {
				toasts.show('Invalid Excel file', 'error');
				return;
			}

			templateFile = {
				name: file.name,
				content: content
			};

			if (reportFiles.length > 0) {
				validateReportsAgainstTemplate();
			}
		} catch (err) {
			toasts.show('Failed to read file', 'error');
		}
	}

	async function setReportsZip(file: any) {
		if (!file) return;

		if (!file.name?.toLowerCase()?.endsWith('.zip')) {
			toasts.show(`${file.name} is not a .zip file`, 'error');
			return;
		}

		reportFiles = [];

		try {
			const xlsxFiles = await extractXlsxFromZip(file);
			if (xlsxFiles.length === 0) {
				toasts.show(`No .xlsx files found in ${file.name}`, 'error');
				return;
			}

			for (const xlsxFile of xlsxFiles) {
				const validation = validateExcelFile(xlsxFile.content);
				reportFiles = [
					...reportFiles,
					{
						file: xlsxFile,
						valid: validation.valid,
						sheets: validation.sheets,
						sheetNames: validation.sheetNames,
						workbook: validation.workbook,
						originalName: file.name
					}
				];
			}
		} catch (err) {
			toasts.show(`Failed to read ${file.name}`, 'error');
			return;
		}

		if (templateFile) {
			validateReportsAgainstTemplate();
		}
	}

	function removeReport(index: number) {
		reportFiles = reportFiles.filter((_, i) => i !== index);
	}

	function validateReportsAgainstTemplate() {
		if (!templateFile || reportFiles.length === 0) return;

		const templateInfo = getTemplateInfo(templateFile);
		if (!templateInfo) return;

		reportFiles = reportFiles.map((report) => {
			const validation = {
				valid: report.valid,
				sheets: report.sheets,
				sheetNames: report.sheetNames,
				workbook: report.workbook
			};

			if (!validation.sheetNames.includes(templateInfo.sheetName)) {
				return { ...report, valid: false, validationError: `Sheet "${templateInfo.sheetName}" not found` };
			}

			const ws = validation.workbook.Sheets[templateInfo.sheetName];
			const reportData = XLSX.utils.sheet_to_json(ws, { header: 1 });

			if (reportData.length === 0) {
				return { ...report, valid: false, validationError: 'Sheet is empty' };
			}

			const reportColumns: any = reportData[0];
			const templateCols = templateInfo.columns;

			const templateColsStr = templateCols.map((c: any) => String(c).trim()).join(',');
			const reportColsStr = reportColumns.map((c: any) => String(c).trim()).join(',');

			if (templateColsStr !== reportColsStr) {
				return { ...report, valid: false, validationError: `Columns don't match template` };
			}

			return { ...report, valid: true, validationError: undefined };
		});
	}

	async function compileAndDownload() {
		const errors: string[] = [];

		if (!templateFile) {
			errors.push('Please upload a template file');
		}

		if (reportFiles.length === 0) {
			errors.push('Please upload a reports ZIP file');
		}

		if (errors.length === 0 && templateFile && reportFiles.length > 0) {
			const templateInfo = getTemplateInfo(templateFile);
			if (!templateInfo) {
				errors.push('Template file is empty or invalid');
			} else {
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

					if (!validation.sheetNames.includes(templateInfo.sheetName)) {
						errors.push(
							`${report.file.name}: Sheet "${templateInfo.sheetName}" not found. Available sheets: ${validation.sheetNames.join(', ')}`
						);
						continue;
					}

					const ws = validation.workbook.Sheets[templateInfo.sheetName];
					const reportData = XLSX.utils.sheet_to_json(ws, { header: 1 });

					if (reportData.length === 0) {
						errors.push(`${report.file.name}: Sheet is empty`);
						continue;
					}

					const reportColumns: any = reportData[0];
					const templateCols = templateInfo.columns;

					const templateColsStr = templateCols.map((c: any) => String(c).trim()).join(',');
					const reportColsStr = reportColumns.map((c: any) => String(c).trim()).join(',');

					if (templateColsStr !== reportColsStr) {
						errors.push(
							`${report.file.name}: Columns do not match template. Expected: "${templateColsStr}", Found: "${reportColsStr}"`
						);
					}
				}
			}
		}

		if (errors.length > 0) {
			errors.forEach((err) => toasts.show(err, 'error'));
			return;
		}

		compiling = true;
		progress = 0;

		try {
			const templateInfo = getTemplateInfo(templateFile!);
			compileReports(templateFile!, reportFiles, templateInfo!, (p) => (progress = p));
			toasts.show('Report compiled and downloaded successfully', 'success');
		} catch (err) {
			toasts.show('Failed to compile reports', 'error');
		}

		compiling = false;
		progress = 0;
	}
</script>

<section class="bg-white border border-gray-200 rounded-md mb-4 overflow-hidden shadow-sm">
	<h2 class="bg-slate-800 text-white px-4 py-2.5 m-0 text-xs font-semibold tracking-wide">
		Compile Report
	</h2>

	<div class="px-4 pt-3.5 pb-4">
		<!-- Upload Template -->
		<div class="mb-3 w-full">
			<h3 class="text-xs my-1.5 mx-0">
				1. Upload Template
				<span class="text-gray-500 text-[10px] ml-1">(Template: .xlsx file)</span>
			</h3>
			<input
				type="file"
				id="templateInput"
				accept=".xlsx"
				class="w-full px-1.5 py-1 mb-1 border border-gray-300 rounded text-xs"
				onchange={(e) => setTemplate((e.target as any).files?.[0])}
			/>
			{#if templateFile}
				<div class="px-2 py-1.5 mt-1.5 bg-green-50 rounded-md flex justify-between items-center text-xs">
					<b>{templateFile.name}</b> ({formatFileSize(templateFile.content.byteLength)})
				</div>
			{/if}
		</div>

		<!-- Upload Reports ZIP -->
		<div class="mb-3 w-full">
			<h3 class="text-xs my-1.5 mx-0">
				2. Upload Reports ZIP
				<span class="text-gray-500 text-[10px] ml-1"
					>(Reports: .zip file containing multiple .xlsx files)</span
				>
			</h3>
			<input
				type="file"
				id="reportsInput"
				accept=".zip"
				class="w-full px-1.5 py-1 mb-1 border border-gray-300 rounded text-xs"
				onchange={(e) => setReportsZip((e.target as any).files?.[0])}
			/>

			{#if reportFiles.length > 0}
				<div class="mt-1.5 space-y-1">
					{#each reportFiles as report, i}
						<div
							class="px-2 py-1.5 rounded-md flex justify-between items-center text-xs {report.valid
								? 'bg-green-50 border-l-2 border-l-emerald-500'
								: 'bg-red-50 border-l-2 border-l-red-500'}"
						>
							<span>
								{report.file.name} ({formatFileSize(report.file.content.byteLength)})
								<span class="text-[10px] ml-1.5 {report.valid
									? 'text-emerald-500'
									: 'text-red-500'}">
									{report.valid ? '✓ Valid' : `✗ ${report.validationError || 'Invalid'}`}
								</span>
								{#if report.originalName}
									<small class="ml-1">from {report.originalName}</small>
								{/if}
							</span>
							<button
								class="bg-transparent border-none text-red-500 text-xs px-1 py-0.5 m-0"
								onclick={() => removeReport(i)}
							>
								×
							</button>
						</div>
					{/each}
				</div>
			{/if}
		</div>

		<!-- Progress and Download -->
		<div class="mt-3 w-full">
			{#if compiling}
				<div class="w-full h-1.5 bg-gray-200 rounded-md mt-1.5 overflow-hidden">
					<div class="h-full bg-blue-600 transition-all duration-300" style="width: {progress}%"></div>
				</div>
			{/if}
			<button
				onclick={compileAndDownload}
				disabled={compiling}
				class="w-full px-4 py-2.5 text-xs font-semibold my-2 min-h-9 bg-emerald-500 hover:bg-emerald-600 disabled:bg-gray-300 disabled:cursor-not-allowed flex items-center justify-center leading-none"
			>
				DOWNLOAD COMPILED XLSX
			</button>
		</div>
	</div>
</section>
