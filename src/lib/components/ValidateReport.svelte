<script lang="ts">
	import { toasts } from '$lib/utils/toast';
	import { validateExcelFile, getTemplateInfo, formatFileSize } from '$lib/utils/excel';
	import * as XLSX from 'xlsx-js-style';
	import type { TemplateFile, ValidationResult } from '$lib/types';

	export let validateTemplateFile: TemplateFile | null = null;
	export let validateReportFile: any | null = null;
	export let validationResult: ValidationResult | null = null;

	async function setValidateTemplate(file: any) {
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

			validateTemplateFile = {
				name: file.name,
				content: content
			};

			if (validateReportFile) {
				validateSingleReport();
			}
		} catch (err) {
			toasts.show('Failed to read file', 'error');
		}
	}

	async function setValidateReport(file: any) {
		if (!file || !file.name?.toLowerCase()?.endsWith('.xlsx')) {
			toasts.show('Report must be a .xlsx file', 'error');
			return;
		}

		try {
			const content = await file.arrayBuffer();
			const validation = validateExcelFile(content);

			if (!validation.valid) {
				toasts.show('Invalid Excel file', 'error');
				return;
			}

			validateReportFile = {
				name: file.name,
				content: content,
				validation: validation
			};

			if (validateTemplateFile) {
				validateSingleReport();
			} else {
				validationResult = { message: 'Please upload a template first', isValid: false };
			}
		} catch (err) {
			toasts.show('Failed to read file', 'error');
		}
	}

	function validateSingleReport() {
		if (!validateTemplateFile) {
			toasts.show('Please upload a template file first', 'error');
			return;
		}

		if (!validateReportFile) {
			toasts.show('Please upload a report file first', 'error');
			return;
		}

		const templateInfo = getTemplateInfo(validateTemplateFile);
		if (!templateInfo) {
			validationResult = { message: 'Template file is empty or invalid', isValid: false };
			return;
		}

		const validation = validateReportFile.validation;
		const errors: string[] = [];

		if (!validation.sheetNames.includes(templateInfo.sheetName)) {
			errors.push(
				`Sheet "${templateInfo.sheetName}" not found. Available sheets: ${validation.sheetNames.join(', ')}`
			);
			validationResult = { message: errors.join('<br>'), isValid: false };
			return;
		}

		const ws = validation.workbook.Sheets[templateInfo.sheetName];
		const reportData = XLSX.utils.sheet_to_json(ws, { header: 1 });

		if (reportData.length === 0) {
			errors.push('Sheet is empty');
			validationResult = { message: errors.join('<br>'), isValid: false };
			return;
		}

		const reportColumns: any = reportData[0];
		const templateCols = templateInfo.columns;

		const templateColsStr = templateCols.map((c: any) => String(c).trim()).join(',');
		const reportColsStr = reportColumns.map((c: any) => String(c).trim()).join(',');

		if (templateColsStr !== reportColsStr) {
			errors.push(
				`Columns don't match template.<br>Expected: "${templateColsStr}"<br>Found: "${reportColsStr}"`
			);
			validationResult = { message: errors.join('<br>'), isValid: false };
			return;
		}

		validationResult = {
			message: '<b>✓ Valid</b><br>Report matches template structure',
			isValid: true
		};
	}
</script>

<section class="bg-white border border-gray-200 rounded-md mb-4 overflow-hidden shadow-sm pb-4">
	<h2 class="bg-slate-800 text-white px-4 py-2.5 m-0 text-xs font-semibold tracking-wide">
		Validate Report
	</h2>

	<div class="px-4 pt-3.5">
		<!-- Upload Template -->
		<div class="mb-3 w-full">
			<h3 class="text-xs my-1.5 mx-0">
				1. Upload Template
				<span class="text-gray-500 text-[10px] ml-1">(Template: .xlsx file)</span>
			</h3>
			<input
				type="file"
				id="validateTemplateInput"
				accept=".xlsx"
				class="w-full px-1.5 py-1 mb-1 border border-gray-300 rounded text-xs"
				onchange={(e) => setValidateTemplate((e.target as any).files?.[0])}
			/>
			{#if validateTemplateFile}
				<div class="px-2 py-1.5 mt-1.5 bg-green-50 rounded-md flex justify-between items-center text-xs">
					<b>{validateTemplateFile.name}</b> ({formatFileSize(validateTemplateFile.content.byteLength)})
				</div>
			{/if}
		</div>

		<!-- Upload Report -->
		<div class="mb-3 w-full">
			<h3 class="text-xs my-1.5 mx-0">
				2. Upload Report <span class="text-gray-500 text-[10px] ml-1">(Report: .xlsx file)</span>
			</h3>
			<input
				type="file"
				id="validateReportInput"
				accept=".xlsx"
				class="w-full px-1.5 py-1 mb-1 border border-gray-300 rounded text-xs"
				onchange={(e) => setValidateReport((e.target as any).files?.[0])}
			/>

			{#if validationResult}
				<div
					class="px-2 py-1.5 mt-1.5 rounded-md flex justify-between items-center text-xs {validationResult.isValid
						? 'bg-green-50 border-l-2 border-l-emerald-500'
						: 'bg-red-50 border-l-2 border-l-red-500'}"
				>
					{@html validationResult.message}
				</div>
			{/if}
		</div>
	</div>
</section>
