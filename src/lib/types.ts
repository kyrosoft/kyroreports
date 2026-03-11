export interface Column {
	value?: string;
}

export interface Toast {
	message: string;
	type: 'error' | 'success';
	id: number;
}

export interface TemplateFile {
	name: string;
	content: ArrayBuffer;
}

export interface ReportFile {
	file: {
		name: string;
		content: ArrayBuffer;
	};
	valid: boolean;
	sheets: number;
	sheetNames: string[];
	workbook: any;
	validationError?: string;
	originalName?: string;
}

export interface ValidationResult {
	message: string;
	isValid: boolean;
}

export interface ExcelValidation {
	valid: boolean;
	sheets: number;
	sheetNames: string[];
	workbook: any;
}

export interface TemplateInfo {
	sheetName: string;
	columns: any[];
	data: any[][];
}
