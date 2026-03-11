<script lang="ts">
	import Navbar from '$lib/components/Navbar.svelte';
	import ToastContainer from '$lib/components/ToastContainer.svelte';
	import CompileReport from '$lib/components/CompileReport.svelte';
	import ValidateReport from '$lib/components/ValidateReport.svelte';
	import TemplateReport from '$lib/components/TemplateReport.svelte';
	import type { TemplateFile, ReportFile, ValidationResult } from '$lib/types';

	let templateFile: TemplateFile | null = null;
	let reportFiles: ReportFile[] = [];
	let validateTemplateFile: TemplateFile | null = null;
	let validateReportFile: any | null = null;
	let validationResult: ValidationResult | null = null;
</script>

<svelte:head>
	<title>Kyro Reports</title>
	<link
		rel="icon"
		href="data:image/svg+xml,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 100 100'><rect fill='%2310b981' width='100' height='100' rx='15'/><path fill='white' d='M25 30h50v8H25zm0 15h50v8H25zm0 15h35v8H25z'/></svg>"
	/>
</svelte:head>

<div class="min-h-screen bg-gray-50">
	<ToastContainer />

	<Navbar />

	<div class="max-w-[1200px] mx-auto px-3 pb-10">
		<CompileReport bind:templateFile bind:reportFiles />
		<ValidateReport bind:validateTemplateFile bind:validateReportFile bind:validationResult />
		<TemplateReport />
	</div>
</div>
