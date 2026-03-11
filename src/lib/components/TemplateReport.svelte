<script lang="ts">
	import { onMount } from 'svelte';
	import { toasts } from '$lib/utils/toast';
	import { createExcelWorkbook } from '$lib/utils/excel';
	import type { Column } from '$lib/types';

	let columns: Column[] = [{ value: '' }];

	function renderTemplate() {
		if (columns.length === 0) {
			columns = [{ value: '' }];
		}
	}

	function autoResizeInput(input: HTMLInputElement) {
		const textLength = input.value.length;
		const calculatedWidth = Math.max(80, Math.min(200, textLength * 8 + 40));
		input.style.width = calculatedWidth + 'px';
	}

	function moveColumnLeft(index: number) {
		if (index > 0) {
			const newColumns = [...columns];
			[newColumns[index], newColumns[index - 1]] = [newColumns[index - 1], newColumns[index]];
			columns = newColumns;
		}
	}

	function moveColumnRight(index: number) {
		if (index < columns.length - 1) {
			const newColumns = [...columns];
			[newColumns[index], newColumns[index + 1]] = [newColumns[index + 1], newColumns[index]];
			columns = newColumns;
		}
	}

	function addColumn() {
		columns = [...columns, { value: '' }];
		setTimeout(() => {
			const inputs = document.querySelectorAll('.column-input-field');
			if (inputs.length > 0) {
				(inputs[inputs.length - 1] as HTMLInputElement).focus();
			}
		}, 0);
	}

	function removeColumn(index: number) {
		if (columns.length <= 1) {
			toasts.show('You must have at least one column', 'error');
			return;
		}
		columns = columns.filter((_, i) => i !== index);
	}

	function createTemplate() {
		const validColumns = columns.map((c: any) => c.value?.trim()).filter((c) => c);

		if (validColumns.length === 0) {
			toasts.show('Please add at least one column', 'error');
			return;
		}

		const errors: string[] = [];
		columns.forEach((col: any, index) => {
			if (!col.value || col.value.trim() === '') {
				errors.push(`Column ${index + 1} is empty`);
			}
		});

		if (errors.length > 0) {
			errors.forEach((err) => toasts.show(err, 'error'));
			return;
		}

		createExcelWorkbook(validColumns as string[], 'KyroReports', 'Template - KyroReports.xlsx');
		toasts.show('Template downloaded successfully', 'success');
	}

	onMount(() => {
		renderTemplate();
	});
</script>

<section class="bg-white border border-gray-200 rounded-md mb-4 overflow-hidden shadow-sm">
	<h2 class="bg-slate-800 text-white px-4 py-2.5 m-0 text-xs font-semibold tracking-wide">
		Template Report
	</h2>

	<div class="px-4 pt-3.5">
		<div class="w-full">
			<div class="my-2 px-2 py-2 border border-gray-300 rounded-md bg-gray-50 overflow-x-auto overflow-y-visible">
				<div class="inline-flex flex-nowrap gap-0 items-center p-0 min-w-max">
					{#each columns as col, index}
						<div class="group relative inline-flex items-center flex-shrink-0">
							<input
								type="text"
								bind:value={col.value}
								placeholder="Column {index + 1}..."
								class="column-input-field min-w-[80px] max-w-[180px] w-auto h-7 px-2 border border-gray-300 border-r-0 text-[11px] bg-white flex-shrink-0 m-0 focus:outline-none focus:border-blue-600 focus:z-10 focus:relative focus:inset-0 focus:inset-[-1px] focus:shadow-[0_0_0_1px_#2563eb] {index ===
								0
									? 'rounded-l-sm'
									: ''} {index === columns.length - 1
									? 'rounded-r-sm !border-r border-gray-300'
									: ''}"
								use:autoResizeInput
							/>
							<button
								onclick={() => removeColumn(index)}
								class="absolute -top-2 -right-px bg-red-500 text-white border-none rounded-full w-[13px] h-[13px] text-[9px] leading-none cursor-pointer flex items-center justify-center transition-all duration-150 opacity-0 group-hover:opacity-100 hover:bg-red-600 hover:scale-110 z-10 p-0 m-0 min-h-0"
								title="Remove column"
							>
								×
							</button>
							{#if index > 0}
								<button
									onclick={() => moveColumnLeft(index)}
									class="absolute -bottom-2 -left-px bg-indigo-500 text-white border-none rounded-full w-[13px] h-[13px] text-[8px] leading-none cursor-pointer flex items-center justify-center transition-all duration-150 opacity-0 group-hover:opacity-100 hover:bg-indigo-600 hover:scale-110 z-10 p-0 m-0 min-h-0 font-bold"
									title="Move left"
								>
									←
								</button>
							{/if}
							{#if index < columns.length - 1}
								<button
									onclick={() => moveColumnRight(index)}
									class="absolute -bottom-2 -right-px bg-indigo-500 text-white border-none rounded-full w-[13px] h-[13px] text-[8px] leading-none cursor-pointer flex items-center justify-center transition-all duration-150 opacity-0 group-hover:opacity-100 hover:bg-indigo-600 hover:scale-110 z-10 p-0 m-0 min-h-0 font-bold"
									title="Move right"
								>
									→
								</button>
							{/if}
						</div>
					{/each}
					<button
						onclick={addColumn}
						class="bg-emerald-500 text-white border border-emerald-600 rounded-md w-[22px] h-[22px] text-sm leading-none cursor-pointer inline-flex items-center justify-center flex-shrink-0 p-0 m-0 ml-1 min-h-0 transition-colors hover:bg-emerald-600"
						title="Add column"
					>
						+
					</button>
				</div>
				<p class="text-gray-500 text-[10px] mt-2 mb-0">{columns.length} column(s)</p>
			</div>
		</div>

		<button
			onclick={createTemplate}
			class="w-full px-4 py-2.5 text-xs font-semibold my-2 min-h-9 bg-emerald-500 hover:bg-emerald-600 flex items-center justify-center leading-none"
		>
			DOWNLOAD TEMPLATE XLSX
		</button>
	</div>
</section>
