<script lang="ts">
	import { toasts } from '$lib/utils/toast';
</script>

<div class="fixed top-12 right-5 z-[1000] flex flex-col gap-2.5 pointer-events-none">
	{#each $toasts as toast (toast.id)}
		<div
			class="bg-white border border-red-500 border-l-4 border-l-red-500 rounded px-4 py-3 shadow-lg text-xs text-red-600 min-w-[250px] max-w-[400px] pointer-events-auto transition-opacity duration-300 {toast.type === 'success'
				? '!border-emerald-500 !border-l-emerald-500 text-emerald-600'
				: ''}"
		>
			{toast.message}
		</div>
	{/each}
</div>
