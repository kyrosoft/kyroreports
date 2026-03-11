<script lang="ts">
	import './layout.css';

	let { children } = $props();
</script>

<svelte:head>
	<link
		rel="icon"
		href="data:image/svg+xml,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 100 100'><rect fill='%2310b981' width='100' height='100' rx='15'/><path fill='white' d='M25 30h50v8H25zm0 15h50v8H25zm0 15h35v8H25z'/></svg>"
	/>
</svelte:head>

{@render children()}
