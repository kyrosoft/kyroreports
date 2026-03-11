import { writable } from 'svelte/store';
import type { Toast } from '$lib/types';

export function createToastStore() {
	const { subscribe, update } = writable<Toast[]>([]);
	let toastId = 0;

	return {
		subscribe,
		show: (message: string, type: 'error' | 'success' = 'error') => {
			const id = toastId++;
			const toast: Toast = { message, type, id };

			update((toasts) => [...toasts, toast]);

			setTimeout(() => {
				update((toasts) => toasts.filter((t) => t.id !== id));
			}, 3000);
		}
	};
}

export const toasts = createToastStore();
