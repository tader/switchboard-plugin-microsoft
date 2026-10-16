import type { Bundle } from './session.js';
export function acquireBrowserRefresh(directory: string, options?: { silent?: boolean; browserPath?: string; signal?: AbortSignal }): Promise<Pick<Bundle, 'clientId' | 'origin' | 'identity' | 'refreshToken'>>;
