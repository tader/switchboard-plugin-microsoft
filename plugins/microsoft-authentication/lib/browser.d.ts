import type { Bundle } from './session.js';
export function acquireBrowserRefresh(directory: string, options?: { silent?: boolean; browserPath?: string; signal?: AbortSignal; sharedProfile?: import('../types/api.d.ts').AuthenticationProfile }): Promise<Pick<Bundle, 'clientId' | 'origin' | 'identity' | 'refreshToken'>>;
