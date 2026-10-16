export const FORMAT: string;
export const CLIENT_ID: string;
export const TEAMS_CLIENT_ID: string;
export const TEAMS_ORIGINS: string[];
export const ORIGINS: string[];
export const RESOURCES: Record<string, string>;
export const GRAPH: string;
export interface Bundle {
  resources?: Partial<Record<'outlook' | 'skype' | 'chatsvcagg' | 'substrate', { accessToken: string; expiresAt: number; scope: string }>>;
  format: string; clientId: string; origin: string; identity: { tenant: string; oid: string };
  accessToken: string; refreshToken: string; expiresAt: number; scope: string; generation: string;
}
export function validateBundle(bundle: unknown, previous?: Bundle): Bundle;
export function fresh(bundle: Bundle): boolean;
export function fail(code: string, message: string, status?: number): Error & { code: string; status: number; reauthenticationRequired?: boolean };
export function identify(bundle: Bundle, options?: { signal?: AbortSignal; fetcher?: typeof fetch }): Promise<{ id: string; label: string }>;
export function refresh(bundle: any, options?: { resource?: 'graph' | 'outlook' | 'skype' | 'chatsvcagg' | 'substrate'; signal?: AbortSignal; fetcher?: typeof fetch; onRotate?: (bundle: Bundle) => Promise<void> }): Promise<Bundle>;
