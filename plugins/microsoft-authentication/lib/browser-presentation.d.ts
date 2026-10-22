import type { DeviceAuthorization, PluginContext } from '../types/api.d.ts';
export function browserAuthorization(capabilities: PluginContext['capabilities'], expiresIn?: number): Extract<DeviceAuthorization, { browser: true }>;
