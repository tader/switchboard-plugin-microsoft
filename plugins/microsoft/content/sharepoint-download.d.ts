import type { PluginContext, ServiceDefinition } from '../types/api.d.ts';
export function createSharePointDownload(ctx: PluginContext, service: ServiceDefinition, options?: { fetcher?: typeof fetch }): Promise<{ service: ServiceDefinition; dispose(): Promise<void> }>;
