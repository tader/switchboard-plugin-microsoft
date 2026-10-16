import type { PluginContext, PluginInstance } from '../types/api.d.ts';
export function createContentPlugin(
  ctx: PluginContext,
  options: { id: string; name: string; kind?: string },
): Promise<PluginInstance>;
