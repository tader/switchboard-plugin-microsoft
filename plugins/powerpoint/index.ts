import type { PluginContext, PluginInstance } from './types/api.d.ts';
export default function setup(ctx: PluginContext): Promise<PluginInstance> {
  return ctx
    .require<{
      createContentPlugin(
        ctx: PluginContext,
        options: { id: string; name: string },
      ): Promise<PluginInstance>;
    }>('microsoft')
    .createContentPlugin(ctx, { id: 'powerpoint', name: 'PowerPoint' });
}
