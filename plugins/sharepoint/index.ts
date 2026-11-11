import type { PluginContext } from './types/api.d.ts';
import type * as Microsoft from './types/microsoft.d.ts';

export default async function setup(ctx: PluginContext) {
  const { microsoftService, createSharePointDownload } = ctx.require<ReturnType<typeof Microsoft.default>['exports']>('microsoft');
  const service = microsoftService({
    id: 'sharepoint', name: 'SharePoint', description: 'Sites, pages, lists and document libraries',
    icon: 'icon.svg', docsUrl: 'https://learn.microsoft.com/graph/api/resources/sharepoint', spec: 'sharepoint',
    fields: [{ key: 'access', label: 'Access', type: 'select', default: 'read', required: true,
      options: [{ value: 'read', label: 'Read accessible sites' }, { value: 'full', label: 'Read and edit accessible sites' }] }],
    scopes: c => c.access === 'full' ? ['Sites.ReadWrite.All'] : ['Sites.Read.All'],
  });
  const download = await createSharePointDownload(ctx, service);
  return { services: [service, download.service], dispose: download.dispose };
}
