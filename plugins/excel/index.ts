import type { PluginContext } from './types/api.d.ts';
import type * as Microsoft from './types/microsoft.d.ts';

const SCOPES: Record<string, string[]> = { read: ['Files.Read'], full: ['Files.ReadWrite'], all: ['Files.ReadWrite.All'] };

export default function setup(ctx: PluginContext) {
  const { microsoftService } = ctx.require<ReturnType<typeof Microsoft.default>['exports']>('microsoft');
  return {
    services: [
      microsoftService({
        id: 'excel',
        name: 'Excel',
        description: 'Excel workbooks in OneDrive for Business and SharePoint',
        icon: 'icon.svg',
        docsUrl: 'https://learn.microsoft.com/graph/api/resources/excel',
        spec: 'excel',
        fields: [
          {
            key: 'access',
            label: 'Access',
            type: 'select',
            default: 'full',
            options: [
              { value: 'read', label: 'View your files' },
              { value: 'full', label: 'View and edit your files' },
              { value: 'all', label: 'View and edit all files you can access, including shared ones' },
            ],
          },
        ],
        scopes: (c) => SCOPES[c.access] ?? SCOPES['full'],
      }),
    ],
  };
}
