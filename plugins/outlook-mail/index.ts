import type { PluginContext } from './types/api.d.ts';
import type * as Microsoft from './types/microsoft.d.ts';

const SCOPES: Record<string, string[]> = { read: ['Mail.Read'], send: ['Mail.Read', 'Mail.Send'], full: ['Mail.ReadWrite', 'Mail.Send'] };

export default function setup(ctx: PluginContext) {
  const { microsoftService } = ctx.require<ReturnType<typeof Microsoft.default>['exports']>('microsoft');
  return {
    services: [
      microsoftService({
        id: 'outlook-mail',
        name: 'Outlook Mail',
        description: 'Email',
        icon: 'icon.svg',
        docsUrl: 'https://learn.microsoft.com/graph/api/resources/mail-api-overview',
        spec: 'mail',
        fields: [
          {
            key: 'access',
            label: 'Access',
            type: 'select',
            default: 'full',
            options: [
              { value: 'read', label: 'Read email' },
              { value: 'send', label: 'Read and send email' },
              { value: 'full', label: 'Read, send, organize and delete email' },
            ],
          },
        ],
        scopes: (c) => SCOPES[c.access] ?? SCOPES['full'],
      }),
    ],
  };
}
