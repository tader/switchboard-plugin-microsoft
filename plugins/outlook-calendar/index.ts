import type { PluginContext } from './types/api.d.ts';
import type * as Microsoft from './types/microsoft.d.ts';

const SCOPES: Record<string, string[]> = { read: ['Calendars.Read'], full: ['Calendars.ReadWrite'] };

export default function setup(ctx: PluginContext) {
  const { microsoftService } = ctx.require<ReturnType<typeof Microsoft.default>['exports']>('microsoft');
  return {
    services: [
      microsoftService({
        id: 'outlook-calendar',
        name: 'Outlook Calendar',
        description: 'Calendars and events',
        icon: 'icon.svg',
        docsUrl: 'https://learn.microsoft.com/graph/api/resources/calendar-overview',
        spec: 'calendar',
        fields: [
          {
            key: 'access',
            label: 'Access',
            type: 'select',
            default: 'full',
            options: [
              { value: 'read', label: 'View calendars and events' },
              { value: 'full', label: 'View and edit calendars and events' },
            ],
          },
        ],
        scopes: (c) => SCOPES[c.access] ?? SCOPES['full'],
      }),
    ],
  };
}
