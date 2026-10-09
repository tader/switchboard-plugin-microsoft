import type { PluginContext } from './types/api.d.ts';
import type * as Microsoft from './types/microsoft.d.ts';

const SCOPES: Record<string, string[]> = { read: ['Tasks.Read'], full: ['Tasks.ReadWrite'] };

export default function setup(ctx: PluginContext) {
  const { microsoftService } = ctx.require<ReturnType<typeof Microsoft.default>['exports']>('microsoft');
  return {
    services: [
      microsoftService({
        id: 'microsoft-todo',
        name: 'Microsoft To Do',
        description: 'Tasks',
        icon: 'icon.svg',
        docsUrl: 'https://learn.microsoft.com/graph/api/resources/todo-overview',
        spec: 'todo',
        fields: [
          {
            key: 'access',
            label: 'Access',
            type: 'select',
            default: 'full',
            options: [
              { value: 'read', label: 'View task lists and tasks' },
              { value: 'full', label: 'View and edit task lists and tasks' },
            ],
          },
        ],
        scopes: (c) => SCOPES[c.access] ?? SCOPES['full'],
      }),
    ],
  };
}
