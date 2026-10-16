const string = { type: 'string' },
  integer = { type: 'integer' },
  object = { type: 'object', additionalProperties: true };
const parameter = (
  name,
  location = 'query',
  required = false,
  schema = string,
) => ({ name, in: location, required, schema });
const response = {
  description: 'Result',
  content: { 'application/json': { schema: object } },
};
const operation = (id, summary, parameters = [], body) => ({
  operationId: id,
  summary,
  parameters,
  ...(body
    ? {
        requestBody: {
          required: true,
          content: { 'application/json': { schema: body } },
        },
      }
    : {}),
  responses: {
    200: response,
    400: { description: 'Invalid input or unsupported file' },
    403: { description: 'Access denied' },
    409: { description: 'Content or structure conflict' },
    412: { description: 'Stale file version' },
    423: { description: 'File locked' },
  },
});
function editSchema(kind) {
  const op = (name, props, required = []) => ({
    type: 'object',
    required: ['op', ...required],
    properties: { op: { type: 'string', enum: [name] }, ...props },
    additionalProperties: false,
  });
  const id = {
    ...string,
    description: 'Element identifier returned by reading this file version.',
  };
  const common = [
    op(
      'replaceText',
      {
        id,
        find: string,
        text: string,
        expectedMatches: {
          type: 'integer',
          default: 1,
          minimum: 1,
          maximum: 1000,
        },
      },
      ['id', 'find', 'text'],
    ),
    op('setCellText', { id, text: string }, ['id', 'text']),
  ];
  const operations =
    kind === 'word'
      ? [
          ...common,
          op('setParagraphText', { id, text: string }, ['id', 'text']),
          op(
            'insertParagraph',
            {
              text: string,
              style: {
                type: 'string',
                description: 'Normal, Title, Subtitle or Heading1–Heading9',
              },
              after: id,
              before: id,
            },
            ['text'],
          ),
          op('deleteParagraph', { id }, ['id']),
          op(
            'insertTable',
            {
              rows: { type: 'array', items: { type: 'array', items: string } },
              after: id,
            },
            ['rows'],
          ),
        ]
      : [
          ...common,
          op('setShapeText', { id, text: string }, ['id', 'text']),
          op('setNotes', { id, text: string }, ['id', 'text']),
          op('addSlide', {
            title: string,
            body: string,
            notes: string,
            index: integer,
          }),
          op('duplicateSlide', { id, index: integer, notes: string }, ['id']),
          op('moveSlide', { id, index: integer }, ['id', 'index']),
          op('deleteSlide', { id }, ['id']),
        ];
  return {
    type: 'array',
    minItems: 1,
    maxItems: 500,
    items: { oneOf: operations },
  };
}
export function contentSpec(kind, base) {
  const paths = {};
  if (kind === 'org-explorer') {
    const id = [
        parameter('person', 'path', true, {
          type: 'string',
          description: 'Person ID, email/UPN, or me.',
        }),
      ],
      page = [
        parameter('limit', 'query', false, {
          type: 'integer',
          minimum: 1,
          maximum: 100,
          default: 100,
        }),
        parameter('cursor'),
      ];
    paths['/people'] = {
      get: operation(
        'searchPeople',
        'Search people by name, email or user principal name',
        [parameter('q', 'query', true), ...page],
      ),
    };
    for (const [suffix, name, extra] of [
      ['', 'Read profile', []],
      ['/manager', 'Read immediate manager', []],
      ['/manager-chain', 'Read manager chain (up to 20 levels)', []],
      ['/reports', 'List direct reports', page],
      ['/peers', 'List peers with the same manager', page],
      [
        '/subtree',
        'Read a bounded reporting subtree',
        [
          parameter('depth', 'query', false, {
            type: 'integer',
            minimum: 1,
            maximum: 5,
            default: 2,
          }),
          parameter('maxPeople', 'query', false, {
            type: 'integer',
            minimum: 1,
            maximum: 500,
            default: 500,
          }),
        ],
      ],
    ])
      paths['/people/{person}' + suffix] = {
        get: operation(
          'org' + (suffix || '/profile').replace(/[^a-z]/gi, ''),
          name,
          [...id, ...extra],
        ),
      };
  } else {
    const plural = kind === 'word' ? 'documents' : 'presentations',
      id = [parameter('drive', 'path', true), parameter('item', 'path', true)],
      edits = editSchema(kind);
    const content =
      kind === 'word'
        ? {
            type: 'object',
            properties: {
              paragraphs: {
                type: 'array',
                items: {
                  oneOf: [
                    string,
                    {
                      type: 'object',
                      required: ['text'],
                      properties: { text: string, style: string },
                    },
                  ],
                },
              },
              tables: {
                type: 'array',
                items: {
                  type: 'array',
                  items: { type: 'array', items: string },
                },
              },
            },
          }
        : {
            type: 'object',
            properties: {
              slides: {
                type: 'array',
                items: {
                  type: 'object',
                  properties: { title: string, body: string, notes: string },
                },
              },
            },
          };
    paths['/capabilities'] = {
      get: operation(
        kind + 'Capabilities',
        'Read supported formats, operations and limits',
      ),
    };
    paths['/files/search'] = {
      get: operation(kind + 'Search', 'Find matching Office files', [
        parameter('q', 'query', true),
        parameter('drive'),
      ]),
    };
    const create = operation(
      kind + 'Create',
      'Create a file or copy and edit a template',
      [],
      {
        type: 'object',
        required: ['drive', 'parent', 'name'],
        properties: {
          drive: string,
          parent: string,
          name: string,
          content,
          template: {
            type: 'object',
            required: ['drive', 'item'],
            properties: { drive: string, item: string },
          },
          operations: edits,
        },
      },
    );
    create.requestBody.content['application/json'].example =
      kind === 'word'
        ? {
            drive: 'drive-id',
            parent: 'folder-id',
            name: 'Document.docx',
            content: {
              paragraphs: [
                { text: 'Project plan', style: 'Heading1' },
                'Introduction',
              ],
              tables: [
                [
                  ['Task', 'Owner'],
                  ['Review', 'Team'],
                ],
              ],
            },
          }
        : {
            drive: 'drive-id',
            parent: 'folder-id',
            name: 'Presentation.pptx',
            content: {
              slides: [
                {
                  title: 'Project overview',
                  body: 'First point\nSecond point',
                  notes: 'Speaker notes',
                },
              ],
            },
          };
    paths['/' + plural] = { post: create };
    paths[`/${plural}/{drive}/{item}`] = {
      get: operation(
        kind + 'Read',
        'Read file version and editable content',
        id,
      ),
    };
    const edit = operation(
      kind + 'Edit',
      'Apply a version-checked batch of content edits',
      id,
      {
        type: 'object',
        required: ['version', 'operations'],
        properties: {
          version: {
            type: 'string',
            description: 'Exact opaque version from the latest read.',
          },
          operations: edits,
        },
      },
    );
    edit.requestBody.content['application/json'].example = {
      version: '"etag-from-read"',
      operations: [
        {
          op: kind === 'word' ? 'setParagraphText' : 'setShapeText',
          id: kind === 'word' ? 'p0' : 'slide256/shape2',
          text: 'Updated text',
        },
      ],
    };
    paths[`/${plural}/{drive}/{item}/edits`] = { post: edit };
  }
  return {
    openapi: '3.0.3',
    info: {
      title:
        kind === 'word'
          ? 'Word'
          : kind === 'powerpoint'
            ? 'PowerPoint'
            : 'Org Explorer',
      version: '1.0.0',
      description:
        kind === 'org-explorer'
          ? 'Read-only organization relationships using existing Microsoft permissions.'
          : 'Content editing through Graph file transfer and server-side Open XML processing. Reads return version-bound element IDs; edits require that version. Untouched package parts are retained. No live Office or browser session is used for data operations.',
    },
    servers: [{ url: base }],
    paths,
  };
}
