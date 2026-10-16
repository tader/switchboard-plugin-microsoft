import http from 'node:http';
import { randomBytes } from 'node:crypto';
import { MAX_FILE, check, fail } from './package.js';
import { GraphClient, itemPath, publicFile, bounded } from './graph.js';
import { readWord, editWord, createWord } from './word.js';
import {
  readPowerPoint,
  editPowerPoint,
  createPowerPoint,
} from './powerpoint.js';
import { OrgExplorer } from './org.js';
import { contentSpec } from './spec.js';
const scopes = {
  read: ['Files.Read'],
  full: ['Files.ReadWrite'],
  all: ['Files.ReadWrite.All'],
};
const validAccess = (value) => {
  check(Object.hasOwn(scopes, value), 'Select read, full or all access.');
  return value;
};
const cleanFilename = (name, kind) => {
  check(
    typeof name === 'string' &&
      name.toLowerCase().endsWith(kind === 'word' ? '.docx' : '.pptx'),
    'Use the matching .docx or .pptx filename.',
  );
  return name;
};
const capabilities = {
  word: [
    'replaceText',
    'setParagraphText',
    'insertParagraph',
    'deleteParagraph',
    'insertTable',
    'setCellText',
  ],
  powerpoint: [
    'replaceText',
    'setShapeText',
    'setCellText',
    'setNotes',
    'addSlide',
    'duplicateSlide',
    'moveSlide',
    'deleteSlide',
  ],
};
export async function dispatch(kind, graph, method, url, input) {
  if (kind === 'org-explorer') {
    check(method === 'GET', 'Org Explorer is read-only.');
    const org = new OrgExplorer(graph),
      query = Object.fromEntries(url.searchParams);
    if (url.pathname === '/people') return org.search(query);
    const match = url.pathname.match(
      /^\/people\/([^/]+)(?:\/(manager|manager-chain|reports|peers|subtree))?$/,
    );
    if (!match) throw fail('not_found', 'Unknown Org Explorer operation.', 404);
    const id = decodeURIComponent(match[1]),
      action = match[2];
    check(
      id.length > 0 && id.length <= 300 && !/[\/\\\x00-\x1f]/.test(id),
      'Invalid person identifier.',
    );
    return action === 'manager'
      ? { person: await org.manager(id) }
      : action === 'manager-chain'
        ? org.chain(id)
        : action === 'reports'
          ? org.reports(id, query)
          : action === 'peers'
            ? org.peers(id, query)
            : action === 'subtree'
              ? org.subtree(id, query)
              : org.profile(id);
  }
  const plural = kind === 'word' ? 'documents' : 'presentations',
    read = kind === 'word' ? readWord : readPowerPoint,
    edit = kind === 'word' ? editWord : editPowerPoint,
    create = kind === 'word' ? createWord : createPowerPoint;
  if (method === 'GET' && url.pathname === '/capabilities')
    return {
      kind,
      operations: capabilities[kind],
      formats: [kind === 'word' ? '.docx' : '.pptx'],
      maxFileBytes: MAX_FILE,
      requiresVersion: true,
    };
  if (method === 'GET' && url.pathname === '/files/search') {
    const q = url.searchParams.get('q'),
      drive = url.searchParams.get('drive');
    check(q && q.length <= 100, 'Search with 1–100 characters.');
    const query = encodeURIComponent(q.replace(/'/g, "''"));
    const data = await graph.request(
      'GET',
      `${drive ? '/drives/' + encodeURIComponent(drive) : '/me/drive'}/root/search(q='${query}')?$select=id,name,eTag,size,webUrl,parentReference&$top=100`,
    );
    return {
      files: (data.value ?? [])
        .filter((f) =>
          f.name?.toLowerCase().endsWith(kind === 'word' ? '.docx' : '.pptx'),
        )
        .map(publicFile),
      truncated: !!data['@odata.nextLink'],
    };
  }
  if (method === 'POST' && url.pathname === '/' + plural) {
    check(
      input &&
        typeof input.drive === 'string' &&
        typeof input.parent === 'string',
      'Provide destination drive and parent item IDs.',
    );
    const name = cleanFilename(input.name, kind);
    let bytes;
    if (input.template) {
      check(
        typeof input.template.drive === 'string' &&
          typeof input.template.item === 'string',
        'Provide template drive and item IDs.',
      );
      const template = await graph.read(
        input.template.drive,
        input.template.item,
      );
      read(template.bytes);
      bytes = input.operations?.length
        ? edit(template.bytes, input.operations).bytes
        : template.bytes;
    } else bytes = create(input.content ?? {});
    const result = await graph.create(input.drive, input.parent, name, bytes);
    return { file: publicFile(result), created: true };
  }
  const match = url.pathname.match(
    new RegExp(`^/${plural}/([^/]+)/([^/]+)(/edits)?$`),
  );
  if (!match) throw fail('not_found', 'Unknown document operation.', 404);
  const drive = decodeURIComponent(match[1]),
    item = decodeURIComponent(match[2]);
  check(
    drive.length > 0 &&
      item.length > 0 &&
      drive.length <= 500 &&
      item.length <= 500,
    'Invalid file identifier.',
  );
  if (method === 'GET' && !match[3]) {
    const value = await graph.read(drive, item);
    cleanFilename(value.metadata.name, kind);
    return { file: publicFile(value.metadata), content: read(value.bytes) };
  }
  if (method === 'POST' && match[3]) {
    check(
      input && typeof input.version === 'string',
      'Provide the version returned by a read.',
    );
    const value = await graph.read(drive, item);
    cleanFilename(value.metadata.name, kind);
    if (value.metadata.eTag !== input.version)
      throw fail(
        'version_conflict',
        'The file has changed. Read it again before editing.',
        412,
      );
    const result = edit(value.bytes, input.operations),
      updated = await graph.replace(drive, item, input.version, result.bytes);
    return { file: publicFile(updated), changes: result.changes };
  }
  throw fail('not_found', 'Unknown document operation.', 404);
}
export async function createContentPlugin(
  ctx,
  { id, name, kind = id },
  { fetcher = fetch } = {},
) {
  const authentication = ctx.require('microsoft-authentication'),
    org = kind === 'org-explorer';
  let disposed = false;
  const invocations = new Map(),
    controllers = new Set();
  const auth = authentication.createConsumerAuth(ctx, {
    id: org ? 'microsoft-graph' : 'onedrive',
    scopes: org
      ? ['User.ReadBasic.All']
      : (config) => scopes[validAccess(config.access ?? 'read')],
    fields: org
      ? []
      : [
          {
            key: 'access',
            label: 'Access',
            type: 'select',
            default: 'read',
            required: true,
            options: [
              { value: 'read', label: 'Read your files' },
              { value: 'full', label: 'Edit your files' },
              {
                value: 'all',
                label: 'Edit all accessible files, including SharePoint',
              },
            ],
          },
        ],
  });
  const server = http.createServer(async (request, response) => {
    const json = (status, body) => {
      if (!response.destroyed) {
        response.writeHead(status, {
          'content-type': 'application/json',
          'cache-control': 'no-store',
        });
        response.end(JSON.stringify(body));
      }
    };
    const url = new URL(request.url, 'http://adapter.localhost'),
      key = /^Bearer ([A-Za-z0-9_-]{43})$/.exec(
        request.headers.authorization ?? '',
      )?.[1],
      invocation = key && invocations.get(key);
    if (key) invocations.delete(key);
    if (
      !invocation ||
      invocation.expires < Date.now() ||
      invocation.method !== request.method ||
      invocation.path !== url.pathname + url.search
    ) {
      request.resume();
      return json(403, {
        error: {
          code: 'invalid_invocation',
          message: 'Invalid or expired API invocation.',
        },
      });
    }
    const controller = new AbortController();
    controllers.add(controller);
    response.once('close', () => controller.abort());
    const signal = invocation.signal
      ? AbortSignal.any([controller.signal, invocation.signal])
      : controller.signal;
    try {
      let input;
      if (request.method === 'POST') {
        const bytes = await bounded(new Response(request), 2 * 1024 * 1024);
        try {
          input = JSON.parse(bytes.toString());
        } catch {
          throw fail('invalid_json', 'Provide a JSON body.');
        }
      }
      const graph = new GraphClient(invocation.bundle.accessToken, {
        access: invocation.access,
        signal,
        fetcher,
        org,
      });
      const result = await dispatch(kind, graph, request.method, url, input);
      json(200, result);
    } catch (error) {
      const known = error.code && Number.isInteger(error.status);
      json(known ? error.status : 500, {
        error: {
          code: known ? error.code : 'adapter_error',
          message: known ? error.message : 'The Microsoft adapter failed.',
        },
      });
    } finally {
      controllers.delete(controller);
    }
  });
  server.requestTimeout = 120_000;
  server.headersTimeout = 15_000;
  await new Promise((resolve, reject) => {
    server.once('error', reject);
    server.listen(0, '127.0.0.1', resolve);
  });
  const origin = `http://127.0.0.1:${server.address().port}`,
    placeholder = `http://${id}.localhost`;
  const prune = setInterval(() => {
    for (const [key, value] of invocations)
      if (value.expires < Date.now()) invocations.delete(key);
  }, 10000);
  prune.unref();
  const { reusable, token, ...baseAuth } = auth;
  const method = {
    ...baseAuth,
    async authorize(request, connection, opts) {
      if (disposed)
        throw fail('disposed', 'The plugin has been unloaded.', 503);
      if (request.url.origin !== placeholder && request.url.origin !== origin)
        throw fail('request_forbidden', 'Invalid adapter destination.', 403);
      if (org && request.method !== 'GET')
        throw fail('request_forbidden', 'Org Explorer is read-only.', 403);
      if (
        !org &&
        (connection.config.access ?? 'read') === 'read' &&
        !['GET', 'HEAD'].includes(request.method)
      )
        throw fail('request_forbidden', 'This connection is read-only.', 403);
      if (!['GET', 'POST'].includes(request.method))
        throw fail('request_forbidden', 'Unsupported adapter method.', 403);
      if (invocations.size >= 1000)
        throw fail('adapter_busy', 'Too many pending operations.', 429);
      const result = await token(connection, opts),
        bundle = result.credentials ?? connection.credentials;
      if (opts.signal?.aborted)
        throw fail('cancelled', 'Request cancelled.', 499);
      const key = randomBytes(32).toString('base64url');
      invocations.set(key, {
        bundle: { accessToken: result.accessToken },
        access: org ? 'read' : validAccess(connection.config.access ?? 'read'),
        method: request.method,
        path: request.url.pathname + request.url.search,
        expires: Date.now() + 10000,
        signal: opts.signal,
      });
      request.url.host = new URL(origin).host;
      request.headers.set('authorization', 'Bearer ' + key);
      return result.credentials ? { credentials: bundle } : undefined;
    },
    async revoke(connection) {
      for (const [key, value] of invocations)
        if (value.bundle.accessToken === connection.credentials.accessToken)
          invocations.delete(key);
    },
  };
  return {
    services: [
      {
        id,
        name,
        description: org
          ? 'People, managers, peers and reporting hierarchies'
          : 'Read and edit ' +
            (kind === 'word' ? 'Word documents' : 'PowerPoint presentations') +
            ' in OneDrive and SharePoint',
        icon: 'icon.svg',
        baseUrl: placeholder,
        allowedHosts: [new URL(placeholder).host, new URL(origin).host],
        openapi: contentSpec(kind, placeholder),
        authMethods: [method],
      },
    ],
    async dispose() {
      disposed = true;
      clearInterval(prune);
      invocations.clear();
      for (const c of controllers) c.abort();
      await new Promise((resolve) => {
        server.close(resolve);
        server.closeAllConnections();
      });
    },
  };
}
