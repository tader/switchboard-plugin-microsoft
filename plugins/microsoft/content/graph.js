import { MAX_FILE, fail, check } from './package.js';
export const encode = (value) => encodeURIComponent(value);
export const itemPath = (drive, item) =>
  `/drives/${encode(drive)}/items/${encode(item)}`;
export async function bounded(response, limit = MAX_FILE) {
  if (!response.body) return Buffer.alloc(0);
  const chunks = [];
  let size = 0;
  const reader = response.body.getReader();
  try {
    for (;;) {
      const { done, value } = await reader.read();
      if (done) break;
      size += value.length;
      if (size > limit)
        throw fail(
          'response_too_large',
          'The response exceeded the size limit.',
          413,
        );
      chunks.push(Buffer.from(value));
    }
    return Buffer.concat(chunks);
  } finally {
    await reader.cancel().catch(() => {});
    reader.releaseLock();
  }
}
export function transferUrl(raw, { conversion = false } = {}) {
  let url;
  try {
    url = new URL(raw);
  } catch {
    throw fail(
      'invalid_transfer',
      'Microsoft returned an invalid file transfer URL.',
      502,
    );
  }
  if (
    url.protocol !== 'https:' ||
    url.username ||
    url.password ||
    url.port ||
    !(
      /(?:^|\.)(?:sharepoint\.com|sharepoint-df\.com|1drv\.com|onedrive\.com)$/.test(
        url.hostname,
      ) ||
      (conversion && /^[a-z0-9-]+-mediap\.svc\.ms$/.test(url.hostname))
    )
  )
    throw fail(
      'invalid_transfer',
      'Microsoft returned an unsupported file transfer host.',
      502,
    );
  return url;
}
export class GraphClient {
  constructor(
    accessToken,
    { access = 'read', signal, fetcher = fetch, org = false } = {},
  ) {
    Object.assign(this, { accessToken, access, signal, fetcher, org });
    this.driveIds = new Set();
  }
  async request(method, relative, body, headers = {}, raw = false) {
    if (this.signal?.aborted)
      throw fail('cancelled', 'Request cancelled.', 499);
    const url = new URL('https://graph.microsoft.com/v1.0' + relative);
    check(
      url.origin === 'https://graph.microsoft.com' &&
        url.pathname.startsWith('/v1.0/') &&
        !url.hash &&
        !/\\|(?:^|\/)\.\.(?:\/|$)/.test(relative),
      'Invalid Graph path.',
    );
    if (this.org) {
      check(
        method === 'GET' &&
          /^\/v1\.0\/(me|users)(?:\/[^/]+)?(?:\/(manager|directReports))?$/.test(
            url.pathname,
          ),
        'Unsupported organization request.',
      );
    } else {
      check(
        /^\/v1\.0\/(me\/drive|drives\/[^/]+)(?:\/|$|:)/.test(url.pathname),
        'Unsupported document request.',
      );
      if (this.access === 'read' && !['GET', 'HEAD'].includes(method))
        throw fail('request_forbidden', 'This connection is read-only.', 403);
      if (this.access !== 'all' && url.pathname.startsWith('/v1.0/drives/')) {
        const drive = decodeURIComponent(url.pathname.split('/')[3]);
        if (!this.driveIds.has(drive)) {
          const own = await this.request('GET', '/me/drive?$select=id');
          this.driveIds.add(own.id);
          if (drive !== own.id)
            throw fail(
              'request_forbidden',
              'Select access to all files for another drive or SharePoint.',
              403,
            );
        }
      }
    }
    let response;
    try {
      response = await this.fetcher(url, {
        method,
        headers: {
          authorization: `Bearer ${this.accessToken}`,
          ...(body !== undefined
            ? {
                'content-type':
                  body instanceof Uint8Array
                    ? 'application/octet-stream'
                    : 'application/json',
              }
            : {}),
          ...headers,
        },
        body:
          body === undefined
            ? undefined
            : body instanceof Uint8Array
              ? body
              : JSON.stringify(body),
        redirect: 'manual',
        signal: this.signal
          ? AbortSignal.any([this.signal, AbortSignal.timeout(60_000)])
          : AbortSignal.timeout(60_000),
      });
    } catch {
      throw fail(
        ['GET', 'HEAD'].includes(method)
          ? 'api_transport_failed'
          : 'write_uncertain',
        ['GET', 'HEAD'].includes(method)
          ? 'Microsoft request failed or timed out.'
          : 'The write result is unknown. Read the file before retrying.',
        502,
      );
    }
    if (raw && [301, 302, 303, 307, 308].includes(response.status))
      return response;
    if (!response.ok) {
      await response.body?.cancel();
      const codes = {
        401: 'api_unauthorized',
        403: 'api_forbidden',
        404: 'not_found',
        409: 'file_conflict',
        412: 'version_conflict',
        423: 'file_locked',
        429: 'rate_limited',
      };
      throw fail(
        codes[response.status] ?? 'api_failed',
        `Microsoft returned HTTP ${response.status}.`,
        response.status,
      );
    }
    if (raw) return response;
    if ([202, 204].includes(response.status)) {
      await response.body?.cancel();
      return {};
    }
    try {
      return JSON.parse((await bounded(response, 8 * 1024 * 1024)).toString());
    } catch (e) {
      if (e.code) throw e;
      throw fail(
        'invalid_response',
        'Microsoft returned an invalid response.',
        502,
      );
    }
  }
  async metadata(drive, item) {
    return this.request(
      'GET',
      itemPath(drive, item) +
        '?$select=id,name,size,eTag,file,webUrl,parentReference',
    );
  }
  async download(drive, item, { format } = {}) {
    check(
      format === undefined || format === 'pdf',
      'Unsupported conversion format.',
    );
    let response = await this.request(
      'GET',
      itemPath(drive, item) + '/content' + (format ? '?format=pdf' : ''),
      undefined,
      {},
      true,
    );
    for (
      let hop = 0;
      [301, 302, 303, 307, 308].includes(response.status);
      hop++
    ) {
      if (hop >= 3) {
        await response.body?.cancel();
        throw fail(
          'invalid_transfer',
          'Microsoft file transfer redirected too many times.',
          502,
        );
      }
      const url = transferUrl(response.headers.get('location'), {
        conversion: format === 'pdf',
      });
      await response.body?.cancel();
      try {
        response = await this.fetcher(url, {
          redirect: 'manual',
          signal: this.signal
            ? AbortSignal.any([this.signal, AbortSignal.timeout(60_000)])
            : AbortSignal.timeout(60_000),
        });
      } catch {
        throw fail(
          'download_failed',
          'File download failed or timed out.',
          502,
        );
      }
    }
    if (!response.ok) {
      await response.body?.cancel();
      throw fail(
        'download_failed',
        `File download returned HTTP ${response.status}.`,
        response.status,
      );
    }
    return bounded(response);
  }
  async read(drive, item) {
    const metadata = await this.metadata(drive, item);
    if (metadata.size > MAX_FILE)
      throw fail(
        'file_too_large',
        'Office packages are limited to 25 MiB.',
        413,
      );
    const bytes = await this.download(drive, item);
    const after = await this.metadata(drive, item);
    if (metadata.eTag !== after.eTag)
      throw fail(
        'version_conflict',
        'The file changed during download. Read it again.',
        409,
      );
    return { metadata, bytes };
  }
  async replace(drive, item, version, bytes) {
    check(
      typeof version === 'string' &&
        version.length > 0 &&
        version.length <= 1000,
      'Provide the version returned by a read.',
    );
    return this.request('PUT', itemPath(drive, item) + '/content', bytes, {
      'If-Match': version,
    });
  }
  async create(drive, parent, name, bytes) {
    check(
      typeof name === 'string' &&
        name.length > 0 &&
        name.length <= 200 &&
        !/[\\/:*?"<>|\x00-\x1f]/.test(name) &&
        !['.', '..'].includes(name),
      'Provide a filename without path separators.',
    );
    return this.request(
      'PUT',
      itemPath(drive, parent) +
        ':/' +
        encode(name) +
        ':/content?%40microsoft.graph.conflictBehavior=fail',
      bytes,
      { 'If-Match': '"switchboard-create-only"' },
    );
  }
}
export const publicFile = (file) => ({
  drive: file.parentReference?.driveId,
  id: file.id,
  name: file.name,
  version: file.eTag,
  size: file.size,
  webUrl: file.webUrl,
});
