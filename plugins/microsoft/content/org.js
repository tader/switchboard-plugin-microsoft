import { check, fail } from './package.js';
import { encode } from './graph.js';
const BASIC =
  'id,displayName,givenName,surname,mail,userPrincipalName,jobTitle,department,officeLocation,companyName';
const person = (value) => {
  const result = {};
  for (const key of BASIC.split(','))
    if (value[key] !== undefined && value[key] !== null)
      result[key] = value[key];
  if (value['@odata.type'] === '#microsoft.graph.orgContact')
    result.type = 'contact';
  return result;
};
const base = (id) => (id === 'me' ? '/me' : '/users/' + encode(id));
const int = (value, def, max) => {
  const number = value === null || value === undefined ? def : Number(value);
  check(
    Number.isInteger(number) && number >= 1 && number <= max,
    `Provide an integer between 1 and ${max}.`,
  );
  return number;
};
export class OrgExplorer {
  constructor(graph) {
    this.graph = graph;
  }
  async list(path, limit = 100) {
    const value = await this.graph.request('GET', path);
    check(Array.isArray(value.value), 'Invalid people response.');
    return {
      people: value.value.slice(0, limit).map(person),
      truncated: value.value.length > limit || !!value['@odata.nextLink'],
      next: value['@odata.nextLink'],
    };
  }
  async page(relative, query = {}) {
    const limit = int(query.limit, 100, 100);
    let cursor = '';
    if (query.cursor) {
      check(
        typeof query.cursor === 'string' && query.cursor.length <= 4000,
        'Invalid cursor.',
      );
      let url;
      try {
        url = new URL(Buffer.from(query.cursor, 'base64url').toString());
      } catch {
        throw fail('invalid_input', 'Invalid cursor.');
      }
      const wanted = new URL('https://graph.microsoft.com/v1.0' + relative);
      check(
        url.origin === wanted.origin &&
          url.pathname === wanted.pathname &&
          !url.hash &&
          !url.username &&
          !url.password &&
          [...url.searchParams.keys()].every((k) =>
            ['$select', '$top', '$skiptoken', '$filter'].includes(k),
          ),
        'Cursor changed the organization resource.',
      );
      cursor = url.pathname.slice('/v1.0'.length) + url.search;
    }
    const result = await this.list(
      cursor ||
        relative +
          '?$select=' +
          BASIC +
          '&$top=' +
          limit +
          (query.search ? `&$filter=${encode(query.search)}` : ''),
      limit,
    );
    return {
      people: result.people,
      truncated: result.truncated,
      nextCursor: result.next
        ? Buffer.from(result.next).toString('base64url')
        : null,
    };
  }
  async profile(id) {
    return person(
      await this.graph.request('GET', base(id) + '?$select=' + BASIC),
    );
  }
  async manager(id) {
    try {
      return person(
        await this.graph.request('GET', base(id) + '/manager?$select=' + BASIC),
      );
    } catch (e) {
      if (e.status === 404) return null;
      throw e;
    }
  }
  async reports(id, query) {
    return this.page(base(id) + '/directReports', query);
  }
  async chain(id) {
    const managers = [],
      seen = new Set();
    let at = id;
    for (let depth = 0; depth < 20; depth++) {
      const p = await this.manager(at);
      if (!p) return { managers, truncated: false, cycle: false };
      if (seen.has(p.id)) return { managers, truncated: true, cycle: true };
      seen.add(p.id);
      managers.push(p);
      at = p.id;
    }
    return { managers, truncated: true, cycle: false };
  }
  async peers(id, query) {
    const current = await this.profile(id),
      manager = await this.manager(id);
    if (!manager) return { people: [], truncated: false, nextCursor: null };
    const page = await this.reports(manager.id, query);
    return { ...page, people: page.people.filter((p) => p.id !== current.id) };
  }
  async search(query) {
    check(
      typeof query.q === 'string' &&
        query.q.trim().length >= 2 &&
        query.q.length <= 100,
      'Search with 2–100 characters.',
    );
    const q = query.q.trim().replace(/'/g, "''");
    return this.page('/users', {
      ...query,
      search: `startswith(displayName,'${q}') or startswith(mail,'${q}') or startswith(userPrincipalName,'${q}')`,
    });
  }
  async subtree(id, query) {
    const depth = int(query.depth, 2, 5),
      max = int(query.maxPeople, 500, 500),
      root = await this.profile(id),
      people = [{ ...root, parent: null, depth: 0 }],
      seen = new Set([root.id]);
    let frontier = [root],
      truncated = false,
      cycle = false;
    for (let level = 1; level <= depth && frontier.length; level++) {
      const next = [];
      for (const parent of frontier) {
        let cursor;
        const cursors = new Set();
        let pages = 0;
        do {
          if (++pages > 20 || (cursor && cursors.has(cursor))) {
            truncated = true;
            break;
          }
          if (cursor) cursors.add(cursor);
          const page = await this.reports(parent.id, {
            limit: Math.min(100, max - people.length || 1),
            cursor,
          });
          for (const p of page.people) {
            if (seen.has(p.id)) {
              cycle = true;
              continue;
            }
            if (people.length >= max) {
              truncated = true;
              break;
            }
            seen.add(p.id);
            people.push({ ...p, parent: parent.id, depth: level });
            if (p.type !== 'contact') next.push(p);
          }
          cursor = page.nextCursor;
          if (people.length >= max) {
            truncated = true;
            break;
          }
        } while (cursor);
        if (people.length >= max) break;
      }
      frontier = next;
      if (people.length >= max) break;
    }
    // Reaching the requested depth is an explicit traversal boundary, not a complete-org claim.
    return { people, truncated, depthLimited: frontier.length > 0, cycle };
  }
}
