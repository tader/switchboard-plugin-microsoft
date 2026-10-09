// OpenAPI schema helpers extracted from Switchboard for the standalone build tool.
export function makeResolver(doc: any) {
  return function resolve(node: any, depth = 0): any {
    while (node && typeof node === 'object' && typeof node.$ref === 'string' && depth++ < 20) {
      const ref: string = node.$ref;
      if (!ref.startsWith('#/')) return undefined;
      node = ref
        .slice(2)
        .split('/')
        .map((s) => s.replace(/~1/g, '/').replace(/~0/g, '~'))
        .reduce((o, k) => (o == null ? o : o[decodeURIComponent(k)]), doc);
    }
    return node;
  };
}

export function exampleJson(schema: any, resolve: (n: any) => any): string | undefined {
  const value = sample(schema, resolve, 0, new Set());
  return value === undefined ? undefined : JSON.stringify(value, null, 2);
}

function sample(schema: any, resolve: (n: any) => any, depth: number, seen: Set<any>): any {
  const s = resolve(schema);
  if (!s || typeof s !== 'object' || depth > 6 || seen.has(s)) return undefined;
  if (s.example !== undefined) return s.example;
  if (s.default !== undefined) return s.default;
  if (Array.isArray(s.enum) && s.enum.length) return s.enum[0];
  const next = new Set(seen).add(s);
  const variant = s.allOf ?? s.oneOf ?? s.anyOf;
  if (Array.isArray(variant)) {
    if (s.allOf) return Object.assign({}, ...variant.map((v: any) => sample(v, resolve, depth + 1, next) ?? {}));
    return sample(variant[0], resolve, depth + 1, next);
  }
  const type = Array.isArray(s.type) ? s.type.find((t: string) => t !== 'null') : s.type;
  if (type === 'object' || s.properties) {
    const out: Record<string, any> = {};
    for (const [k, v] of Object.entries<any>(s.properties ?? {})) {
      const r = resolve(v);
      if (r?.readOnly) continue;
      const val = sample(v, resolve, depth + 1, next);
      if (val !== undefined) out[k] = val;
    }
    return out;
  }
  if (type === 'array') {
    const item = sample(s.items, resolve, depth + 1, next);
    return item === undefined ? [] : [item];
  }
  if (type === 'integer' || type === 'number') return 0;
  if (type === 'boolean') return false;
  if (type === 'string') {
    if (s.format === 'date-time') return new Date(0).toISOString();
    if (s.format === 'date') return '1970-01-01';
    return '';
  }
  return undefined;
}
