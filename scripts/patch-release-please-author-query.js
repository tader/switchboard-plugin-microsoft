import fs from 'node:fs/promises';
import { pathToFileURL } from 'node:url';

// Work around GitHub returning NOT_FOUND for an unavailable commit author's user.
// Release Please accepts an absent username; retain the commit's name and email.
// Fail closed if an action update changes the query rather than silently skipping it.
export function patchAuthorQuery(source) {
  const author = /author \{\s+name\s+email\s+user \{\s+login\s+\}\s+\}/g;
  if ([...source.matchAll(author)].length !== 1) {
    throw new Error('Expected exactly one Release Please commit author query; review the pinned action.');
  }
  return source.replace(author, 'author { name email }');
}

if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href) {
  const file = process.argv[2];
  if (!file) throw new Error('Usage: node scripts/patch-release-please-author-query.js <action-bundle>');
  await fs.writeFile(file, patchAuthorQuery(await fs.readFile(file, 'utf8')));
}
