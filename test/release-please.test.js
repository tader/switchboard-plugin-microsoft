import test from 'node:test';
import assert from 'node:assert/strict';
import { patchAuthorQuery } from '../scripts/patch-release-please-author-query.js';

test('Release Please workaround removes only the optional commit author account lookup', () => {
  const source = `const query = \`query { history { nodes {
    associatedPullRequests { nodes { number title files { nodes { path } } } }
    sha: oid message
    author {
      name
      email
      user {
        login
      }
    }
  } } }\`;`;
  const patched = patchAuthorQuery(source);
  assert.equal(patched, source.replace(/author \{[\s\S]*?\n    \}/, 'author { name email }'));
});

test('Release Please workaround fails closed for missing, changed or duplicate queries', () => {
  const author = 'author { name email user { login } }';
  for (const source of ['', author.replace('email', 'email extraField'), author + author, patchAuthorQuery(author)]) {
    assert.throws(() => patchAuthorQuery(source), /Expected exactly one/);
  }
});
