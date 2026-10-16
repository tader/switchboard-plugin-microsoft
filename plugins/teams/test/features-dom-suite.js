import { fixtureHTML } from './dom-suite.js';

function initializePeopleFixture() {
  const directory = [
    { name: 'Alice Adams', email: 'alice@example.com' },
    { name: 'Alice Adams', email: 'alice.other@example.com' },
    { name: 'Bob Brown', email: 'bob@example.com' },
  ];
  const send = document.querySelector('[data-tid="sendMessageCommands-send"]');
  document.querySelector('[data-testid="simple-collab-left-rail-header-new-message-button"]').onclick = () => {
    document.querySelector('#fixture-picker')?.remove();
    const picker = document.createElement('div'); picker.id = 'fixture-picker';
    const input = document.createElement('input'); input.dataset.tid = 'people-picker-search';
    const list = document.createElement('div'); list.setAttribute('role', 'listbox');
    const render = () => {
      list.replaceChildren();
      const query = input.value.toLowerCase();
      let found = directory.filter(person => person.name.toLowerCase().includes(query) || person.email.includes(query));
      if (window.duplicatePerson && found.length) found = [...found, found[0]];
      for (const person of found) {
        const row = document.createElement('div'); row.dataset.tid = 'people-picker-entry-' + person.email; row.dataset.type = query.includes('@') ? 'Person' : 'ADUser'; row.setAttribute('role', 'option');
        const name = document.createElement('div'); name.textContent = person.name;
        const email = document.createElement('div'); email.textContent = person.email;
        row.append(name, email);
        row.onclick = () => {
          picker.remove();
          document.querySelector('[data-tid="chat-title"]').textContent = person.name;
          document.querySelector('[data-tid="message-pane-list-viewport"]')?.remove();
          document.querySelector('[data-tid="chat-pane-new-chat"]')?.remove();
          const empty = document.createElement('div'); empty.dataset.tid = 'chat-pane-new-chat'; empty.textContent = 'Start your conversation'; document.body.append(empty);
          send.removeAttribute('data-track-thread-id');
          window.selectedEmail = person.email;
        };
        list.append(row);
      }
      const group = document.createElement('div'); group.dataset.tid = 'people-picker-entry-group@example.com'; group.dataset.type = 'Chat'; group.textContent = 'Group suggestion'; list.append(group);
    };
    input.addEventListener('input', render); picker.append(input, list); document.body.append(picker); render();
  };
  send.onclick = () => {
    const editor = document.querySelector('[contenteditable]');
    let list = document.querySelector('[data-tid="message-pane-list-viewport"]');
    if (!list) { list = document.createElement('div'); list.dataset.tid = 'message-pane-list-viewport'; document.body.append(list); }
    const message = document.createElement('div'); message.dataset.tid = 'chat-pane-message'; message.dataset.mid = String(++window.messageCount);
    const body = document.createElement('div'); body.setAttribute('data-message-content', ''); body.textContent = editor.innerText;
    message.append(body); list.append(message); editor.textContent = '';
    document.querySelector('[data-tid="chat-pane-new-chat"]')?.remove();
    send.setAttribute('data-track-thread-id', 'thread-for-' + window.selectedEmail);
    window.sentCount++;
  };
}

export const featuresFixtureHTML = fixtureHTML.replace('</body>', '<button data-testid="simple-collab-left-rail-header-new-message-button">New message</button><script>(' + initializePeopleFixture.toString() + ')()</script></body>');

export async function runFeaturesDOMSuite(page, dom, selectors, html) {
  const results = [];
  const check = (condition, label) => { if (!condition) throw new Error(label); results.push(label); };
  await page.evaluate(html => { document.open(); document.write(html); document.close(); }, html);
  const run = (action, args = {}) => page.evaluate(dom, action, args, selectors);
  await page.evaluate(() => {
    window.openedChats = 0;
    const alice = document.querySelector('[data-item-type="chat"]');
    alice.setAttribute('aria-labelledby', 'chat_list_unread_text');
    alice.addEventListener('click', () => window.openedChats++);
    const preview = document.createElement('span'); preview.id = 'message-preview-chat-list-item_a'; preview.textContent = 'Unread preview'; alice.append(preview);
    const list = document.querySelector('[data-tid="message-pane-list-viewport"]');
    const marker = document.createElement('div'); marker.dataset.tid = 'last-read-line'; marker.textContent = 'Last read'; list.append(marker);
    for (const id of ['101', '102']) {
      const message = document.createElement('div'); message.dataset.tid = 'chat-pane-message'; message.dataset.mid = id;
      const body = document.createElement('div'); body.setAttribute('data-message-content', ''); body.textContent = 'Unread ' + id; message.append(body); list.append(message);
    }
  });
  const chats = await run('chats'), alice = chats.items[0];
  check(alice.unread && !chats.items[1].unread && alice.preview === 'Unread preview', 'detects shared accessibility unread markers and reads preview');
  check(await page.evaluate(() => window.openedChats === 0), 'unread summaries do not open chats');
  const unread = await run('unreadMessages', { ...alice, limit: 1 });
  check(unread.items.length === 1 && unread.items[0].id === '101' && unread.truncated && unread.boundaryFound, 'returns only messages after Last read and reports truncation');
  await page.evaluate(() => document.querySelector('[data-tid="last-read-line"]').remove());
  const fallback = await run('unreadMessages', { ...alice, limit: 10 });
  check(!fallback.boundaryFound && fallback.items.length === 0 && fallback.recentMessages.length === 3, 'never labels recent fallback messages as confirmed unread');
  await page.evaluate(() => document.querySelector('[contenteditable]').textContent = 'Keep my draft');
  check((await run('peopleFocus')).error.code === 'draft_exists', 'people search preserves manual message drafts');
  await page.evaluate(() => document.querySelector('[contenteditable]').textContent = '');
  const search = async query => {
    const before = await run('peopleFocus');
    await page.call('Input.insertText', { text: query });
    return run('peopleResults', { ...before, query });
  };
  const people = await search('Alice');
  check(people.items.length === 2 && people.items.every(p => p.email.includes('alice')), 'search disambiguates same-name people and excludes group suggestions');
  check((await run('selectPerson', { email: 'alice@example.com' })).error.code === 'search_changed', 'recipient selection requires the exact email query');
  await page.evaluate(() => window.duplicatePerson = true);
  await search('alice@example.com');
  check((await run('selectPerson', { email: 'alice@example.com' })).error.code === 'recipient_unavailable', 'duplicate exact-email results fail closed');
  await page.evaluate(() => window.duplicatePerson = false);
  await search('alice@example.com');
  const person = await run('selectPerson', { email: 'alice@example.com' });
  check(person.key.value === 'alice@example.com' && person.title === 'Alice Adams' && person.threadId === null, 'selects verified email into a new empty conversation');
  check(await page.evaluate(() => window.sentCount === 0), 'opening a new conversation sends nothing');
  const prepared = await run('prepare', person);
  check(prepared.beforeIds.length === 0, 'prepares a new chat without message history or thread metadata');
  await page.call('Input.insertText', { text: 'Hello Alice' });
  const point = await run('sendPoint', { ...person, text: 'Hello Alice' });
  await page.clickPoint(point);
  const confirmed = await run('confirm', { ...person, text: 'Hello Alice', beforeIds: [] });
  check(confirmed.status === 'observed_in_chat' && confirmed.message.text === 'Hello Alice', 'confirms first message as a new thread materializes');
  check(await page.evaluate(() => window.sentCount === 1 && window.selectedEmail === 'alice@example.com'), 'first message goes to the verified person exactly once');
  await search('bob@example.com');
  const bob = await run('selectPerson', { email: 'bob@example.com' });
  await page.evaluate(() => document.querySelector('[contenteditable]').id = 'other-recipient-composer');
  check((await run('prepare', bob)).error.code === 'recipient_changed', 'refuses a reused composer whose identity changed');
  await page.evaluate(() => document.querySelector('[contenteditable]').id = 'fixture-composer');
  await page.evaluate(() => { const composer = document.querySelector('[contenteditable]'); composer.replaceWith(composer.cloneNode(true)); });
  check((await run('prepare', bob)).error.code === 'recipient_changed', 'refuses a replaced composer without fresh recipient verification');
  return results;
}
