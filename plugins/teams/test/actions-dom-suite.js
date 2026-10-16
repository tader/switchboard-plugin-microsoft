import { fixtureHTML } from './dom-suite.js';
function initializeActionsFixture() {
  const list = document.querySelector('[data-tid="message-pane-list-viewport"]'); list.style.height = '250px';
  const ownWrapper = document.createElement('div'); ownWrapper.dataset.tid = 'chat-pane-item';
  ownWrapper.innerHTML = '<div class="fui-ChatMyMessage"><div data-tid="chat-pane-message" data-mid="200"><div data-message-content>Original note</div></div><div data-tid="diverse-reaction-summary"></div></div>';
  list.append(ownWrapper);
  const states = { '100': { like: false, heart: false }, '200': { like: false, heart: false } };
  window.actionCounts = { reaction: 0, edit: 0, delete: 0, unread: 0 };
  const reactionSummary = (mid, root) => {
    root.querySelector('[data-tid="diverse-reaction-summary"]')?.remove();
    const summary = document.createElement('div'); summary.dataset.tid = 'diverse-reaction-summary';
    for (const [reaction, selected] of Object.entries(states[mid])) {
      if (!selected) continue;
      const pill = document.createElement('button'); pill.dataset.tid = 'diverse-reaction-pill-button'; pill.setAttribute('aria-pressed', 'true');
      const emoji = document.createElement('img'); emoji.setAttribute('itemtype', 'http://schema.skype.com/Emoji'); emoji.setAttribute('itemid', reaction === 'like' ? 'yes' : reaction); pill.append(emoji);
      pill.onclick = () => toggleReaction(mid, reaction, root); summary.append(pill);
    }
    root.append(summary);
  };
  const toggleReaction = (mid, reaction, root) => {
    states[mid][reaction] = !states[mid][reaction]; window.actionCounts.reaction++;
    reactionSummary(mid, root);
    const button = document.getElementById(mid + '-popover-surface')?.querySelector('[data-tid="message-actions-' + reaction + '"]');
    button?.setAttribute('aria-pressed', String(states[mid][reaction]));
    document.querySelector('[data-tid="message-reactions-on-menu"] [data-tid="message-actions-' + reaction + '"]')?.setAttribute('aria-pressed', String(states[mid][reaction]));
  };
  const toolbar = (mid, root) => {
    document.querySelector('[data-tid="message-actions-container"]')?.remove();
    const surface = document.createElement('div'); surface.id = mid + '-popover-surface'; surface.dataset.tid = 'message-actions-container';
    const bar = document.createElement('div'); bar.dataset.tid = 'message-reactions-toolbar';
    for (const reaction of ['like', 'heart']) {
      const button = document.createElement('button'); button.dataset.tid = 'message-actions-' + reaction; button.setAttribute('aria-pressed', String(states[mid][reaction])); button.setAttribute('aria-label', reaction);
      const emoji = document.createElement('img'); emoji.setAttribute('itemtype', 'http://schema.skype.com/Emoji'); emoji.setAttribute('itemid', reaction === 'like' ? 'yes' : reaction); button.append(emoji);
      button.onclick = () => toggleReaction(mid, reaction, root); bar.append(button);
    }
    surface.append(bar); document.body.append(surface);
  };
  const attachMessage = node => {
    const mid = node.dataset.mid, wrapper = node.closest('[data-tid="chat-pane-item"]'), owned = !!node.closest('.fui-ChatMyMessage');
    node.addEventListener('mouseenter', () => toolbar(mid, wrapper));
    node.addEventListener('contextmenu', event => {
      event.preventDefault(); document.querySelector('[data-tid="message-actions-menu-renderer-getV9MessageActions"]')?.remove();
      const menu = document.createElement('div'); menu.dataset.tid = 'message-actions-menu-renderer-getV9MessageActions'; menu.setAttribute('role', 'menu');
      toolbar(mid, wrapper);
      const bar = document.querySelector('[data-tid="message-reactions-toolbar"]'); bar.dataset.tid = 'message-reactions-on-menu'; menu.append(bar);
      document.querySelector('[data-tid="message-actions-container"]')?.remove();
      if (owned) {
        const edit = document.createElement('button'); edit.dataset.tid = 'message-actions-edit'; edit.textContent = 'Edit';
        edit.onclick = () => {
          const original = node.querySelector('[data-message-content]').innerText;
          const originalClone = node.cloneNode(true); node.remove(); menu.remove();
          const container = wrapper.querySelector('.fui-ChatMyMessage');
          const editor = document.createElement('div'); editor.dataset.tid = 'ckeditor'; editor.contentEditable = 'true'; editor.textContent = original; editor.id = 'edit-message-' + mid; container.append(editor);
          const done = document.createElement('button'); done.dataset.tid = 'newMessageCommands-send'; done.textContent = 'Done'; done.setAttribute('data-track-thread-id', 'thread-a');
          done.onclick = () => { originalClone.querySelector('[data-message-content]').innerHTML = editor.innerHTML; editor.remove(); done.remove(); container.append(originalClone); attachMessage(originalClone); window.actionCounts.edit++; };
          container.append(done);
        };
        const remove = document.createElement('button'); remove.dataset.tid = 'message-actions-delete'; remove.textContent = 'Delete';
        remove.onclick = () => { menu.remove(); node.remove(); const marker = document.createElement('div'); marker.dataset.tid = 'deleted-message-placeholder'; marker.textContent = 'You deleted this message.'; wrapper.append(marker); window.actionCounts.delete++; };
        menu.append(edit, remove);
      }
      document.body.append(menu);
    });
  };
  document.querySelectorAll('[data-tid="chat-pane-message"]').forEach(attachMessage);
  document.querySelectorAll('[role="treeitem"][data-item-type="chat"]').forEach(row => row.addEventListener('contextmenu', event => {
    event.preventDefault(); document.querySelector('[data-tid="change-unread-status-menu-item"]')?.remove();
    const button = document.createElement('button'); button.dataset.tid = 'change-unread-status-menu-item'; button.textContent = row.getAttribute('data-is-unread') === 'true' ? 'Mark as read' : 'Mark as unread';
    button.onclick = () => { row.setAttribute('data-is-unread', String(row.getAttribute('data-is-unread') !== 'true')); button.remove(); window.actionCounts.unread++; }; document.body.append(button);
  }));
}
export const actionsFixtureHTML = fixtureHTML.replace('</body>', '<script>(' + initializeActionsFixture.toString() + ')()</script></body>');
export async function runActionsDOMSuite(page, dom, selectors, html) {
  const results = [], check = (ok, label) => { if (!ok) throw new Error(label); results.push(label); };
  const reset = () => page.evaluate(html => { document.open(); document.write(html); document.close(); }, html);
  const run = (action, args) => page.evaluate(dom, action, args, selectors);
  const base = { title: 'Alice', threadId: 'thread-a', key: { attribute: 'data-fui-tree-item-value', value: 'Folder|folder/OneGQL_ChatConversation|thread-a' } };
  const args = (kind, messageId = '200', extra = {}) => ({ ...base, kind, messageId, token: kind + '-fixture', ...extra });
  const context = async a => {
    const point = await run('messageTarget', a); if (point.error) return point;
    await page.call('Input.dispatchMouseEvent', { type: 'mousePressed', ...point, button: 'right', clickCount: 1 });
    await page.call('Input.dispatchMouseEvent', { type: 'mouseReleased', ...point, button: 'right', clickCount: 1 });
    return run('menuPrepare', a);
  };
  await reset();
  const reaction = args('reaction', '100', { reaction: 'like', selected: true });
  const menuPoint = await run('messageTarget', reaction);
  for (const type of ['mousePressed', 'mouseReleased']) await page.call('Input.dispatchMouseEvent', { type, ...menuPoint, button: 'right', clickCount: 1 });
  const menuCatalog = await run('reactionReady', reaction);
  check(menuCatalog.available.length === 2 && await page.evaluate(() => !document.querySelector('[data-tid="message-reactions-toolbar"]')), 'discovers reactions from the native menu without a hover toolbar');
  await run('reactionPrepare', reaction);
  check((await run('reactionCommit', reaction)).selected, 'adds a quick reaction through its guarded native message menu');
  await run('release', reaction); await reset();
  const point = await run('messageTarget', reaction);
  await page.call('Input.dispatchMouseEvent', { type: 'mouseMoved', ...point });
  const catalog = await run('reactionReady', reaction);
  check(catalog.available.length === 2 && catalog.available[0].reaction === 'like', 'lists native quick reactions and maps yes to like');
  await run('reactionPrepare', reaction);
  check((await run('reactionCommit', reaction)).selected === true, 'adds a reaction and observes the selected state');
  check((await run('reactionCommit', reaction)).changed === false && await page.evaluate(() => window.actionCounts.reaction === 1), 'setting an existing reaction does not toggle it off');
  const remove = { ...reaction, selected: false };
  await run('reactionPrepare', remove);
  check((await run('reactionCommit', remove)).selected === false, 'removes an existing reaction explicitly');
  check((await run('reactionPrepare', { ...remove, reaction: 'unavailable' })).error.code === 'reaction_unavailable', 'refuses unavailable reaction IDs');
  await reset();
  check((await run('messageTarget', args('delete', '100', { expectedText: 'Hello Thomas' }))).error.code === 'not_own_message', 'refuses deletion of another person’s message');
  check((await run('messageTarget', args('edit', '200', { expectedText: 'Stale text' }))).error.code === 'message_changed', 'refuses stale expected text');
  const edit = args('edit', '200', { expectedText: 'Original note', text: 'Updated note' });
  await context(edit); const prepared = await run('editPrepare', edit);
  check(prepared.ready, 'opens the inline editor for the exact own message');
  await run('editSelect', edit); await page.call('Input.insertText', { text: 'Updated note' });
  check((await run('editCommit', edit)).status === 'edited_in_ui', 'saves an inline edit and observes the same message ID with new text');
  check(await page.evaluate(() => !document.querySelector('#fixture-composer').innerText.trim()), 'editing does not type in the main composer');
  await reset(); await context(edit); await run('editPrepare', edit); await run('editSelect', edit);
  await page.evaluate(() => document.querySelector('#fixture-composer').focus());
  await page.call('Input.insertText', { text: 'Updated note' });
  check(await page.evaluate(() => !document.querySelector('#fixture-composer').innerText.trim()), 'blocks edit replacement text if focus shifts to the main composer');
  check((await run('editCommit', edit)).error.code === 'edit_changed', 'does not save after a blocked or mismatched replacement');
  await reset();
  await page.evaluate(() => document.querySelector('#fixture-composer').textContent = 'Keep this draft');
  check((await run('messageTarget', edit)).error.code === 'draft_exists', 'editing preserves a manual message draft');
  await page.evaluate(() => { document.querySelector('#fixture-composer').textContent = ''; const img = document.createElement('img'); document.querySelector('[data-mid="200"] [data-message-content]').append(img); });
  check((await run('messageTarget', edit)).error.code === 'edit_content_unsupported', 'refuses to flatten attachments or rich quote/emoji content');
  await reset();
  const deletion = args('delete', '200', { expectedText: 'Original note' });
  await context(deletion);
  await page.evaluate(() => document.querySelector('[data-mid="200"] [data-message-content]').textContent = 'Changed meanwhile');
  check((await run('deleteCommit', deletion)).error.code === 'message_changed' && await page.evaluate(() => window.actionCounts.delete === 0), 'blocks deletion if the message changes after preparation');
  await reset(); await context(deletion);
  check((await run('deleteCommit', deletion)).status === 'deleted_in_ui', 'deletion requires a positive native tombstone');
  await reset(); await context(deletion);
  await page.evaluate(() => document.querySelector('[data-mid="100"]').click());
  check((await run('deleteCommit', deletion)).error.code === 'action_changed', 'invalidates a reused menu after a click on a different message');
  await run('release', deletion);
  await reset(); await context(deletion);
  await page.evaluate(() => { const other = document.querySelector('[data-mid="100"]'); other.dispatchEvent(new MouseEvent('contextmenu', { bubbles: true, cancelable: true })); });
  check((await run('deleteCommit', deletion)).error.code === 'action_changed', 'invalidates preparation if another message context menu is requested');
  await run('release', deletion);
  await reset(); await context(deletion);
  await page.evaluate(() => { const button = document.querySelector('[data-tid="message-actions-delete"]'); button.onclick = () => document.querySelector('[data-mid="200"]').closest('[data-tid="chat-pane-item"]').remove(); });
  check((await run('deleteCommit', deletion)).error.code === 'mutation_uncertain', 'never confirms deletion merely because virtualization removed the message');
  await run('release', deletion);
  await reset();
  const guarded = await run('messageTarget', deletion);
  await page.evaluate(() => document.querySelector('[data-tid="sendMessageCommands-send"]').setAttribute('data-track-thread-id', 'thread-b'));
  await page.call('Input.dispatchMouseEvent', { type: 'mousePressed', ...guarded, button: 'right', clickCount: 1 });
  await page.call('Input.dispatchMouseEvent', { type: 'mouseReleased', ...guarded, button: 'right', clickCount: 1 });
  check(await page.evaluate(() => !document.querySelector('[data-tid="message-actions-menu-renderer-getV9MessageActions"]')), 'blocks the native context menu when the chat changes before right-click');
  await reset();
  const unread = { ...base, token: 'unread-fixture', kind: 'unread' };
  const sidebar = await run('unreadTarget', unread);
  await page.call('Input.dispatchMouseEvent', { type: 'mousePressed', x: sidebar.x, y: sidebar.y, button: 'right', clickCount: 1 });
  await page.call('Input.dispatchMouseEvent', { type: 'mouseReleased', x: sidebar.x, y: sidebar.y, button: 'right', clickCount: 1 });
  await run('unreadPrepare', unread);
  check((await run('unreadCommit', unread)).status === 'unread_observed', 'marks the exact sidebar chat unread and verifies its flag');
  check((await run('unreadCommit', unread)).changed === false && await page.evaluate(() => window.actionCounts.unread === 1), 'mark unread never toggles an already-unread chat back to read');
  return results;
}
