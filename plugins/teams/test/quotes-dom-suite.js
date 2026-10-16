import { fixtureHTML } from './dom-suite.js';
function initializeQuoteFixture() {
  const message = document.querySelector('[data-mid="100"]');
  message.addEventListener('contextmenu', event => {
    event.preventDefault(); document.querySelector('[data-tid="message-actions-quoted-reply"]')?.remove();
    const button = document.createElement('button'); button.dataset.tid = 'message-actions-quoted-reply'; button.textContent = 'Reply with quote';
    button.onclick = () => {
      const card = document.createElement('card'); card.setAttribute('itemscope', 'QuotedReplyCard'); card.setAttribute('contenteditable', 'false');
      card.setAttribute('data-itemprops', JSON.stringify({ messageId: window.wrongQuote ? 'other' : '100', preview: 'Hello Thomas' }));
      card.innerHTML = '<div data-tid="quoted-reply-card"><span>Alice</span><span data-tid="quoted-reply-preview-content">Hello Thomas</span></div>';
      const editor = document.querySelector('[contenteditable="true"]'); editor.append(card, document.createElement('p')); button.remove();
    };
    document.body.append(button);
  });
  document.querySelector('[data-tid="sendMessageCommands-send"]').onclick = () => {
    const editor = document.querySelector('[contenteditable="true"]'), card = editor.querySelector('card');
    const clone = editor.cloneNode(true); clone.querySelector('card')?.remove();
    const message = document.createElement('div'); message.dataset.tid = 'chat-pane-message'; message.dataset.mid = String(++window.messageCount);
    const body = document.createElement('div'); body.setAttribute('data-message-content', '');
    if (card) { const quote = document.createElement('div'); quote.setAttribute('data-track-module-name', 'messageQuotedReply'); quote.innerHTML = card.innerHTML; body.append(quote); }
    const reply = document.createElement('p'); reply.textContent = clone.textContent; body.append(reply); message.append(body);
    document.querySelector('[data-tid="message-pane-list-viewport"]').append(message); editor.replaceChildren(); window.sentCount++;
  };
}
export const quotesFixtureHTML = fixtureHTML.replace('</body>', '<script>(' + initializeQuoteFixture.toString() + ')()</script></body>');
export async function runQuotesDOMSuite(page, dom, selectors, html) {
  const results = [], check = (condition, label) => { if (!condition) throw new Error(label); results.push(label); };
  const run = (action, args = {}) => page.evaluate(dom, action, args, selectors);
  await page.evaluate(html => { document.open(); document.write(html); document.close(); }, html);
  const chat = (await run('chats')).items[0], args = { ...chat, replyToMessageId: '100' };
  check((await run('quoteTarget', { ...chat, replyToMessageId: 'missing' })).error.code === 'message_unavailable', 'refuses to quote a missing message ID');
  const prepare = async () => {
    const point = await run('quoteTarget', args);
    await page.call('Input.dispatchMouseEvent', { type: 'mousePressed', ...point, button: 'right', clickCount: 1 });
    await page.call('Input.dispatchMouseEvent', { type: 'mouseReleased', ...point, button: 'right', clickCount: 1 });
    return run('quoteSelect', args);
  };
  const prepared = await prepare();
  check(prepared.replyToMessageId === '100' && prepared.quotePreview === 'Hello Thomas', 'uses native Reply with quote and verifies the selected message ID');
  check((await run('prepare', chat)).error.code === 'draft_exists', 'ordinary sends preserve existing quote drafts');
  await page.call('Input.insertText', { text: 'My reply' });
  const point = await run('sendPoint', { ...args, text: 'My reply' });
  check(!point.error, 'verifies reply text separately from quote preview');
  await page.clickPoint(point);
  const confirmed = await run('confirm', { ...args, text: 'My reply', ...prepared });
  check(confirmed.status === 'observed_in_chat' && confirmed.replyToMessageId === '100' && confirmed.quoteObserved && confirmed.message.hasQuote, 'confirms a new quoted reply using rendered quote preview and message ID');
  await prepare(); await page.call('Input.insertText', { text: 'Guarded reply' });
  const guarded = await run('sendPoint', { ...args, text: 'Guarded reply' });
  await page.evaluate(() => document.querySelector('card').setAttribute('data-itemprops', JSON.stringify({ messageId: 'other', preview: 'Hello Thomas' })));
  await page.clickPoint(guarded);
  check(await page.evaluate(() => window.sentCount === 1), 'blocks send when the quoted message changes before the click');
  await page.evaluate(() => { document.querySelector('[contenteditable="true"]').replaceChildren(); window.wrongQuote = true; });
  check((await prepare()).error.code === 'quote_mismatch', 'refuses a native menu action that inserts a different quote');
  return results;
}
