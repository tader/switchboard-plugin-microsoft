/** Native Outlook web operations. These are intentionally separate from the Graph schemas. */
export function outlookSpec(service: string) {
  const paths: Record<string, any> = {};
  const add = (url: string, method: string, summary: string, example?: object) => {
    const parameters = [...url.matchAll(/\{([^}]+)\}/g)].map(match => ({ name: match[1], in: 'path', required: true, schema: { type: 'string' } }));
    if (method === 'get') for (const name of ['$top', '$skip', '$select', '$filter', '$orderby']) parameters.push({ name, in: 'query', required: false, schema: { type: name === '$top' || name === '$skip' ? 'integer' : 'string' } });
    (paths[url] ??= {})[method] = { operationId: `outlook.${service}.${method}.${url.replace(/[^a-zA-Z0-9]+/g, '.')}`, summary,
      description: 'Outlook web API. JSON property names and OData fields use PascalCase; Graph request bodies and IDs are not interchangeable.', parameters,
      ...(example ? { requestBody: { required: true, content: { 'application/json': { example } } } } : {}) };
  };
  if (service === 'outlook-mail') {
    for (const root of ['/me/messages', '/me/mailfolders/{folder-id}/messages']) {
      add(root, 'get', 'List messages'); add(root, 'post', 'Create a draft', { Subject: 'Draft', Body: { ContentType: 'Text', Content: 'Hello' } });
    }
    add('/me/mailfolders', 'get', 'List mail folders');
    add('/me/mailfolders/{folder-id}/childfolders', 'get', 'List child mail folders');
    add('/me/messages/{message-id}', 'get', 'Read message');
    add('/me/messages/{message-id}', 'patch', 'Update message', { IsRead: true });
    add('/me/messages/{message-id}', 'delete', 'Delete message');
    add('/me/messages/{message-id}/attachments', 'get', 'List attachments');
    add('/me/messages/{message-id}/attachments/{attachment-id}', 'get', 'Read attachment');
    add('/me/messages/{message-id}/attachments', 'post', 'Attach a file to a draft', { '@odata.type': '#Microsoft.OutlookServices.FileAttachment', Name: 'example.txt', ContentBytes: 'SGVsbG8=' });
    add('/me/messages/{message-id}/move', 'post', 'Move message', { DestinationId: 'archive' });
    add('/me/messages/{message-id}/reply', 'post', 'Reply to message', { Comment: 'Thank you' });
    add('/me/messages/{message-id}/replyall', 'post', 'Reply to all', { Comment: 'Thank you' });
    add('/me/messages/{message-id}/forward', 'post', 'Forward message', { Comment: 'For your information', ToRecipients: [{ EmailAddress: { Address: 'someone@example.com' } }] });
    add('/me/messages/{message-id}/send', 'post', 'Send a draft');
    add('/me/sendmail', 'post', 'Send mail', { Message: { Subject: 'Hello', Body: { ContentType: 'Text', Content: 'Hello' }, ToRecipients: [{ EmailAddress: { Address: 'someone@example.com' } }] }, SaveToSentItems: true });
  } else if (service === 'outlook-calendar') {
    add('/me/calendars', 'get', 'List calendars');
    for (const root of ['/me/events', '/me/calendars/{calendar-id}/events']) {
      add(root, 'get', 'List events');
      add(root, 'post', 'Create event', { Subject: 'Planning', Start: { DateTime: '2027-01-15T10:00:00', TimeZone: 'UTC' }, End: { DateTime: '2027-01-15T10:30:00', TimeZone: 'UTC' } });
    }
    add('/me/events/{event-id}', 'get', 'Read event');
    add('/me/events/{event-id}', 'patch', 'Update event', { Subject: 'Updated planning' });
    add('/me/events/{event-id}', 'delete', 'Delete event');
    add('/me/calendarview', 'get', 'List events in a date range');
    paths['/me/calendarview'].get.parameters.push(...['startDateTime', 'endDateTime'].map(name => ({ name, in: 'query', required: true, schema: { type: 'string' } })));
  } else if (service === 'microsoft-todo') {
    add('/me/taskfolders', 'get', 'List Outlook task folders');
    add('/me/taskfolders', 'post', 'Create task folder', { Name: 'New list' });
    add('/me/taskfolders/{folder-id}', 'patch', 'Rename task folder', { Name: 'Updated list' });
    add('/me/taskfolders/{folder-id}', 'delete', 'Delete task folder');
    for (const root of ['/me/tasks', '/me/taskfolders/{folder-id}/tasks']) {
      add(root, 'get', 'List tasks'); add(root, 'post', 'Create task', { Subject: 'New task', Body: { ContentType: 'Text', Content: '' } });
    }
    add('/me/tasks/{task-id}', 'get', 'Read task');
    add('/me/tasks/{task-id}', 'patch', 'Update task', { Subject: 'Updated task' });
    add('/me/tasks/{task-id}/complete', 'post', 'Complete task');
    add('/me/tasks/{task-id}', 'delete', 'Delete task');
  }
  return { openapi: '3.0.3', info: { title: 'Outlook web API', version: '2.0', description: 'Private compatibility backend for existing Outlook web sessions. Microsoft has retired the public Outlook REST API; these operations are verified separately against the web-session audience and may change.' }, servers: [{ url: 'https://outlook.office.com/api/v2.0' }], paths };
}
