import { AdapterError } from './errors.js';

export function recipientEmail(value) {
  if (typeof value !== 'string' || value.length > 254 || !/^[A-Za-z0-9.!#$%&'*+/=?^_`{|}~-]+@[A-Za-z0-9](?:[A-Za-z0-9.-]*[A-Za-z0-9])?$/.test(value.trim()) || value.includes(',')) {
    throw new AdapterError('invalid_recipient', 'email must be one exact Teams sign-in email/UPN, not a display name or a list of people.', 400);
  }
  return value.trim().toLowerCase();
}

export function peopleQuery(value) {
  if (typeof value !== 'string' || value.trim().length < 2 || value.length > 150 || /[\u0000-\u001f]/.test(value)) throw new AdapterError('invalid_query', 'query must contain 2–150 characters without control characters.', 400);
  return value.trim();
}

export function messageQuery(value) {
  if (typeof value !== 'string' || !value.trim() || value.length > 500 || /[\u0000-\u001f]/.test(value)) throw new AdapterError('invalid_query', 'query must contain 1–500 characters without control characters.', 400);
  return value.trim();
}
export function messageId(value) {
  if (typeof value !== 'string' || !value || value.length > 512 || /[\u0000-\u001f]/.test(value)) throw new AdapterError('invalid_message_id', 'messageId must be an exact message ID returned by readTeamsMessages.', 400);
  return value;
}

export function messageText(value, field = 'text', allowEmpty = false) {
  if (typeof value !== 'string' || (!allowEmpty && !value.trim()) || value.length > 20_000 || /[\u0000-\u0008\u000b\u000c\u000e-\u001f]/.test(value)) throw new AdapterError('invalid_message', `${field} must be plain text, at most 20,000 characters${allowEmpty ? '' : ', and nonempty'}.`, 400);
  return value;
}
export function reactionInput(value, selected) {
  if (typeof value !== 'string' || !/^[a-z][a-z0-9_-]{0,63}$/.test(value) || typeof selected !== 'boolean') throw new AdapterError('invalid_reaction', 'reaction must be an ID returned by listTeamsMessageReactions and selected must be true or false.', 400);
  return { reaction: value === 'yes' ? 'like' : value, selected };
}
