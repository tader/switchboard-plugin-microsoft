import { AdapterError } from './errors.js';

/** A browser launched on the Switchboard host needs instructions, not a device code. */
export function browserAuthorization(capabilities, expiresIn = 600) {
  if (!capabilities?.deviceBrowser) throw new AdapterError('core_upgrade_required', 'Upgrade Switchboard to use browser sign-in.', 503);
  return {
    browser: true,
    instructions: 'Microsoft sign-in is starting. If needed, a browser window will open automatically on the computer running this Switchboard. Complete sign-in there; the window closes when authentication finishes.',
    expiresIn,
    interval: 3,
  };
}
