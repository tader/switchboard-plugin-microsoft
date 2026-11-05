export class AdapterError extends Error {
  constructor(code, message, status = 409) {
    super(message);
    this.code = code;
    if (["authentication_required","token_import_required","signin_required","signin_failed","account_mismatch","teams_signin_required","invalid_credentials"].includes(code)) this.reauthenticationRequired=true;
    this.status = status;
  }
}
