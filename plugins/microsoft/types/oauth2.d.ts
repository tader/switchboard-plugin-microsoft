import type { AccountInfo, AuthMethod, Connection, Field, PluginContext } from './api.d.ts';
type Cfg = Record<string, any>;
type Value<T> = T | ((config: Cfg) => T);
export interface OAuthCredentials {
    accessToken: string;
    refreshToken?: string;
    /** Epoch milliseconds. */
    expiresAt?: number;
    tokenType?: string;
    scope?: string;
    idToken?: string;
}
export interface OAuthOptions {
    id?: string;
    name?: string;
    description?: string;
    unavailable?: string;
    fields?: Field[];
    clientId: Value<string | undefined>;
    clientSecret?: Value<string | undefined>;
    tokenUrl: Value<string>;
    scopes?: Value<string[]>;
    scopeSeparator?: string;
    /** How the client authenticates at the token endpoint. Default: "post". */
    tokenAuth?: Value<'post' | 'basic' | undefined>;
    extraTokenHeaders?: Record<string, string>;
    /** Looks up who signed in. */
    identify?(credentials: OAuthCredentials, config: Cfg, raw: any): Promise<AccountInfo | undefined> | AccountInfo | undefined;
    revoke?(credentials: OAuthCredentials, config: Cfg): Promise<void>;
}
export interface AuthorizationCodeOptions extends OAuthOptions {
    authorizeUrl: Value<string>;
    authorizeParams?: Value<Record<string, string>>;
    /** Default: true. */
    pkce?: Value<boolean>;
}
export interface DeviceCodeOptions extends OAuthOptions {
    deviceAuthorizationUrl: Value<string>;
}
export declare class OAuthError extends Error {
    code: string;
    constructor(code: string, description?: string);
}
/** POSTs to a token endpoint and returns the parsed response. Throws OAuthError on an error response. */
export declare function tokenRequest(o: OAuthOptions, cfg: Cfg, params: Record<string, string>): Promise<any>;
export declare function toCredentials(data: any, previous?: OAuthCredentials): OAuthCredentials;
/**
 * Returns credentials that are valid for at least another minute, refreshing when needed.
 * `refreshed` is true when the caller should persist them.
 */
export declare function freshCredentials(o: OAuthOptions, conn: Connection, force?: boolean, renew?: () => Promise<OAuthCredentials>): Promise<{
    credentials: OAuthCredentials;
    refreshed: boolean;
}>;
/** The browser-based authorization code flow, with PKCE. */
export declare function authorizationCode(o: AuthorizationCodeOptions): AuthMethod;
/** RFC 8628 device authorization: the user enters a code on another device. */
export declare function deviceCode(o: DeviceCodeOptions): AuthMethod;
/** Machine-to-machine: tokens are requested with the client's own credentials. */
export declare function clientCredentials(o: OAuthOptions & {
    label?: Value<string | undefined>;
}): AuthMethod;
export default function setup(_ctx: PluginContext): {
    services: {
        id: string;
        name: string;
        description: string;
        icon: string;
        baseUrl: (conn: Connection) => any;
        allowedHosts: (conn: Connection) => string[];
        openapi: (conn: Connection) => any;
        authMethods: AuthMethod[];
    }[];
    exports: {
        authorizationCode: typeof authorizationCode;
        deviceCode: typeof deviceCode;
        clientCredentials: typeof clientCredentials;
        tokenRequest: typeof tokenRequest;
        toCredentials: typeof toCredentials;
        freshCredentials: typeof freshCredentials;
        OAuthError: typeof OAuthError;
    };
};
export {};
