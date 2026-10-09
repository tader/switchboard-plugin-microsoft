export interface PluginManifest {
    id: string;
    name: string;
    version: string;
    description?: string;
    /** Entry module relative to the plugin directory. Defaults to index.ts, then index.js. */
    main?: string;
    /** Ids of plugins whose exports this plugin uses via ctx.require(). */
    dependencies?: string[];
    /** Optional npm semver ranges, keyed by an id declared in dependencies. */
    dependencyVersions?: Record<string, string>;
    /** Instance-wide settings configured by an administrator (e.g. an OAuth client). */
    settings?: Field[];
    /** Icon file relative to the plugin directory (svg/png). */
    icon?: string;
}
export interface Field {
    key: string;
    label: string;
    type?: 'text' | 'secret' | 'url' | 'textarea' | 'select' | 'boolean';
    description?: string;
    placeholder?: string;
    required?: boolean;
    default?: string | boolean;
    options?: {
        value: string;
        label: string;
    }[];
    /** Shown under "Advanced" in forms. */
    advanced?: boolean;
}
export interface Logger {
    info(...args: unknown[]): void;
    warn(...args: unknown[]): void;
    error(...args: unknown[]): void;
}
export interface PluginContext {
    manifest: PluginManifest;
    /** Values of the manifest's `settings`, as configured by an administrator. */
    settings: Record<string, any>;
    /** Exports of a plugin listed in `dependencies`. */
    require<T = any>(pluginId: string): T;
    log: Logger;
    /** Public base URL of Switchboard, without trailing slash. */
    publicUrl: string;
    /** The single OAuth redirect URI of Switchboard. Register this at providers. */
    callbackUrl: string;
    /** Absolute path of the plugin's directory. */
    dir: string;
    /** Persistent directory for this plugin's own files. */
    dataDir: string;
    /** True when this instance has a peer connection configured. */
    peer: boolean;
    /** Core MCP OAuth discovery and encrypted connect-flow integration. */
    mcp: {
        oauth(): AuthMethod;
    };
}
export interface PluginInstance {
    services?: ServiceDefinition[];
    /** Made available to dependent plugins via ctx.require(). */
    exports?: unknown;
    dispose?(): void | Promise<void>;
}
export type PluginSetup = (ctx: PluginContext) => PluginInstance | Promise<PluginInstance>;
export interface AccountInfo {
    /** Stable id at the provider. Reconnecting the same account updates the existing connection. */
    id?: string;
    /** Human readable, e.g. an email address or username. */
    label: string;
    avatarUrl?: string;
}
/** What Switchboard knows about one connected account. */
export type ConnectionKind = 'http' | 'mcp';
export interface Connection {
    /** Omitted by older plugins; defaults to HTTP. */
    kind?: ConnectionKind;
    id: string;
    name: string;
    serviceId: string;
    methodId: string;
    /** Values of the auth method's fields entered by the user. */
    config: Record<string, any>;
    credentials: any;
    account?: AccountInfo;
}
export interface OutgoingRequest {
    method: string;
    url: URL;
    headers: Headers;
}
export interface Connected {
    credentials: any;
    account?: AccountInfo;
    /** Replaces the user's config (e.g. to drop values that should not be kept). */
    config?: Record<string, any>;
}
export interface DeviceAuthorization {
    userCode: string;
    verificationUri: string;
    verificationUriComplete?: string;
    /** Seconds until the code expires. */
    expiresIn?: number;
    /** Seconds between polls. */
    interval?: number;
}
/**
 * Result of a connect step:
 *  - Connected: done
 *  - { redirect }: send the user's browser there; the provider returns to ctx.callbackUrl
 *  - { device }: show the code to the user and poll() until done
 */
export type ConnectStep = Connected | {
    redirect: string;
    pending?: any;
} | {
    device: DeviceAuthorization;
    pending?: any;
};
export type PollResult = Connected | {
    wait: true;
    pending?: any;
    interval?: number;
};
export interface ConnectArgs {
    config: Record<string, any>;
    callbackUrl: string;
    /** Opaque value to pass as OAuth `state`. */
    state: string;
    /** Set when re-authenticating an existing connection. */
    connection?: Connection;
}
export interface TokenResult {
    accessToken: string;
    tokenType?: string;
    expiresAt?: number;
    /** Updated credentials to persist (e.g. after a refresh). */
    credentials?: any;
}
export interface AuthMethod {
    /** Unique within the service, e.g. "oauth", "device", "token". */
    id: string;
    name: string;
    description?: string;
    /** Asked from the user when connecting with this method. */
    fields?: Field[];
    /** When set, the method is shown but cannot be used (e.g. "Not configured by an administrator"). */
    unavailable?: string;
    connect(args: ConnectArgs): Promise<ConnectStep> | ConnectStep;
    /** Completes a redirect flow. `params` are the query parameters the provider returned. */
    callback?(args: ConnectArgs & {
        pending: any;
        params: Record<string, string>;
    }): Promise<Connected>;
    /** Advances a device flow. */
    poll?(args: ConnectArgs & {
        pending: any;
    }): Promise<PollResult>;
    /**
     * Adds credentials to an outgoing request. `force` is set when a previous attempt got a 401,
     * so cached tokens should be refreshed. Return new credentials to persist them.
     */
    authorize(req: OutgoingRequest, conn: Connection, opts: {
        force?: boolean;
        signal?: AbortSignal;
    }): Promise<void | {
        credentials?: any;
    }> | void | {
        credentials?: any;
    };
    /** Hands out a usable bearer token, for scripts that use a provider SDK directly. */
    token?(conn: Connection, opts: {
        force?: boolean;
        signal?: AbortSignal;
    }): Promise<TokenResult>;
    /** Called when the connection is deleted. */
    revoke?(conn: Connection): Promise<void>;
}
export interface ServiceDefinition {
    /** Omitted by existing plugins; defaults to HTTP. */
    kind?: ConnectionKind;
    /** Globally unique, e.g. "gmail". */
    id: string;
    name: string;
    description?: string;
    /** SVG markup, a data: URI, an https URL, or a file relative to the plugin directory. */
    icon?: string;
    docsUrl?: string;
    /** Relative request URLs resolve against this. */
    baseUrl?: string | ((conn: Connection) => string | undefined);
    /**
     * Hosts credentials may be sent to. Entries may start with "*." for subdomains.
     * Defaults to the host of baseUrl.
     */
    allowedHosts?: string[] | ((conn: Connection) => string[]);
    /**
     * OpenAPI (3.x) or Swagger (2.0) document: a URL, or the parsed document.
     * When its server is a placeholder or not an allowed host, the connection's base URL is used instead.
     */
    openapi?: string | object | ((conn: Connection) => string | object | undefined | Promise<string | object | undefined>);
    authMethods: AuthMethod[];
}
