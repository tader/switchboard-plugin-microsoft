import type { AuthMethod, Connection, Field, PluginContext } from './api.d.ts';
export interface AuthenticationExports {
 createConsumerAuth(ctx: PluginContext, options: {id:string;fields?:Field[];scopes:string[]|((config:Record<string,any>)=>string[])}): AuthMethod;
 baseUrl(conn:Connection): string;
 backend(conn:Connection): string;
 outlookSpec(service:string): object;
}
