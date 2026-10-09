import type { Field, PluginContext, ServiceDefinition } from './api.d.ts';
type Cfg = Record<string, any>;
export interface MicrosoftServiceOptions {
    id: string;
    name: string;
    description?: string;
    icon?: string;
    docsUrl?: string;
    /** Microsoft Graph permissions, e.g. Mail.ReadWrite. */
    scopes: string[] | ((config: Cfg) => string[]);
    fields?: Field[];
    /** Name of a slim API description in openapi/, built by openapi/build.ts. */
    spec?: string;
}
export default function setup(ctx: PluginContext): {
    services: ServiceDefinition[];
    exports: {
        microsoftService: (o: MicrosoftServiceOptions) => ServiceDefinition;
        spec: (name: string) => object;
    };
};
export {};
