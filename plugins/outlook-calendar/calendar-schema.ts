// Canonical, provider-independent schemas. Vendored by adapter releases until an SDK is published.
import spec from "./calendar-v1.json" with { type: "json" };
export const calendarOperations: Record<
  string,
  {
    inputSchema: Record<string, unknown>;
    outputSchema: Record<string, unknown>;
  }
> = Object.fromEntries(
  Object.entries(spec.operations).map(([name, op]) => [
    name,
    {
      inputSchema: { ...op.input, definitions: spec.definitions },
      outputSchema: { ...op.output, definitions: spec.definitions },
    },
  ]),
);
