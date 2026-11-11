// Canonical Email contract; adapter snapshots must match this JSON revision.
import spec from "./email-v1.json" with { type: "json" };
export const emailOperations: Record<
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
