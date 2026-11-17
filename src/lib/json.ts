/** JSON for display: Dates as ISO strings (their toJSON), functions dropped. */
export function toJson(value: unknown): string {
  return JSON.stringify(value, (_k, v: unknown) => (typeof v === "function" ? undefined : v), 2) ?? "undefined";
}
