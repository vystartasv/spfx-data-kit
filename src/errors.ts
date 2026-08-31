import { DataError } from "./contracts.js";
export function mapSharePointError(cause: unknown): DataError {
  const status = typeof cause === "object" && cause !== null && "status" in cause ? (cause as { status?: unknown }).status : undefined;
  if (status === 404) return new DataError("not-found", "SharePoint item was not found", cause);
  if (status === 401 || status === 403) return new DataError("permission", "SharePoint permission denied", cause);
  if (status === 409 || status === 412) return new DataError("conflict", "SharePoint ETag conflict", cause);
  if (status === 400) return new DataError("validation", "SharePoint rejected the input", cause);
  if (status === 408 || status === 429 || (typeof status === "number" && status >= 500)) return new DataError("transient", "SharePoint request failed temporarily", cause);
  return new DataError("unknown", "SharePoint request failed", cause);
}
