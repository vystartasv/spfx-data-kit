import type { ErrorKind, RequestHeaders, ResponseHeaders } from "./contracts.js";

export type ErrorDetails = Readonly<Record<string, unknown>>;

export class DataError extends Error {
  readonly name = "DataError";

  constructor(
    public readonly kind: ErrorKind,
    message: string,
    public readonly cause?: unknown,
    public readonly status?: number,
    public readonly retryAfterMs?: number,
    public readonly code?: string,
    public readonly details?: ErrorDetails,
  ) {
    super(message);
    Object.setPrototypeOf(this, new.target.prototype);
  }
}

export function headerValue(headers: ResponseHeaders | undefined, name: string): string | undefined {
  if (!headers) return undefined;
  if (typeof (headers as HeaderCollectionLike).get === "function") return (headers as HeaderCollectionLike).get(name) ?? undefined;
  const wanted = name.toLowerCase();
  const key = Object.keys(headers).find((candidate) => candidate.toLowerCase() === wanted);
  return key === undefined ? undefined : (headers as RequestHeaders)[key];
}

type HeaderCollectionLike = { get(name: string): string | null };

export function retryAfterMilliseconds(value: string | undefined, now = Date.now): number | undefined {
  if (value === undefined) return undefined;
  const seconds = Number(value.trim());
  if (Number.isFinite(seconds) && seconds >= 0) return Math.round(seconds * 1000);
  const date = Date.parse(value);
  return Number.isNaN(date) ? undefined : Math.max(0, date - now());
}

function asDetails(value: unknown): ErrorDetails | undefined {
  return typeof value === "object" && value !== null && !Array.isArray(value) ? value as ErrorDetails : undefined;
}

function detailCode(details: ErrorDetails | undefined): string | undefined {
  const error = asDetails(details?.error);
  const code = error?.code ?? details?.code;
  return typeof code === "string" ? code : undefined;
}

export function mapHttpError(status: number, cause?: unknown, headers?: ResponseHeaders, now = Date.now, details?: unknown): DataError {
  const retryAfter = retryAfterMilliseconds(headerValue(headers, "retry-after"), now);
  const structured = asDetails(details);
  const code = detailCode(structured);
  if (status === 404) return new DataError("not-found", "The requested resource was not found", cause, status, undefined, code, structured);
  if (status === 401 || status === 403) return new DataError("permission", "The request is not permitted", cause, status, undefined, code, structured);
  if (status === 409 || status === 412) return new DataError("conflict", "The resource changed before the request completed", cause, status, undefined, code, structured);
  if (status === 400 || status === 422) return new DataError("validation", "The service rejected the request", cause, status, undefined, code, structured);
  if (status === 429) return new DataError("throttled", "The service throttled the request", cause, status, retryAfter, code, structured);
  if (status === 408 || status >= 500) return new DataError("transient", "The service failed temporarily", cause, status, retryAfter, code, structured);
  return new DataError("unknown", "The service returned an unexpected response", cause, status, retryAfter, code, structured);
}

export function mapSharePointError(cause: unknown): DataError {
  const status = typeof cause === "object" && cause !== null && "status" in cause
    ? (cause as { status?: unknown }).status
    : undefined;
  if (typeof status === "number") return mapHttpError(status, cause);
  return new DataError("unknown", "SharePoint request failed", cause);
}

export function parseJson<T>(text: string, status?: number): T {
  try {
    return JSON.parse(text) as T;
  } catch (error) {
    throw new DataError("unknown", "The service returned invalid JSON", error, status);
  }
}
