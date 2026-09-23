import { DataError, mapHttpError } from "./errors.js";
import type { DataRequest, RequestHeaders, RequestTransport, TransportResponse } from "./contracts.js";

export interface CacheOptions { readonly maxEntries: number; readonly ttlMs: number; }
export interface RetryOptions { readonly maxRetries?: number; readonly sleep?: (milliseconds: number) => Promise<void>; readonly now?: () => number; }
export interface DataClientOptions { readonly cache?: CacheOptions; readonly retry?: RetryOptions; }
export interface ClientDiagnostics { readonly requests: number; readonly cacheHits: number; readonly deduplicated: number; readonly retries: number; readonly failures: number; }
export interface RawDataResponse { readonly status: number; readonly headers: RequestHeaders; readonly text: string; }

type CacheEntry = { expiresAt: number; value: unknown };
const defaultSleep = (milliseconds: number): Promise<void> => new Promise((resolve) => setTimeout(resolve, milliseconds));

function validateCache(options: CacheOptions): void {
  if (!Number.isInteger(options.maxEntries) || options.maxEntries < 1) throw new DataError("validation", "cache.maxEntries must be a positive integer");
  if (!Number.isFinite(options.ttlMs) || options.ttlMs <= 0) throw new DataError("validation", "cache.ttlMs must be positive");
}
function validateRetries(options: RetryOptions | undefined): void {
  if (options?.maxRetries !== undefined && (!Number.isInteger(options.maxRetries) || options.maxRetries < 0)) throw new DataError("validation", "retry.maxRetries must be a non-negative integer");
}
function headerEntries(headers: RequestHeaders | undefined, selected: readonly string[] | undefined): string[] {
  if (!headers) return [];
  const names = selected ?? Object.keys(headers);
  return [...new Set(names.map((name) => name.toLowerCase()))].sort().map((name) => {
    const source = Object.keys(headers).find((candidate) => candidate.toLowerCase() === name);
    return `${name}=${source === undefined ? "" : headers[source]}`;
  });
}
export function requestCacheKey(request: DataRequest): string {
  const method = (request.method ?? "GET").toUpperCase();
  return JSON.stringify([method, request.url, request.body ?? "", headerEntries(request.headers, request.cacheKeyHeaders)]);
}
function responseHeaders(response: TransportResponse): RequestHeaders { return response.headers ?? {}; }

export class DataClient {
  private readonly cache?: CacheOptions;
  private readonly retry: Required<Pick<RetryOptions, "maxRetries" | "sleep" | "now">>;
  private readonly entries = new Map<string, CacheEntry>();
  private readonly inflight = new Map<string, Promise<unknown>>();
  private counters: ClientDiagnostics = { requests: 0, cacheHits: 0, deduplicated: 0, retries: 0, failures: 0 };

  constructor(private readonly transport: RequestTransport, options: DataClientOptions = {}) {
    if (options.cache) validateCache(options.cache);
    validateRetries(options.retry);
    this.cache = options.cache;
    this.retry = { maxRetries: options.retry?.maxRetries ?? 2, sleep: options.retry?.sleep ?? defaultSleep, now: options.retry?.now ?? Date.now };
  }
  diagnostics(): ClientDiagnostics { return { ...this.counters }; }
  async get<T>(url: string, options: Omit<DataRequest, "url" | "method"> = {}): Promise<T> { return this.request<T>({ ...options, url, method: "GET" }); }

  async request<T>(request: DataRequest): Promise<T> {
    const normalized = { ...request, method: (request.method ?? "GET").toUpperCase() };
    if (normalized.method !== "GET") return this.parse<T>(await this.requestRaw(normalized));
    const key = requestCacheKey(normalized);
    const now = this.retry.now();
    const cached = this.entries.get(key);
    if (cached && cached.expiresAt > now) { this.counters = { ...this.counters, cacheHits: this.counters.cacheHits + 1 }; return cached.value as T; }
    if (cached) this.entries.delete(key);
    const pending = this.inflight.get(key);
    if (pending) { this.counters = { ...this.counters, deduplicated: this.counters.deduplicated + 1 }; return pending as Promise<T>; }
    const operation = this.requestRaw(normalized).then((response) => this.parse<T>(response));
    this.inflight.set(key, operation);
    try {
      const value = await operation;
      if (this.cache) {
        this.entries.set(key, { value, expiresAt: this.retry.now() + this.cache.ttlMs });
        while (this.entries.size > this.cache.maxEntries) this.entries.delete(this.entries.keys().next().value as string);
      }
      return value;
    } finally { this.inflight.delete(key); }
  }

  async requestRaw(request: DataRequest): Promise<RawDataResponse> {
    const method = (request.method ?? "GET").toUpperCase();
    const canRetry = method === "GET";
    let attempt = 0;
    while (true) {
      this.counters = { ...this.counters, requests: this.counters.requests + 1 };
      let response: TransportResponse;
      try {
        response = await this.transport.request(request.url, { method, headers: request.headers, body: request.body });
      } catch (cause) {
        if (!canRetry || attempt >= this.retry.maxRetries) {
          this.counters = { ...this.counters, failures: this.counters.failures + 1 };
          if (cause instanceof DataError) throw cause;
          throw new DataError("unknown", "The transport failed before receiving a response", cause);
        }
        attempt++;
        this.counters = { ...this.counters, retries: this.counters.retries + 1 };
        await this.retry.sleep(Math.min(1000 * 2 ** (attempt - 1), 8000));
        continue;
      }
      const headers = responseHeaders(response);
      const text = await response.text();
      if (response.status >= 200 && response.status < 300) return { status: response.status, headers, text };
      const error = mapHttpError(response.status, { status: response.status }, headers, this.retry.now);
      const retryable = canRetry && (response.status === 408 || response.status === 429 || response.status >= 500);
      if (!retryable || attempt >= this.retry.maxRetries) { this.counters = { ...this.counters, failures: this.counters.failures + 1 }; throw error; }
      attempt++;
      this.counters = { ...this.counters, retries: this.counters.retries + 1 };
      await this.retry.sleep(error.retryAfterMs ?? Math.min(1000 * 2 ** (attempt - 1), 8000));
    }
  }

  invalidate(target: string | { readonly key?: string; readonly prefix?: string }, mode: "key" | "prefix" = "prefix"): number {
    const exact = typeof target === "string" && mode === "key" ? target : typeof target === "object" ? target.key : undefined;
    const prefix = typeof target === "string" && mode === "prefix" ? target : typeof target === "object" ? target.prefix : undefined;
    if (exact === undefined && prefix === undefined) throw new DataError("validation", "invalidate requires a key or prefix");
    let removed = 0;
    for (const key of this.entries.keys()) if ((exact !== undefined && key === exact) || (prefix !== undefined && key.startsWith(prefix))) { this.entries.delete(key); removed++; }
    return removed;
  }
  invalidateKey(key: string): number { return this.invalidate(key, "key"); }
  invalidatePrefix(prefix: string): number { return this.invalidate(prefix, "prefix"); }
  invalidateUrl(url: string): number {
    let removed = 0;
    for (const key of this.entries.keys()) { try { if ((JSON.parse(key) as unknown[])[1] === url) { this.entries.delete(key); removed++; } } catch { /* internal key */ } }
    return removed;
  }
  invalidateUrlPrefix(prefix: string): number {
    let removed = 0;
    for (const key of this.entries.keys()) {
      try { if (String((JSON.parse(key) as unknown[])[1]).startsWith(prefix)) { this.entries.delete(key); removed++; } } catch { /* internal key */ }
    }
    return removed;
  }
  clearCache(): number { const count = this.entries.size; this.entries.clear(); return count; }
  private parse<T>(response: RawDataResponse): T {
    if (response.status === 204 || response.text.trim() === "") return undefined as T;
    try { return JSON.parse(response.text) as T; } catch (cause) { throw new DataError("unknown", "The service returned invalid JSON", cause, response.status); }
  }
}
