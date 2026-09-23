import { DataError, mapHttpError } from "./errors.js";
import type { DataRequest, RequestHeaders, RequestTransport, ResponseHeaders, TransportResponse } from "./contracts.js";

export interface CacheOptions { readonly maxEntries: number; readonly ttlMs: number; }
export interface RetryOptions { readonly maxRetries?: number; readonly sleep?: (milliseconds: number) => Promise<void>; readonly now?: () => number; }
export interface DataClientOptions { readonly cache?: CacheOptions; readonly retry?: RetryOptions; readonly signal?: AbortSignal; readonly timeoutMs?: number; }
export interface ClientDiagnostics { readonly requests: number; readonly cacheHits: number; readonly deduplicated: number; readonly retries: number; readonly failures: number; }
export interface RawDataResponse { readonly status: number; readonly headers: ResponseHeaders; readonly text: string; readonly bytes?: Uint8Array; }

type CacheEntry = { expiresAt: number; value: unknown };
const defaultSleep = (milliseconds: number): Promise<void> => new Promise((resolve) => setTimeout(resolve, milliseconds));

function validateCache(options: CacheOptions): void {
  if (!Number.isInteger(options.maxEntries) || options.maxEntries < 1) throw new DataError("validation", "cache.maxEntries must be a positive integer");
  if (!Number.isFinite(options.ttlMs) || options.ttlMs <= 0) throw new DataError("validation", "cache.ttlMs must be positive");
}
function validateRetries(options: RetryOptions | undefined): void {
  if (options?.maxRetries !== undefined && (!Number.isInteger(options.maxRetries) || options.maxRetries < 0)) throw new DataError("validation", "retry.maxRetries must be a non-negative integer");
}
function validateTimeout(timeoutMs: number | undefined): void {
  if (timeoutMs !== undefined && (!Number.isInteger(timeoutMs) || timeoutMs < 0)) throw new DataError("validation", "timeoutMs must be a non-negative integer");
}
function isAbortError(cause: unknown): boolean {
  return typeof cause === "object" && cause !== null && "name" in cause && (cause as { name?: unknown }).name === "AbortError";
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
function responseHeaders(response: TransportResponse): ResponseHeaders { return response.headers ?? {}; }

export class DataClient {
  private readonly cache?: CacheOptions;
  private readonly signal?: AbortSignal;
  private readonly timeoutMs?: number;
  private readonly retry: Required<Pick<RetryOptions, "maxRetries" | "sleep" | "now">>;
  private readonly entries = new Map<string, CacheEntry>();
  private readonly inflight = new Map<string, Promise<unknown>>();
  private cacheVersion = 0;
  private counters: ClientDiagnostics = { requests: 0, cacheHits: 0, deduplicated: 0, retries: 0, failures: 0 };

  constructor(private readonly transport: RequestTransport, options: DataClientOptions = {}) {
    if (options.cache) validateCache(options.cache);
    validateRetries(options.retry);
    validateTimeout(options.timeoutMs);
    this.cache = options.cache;
    this.signal = options.signal;
    this.timeoutMs = options.timeoutMs;
    this.retry = { maxRetries: options.retry?.maxRetries ?? 2, sleep: options.retry?.sleep ?? defaultSleep, now: options.retry?.now ?? Date.now };
  }
  diagnostics(): ClientDiagnostics { return { ...this.counters }; }
  async get<T>(url: string, options: Omit<DataRequest, "url" | "method"> = {}): Promise<T> { return this.request<T>({ ...options, url, method: "GET" }); }
  async requestBytes(url: string, options: Omit<DataRequest, "url" | "method" | "responseType"> = {}): Promise<Uint8Array> {
    const response = await this.requestRaw({ ...options, url, method: "GET", responseType: "binary" });
    if (!response.bytes) throw new DataError("unknown", "The transport did not provide a binary response", undefined, response.status);
    return response.bytes;
  }

  async request<T>(request: DataRequest): Promise<T> {
    const normalized = { ...request, method: (request.method ?? "GET").toUpperCase() };
    if (normalized.method !== "GET") return this.parse<T>(await this.requestRaw(normalized));
    return this.parse<T>(await this.requestRaw(normalized, { cache: true }));
  }

  async requestRaw(request: DataRequest, options: { readonly cache?: boolean } = {}): Promise<RawDataResponse> {
    const method = (request.method ?? "GET").toUpperCase();
    const normalized = { ...request, method };
    validateTimeout(request.timeoutMs);
    if (options.cache && method === "GET") return this.cachedGet(normalized);
    return this.executeRaw(normalized);
  }

  private async cachedGet(request: DataRequest): Promise<RawDataResponse> {
    const key = requestCacheKey(request);
    const now = this.retry.now();
    const cached = this.entries.get(key);
    if (cached && cached.expiresAt > now) { this.counters = { ...this.counters, cacheHits: this.counters.cacheHits + 1 }; return cached.value as RawDataResponse; }
    if (cached) this.entries.delete(key);
    const pending = this.inflight.get(key);
    if (pending) { this.counters = { ...this.counters, deduplicated: this.counters.deduplicated + 1 }; return pending as Promise<RawDataResponse>; }
    const cacheVersion = this.cacheVersion;
    const operation = this.executeRaw(request);
    this.inflight.set(key, operation);
    try {
      const response = await operation;
      if (this.cache && cacheVersion === this.cacheVersion) {
        this.entries.set(key, { value: response, expiresAt: this.retry.now() + this.cache.ttlMs });
        while (this.entries.size > this.cache.maxEntries) this.entries.delete(this.entries.keys().next().value as string);
      }
      return response;
    } finally { if (this.inflight.get(key) === operation) this.inflight.delete(key); }
  }

  private async executeRaw(request: DataRequest): Promise<RawDataResponse> {
    const method = (request.method ?? "GET").toUpperCase();
    const canRetry = method === "GET";
    const signal = request.signal ?? this.signal;
    const timeoutMs = request.timeoutMs ?? this.timeoutMs;
    if (signal?.aborted) throw signal.reason ?? new DOMException("The request was aborted", "AbortError");
    let attempt = 0;
    while (true) {
      this.counters = { ...this.counters, requests: this.counters.requests + 1 };
      let response: TransportResponse;
      try {
        response = await this.transport.request(request.url, { method, headers: request.headers, body: request.body, responseType: request.responseType, signal, timeoutMs });
      } catch (cause) {
        if (signal?.aborted) {
          this.counters = { ...this.counters, failures: this.counters.failures + 1 };
          throw signal.reason ?? cause;
        }
        if (isAbortError(cause) || !canRetry || attempt >= this.retry.maxRetries) {
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
      let bytes: Uint8Array | undefined;
      let text: string;
      if (request.responseType === "binary" && response.arrayBuffer) {
        bytes = new Uint8Array(await response.arrayBuffer());
        text = response.status >= 200 && response.status < 300 ? "" : new TextDecoder().decode(bytes);
      } else if (request.responseType === "binary" && response.status >= 200 && response.status < 300) {
        throw new DataError("unknown", "The transport does not support binary responses", undefined, response.status);
      } else {
        text = await response.text();
      }
      if (response.status >= 200 && response.status < 300) return { status: response.status, headers, text, ...(bytes === undefined ? {} : { bytes }) };
      let details: unknown;
      if (text.trim()) {
        try { details = JSON.parse(text); } catch { /* keep the response body out of structured details */ }
      }
      const error = mapHttpError(response.status, { status: response.status }, headers, this.retry.now, details);
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
    return this.invalidateMatching((key) => (exact !== undefined && key === exact) || (prefix !== undefined && key.startsWith(prefix)));
  }
  invalidateKey(key: string): number { return this.invalidate(key, "key"); }
  invalidatePrefix(prefix: string): number { return this.invalidate(prefix, "prefix"); }
  invalidateUrl(url: string): number {
    return this.invalidateMatching((key) => this.cacheUrl(key) === url);
  }
  invalidateUrlPrefix(prefix: string): number {
    return this.invalidateMatching((key) => this.cacheUrl(key)?.startsWith(prefix) ?? false);
  }
  clearCache(): number { const count = this.entries.size; this.entries.clear(); this.cacheVersion++; return count; }
  private invalidateMatching(matches: (key: string) => boolean): number {
    this.cacheVersion++;
    let removed = 0;
    for (const key of this.entries.keys()) if (matches(key)) { this.entries.delete(key); removed++; }
    for (const key of this.inflight.keys()) if (matches(key)) this.inflight.delete(key);
    return removed;
  }
  private cacheUrl(key: string): string | undefined {
    try { return String((JSON.parse(key) as unknown[])[1]); } catch { return undefined; }
  }
  private parse<T>(response: RawDataResponse): T {
    if (response.status === 204 || response.text.trim() === "") return undefined as T;
    try { return JSON.parse(response.text) as T; } catch (cause) { throw new DataError("unknown", "The service returned invalid JSON", cause, response.status); }
  }
}
