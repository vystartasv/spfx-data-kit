import { DataError, headerValue, parseJson } from "./errors.js";
import { DataClient } from "./client.js";
import type { DataResult, RequestHeaders } from "./contracts.js";

export interface GraphRequestOptions { readonly method?: string; readonly headers?: RequestHeaders; readonly body?: unknown; readonly etag?: string; readonly signal?: AbortSignal; readonly timeoutMs?: number; }
export interface GraphBatchRequest { readonly id: string; readonly method: string; readonly url: string; readonly headers?: RequestHeaders; readonly body?: unknown; }
export interface GraphBatchResponse { readonly id: string; readonly status: number; readonly headers: RequestHeaders; readonly body?: unknown; readonly ok: boolean; }
export interface GraphBatchResult { readonly responses: readonly GraphBatchResponse[]; readonly failures: readonly GraphBatchResponse[]; }
export interface GraphPage<T> { readonly value: readonly T[]; readonly nextLink?: string; readonly deltaLink?: string; }
export interface GraphDelta<T> { readonly value: readonly T[]; readonly deltaLink?: string; }
export interface GraphIterationOptions { readonly maxPages?: number; readonly maxItems?: number; }
export interface GraphAdapterOptions { readonly baseUrl?: string; }

const jsonHeaders = { Accept: "application/json", "Content-Type": "application/json" };
const batchMethods = new Set(["GET", "POST", "PUT", "PATCH", "DELETE"]);
const asRecord = (value: unknown): Record<string, unknown> => typeof value === "object" && value !== null ? value as Record<string, unknown> : {};
const etagOf = (value: unknown): string | undefined => {
  const row = asRecord(value);
  const etag = row["@odata.etag"] ?? row["odata.etag"];
  return typeof etag === "string" ? etag : undefined;
};
const absolute = (base: string, path: string): string => /^[a-z][a-z\d+.-]*:/i.test(path) ? path : `${base.replace(/\/+$/, "")}/${path.replace(/^\/+/, "")}`;
const pagePayload = <T>(value: unknown): GraphPage<T> => {
  const root = asRecord(value);
  const rows = Array.isArray(root.value) ? root.value as T[] : [];
  const nextLink = typeof root["@odata.nextLink"] === "string" ? root["@odata.nextLink"] : undefined;
  const deltaLink = typeof root["@odata.deltaLink"] === "string" ? root["@odata.deltaLink"] : undefined;
  return { value: rows, ...(nextLink === undefined ? {} : { nextLink }), ...(deltaLink === undefined ? {} : { deltaLink }) };
};
const validateIteration = (options: GraphIterationOptions): void => {
  for (const [name, value] of [["maxPages", options.maxPages], ["maxItems", options.maxItems]] as const) {
    if (value !== undefined && (!Number.isInteger(value) || value < 0)) throw new DataError("validation", `${name} must be a non-negative integer`);
  }
};

export class GraphAdapter {
  private readonly baseUrl: string;
  constructor(private readonly client: DataClient, options: GraphAdapterOptions = {}) { this.baseUrl = options.baseUrl ?? "https://graph.microsoft.com/v1.0"; }

  async request<T>(path: string, options: GraphRequestOptions = {}): Promise<DataResult<T>> {
    const method = (options.method ?? "GET").toUpperCase();
    const headers: RequestHeaders = { ...jsonHeaders, ...options.headers, ...(options.etag === undefined ? {} : { "If-Match": options.etag }) };
    const url = absolute(this.baseUrl, path);
    if (method === "GET") {
      const response = await this.client.requestRaw({ url, method, headers, signal: options.signal, timeoutMs: options.timeoutMs }, { cache: true });
      const value = response.text.trim() ? parseJson<T>(response.text, response.status) : undefined as T;
      return { data: value, etag: etagOf(value) ?? headerValue(response.headers, "etag") };
    }
    const response = await this.client.requestRaw({ url, method, headers, body: options.body === undefined ? undefined : JSON.stringify(options.body), signal: options.signal, timeoutMs: options.timeoutMs });
    const value = response.text.trim() ? parseJson<T>(response.text, response.status) : undefined as T;
    return { data: value, etag: etagOf(value) ?? headerValue(response.headers, "etag") };
  }

  async page<T>(path: string, options: GraphRequestOptions = {}): Promise<GraphPage<T>> {
    const result = await this.request<unknown>(path, options);
    return pagePayload<T>(result.data);
  }

  async *pages<T>(path: string, options: GraphIterationOptions & GraphRequestOptions = {}): AsyncIterable<GraphPage<T>> {
    validateIteration(options);
    const maxPages = options.maxPages ?? 100;
    const maxItems = options.maxItems ?? Number.MAX_SAFE_INTEGER;
    let url = path;
    let pages = 0;
    let items = 0;
    const seen = new Set<string>();
    while (pages < maxPages && items < maxItems && !seen.has(url)) {
      seen.add(url);
      const page = await this.page<T>(url, options);
      const value = page.value.slice(0, maxItems - items);
      items += value.length;
      pages++;
      yield { ...page, value };
      if (value.length < page.value.length || !page.nextLink) break;
      url = page.nextLink;
    }
  }

  async *iterate<T>(path: string, options: GraphIterationOptions & GraphRequestOptions = {}): AsyncIterable<T> {
    for await (const page of this.pages<T>(path, options)) yield* page.value;
  }

  async delta<T>(path: string, options: GraphIterationOptions & GraphRequestOptions = {}): Promise<GraphDelta<T>> {
    const value: T[] = [];
    let deltaLink: string | undefined;
    for await (const page of this.pages<T>(path, options)) {
      value.push(...page.value);
      deltaLink = page.deltaLink ?? deltaLink;
    }
    return deltaLink === undefined ? { value } : { value, deltaLink };
  }

  async batch(requests: readonly GraphBatchRequest[]): Promise<GraphBatchResult> {
    if (requests.length > 20) throw new DataError("validation", "Graph JSON batches support at most 20 requests");
    const ids = new Set<string>();
    for (const request of requests) {
      if (!request.id || !request.id.trim() || ids.has(request.id)) throw new DataError("validation", "Graph batch request ids must be non-empty and unique");
      if (!request.url || !request.url.trim() || request.url.startsWith("//") || /^[a-z][a-z\d+.-]*:/i.test(request.url)) throw new DataError("validation", "Graph batch child URLs must be relative");
      if (!batchMethods.has(request.method.toUpperCase())) throw new DataError("validation", "Graph batch request methods are invalid");
      ids.add(request.id);
    }
    const response = await this.client.requestRaw({
      url: absolute(this.baseUrl, "$batch"),
      method: "POST",
      headers: { ...jsonHeaders },
      body: JSON.stringify({ requests: requests.map(({ id, method, url, headers, body }) => ({ id, method: method.toUpperCase(), url, ...(headers ? { headers } : {}), ...(body === undefined ? {} : { body }) })) }),
    });
    const payload = asRecord(parseJson<unknown>(response.text, response.status));
    if (!Array.isArray(payload.responses)) throw new DataError("unknown", "Graph batch response did not contain responses", payload, response.status);
    const responseIds = new Set<string>();
    const responses = payload.responses.map((value): GraphBatchResponse => {
      const item = asRecord(value);
      const id = typeof item.id === "string" ? item.id : "";
      const status = typeof item.status === "number" ? item.status : 0;
      if (!id || !ids.has(id) || responseIds.has(id) || !Number.isInteger(status) || status < 100 || status > 599) throw new DataError("unknown", "Graph batch response contained an invalid child response", payload, response.status);
      responseIds.add(id);
      const headers = asRecord(item.headers) as RequestHeaders;
      return { id, status, headers, body: item.body, ok: status >= 200 && status < 300 };
    });
    if (responseIds.size !== ids.size) throw new DataError("unknown", "Graph batch response did not contain every child response", payload, response.status);
    return { responses, failures: responses.filter((item) => !item.ok) };
  }
}
