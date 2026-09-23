import { DataError, headerValue, parseJson } from "./errors.js";
import { DataClient } from "./client.js";
import type { DataResult, RequestHeaders } from "./contracts.js";

export interface GraphRequestOptions { readonly method?: string; readonly headers?: RequestHeaders; readonly body?: unknown; readonly etag?: string; }
export interface GraphBatchRequest { readonly id: string; readonly method: string; readonly url: string; readonly headers?: RequestHeaders; readonly body?: unknown; }
export interface GraphBatchResponse { readonly id: string; readonly status: number; readonly headers: RequestHeaders; readonly body?: unknown; readonly ok: boolean; }
export interface GraphBatchResult { readonly responses: readonly GraphBatchResponse[]; readonly failures: readonly GraphBatchResponse[]; }
export interface GraphAdapterOptions { readonly baseUrl?: string; }

const jsonHeaders = { Accept: "application/json", "Content-Type": "application/json" };
const asRecord = (value: unknown): Record<string, unknown> => typeof value === "object" && value !== null ? value as Record<string, unknown> : {};
const etagOf = (value: unknown): string | undefined => {
  const row = asRecord(value);
  const etag = row["@odata.etag"] ?? row["odata.etag"];
  return typeof etag === "string" ? etag : undefined;
};
const absolute = (base: string, path: string): string => /^[a-z][a-z\d+.-]*:/i.test(path) ? path : `${base.replace(/\/+$/, "")}/${path.replace(/^\/+/, "")}`;

export class GraphAdapter {
  private readonly baseUrl: string;
  constructor(private readonly client: DataClient, options: GraphAdapterOptions = {}) { this.baseUrl = options.baseUrl ?? "https://graph.microsoft.com/v1.0"; }

  async request<T>(path: string, options: GraphRequestOptions = {}): Promise<DataResult<T>> {
    const method = (options.method ?? "GET").toUpperCase();
    const headers: RequestHeaders = { ...jsonHeaders, ...options.headers, ...(options.etag === undefined ? {} : { "If-Match": options.etag }) };
    const url = absolute(this.baseUrl, path);
    if (method === "GET") {
      const value = await this.client.get<T>(url, { headers });
      return { data: value, etag: etagOf(value) };
    }
    const response = await this.client.requestRaw({ url, method, headers, body: options.body === undefined ? undefined : JSON.stringify(options.body) });
    const value = response.text.trim() ? parseJson<T>(response.text, response.status) : undefined as T;
    return { data: value, etag: etagOf(value) ?? headerValue(response.headers, "etag") };
  }

  async batch(requests: readonly GraphBatchRequest[]): Promise<GraphBatchResult> {
    if (requests.length > 20) throw new DataError("validation", "Graph JSON batches support at most 20 requests");
    const ids = new Set<string>();
    for (const request of requests) {
      if (!request.id || ids.has(request.id)) throw new DataError("validation", "Graph batch request ids must be non-empty and unique");
      if (/^[a-z][a-z\d+.-]*:/i.test(request.url)) throw new DataError("validation", "Graph batch child URLs must be relative");
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
    const responses = payload.responses.map((value): GraphBatchResponse => {
      const item = asRecord(value);
      const id = typeof item.id === "string" ? item.id : "";
      const status = typeof item.status === "number" ? item.status : 0;
      const headers = asRecord(item.headers) as RequestHeaders;
      return { id, status, headers, body: item.body, ok: status >= 200 && status < 300 };
    });
    return { responses, failures: responses.filter((item) => !item.ok) };
  }
}
