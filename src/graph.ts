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
export interface GraphDriveMetadata { readonly id: string; readonly driveType?: string; readonly name?: string; readonly webUrl?: string; readonly createdDateTime?: string; readonly lastModifiedDateTime?: string; readonly owner?: unknown; readonly quota?: unknown; readonly [key: string]: unknown; }
export interface GraphDriveItemMetadata { readonly id: string; readonly name?: string; readonly size?: number; readonly eTag?: string; readonly cTag?: string; readonly webUrl?: string; readonly createdDateTime?: string; readonly lastModifiedDateTime?: string; readonly file?: unknown; readonly folder?: unknown; readonly parentReference?: unknown; readonly [key: string]: unknown; }
export interface GraphDriveItemUpdate { readonly name?: string; readonly description?: string; readonly fileSystemInfo?: { readonly createdDateTime?: string; readonly lastModifiedDateTime?: string }; readonly [key: string]: unknown; }
export type GraphDriveRequestOptions = GraphRequestOptions;
export interface GraphDriveWriteOptions extends GraphRequestOptions { readonly etag?: string; }
export interface GraphSiteMetadata { readonly id: string; readonly name?: string; readonly displayName?: string; readonly webUrl?: string; readonly siteCollection?: unknown; readonly [key: string]: unknown; }
export interface GraphListMetadata { readonly id: string; readonly name?: string; readonly displayName?: string; readonly webUrl?: string; readonly createdDateTime?: string; readonly lastModifiedDateTime?: string; readonly list?: unknown; readonly [key: string]: unknown; }
export interface GraphListItem { readonly id: string; readonly fields?: Record<string, unknown>; readonly [key: string]: unknown; }
export interface GraphListItemCreate { readonly fields: Readonly<Record<string, unknown>>; }
export type GraphListItemUpdate = Readonly<Record<string, unknown>>;
export interface GraphListQuery extends GraphIterationOptions, GraphRequestOptions {
  readonly select?: readonly string[];
  readonly expand?: readonly string[];
  readonly filter?: string;
  readonly orderBy?: string | readonly [string, boolean][];
  readonly top?: number;
}
export interface GraphDirectoryObject { readonly id: string; readonly [key: string]: unknown; }
export interface GraphUser extends GraphDirectoryObject { readonly displayName?: string; readonly userPrincipalName?: string; readonly mail?: string; }
export interface GraphGroup extends GraphDirectoryObject { readonly displayName?: string; readonly description?: string; readonly mail?: string; readonly mailEnabled?: boolean; readonly securityEnabled?: boolean; }
export interface GraphDirectoryQuery extends GraphIterationOptions, GraphRequestOptions {
  readonly select?: readonly string[];
  readonly filter?: string;
  readonly orderBy?: string | readonly [string, boolean][];
  readonly top?: number;
}
export interface GraphSearchSortProperty { readonly name: string; readonly isDescending?: boolean; }
export interface GraphSearchRequest extends GraphIterationOptions, GraphRequestOptions {
  readonly entityTypes: readonly string[];
  readonly queryText: string;
  readonly fields?: readonly string[];
  readonly from?: number;
  readonly size?: number;
  readonly sortProperties?: readonly GraphSearchSortProperty[];
}
export interface GraphSearchHit<T = unknown> {
  readonly hitId: string;
  readonly rank?: number;
  readonly summary?: string;
  readonly resource?: T;
  readonly resultTemplateId?: string;
  readonly [key: string]: unknown;
}
export interface GraphSearchResult<T = unknown> {
  readonly data: readonly GraphSearchHit<T>[];
  readonly total: number;
  readonly moreResultsAvailable: boolean;
  readonly from: number;
  readonly size: number;
}
export type GraphDirectoryRequestOptions = GraphDirectoryQuery;
export type GraphDirectoryWriteOptions = GraphRequestOptions;
export type GraphSiteRequestOptions = GraphRequestOptions;
export type GraphListWriteOptions = GraphRequestOptions;

const jsonHeaders = { Accept: "application/json", "Content-Type": "application/json" };
const batchMethods = new Set(["GET", "POST", "PUT", "PATCH", "DELETE"]);
const asRecord = (value: unknown): Record<string, unknown> => typeof value === "object" && value !== null ? value as Record<string, unknown> : {};
const etagOf = (value: unknown): string | undefined => {
  const row = asRecord(value);
  const etag = row["@odata.etag"] ?? row["odata.etag"] ?? row.eTag ?? row.etag;
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
const graphSegment = (value: string, label: string): string => {
  if (typeof value !== "string" || value.length === 0) throw new DataError("validation", `${label} must be a non-empty path segment`);
  let decoded: string;
  try { decoded = decodeURIComponent(value); } catch (cause) { throw new DataError("validation", `${label} encoding is invalid`, cause); }
  if (decoded === "." || decoded === ".." || /[\\/?#]/.test(decoded)) throw new DataError("validation", `${label} must be a single path segment`);
  return encodeURIComponent(decoded);
};
const graphPath = (value: string): string[] => {
  if (typeof value !== "string" || value.length === 0 || value.startsWith("/") || value.endsWith("/")) throw new DataError("validation", "Graph drive file path must be relative and non-empty");
  const segments = value.split("/");
  return segments.map((segment) => graphSegment(segment, "Graph drive file path segment"));
};
const graphOptions = (options: GraphRequestOptions, accept = "application/json"): GraphRequestOptions => ({ ...options, headers: { Accept: accept, ...options.headers } });
const graphQuery = (path: string, options: GraphListQuery | GraphDirectoryQuery): string => {
  validateIteration(options);
  if (options.top !== undefined && (!Number.isInteger(options.top) || options.top < 0)) throw new DataError("validation", "top must be a non-negative integer");
  const values: string[] = [];
  const add = (name: string, value: string | number): void => { values.push(`${name}=${encodeURIComponent(String(value)).replaceAll("'", "%27")}`); };
  if (options.select?.length) add("$select", options.select.join(","));
  if ("expand" in options && options.expand?.length) add("$expand", options.expand.join(","));
  if (options.filter) add("$filter", options.filter);
  if (options.orderBy) {
    const orders = typeof options.orderBy === "string" ? [[options.orderBy, true] as [string, boolean]] : options.orderBy;
    add("$orderby", orders.map(([field, ascending]) => `${field} ${ascending ? "asc" : "desc"}`).join(","));
  }
  if (options.top !== undefined) add("$top", options.top);
  return values.length === 0 ? path : `${path}?${values.join("&")}`;
};
const emptyGraphPage = <T>(): GraphPage<T> => ({ value: [] });
const validateSearch = (request: GraphSearchRequest): void => {
  if (!Array.isArray(request.entityTypes) || request.entityTypes.length === 0 || request.entityTypes.some((type) => typeof type !== "string" || type.trim() === "")) throw new DataError("validation", "Graph search entityTypes must contain at least one non-empty string");
  if (typeof request.queryText !== "string" || request.queryText.trim() === "") throw new DataError("validation", "Graph search queryText must be a non-empty string");
  if (request.fields?.some((field) => typeof field !== "string" || field.trim() === "")) throw new DataError("validation", "Graph search fields must contain non-empty strings");
  if (request.sortProperties?.some((property) => typeof property.name !== "string" || property.name.trim() === "" || (property.isDescending !== undefined && typeof property.isDescending !== "boolean"))) throw new DataError("validation", "Graph search sortProperties are invalid");
  validateIteration(request);
  for (const [name, value] of [["from", request.from], ["size", request.size]] as const) {
    if (value !== undefined && (!Number.isInteger(value) || value < (name === "from" ? 0 : 1) || (name === "size" && value > 1000))) throw new DataError("validation", `Graph search ${name} is out of bounds`);
  }
};
const searchRequestBody = (request: GraphSearchRequest, from: number, size: number): Record<string, unknown> => ({
  requests: [{
    entityTypes: [...request.entityTypes],
    query: { queryString: request.queryText },
    ...(request.fields?.length ? { fields: [...request.fields] } : {}),
    from,
    size,
    ...(request.sortProperties?.length ? { sortProperties: request.sortProperties.map(({ name, isDescending }) => ({ name, ...(isDescending === undefined ? {} : { isDescending }) })) } : {}),
  }],
});
const searchResult = <T>(value: unknown, from: number, size: number): GraphSearchResult<T> => {
  const root = asRecord(value);
  const response = Array.isArray(root.value) ? asRecord(root.value[0]) : {};
  const containers = Array.isArray(response.hitsContainers) ? response.hitsContainers.map(asRecord) : [];
  const data = containers.flatMap((container) => Array.isArray(container.hits) ? container.hits as GraphSearchHit<T>[] : []);
  const total = containers.reduce((sum, container) => sum + (typeof container.total === "number" && Number.isInteger(container.total) ? container.total : 0), 0);
  const moreResultsAvailable = containers.some((container) => container.moreResultsAvailable === true);
  return { data, total, moreResultsAvailable, from, size };
};

export function graphDriveUrl(driveId: string): string { return `/drives/${graphSegment(driveId, "Graph drive id")}`; }
export function graphDriveRootUrl(driveId: string): string { return `${graphDriveUrl(driveId)}/root`; }
export function graphDriveItemUrl(driveId: string, itemId: string): string { return `${graphDriveUrl(driveId)}/items/${graphSegment(itemId, "Graph drive item id")}`; }
export function graphDriveChildrenUrl(driveId: string, itemId?: string): string { return `${itemId === undefined ? graphDriveRootUrl(driveId) : graphDriveItemUrl(driveId, itemId)}/children`; }
export function graphDriveContentUrl(driveId: string, itemId: string): string { return `${graphDriveItemUrl(driveId, itemId)}/content`; }
export function graphDriveUploadUrl(driveId: string, filePath: string): string { return `${graphDriveUrl(driveId)}/root:/${graphPath(filePath).join("/")}:/content`; }
export function graphSiteUrl(siteId: string): string { return `/sites/${graphSegment(siteId, "Graph site id")}`; }
export function graphSiteByPathUrl(hostname: string, sitePath: string): string {
  if (typeof sitePath !== "string" || !sitePath.startsWith("/") || sitePath.includes("?") || sitePath.includes("#")) throw new DataError("validation", "Graph site path must be an absolute server-relative path");
  let decoded: string;
  try { decoded = decodeURIComponent(sitePath); } catch (cause) { throw new DataError("validation", "Graph site path encoding is invalid", cause); }
  const encodedHostname = graphSegment(hostname, "Graph site hostname");
  if (decoded === "/") return `/sites/${encodedHostname}:/`;
  const segments = decoded.split("/");
  if (segments.length > 1 && segments.some((segment, index) => index > 0 && segment.length === 0)) throw new DataError("validation", "Graph site path contains an empty segment");
  return `/sites/${encodedHostname}:/${segments.slice(1).map((segment) => graphSegment(segment, "Graph site path segment")).join("/")}`;
}
export function graphSiteListsUrl(siteId: string): string { return `${graphSiteUrl(siteId)}/lists`; }
export function graphListUrl(siteId: string, listId: string): string { return `${graphSiteListsUrl(siteId)}/${graphSegment(listId, "Graph list id")}`; }
export function graphListItemsUrl(siteId: string, listId: string): string { return `${graphListUrl(siteId, listId)}/items`; }
export function graphListItemUrl(siteId: string, listId: string, itemId: string): string { return `${graphListItemsUrl(siteId, listId)}/${graphSegment(itemId, "Graph list item id")}`; }
export function graphListItemFieldsUrl(siteId: string, listId: string, itemId: string): string { return `${graphListItemUrl(siteId, listId, itemId)}/fields`; }
export function graphCurrentUserUrl(): string { return "/me"; }
export function graphUsersUrl(): string { return "/users"; }
export function graphUserUrl(userId: string): string { return `${graphUsersUrl()}/${graphSegment(userId, "Graph user id")}`; }
export function graphGroupsUrl(): string { return "/groups"; }
export function graphGroupUrl(groupId: string): string { return `${graphGroupsUrl()}/${graphSegment(groupId, "Graph group id")}`; }
export function graphGroupMembersUrl(groupId: string): string { return `${graphGroupUrl(groupId)}/members`; }
export function graphGroupMembersRefUrl(groupId: string, memberId: string): string { return `${graphGroupMembersUrl(groupId)}/${graphSegment(memberId, "Graph group member id")}/$ref`; }
export function graphDirectoryObjectUrl(objectId: string): string { return `/directoryObjects/${graphSegment(objectId, "Graph directory object id")}`; }

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

export class GraphSearchAdapter {
  private readonly graph: GraphAdapter;
  constructor(client: DataClient, options: GraphAdapterOptions = {}) { this.graph = new GraphAdapter(client, options); }

  async searchPage<T = unknown>(request: GraphSearchRequest): Promise<GraphSearchResult<T>> {
    validateSearch(request);
    const from = request.from ?? 0;
    const size = Math.min(request.size ?? 25, request.maxItems ?? Number.MAX_SAFE_INTEGER);
    if (request.maxItems === 0) return { data: [], total: 0, moreResultsAvailable: false, from, size };
    const result = await this.graph.request<unknown>("/search/query", {
      method: "POST",
      headers: request.headers,
      body: searchRequestBody(request, from, size),
      signal: request.signal,
      timeoutMs: request.timeoutMs,
    });
    return searchResult<T>(result.data, from, size);
  }

  async *pages<T = unknown>(request: GraphSearchRequest): AsyncIterable<GraphSearchResult<T>> {
    validateSearch(request);
    const maxPages = request.maxPages ?? 100;
    const maxItems = request.maxItems ?? Number.MAX_SAFE_INTEGER;
    if (maxPages === 0 || maxItems === 0) return;
    const seen = new Set<number>();
    let from = request.from ?? 0;
    let count = 0;
    let pageCount = 0;
    while (pageCount < maxPages && count < maxItems && !seen.has(from)) {
      seen.add(from);
      const page = await this.searchPage<T>({ ...request, from, maxItems: Math.min(maxItems - count, request.size ?? 25) });
      pageCount++;
      const data = page.data.slice(0, maxItems - count);
      count += data.length;
      yield data.length === page.data.length ? page : { ...page, data };
      if (!page.moreResultsAvailable || page.data.length === 0) break;
      from += page.data.length;
    }
  }

  async search<T = unknown>(request: GraphSearchRequest): Promise<GraphSearchResult<T>> {
    validateSearch(request);
    const from = request.from ?? 0;
    const size = Math.min(request.size ?? 25, request.maxItems ?? Number.MAX_SAFE_INTEGER);
    if (request.maxPages === 0 || request.maxItems === 0) return { data: [], total: 0, moreResultsAvailable: false, from, size };
    const data: GraphSearchHit<T>[] = [];
    let total = 0;
    let moreResultsAvailable = false;
    for await (const page of this.pages<T>(request)) {
      data.push(...page.data);
      total = page.total;
      moreResultsAvailable = page.moreResultsAvailable;
    }
    return { data, total, moreResultsAvailable, from, size };
  }
}

export const graphSmallUploadMaxBytes = 1_500_000;

export class GraphDriveAdapter {
  private readonly baseUrl: string;
  private readonly graph: GraphAdapter;
  constructor(private readonly client: DataClient, options: GraphAdapterOptions = {}) {
    this.baseUrl = options.baseUrl ?? "https://graph.microsoft.com/v1.0";
    this.graph = new GraphAdapter(client, options);
  }

  getDrive(driveId: string, options: GraphDriveRequestOptions = {}): Promise<DataResult<GraphDriveMetadata>> {
    return this.graph.request<GraphDriveMetadata>(graphDriveUrl(driveId), graphOptions(options));
  }
  getRoot(driveId: string, options: GraphDriveRequestOptions = {}): Promise<DataResult<GraphDriveItemMetadata>> {
    return this.graph.request<GraphDriveItemMetadata>(graphDriveRootUrl(driveId), graphOptions(options));
  }
  getItem(driveId: string, itemId: string, options: GraphDriveRequestOptions = {}): Promise<DataResult<GraphDriveItemMetadata>> {
    return this.graph.request<GraphDriveItemMetadata>(graphDriveItemUrl(driveId, itemId), graphOptions(options));
  }
  listChildren<T extends GraphDriveItemMetadata = GraphDriveItemMetadata>(driveId: string, itemIdOrOptions?: string | (GraphIterationOptions & GraphDriveRequestOptions), options: GraphIterationOptions & GraphDriveRequestOptions = {}): Promise<GraphPage<T>> {
    const itemId = typeof itemIdOrOptions === "string" ? itemIdOrOptions : undefined;
    const requestOptions = typeof itemIdOrOptions === "string" ? options : itemIdOrOptions ?? {};
    return this.graph.page<T>(graphDriveChildrenUrl(driveId, itemId), graphOptions(requestOptions));
  }
  children<T extends GraphDriveItemMetadata = GraphDriveItemMetadata>(driveId: string, itemIdOrOptions?: string | (GraphIterationOptions & GraphDriveRequestOptions), options: GraphIterationOptions & GraphDriveRequestOptions = {}): Promise<GraphPage<T>> {
    return this.listChildren<T>(driveId, itemIdOrOptions, options);
  }
  childrenPages<T extends GraphDriveItemMetadata = GraphDriveItemMetadata>(driveId: string, itemIdOrOptions?: string | (GraphIterationOptions & GraphDriveRequestOptions), options: GraphIterationOptions & GraphDriveRequestOptions = {}): AsyncIterable<GraphPage<T>> {
    const itemId = typeof itemIdOrOptions === "string" ? itemIdOrOptions : undefined;
    const requestOptions = typeof itemIdOrOptions === "string" ? options : itemIdOrOptions ?? {};
    return this.graph.pages<T>(graphDriveChildrenUrl(driveId, itemId), graphOptions(requestOptions));
  }
  iterateChildren<T extends GraphDriveItemMetadata = GraphDriveItemMetadata>(driveId: string, itemIdOrOptions?: string | (GraphIterationOptions & GraphDriveRequestOptions), options: GraphIterationOptions & GraphDriveRequestOptions = {}): AsyncIterable<T> {
    const itemId = typeof itemIdOrOptions === "string" ? itemIdOrOptions : undefined;
    const requestOptions = typeof itemIdOrOptions === "string" ? options : itemIdOrOptions ?? {};
    return this.graph.iterate<T>(graphDriveChildrenUrl(driveId, itemId), graphOptions(requestOptions));
  }
  downloadFile(driveId: string, itemId: string, options: GraphDriveRequestOptions = {}): Promise<Uint8Array> {
    const headers: RequestHeaders = { Accept: "*/*", ...options.headers, ...(options.etag === undefined ? {} : { "If-Match": options.etag }) };
    return this.client.requestBytes(absolute(this.baseUrl, graphDriveContentUrl(driveId, itemId)), { headers, signal: options.signal, timeoutMs: options.timeoutMs });
  }
  async uploadFile(driveId: string, filePath: string, content: Uint8Array, options: GraphDriveWriteOptions = {}): Promise<DataResult<GraphDriveItemMetadata>> {
    if (!(content instanceof Uint8Array)) throw new DataError("validation", "Graph drive file content must be a Uint8Array");
    if (content.byteLength > graphSmallUploadMaxBytes) throw new DataError("validation", `Graph simple uploads are limited to ${graphSmallUploadMaxBytes} bytes`);
    const response = await this.client.requestRaw({ url: absolute(this.baseUrl, graphDriveUploadUrl(driveId, filePath)), method: "PUT", headers: { Accept: "application/json", "Content-Type": "application/octet-stream", ...options.headers, ...(options.etag === undefined ? {} : { "If-Match": options.etag }) }, body: content, signal: options.signal, timeoutMs: options.timeoutMs });
    const value = parseJson<GraphDriveItemMetadata>(response.text, response.status);
    return { data: value, ...(etagOf(value) === undefined && headerValue(response.headers, "etag") === undefined ? {} : { etag: etagOf(value) ?? headerValue(response.headers, "etag") }) };
  }
  updateItem(driveId: string, itemId: string, input: GraphDriveItemUpdate, options: GraphDriveWriteOptions = {}): Promise<DataResult<GraphDriveItemMetadata>> {
    return this.graph.request<GraphDriveItemMetadata>(graphDriveItemUrl(driveId, itemId), { ...graphOptions(options), method: "PATCH", body: input });
  }
  deleteItem(driveId: string, itemId: string, options: GraphDriveWriteOptions = {}): Promise<DataResult<void>> {
    return this.graph.request<void>(graphDriveItemUrl(driveId, itemId), { ...graphOptions(options), method: "DELETE" });
  }
}

export class GraphSitesAdapter {
  private readonly graph: GraphAdapter;
  constructor(client: DataClient, options: GraphAdapterOptions = {}) { this.graph = new GraphAdapter(client, options); }

  getSite(siteId: string, options: GraphSiteRequestOptions = {}): Promise<DataResult<GraphSiteMetadata>> {
    return this.graph.request<GraphSiteMetadata>(graphSiteUrl(siteId), graphOptions(options));
  }
  getSiteByPath(hostname: string, sitePath: string, options: GraphSiteRequestOptions = {}): Promise<DataResult<GraphSiteMetadata>> {
    return this.graph.request<GraphSiteMetadata>(graphSiteByPathUrl(hostname, sitePath), graphOptions(options));
  }
  listLists<T extends GraphListMetadata = GraphListMetadata>(siteId: string, options: GraphListQuery = {}): Promise<GraphPage<T>> {
    const path = graphQuery(graphSiteListsUrl(siteId), options);
    return options.top === 0 || options.maxPages === 0 || options.maxItems === 0 ? Promise.resolve(emptyGraphPage<T>()) : this.graph.page<T>(path, graphOptions(options));
  }
  listsPages<T extends GraphListMetadata = GraphListMetadata>(siteId: string, options: GraphListQuery = {}): AsyncIterable<GraphPage<T>> {
    return this.graph.pages<T>(graphQuery(graphSiteListsUrl(siteId), options), graphOptions(options));
  }
  iterateLists<T extends GraphListMetadata = GraphListMetadata>(siteId: string, options: GraphListQuery = {}): AsyncIterable<T> {
    return this.graph.iterate<T>(graphQuery(graphSiteListsUrl(siteId), options), graphOptions(options));
  }
  getList<T extends GraphListMetadata = GraphListMetadata>(siteId: string, listId: string, options: GraphListQuery = {}): Promise<DataResult<T>> {
    return this.graph.request<T>(graphQuery(graphListUrl(siteId, listId), options), graphOptions(options));
  }
  listItems<T extends GraphListItem = GraphListItem>(siteId: string, listId: string, options: GraphListQuery = {}): Promise<GraphPage<T>> {
    const path = graphQuery(graphListItemsUrl(siteId, listId), options);
    return options.top === 0 || options.maxPages === 0 || options.maxItems === 0 ? Promise.resolve(emptyGraphPage<T>()) : this.graph.page<T>(path, graphOptions(options));
  }
  itemsPages<T extends GraphListItem = GraphListItem>(siteId: string, listId: string, options: GraphListQuery = {}): AsyncIterable<GraphPage<T>> {
    return this.graph.pages<T>(graphQuery(graphListItemsUrl(siteId, listId), options), graphOptions(options));
  }
  iterateItems<T extends GraphListItem = GraphListItem>(siteId: string, listId: string, options: GraphListQuery = {}): AsyncIterable<T> {
    return this.graph.iterate<T>(graphQuery(graphListItemsUrl(siteId, listId), options), graphOptions(options));
  }
  getItem<T extends GraphListItem = GraphListItem>(siteId: string, listId: string, itemId: string, options: GraphListQuery = {}): Promise<DataResult<T>> {
    return this.graph.request<T>(graphQuery(graphListItemUrl(siteId, listId, itemId), options), graphOptions(options));
  }
  createItem<T extends GraphListItem = GraphListItem>(siteId: string, listId: string, input: GraphListItemCreate, options: GraphListWriteOptions = {}): Promise<DataResult<T>> {
    return this.graph.request<T>(graphListItemsUrl(siteId, listId), { ...graphOptions(options), method: "POST", body: input });
  }
  updateItem<T extends Record<string, unknown> = Record<string, unknown>>(siteId: string, listId: string, itemId: string, input: GraphListItemUpdate, options: GraphListWriteOptions = {}): Promise<DataResult<T>> {
    return this.graph.request<T>(graphListItemFieldsUrl(siteId, listId, itemId), { ...graphOptions(options), method: "PATCH", body: input });
  }
  deleteItem(siteId: string, listId: string, itemId: string, options: GraphListWriteOptions = {}): Promise<DataResult<void>> {
    return this.graph.request<void>(graphListItemUrl(siteId, listId, itemId), { ...graphOptions(options), method: "DELETE" });
  }
}

export class GraphDirectoryAdapter {
  private readonly baseUrl: string;
  private readonly graph: GraphAdapter;
  constructor(client: DataClient, options: GraphAdapterOptions = {}) {
    this.baseUrl = options.baseUrl ?? "https://graph.microsoft.com/v1.0";
    this.graph = new GraphAdapter(client, options);
  }

  getCurrentUser<T extends GraphUser = GraphUser>(options: GraphDirectoryRequestOptions = {}): Promise<DataResult<T>> {
    return this.graph.request<T>(graphQuery(graphCurrentUserUrl(), options), graphOptions(options));
  }
  getUser<T extends GraphUser = GraphUser>(userId: string, options: GraphDirectoryRequestOptions = {}): Promise<DataResult<T>> {
    return this.graph.request<T>(graphQuery(graphUserUrl(userId), options), graphOptions(options));
  }
  listUsers<T extends GraphUser = GraphUser>(options: GraphDirectoryQuery = {}): Promise<GraphPage<T>> {
    const path = graphQuery(graphUsersUrl(), options);
    return options.top === 0 || options.maxPages === 0 || options.maxItems === 0 ? Promise.resolve(emptyGraphPage<T>()) : this.graph.page<T>(path, graphOptions(options));
  }
  usersPages<T extends GraphUser = GraphUser>(options: GraphDirectoryQuery = {}): AsyncIterable<GraphPage<T>> {
    return this.graph.pages<T>(graphQuery(graphUsersUrl(), options), graphOptions(options));
  }
  iterateUsers<T extends GraphUser = GraphUser>(options: GraphDirectoryQuery = {}): AsyncIterable<T> {
    return this.graph.iterate<T>(graphQuery(graphUsersUrl(), options), graphOptions(options));
  }
  listGroups<T extends GraphGroup = GraphGroup>(options: GraphDirectoryQuery = {}): Promise<GraphPage<T>> {
    const path = graphQuery(graphGroupsUrl(), options);
    return options.top === 0 || options.maxPages === 0 || options.maxItems === 0 ? Promise.resolve(emptyGraphPage<T>()) : this.graph.page<T>(path, graphOptions(options));
  }
  groupsPages<T extends GraphGroup = GraphGroup>(options: GraphDirectoryQuery = {}): AsyncIterable<GraphPage<T>> {
    return this.graph.pages<T>(graphQuery(graphGroupsUrl(), options), graphOptions(options));
  }
  iterateGroups<T extends GraphGroup = GraphGroup>(options: GraphDirectoryQuery = {}): AsyncIterable<T> {
    return this.graph.iterate<T>(graphQuery(graphGroupsUrl(), options), graphOptions(options));
  }
  getGroup<T extends GraphGroup = GraphGroup>(groupId: string, options: GraphDirectoryRequestOptions = {}): Promise<DataResult<T>> {
    return this.graph.request<T>(graphQuery(graphGroupUrl(groupId), options), graphOptions(options));
  }
  listMembers<T extends GraphDirectoryObject = GraphDirectoryObject>(groupId: string, options: GraphDirectoryQuery = {}): Promise<GraphPage<T>> {
    const path = graphQuery(graphGroupMembersUrl(groupId), options);
    return options.top === 0 || options.maxPages === 0 || options.maxItems === 0 ? Promise.resolve(emptyGraphPage<T>()) : this.graph.page<T>(path, graphOptions(options));
  }
  membersPages<T extends GraphDirectoryObject = GraphDirectoryObject>(groupId: string, options: GraphDirectoryQuery = {}): AsyncIterable<GraphPage<T>> {
    return this.graph.pages<T>(graphQuery(graphGroupMembersUrl(groupId), options), graphOptions(options));
  }
  iterateMembers<T extends GraphDirectoryObject = GraphDirectoryObject>(groupId: string, options: GraphDirectoryQuery = {}): AsyncIterable<T> {
    return this.graph.iterate<T>(graphQuery(graphGroupMembersUrl(groupId), options), graphOptions(options));
  }
  addMember(groupId: string, memberId: string, options: GraphDirectoryWriteOptions = {}): Promise<DataResult<void>> {
    return this.graph.request<void>(graphGroupMembersUrl(groupId) + "/$ref", {
      ...graphOptions(options),
      method: "POST",
      body: { "@odata.id": absolute(this.baseUrl, graphDirectoryObjectUrl(memberId)) },
    });
  }
  removeMember(groupId: string, memberId: string, options: GraphDirectoryWriteOptions = {}): Promise<DataResult<void>> {
    return this.graph.request<void>(graphGroupMembersRefUrl(groupId, memberId), { ...graphOptions(options), method: "DELETE" });
  }
}
