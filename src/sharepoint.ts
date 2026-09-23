import { DataError, headerValue, parseJson } from "./errors.js";
import { DataClient } from "./client.js";
import type { CrudResource, DataResult, ListQuery, ListResult, RemoveOptions, RequestHeaders, ResponseHeaders, UpdateOptions } from "./contracts.js";

type Row = Record<string, unknown>;
export type SharePointFieldMap<TEntity> = Partial<Record<keyof TEntity, string>>;
export interface SharePointRestOptions<TEntity, TCreate, TUpdate> {
  readonly siteUrl: string;
  readonly listTitle: string;
  readonly map?: SharePointFieldMap<TEntity>;
  readonly createMap?: (input: TCreate) => Row;
  readonly updateMap?: (input: TUpdate) => Row;
}
export interface SharePointBatchRequest { readonly id?: string; readonly method: "GET"; readonly url: string; readonly headers?: RequestHeaders; }
export interface SharePointBatchResponse { readonly id?: string; readonly status: number; readonly headers: RequestHeaders; readonly body?: unknown; readonly ok: boolean; }
export interface SharePointBatchResult { readonly responses: readonly SharePointBatchResponse[]; readonly failures: readonly SharePointBatchResponse[]; }
export interface SharePointRequestOptions { readonly signal?: AbortSignal; readonly timeoutMs?: number; }
export interface SharePointWriteOptions extends SharePointRequestOptions { readonly etag?: string; }
export interface SharePointFileMetadata { readonly Name: string; readonly ServerRelativeUrl: string; readonly Length?: number; readonly TimeCreated?: string; readonly TimeLastModified?: string; readonly UniqueId?: string; readonly [key: string]: unknown; }
export interface SharePointFolderMetadata { readonly Name: string; readonly ServerRelativeUrl: string; readonly ItemCount?: number; readonly TimeCreated?: string; readonly TimeLastModified?: string; readonly UniqueId?: string; readonly [key: string]: unknown; }
export interface SharePointAttachmentMetadata { readonly FileName: string; readonly ServerRelativeUrl: string; readonly TimeLastModified?: string; readonly UniqueId?: string; readonly [key: string]: unknown; }
export interface SharePointFolderChildren { readonly files: readonly SharePointFileMetadata[]; readonly folders: readonly SharePointFolderMetadata[]; }
export interface SharePointFilesOptions { readonly siteUrl: string; readonly listTitle?: string; }

const jsonAccept = "application/json;odata=nometadata";
const asRow = (value: unknown): Row => typeof value === "object" && value !== null ? value as Row : {};
const asText = (value: unknown): string | undefined => typeof value === "string" ? value : undefined;

export function sharePointListItemsUrl(siteUrl: string, listTitle: string): string {
  const base = siteUrl.replace(/\/+$/, "");
  const escaped = encodeURIComponent(listTitle.replaceAll("'", "''"));
  return `${base}/_api/web/lists/getbytitle('${escaped}')/items`;
}
export function sharePointItemUrl(siteUrl: string, listTitle: string, id: number): string {
  if (!Number.isInteger(id) || id < 1) throw new DataError("validation", "SharePoint item id must be a positive integer");
  return `${sharePointListItemsUrl(siteUrl, listTitle)}(${id})`;
}

function sitePath(siteUrl: string): { readonly base: string; readonly path: string } {
  let parsed: URL;
  try { parsed = new URL(siteUrl); } catch (cause) { throw new DataError("validation", "SharePoint site URL is invalid", cause); }
  const path = parsed.pathname.replace(/\/+$/, "");
  return { base: `${parsed.origin}${path}`, path };
}
function encodedServerRelativePath(siteUrl: string, value: string): string {
  const root = sitePath(siteUrl).path;
  if (/^[a-z][a-z\d+.-]*:\/\//i.test(value)) throw new DataError("validation", "SharePoint resource paths must be server-relative");
  const candidate = value.startsWith("/") ? value : `${root}/${value}`;
  let decoded: string;
  try { decoded = decodeURIComponent(candidate); } catch (cause) { throw new DataError("validation", "SharePoint resource path encoding is invalid", cause); }
  const segments = decoded.split("/");
  if (segments.some((segment) => segment === "." || segment === "..")) throw new DataError("validation", "SharePoint resource paths cannot contain dot segments");
  if (root && decoded !== root && !decoded.startsWith(`${root}/`)) throw new DataError("validation", "SharePoint resource path is outside the configured site");
  return segments.map((segment) => encodeURIComponent(segment.replaceAll("'", "''"))).join("/");
}
function odataString(value: string): string { return encodeURIComponent(value.replaceAll("'", "''")); }
function fileName(value: string): string {
  if (!value || value.includes("/") || value.includes("\\")) throw new DataError("validation", "SharePoint file names must be non-empty and contain no path separators");
  return odataString(value);
}
function resourcePath(siteUrl: string, value: string): string { return `'${encodedServerRelativePath(siteUrl, value)}'`; }

export function sharePointFileUrl(siteUrl: string, serverRelativeUrl: string): string {
  return `${sitePath(siteUrl).base}/_api/web/GetFileByServerRelativeUrl(${resourcePath(siteUrl, serverRelativeUrl)})`;
}
export function sharePointFolderUrl(siteUrl: string, serverRelativeUrl: string): string {
  return `${sitePath(siteUrl).base}/_api/web/GetFolderByServerRelativeUrl(${resourcePath(siteUrl, serverRelativeUrl)})`;
}
export function sharePointFolderChildrenUrl(siteUrl: string, serverRelativeUrl: string): string {
  return `${sharePointFolderUrl(siteUrl, serverRelativeUrl)}?$expand=Folders,Files`;
}
export function sharePointFileDownloadUrl(siteUrl: string, serverRelativeUrl: string): string { return `${sharePointFileUrl(siteUrl, serverRelativeUrl)}/$value`; }
export function sharePointFileUploadUrl(siteUrl: string, folderServerRelativeUrl: string, name: string, overwrite = false): string {
  return `${sharePointFolderUrl(siteUrl, folderServerRelativeUrl)}/Files/add(url='${fileName(name)}',overwrite=${overwrite ? "true" : "false"})`;
}
export function sharePointAttachmentFilesUrl(siteUrl: string, listTitle: string, id: number): string {
  return `${sharePointItemUrl(siteUrl, listTitle, id)}/AttachmentFiles`;
}
export function sharePointAttachmentUrl(siteUrl: string, listTitle: string, id: number, name: string): string {
  return `${sharePointAttachmentFilesUrl(siteUrl, listTitle, id)}('${fileName(name)}')`;
}
export function sharePointAttachmentDownloadUrl(siteUrl: string, listTitle: string, id: number, name: string): string {
  return `${sharePointAttachmentUrl(siteUrl, listTitle, id, name)}/$value`;
}
export function sharePointAttachmentUploadUrl(siteUrl: string, listTitle: string, id: number, name: string): string {
  return `${sharePointAttachmentFilesUrl(siteUrl, listTitle, id)}/add(FileName='${fileName(name)}')`;
}

function validateBounds(query: ListQuery | undefined): void {
  for (const [name, value] of [["pageSize", query?.pageSize], ["maxPages", query?.maxPages]] as const) {
    if (value !== undefined && (!Number.isInteger(value) || value < 0 || (name === "pageSize" && value < 1))) throw new DataError("validation", `${name} must be ${name === "maxPages" ? "a non-negative" : "a positive"} integer`);
  }
  if (query?.top !== undefined && (!Number.isInteger(query.top) || query.top < 0)) throw new DataError("validation", "top must be a non-negative integer");
}
function queryUrl(base: string, query: ListQuery): string {
  const values: string[] = [];
  const add = (name: string, value: string | number) => values.push(`${name}=${encodeURIComponent(String(value))}`);
  if (query.select?.length) add("$select", query.select.join(","));
  if (query.expand?.length) add("$expand", query.expand.join(","));
  if (query.filter) add("$filter", query.filter);
  if (query.orderBy) {
    const orders = typeof query.orderBy === "string" ? [[query.orderBy, true] as [string, boolean]] : query.orderBy;
    add("$orderby", orders.map(([field, ascending]) => `${field} ${ascending ? "asc" : "desc"}`).join(","));
  }
  const pageSize = query.pageSize ?? 100;
  add("$top", query.top === undefined ? pageSize : Math.min(query.top, pageSize));
  return `${base}?${values.join("&")}`;
}
function pageRows(value: unknown): Row[] {
  const root = asRow(value);
  const d = asRow(root.d);
  const rows = root.value ?? d.results;
  return Array.isArray(rows) ? rows.map(asRow) : [];
}
function nextLink(value: unknown): string | undefined {
  const root = asRow(value);
  return asText(root["@odata.nextLink"]) ?? asText(root["odata.nextLink"]) ?? asText(asRow(root.d).__next);
}
function rowId(row: Row): number | undefined {
  const id = row.Id ?? row.ID;
  return typeof id === "number" && Number.isInteger(id) ? id : undefined;
}
function rowEtag(row: Row): string | undefined { return asText(row["@odata.etag"]) ?? asText(row["odata.etag"]) ?? asText(row.ETag); }
function itemRow(value: unknown): Row { const row = asRow(value); return asRow(row.d ?? row); }

function parseBatchHeaders(value: string): RequestHeaders {
  const headers: Record<string, string> = {};
  for (const line of value.split(/\r?\n/)) { const index = line.indexOf(":"); if (index > 0) headers[line.slice(0, index).trim()] = line.slice(index + 1).trim(); }
  return headers;
}
function parseBatchResponse(text: string, contentType: string | undefined, requests: readonly SharePointBatchRequest[]): SharePointBatchResult {
  const match = contentType?.match(/boundary\s*=\s*(?:"([^"]+)"|([^;\s]+))/i);
  if (!match) throw new DataError("unknown", "SharePoint batch response did not include a boundary");
  const boundary = match[1] ?? match[2];
  const responses: SharePointBatchResponse[] = [];
  for (const part of text.split(`--${boundary}`).slice(1)) {
    if (part.trim() === "" || part.trim() === "--") continue;
    const http = part.match(/HTTP\/\d(?:\.\d)?\s+(\d{3})[^\r\n]*\r?\n([\s\S]*?)(?:\r?\n\r?\n)([\s\S]*)/i);
    if (!http) continue;
    const status = Number(http[1]);
    const headers = parseBatchHeaders(http[2]);
    const rawBody = http[3].replace(/\r?\n--?\s*$/, "").trim();
    let body: unknown = rawBody;
    if (rawBody) { try { body = JSON.parse(rawBody); } catch { /* non-JSON child response */ } }
    const index = responses.length;
    responses.push({ id: requests[index]?.id, status, headers, body, ok: status >= 200 && status < 300 });
  }
  return { responses, failures: responses.filter((response) => !response.ok) };
}

export class SharePointRestAdapter<TEntity = Row, TCreate = Partial<TEntity>, TUpdate = Partial<TEntity>> implements CrudResource<TEntity, TCreate, TUpdate> {
  private readonly base: string;
  constructor(private readonly client: DataClient, private readonly options: SharePointRestOptions<TEntity, TCreate, TUpdate>) {
    this.base = sharePointListItemsUrl(options.siteUrl, options.listTitle);
  }
  private map(raw: Row): TEntity {
    if (!this.options.map) return raw as TEntity;
    const output = {} as Row;
    for (const key of Object.keys(this.options.map) as (keyof TEntity)[]) { const field = this.options.map[key]; if (field) output[String(key)] = raw[field]; }
    return output as TEntity;
  }
  private result(raw: Row, responseEtag?: string): DataResult<TEntity> { const etag = rowEtag(raw) ?? responseEtag; return etag === undefined ? { data: this.map(raw) } : { data: this.map(raw), etag }; }
  private async read(url: string): Promise<DataResult<TEntity>> {
    const response = await this.client.requestRaw({ url, method: "GET", headers: { Accept: jsonAccept } });
    return this.result(itemRow(parseJson<unknown>(response.text, response.status)), headerValue(response.headers, "etag"));
  }
  async list(query: ListQuery = {}): Promise<ListResult<TEntity>> {
    validateBounds(query);
    if (query.top === 0 || query.maxPages === 0) return { data: [], etags: {} };
    const data: TEntity[] = [];
    const etags: Record<string, string> = {};
    let url = queryUrl(this.base, query);
    let pages = 0;
    const seenUrls = new Set<string>();
    while (url) {
      if (seenUrls.has(url)) break;
      seenUrls.add(url);
      const response = await this.client.requestRaw({ url, method: "GET", headers: { Accept: jsonAccept } });
      const payload = parseJson<unknown>(response.text, response.status);
      for (const raw of pageRows(payload)) {
        if (query.top !== undefined && data.length >= query.top) break;
        data.push(this.map(raw));
        const id = rowId(raw); const etag = rowEtag(raw); if (id !== undefined && etag !== undefined) etags[String(id)] = etag;
      }
      pages++;
      if ((query.maxPages !== undefined && pages >= query.maxPages) || (query.top !== undefined && data.length >= query.top)) break;
      const link = nextLink(payload); if (!link) break;
      url = /^[a-z][a-z\d+.-]*:/i.test(link) ? link : new URL(link, url).toString();
    }
    return { data, etags };
  }
  get(id: number): Promise<DataResult<TEntity>> { return this.read(sharePointItemUrl(this.options.siteUrl, this.options.listTitle, id)); }
  async create(input: TCreate): Promise<DataResult<TEntity>> {
    const response = await this.client.requestRaw({ url: this.base, method: "POST", headers: { Accept: jsonAccept, "Content-Type": "application/json;odata=nometadata" }, body: JSON.stringify(this.options.createMap?.(input) ?? input) });
    const raw = response.text.trim() ? asRow(parseJson<unknown>(response.text, response.status)) : {};
    const payload = asRow(raw.d ?? raw);
    const location = headerValue(response.headers, "location");
    const id = rowId(payload) ?? Number(location?.match(/(?:items\(|\/)(\d+)\)?(?:$|\?)/i)?.[1]);
    if (!Number.isInteger(id) || id < 1) throw new DataError("unknown", "SharePoint did not return the new item id", undefined, response.status);
    this.client.invalidateUrlPrefix(this.base);
    return this.get(id);
  }
  async update(id: number, input: TUpdate, options?: UpdateOptions): Promise<DataResult<TEntity>> {
    const url = sharePointItemUrl(this.options.siteUrl, this.options.listTitle, id);
    await this.client.requestRaw({ url, method: "POST", headers: { Accept: jsonAccept, "Content-Type": "application/json;odata=nometadata", "X-HTTP-Method": "MERGE", "IF-MATCH": options?.etag ?? "*" }, body: JSON.stringify(this.options.updateMap?.(input) ?? input) });
    this.client.invalidateUrlPrefix(this.base);
    return this.get(id);
  }
  async remove(id: number, options?: RemoveOptions): Promise<DataResult<void>> {
    const url = sharePointItemUrl(this.options.siteUrl, this.options.listTitle, id);
    await this.client.requestRaw({ url, method: "DELETE", headers: { Accept: jsonAccept, "IF-MATCH": options?.etag ?? "*" } });
    this.client.invalidateUrlPrefix(this.base);
    return { data: undefined };
  }
  async batch(requests: readonly SharePointBatchRequest[], boundary = "spfx-data-kit-batch"): Promise<SharePointBatchResult> {
    if (requests.some((request) => request.method !== "GET")) throw new DataError("validation", "SharePoint REST batches support GET children only");
    const lines: string[] = [];
    for (const request of requests) {
      const url = /^[a-z][a-z\d+.-]*:/i.test(request.url) ? request.url : new URL(request.url, this.options.siteUrl).toString();
      const childHeaders = Object.keys(request.headers ?? {}).some((name) => name.toLowerCase() === "accept") ? request.headers! : { Accept: jsonAccept, ...request.headers };
      lines.push(`--${boundary}`, "Content-Type: application/http", "Content-Transfer-Encoding: binary", "", `GET ${url} HTTP/1.1`, ...Object.entries(childHeaders).map(([name, value]) => `${name}: ${value}`), "", "");
    }
    lines.push(`--${boundary}--`, "");
    const response = await this.client.requestRaw({ url: `${this.options.siteUrl.replace(/\/+$/, "")}/_api/$batch`, method: "POST", headers: { Accept: jsonAccept, "Content-Type": `multipart/mixed; boundary=${boundary}` }, body: lines.join("\r\n") });
    return parseBatchResponse(response.text, headerValue(response.headers, "content-type"), requests);
  }
}

export const sharePointSmallUploadMaxBytes = 1_500_000;

function malformed(label: string, status: number): DataError { return new DataError("unknown", `SharePoint returned malformed ${label} response`, undefined, status); }
function metadataRow<T extends Row>(value: unknown, label: string, required: string, status: number): T {
  const row = itemRow(value);
  if (typeof row[required] !== "string" || typeof row.ServerRelativeUrl !== "string") throw malformed(label, status);
  return row as T;
}
function metadataResult<T extends Row>(response: { readonly status: number; readonly headers: ResponseHeaders; readonly text: string }, label: string, required: string): DataResult<T> {
  const row = metadataRow<T>(parseJson<unknown>(response.text, response.status), label, required, response.status);
  const etag = rowEtag(row) ?? headerValue(response.headers, "etag");
  return etag === undefined ? { data: row } : { data: row, etag };
}
function collectionRows(value: unknown, label: string, status: number): Row[] {
  const collection = itemRow(value);
  const rows = Array.isArray(collection.value) ? collection.value : collection.results;
  if (!Array.isArray(rows)) throw malformed(label, status);
  return rows.map((row) => {
    if (typeof row !== "object" || row === null) throw malformed(label, status);
    return row as Row;
  });
}
function withRequestOptions(options?: SharePointRequestOptions): Pick<SharePointRequestOptions, "signal" | "timeoutMs"> {
  return options === undefined ? {} : { signal: options.signal, timeoutMs: options.timeoutMs };
}
function withEtag(headers: RequestHeaders, etag: string | undefined): RequestHeaders { return etag === undefined ? headers : { ...headers, "IF-MATCH": etag }; }

export class SharePointFilesAdapter {
  constructor(private readonly client: DataClient, private readonly options: SharePointFilesOptions) { sitePath(options.siteUrl); }

  private async json(url: string, options?: SharePointRequestOptions) {
    return this.client.requestRaw({ url, method: "GET", headers: { Accept: jsonAccept }, ...withRequestOptions(options) });
  }
  private listTitle(): string {
    if (!this.options.listTitle) throw new DataError("validation", "SharePoint listTitle is required for attachments");
    return this.options.listTitle;
  }
  async getFile(serverRelativeUrl: string, options?: SharePointRequestOptions): Promise<DataResult<SharePointFileMetadata>> {
    return metadataResult(await this.json(sharePointFileUrl(this.options.siteUrl, serverRelativeUrl), options), "file", "Name");
  }
  async getFolder(serverRelativeUrl: string, options?: SharePointRequestOptions): Promise<DataResult<SharePointFolderMetadata>> {
    return metadataResult(await this.json(sharePointFolderUrl(this.options.siteUrl, serverRelativeUrl), options), "folder", "Name");
  }
  async folderChildren(serverRelativeUrl: string, options?: SharePointRequestOptions): Promise<DataResult<SharePointFolderChildren>> {
    const response = await this.json(sharePointFolderChildrenUrl(this.options.siteUrl, serverRelativeUrl), options);
    const root = itemRow(parseJson<unknown>(response.text, response.status));
    if (!("Files" in root) || !("Folders" in root)) throw malformed("folder children", response.status);
    const files = collectionRows(root.Files, "folder files", response.status).map((row) => metadataRow<SharePointFileMetadata>(row, "file", "Name", response.status));
    const folders = collectionRows(root.Folders, "folder folders", response.status).map((row) => metadataRow<SharePointFolderMetadata>(row, "folder", "Name", response.status));
    const etag = rowEtag(root) ?? headerValue(response.headers, "etag");
    const data = { files, folders };
    return etag === undefined ? { data } : { data, etag };
  }
  async downloadFile(serverRelativeUrl: string, options?: SharePointRequestOptions): Promise<Uint8Array> {
    return this.client.requestBytes(sharePointFileDownloadUrl(this.options.siteUrl, serverRelativeUrl), { headers: { Accept: "*/*" }, ...withRequestOptions(options) });
  }
  async uploadFile(folderServerRelativeUrl: string, name: string, content: Uint8Array, options: SharePointWriteOptions & { readonly overwrite?: boolean } = {}): Promise<DataResult<SharePointFileMetadata>> {
    if (!(content instanceof Uint8Array)) throw new DataError("validation", "SharePoint file content must be a Uint8Array");
    if (content.byteLength > sharePointSmallUploadMaxBytes) throw new DataError("validation", `SharePoint small uploads are limited to ${sharePointSmallUploadMaxBytes} bytes`);
    const url = sharePointFileUploadUrl(this.options.siteUrl, folderServerRelativeUrl, name, options.overwrite ?? false);
    const response = await this.client.requestRaw({ url, method: "POST", headers: withEtag({ Accept: jsonAccept, "Content-Type": "application/octet-stream" }, options.etag), body: content, ...withRequestOptions(options) });
    return metadataResult(response, "file upload", "Name");
  }
  async listAttachments(itemId: number, options?: SharePointRequestOptions): Promise<ListResult<SharePointAttachmentMetadata>> {
    const response = await this.json(sharePointAttachmentFilesUrl(this.options.siteUrl, this.listTitle(), itemId), options);
    const rows = collectionRows(parseJson<unknown>(response.text, response.status), "attachments", response.status).map((row) => metadataRow<SharePointAttachmentMetadata>(row, "attachment", "FileName", response.status));
    const etags: Record<string, string> = {};
    for (const row of rows) { const etag = rowEtag(row); if (etag) etags[row.FileName] = etag; }
    return { data: rows, etags };
  }
  async downloadAttachment(itemId: number, name: string, options?: SharePointRequestOptions): Promise<Uint8Array> {
    return this.client.requestBytes(sharePointAttachmentDownloadUrl(this.options.siteUrl, this.listTitle(), itemId, name), { headers: { Accept: "*/*" }, ...withRequestOptions(options) });
  }
  async uploadAttachment(itemId: number, name: string, content: Uint8Array, options: SharePointWriteOptions = {}): Promise<DataResult<SharePointAttachmentMetadata>> {
    if (!(content instanceof Uint8Array)) throw new DataError("validation", "SharePoint attachment content must be a Uint8Array");
    if (content.byteLength > sharePointSmallUploadMaxBytes) throw new DataError("validation", `SharePoint small uploads are limited to ${sharePointSmallUploadMaxBytes} bytes`);
    const response = await this.client.requestRaw({ url: sharePointAttachmentUploadUrl(this.options.siteUrl, this.listTitle(), itemId, name), method: "POST", headers: withEtag({ Accept: jsonAccept, "Content-Type": "application/octet-stream" }, options.etag), body: content, ...withRequestOptions(options) });
    return metadataResult(response, "attachment upload", "FileName");
  }
  async deleteAttachment(itemId: number, name: string, options: SharePointWriteOptions = {}): Promise<DataResult<void>> {
    await this.client.requestRaw({ url: sharePointAttachmentUrl(this.options.siteUrl, this.listTitle(), itemId, name), method: "DELETE", headers: { Accept: jsonAccept, "IF-MATCH": options.etag ?? "*" }, ...withRequestOptions(options) });
    return { data: undefined };
  }
}
